#!/usr/bin/env python3
"""
M365 Session Pool - Native device code auth with PowerShell sessions.

All modules use their native device code flow - no MSAL token juggling.
"""

import json
import logging
import os
import re
import subprocess
import threading
import time
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional, Dict, Any, List
from flask import Flask, jsonify, request

# Metrics
class Metrics:
    def __init__(self):
        self.start_time = datetime.now()
        self.request_count = 0
        self.error_count = 0
        self.auth_count = 0
        self.response_times: List[float] = []  # Last 100 response times
        self.lock = threading.Lock()

    def record_request(self, duration: float, error: bool = False):
        with self.lock:
            self.request_count += 1
            if error:
                self.error_count += 1
            self.response_times.append(duration)
            if len(self.response_times) > 100:
                self.response_times.pop(0)

    def record_auth(self):
        with self.lock:
            self.auth_count += 1

    def get_stats(self) -> dict:
        with self.lock:
            uptime = (datetime.now() - self.start_time).total_seconds()
            avg_response = sum(self.response_times) / len(self.response_times) if self.response_times else 0
            return {
                "uptime_seconds": round(uptime, 1),
                "uptime_human": f"{int(uptime // 3600)}h {int((uptime % 3600) // 60)}m {int(uptime % 60)}s",
                "total_requests": self.request_count,
                "total_errors": self.error_count,
                "error_rate": round(self.error_count / self.request_count * 100, 2) if self.request_count else 0,
                "total_auths": self.auth_count,
                "avg_response_ms": round(avg_response * 1000, 1),
                "last_100_responses": len(self.response_times),
            }

metrics = Metrics()

# Configuration
HOST = os.getenv("SESSION_POOL_HOST", "0.0.0.0")
PORT = int(os.getenv("SESSION_POOL_PORT", "5200"))
# Single-connection mode: if set, only handle this connection
SINGLE_CONNECTION = os.getenv("M365_CONNECTION", "")
LOG_LEVEL = os.getenv("SESSION_POOL_LOG_LEVEL", "INFO")
COMMAND_TIMEOUT = int(os.getenv("COMMAND_TIMEOUT", "300"))

# Logging
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL),
    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger("session_pool")

# Connection registry
REGISTRY_PATH = os.path.expanduser("~/.m365-connections.json")

# Module configurations - all use native device code
MODULES = {
    "exo": {
        "name": "Exchange Online",
        "connect_cmd": "Connect-ExchangeOnline -Device -ShowBanner:$false",
        "health_cmd": "Get-ConnectionInformation | Select-Object -First 1 | ConvertTo-Json",
        "health_pattern": r"(Organization|TenantId)",
        "device_code_pattern": r"code\s+([A-Z0-9]{8,})",
    },
    "azure": {
        "name": "Azure PowerShell",
        "connect_cmd": "Connect-AzAccount -UseDeviceAuthentication -TenantId {tenant_id}",
        "health_cmd": "(Get-AzContext) | Select-Object Name, Account | ConvertTo-Json",
        "health_pattern": r"(Name|Account)",
        "device_code_pattern": r"code\s+([A-Z0-9]{8,})",
    },
    "teams": {
        "name": "Microsoft Teams",
        "connect_cmd": "Connect-MicrosoftTeams -UseDeviceAuthentication -TenantId {tenant_id}",
        "health_cmd": "Get-CsTenant | Select-Object TenantId, DisplayName | ConvertTo-Json",
        "health_pattern": r"(TenantId|DisplayName)",
        "device_code_pattern": r"code\s+([A-Z0-9]{8,})",
    },
    "pnp": {
        "name": "PnP PowerShell",
        "connect_cmd": 'Connect-PnPOnline -Url "https://{sharepoint_host}.sharepoint.com" -DeviceLogin -ClientId "{app_id}" -Tenant "{tenant}"',
        "health_cmd": "Get-PnPConnection | Select-Object Url, ConnectionType | ConvertTo-Json",
        "health_pattern": r"(Url|ConnectionType)",
        "device_code_pattern": r"code\s+([A-Z0-9]{8,})",
    },
    # Power Platform disabled - PAC CLI device code only works with its first-party app
    # To re-enable, would need service principal auth (--applicationId + --clientSecret + --tenant)
    # "powerplatform": {
    #     "name": "Power Platform (PAC CLI)",
    #     "use_pac": True,
    #     "connect_cmd": "pac auth create --deviceCode --tenant {tenant_id}",
    #     "health_cmd": "pac auth list",
    #     "health_pattern": r"(UNIVERSAL|Active)",
    #     "device_code_pattern": r"code\s+([A-Z0-9]{8,})",
    # },
}

# Command marker for PowerShell output sync
MARKER = "___M365_DONE___"


def load_connection_registry() -> Dict[str, Any]:
    """Load connection registry."""
    try:
        if os.path.exists(REGISTRY_PATH):
            with open(REGISTRY_PATH) as f:
                return json.load(f)
    except Exception as e:
        logger.error(f"Failed to load registry: {e}")
    return {"connections": {}}


def get_connection_config(connection_name: str) -> Optional[Dict[str, Any]]:
    """Get config for a named connection."""
    registry = load_connection_registry()
    return registry.get("connections", {}).get(connection_name)


@dataclass
class Session:
    """A PowerShell session with native device code auth."""

    tenant: str
    module: str
    connection_name: str
    app_id: str

    # Process
    process: Optional[subprocess.Popen] = None
    process_lock: threading.Lock = field(default_factory=threading.Lock)

    # State
    state: str = "initializing"  # initializing, ready, auth_pending, authenticated, error
    device_code: Optional[str] = None
    auth_initiated_by: Optional[str] = None

    # Identity
    authenticated_as: Optional[str] = None

    # Tracking
    last_command: Optional[datetime] = None
    last_error: Optional[str] = None

    def __post_init__(self):
        self.process_lock = threading.Lock()

    @property
    def session_id(self) -> str:
        return f"{self.connection_name}/{self.module}"

    def start_process(self) -> bool:
        """Start the PowerShell process."""
        try:
            module_config = MODULES.get(self.module)
            if not module_config:
                self.last_error = f"Unknown module: {self.module}"
                self.state = "error"
                return False

            if module_config.get("use_pac"):
                self.state = "ready"
                return True

            # Use stdbuf to force line-buffered stdout, or pwsh will buffer in non-interactive mode
            self.process = subprocess.Popen(
                ["stdbuf", "-oL", "pwsh", "-NoLogo", "-NoProfile", "-NoExit", "-Command", "-"],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
            )

            self._send_raw('$ErrorActionPreference = "Continue"')
            self._send_raw('function prompt { "" }')

            self.state = "ready"
            logger.info(f"[{self.session_id}] PowerShell started (PID: {self.process.pid})")
            return True

        except Exception as e:
            self.last_error = str(e)
            self.state = "error"
            logger.error(f"[{self.session_id}] Failed to start: {e}")
            return False

    def _send_raw(self, command: str, timeout: int = 30) -> str:
        """Send command and read output until marker."""
        if not self.process or self.process.poll() is not None:
            raise RuntimeError("Process not running")

        # Force flush stdout after command to prevent buffering issues
        full_cmd = f'{command}; [Console]::Out.Flush(); Write-Host "{MARKER}"\n'
        logger.info(f"[{self.session_id}] Sending command: {command[:80]}")
        self.process.stdin.write(full_cmd)
        self.process.stdin.flush()

        output_lines = []
        start = time.time()

        while True:
            if time.time() - start > timeout:
                logger.error(f"[{self.session_id}] Command timed out. Output so far: {output_lines[:5]}")
                raise TimeoutError(f"Command timed out after {timeout}s")

            line = self.process.stdout.readline()
            if not line:
                time.sleep(0.01)
                continue

            line = line.rstrip('\n\r')
            logger.debug(f"[{self.session_id}] Output line: {line[:100]}")
            if MARKER in line:
                break
            output_lines.append(line)

        logger.info(f"[{self.session_id}] Command completed, {len(output_lines)} lines")
        return '\n'.join(output_lines)

    def initiate_auth(self, caller_id: str) -> Dict[str, Any]:
        """Start native device code authentication."""
        module_config = MODULES.get(self.module)
        conn_config = get_connection_config(self.connection_name)

        self.state = "auth_pending"
        self.auth_initiated_by = caller_id
        self.device_code = None

        try:
            # Build connect command with placeholders
            sharepoint_host = conn_config.get("sharepoint_host", self.tenant.split('.')[0]) if conn_config else self.tenant.split('.')[0]
            tenant_id = conn_config.get("tenantId", self.tenant) if conn_config else self.tenant
            connect_cmd = module_config["connect_cmd"].format(
                tenant=self.tenant,
                tenant_id=tenant_id,
                sharepoint_host=sharepoint_host,
                app_id=self.app_id,
            )

            logger.info(f"[{self.session_id}] Starting auth: {connect_cmd}")

            if module_config.get("use_pac"):
                return self._initiate_pac_auth(connect_cmd, module_config)

            # Start connect command in background thread to capture device code
            device_code = None
            auth_complete = threading.Event()
            output_buffer = []

            # Flag to signal reader thread to stop
            stop_reader = threading.Event()
            self._auth_stop_reader = stop_reader

            def reader_thread():
                nonlocal device_code
                try:
                    # Send connect command with marker
                    self.process.stdin.write(f'{connect_cmd}; Write-Host "{MARKER}"\n')
                    self.process.stdin.flush()

                    start = time.time()
                    while time.time() - start < 120 and not stop_reader.is_set():
                        line = self.process.stdout.readline()
                        if not line:
                            time.sleep(0.1)
                            continue

                        line = line.strip()
                        output_buffer.append(line)
                        logger.info(f"[{self.session_id}] Auth output: {line[:100]}")

                        # Look for device code
                        if not device_code:
                            match = re.search(module_config["device_code_pattern"], line, re.IGNORECASE)
                            if match:
                                device_code = match.group(1)
                                self.device_code = device_code
                                logger.info(f"[{self.session_id}] Device code: {device_code}")

                        # Check for auth completion - MARKER means connect command finished
                        if MARKER in line:
                            if device_code:  # Only after we've seen device code
                                auth_complete.set()
                                logger.info(f"[{self.session_id}] Auth connect completed")
                            break

                except Exception as e:
                    logger.error(f"[{self.session_id}] Reader error: {e}")

            reader = threading.Thread(target=reader_thread, daemon=True)
            reader.start()

            # Wait for device code to appear (up to 30 sec)
            for _ in range(60):
                if self.device_code:
                    break
                time.sleep(0.5)

            if self.device_code:
                return {
                    "status": "auth_required",
                    "device_code": self.device_code,
                    "auth_url": "https://microsoft.com/devicelogin",
                    "message": f"To sign in, use a web browser to open https://microsoft.com/devicelogin and enter the code {self.device_code}",
                    "expires_in": 900,
                }
            else:
                self.state = "error"
                self.last_error = "No device code received"
                return {"status": "error", "error": "No device code received", "output": '\n'.join(output_buffer[:10])}

        except Exception as e:
            self.state = "error"
            self.last_error = str(e)
            logger.error(f"[{self.session_id}] Auth initiation failed: {e}")
            return {"status": "error", "error": str(e)}

    def _initiate_pac_auth(self, connect_cmd: str, module_config: dict) -> Dict[str, Any]:
        """Handle PAC CLI auth - uses Popen since PAC waits for auth completion."""
        try:
            # Start PAC in background - it blocks until auth completes
            self.process = subprocess.Popen(
                connect_cmd,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
            )

            # Read output until we find device code (or timeout)
            output_lines = []
            start = time.time()
            while time.time() - start < 30:
                line = self.process.stdout.readline()
                if not line:
                    if self.process.poll() is not None:
                        break
                    time.sleep(0.1)
                    continue

                output_lines.append(line.strip())
                logger.info(f"[{self.session_id}] PAC output: {line.strip()}")

                match = re.search(module_config["device_code_pattern"], line, re.IGNORECASE)
                if match:
                    self.device_code = match.group(1)
                    return {
                        "status": "auth_required",
                        "device_code": self.device_code,
                        "auth_url": "https://microsoft.com/devicelogin",
                        "message": f"Complete PAC CLI auth with code {self.device_code}",
                        "expires_in": 900,
                    }

            return {"status": "error", "error": "No device code in PAC output", "output": '\n'.join(output_lines)}

        except Exception as e:
            return {"status": "error", "error": str(e)}

    def check_auth_complete(self) -> bool:
        """Check if authentication completed."""
        module_config = MODULES.get(self.module)

        if module_config.get("use_pac"):
            return self._check_pac_auth()

        try:
            # Stop the auth reader thread so it doesn't compete for stdout
            if hasattr(self, '_auth_stop_reader'):
                self._auth_stop_reader.set()
                time.sleep(0.5)  # Give reader thread time to exit

            # Try health check to see if connected
            with self.process_lock:
                # Drain any leftover output from auth
                import select
                while select.select([self.process.stdout], [], [], 0.1)[0]:
                    self.process.stdout.readline()

                # First sync with a marker
                self._send_raw("Write-Host 'sync'", timeout=5)

                # Now check health
                health_output = self._send_raw(module_config["health_cmd"], timeout=30)
                logger.info(f"[{self.session_id}] Health check: {health_output[:200]}")

                if re.search(module_config["health_pattern"], health_output):
                    self.state = "authenticated"
                    self.device_code = None
                    self.auth_initiated_by = None

                    # Extract authenticated identity from health output
                    try:
                        health_data = json.loads(health_output.strip())
                        if self.module == "exo":
                            self.authenticated_as = health_data.get("UserPrincipalName") or health_data.get("Organization")
                        elif self.module == "azure":
                            account = health_data.get("Account", {})
                            self.authenticated_as = account.get("Id") if isinstance(account, dict) else str(account)
                        elif self.module == "teams":
                            self.authenticated_as = health_data.get("DisplayName")
                        elif self.module == "pnp":
                            self.authenticated_as = health_data.get("Url")
                    except (json.JSONDecodeError, AttributeError):
                        pass  # Health output may not be clean JSON

                    logger.info(f"[{self.session_id}] Auth completed! Identity: {self.authenticated_as}")
                    return True

        except Exception as e:
            logger.debug(f"[{self.session_id}] Auth check: {e}")

        return False

    def _check_pac_auth(self) -> bool:
        """Check if PAC auth completed."""
        try:
            result = subprocess.run("pac auth list", shell=True, capture_output=True, text=True, timeout=10)
            if "Active" in result.stdout or "UNIVERSAL" in result.stdout:
                self.state = "authenticated"
                self.device_code = None
                logger.info(f"[{self.session_id}] PAC auth completed")
                return True
        except:
            pass
        return False

    def run_command(self, command: str, caller_id: str, timeout: int = COMMAND_TIMEOUT) -> Dict[str, Any]:
        """Execute a command."""
        module_config = MODULES.get(self.module)

        # Check auth state
        if self.state == "auth_pending":
            if self.auth_initiated_by == caller_id:
                if self.check_auth_complete():
                    pass  # Continue to execute
                else:
                    return {
                        "status": "auth_required",
                        "device_code": self.device_code,
                        "auth_url": "https://microsoft.com/devicelogin",
                        "message": "Complete device code auth, then retry",
                    }
            else:
                return {
                    "status": "auth_in_progress",
                    "message": "Auth in progress by another caller, retry shortly"
                }

        if self.state != "authenticated":
            return self.initiate_auth(caller_id)

        # Execute command
        with self.process_lock:
            try:
                if module_config.get("use_pac"):
                    result = subprocess.run(command, shell=True, capture_output=True, text=True, timeout=timeout)
                    output = result.stdout + result.stderr
                else:
                    output = self._send_raw(command, timeout=timeout)

                self.last_command = datetime.now()
                response = {"status": "success", "output": output}
                if self.authenticated_as:
                    response["authenticated_as"] = self.authenticated_as
                return response

            except TimeoutError as e:
                return {"status": "error", "error": f"Timeout: {e}"}
            except Exception as e:
                self.last_error = str(e)
                return {"status": "error", "error": str(e)}

    def stop(self):
        """Stop the session."""
        if self.process:
            try:
                self.process.terminate()
                self.process.wait(timeout=5)
            except:
                self.process.kill()
        self.state = "stopped"


class SessionPool:
    """Manages pool of sessions."""

    def __init__(self):
        self.sessions: Dict[str, Session] = {}
        self.lock = threading.Lock()

    def get_or_create_session(self, connection_name: str, module: str) -> Session:
        """Get or create a session."""
        session_id = f"{connection_name}/{module}"

        with self.lock:
            if session_id not in self.sessions:
                registry = load_connection_registry()
                conn_config = registry.get("connections", {}).get(connection_name)
                if not conn_config:
                    raise ValueError(f"Connection '{connection_name}' not found")

                # Check for module-specific app (e.g., EXO uses Microsoft's built-in app)
                known_module_apps = registry.get("_knownModuleApps", {})
                module_apps = conn_config.get("moduleApps", {})

                # Priority: connection's moduleApps > registry's _knownModuleApps > connection's appId
                app_id = module_apps.get(module) or known_module_apps.get(module) or conn_config.get("appId", "")

                logger.info(f"[{session_id}] Using app_id: {app_id}")

                session = Session(
                    tenant=conn_config.get("tenant", ""),
                    module=module,
                    connection_name=connection_name,
                    app_id=app_id,
                )
                session.start_process()
                self.sessions[session_id] = session
                logger.info(f"Created session: {session_id}")

            return self.sessions[session_id]

    def run_command(self, connection_name: str, module: str, command: str, caller_id: str) -> Dict[str, Any]:
        """Run a command."""
        try:
            session = self.get_or_create_session(connection_name, module)
            return session.run_command(command, caller_id)
        except Exception as e:
            return {"status": "error", "error": str(e)}

    def reset_connection(self, connection_name: str, module: str = None) -> Dict[str, Any]:
        """Reset sessions for a connection. If module specified, reset only that session."""
        with self.lock:
            reset = []
            to_remove = []
            for session_id, session in self.sessions.items():
                if session.connection_name == connection_name:
                    if module and session.module != module:
                        continue
                    session.stop()
                    to_remove.append(session_id)
                    reset.append(session_id)
            for sid in to_remove:
                del self.sessions[sid]
            return {"reset": reset, "count": len(reset)}

    def get_status(self) -> Dict[str, Any]:
        """Get status of all sessions."""
        with self.lock:
            return {
                "sessions": [
                    {
                        "session_id": s.session_id,
                        "state": s.state,
                        "authenticated": s.state == "authenticated",
                        "last_command": s.last_command.isoformat() if s.last_command else None,
                    }
                    for s in self.sessions.values()
                ]
            }


# Global pool
pool = SessionPool()


# Keepalive thread - pings authenticated sessions to prevent token expiry
class SessionKeepalive:
    def __init__(self, pool: SessionPool, interval: int = 300):
        self.pool = pool
        self.interval = interval  # seconds between keepalive pings
        self.running = True
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.last_ping = {}
        self.ping_count = 0
        self.ping_failures = 0

    def start(self):
        self.thread.start()
        logger.info(f"Keepalive started (interval: {self.interval}s)")

    def stop(self):
        self.running = False

    def _run(self):
        while self.running:
            time.sleep(self.interval)
            self._ping_sessions()

    def _ping_sessions(self):
        with self.pool.lock:
            sessions = list(self.pool.sessions.values())

        for session in sessions:
            if session.state != "authenticated":
                continue

            try:
                module_config = MODULES.get(session.module)
                if not module_config or module_config.get("use_pac"):
                    continue

                # Run health check to keep session alive
                with session.process_lock:
                    health_output = session._send_raw(module_config["health_cmd"], timeout=30)

                if re.search(module_config["health_pattern"], health_output):
                    self.last_ping[session.session_id] = datetime.now()
                    self.ping_count += 1
                    logger.debug(f"[{session.session_id}] Keepalive OK")
                else:
                    self.ping_failures += 1
                    logger.warning(f"[{session.session_id}] Keepalive failed: {health_output[:100]}")

            except Exception as e:
                self.ping_failures += 1
                logger.error(f"[{session.session_id}] Keepalive error: {e}")

    def get_stats(self) -> dict:
        return {
            "interval_seconds": self.interval,
            "total_pings": self.ping_count,
            "failures": self.ping_failures,
            "last_pings": {k: v.isoformat() for k, v in self.last_ping.items()},
        }


# Start keepalive (5 min interval)
keepalive = SessionKeepalive(pool, interval=300)
keepalive.start()

# Flask app
app = Flask(__name__)


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "healthy", "timestamp": datetime.now().isoformat()})


@app.route("/status", methods=["GET"])
def status():
    return jsonify(pool.get_status())


@app.route("/connections", methods=["GET"])
def list_connections():
    registry = load_connection_registry()
    return jsonify({"connections": registry.get("connections", {})})


@app.route("/run", methods=["POST"])
def run_command():
    start = time.time()
    data = request.get_json()
    module = data.get("module")
    command = data.get("command")
    caller_id = data.get("caller_id", "anonymous")

    # Single-connection mode: use env var, ignore request connection
    if SINGLE_CONNECTION:
        connection = SINGLE_CONNECTION
    else:
        connection = data.get("connection")

    if not all([connection, module, command]):
        metrics.record_request(time.time() - start, error=True)
        return jsonify({"status": "error", "error": "Missing connection, module, or command"}), 400

    # In single-connection mode, reject requests for other connections
    if SINGLE_CONNECTION and connection != SINGLE_CONNECTION:
        metrics.record_request(time.time() - start, error=True)
        return jsonify({"status": "error", "error": f"This container only handles {SINGLE_CONNECTION}"}), 400

    if module not in MODULES:
        metrics.record_request(time.time() - start, error=True)
        return jsonify({"status": "error", "error": f"Unknown module: {module}"}), 400

    result = pool.run_command(connection, module, command, caller_id)
    is_error = result.get("status") == "error"
    is_auth = result.get("status") == "auth_required"
    metrics.record_request(time.time() - start, error=is_error)
    if is_auth:
        metrics.record_auth()
    return jsonify(result)


@app.route("/reset", methods=["POST"])
def reset_connection():
    """Reset (kill + remove) sessions for a specific connection."""
    data = request.get_json()
    connection = data.get("connection")
    module = data.get("module")  # Optional: reset only one module

    if not connection:
        return jsonify({"status": "error", "error": "Missing connection parameter"}), 400

    result = pool.reset_connection(connection, module)
    logger.info(f"Reset connection {connection} (module={module}): {result}")
    return jsonify({"status": "success", **result})


@app.route("/metrics", methods=["GET"])
def get_metrics():
    stats = metrics.get_stats()
    stats["sessions"] = pool.get_status()["sessions"]
    stats["active_sessions"] = len([s for s in stats["sessions"] if s["authenticated"]])
    stats["keepalive"] = keepalive.get_stats()
    return jsonify(stats)


if __name__ == "__main__":
    logger.info(f"Starting M365 Session Pool on {HOST}:{PORT}")
    app.run(host=HOST, port=PORT, threaded=True)
