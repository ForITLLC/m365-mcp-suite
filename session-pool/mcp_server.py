#!/usr/bin/env python3
"""
M365 Session Pool MCP Server

Thin MCP wrapper around the session pool HTTP API.
The session pool does all the heavy lifting - this just exposes it via MCP.
"""

import hashlib
import json
import logging
import os
import sys
import urllib.request
import urllib.error
from datetime import datetime

# Logging to stderr (MCP uses stdout for protocol)
LOG_FILE = "/tmp/mm_mcp.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler(sys.stderr)
    ]
)
logger = logging.getLogger("mm_mcp")

# Session pool endpoint
SESSION_POOL_URL = os.getenv("SESSION_POOL_URL", "http://localhost:5200")

# Generate caller_id from machine + user for auth coordination
def get_caller_id() -> str:
    """Generate stable caller ID for auth coordination."""
    identifier = f"{os.getenv('USER', 'unknown')}@{os.uname().nodename}"
    return hashlib.sha256(identifier.encode()).hexdigest()[:12]

CALLER_ID = get_caller_id()
logger.info(f"MCP Server starting. Caller ID: {CALLER_ID}, Pool URL: {SESSION_POOL_URL}")


def call_pool_api(endpoint: str, method: str = "GET", data: dict = None) -> dict:
    """Call the session pool HTTP API."""
    url = f"{SESSION_POOL_URL}{endpoint}"
    logger.info(f"API call: {method} {url}")
    if data:
        logger.info(f"Request data: {json.dumps(data)}")

    try:
        if data:
            req = urllib.request.Request(
                url,
                data=json.dumps(data).encode(),
                headers={"Content-Type": "application/json"},
                method=method
            )
        else:
            req = urllib.request.Request(url, method=method)

        logger.info(f"Opening connection to {url}...")
        resp = urllib.request.urlopen(req, timeout=30)
        logger.info(f"Connection opened, status: {resp.status}")
        raw = resp.read()
        logger.info(f"Read {len(raw)} bytes")
        result = json.loads(raw.decode())
        logger.info(f"API response: {json.dumps(result)[:500]}")
        resp.close()
        return result

    except urllib.error.URLError as e:
        logger.error(f"API URLError: {e.reason}")
        return {"status": "error", "error": f"Session pool unavailable: {e.reason}"}
    except urllib.error.HTTPError as e:
        logger.error(f"API HTTPError: {e.code} {e.reason}")
        try:
            body = json.loads(e.read().decode())
            return body
        except:
            return {"status": "error", "error": f"HTTP {e.code}: {e.reason}"}
    except Exception as e:
        logger.error(f"API Exception: {e}", exc_info=True)
        return {"status": "error", "error": str(e)}


def format_auth_required(result: dict) -> str:
    """Format auth_required response with prominent device code."""
    device_code = result.get("device_code", "")
    auth_url = result.get("auth_url", "https://microsoft.com/devicelogin")
    message = result.get("message", "")

    logger.info(f"Auth required - device_code: {device_code}")

    if not device_code:
        return f"""
**AUTH REQUIRED** (device code not captured)

Go to: {auth_url}

{message}

Check container logs: docker logs m365-pool
"""

    return f"""
**DEVICE CODE: {device_code}**
Go to: {auth_url}

{message}

After completing authentication, retry your command.
"""


def format_result(result: dict) -> str:
    """Format API result for display."""
    status = result.get("status", "unknown")
    logger.info(f"Formatting result with status: {status}")

    if status == "auth_required":
        return format_auth_required(result)

    if status == "auth_in_progress":
        return f"Auth in progress by another session. {result.get('message', 'Retry shortly.')}"

    if status == "success":
        output = result.get("output", "OK")
        logger.info(f"Command success, output length: {len(output)}")
        return output

    if status == "error":
        error = result.get('error', 'Unknown error')
        logger.error(f"Command error: {error}")
        return f"Error: {error}"

    # Fallback - return as JSON
    return json.dumps(result, indent=2)


# MCP Protocol Implementation
def handle_initialize(params: dict) -> dict:
    """Handle MCP initialize request."""
    logger.info("MCP initialize")
    return {
        "protocolVersion": "2024-11-05",
        "capabilities": {"tools": {}},
        "serverInfo": {
            "name": "mm",
            "version": "1.0.0",
        },
    }


def handle_list_tools() -> dict:
    """Return available tools."""
    logger.info("MCP tools/list")
    return {
        "tools": [
            {
                "name": "run",
                "description": "Microsoft PowerShell. Omit all params to list connections. Provide connection+module+command to execute.",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "connection": {
                            "type": "string",
                            "description": "Connection name (e.g., 'ForIT-GA')",
                        },
                        "module": {
                            "type": "string",
                            "enum": ["exo", "pnp", "azure", "teams", "powerplatform"],
                            "description": "exo=Exchange, pnp=SharePoint, azure, teams, powerplatform",
                        },
                        "command": {
                            "type": "string",
                            "description": "PowerShell command",
                        },
                    },
                },
            },
        ]
    }


def handle_call_tool(name: str, arguments: dict) -> dict:
    """Execute a tool call."""
    logger.info(f"MCP tools/call: {name} with args: {json.dumps(arguments)}")

    if name == "run":
        connection = arguments.get("connection", "").strip()
        module = arguments.get("module", "").strip()
        command = arguments.get("command", "").strip()

        # No params = list connections + status
        if not connection and not module and not command:
            logger.info("Listing connections (no params)")
            conns = call_pool_api("/connections", "GET").get("connections", {})
            status = call_pool_api("/status", "GET").get("sessions", [])

            lines = ["Connections:"]
            for n, info in conns.items():
                lines.append(f"  {n} ({info.get('tenant', '?')}) - {info.get('description', '')}")

            if status:
                lines.append("\nActive sessions:")
                for s in status:
                    state = "[OK]" if s.get("authenticated") else "[PENDING]"
                    lines.append(f"  {state} {s.get('session_id')}")

            lines.append("\nModules: exo, pnp, azure, teams, powerplatform")
            return {"content": [{"type": "text", "text": "\n".join(lines)}]}

        # Partial params = error with guidance
        if not all([connection, module, command]):
            missing = [p for p, v in [("connection", connection), ("module", module), ("command", command)] if not v]
            logger.warning(f"Missing params: {missing}")
            return {"content": [{"type": "text", "text": f"Missing: {', '.join(missing)}. Omit all params to list connections."}], "isError": True}

        # Execute command
        logger.info(f"Executing: {connection}/{module} -> {command[:100]}")
        result = call_pool_api("/run", "POST", {
            "connection": connection,
            "module": module,
            "command": command,
            "caller_id": CALLER_ID,
        })

        formatted = format_result(result)
        is_error = result.get("status") == "error"
        logger.info(f"Result formatted, is_error: {is_error}")
        return {"content": [{"type": "text", "text": formatted}], "isError": is_error}

    logger.error(f"Unknown tool: {name}")
    return {"content": [{"type": "text", "text": f"Unknown tool: {name}"}], "isError": True}


def main():
    """Main MCP server loop."""
    logger.info("MCP server main loop starting")
    request_id = None

    while True:
        try:
            line = sys.stdin.readline()
            if not line:
                logger.info("EOF on stdin, exiting")
                break

            logger.debug(f"Received: {line[:200]}")
            request = json.loads(line)
            method = request.get("method", "")
            params = request.get("params", {})
            request_id = request.get("id")

            logger.info(f"MCP method: {method}, id: {request_id}")
            result = None

            if method == "initialize":
                result = handle_initialize(params)
            elif method == "notifications/initialized":
                logger.info("Client initialized notification")
                continue  # No response needed
            elif method == "tools/list":
                result = handle_list_tools()
            elif method == "tools/call":
                tool_name = params.get("name", "")
                arguments = params.get("arguments", {})
                result = handle_call_tool(tool_name, arguments)
            else:
                logger.warning(f"Unknown MCP method: {method}")
                result = {"error": {"code": -32601, "message": f"Unknown method: {method}"}}

            if request_id is not None:
                response = {"jsonrpc": "2.0", "id": request_id, "result": result}
                response_str = json.dumps(response)
                logger.debug(f"Sending: {response_str[:200]}")
                sys.stdout.write(response_str + "\n")
                sys.stdout.flush()

        except json.JSONDecodeError as e:
            logger.error(f"JSON decode error: {e}")
            continue
        except Exception as e:
            logger.error(f"Exception in main loop: {e}", exc_info=True)
            if request_id is not None:
                error_response = {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "error": {"code": -32603, "message": str(e)},
                }
                sys.stdout.write(json.dumps(error_response) + "\n")
                sys.stdout.flush()


if __name__ == "__main__":
    main()
