#!/usr/bin/env python3
"""
M365 Registry MCP Server.

Central registry for M365 connections. This is the ONLY MCP that writes to
~/.m365-connections.json. Other MCPs read from it.

Tools:
- registry_add_connection: Add a new connection (requires appId)
- registry_remove_connection: Remove a connection by name
- registry_update_connection: Update an existing connection
- registry_list_connections: List all connections
- registry_setup_instructions: Show how to create an Azure AD app
"""

import asyncio
import fcntl
import json
import re
import sys
import time
from pathlib import Path

# Add parent dir for shared logger
sys.path.insert(0, str(Path(__file__).parent.parent))
try:
    from mcp_logger import log_tool_call
except ImportError:
    def log_tool_call(*args, **kwargs): pass

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

# Registry file location
CONNECTIONS_FILE = Path.home() / ".m365-connections.json"

# Valid MCP names
VALID_MCPS = ["pnp-m365", "microsoft-graph", "exo", "pwsh-manager", "pp-admin", "onenote"]

server = Server("m365-registry")


def load_registry() -> dict:
    """Load the registry file."""
    try:
        return json.loads(CONNECTIONS_FILE.read_text())
    except FileNotFoundError:
        return {
            "_schema": "Universal M365 connection registry - shared across MCPs",
            "_rules": [
                "connectionName is ALWAYS required - NO DEFAULTS EVER",
                "Each connection = tenant + appId + description + mcps array",
                "mcps array controls which MCP servers can use this connection",
                "description is REQUIRED - must explain what this app/account combination is for"
            ],
            "connections": {}
        }
    except Exception as e:
        return {"error": str(e), "connections": {}}


def save_registry(data: dict) -> bool:
    """Save the registry file with file locking."""
    try:
        with open(CONNECTIONS_FILE, 'w') as f:
            fcntl.flock(f.fileno(), fcntl.LOCK_EX)
            try:
                json.dump(data, f, indent=2)
                f.write('\n')
            finally:
                fcntl.flock(f.fileno(), fcntl.LOCK_UN)
        return True
    except Exception as e:
        return False


def is_valid_guid(s: str) -> bool:
    """Check if string is a valid GUID format."""
    pattern = r'^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
    return bool(re.match(pattern, s))


def get_setup_instructions(tenant: str = None) -> str:
    """Return instructions for setting up an Azure AD app."""
    tenant_part = f" --tenant {tenant}" if tenant else ""
    return f"""## Create Azure AD App Registration

**The PnP multi-tenant app was retired Sept 9, 2024. You must create your own app.**

### Option 1: CLI for Microsoft 365 (Recommended)

```bash
# Install CLI for Microsoft 365
npm install -g @pnp/cli-microsoft365

# Create app registration (requires Global Admin)
m365 setup{tenant_part}
```

After setup completes, note the **App ID** (Client ID) - you'll need it.

### Option 2: Manual Azure Portal

1. Go to https://portal.azure.com
2. Azure Active Directory > App Registrations > New Registration
3. Name: "My M365 App"
4. Supported account types: Single tenant
5. Note the Application (client) ID

### After Creating the App

Add the connection with:
```
registry_add_connection(
    name="YourConnectionName",
    tenant="{tenant or 'yourtenant.onmicrosoft.com'}",
    appId="your-app-id-here",
    description="What this connection is for",
    mcps=["pnp-m365", "exo"]  # which MCPs need access
)
```

See docs/M365-CLI-SETUP.md for detailed instructions.
"""


@server.list_tools()
async def list_tools():
    return [
        Tool(
            name="registry_add_connection",
            description="Add a new M365 connection to the registry. Requires appId - use registry_setup_instructions if you need to create one.",
            inputSchema={
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Connection name (e.g., 'ForIT', 'ClientX')"},
                    "tenant": {"type": "string", "description": "Tenant domain (e.g., 'forit.io', 'clientx.onmicrosoft.com')"},
                    "appId": {"type": "string", "description": "Azure AD app registration ID (GUID). REQUIRED - no default exists."},
                    "description": {"type": "string", "description": "REQUIRED: What this app/account combination is used for"},
                    "mcps": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": f"Which MCPs can use this connection. Valid: {VALID_MCPS}"
                    }
                },
                "required": ["name", "tenant", "appId", "description", "mcps"]
            }
        ),
        Tool(
            name="registry_remove_connection",
            description="Remove a connection from the registry by name.",
            inputSchema={
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Connection name to remove"}
                },
                "required": ["name"]
            }
        ),
        Tool(
            name="registry_update_connection",
            description="Update an existing connection's properties.",
            inputSchema={
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Connection name to update"},
                    "appId": {"type": "string", "description": "New app ID (optional)"},
                    "tenant": {"type": "string", "description": "New tenant (optional)"},
                    "description": {"type": "string", "description": "New description (optional)"},
                    "mcps": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "New MCP list (optional)"
                    }
                },
                "required": ["name"]
            }
        ),
        Tool(
            name="registry_list_connections",
            description="List all connections in the registry, optionally filtered by MCP.",
            inputSchema={
                "type": "object",
                "properties": {
                    "mcp": {"type": "string", "description": f"Filter to connections for this MCP. Valid: {VALID_MCPS}"}
                },
                "required": []
            }
        ),
        Tool(
            name="registry_setup_instructions",
            description="Show instructions for creating an Azure AD app registration using CLI for Microsoft 365.",
            inputSchema={
                "type": "object",
                "properties": {
                    "tenant": {"type": "string", "description": "Tenant domain to include in instructions (optional)"}
                },
                "required": []
            }
        ),
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict):
    start_time = time.time()
    error_msg = None
    result_summary = None

    try:
        result = await _call_tool_impl(name, arguments)
        if result and len(result) > 0:
            text = result[0].text[:100] if hasattr(result[0], 'text') else str(result[0])[:100]
            if "error" in text.lower():
                error_msg = text
            else:
                result_summary = text
        return result
    except Exception as e:
        error_msg = str(e)
        raise
    finally:
        duration_ms = int((time.time() - start_time) * 1000)
        log_tool_call(
            mcp_name="m365-registry",
            tool_name=name,
            arguments=arguments,
            connection_name=arguments.get("name"),
            result=result_summary,
            error=error_msg,
            duration_ms=duration_ms
        )


async def _call_tool_impl(name: str, arguments: dict):
    """Tool implementation."""

    if name == "registry_list_connections":
        registry = load_registry()
        connections = registry.get("connections", {})
        mcp_filter = arguments.get("mcp")

        if mcp_filter:
            connections = {
                n: c for n, c in connections.items()
                if mcp_filter in c.get("mcps", [])
            }

        if not connections:
            msg = f"No connections found"
            if mcp_filter:
                msg += f" for MCP '{mcp_filter}'"
            return [TextContent(type="text", text=json.dumps({"message": msg, "connections": []}, indent=2))]

        results = []
        for conn_name, conn in connections.items():
            results.append({
                "name": conn_name,
                "tenant": conn.get("tenant", ""),
                "appId": conn.get("appId", ""),
                "description": conn.get("description", ""),
                "mcps": conn.get("mcps", [])
            })

        return [TextContent(type="text", text=json.dumps({"connections": results}, indent=2))]

    elif name == "registry_add_connection":
        conn_name = arguments.get("name", "").strip()
        tenant = arguments.get("tenant", "").strip()
        app_id = arguments.get("appId", "").strip()
        description = arguments.get("description", "").strip()
        mcps = arguments.get("mcps", [])

        # Validate required fields
        missing = []
        if not conn_name:
            missing.append("name")
        if not tenant:
            missing.append("tenant")
        if not app_id:
            missing.append("appId")
        if not description:
            missing.append("description")
        if not mcps:
            missing.append("mcps")

        if missing:
            error = {"error": f"Missing required fields: {', '.join(missing)}"}
            if "appId" in missing:
                error["hint"] = "No default appId exists. Use registry_setup_instructions to learn how to create one."
            return [TextContent(type="text", text=json.dumps(error, indent=2))]

        # Validate appId format
        if not is_valid_guid(app_id):
            return [TextContent(type="text", text=json.dumps({
                "error": f"Invalid appId format: {app_id}",
                "hint": "appId must be a valid GUID (e.g., '9bc3ab49-b65d-410a-85ad-de819febfddc')"
            }, indent=2))]

        # Validate MCPs
        invalid_mcps = [m for m in mcps if m not in VALID_MCPS]
        if invalid_mcps:
            return [TextContent(type="text", text=json.dumps({
                "error": f"Invalid MCP names: {invalid_mcps}",
                "valid_mcps": VALID_MCPS
            }, indent=2))]

        # Load and update registry
        registry = load_registry()
        if conn_name in registry.get("connections", {}):
            return [TextContent(type="text", text=json.dumps({
                "error": f"Connection '{conn_name}' already exists",
                "hint": "Use registry_update_connection to modify, or registry_remove_connection first"
            }, indent=2))]

        registry.setdefault("connections", {})[conn_name] = {
            "appId": app_id,
            "tenant": tenant,
            "description": description,
            "mcps": mcps
        }

        if save_registry(registry):
            return [TextContent(type="text", text=json.dumps({
                "success": True,
                "message": f"Connection '{conn_name}' added",
                "connection": registry["connections"][conn_name]
            }, indent=2))]
        else:
            return [TextContent(type="text", text=json.dumps({
                "error": "Failed to save registry file"
            }, indent=2))]

    elif name == "registry_remove_connection":
        conn_name = arguments.get("name", "").strip()
        if not conn_name:
            return [TextContent(type="text", text=json.dumps({"error": "name is required"}, indent=2))]

        registry = load_registry()
        if conn_name not in registry.get("connections", {}):
            return [TextContent(type="text", text=json.dumps({
                "error": f"Connection '{conn_name}' not found",
                "available": list(registry.get("connections", {}).keys())
            }, indent=2))]

        removed = registry["connections"].pop(conn_name)
        if save_registry(registry):
            return [TextContent(type="text", text=json.dumps({
                "success": True,
                "message": f"Connection '{conn_name}' removed",
                "removed": removed
            }, indent=2))]
        else:
            return [TextContent(type="text", text=json.dumps({
                "error": "Failed to save registry file"
            }, indent=2))]

    elif name == "registry_update_connection":
        conn_name = arguments.get("name", "").strip()
        if not conn_name:
            return [TextContent(type="text", text=json.dumps({"error": "name is required"}, indent=2))]

        registry = load_registry()
        if conn_name not in registry.get("connections", {}):
            return [TextContent(type="text", text=json.dumps({
                "error": f"Connection '{conn_name}' not found",
                "available": list(registry.get("connections", {}).keys())
            }, indent=2))]

        conn = registry["connections"][conn_name]
        updated = []

        if "appId" in arguments and arguments["appId"]:
            app_id = arguments["appId"].strip()
            if not is_valid_guid(app_id):
                return [TextContent(type="text", text=json.dumps({
                    "error": f"Invalid appId format: {app_id}"
                }, indent=2))]
            conn["appId"] = app_id
            updated.append("appId")

        if "tenant" in arguments and arguments["tenant"]:
            conn["tenant"] = arguments["tenant"].strip()
            updated.append("tenant")

        if "description" in arguments and arguments["description"]:
            conn["description"] = arguments["description"].strip()
            updated.append("description")

        if "mcps" in arguments and arguments["mcps"]:
            invalid_mcps = [m for m in arguments["mcps"] if m not in VALID_MCPS]
            if invalid_mcps:
                return [TextContent(type="text", text=json.dumps({
                    "error": f"Invalid MCP names: {invalid_mcps}",
                    "valid_mcps": VALID_MCPS
                }, indent=2))]
            conn["mcps"] = arguments["mcps"]
            updated.append("mcps")

        if not updated:
            return [TextContent(type="text", text=json.dumps({
                "message": "No changes made - no update fields provided"
            }, indent=2))]

        if save_registry(registry):
            return [TextContent(type="text", text=json.dumps({
                "success": True,
                "message": f"Connection '{conn_name}' updated",
                "updated_fields": updated,
                "connection": conn
            }, indent=2))]
        else:
            return [TextContent(type="text", text=json.dumps({
                "error": "Failed to save registry file"
            }, indent=2))]

    elif name == "registry_setup_instructions":
        tenant = arguments.get("tenant", "")
        return [TextContent(type="text", text=get_setup_instructions(tenant))]

    return [TextContent(type="text", text=f"Unknown tool: {name}")]


async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())


if __name__ == "__main__":
    asyncio.run(main())
