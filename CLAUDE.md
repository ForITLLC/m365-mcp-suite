# Microsoft Master (MM)

## Architecture
One MCP server — `mm` — for ALL Microsoft operations.
- **mm** - Microsoft Master (Exchange, SharePoint, Teams, Azure, Power Platform, Graph API)
- Two tools: `mm__run` (PowerShell) and `mm__graph_request` (Graph REST API)
- PowerShell backed by `session-pool/` — isolated Docker containers per connection
- Graph backed by MSAL with per-connection token cache (`~/.mm-graph-tokens/`)
- Connection registry: `~/.m365-connections.json` (READ-ONLY to MCPs)

Old servers (pnp, graph, pwsh-manager, registry) are archived in `_archived/`.

## Connection Rules - NO DEFAULTS EVER
- Universal registry: `~/.m365-connections.json`
- Every command REQUIRES `connection` parameter
- NEVER use the word "default" in any MM code
- Connection = Account + AppId + Tenant + Description
- Connections: ForIT-GA, ForIT-Personal, Pivot, GreatNorth-GA, GreatNorth-Personal, WMA

## Authentication
- ALWAYS use MM MCP tools, NEVER Bash for M365 auth
- Both tools handle device codes and auth automatically
- Display device codes as:
  ```
  **DEVICE CODE: XXXXXXXX**
  Go to: https://microsoft.com/devicelogin
  ```

## MM Tools

### `mm__run` — PowerShell via Session Pool
- No params → list connections
- `connection` + `module` + `command` → execute PowerShell
- Modules: `exo` (Exchange), `pnp` (SharePoint), `azure`, `teams`

### `mm__graph_request` — Microsoft Graph REST API
- No params → list connections
- `connection` + `endpoint` → GET request to Graph
- `connection` + `endpoint` + `method` + `body` → any Graph operation
- Supports `/v1.0/` and `/beta/` endpoints
- OData noise stripped automatically
- Token cache: `~/.mm-graph-tokens/` (per connection, MSAL)

## Session Pool (Docker)
- Each connection gets an isolated Docker container
- Router on port 5200, individual containers on 5210-5215
- `session-pool/docker-compose.yml` manages the stack
- Health checks built in

## Testing
```bash
# Check all containers healthy
docker ps --format "{{.Names}}\t{{.Status}}" | grep m365

# Test PowerShell
mcpjungle invoke mm run '{"connection":"ForIT-GA","module":"exo","command":"Get-Mailbox -ResultSize 1"}'

# Test Graph API
mcpjungle invoke mm graph_request '{"connection":"ForIT-GA","endpoint":"/me"}'
```

## MCPJungle Registration
```bash
mcpjungle register --conf mm/mcpjungle-config.json
```
