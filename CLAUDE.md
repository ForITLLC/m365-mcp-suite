# Microsoft Master (MM)

## Architecture
One MCP server — `mm` — for ALL Microsoft operations.
- **mm** - Microsoft Master (Exchange, SharePoint, Teams, Azure, Power Platform, Graph API)
- Two tools: `mm__run` (PowerShell) and `mm__graph_request` (Graph REST API)
- PowerShell backed by `session-pool/` — Docker containers with persistent sessions
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
- Both tools auto-trigger device codes when a connection isn't authenticated — no pre-auth step needed
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
- Command guardrails block dangerous operations (see below)

### `mm__graph_request` — Microsoft Graph REST API
- No params → list connections
- `connection` + `endpoint` → GET request to Graph
- `connection` + `endpoint` + `method` + `body` → any Graph operation
- Optional `resource`: `graph` (default) or `flow` (Power Automate)
- Supports `/v1.0/` and `/beta/` endpoints
- OData noise stripped automatically
- Token cache: `~/.mm-graph-tokens/` (per connection, MSAL)

## Session Pool

### Deployment Modes
Both modes use the same `session_pool.py` — only `SINGLE_CONNECTION` env var differs.

| Mode | File | Use Case |
|------|------|----------|
| **Unified** | `docker-compose.unified.yml` | Dev, single container, all connections |
| **Isolated** | `docker-compose.isolated.yml` | Production, one container per connection |

### Command Guardrails
The session pool blocks dangerous operations before execution:

**Blocked** (hard reject):
- `Install-Module`, `Uninstall-Module`, `Update-Module` — container integrity
- `New-AzRoleAssignment`, `Remove-AzRoleAssignment` — access escalation
- `Set-AzKeyVaultAccessPolicy`, `Remove-AzKeyVaultAccessPolicy` — vault security
- `New-AzADApplication`, `New-AzADServicePrincipal` — app registration
- Raw OAuth requests to `login.microsoftonline` — auth bypass

**Warned** (logged but allowed):
- `Remove-Az*` — Azure resource deletion
- `Remove-Mailbox`, `Remove-PnP*`, `Remove-Team` — destructive operations
- `ForwardingSmtpAddress`, `Set-InboxRule` with Forward — mail forwarding

### Session Persistence
- Authenticated session metadata saved to `~/.m365-state/` on every auth completion
- On container restart, sessions restore automatically:
  - **Azure**: Tokens cached in `~/.Azure` Docker volume — restores without device code
  - **EXO/Teams**: In-process only — requires re-auth on next use (by design)
  - **PnP**: Token cache in `~/.config` Docker volume — may restore depending on token expiry

### Azure Context Isolation (Unified Mode)
In unified mode, all Azure sessions share `~/.Azure`. To prevent cross-tenant contamination:
- `Disable-AzContextAutosave -Scope Process` prevents writes to shared cache
- `Select-AzContext` by `expectedEmail` from connection config picks the correct tenant

### Logging
Dual output: stdout (`docker logs`) + persistent files (`~/.m365-logs/`).
- Every command's output content logged (first 500 chars preview)
- AADSTS and auth error patterns flagged at WARNING level
- `mm/server.py` also scans "successful" responses for buried auth errors

### Host Directories
| Path | Purpose |
|------|---------|
| `~/.m365-connections.json` | Connection registry (read-only) |
| `~/.m365-logs/` | Persistent log files |
| `~/.m365-state/` | Session state for restart persistence |
| `~/.mm-graph-tokens/` | MSAL token cache for Graph API |

## Testing
```bash
# Check all containers healthy
docker ps --format "{{.Names}}\t{{.Status}}" | grep m365

# Test PowerShell
mcpjungle invoke mm run '{"connection":"ForIT-GA","module":"exo","command":"Get-Mailbox -ResultSize 1"}'

# Test Graph API
mcpjungle invoke mm graph_request '{"connection":"ForIT-GA","endpoint":"/me"}'

# Session status
curl http://localhost:5200/status | jq

# Metrics
curl http://localhost:5200/metrics | jq

# Persistent logs
tail -f ~/.m365-logs/session-pool-unified.log
```

## MCPJungle Registration
```bash
mcpjungle register --conf mm/mcpjungle-config.json
```
