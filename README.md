# Microsoft Master (MM)

Unified Microsoft 365 MCP server for AI assistants. One server, two tools, all of M365.

## History

This repo originally contained four separate MCP servers:

| Server | Purpose | Status |
|--------|---------|--------|
| **graph** | Microsoft Graph REST API (TypeScript) | Archived |
| **pnp** | CLI for Microsoft 365 / SharePoint (TypeScript) | Archived |
| **pwsh-manager** | PowerShell session management (Python/Docker) | Archived |
| **registry** | Connection registry management (Python) | Archived |

In February 2026, these were consolidated into a single **mm** server with two tools:
- `mm__run` — PowerShell commands via a Docker session pool
- `mm__graph_request` — Direct Microsoft Graph REST API calls via MSAL

The old servers are preserved in `_archived/` for reference.

## Architecture

```
┌─────────────────────────────────────────────────────────┐
│                    AI Assistant                          │
└─────────────────────┬───────────────────────────────────┘
                      │ MCP Protocol (stdio)
                      ▼
               ┌─────────────┐
               │  mm/server  │  Python MCP server
               │             │  - graph_request (MSAL → Graph API)
               │             │  - run (HTTP → session pool)
               └──────┬──────┘
                      │ HTTP :5200
                      ▼
          ┌───────────────────────┐
          │   session-pool/       │  Docker container(s)
          │   session_pool.py     │  - PowerShell processes per module
          │                       │  - Native device code auth
          │   Modules:            │  - Session persistence
          │   exo, pnp, azure,    │  - Command guardrails
          │   teams               │  - Comprehensive logging
          └───────────────────────┘
                      │
          ┌───────────┴───────────┐
          │ ~/.m365-connections   │  Connection registry (READ-ONLY)
          │ ~/.mm-graph-tokens/   │  Graph MSAL token cache
          │ ~/.m365-logs/         │  Persistent logs
          │ ~/.m365-state/        │  Session state persistence
          └───────────────────────┘
```

## Quick Start

### 1. Start the session pool

```bash
# Unified mode — single container, all connections (recommended for dev/small servers)
cd session-pool
docker compose -p m365-session-pool -f docker-compose.unified.yml up -d

# Isolated mode — one container per connection (recommended for production/32GB+ RAM)
docker compose -p m365-session-pool -f docker-compose.isolated.yml up -d
```

### 2. Register with MCPJungle

```bash
mcpjungle register --conf mm/mcpjungle-config.json
```

### 3. Use it

```bash
# List connections
mcpjungle invoke mm run '{}'

# PowerShell (Exchange)
mcpjungle invoke mm run '{"connection":"ForIT-GA","module":"exo","command":"Get-Mailbox -ResultSize 1"}'

# Graph API
mcpjungle invoke mm graph_request '{"connection":"ForIT-GA","endpoint":"/me"}'

# Power Automate (Flow API)
mcpjungle invoke mm graph_request '{"connection":"ForIT-GA","endpoint":"/providers/Microsoft.ProcessSimple/environments","resource":"flow"}'
```

Auth is automatic — if a connection isn't authenticated, the tool returns a device code. No pre-auth step needed.

## Connection Registry

All connections live in `~/.m365-connections.json` (read-only to MCPs):

```json
{
  "connections": {
    "ForIT-GA": {
      "appId": "your-app-id",
      "tenant": "forit.io",
      "tenantId": "guid-here",
      "expectedEmail": "user@domain.com",
      "description": "ForIT Global Admin"
    }
  }
}
```

Every command requires a `connection` parameter. There are no defaults.

## Tools

### `mm__run` — PowerShell via Session Pool

Execute PowerShell commands through persistent Docker-hosted sessions.

| Parameter | Description |
|-----------|-------------|
| `connection` | Connection name (e.g., `ForIT-GA`) |
| `module` | `exo` (Exchange), `pnp` (SharePoint), `azure`, `teams` |
| `command` | PowerShell command to execute |

Omit all parameters to list available connections.

### `mm__graph_request` — Microsoft Graph REST API

Direct HTTP requests to Microsoft Graph (or Flow API) via MSAL tokens.

| Parameter | Description |
|-----------|-------------|
| `connection` | Connection name |
| `endpoint` | API path (e.g., `/me/messages`) |
| `method` | `GET`, `POST`, `PATCH`, `PUT`, `DELETE` (default: GET) |
| `body` | Request body for POST/PATCH/PUT |
| `resource` | `graph` (default) or `flow` for Power Automate |

## Session Pool

The session pool manages PowerShell processes with native device code authentication.

### Deployment Modes

| Mode | File | Use Case | RAM |
|------|------|----------|-----|
| **Unified** | `docker-compose.unified.yml` | Dev, small servers | ~1-2GB total |
| **Isolated** | `docker-compose.isolated.yml` | Production, multi-user | ~512MB per connection |

### Features

- **Session persistence** — Authenticated sessions survive container restarts. Session metadata saved to `~/.m365-state/`, PowerShell token caches persisted via Docker volumes. Azure sessions restore from cached tokens; EXO/Teams require re-auth (in-process only).
- **Command guardrails** — Blocks dangerous operations: `Install-Module` (container integrity), `New-AzRoleAssignment` (access escalation), raw OAuth requests, app registration modifications. Warned but allowed: `Remove-Az*`, mail forwarding rules.
- **Azure context isolation** — In unified mode, `Disable-AzContextAutosave` + `Select-AzContext` by expectedEmail prevents cross-tenant context contamination when multiple Azure sessions share `~/.Azure`.
- **Comprehensive logging** — Dual output: stdout (docker logs) + persistent files (`~/.m365-logs/`). Every command's output content logged. AADSTS and auth error patterns flagged at WARNING level.
- **Keepalive** — Background thread pings authenticated sessions every 5 minutes to prevent token expiry. Stale sessions (auth_pending > 15 min) automatically reaped.
- **Metrics** — `/metrics` endpoint with request counts, error rates, response times, session states.

### Session Pool API

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/health` | GET | Health check |
| `/status` | GET | All session states |
| `/connections` | GET | List registry connections |
| `/run` | POST | Execute a command |
| `/reset` | POST | Reset a connection's sessions |
| `/metrics` | GET | Performance metrics |

### Host Directories

| Path | Purpose |
|------|---------|
| `~/.m365-connections.json` | Connection registry (read-only) |
| `~/.m365-logs/` | Persistent log files |
| `~/.m365-state/` | Session state for restart persistence |
| `~/.mm-graph-tokens/` | MSAL token cache for Graph API |

## Monitoring

```bash
# Container health
docker ps --format "{{.Names}}\t{{.Status}}" | grep m365

# Live logs
docker logs -f m365-pool

# Persistent logs (survive container restarts)
tail -f ~/.m365-logs/session-pool-unified.log

# Session status
curl http://localhost:5200/status | jq

# Metrics
curl http://localhost:5200/metrics | jq
```

## App Registration

The PnP multi-tenant app was retired September 9, 2024. You must create your own Azure AD app registration. See [docs/M365-CLI-SETUP.md](docs/M365-CLI-SETUP.md) for instructions.

## License

MIT
