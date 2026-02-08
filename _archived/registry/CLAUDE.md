# M365 Registry MCP

Central registry for M365 connections. This MCP is the ONLY writer to `~/.m365-connections.json`.

## Purpose

- Add new M365 connections (tenant + appId + description)
- Remove connections
- Update existing connections
- List connections (optionally filtered by MCP)
- Show setup instructions for creating Azure AD apps

## Key Rule

**appId is REQUIRED.** There is no default multi-tenant app - the PnP Management Shell app was retired Sept 9, 2024. Every connection needs a custom app registration.

## Tools

| Tool | Purpose |
|------|---------|
| `registry_add_connection` | Add new connection (all fields required) |
| `registry_remove_connection` | Remove by name |
| `registry_update_connection` | Update existing connection |
| `registry_list_connections` | List all, optionally filter by MCP |
| `registry_setup_instructions` | Show how to create an Azure AD app |

## Adding a New Connection

1. Create an Azure AD app first (see `registry_setup_instructions`)
2. Call `registry_add_connection` with:
   - `name`: Friendly identifier (e.g., "ClientX")
   - `tenant`: Domain (e.g., "clientx.onmicrosoft.com")
   - `appId`: The app registration ID (GUID)
   - `description`: REQUIRED - what this connection is for
   - `mcps`: Array of MCPs that can use it

## Valid MCPs

- `pnp-m365` - PnP CLI for Microsoft 365
- `microsoft-graph` - Microsoft Graph API
- `exo` - Exchange Online
- `pwsh-manager` - PowerShell session manager
- `pp-admin` - Power Platform Admin
- `onenote` - OneNote API

## Registry Location

`~/.m365-connections.json`

Other MCPs read this file directly. Only this MCP writes to it.
