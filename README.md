# ForIT Microsoft Graph

A lean MCP (Model Context Protocol) server for raw Microsoft Graph API access with **multi-tenant account management**.

> Fork of [@softeria/ms-365-mcp-server](https://github.com/Softeria/ms-365-mcp-server) - stripped down to essentials.

## Why This Fork?

The original Softeria MCP exposes **37 specialized tools** (~12KB context). This fork exposes **7 tools** (~1KB context):

| Tool | Description |
|------|-------------|
| `login` | Authenticate with Microsoft (device code flow) |
| `logout` | Log out from Microsoft account |
| `verify-login` | Check authentication status |
| `list-accounts` | List all cached Microsoft accounts |
| `select-account` | Switch between accounts (multi-tenant) |
| `remove-account` | Remove an account from cache |
| `graph-request` | **Execute any Graph API request** |

## The `graph-request` Tool

Instead of 30+ specialized tools for mail, calendar, users, etc., use one flexible tool:

```json
{
  "endpoint": "/me/messages",
  "method": "GET",
  "queryParams": {
    "$select": "subject,from,receivedDateTime",
    "$top": "10"
  }
}
```

### Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `endpoint` | string | Graph API path (e.g., `/me`, `/users`, `/me/calendar/events`) |
| `method` | enum | `GET`, `POST`, `PUT`, `PATCH`, `DELETE` |
| `body` | object | Request body for POST/PUT/PATCH |
| `queryParams` | object | OData query params (`$select`, `$filter`, `$top`, etc.) |
| `headers` | object | Additional HTTP headers |
| `apiVersion` | enum | `v1.0` (default) or `beta` |
| `accountId` | string | Target specific account without switching (multi-tenant) |

### Examples

**Get current user:**
```json
{ "endpoint": "/me" }
```

**List emails with filter:**
```json
{
  "endpoint": "/me/messages",
  "queryParams": {
    "$filter": "isRead eq false",
    "$select": "subject,from",
    "$top": "5"
  }
}
```

**Create calendar event:**
```json
{
  "endpoint": "/me/calendar/events",
  "method": "POST",
  "body": {
    "subject": "Team Meeting",
    "start": { "dateTime": "2025-01-15T10:00:00", "timeZone": "UTC" },
    "end": { "dateTime": "2025-01-15T11:00:00", "timeZone": "UTC" }
  }
}
```

**Use beta API:**
```json
{
  "endpoint": "/me/insights/used",
  "apiVersion": "beta"
}
```

**Query specific tenant (multi-account):**
```json
{
  "endpoint": "/me",
  "accountId": "abc123-def456-..."
}
```
No need to switch accounts - just reference by ID!

## Installation

```bash
npm install -g forit-microsoft-graph
```

## Usage

### With Claude Desktop / MCP Client

Add to your MCP config:

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "npx",
      "args": ["-y", "forit-microsoft-graph"]
    }
  }
}
```

### Multi-Account Support

```bash
# Login to first account
microsoft-graph-mcp --login

# Login to second account (adds to cache)
microsoft-graph-mcp --login

# List accounts
microsoft-graph-mcp --list-accounts

# Select specific account
microsoft-graph-mcp --select-account <account-id>
```

## Use With PnP CLI

This MCP is designed to complement [PnP CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/):

- **Microsoft Graph MCP**: Raw Graph API access, multi-account auth
- **PnP CLI**: 500+ specialized commands for SharePoint, Teams, Power Platform

## Environment Variables

| Variable | Description |
|----------|-------------|
| `MS365_MCP_CLIENT_ID` | Azure AD app client ID (optional, uses default) |
| `MS365_MCP_CLIENT_SECRET` | Client secret for confidential apps |
| `MS365_MCP_TENANT_ID` | Azure AD tenant ID (default: `common`) |

## License

MIT - Original work by [Softeria](https://github.com/Softeria/ms-365-mcp-server)

## Credits

This project is a fork of [@softeria/ms-365-mcp-server](https://github.com/Softeria/ms-365-mcp-server), modified to provide a leaner, raw-API-focused experience.
