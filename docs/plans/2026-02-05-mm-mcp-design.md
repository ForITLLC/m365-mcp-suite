# MM MCP Design - Unified M365 PowerShell Management

## Overview

Single MCP for all Microsoft 365 PowerShell operations. Designed for MSP use - fast tenant onboarding via device code flow.

## Modules

| Module | Connect Command | Device Code Source | App ID Required |
|--------|-----------------|-------------------|-----------------|
| EXO | `Connect-ExchangeOnline -Device` | Native (PS7) | No - uses Microsoft's built-in app |
| Azure | `Connect-AzAccount -UseDeviceAuthentication` | Native | No - uses Microsoft's built-in app |
| Teams | `Connect-MicrosoftTeams -UseDeviceAuthentication` | Native | No - uses Microsoft's built-in app |
| PnP | `Connect-PnPOnline -Url <url> -DeviceLogin -ClientId <id>` | Native | **YES** - custom app per tenant (multi-tenant app deleted Sept 2024) |
| Power Platform | `pac auth create --deviceCode --tenant <tenant>` | Native (PAC CLI) | No - uses PAC CLI's built-in flow |

## Architecture

```
┌─────────────────────────────────────────┐
│              MM MCP Server              │
│         (Python + Flask/Gunicorn)       │
├─────────────────────────────────────────┤
│  Session Pool                           │
│  ┌─────────┐ ┌─────────┐ ┌─────────┐   │
│  │ EXO     │ │ Azure   │ │ Teams   │   │
│  │ Session │ │ Session │ │ Session │   │
│  └────┬────┘ └────┬────┘ └────┬────┘   │
│       │           │           │         │
│       └───────────┼───────────┘         │
│                   │                     │
│            PowerShell Process           │
│            (per connection/module)      │
└─────────────────────────────────────────┘
                    │
                    ▼
        ┌───────────────────────┐
        │ ~/.m365-connections   │
        │        .json          │
        └───────────────────────┘
```

## Connection Registry

Single app per tenant with all required permissions:

```json
{
  "connections": {
    "TenantName": {
      "appId": "your-app-id",
      "tenant": "tenant.onmicrosoft.com",
      "description": "What this connection is for",
      "mcps": ["mm"]
    }
  }
}
```

## Authentication Flow

1. User calls `mm run --connection TenantName --module exo --command "Get-Mailbox"`
2. If no active session: start PowerShell, run native device code connect
3. Capture device code from stdout, return to user
4. User authenticates in browser
5. Poll stdout for auth completion
6. Run command, return output

## App Permissions Required

Single Azure AD app registration per tenant needs:

**Microsoft Graph (Delegated):**
- User.Read.All
- Group.ReadWrite.All
- AppCatalog.ReadWrite.All
- TeamSettings.ReadWrite.All
- Channel.Delete.All
- ChannelSettings.ReadWrite.All
- ChannelMember.ReadWrite.All

**Office 365 Exchange Online (Delegated):**
- Exchange.Manage

**Skype and Teams Tenant Admin API (Delegated):**
- user_impersonation

**SharePoint (Delegated):**
- AllSites.FullControl (or as needed)

**Azure Management (Delegated):**
- user_impersonation

## Session Management

- One PowerShell process per connection/module pair
- Sessions persist until idle timeout (default: 24h)
- Health checks via module-specific commands
- Automatic reconnection on session failure

## MCP Tool Interface

```
mm_run(connection, module, command) -> output | device_code_prompt
mm_status() -> list of active sessions
mm_disconnect(connection, module) -> success
```
