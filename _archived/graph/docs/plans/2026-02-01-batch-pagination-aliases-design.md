# Design: Batch Requests, Auto-Pagination, and Account Aliases

**Date:** 2026-02-01
**Version:** 1.3.0
**Status:** Approved

## Overview

Three features to enhance the ForIT Microsoft Graph MCP:

1. **`graph-batch`** - Execute up to 20 Graph API requests in a single call
2. **`fetchAllPages`** - Auto-paginate `graph-request` results
3. **Account Aliases** - Use friendly names instead of GUIDs

All maintain the "lean" philosophy: minimal context footprint, maximum flexibility.

## 1. `graph-batch` Tool

### Purpose

Execute multiple Graph API requests in one network round-trip. Reduces latency and token usage when querying multiple endpoints.

### Parameters

```typescript
{
  requests: Array<{
    id: string;           // Unique ID for matching responses
    url: string;          // Graph API path (e.g., "/me/messages")
    method?: string;      // GET, POST, PUT, PATCH, DELETE (default: GET)
    body?: object;        // Request body for POST/PUT/PATCH
    headers?: object;     // Additional headers per request
  }>;
  accountId?: string;     // Required when multiple accounts exist (accepts alias)
}
```

### Response

```typescript
{
  responses: Array<{
    id: string;           // Matches request ID
    status: number;       // HTTP status code
    body: object;         // Response body (parsed JSON)
    headers?: object;     // Response headers if relevant
  }>;
}
```

### Constraints

- Maximum 20 requests per batch (Microsoft Graph API limit)
- All requests execute in parallel (no `dependsOn` in v1)
- Individual request failures don't fail the entire batch
- Each response includes its own status code

### Example

```json
{
  "requests": [
    { "id": "mail", "url": "/me/messages?$top=5" },
    { "id": "calendar", "url": "/me/calendar/events?$top=5" },
    { "id": "profile", "url": "/me" }
  ],
  "accountId": "work"
}
```

### Implementation

Uses Microsoft Graph `/$batch` endpoint:
- POST to `https://graph.microsoft.com/v1.0/$batch`
- Body: `{ "requests": [...] }`
- Single auth token for all requests

## 2. `fetchAllPages` Parameter

### Purpose

Automatically follow `@odata.nextLink` pagination without manual handling.

### Addition to `graph-request`

```typescript
fetchAllPages?: boolean;  // Default: false
```

### Behavior

When `fetchAllPages: true`:

1. Execute initial request
2. If response contains `@odata.nextLink`, fetch next page
3. Concatenate `value` arrays from all pages
4. Repeat until no more `@odata.nextLink` or 100 page limit reached
5. Return combined response with aggregated `value` array
6. Remove `@odata.nextLink` from final response
7. Update `@odata.count` if present

### Safety

- Maximum 100 pages (prevents runaway pagination)
- Logs warning if limit reached
- Respects rate limiting between requests

### Example

```json
{
  "endpoint": "/me/messages",
  "queryParams": { "$select": "subject,from" },
  "fetchAllPages": true,
  "accountId": "work"
}
```

Returns all messages, not just first page.

## 3. Account Aliases

### Purpose

Use friendly names like "work" or "personal" instead of account GUIDs.

### New Tool: `set-account-alias`

```typescript
{
  alias: string;      // Friendly name (e.g., "work", "personal", "forit")
  accountId: string;  // Full account ID to alias
}
```

### New Tool: `remove-account-alias`

```typescript
{
  alias: string;      // Alias to remove
}
```

### Updated `list-accounts` Response

```typescript
{
  accounts: Array<{
    id: string;
    username: string;
    name: string;
    appId: string;
    tenantId: string;
    aliases: string[];    // NEW: List of aliases for this account
  }>;
}
```

### Alias Resolution

All tools accepting `accountId` now resolve aliases:
- `graph-request`
- `graph-batch`
- `remove-account`

Resolution order:
1. Check if input matches an alias → use mapped accountId
2. Otherwise, treat as literal accountId

### Storage

Aliases stored in token cache alongside account metadata:

```json
{
  "accountAliases": {
    "work": "abc123-def456-...",
    "personal": "xyz789-..."
  }
}
```

### Example Usage

```bash
# Set up alias
set-account-alias alias="work" accountId="abc123-def456-..."

# Use alias in requests
graph-request endpoint="/me" accountId="work"
graph-batch requests=[...] accountId="work"
```

## File Changes

### `src/auth-tools.ts`

- Add `graph-batch` tool
- Add `fetchAllPages` to `graph-request`
- Add `set-account-alias` tool
- Add `remove-account-alias` tool
- Update `list-accounts` to include aliases

### `src/auth.ts`

- Add `resolveAccountId(idOrAlias)` method
- Add `setAlias(alias, accountId)` method
- Add `removeAlias(alias)` method
- Add `getAliasesForAccount(accountId)` method
- Update cache schema for aliases

### `src/graph-client.ts`

- Add `batchRequest(requests, options)` method
- Reuse pagination logic from `graph-tools.ts`

## Version Bump

- Current: 1.2.1
- New: 1.3.0 (new features, backward compatible)

## Testing

1. **Batch requests:** Verify 20-request limit, error handling, response mapping
2. **Pagination:** Test with large mailboxes, verify 100-page cap
3. **Aliases:** CRUD operations, resolution in all tools, persistence across restarts

## Future Considerations (Not in v1.3)

- `dependsOn` for batch request chaining
- `maxPages` parameter for fine-grained pagination control
- Auto-alias suggestion based on email domain
