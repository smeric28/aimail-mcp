# AI Mail MCP — Microsoft 365 / Graph Connector

A Model Context Protocol (MCP) server that gives Claude full read **and write**
access to Microsoft 365 through the Microsoft Graph API — mirroring the official
Microsoft 365 connector surface and then some. Built for **Fireball Industries**
to run as a shared, multi-user remote connector available in the Claude apps on
**iPhone, desktop, and web**, with a local (stdio) mode for development.

> Status: `v3.0` adds the remote OAuth transport, hosting artifacts, and the
> full Graph surface (files, contacts, tasks). See [Roadmap](#roadmap).

---

## What it can do

| Surface | Tools |
| --- | --- |
| **Mail** | `search_emails`, `get_email`, `create_draft`, `update_draft`, `delete_draft`, `send_email`, `send_draft`, `reply_to_email`, `forward_email`, `move_email`, `set_email_read`, `flag_email`, `list_mail_folders` |
| **Attachments** | `find_attachments`, `download_attachment`, `copy_attachment_to_draft`, `add_attachment_to_draft` |
| **Calendar** | `list_calendar_events`, `create_calendar_event`, `update_calendar_event`, `cancel_calendar_event`, `delete_calendar_event`, `respond_to_calendar_event`, `find_available_times` |
| **Files (OneDrive + SharePoint)** | `list_files`, `search_files`, `read_file`, `upload_file` (large files via upload session), `create_folder`, `delete_file`, `share_file`, `list_sharepoint_sites`, `list_site_drives` |
| **Contacts** | `list_contacts`, `search_contacts`, `create_contact`, `update_contact`, `delete_contact` |
| **Tasks (Microsoft To Do)** | `list_task_lists`, `list_tasks`, `create_task`, `update_task`, `complete_task`, `delete_task` |
| **Teams chat** | `list_chats`, `list_chat_messages`, `send_chat_message` |
| **Accounts** | `list_accounts` |

Gmail is also supported for mail in local mode via a shared `MailProvider`
interface; the Microsoft 365 surfaces above are Graph-only.

---

## Architecture

There is **one server, two transports**, sharing a single tool registry
(`src/tools/registry.ts`) so capabilities are identical regardless of how it is
reached.

```
                     ┌───────────────────────────────────────────┐
                     │            aimail-mcp server                │
   Claude apps ──────┤  buildMcpServer() + shared tool registry    │
  (iOS/desktop/web)  │     │                         │             │
        │            │  stdio transport         HTTP transport      │
        │            │  (local dev)         (remote connector)      │
        │            └─────┼─────────────────────────┼─────────────┘
        │                  │                          │
        │            MSAL local cache         OAuth bridge → Entra ID
        │            (~/.aimail-mcp-cache)     (per-user delegated token)
        ▼                                            │
   Custom connector  ───────────────────────────────┘
   over HTTPS + OAuth                         Microsoft Graph API
```

- **stdio transport** (`MCP_TRANSPORT=stdio`, default): the original local mode.
  Launched as a subprocess by Claude Desktop / OpenCode. Uses MSAL with a local
  browser sign-in and an on-disk token cache. Single user, single machine.
- **HTTP transport** (`MCP_TRANSPORT=http`): a hosted, internet-reachable server
  that Claude adds as a **custom connector**. Each request carries the
  signed-in user's delegated Graph token, so it is genuinely multi-user. **This
  is the only mode the Claude iPhone app can use.**

### Why the remote transport is required for iPhone

The Claude mobile app (and the cleanest desktop/web path) can only talk to
**remote MCP servers** — HTTPS endpoints with their own OAuth flow. A local
stdio server that opens `localhost` and writes a token file works on exactly one
desktop. To serve the whole Fireball team across devices, the server is hosted
and each person signs in with their own Microsoft account.

---

## Remote connector setup (Fireball, multi-user)

### 1. Entra app registration

The app registration already exists in the Fireball tenant. Add the full Graph
scope set and the hosted OAuth callback, then grant admin consent once for all
users:

```powershell
./scripts/update-azure-app-graph-scopes.ps1 `
    -AppId <application-client-id> `
    -PublicBaseUrl https://aimail.fireballz.ai `
    -GrantAdminConsent
```

This requests the delegated scopes in `mcp-server/src/config/scopes.ts`
(`Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `Contacts.ReadWrite`,
`Files.ReadWrite.All`, `Sites.ReadWrite.All`, `Tasks.ReadWrite`, …) and adds the
redirect URI `https://aimail.fireballz.ai/auth/callback` (type **Web**).

> Need a fresh registration instead? Use `scripts/setup-azure-ad-graph.ps1`.

### 2. Deploy the server

```powershell
./scripts/deploy-azure-containerapp.ps1 `
    -ResourceGroup aimail-rg -Location eastus `
    -ClientId <app-id> -ClientSecret <secret> -TenantId <tenant-guid> `
    -PublicBaseUrl https://aimail.fireballz.ai
```

Azure Container Apps is the recommended host (same directory as the app
registration, scales to zero, simple HTTPS ingress). The container runs
`MCP_TRANSPORT=http`. Any container host works — see
[`mcp-server/Dockerfile`](mcp-server/Dockerfile).

Required runtime env: `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`,
`AZURE_TENANT_ID`, `PUBLIC_BASE_URL`, `PORT` (see `.env.example`).

### 3. Add the connector in Claude

In Claude (iPhone, desktop, or web): **Settings → Connectors → Add custom
connector**, then enter:

```
https://aimail.fireballz.ai/mcp
```

Claude discovers the OAuth metadata, registers itself, and sends the user
through Microsoft sign-in. After consent, the connector's tools are available in
chat. Each Fireball member signs in as themselves and only ever sees their own
mailbox, files, and calendar.

### How the OAuth flow works

Microsoft Entra ID does not support open Dynamic Client Registration, which the
Claude connector flow expects. So the server presents **itself** as the
authorization server and brokers the real flow to Entra
(`src/server/oauthBridge.ts`):

1. Claude discovers `/.well-known/oauth-protected-resource` and
   `/.well-known/oauth-authorization-server`.
2. Claude registers via `/register` (DCR shim) and starts Authorization Code +
   PKCE at `/authorize`.
3. `/authorize` redirects the user to Entra (requesting Graph scopes). Entra
   calls back to `/auth/callback`; the server exchanges the code and mints its
   own short-lived code for Claude.
4. `/token` validates Claude's PKCE and returns the Entra-issued tokens. Because
   we request **Graph** scopes, those access tokens are forwarded straight to
   Microsoft Graph (token passthrough).

> **Security note:** token passthrough is the simplest correct model for an
> internal connector. If you later want the server to hold its own audience,
> switch to the On-Behalf-Of flow in the bridge. Tokens are never persisted by
> the server; the in-memory auth-code/state maps are short-lived (10 min) and
> should be backed by Redis if you scale beyond one replica.

---

## Local (stdio) setup — development

```bash
cd mcp-server
npm install
npm run build
cp .env.example .env   # set AZURE_CLIENT_ID / SECRET / TENANT_ID
npm start              # MCP_TRANSPORT=stdio
```

Register with Claude Desktop / OpenCode by pointing the MCP client at
`node /absolute/path/mcp-server/dist/index.js`. First call opens a browser for
Microsoft sign-in; the token is cached under `~/.aimail-mcp-cache`.

For multiple accounts/shared mailboxes, set `ACCOUNTS_CONFIG` (JSON) — see
`src/auth/providerManager.ts`.

---

## Development

```bash
cd mcp-server
npm run build      # tsc
npm test           # jest
npm run start:http # run the remote server locally on :8080
```

Smoke-test the remote endpoints locally:

```bash
curl localhost:8080/healthz
curl localhost:8080/.well-known/oauth-protected-resource
```

---

## Roadmap

- [x] Remote HTTP transport + Entra OAuth bridge (iPhone/desktop/web)
- [x] Full Graph surface: mail send/reply/forward + attachment download,
      calendar update/cancel/RSVP, OneDrive + SharePoint files, contacts,
      To Do tasks, Teams chat
- [x] Large-file upload sessions (>4 MB)
- [x] Hosting artifacts (Dockerfile, Azure Container Apps deploy)
- [ ] On-Behalf-Of token exchange (own API audience) + Redis-backed auth state
- [ ] SharePoint list items; Planner; Teams channel messages
- [ ] Per-tool scope minimization & consent screens

---

## Repos

This connector is mirrored to the `Fireball-WhiteTeam` org for team use. The
canonical development repo is `smeric28/aimail-mcp`.
