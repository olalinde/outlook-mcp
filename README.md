[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/ryaker-outlook-mcp-badge.png)](https://mseep.ai/app/ryaker-outlook-mcp)

# M365 Assistant MCP Server

A comprehensive MCP (Model Context Protocol) server that connects Claude with Microsoft 365 services through the Microsoft Graph API and Power Automate API.

## Supported Services

- **Outlook** - Email, calendar, folders, and rules
- **OneDrive** - Files, folders, search, and sharing
- **Power Automate** - Flows, environments, and run history

## Directory Structure

```
├── index.js                 # Main entry point
├── config.js                # Configuration settings
├── auth/                    # Authentication modules
│   ├── index.js             # Authentication exports
│   ├── token-storage.js     # Token store: device code flow, Graph and Flow token refresh
│   ├── token-manager.js     # Test-mode tokens
│   └── tools.js             # Auth-related tools
├── calendar/                # Calendar functionality
│   ├── index.js             # Calendar exports
│   ├── list.js              # List events
│   ├── create.js            # Create event
│   ├── delete.js            # Delete event
│   ├── cancel.js            # Cancel event
│   ├── accept.js            # Accept event
│   └── decline.js           # Decline event
├── email/                   # Email functionality
│   ├── index.js             # Email exports
│   ├── list.js              # List emails
│   ├── search.js            # Search emails
│   ├── read.js              # Read email
│   ├── send.js              # Send email
│   ├── draft.js             # Create draft
│   ├── attachments.js       # File attachments (send/draft) and download-attachments
│   └── mark-as-read.js      # Mark email read/unread
├── folder/                  # Folder functionality
│   ├── index.js             # Folder exports
│   ├── list.js              # List folders
│   ├── create.js            # Create folder
│   └── move.js              # Move emails
├── rules/                   # Email rules functionality
│   ├── index.js             # Rules exports
│   ├── list.js              # List rules
│   └── create.js            # Create rule
├── onedrive/                # OneDrive functionality
│   ├── index.js             # OneDrive exports
│   ├── list.js              # List files/folders
│   ├── search.js            # Search files
│   ├── download.js          # Get download URL
│   ├── upload.js            # Simple upload (<4MB)
│   ├── upload-large.js      # Chunked upload (>4MB)
│   ├── share.js             # Create sharing link
│   └── folder.js            # Create/delete folders
├── power-automate/          # Power Automate functionality
│   ├── index.js             # Power Automate exports
│   ├── flow-api.js          # Flow API client
│   ├── list-environments.js # List environments
│   ├── list-flows.js        # List flows
│   ├── run-flow.js          # Trigger flow
│   ├── list-runs.js         # Run history
│   └── toggle-flow.js       # Enable/disable flow
├── scripts/
│   └── login.mjs            # Sign-in CLI (device code) and status
└── utils/                   # Utility functions
    ├── graph-api.js         # Microsoft Graph API helper
    ├── odata-helpers.js     # OData query building
    └── mock-data.js         # Test mode data
```

## Features

- **Authentication**: OAuth 2.0 authentication with Microsoft Graph API (+ Flow API for Power Automate)
- **Email Management**: List, search, read, send, and organize emails
- **Calendar Management**: List, create, accept, decline, and delete calendar events
- **OneDrive Integration**: List, search, upload, download, and share files
- **Power Automate**: List environments/flows, trigger flows, view run history
- **Modular Structure**: Clean separation of concerns for maintainability
- **Test Mode**: Simulated responses for testing without real API calls

## Available Tools

### Outlook (Email & Calendar)
| Tool | Description |
|------|-------------|
| `list-emails` | List recent emails (`folder`: inbox, `sent`/`sentitems`, drafts, ...) |
| `search-emails` | Search emails with filters |
| `read-email` | Read email content (`format: "html"` for the unmodified HTML + attachments) |
| `send-email` | Send a new email (optional `attachments`) |
| `draft-email` | Create a draft (optional `attachments`, `isHtml`) |
| `download-attachments` | Save an email's file attachments to a local folder |
| `mark-as-read` | Mark email as read/unread |
| `list-events` | List calendar events |
| `create-event` | Create calendar event |
| `accept-event` | Accept event invitation |
| `decline-event` | Decline event invitation |
| `delete-event` | Delete calendar event |
| `list-folders` | List mail folders |
| `create-folder` | Create mail folder |
| `move-emails` | Move emails between folders |
| `list-rules` | List inbox rules |
| `create-rule` | Create inbox rule |

### Attachments and HTML email

`send-email` and `draft-email` take an optional `attachments` array:

```json
"attachments": [
  { "path": "C:/temp/diagram.png", "isInline": true, "contentId": "diagram1" },
  { "path": "C:/temp/rapport.pdf", "name": "Rapport.pdf" }
]
```

- `path` (required): absolute local file path. The file is read and sent as a Graph `#microsoft.graph.fileAttachment`.
- `name`: defaults to the file name.
- `contentType`: defaults from the extension (png, jpg/jpeg, gif, svg, pdf), otherwise `application/octet-stream`.
- `isInline` (default false) and `contentId`: for embedded images, referenced in the HTML body as `<img src="cid:diagram1">`.
- Max 3 MB per file; larger files are rejected with an error before anything is sent.
- The HTML body is sent unchanged (inline styles are kept). Use `isHtml: true` when the body has no `<html>` tag.

`download-attachments` — `{ messageId, saveDir, inlineOnly? }`: saves the message's file attachments into `saveDir`
(absolute; created if missing). Unsafe characters in names are replaced, files with the same name in `saveDir` are
overwritten, and duplicate names within one message get `-1`, `-2` .... Returns JSON:

```json
[{ "path": "C:/temp/ut/diagram.png", "name": "diagram.png", "contentType": "image/png", "contentId": "diagram1", "isInline": true, "size": 12345 }]
```

`read-email` with `format: "html"` returns JSON instead of sanitized text:
`{ id, subject, from, to, cc, sentDateTime, receivedDateTime, bodyContentType, body, attachments: [{ id, name, contentType, contentId, isInline, size }] }`,
where `body` is the unmodified HTML. Note that this bypasses the prompt-injection sanitizing of the default text format.
To find a sent message: `list-emails` with `folder: "sentitems"` (or `"sent"`) lists subjects and ids.
Folder names may be paths such as `"Inbox/Projekt/2026"`. A default folder as the first part is found whatever its
display name, in English or Swedish: Inbox/Inkorg, Sent Items/Skickat, Drafts/Utkast, Deleted Items/Borttaget,
Archive/Arkiv, Junk/Skräppost. `search-emails` never returns non-matching emails: with no hits it says so, and hits that
match only part of the criteria are labelled as such.

### OneDrive
| Tool | Description |
|------|-------------|
| `onedrive-list` | List files in a path |
| `onedrive-search` | Search files by query |
| `onedrive-download` | Get download URL |
| `onedrive-upload` | Upload small file (<4MB) |
| `onedrive-upload-large` | Chunked upload (>4MB) |
| `onedrive-share` | Create sharing link |
| `onedrive-create-folder` | Create folder |
| `onedrive-delete` | Delete file or folder |

### Power Automate
| Tool | Description |
|------|-------------|
| `flow-list-environments` | List Power Platform environments |
| `flow-list` | List flows in environment |
| `flow-run` | Trigger a manual flow |
| `flow-list-runs` | Get flow run history |
| `flow-toggle` | Enable/disable a flow |

## Quick Start

1. **Install dependencies**: `npm install`
2. **Azure setup**: Register app in Azure Portal (see detailed steps below)
3. **Configure environment**: Copy `.env.example` to `.env` and add your Azure credentials
4. **Configure Claude**: Update your Claude Desktop config with the server path
5. **Authenticate**: Use the `authenticate` tool in Claude, or run `node scripts/login.mjs` (device code; no auth server needed)
6. **Start using**: Access your M365 data through Claude!

## Installation

### Prerequisites
- Node.js 14.0.0 or higher
- npm or yarn package manager
- Azure account for app registration

### Install Dependencies

```bash
npm install
```

## Azure App Registration & Configuration

### App Registration

1. Open [Azure Portal](https://portal.azure.com/)
2. Search for "App registrations"
3. Click "New registration"
4. Name: "M365 MCP Server"
5. Account type: "Accounts in any organizational directory and personal Microsoft accounts"
6. Redirect URI: leave empty (sign-in uses the device code flow)
7. Click "Register"
8. Under **Authentication**, set **Allow public client flows** to **Yes** and save
9. Copy the "Application (client) ID" for your `.env` file

### App Permissions

1. Go to "API permissions" under Manage
2. Click "Add a permission" → "Microsoft Graph" → "Delegated permissions"
3. Add these permissions:
   - `offline_access`
   - `User.Read`
   - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.Read`, `Calendars.ReadWrite`
   - `Files.Read`, `Files.ReadWrite`
4. Click "Add permissions"

**For Power Automate** (optional):
- Requires additional Azure AD configuration with Flow API scope
- See Power Automate section below for details

No client secret is needed: the server is a public client and signs in with the device code flow.

## Configuration

### 1. Environment Variables

```bash
cp .env.example .env
```

Edit `.env`:
```bash
# Get these values from Azure Portal > App Registrations > Your App
MS_CLIENT_ID=your-application-client-id-here
MS_TENANT_ID=your-tenant-id-here
USE_TEST_MODE=false
```

**Important Notes:**
- Use `MS_CLIENT_ID` in the `.env` file (or `OUTLOOK_CLIENT_ID` in the Claude Desktop config)
- Set `MS_TENANT_ID` for single-tenant apps to avoid `/common` endpoint errors

### 2. Claude Desktop Configuration

Add to your Claude Desktop config:

```json
{
  "mcpServers": {
    "m365-assistant": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/index.js"],
      "env": {
        "USE_TEST_MODE": "false",
        "OUTLOOK_CLIENT_ID": "your-client-id",
        "MS_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

## Authentication

### Graph API (Outlook + OneDrive)

Sign-in uses the OAuth 2.0 device code flow; no auth server or client secret is needed.

1. Use the `authenticate` tool in Claude (`force: true` signs in again from scratch), or run `node scripts/login.mjs`
2. Open the URL shown, enter the code and sign in
3. Tokens are saved to `~/.outlook-mcp-tokens.json`; `check-auth-status` confirms (and refreshes an expired access token)

A running server picks up tokens that `scripts/login.mjs` (or another server) has written to the token file.

### Sign-in CLI (`scripts/login.mjs`)

The server uses `auth/token-storage.js` for all Graph calls: tokens in `~/.outlook-mcp-tokens.json`, app registration
from `MS_CLIENT_ID` (or `OUTLOOK_CLIENT_ID`) and `MS_TENANT_ID`. An expired access token is refreshed automatically
with the refresh token; only when that fails do tools answer "Authentication required".

Sign in outside Claude (works from any working directory; stdout is JSON lines only, logs go to stderr):

```bash
node scripts/login.mjs
# {"type":"code","userCode":"ABCD1234","verificationUri":"https://microsoft.com/devicelogin","message":"...","expiresIn":900}
# ... user signs in in the browser ...
# {"type":"done","account":"name@company.com"}
# on failure: {"type":"error","message":"..."} and exit code 1
```

Status (tries a refresh if the access token has expired, and saves the new tokens if it works; always exit code 0):

```bash
node scripts/login.mjs --status
# {"authenticated":true,"expiresAt":"2026-10-01T12:00:00.000Z","hasRefreshToken":true,"account":"name@company.com"}
```

**Requirement: public client flows.** The device code flow needs the app registration to allow it: Azure Portal →
App registrations → the app → **Authentication** → **Allow public client flows** ("Tillåt offentliga klientflöden")
= **Yes** → Save. Otherwise sign-in fails with AADSTS7000218, and `login.mjs` prints that instruction.

### Power Automate (Optional)

Power Automate uses a separate token for the Flow API scope (`https://service.flow.microsoft.com/.default`). Add the
delegated Power Automate (Flow service) permission to the app registration. The Flow token is obtained and renewed
with the refresh token from the normal sign-in, independently of the Graph token; no extra sign-in is needed.

**Limitations:**
- Only solution-aware flows are accessible
- Only manual trigger flows can be run via API
- Requires environment ID for most operations

## Troubleshooting

### Common Issues

**"Cannot find module"**
```bash
npm install
```

**AADSTS7000218 / "client_assertion or client_secret" during sign-in**
- Allow public client flows in the app registration (see "Requirement: public client flows" above)

**"Authentication required"**
- The refresh token no longer works: run `node scripts/login.mjs` (or the `authenticate` tool) to sign in again

## Testing

```bash
# Run with MCP Inspector
npm run inspect

# Run in test mode (mock data)
npm run test-mode

# Run Jest tests
npm test
```

## Extending the Server

1. Create new module directory
2. Implement tool handlers in separate files
3. Export tool definitions from module index
4. Import and add to `TOOLS` array in `index.js`
