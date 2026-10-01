# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

- `npm install` - **ALWAYS run first** to install dependencies
- `npm start` - Start the MCP server
- `node scripts/login.mjs` - Sign in with the device code flow (JSON lines on stdout); `--status` shows sign-in status
- `npm run test-mode` - Start the server in test mode with mock data
- `npm run inspect` - Use MCP Inspector to test the server interactively
- `npm test` - Run Jest tests

## Architecture Overview

This is a modular MCP (Model Context Protocol) server that provides Claude with access to Microsoft 365 services:
- **Outlook** - Email, calendar, folders, rules
- **OneDrive** - Files, folders, sharing
- **Power Automate** - Flows, environments, runs

### Core Structure
- `index.js` - Main entry point that combines all module tools and handles MCP protocol
- `config.js` - Centralized configuration (API endpoints, scopes, field selections)
- `scripts/login.mjs` - Sign-in CLI (device code flow) using the same token store as the server

### Modules
Each module exports tools and handlers:
- `auth/` - OAuth 2.0 authentication with token management (Graph + Flow)
- `calendar/` - Calendar operations (list, create, accept, decline, delete events)
- `email/` - Email management (list, search, read, send, mark as read)
- `folder/` - Folder operations (list, create, move)
- `rules/` - Email rules management
- `onedrive/` - OneDrive operations (list, search, download, upload, share, folder ops)
- `power-automate/` - Flow operations (list environments, list/run/toggle flows, run history)
- `utils/` - Shared utilities including Graph API client and OData helpers

### Key Components
- **Token Management**: Tokens stored in `~/.outlook-mcp-tokens.json` (both Graph and Flow tokens)
- **Graph API Client**: `utils/graph-api.js` handles Microsoft Graph API calls (Outlook, OneDrive)
- **Flow API Client**: `power-automate/flow-api.js` handles Power Automate API calls
- **Test Mode**: Mock data responses when `USE_TEST_MODE=true`
- **Modular Tools**: Each module exports tools array that gets combined in main server

## Authentication

### Graph API (Outlook + OneDrive)
1. Azure app registration (public client, no client secret) with **Allow public client flows** = Yes and delegated permissions:
   - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.Read`, `Calendars.ReadWrite`
   - `Files.Read`, `Files.ReadWrite`
   - `User.Read`, `offline_access`
2. Sign in with the device code flow: the `authenticate` tool (returns URL + code) or `node scripts/login.mjs`.
   No auth server is needed.
3. Tokens are stored in `~/.outlook-mcp-tokens.json` and refreshed automatically (`auth/token-storage.js`);
   a running server re-reads the file when another process has written new tokens.

### Power Automate (Optional)
- Requires the Flow API scope `https://service.flow.microsoft.com/.default` in the app registration
- The Flow token is obtained/renewed with the refresh token (independent of the Graph token) and stored in the same token file
- Only solution-aware flows accessible via API
- Only manual trigger flows can be triggered

## Configuration

### Environment Variables
- **For .env file**: Use `MS_CLIENT_ID` (and `MS_TENANT_ID` for single-tenant apps)
- **For Claude Desktop config**: Use `OUTLOOK_CLIENT_ID`
- No client secret: the server is a public client

### Config Constants
- `GRAPH_API_ENDPOINT`: `https://graph.microsoft.com/v1.0/`
- `FLOW_API_ENDPOINT`: `https://api.flow.microsoft.com`
- `ONEDRIVE_UPLOAD_THRESHOLD`: 4MB (files larger need chunked upload)
- Default page size: 25, max results: 50

### Common Setup Issues
1. **Missing dependencies**: Always run `npm install` first
2. **AADSTS7000218 at sign-in**: Allow public client flows in the app registration
3. **stdout is the MCP stream**: log with `console.error`, never `console.log`, in server code

## Test Mode

Set `USE_TEST_MODE=true` to use mock data instead of real API calls. Mock responses defined in:
- `utils/mock-data.js` - Graph API mocks
- `power-automate/flow-api.js` - Flow API mocks (inline)

## Error Handling

- Graph API auth failures: "UNAUTHORIZED" error
- Flow API auth failures: "FLOW_UNAUTHORIZED" error
- API errors include status codes and response details
- Token expiration triggers re-authentication flow
- Empty API responses handled gracefully
