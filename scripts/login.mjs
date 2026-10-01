#!/usr/bin/env node
/**
 * Sign in to Microsoft Graph with the device code flow, or report sign-in status.
 *
 *   node scripts/login.mjs           -> device code sign-in, JSON lines on stdout
 *   node scripts/login.mjs --status  -> one JSON line with the current status
 *
 * Uses auth/token-storage.js, the same token store (~/.outlook-mcp-tokens.json),
 * app registration (MS_CLIENT_ID / MS_TENANT_ID) and scopes as the MCP server.
 * Token values are never printed.
 */
import { createRequire } from 'module';

const require = createRequire(import.meta.url);

// TokenStorage logs with console.log; keep stdout reserved for the JSON lines.
console.log = console.error;

const TokenStorage = require('../auth/token-storage.js');
const { callGraphAPI } = require('../utils/graph-api.js');

function emit(obj) {
  process.stdout.write(JSON.stringify(obj) + '\n');
}

async function fetchAccount(accessToken) {
  try {
    const me = await callGraphAPI(accessToken, 'GET', 'me', null, { $select: 'userPrincipalName,mail' });
    return me.userPrincipalName || me.mail || null;
  } catch {
    return null;
  }
}

// Error text for app registrations that don't allow public client flows.
const PUBLIC_CLIENT_HELP =
  'Appregistreringen tillåter inte offentliga klientflöden (device code). Öppna Azure Portal > App registrations > ' +
  'appen > Authentication och sätt "Tillåt offentliga klientflöden" / "Allow public client flows" till Ja/Yes, spara och försök igen.';

function explain(error) {
  const message = (error && error.message) || String(error);
  if (/AADSTS7000218|AADSTS70002\b|client_assertion|client_secret/i.test(message)) {
    return `${PUBLIC_CLIENT_HELP} (${message})`;
  }
  return message;
}

async function status() {
  const storage = new TokenStorage();
  const tokens = await storage.getTokens();
  const result = {
    authenticated: false,
    expiresAt: tokens && tokens.expires_at ? new Date(tokens.expires_at).toISOString() : null,
    hasRefreshToken: !!(tokens && tokens.refresh_token),
    account: null
  };

  if (tokens && tokens.access_token) {
    let accessToken = null;
    if (!storage.isTokenExpired()) {
      accessToken = tokens.access_token;
    } else if (tokens.refresh_token) {
      try {
        accessToken = await storage.refreshAccessToken(); // saves the new tokens
      } catch (error) {
        console.error(`Refresh failed: ${error.message}`);
      }
    }
    if (accessToken) {
      result.authenticated = true;
      result.expiresAt = new Date(storage.tokens.expires_at).toISOString();
      result.hasRefreshToken = !!storage.tokens.refresh_token;
      result.account = await fetchAccount(accessToken);
    }
  }

  emit(result);
}

async function login() {
  const storage = new TokenStorage();
  try {
    const dc = await storage.startDeviceCode();
    emit({
      type: 'code',
      userCode: dc.user_code,
      verificationUri: dc.verification_uri,
      message: dc.message,
      expiresIn: dc.expires_in
    });
    const tokens = await storage.pollDeviceCode(dc.device_code, dc.interval, dc.expires_in);
    emit({ type: 'done', account: await fetchAccount(tokens.access_token) });
  } catch (error) {
    emit({ type: 'error', message: explain(error) });
    process.exit(1);
  }
}

if (process.argv.includes('--status')) {
  status().then(() => process.exit(0), (error) => {
    console.error(error);
    emit({ authenticated: false, expiresAt: null, hasRefreshToken: false, account: null });
    process.exit(0);
  });
} else {
  login().then(() => process.exit(0));
}
