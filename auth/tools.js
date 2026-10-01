/**
 * Authentication-related tools for the Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');
const TokenStorage = require('./token-storage');

// Shared token store: ensureAuthenticated (auth/index.js), authenticate and
// check-auth-status all use this instance, so they see the same tokens.
const tokenStorage = new TokenStorage();

/**
 * About tool handler
 * @returns {object} - MCP response
 */
async function handleAbout() {
  return {
    content: [{
      type: "text",
      text: `M365 Assistant MCP Server v${config.SERVER_VERSION}\n\nProvides access to Microsoft 365 services through Microsoft Graph API:\n- Outlook (email, calendar, folders, rules)\n- OneDrive (files, folders, sharing)\n- Power Automate (flows, environments, runs)\n\nModular architecture for improved maintainability.`
    }]
  };
}

/**
 * Authentication tool handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAuthenticate(args) {
  const force = args && args.force === true;

  // For test mode, create a test token
  if (config.USE_TEST_MODE) {
    // Create a test token with a 1-hour expiry
    tokenManager.createTestTokens();

    return {
      content: [{
        type: "text",
        text: 'Successfully authenticated with Microsoft Graph API (test mode)'
      }]
    };
  }

  // Real authentication via OAuth 2.0 device code flow (public client, no secret).
  try {
    if (force) {
      await tokenStorage.clearTokens();
    }

    const dc = await tokenStorage.startDeviceCode();

    // Poll in the background; tokens are persisted to disk when sign-in completes.
    tokenStorage.pollDeviceCode(dc.device_code, dc.interval, dc.expires_in)
      .then(() => console.error('[AUTHENTICATE] Device code sign-in complete; tokens saved.'))
      .catch((err) => console.error(`[AUTHENTICATE] Device code polling failed: ${err.message}`));

    return {
      content: [{
        type: "text",
        text: `To sign in, open ${dc.verification_uri} in a browser and enter this code:\n\n    ${dc.user_code}\n\nSign-in completes automatically a few seconds after you finish in the browser. Then run "check-auth-status" to confirm, or just use any Outlook tool.`
      }]
    };
  } catch (error) {
    return {
      content: [{
        type: "text",
        text: `Failed to start authentication: ${error.message}`
      }]
    };
  }
}

/**
 * Check authentication status tool handler.
 * Uses the same path as real tool calls: an expired access token is refreshed
 * (and the refreshed tokens saved) when a refresh token is available.
 * @returns {object} - MCP response
 */
async function handleCheckAuthStatus() {
  console.error('[CHECK-AUTH-STATUS] Starting authentication status check');

  if (config.USE_TEST_MODE) {
    const testTokens = tokenManager.loadTokenCache();
    return {
      content: [{ type: "text", text: testTokens ? "Authenticated and ready (test mode)" : "Not authenticated (test mode)" }]
    };
  }

  const tokens = await tokenStorage.getTokens();
  if (!tokens || !tokens.access_token) {
    console.error('[CHECK-AUTH-STATUS] No tokens found');
    return {
      content: [{ type: "text", text: "Not authenticated. Use the 'authenticate' tool to sign in." }]
    };
  }

  const wasExpired = tokenStorage.isTokenExpired();
  const hadRefreshToken = !!tokens.refresh_token;
  const accessToken = await tokenStorage.getValidAccessToken();

  if (!accessToken) {
    console.error('[CHECK-AUTH-STATUS] Access token expired and could not be refreshed');
    return {
      content: [{
        type: "text",
        text: `Not authenticated: the access token has expired and ${hadRefreshToken ? 'the refresh failed' : 'there is no refresh token'}. Use the 'authenticate' tool to sign in again.`
      }]
    };
  }

  const expiresAt = new Date(tokenStorage.getExpiryTime()).toISOString();
  console.error(`[CHECK-AUTH-STATUS] Authenticated; access token ${wasExpired ? 'refreshed' : 'valid'}, expires at ${expiresAt}`);
  return {
    content: [{
      type: "text",
      text: `Authenticated and ready${wasExpired ? ' (access token was expired and has been refreshed)' : ''}. Access token expires at ${expiresAt}.`
    }]
  };
}

// Tool definitions
const authTools = [
  {
    name: "about",
    description: "Returns information about this M365 Assistant server",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleAbout
  },
  {
    name: "authenticate",
    description: "Authenticate with Microsoft Graph API to access Outlook data",
    inputSchema: {
      type: "object",
      properties: {
        force: {
          type: "boolean",
          description: "Force re-authentication even if already authenticated"
        }
      },
      required: []
    },
    handler: handleAuthenticate
  },
  {
    name: "check-auth-status",
    description: "Check the current authentication status with Microsoft Graph API",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleCheckAuthStatus
  }
];

module.exports = {
  tokenStorage,
  authTools,
  handleAbout,
  handleAuthenticate,
  handleCheckAuthStatus
};
