/**
 * Authentication-related tools for the Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');

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
  const TokenStorage = require('./token-storage');
  const storage = new TokenStorage();

  try {
    if (force) {
      await storage.clearTokens();
    }

    const dc = await storage.startDeviceCode();

    // Poll in the background; tokens are persisted to disk when sign-in completes.
    storage.pollDeviceCode(dc.device_code, dc.interval, dc.expires_in)
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
 * Check authentication status tool handler
 * @returns {object} - MCP response
 */
async function handleCheckAuthStatus() {
  console.error('[CHECK-AUTH-STATUS] Starting authentication status check');
  
  const tokens = tokenManager.loadTokenCache();
  
  console.error(`[CHECK-AUTH-STATUS] Tokens loaded: ${tokens ? 'YES' : 'NO'}`);
  
  if (!tokens || !tokens.access_token) {
    console.error('[CHECK-AUTH-STATUS] No valid access token found');
    return {
      content: [{ type: "text", text: "Not authenticated" }]
    };
  }
  
  console.error('[CHECK-AUTH-STATUS] Access token present');
  console.error(`[CHECK-AUTH-STATUS] Token expires at: ${tokens.expires_at}`);
  console.error(`[CHECK-AUTH-STATUS] Current time: ${Date.now()}`);
  
  return {
    content: [{ type: "text", text: "Authenticated and ready" }]
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
  authTools,
  handleAbout,
  handleAuthenticate,
  handleCheckAuthStatus
};
