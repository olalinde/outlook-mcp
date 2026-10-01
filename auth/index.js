/**
 * Authentication module for Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');
const { authTools, tokenStorage } = require('./tools');

/**
 * Ensures the user is authenticated and returns an access token.
 * Automatically refreshes expired tokens using the refresh_token grant.
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(forceNew = false) {
  if (forceNew) {
    throw new Error('Authentication required');
  }

  // Use TokenStorage which handles automatic refresh
  const accessToken = await tokenStorage.getValidAccessToken();
  if (!accessToken) {
    throw new Error('Authentication required');
  }

  return accessToken;
}

/**
 * Returns a valid Power Automate (Flow API) access token, or null if none can be obtained.
 * @returns {Promise<string|null>}
 */
async function getFlowAccessToken() {
  if (config.USE_TEST_MODE) {
    return tokenManager.getAccessToken(); // test_access_token_..., simulated by callFlowAPI
  }
  return tokenStorage.getValidFlowAccessToken();
}

module.exports = {
  tokenManager,
  getFlowAccessToken,
  authTools,
  ensureAuthenticated
};
