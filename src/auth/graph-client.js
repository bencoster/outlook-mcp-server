/**
 * Microsoft Graph API Client
 * Manages OAuth tokens and provides an authenticated fetch wrapper
 */

import fs from 'fs';
import path from 'path';

let cachedTokens = null;

/**
 * Load stored tokens from disk
 */
function loadTokens(tokenPath) {
    try {
          if (fs.existsSync(tokenPath)) {
                  const data = fs.readFileSync(tokenPath, 'utf-8');
                  return JSON.parse(data);
          }
    } catch (err) {
          console.error('Failed to load tokens:', err.message);
    }
    return null;
}

/**
 * Save tokens to disk
 */
function saveTokens(tokenPath, tokens) {
    try {
          const dir = path.dirname(tokenPath);
          if (!fs.existsSync(dir)) {
                  fs.mkdirSync(dir, { recursive: true });
          }
          fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2));
    } catch (err) {
          console.error('Failed to save tokens:', err.message);
    }
}

/**
 * Refresh the access token using the refresh token
 */
async function refreshAccessToken(config, refreshToken) {
    const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;

  const params = new URLSearchParams({
        client_id: config.clientId,
        client_secret: config.clientSecret,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
        scope: config.scopes.join(' '),
  });

  const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: params.toString(),
  });

  if (!response.ok) {
        const errorData = await response.text();
        throw new Error(`Token refresh failed: ${errorData}`);
  }

  const data = await response.json();
    return {
          accessToken: data.access_token,
          refreshToken: data.refresh_token || refreshToken,
          expiresAt: Date.now() + (data.expires_in * 1000),
    };
}

/**
 * Get a valid access token, refreshing if needed
 */
async function getAccessToken(config) {
    if (!cachedTokens) {
          cachedTokens = loadTokens(config.tokenStorePath);
    }

  if (!cachedTokens) {
        return null;
  }

  // If token expires within 5 minutes, refresh it
  if (cachedTokens.expiresAt && cachedTokens.expiresAt < Date.now() + 300000) {
        try {
                cachedTokens = await refreshAccessToken(config, cachedTokens.refreshToken);
                saveTokens(config.tokenStorePath, cachedTokens);
        } catch (err) {
                console.error('Token refresh failed:', err.message);
                return null;
        }
  }

  return cachedTokens.accessToken;
}

/**
 * Create a Graph API client wrapper
 */
export async function getGraphClient(config) {
    const accessToken = await getAccessToken(config);
    if (!accessToken) return null;

  return {
        /**
               * Make an authenticated request to the Microsoft Graph API
         */
        async api(endpoint, options = {}) {
                const url = endpoint.startsWith('http')
                  ? endpoint
                          : `https://graph.microsoft.com/v1.0${endpoint}`;

          const response = await fetch(url, {
                    ...options,
                    headers: {
                                Authorization: `Bearer ${accessToken}`,
                                'Content-Type': 'application/json',
                                ...options.headers,
                    },
          });

          if (!response.ok) {
                    const errorText = await response.text();
                    let errorMessage;
                    try {
                                const errorJson = JSON.parse(errorText);
                                errorMessage = errorJson.error?.message || errorText;
                    } catch {
                                errorMessage = errorText;
                    }
                    throw new Error(`Graph API error (${response.status}): ${errorMessage}`);
          }

          // Some endpoints return no content
          if (response.status === 204) return null;

          const contentType = response.headers.get('content-type');
                if (contentType && contentType.includes('application/json')) {
                          return response.json();
                }

          return response.text();
        },

        /**
               * GET request shorthand
         */
        async get(endpoint) {
                return this.api(endpoint);
        },

        /**
               * POST request shorthand
         */
        async post(endpoint, body) {
                return this.api(endpoint, {
                          method: 'POST',
                          body: JSON.stringify(body),
                });
        },

        /**
               * PATCH request shorthand
         */
        async patch(endpoint, body) {
                return this.api(endpoint, {
                          method: 'PATCH',
                          body: JSON.stringify(body),
                });
        },

        /**
               * DELETE request shorthand
         */
        async delete(endpoint) {
                return this.api(endpoint, { method: 'DELETE' });
        },
  };
}

/**
 * Store tokens after OAuth callback
 */
export function storeTokens(config, tokens) {
    cachedTokens = tokens;
    saveTokens(config.tokenStorePath, tokens);
}
