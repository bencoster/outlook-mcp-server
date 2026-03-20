#!/usr/bin/env node

/**
   * OAuth2 Authentication Server
   * Handles the OAuth callback from Microsoft to exchange the auth code for tokens.
   * Run with: npm run auth-server
   */

import express from 'express';
import { getConfig } from '../config.js';
import { storeTokens } from './graph-client.js';

const config = getConfig();
const app = express();
const PORT = 3333;

app.get('/auth/callback', async (req, res) => {
    const { code, error, error_description } = req.query;

          if (error) {
                res.status(400).send(`Authentication error: ${error_description || error}`);
                return;
          }

          if (!code) {
                res.status(400).send('No authorization code received.');
                return;
          }

          try {
                const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;

      const params = new URLSearchParams({
              client_id: config.clientId,
              client_secret: config.clientSecret,
              code: code,
              redirect_uri: config.redirectUri,
              grant_type: 'authorization_code',
              scope: config.scopes.join(' '),
      });

      const response = await fetch(tokenUrl, {
              method: 'POST',
              headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
              body: params.toString(),
      });

      if (!response.ok) {
              const errorText = await response.text();
              res.status(500).send(`Token exchange failed: ${errorText}`);
              return;
      }

      const data = await response.json();
                const tokens = {
                        accessToken: data.access_token,
                        refreshToken: data.refresh_token,
                        expiresAt: Date.now() + (data.expires_in * 1000),
                };

      storeTokens(config, tokens);

      res.send(`
            <html>
                    <body style="font-family: sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: #1a1a2e; color: #e0e0e0;">
                              <div style="text-align: center; padding: 40px; background: #16213e; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.3);">
                                          <h1 style="color: #0078d4;">Authentication Successful!</h1>
                                                      <p>Your Outlook MCP server is now connected.</p>
                                                                  <p>You can close this window and return to your AI assistant.</p>
                                                                            </div>
                                                                                    </body>
                                                                                          </html>
                                                                                              `);

      console.log('Authentication successful! Tokens stored.');
                setTimeout(() => process.exit(0), 2000);
          } catch (err) {
                res.status(500).send(`Error: ${err.message}`);
          }
});

app.listen(PORT, () => {
    console.log(`Auth callback server listening on http://localhost:${PORT}`);
    console.log('Waiting for OAuth callback...');
});
