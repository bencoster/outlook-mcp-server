/**
 * Authentication Tools
 * Tools for managing OAuth2 authentication with Microsoft
 */

export const authTools = [
  {
        name: 'authenticate',
        description: 'Generate an OAuth2 authentication URL. The user must visit this URL to authorize the app, then run the auth-server to handle the callback. Usage: 1) Start auth server with `npm run auth-server`, 2) Visit the returned URL, 3) Sign in with Microsoft account.',
        inputSchema: {
                type: 'object',
                properties: {},
        },
  },
  {
        name: 'check-auth-status',
        description: 'Check whether the server currently has valid authentication tokens for Microsoft Graph API.',
        inputSchema: {
                type: 'object',
                properties: {},
        },
  },
  ];

export async function handleAuthTool(name, args, config) {
    switch (name) {
      case 'authenticate':
              return authenticate(config);
      case 'check-auth-status':
              return checkAuthStatus(config);
      default:
              return { content: [{ type: 'text', text: `Unknown auth tool: ${name}` }], isError: true };
    }
}

function authenticate(config) {
    if (!config.clientId) {
          return {
                  content: [{
                            type: 'text',
                            text: 'Error: No client ID configured. Set OUTLOOK_CLIENT_ID or MS_CLIENT_ID in your environment variables.',
                  }],
                  isError: true,
          };
    }

  const authUrl = new URL(`https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/authorize`);
    authUrl.searchParams.set('client_id', config.clientId);
    authUrl.searchParams.set('response_type', 'code');
    authUrl.searchParams.set('redirect_uri', config.redirectUri);
    authUrl.searchParams.set('scope', config.scopes.join(' '));
    authUrl.searchParams.set('response_mode', 'query');

  return {
        content: [{
                type: 'text',
                text: [
                          'To authenticate with Microsoft:',
                          '',
                          '1. First, start the auth callback server in a terminal:',
                          '   npm run auth-server',
                          '',
                          '2. Then visit this URL to sign in:',
                          `   ${authUrl.toString()}`,
                          '',
                          '3. After signing in, you will be redirected and tokens will be saved automatically.',
                          '',
                          'Once complete, all Outlook tools will be available.',
                        ].join('\n'),
        }],
  };
}

async function checkAuthStatus(config) {
    const fs = await import('fs');
    try {
          if (fs.existsSync(config.tokenStorePath)) {
                  const data = JSON.parse(fs.readFileSync(config.tokenStorePath, 'utf-8'));
                  const isExpired = data.expiresAt && data.expiresAt < Date.now();
                  const hasRefreshToken = !!data.refreshToken;

            return {
                      content: [{
                                  type: 'text',
                                  text: [
                                                'Authentication Status:',
                                                `  Token file: ${config.tokenStorePath}`,
                                                `  Access token: ${isExpired ? 'EXPIRED' : 'Valid'}`,
                                                `  Refresh token: ${hasRefreshToken ? 'Present (can auto-refresh)' : 'Missing'}`,
                                                `  Expires: ${data.expiresAt ? new Date(data.expiresAt).toISOString() : 'Unknown'}`,
                                                '',
                                                isExpired && hasRefreshToken
                                                  ? 'Token is expired but will auto-refresh on next API call.'
                                                  : isExpired
                                                  ? 'Token is expired. Please re-authenticate.'
                                                  : 'Ready to use.',
                                              ].join('\n'),
                      }],
            };
          }

      return {
              content: [{
                        type: 'text',
                        text: 'Not authenticated. No token file found. Please run the "authenticate" tool first.',
              }],
      };
    } catch (err) {
          return {
                  content: [{
                            type: 'text',
                            text: `Error checking auth status: ${err.message}`,
                  }],
                  isError: true,
          };
    }
}
