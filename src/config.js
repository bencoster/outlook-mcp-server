/**
 * Configuration module
 * Reads settings from environment variables
 */

export function getConfig() {
    return {
          clientId: process.env.OUTLOOK_CLIENT_ID || process.env.MS_CLIENT_ID || '',
          clientSecret: process.env.OUTLOOK_CLIENT_SECRET || process.env.MS_CLIENT_SECRET || '',
          tenantId: process.env.MS_TENANT_ID || 'common',
          redirectUri: process.env.REDIRECT_URI || 'http://localhost:3333/auth/callback',
          tokenStorePath: process.env.TOKEN_STORE_PATH || getDefaultTokenPath(),
          scopes: [
                  'offline_access',
                  'User.Read',
                  'Mail.Read',
                  'Mail.ReadWrite',
                  'Mail.Send',
                  'Calendars.Read',
                  'Calendars.ReadWrite',
                  'Contacts.Read',
                  'Contacts.ReadWrite',
                  'MailboxSettings.ReadWrite',
                ],
    };
}

function getDefaultTokenPath() {
    const homeDir = process.env.HOME || process.env.USERPROFILE || '';
    return `${homeDir}/.outlook-mcp-tokens.json`;
}
