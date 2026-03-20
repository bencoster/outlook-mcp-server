# Outlook MCP Server

A comprehensive **Model Context Protocol (MCP)** server that connects AI assistants like **Claude, Cursor, and others** directly to your Microsoft Outlook account via the **Microsoft Graph API**.

It provides structured and secure access to your **email, calendar, contacts, folders, attachments, rules, and categories** so your AI agent can manage your Outlook on your behalf.

## Features

- **Email Management** - List, search, read, send, reply, forward, and organize emails
- - **Folder Organization** - Create, rename, delete, and move mail folders to keep your inbox tidy
  - - **Calendar & Events** - List calendars, create events with attendees, recurrence, Teams meetings, accept/decline invitations
    - - **Smart Attachments** - List, download, and add attachments to emails; create drafts with attachments
      - - **Inbox Rules & Filtering** - Create automated rules to filter, sort, forward, or act on messages as they arrive
        - - **Contact Management** - List, create, update, and organize contacts and contact folders
          - - **Category Organization** - Create master categories and apply them to emails and events
           
            - ## Available Tools (35+)
           
            - ### Authentication
            - | Tool | Description |
            - |------|-------------|
            - | `authenticate` | Generate OAuth2 URL to connect your Microsoft account |
            - | `check-auth-status` | Check if authentication tokens are valid |
           
            - ### Email
            - | Tool | Description |
            - |------|-------------|
            - | `list-emails` | List recent emails from inbox or specific folder |
            - | `search-emails` | Search emails by keyword |
            - | `read-email` | Read full email content with attachments |
            - | `send-email` | Send email with TO/CC/BCC, HTML body, importance |
            - | `reply-email` | Reply or reply-all to an email |
            - | `forward-email` | Forward email to new recipients |
            - | `mark-as-read` | Mark emails as read/unread |
            - | `move-email` | Move email between folders |
            - | `delete-email` | Delete an email |
           
            - ### Calendar
            - | Tool | Description |
            - |------|-------------|
            - | `list-calendars` | List all calendars |
            - | `create-calendar` | Create a new calendar |
            - | `list-events` | List events in a date range |
            - | `create-event` | Create event with attendees, recurrence, Teams |
            - | `update-event` | Update an existing event |
            - | `delete-event` | Delete an event |
            - | `accept-event` | Accept an event invitation |
            - | `decline-event` | Decline an event invitation |
           
            - ### Folders
            - | Tool | Description |
            - |------|-------------|
            - | `list-folders` | List mail folders with unread counts |
            - | `create-folder` | Create a new mail folder |
            - | `rename-folder` | Rename a folder |
            - | `delete-folder` | Delete a folder |
            - | `move-folder` | Move folder to another parent |
           
            - ### Contacts
            - | Tool | Description |
            - |------|-------------|
            - | `list-contacts` | List/search contacts |
            - | `get-contact` | Get contact details |
            - | `create-contact` | Create a new contact |
            - | `update-contact` | Update contact info |
            - | `delete-contact` | Delete a contact |
            - | `list-contact-folders` | List contact folders |
            - | `create-contact-folder` | Create contact folder |
           
            - ### Attachments
            - | Tool | Description |
            - |------|-------------|
            - | `list-attachments` | List attachments on a message |
            - | `get-attachment` | Get attachment content |
            - | `add-attachment` | Add attachment to a draft |
            - | `delete-attachment` | Remove attachment from draft |
            - | `create-draft-with-attachment` | Create draft with attachment |
           
            - ### Rules
            - | Tool | Description |
            - |------|-------------|
            - | `list-rules` | List inbox rules |
            - | `get-rule` | Get rule details |
            - | `create-rule` | Create inbox rule with conditions/actions |
            - | `update-rule` | Enable/disable or rename a rule |
            - | `delete-rule` | Delete a rule |
           
            - ### Categories
            - | Tool | Description |
            - |------|-------------|
            - | `list-categories` | List master categories |
            - | `create-category` | Create category with color |
            - | `update-category` | Update category |
            - | `delete-category` | Delete category |
            - | `categorize-message` | Apply categories to email |
            - | `categorize-event` | Apply categories to event |
           
            - ## Quick Start
           
            - ### 1. Install Dependencies
           
            - ```bash
              git clone https://github.com/bencoster/outlook-mcp-server.git
              cd outlook-mcp-server
              npm install
              ```

              ### 2. Azure App Registration

              1. Go to [Azure Portal](https://portal.azure.com) > **App registrations** > **New registration**
              2. 2. Name: `Outlook MCP Server`
                 3. 3. Account type: **Accounts in any organizational directory and personal Microsoft accounts**
                    4. 4. Redirect URI: **Web** > `http://localhost:3333/auth/callback`
                       5. 5. Click **Register**
                          6. 6. Copy the **Application (client) ID**
                            
                             7. **API Permissions** (under Manage > API permissions):
                             8. - `offline_access`
                                - - `User.Read`
                                  - - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
                                    - - `Calendars.Read`, `Calendars.ReadWrite`
                                      - - `Contacts.Read`, `Contacts.ReadWrite`
                                        - - `MailboxSettings.ReadWrite`
                                         
                                          - **Client Secret** (under Certificates & secrets):
                                          - - Click **New client secret** and copy the **VALUE** (not the Secret ID)
                                           
                                            - ### 3. Configure Environment
                                           
                                            - ```bash
                                              cp .env.example .env
                                              ```

                                              Edit `.env` with your Azure credentials:
                                              ```
                                              OUTLOOK_CLIENT_ID=your-client-id
                                              OUTLOOK_CLIENT_SECRET=your-client-secret-value
                                              MS_TENANT_ID=common
                                              ```

                                              ### 4. Configure Your AI Assistant

                                              #### Claude Desktop
                                              Add to your Claude Desktop config (`claude_desktop_config.json`):

                                              ```json
                                              {
                                                "mcpServers": {
                                                  "outlook": {
                                                    "command": "node",
                                                    "args": ["/path/to/outlook-mcp-server/src/index.js"],
                                                    "env": {
                                                      "OUTLOOK_CLIENT_ID": "your-client-id",
                                                      "OUTLOOK_CLIENT_SECRET": "your-client-secret"
                                                    }
                                                  }
                                                }
                                              }
                                              ```

                                              #### Claude Code (CLI)
                                              ```bash
                                              claude mcp add outlook node /path/to/outlook-mcp-server/src/index.js
                                              ```

                                              #### Cursor
                                              Add to `.cursor/mcp.json`:
                                              ```json
                                              {
                                                "mcpServers": {
                                                  "outlook": {
                                                    "command": "node",
                                                    "args": ["/path/to/outlook-mcp-server/src/index.js"],
                                                    "env": {
                                                      "OUTLOOK_CLIENT_ID": "your-client-id",
                                                      "OUTLOOK_CLIENT_SECRET": "your-client-secret"
                                                    }
                                                  }
                                                }
                                              }
                                              ```

                                              ### 5. Authenticate

                                              1. Start the auth callback server:
                                              2.    ```bash
                                                       npm run auth-server
                                                       ```
                                                    2. Use the `authenticate` tool in your AI assistant to get the OAuth URL
                                                    3. 3. Visit the URL and sign in with your Microsoft account
                                                       4. 4. Tokens are saved automatically to `~/.outlook-mcp-tokens.json`
                                                         
                                                          5. ## Project Structure
                                                         
                                                          6. ```
                                                             outlook-mcp-server/
                                                             ├── package.json
                                                             ├── .env.example
                                                             ├── src/
                                                             │   ├── index.js              # Main MCP server entry point
                                                             │   ├── config.js              # Configuration from environment
                                                             │   ├── auth/
                                                             │   │   ├── graph-client.js    # Graph API client with token management
                                                             │   │   └── auth-server.js     # OAuth callback server
                                                             │   └── tools/
                                                             │       ├── auth-tools.js      # Authentication tools
                                                             │       ├── email.js           # Email management tools
                                                             │       ├── calendar.js        # Calendar & event tools
                                                             │       ├── folders.js         # Mail folder tools
                                                             │       ├── contacts.js        # Contact management tools
                                                             │       ├── attachments.js     # Attachment handling tools
                                                             │       ├── rules.js           # Inbox rule tools
                                                             │       └── categories.js      # Category management tools
                                                             ```

                                                             ## Development

                                                             ```bash
                                                             # Run with auto-reload
                                                             npm run dev

                                                             # Test with MCP Inspector
                                                             npm run inspect
                                                             ```

                                                             ## Troubleshooting

                                                             **"Not authenticated"** - Run `npm run auth-server` and use the `authenticate` tool

                                                             **"Invalid client secret" (AADSTS7000215)** - Use the secret **VALUE**, not the Secret ID

                                                             **"Port 3333 in use"** - Run `npx kill-port 3333` then retry

                                                             **Token expired** - Tokens auto-refresh. If issues persist, delete `~/.outlook-mcp-tokens.json` and re-authenticate

                                                             ## License

                                                             MIT
