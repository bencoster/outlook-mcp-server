#!/usr/bin/env node

/**
   * Outlook MCP Server - Main Entry Point
   * 
   * A comprehensive MCP server that connects AI assistants to Microsoft Outlook
   * via the Microsoft Graph API. Supports email, calendar, contacts, folders,
   * attachments, rules, and categories.
   */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
    CallToolRequestSchema,
    ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';

import { getConfig } from './config.js';
import { getGraphClient } from './auth/graph-client.js';

// Import tool definitions and handlers
import { emailTools, handleEmailTool } from './tools/email.js';
import { calendarTools, handleCalendarTool } from './tools/calendar.js';
import { folderTools, handleFolderTool } from './tools/folders.js';
import { contactTools, handleContactTool } from './tools/contacts.js';
import { attachmentTools, handleAttachmentTool } from './tools/attachments.js';
import { ruleTools, handleRuleTool } from './tools/rules.js';
import { categoryTools, handleCategoryTool } from './tools/categories.js';
import { authTools, handleAuthTool } from './tools/auth-tools.js';

const config = getConfig();

// Collect all tool definitions
const ALL_TOOLS = [
    ...authTools,
    ...emailTools,
    ...calendarTools,
    ...folderTools,
    ...contactTools,
    ...attachmentTools,
    ...ruleTools,
    ...categoryTools,
  ];

// Map tool names to their handler modules
function getHandler(toolName) {
    if (authTools.some(t => t.name === toolName)) return handleAuthTool;
    if (emailTools.some(t => t.name === toolName)) return handleEmailTool;
    if (calendarTools.some(t => t.name === toolName)) return handleCalendarTool;
    if (folderTools.some(t => t.name === toolName)) return handleFolderTool;
    if (contactTools.some(t => t.name === toolName)) return handleContactTool;
    if (attachmentTools.some(t => t.name === toolName)) return handleAttachmentTool;
    if (ruleTools.some(t => t.name === toolName)) return handleRuleTool;
    if (categoryTools.some(t => t.name === toolName)) return handleCategoryTool;
    return null;
}

// Create and configure the MCP server
const server = new Server(
  {
        name: 'outlook-mcp-server',
        version: '1.0.0',
  },
  {
        capabilities: {
                tools: {},
        },
  }
  );

// List available tools
server.setRequestHandler(ListToolsRequestSchema, async () => {
    return { tools: ALL_TOOLS };
});

// Handle tool calls
server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

                           try {
                                 const handler = getHandler(name);
                                 if (!handler) {
                                         return {
                                                   content: [{ type: 'text', text: `Unknown tool: ${name}` }],
                                                   isError: true,
                                         };
                                 }

      // Auth tools don't need a graph client
      if (authTools.some(t => t.name === name)) {
              return await handler(name, args, config);
      }

      // All other tools need an authenticated graph client
      const client = await getGraphClient(config);
                                 if (!client) {
                                         return {
                                                   content: [{
                                                               type: 'text',
                                                               text: 'Not authenticated. Please run the "authenticate" tool first to connect your Microsoft account.',
                                                   }],
                                                   isError: true,
                                         };
                                 }

      return await handler(name, args, client);
                           } catch (error) {
                                 return {
                                         content: [{
                                                   type: 'text',
                                                   text: `Error executing ${name}: ${error.message}`,
                                         }],
                                         isError: true,
                                 };
                           }
});

// Start the server
async function main() {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error('Outlook MCP Server running on stdio');
}

main().catch((error) => {
    console.error('Fatal error:', error);
    process.exit(1);
});
