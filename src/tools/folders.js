/**
 * Folder Tools
 * Tools for managing Outlook mail folders via Microsoft Graph API
 */

export const folderTools = [
  {
        name: 'list-folders',
        description: 'List all mail folders including child folders. Returns folder names, IDs, unread counts, and total counts.',
        inputSchema: {
                type: 'object',
                properties: {
                          parentFolderId: { type: 'string', description: 'Parent folder ID to list children of. If omitted, lists top-level folders.' },
                },
        },
  },
  {
        name: 'create-folder',
        description: 'Create a new mail folder. Can be top-level or nested under a parent folder.',
        inputSchema: {
                type: 'object',
                properties: {
                          displayName: { type: 'string', description: 'Name for the new folder.' },
                          parentFolderId: { type: 'string', description: 'Parent folder ID. If omitted, creates at top level.' },
                          isHidden: { type: 'boolean', description: 'Whether the folder should be hidden. Defaults to false.' },
                },
                required: ['displayName'],
        },
  },
  {
        name: 'rename-folder',
        description: 'Rename an existing mail folder.',
        inputSchema: {
                type: 'object',
                properties: {
                          folderId: { type: 'string', description: 'The folder ID to rename.' },
                          displayName: { type: 'string', description: 'New name for the folder.' },
                },
                required: ['folderId', 'displayName'],
        },
  },
  {
        name: 'delete-folder',
        description: 'Delete a mail folder and all its contents.',
        inputSchema: {
                type: 'object',
                properties: {
                          folderId: { type: 'string', description: 'The folder ID to delete.' },
                },
                required: ['folderId'],
        },
  },
  {
        name: 'move-folder',
        description: 'Move a mail folder to be a child of another folder.',
        inputSchema: {
                type: 'object',
                properties: {
                          folderId: { type: 'string', description: 'The folder ID to move.' },
                          destinationFolderId: { type: 'string', description: 'The destination parent folder ID.' },
                },
                required: ['folderId', 'destinationFolderId'],
        },
  },
  ];

export async function handleFolderTool(name, args, client) {
    switch (name) {
      case 'list-folders': return listFolders(args, client);
      case 'create-folder': return createFolder(args, client);
      case 'rename-folder': return renameFolder(args, client);
      case 'delete-folder': return deleteFolder(args, client);
      case 'move-folder': return moveFolder(args, client);
      default:
              return { content: [{ type: 'text', text: `Unknown folder tool: ${name}` }], isError: true };
    }
}

async function listFolders(args, client) {
    const endpoint = args.parentFolderId
      ? `/me/mailFolders/${args.parentFolderId}/childFolders?$select=id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount`
          : '/me/mailFolders?$select=id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount&$top=100';

  const result = await client.get(endpoint);
    const folders = result.value || [];

  if (folders.length === 0) {
        return { content: [{ type: 'text', text: 'No folders found.' }] };
  }

  const formatted = folders.map(f => [
        `  Name: ${f.displayName}`,
        `  ID: ${f.id}`,
        `  Total items: ${f.totalItemCount}`,
        `  Unread: ${f.unreadItemCount}`,
        `  Child folders: ${f.childFolderCount}`,
        '  ---',
      ].join('\n')).join('\n');

  return { content: [{ type: 'text', text: `Mail Folders (${folders.length}):\n\n${formatted}` }] };
}

async function createFolder(args, client) {
    const body = { displayName: args.displayName };
    if (args.isHidden) body.isHidden = true;

  const endpoint = args.parentFolderId
      ? `/me/mailFolders/${args.parentFolderId}/childFolders`
        : '/me/mailFolders';

  const result = await client.post(endpoint, body);
    return {
          content: [{
                  type: 'text',
                  text: `Folder "${result.displayName}" created successfully.\nID: ${result.id}`,
          }],
    };
}

async function renameFolder(args, client) {
    await client.patch(`/me/mailFolders/${args.folderId}`, {
          displayName: args.displayName,
    });
    return { content: [{ type: 'text', text: `Folder renamed to "${args.displayName}".` }] };
}

async function deleteFolder(args, client) {
    await client.delete(`/me/mailFolders/${args.folderId}`);
    return { content: [{ type: 'text', text: 'Folder deleted successfully.' }] };
}

async function moveFolder(args, client) {
    await client.post(`/me/mailFolders/${args.folderId}/move`, {
          destinationId: args.destinationFolderId,
    });
    return { content: [{ type: 'text', text: 'Folder moved successfully.' }] };
}
