/**
 * Category Tools
 * Tools for managing Outlook master categories via Microsoft Graph API
 */

export const categoryTools = [
  {
        name: 'list-categories',
        description: 'List all master categories defined in the mailbox. Categories can be applied to emails, events, and contacts.',
        inputSchema: { type: 'object', properties: {} },
  },
  {
        name: 'create-category',
        description: 'Create a new master category with a name and color.',
        inputSchema: {
                type: 'object',
                properties: {
                          displayName: { type: 'string', description: 'Category name.' },
                          color: {
                                      type: 'string',
                                      description: 'Category color preset.',
                                      enum: ['none', 'preset0', 'preset1', 'preset2', 'preset3', 'preset4', 'preset5', 'preset6', 'preset7', 'preset8', 'preset9', 'preset10', 'preset11', 'preset12', 'preset13', 'preset14', 'preset15', 'preset16', 'preset17', 'preset18', 'preset19', 'preset20', 'preset21', 'preset22', 'preset23', 'preset24'],
                          },
                },
                required: ['displayName'],
        },
  },
  {
        name: 'update-category',
        description: 'Update a master category (change name or color).',
        inputSchema: {
                type: 'object',
                properties: {
                          categoryId: { type: 'string', description: 'The category ID to update.' },
                          displayName: { type: 'string', description: 'New category name.' },
                          color: { type: 'string', description: 'New color preset.' },
                },
                required: ['categoryId'],
        },
  },
  {
        name: 'delete-category',
        description: 'Delete a master category.',
        inputSchema: {
                type: 'object',
                properties: {
                          categoryId: { type: 'string', description: 'The category ID to delete.' },
                },
                required: ['categoryId'],
        },
  },
  {
        name: 'categorize-message',
        description: 'Apply one or more categories to an email message.',
        inputSchema: {
                type: 'object',
                properties: {
                          messageId: { type: 'string', description: 'The email message ID.' },
                          categories: { type: 'array', items: { type: 'string' }, description: 'Array of category names to apply.' },
                },
                required: ['messageId', 'categories'],
        },
  },
  {
        name: 'categorize-event',
        description: 'Apply one or more categories to a calendar event.',
        inputSchema: {
                type: 'object',
                properties: {
                          eventId: { type: 'string', description: 'The event ID.' },
                          categories: { type: 'array', items: { type: 'string' }, description: 'Array of category names to apply.' },
                },
                required: ['eventId', 'categories'],
        },
  },
  ];

const COLOR_MAP = {
    none: 'None', preset0: 'Red', preset1: 'Orange', preset2: 'Brown',
    preset3: 'Yellow', preset4: 'Green', preset5: 'Teal', preset6: 'Olive',
    preset7: 'Blue', preset8: 'Purple', preset9: 'Cranberry', preset10: 'Steel',
    preset11: 'DarkSteel', preset12: 'Gray', preset13: 'DarkGray',
    preset14: 'Black', preset15: 'DarkRed', preset16: 'DarkOrange',
    preset17: 'DarkBrown', preset18: 'DarkYellow', preset19: 'DarkGreen',
    preset20: 'DarkTeal', preset21: 'DarkOlive', preset22: 'DarkBlue',
    preset23: 'DarkPurple', preset24: 'DarkCranberry',
};

export async function handleCategoryTool(name, args, client) {
    switch (name) {
      case 'list-categories': return listCategories(client);
      case 'create-category': return createCategory(args, client);
      case 'update-category': return updateCategory(args, client);
      case 'delete-category': return deleteCategory(args, client);
      case 'categorize-message': return categorizeMessage(args, client);
      case 'categorize-event': return categorizeEvent(args, client);
      default:
              return { content: [{ type: 'text', text: `Unknown category tool: ${name}` }], isError: true };
    }
}

async function listCategories(client) {
    const result = await client.get('/me/outlook/masterCategories');
    const categories = result.value || [];

  if (categories.length === 0) {
        return { content: [{ type: 'text', text: 'No categories defined.' }] };
  }

  const formatted = categories.map(c => [
        `  Name: ${c.displayName}`,
        `  ID: ${c.id}`,
        `  Color: ${c.color} (${COLOR_MAP[c.color] || c.color})`,
        '  ---',
      ].join('\n')).join('\n');

  return { content: [{ type: 'text', text: `Categories (${categories.length}):\n\n${formatted}` }] };
}

async function createCategory(args, client) {
    const body = {
          displayName: args.displayName,
          color: args.color || 'none',
    };

  const result = await client.post('/me/outlook/masterCategories', body);
    return {
          content: [{
                  type: 'text',
                  text: `Category "${result.displayName}" created (${COLOR_MAP[result.color] || result.color}).\nID: ${result.id}`,
          }],
    };
}

async function updateCategory(args, client) {
    const update = {};
    if (args.displayName) update.displayName = args.displayName;
    if (args.color) update.color = args.color;

  await client.patch(`/me/outlook/masterCategories/${args.categoryId}`, update);
    return { content: [{ type: 'text', text: 'Category updated successfully.' }] };
}

async function deleteCategory(args, client) {
    await client.delete(`/me/outlook/masterCategories/${args.categoryId}`);
    return { content: [{ type: 'text', text: 'Category deleted.' }] };
}

async function categorizeMessage(args, client) {
    await client.patch(`/me/messages/${args.messageId}`, {
          categories: args.categories,
    });
    return {
          content: [{
                  type: 'text',
                  text: `Applied categories [${args.categories.join(', ')}] to message.`,
          }],
    };
}

async function categorizeEvent(args, client) {
    await client.patch(`/me/events/${args.eventId}`, {
          categories: args.categories,
    });
    return {
          content: [{
                  type: 'text',
                  text: `Applied categories [${args.categories.join(', ')}] to event.`,
          }],
    };
}
