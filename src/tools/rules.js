/**
 * Rules Tools
 * Tools for managing Outlook inbox rules via Microsoft Graph API
 */

export const ruleTools = [
  {
        name: 'list-rules',
        description: 'List all inbox rules (message rules) configured for the mailbox.',
        inputSchema: { type: 'object', properties: {} },
  },
  {
        name: 'get-rule',
        description: 'Get details of a specific inbox rule.',
        inputSchema: {
                type: 'object',
                properties: {
                          ruleId: { type: 'string', description: 'The rule ID.' },
                },
                required: ['ruleId'],
        },
  },
  {
        name: 'create-rule',
        description: 'Create a new inbox rule to automatically process incoming messages. Supports conditions like sender, subject contains, and actions like move, mark as read, forward.',
        inputSchema: {
                type: 'object',
                properties: {
                          displayName: { type: 'string', description: 'Name for the rule.' },
                          isEnabled: { type: 'boolean', description: 'Whether the rule is active. Defaults to true.' },
                          sequence: { type: 'number', description: 'Rule priority/order.' },
                          conditions: {
                                      type: 'object',
                                      description: 'Conditions that trigger the rule.',
                                      properties: {
                                                    senderContains: { type: 'array', items: { type: 'string' }, description: 'Sender address contains these strings.' },
                                                    subjectContains: { type: 'array', items: { type: 'string' }, description: 'Subject contains these strings.' },
                                                    bodyContains: { type: 'array', items: { type: 'string' }, description: 'Body contains these strings.' },
                                                    fromAddresses: { type: 'array', items: { type: 'string' }, description: 'From these specific email addresses.' },
                                                    hasAttachments: { type: 'boolean', description: 'Message has attachments.' },
                                                    importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Message importance level.' },
                                                    isReadReceipt: { type: 'boolean', description: 'Is a read receipt.' },
                                      },
                          },
                          actions: {
                                      type: 'object',
                                      description: 'Actions to perform when conditions are met.',
                                      properties: {
                                                    moveToFolder: { type: 'string', description: 'Folder ID to move message to.' },
                                                    copyToFolder: { type: 'string', description: 'Folder ID to copy message to.' },
                                                    delete: { type: 'boolean', description: 'Delete the message.' },
                                                    markAsRead: { type: 'boolean', description: 'Mark as read.' },
                                                    markImportance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Set importance level.' },
                                                    forwardTo: { type: 'array', items: { type: 'string' }, description: 'Email addresses to forward to.' },
                                                    stopProcessingRules: { type: 'boolean', description: 'Stop processing more rules.' },
                                      },
                          },
                },
                required: ['displayName', 'conditions', 'actions'],
        },
  },
  {
        name: 'update-rule',
        description: 'Update an existing inbox rule (enable/disable, change conditions or actions).',
        inputSchema: {
                type: 'object',
                properties: {
                          ruleId: { type: 'string', description: 'The rule ID to update.' },
                          displayName: { type: 'string', description: 'New name for the rule.' },
                          isEnabled: { type: 'boolean', description: 'Enable or disable the rule.' },
                },
                required: ['ruleId'],
        },
  },
  {
        name: 'delete-rule',
        description: 'Delete an inbox rule.',
        inputSchema: {
                type: 'object',
                properties: {
                          ruleId: { type: 'string', description: 'The rule ID to delete.' },
                },
                required: ['ruleId'],
        },
  },
  ];

export async function handleRuleTool(name, args, client) {
    switch (name) {
      case 'list-rules': return listRules(client);
      case 'get-rule': return getRule(args, client);
      case 'create-rule': return createRule(args, client);
      case 'update-rule': return updateRule(args, client);
      case 'delete-rule': return deleteRule(args, client);
      default:
              return { content: [{ type: 'text', text: `Unknown rule tool: ${name}` }], isError: true };
    }
}

function formatRule(rule) {
    const conditions = [];
    const c = rule.conditions || {};
    if (c.senderContains?.length) conditions.push(`Sender contains: ${c.senderContains.join(', ')}`);
    if (c.subjectContains?.length) conditions.push(`Subject contains: ${c.subjectContains.join(', ')}`);
    if (c.bodyContains?.length) conditions.push(`Body contains: ${c.bodyContains.join(', ')}`);
    if (c.fromAddresses?.length) conditions.push(`From: ${c.fromAddresses.map(a => a.emailAddress?.address).join(', ')}`);
    if (c.hasAttachments) conditions.push('Has attachments');
    if (c.importance) conditions.push(`Importance: ${c.importance}`);

  const actions = [];
    const a = rule.actions || {};
    if (a.moveToFolder) actions.push(`Move to folder: ${a.moveToFolder}`);
    if (a.copyToFolder) actions.push(`Copy to folder: ${a.copyToFolder}`);
    if (a.delete) actions.push('Delete');
    if (a.markAsRead) actions.push('Mark as read');
    if (a.markImportance) actions.push(`Set importance: ${a.markImportance}`);
    if (a.forwardTo?.length) actions.push(`Forward to: ${a.forwardTo.map(r => r.emailAddress?.address).join(', ')}`);
    if (a.stopProcessingRules) actions.push('Stop processing more rules');

  return [
        `  Name: ${rule.displayName}`,
        `  ID: ${rule.id}`,
        `  Enabled: ${rule.isEnabled ? 'Yes' : 'No'}`,
        `  Sequence: ${rule.sequence}`,
        `  Conditions: ${conditions.length ? conditions.join('; ') : 'None'}`,
        `  Actions: ${actions.length ? actions.join('; ') : 'None'}`,
        '  ---',
      ].join('\n');
}

async function listRules(client) {
    const result = await client.get('/me/mailFolders/inbox/messageRules');
    const rules = result.value || [];

  if (rules.length === 0) {
        return { content: [{ type: 'text', text: 'No inbox rules configured.' }] };
  }

  const formatted = rules.map(formatRule).join('\n');
    return { content: [{ type: 'text', text: `Inbox Rules (${rules.length}):\n\n${formatted}` }] };
}

async function getRule(args, client) {
    const rule = await client.get(`/me/mailFolders/inbox/messageRules/${args.ruleId}`);
    return { content: [{ type: 'text', text: `Rule Details:\n\n${formatRule(rule)}` }] };
}

async function createRule(args, client) {
    const rule = {
          displayName: args.displayName,
          isEnabled: args.isEnabled !== false,
          sequence: args.sequence || 1,
          conditions: {},
          actions: {},
    };

  // Build conditions
  const c = args.conditions;
    if (c.senderContains) rule.conditions.senderContains = c.senderContains;
    if (c.subjectContains) rule.conditions.subjectContains = c.subjectContains;
    if (c.bodyContains) rule.conditions.bodyContains = c.bodyContains;
    if (c.fromAddresses) {
          rule.conditions.fromAddresses = c.fromAddresses.map(addr => ({
                  emailAddress: { address: addr },
          }));
    }
    if (c.hasAttachments !== undefined) rule.conditions.hasAttachments = c.hasAttachments;
    if (c.importance) rule.conditions.importance = c.importance;

  // Build actions
  const a = args.actions;
    if (a.moveToFolder) rule.actions.moveToFolder = a.moveToFolder;
    if (a.copyToFolder) rule.actions.copyToFolder = a.copyToFolder;
    if (a.delete) rule.actions.delete = true;
    if (a.markAsRead) rule.actions.markAsRead = true;
    if (a.markImportance) rule.actions.markImportance = a.markImportance;
    if (a.forwardTo) {
          rule.actions.forwardTo = a.forwardTo.map(addr => ({
                  emailAddress: { address: addr },
          }));
    }
    if (a.stopProcessingRules) rule.actions.stopProcessingRules = true;

  const result = await client.post('/me/mailFolders/inbox/messageRules', rule);
    return {
          content: [{
                  type: 'text',
                  text: `Rule "${result.displayName}" created successfully.\nID: ${result.id}`,
          }],
    };
}

async function updateRule(args, client) {
    const update = {};
    if (args.displayName) update.displayName = args.displayName;
    if (args.isEnabled !== undefined) update.isEnabled = args.isEnabled;

  await client.patch(`/me/mailFolders/inbox/messageRules/${args.ruleId}`, update);
    return { content: [{ type: 'text', text: 'Rule updated successfully.' }] };
}

async function deleteRule(args, client) {
    await client.delete(`/me/mailFolders/inbox/messageRules/${args.ruleId}`);
    return { content: [{ type: 'text', text: 'Rule deleted successfully.' }] };
}
