/**
 * Email Tools
 * Tools for managing Outlook emails via Microsoft Graph API
 */

export const emailTools = [
  {
        name: 'list-emails',
        description: 'List recent emails from your inbox or a specific folder. Returns subject, sender, date, and preview.',
        inputSchema: {
                type: 'object',
                properties: {
                          folderId: { type: 'string', description: 'Mail folder ID or well-known name (e.g., "inbox", "drafts", "sentitems", "deleteditems"). Defaults to inbox.' },
                          count: { type: 'number', description: 'Number of emails to return (max 50, default 10).' },
                          skip: { type: 'number', description: 'Number of emails to skip for pagination.' },
                          filter: { type: 'string', description: 'OData filter expression (e.g., "isRead eq false").' },
                },
        },
  },
  {
        name: 'search-emails',
        description: 'Search emails using a keyword query. Searches across subject, body, and sender fields.',
        inputSchema: {
                type: 'object',
                properties: {
                          query: { type: 'string', description: 'Search query string.' },
                          count: { type: 'number', description: 'Number of results to return (max 50, default 10).' },
                },
                required: ['query'],
        },
  },
  {
        name: 'read-email',
        description: 'Read the full content of a specific email by its ID. Returns subject, body, sender, recipients, and attachments info.',
        inputSchema: {
                type: 'object',
                properties: {
                          messageId: { type: 'string', description: 'The ID of the email message to read.' },
                },
                required: ['messageId'],
        },
  },
  {
        name: 'send-email',
        description: 'Send a new email. Supports TO, CC, BCC recipients, HTML or text body, and importance level.',
        inputSchema: {
                type: 'object',
                properties: {
                          to: {
                                      type: 'array',
                                      items: { type: 'string' },
                                      description: 'Array of recipient email addresses.',
                          },
                          cc: {
                                      type: 'array',
                                      items: { type: 'string' },
                                      description: 'Array of CC email addresses.',
                          },
                          bcc: {
                                      type: 'array',
                                      items: { type: 'string' },
                                      description: 'Array of BCC email addresses.',
                          },
                          subject: { type: 'string', description: 'Email subject line.' },
                          body: { type: 'string', description: 'Email body content.' },
                          bodyType: { type: 'string', enum: ['text', 'html'], description: 'Body content type. Defaults to "html".' },
                          importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Email importance level.' },
                          saveToSentItems: { type: 'boolean', description: 'Save to Sent Items folder. Defaults to true.' },
                },
                required: ['to', 'subject', 'body'],
        },
  },
  {
        name: 'reply-email',
        description: 'Reply to an existing email. Can reply to sender only or reply-all.',
        inputSchema: {
                type: 'object',
                properties: {
                          messageId: { type: 'string', description: 'The ID of the email to reply to.' },
                          body: { type: 'string', description: 'Reply body content.' },
                          replyAll: { type: 'boolean', description: 'If true, reply to all recipients. Defaults to false.' },
                },
                required: ['messageId', 'body'],
        },
  },
  {
        name: 'forward-email',
        description: 'Forward an existing email to new recipients.',
        inputSchema: {
                type: 'object',
                properties: {
                          messageId: { type: 'string', description: 'The ID of the email to forward.' },
                          to: {
                                      type: 'array',
                                      items: { type: 'string' },
                                      description: 'Array of recipient email addresses to forward to.',
                          },
                          comment: { type: 'string', description: 'Optional comment to include with the forwarded email.' },
                },
                required: ['messageId', 'to'],
        },
  },
  {
        name: 'mark-as-read',
        description: 'Mark one or more emails as read or unread.',
        inputSchema: {
                type: 'object',
                properties: {
                          messageIds: {
                                      type: 'array',
                                      items: { type: 'string' },
                                      description: 'Array of message IDs to update.',
                          },
                          isRead: { type: 'boolean', description: 'Set to true to mark as read, false for unread.' },
                },
                required: ['messageIds', 'isRead'],
        },
  },
  {
        name: 'move-email',
        description: 'Move an email to a different folder.',
        inputSchema: {
                type: 'object',
                properties: {
                          messageId: { type: 'string', description: 'The ID of the email to move.' },
                          destinationFolderId: { type: 'string', description: 'Destination folder ID or well-known name.' },
                },
                required: ['messageId', 'destinationFolderId'],
        },
  },
  {
        name: 'delete-email',
        description: 'Delete an email (moves to Deleted Items).',
        inputSchema: {
                type: 'object',
                properties: {
                          messageId: { type: 'string', description: 'The ID of the email to delete.' },
                },
                required: ['messageId'],
        },
  },
  ];

export async function handleEmailTool(name, args, client) {
    switch (name) {
      case 'list-emails': return listEmails(args, client);
      case 'search-emails': return searchEmails(args, client);
      case 'read-email': return readEmail(args, client);
      case 'send-email': return sendEmail(args, client);
      case 'reply-email': return replyEmail(args, client);
      case 'forward-email': return forwardEmail(args, client);
      case 'mark-as-read': return markAsRead(args, client);
      case 'move-email': return moveEmail(args, client);
      case 'delete-email': return deleteEmail(args, client);
      default:
              return { content: [{ type: 'text', text: `Unknown email tool: ${name}` }], isError: true };
    }
}

function formatRecipient(email) {
    return { emailAddress: { address: email } };
}

function formatEmailSummary(msg) {
    return [
          `ID: ${msg.id}`,
          `Subject: ${msg.subject || '(no subject)'}`,
          `From: ${msg.from?.emailAddress?.name || ''} <${msg.from?.emailAddress?.address || 'unknown'}>`,
          `Date: ${msg.receivedDateTime}`,
          `Read: ${msg.isRead ? 'Yes' : 'No'}`,
          `Preview: ${msg.bodyPreview?.substring(0, 150) || ''}`,
          '---',
        ].join('\n');
}

async function listEmails(args, client) {
    const folder = args.folderId || 'inbox';
    const count = Math.min(args.count || 10, 50);
    const skip = args.skip || 0;

  let endpoint = `/me/mailFolders/${folder}/messages?$top=${count}&$skip=${skip}&$orderby=receivedDateTime desc&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments,importance`;

  if (args.filter) {
        endpoint += `&$filter=${encodeURIComponent(args.filter)}`;
  }

  const result = await client.get(endpoint);
    const emails = result.value || [];

  if (emails.length === 0) {
        return { content: [{ type: 'text', text: 'No emails found.' }] };
  }

  const formatted = emails.map(formatEmailSummary).join('\n');
    return {
          content: [{
                  type: 'text',
                  text: `Found ${emails.length} email(s):\n\n${formatted}`,
          }],
    };
}

async function searchEmails(args, client) {
    const count = Math.min(args.count || 10, 50);
    const endpoint = `/me/messages?$search="${encodeURIComponent(args.query)}"&$top=${count}&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments`;

  const result = await client.get(endpoint);
    const emails = result.value || [];

  if (emails.length === 0) {
        return { content: [{ type: 'text', text: `No emails found matching "${args.query}".` }] };
  }

  const formatted = emails.map(formatEmailSummary).join('\n');
    return {
          content: [{
                  type: 'text',
                  text: `Search results for "${args.query}" (${emails.length} found):\n\n${formatted}`,
          }],
    };
}

async function readEmail(args, client) {
    const msg = await client.get(`/me/messages/${args.messageId}?$expand=attachments($select=id,name,contentType,size)`);

  const attachments = msg.attachments?.map(a =>
        `  - ${a.name} (${a.contentType}, ${(a.size / 1024).toFixed(1)} KB, ID: ${a.id})`
                                             ).join('\n') || '  None';

  const recipients = msg.toRecipients?.map(r => `${r.emailAddress.name || ''} <${r.emailAddress.address}>`).join(', ') || 'None';
    const cc = msg.ccRecipients?.map(r => `${r.emailAddress.name || ''} <${r.emailAddress.address}>`).join(', ') || 'None';

  const text = [
        `Subject: ${msg.subject}`,
        `From: ${msg.from?.emailAddress?.name || ''} <${msg.from?.emailAddress?.address}>`,
        `To: ${recipients}`,
        `CC: ${cc}`,
        `Date: ${msg.receivedDateTime}`,
        `Importance: ${msg.importance}`,
        `Read: ${msg.isRead ? 'Yes' : 'No'}`,
        `Has Attachments: ${msg.hasAttachments ? 'Yes' : 'No'}`,
        `\nAttachments:\n${attachments}`,
        `\nBody:\n${msg.body?.content || '(empty)'}`,
      ].join('\n');

  return { content: [{ type: 'text', text }] };
}

async function sendEmail(args, client) {
    const message = {
          subject: args.subject,
          body: {
                  contentType: args.bodyType === 'text' ? 'text' : 'html',
                  content: args.body,
          },
          toRecipients: args.to.map(formatRecipient),
          importance: args.importance || 'normal',
    };

  if (args.cc) message.ccRecipients = args.cc.map(formatRecipient);
    if (args.bcc) message.bccRecipients = args.bcc.map(formatRecipient);

  await client.post('/me/sendMail', {
        message,
        saveToSentItems: args.saveToSentItems !== false,
  });

  return {
        content: [{
                type: 'text',
                text: `Email sent successfully to ${args.to.join(', ')}`,
        }],
  };
}

async function replyEmail(args, client) {
    const endpoint = args.replyAll
      ? `/me/messages/${args.messageId}/replyAll`
          : `/me/messages/${args.messageId}/reply`;

  await client.post(endpoint, { comment: args.body });

  return {
        content: [{
                type: 'text',
                text: `Reply${args.replyAll ? ' all' : ''} sent successfully.`,
        }],
  };
}

async function forwardEmail(args, client) {
    await client.post(`/me/messages/${args.messageId}/forward`, {
          comment: args.comment || '',
          toRecipients: args.to.map(formatRecipient),
    });

  return {
        content: [{
                type: 'text',
                text: `Email forwarded successfully to ${args.to.join(', ')}`,
        }],
  };
}

async function markAsRead(args, client) {
    const promises = args.messageIds.map(id =>
          client.patch(`/me/messages/${id}`, { isRead: args.isRead })
                                           );

  await Promise.all(promises);

  return {
        content: [{
                type: 'text',
                text: `Marked ${args.messageIds.length} email(s) as ${args.isRead ? 'read' : 'unread'}.`,
        }],
  };
}

async function moveEmail(args, client) {
    await client.post(`/me/messages/${args.messageId}/move`, {
          destinationId: args.destinationFolderId,
    });

  return {
        content: [{
                type: 'text',
                text: `Email moved to folder ${args.destinationFolderId}.`,
        }],
  };
}

async function deleteEmail(args, client) {
    await client.delete(`/me/messages/${args.messageId}`);

  return {
        content: [{
                type: 'text',
                text: 'Email deleted (moved to Deleted Items).',
        }],
  };
}
