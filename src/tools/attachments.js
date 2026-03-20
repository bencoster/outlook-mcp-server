/**
 * Attachment Tools
 * Tools for managing email attachments via Microsoft Graph API
 */

export const attachmentTools = [
  {
        name: 'list-attachments',
        description: 'List all attachments on a specific email message.',
        inputSchema: {
                type: 'object',
                properties: {
                          messageId: { type: 'string', description: 'The email message ID.' },
                },
                required: ['messageId'],
        },
  },
  {
        name: 'get-attachment',
        description: 'Get the content of a specific attachment. Returns base64-encoded content for file attachments.',
        inputSchema: {
                type: 'object',
                properties: {
                          messageId: { type: 'string', description: 'The email message ID.' },
                          attachmentId: { type: 'string', description: 'The attachment ID.' },
                },
                required: ['messageId', 'attachmentId'],
        },
  },
  {
        name: 'add-attachment',
        description: 'Add a file attachment to a draft email message. Content must be base64-encoded.',
        inputSchema: {
                type: 'object',
                properties: {
                          messageId: { type: 'string', description: 'The draft message ID to attach the file to.' },
                          name: { type: 'string', description: 'File name (e.g., "report.pdf").' },
                          contentType: { type: 'string', description: 'MIME type (e.g., "application/pdf").' },
                          contentBytes: { type: 'string', description: 'Base64-encoded file content.' },
                },
                required: ['messageId', 'name', 'contentBytes'],
        },
  },
  {
        name: 'delete-attachment',
        description: 'Remove an attachment from a draft email message.',
        inputSchema: {
                type: 'object',
                properties: {
                          messageId: { type: 'string', description: 'The message ID.' },
                          attachmentId: { type: 'string', description: 'The attachment ID to remove.' },
                },
                required: ['messageId', 'attachmentId'],
        },
  },
  {
        name: 'create-draft-with-attachment',
        description: 'Create a new draft email with an attachment in one step.',
        inputSchema: {
                type: 'object',
                properties: {
                          to: { type: 'array', items: { type: 'string' }, description: 'Recipient email addresses.' },
                          subject: { type: 'string', description: 'Email subject.' },
                          body: { type: 'string', description: 'Email body.' },
                          bodyType: { type: 'string', enum: ['text', 'html'], description: 'Body type (default: html).' },
                          attachmentName: { type: 'string', description: 'Attachment file name.' },
                          attachmentContentType: { type: 'string', description: 'Attachment MIME type.' },
                          attachmentContent: { type: 'string', description: 'Base64-encoded attachment content.' },
                },
                required: ['to', 'subject', 'body', 'attachmentName', 'attachmentContent'],
        },
  },
  ];

export async function handleAttachmentTool(name, args, client) {
    switch (name) {
      case 'list-attachments': return listAttachments(args, client);
      case 'get-attachment': return getAttachment(args, client);
      case 'add-attachment': return addAttachment(args, client);
      case 'delete-attachment': return deleteAttachment(args, client);
      case 'create-draft-with-attachment': return createDraftWithAttachment(args, client);
      default:
              return { content: [{ type: 'text', text: `Unknown attachment tool: ${name}` }], isError: true };
    }
}

async function listAttachments(args, client) {
    const result = await client.get(
          `/me/messages/${args.messageId}/attachments?$select=id,name,contentType,size,isInline`
        );
    const attachments = result.value || [];

  if (attachments.length === 0) {
        return { content: [{ type: 'text', text: 'No attachments found on this message.' }] };
  }

  const formatted = attachments.map(a => [
        `  Name: ${a.name}`,
        `  ID: ${a.id}`,
        `  Type: ${a.contentType}`,
        `  Size: ${(a.size / 1024).toFixed(1)} KB`,
        `  Inline: ${a.isInline ? 'Yes' : 'No'}`,
        '  ---',
      ].join('\n')).join('\n');

  return { content: [{ type: 'text', text: `Attachments (${attachments.length}):\n\n${formatted}` }] };
}

async function getAttachment(args, client) {
    const attachment = await client.get(
          `/me/messages/${args.messageId}/attachments/${args.attachmentId}`
        );

  const text = [
        `Name: ${attachment.name}`,
        `Type: ${attachment.contentType}`,
        `Size: ${(attachment.size / 1024).toFixed(1)} KB`,
        '',
        attachment.contentBytes
          ? `Content (base64): ${attachment.contentBytes.substring(0, 200)}${attachment.contentBytes.length > 200 ? '...' : ''}`
          : 'No content bytes available (may be a reference attachment).',
      ].join('\n');

  return { content: [{ type: 'text', text }] };
}

async function addAttachment(args, client) {
    const attachment = {
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: args.name,
          contentType: args.contentType || 'application/octet-stream',
          contentBytes: args.contentBytes,
    };

  await client.post(`/me/messages/${args.messageId}/attachments`, attachment);
    return {
          content: [{
                  type: 'text',
                  text: `Attachment "${args.name}" added to message.`,
          }],
    };
}

async function deleteAttachment(args, client) {
    await client.delete(`/me/messages/${args.messageId}/attachments/${args.attachmentId}`);
    return { content: [{ type: 'text', text: 'Attachment removed from message.' }] };
}

async function createDraftWithAttachment(args, client) {
    // First create the draft
  const message = {
        subject: args.subject,
        body: {
                contentType: args.bodyType === 'text' ? 'text' : 'html',
                content: args.body,
        },
        toRecipients: args.to.map(email => ({
                emailAddress: { address: email },
        })),
  };

  const draft = await client.post('/me/messages', message);

  // Then add the attachment
  const attachment = {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: args.attachmentName,
        contentType: args.attachmentContentType || 'application/octet-stream',
        contentBytes: args.attachmentContent,
  };

  await client.post(`/me/messages/${draft.id}/attachments`, attachment);

  return {
        content: [{
                type: 'text',
                text: [
                          `Draft email created with attachment "${args.attachmentName}".`,
                          `Draft ID: ${draft.id}`,
                          '',
                          'You can send this draft using the send-email tool or add more attachments.',
                        ].join('\n'),
        }],
  };
}
