/**
 * Contact Tools
 * Tools for managing Outlook contacts via Microsoft Graph API
 */

export const contactTools = [
  {
        name: 'list-contacts',
        description: 'List contacts from the default contacts folder or a specific folder.',
        inputSchema: {
                type: 'object',
                properties: {
                          folderId: { type: 'string', description: 'Contact folder ID. Defaults to main contacts folder.' },
                          count: { type: 'number', description: 'Number of contacts to return (max 100, default 25).' },
                          search: { type: 'string', description: 'Search query to filter contacts by name or email.' },
                },
        },
  },
  {
        name: 'get-contact',
        description: 'Get detailed information about a specific contact.',
        inputSchema: {
                type: 'object',
                properties: {
                          contactId: { type: 'string', description: 'The contact ID.' },
                },
                required: ['contactId'],
        },
  },
  {
        name: 'create-contact',
        description: 'Create a new contact with name, email, phone, company, and address details.',
        inputSchema: {
                type: 'object',
                properties: {
                          givenName: { type: 'string', description: 'First name.' },
                          surname: { type: 'string', description: 'Last name.' },
                          emailAddresses: { type: 'array', items: { type: 'string' }, description: 'Email addresses.' },
                          businessPhones: { type: 'array', items: { type: 'string' }, description: 'Business phone numbers.' },
                          mobilePhone: { type: 'string', description: 'Mobile phone number.' },
                          companyName: { type: 'string', description: 'Company name.' },
                          jobTitle: { type: 'string', description: 'Job title.' },
                          department: { type: 'string', description: 'Department.' },
                          officeLocation: { type: 'string', description: 'Office location.' },
                          personalNotes: { type: 'string', description: 'Personal notes about the contact.' },
                          folderId: { type: 'string', description: 'Contact folder ID. Defaults to main folder.' },
                },
                required: ['givenName'],
        },
  },
  {
        name: 'update-contact',
        description: 'Update an existing contact.',
        inputSchema: {
                type: 'object',
                properties: {
                          contactId: { type: 'string', description: 'The contact ID to update.' },
                          givenName: { type: 'string' },
                          surname: { type: 'string' },
                          emailAddresses: { type: 'array', items: { type: 'string' } },
                          businessPhones: { type: 'array', items: { type: 'string' } },
                          mobilePhone: { type: 'string' },
                          companyName: { type: 'string' },
                          jobTitle: { type: 'string' },
                          personalNotes: { type: 'string' },
                },
                required: ['contactId'],
        },
  },
  {
        name: 'delete-contact',
        description: 'Delete a contact.',
        inputSchema: {
                type: 'object',
                properties: {
                          contactId: { type: 'string', description: 'The contact ID to delete.' },
                },
                required: ['contactId'],
        },
  },
  {
        name: 'list-contact-folders',
        description: 'List all contact folders.',
        inputSchema: { type: 'object', properties: {} },
  },
  {
        name: 'create-contact-folder',
        description: 'Create a new contact folder to organize contacts.',
        inputSchema: {
                type: 'object',
                properties: {
                          displayName: { type: 'string', description: 'Name for the new contact folder.' },
                          parentFolderId: { type: 'string', description: 'Parent folder ID for nested folders.' },
                },
                required: ['displayName'],
        },
  },
  ];

export async function handleContactTool(name, args, client) {
    switch (name) {
      case 'list-contacts': return listContacts(args, client);
      case 'get-contact': return getContact(args, client);
      case 'create-contact': return createContact(args, client);
      case 'update-contact': return updateContact(args, client);
      case 'delete-contact': return deleteContact(args, client);
      case 'list-contact-folders': return listContactFolders(client);
      case 'create-contact-folder': return createContactFolder(args, client);
      default:
              return { content: [{ type: 'text', text: `Unknown contact tool: ${name}` }], isError: true };
    }
}

function formatContact(c) {
    const emails = c.emailAddresses?.map(e => e.address).join(', ') || 'None';
    const phones = [...(c.businessPhones || []), c.mobilePhone].filter(Boolean).join(', ') || 'None';

  return [
        `  Name: ${c.displayName || `${c.givenName || ''} ${c.surname || ''}`.trim()}`,
        `  ID: ${c.id}`,
        `  Email: ${emails}`,
        `  Phone: ${phones}`,
        `  Company: ${c.companyName || 'N/A'}`,
        `  Title: ${c.jobTitle || 'N/A'}`,
        '  ---',
      ].join('\n');
}

async function listContacts(args, client) {
    const count = Math.min(args.count || 25, 100);
    let endpoint;

  if (args.search) {
        endpoint = args.folderId
          ? `/me/contactFolders/${args.folderId}/contacts?$search="${encodeURIComponent(args.search)}"&$top=${count}`
                : `/me/contacts?$search="${encodeURIComponent(args.search)}"&$top=${count}`;
  } else {
        endpoint = args.folderId
          ? `/me/contactFolders/${args.folderId}/contacts?$top=${count}&$orderby=displayName`
                : `/me/contacts?$top=${count}&$orderby=displayName`;
  }

  endpoint += '&$select=id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,companyName,jobTitle';

  const result = await client.get(endpoint);
    const contacts = result.value || [];

  if (contacts.length === 0) {
        return { content: [{ type: 'text', text: 'No contacts found.' }] };
  }

  const formatted = contacts.map(formatContact).join('\n');
    return { content: [{ type: 'text', text: `Contacts (${contacts.length}):\n\n${formatted}` }] };
}

async function getContact(args, client) {
    const c = await client.get(`/me/contacts/${args.contactId}`);

  const emails = c.emailAddresses?.map(e => `${e.name || ''} <${e.address}>`).join('\n    ') || 'None';
    const text = [
          `Name: ${c.displayName}`,
          `First: ${c.givenName || 'N/A'}`,
          `Last: ${c.surname || 'N/A'}`,
          `Emails:\n    ${emails}`,
          `Business Phones: ${c.businessPhones?.join(', ') || 'None'}`,
          `Mobile: ${c.mobilePhone || 'N/A'}`,
          `Company: ${c.companyName || 'N/A'}`,
          `Job Title: ${c.jobTitle || 'N/A'}`,
          `Department: ${c.department || 'N/A'}`,
          `Office: ${c.officeLocation || 'N/A'}`,
          `Notes: ${c.personalNotes || 'None'}`,
          `Created: ${c.createdDateTime}`,
          `Modified: ${c.lastModifiedDateTime}`,
        ].join('\n');

  return { content: [{ type: 'text', text }] };
}

async function createContact(args, client) {
    const contact = {};
    if (args.givenName) contact.givenName = args.givenName;
    if (args.surname) contact.surname = args.surname;
    if (args.companyName) contact.companyName = args.companyName;
    if (args.jobTitle) contact.jobTitle = args.jobTitle;
    if (args.department) contact.department = args.department;
    if (args.officeLocation) contact.officeLocation = args.officeLocation;
    if (args.mobilePhone) contact.mobilePhone = args.mobilePhone;
    if (args.personalNotes) contact.personalNotes = args.personalNotes;
    if (args.businessPhones) contact.businessPhones = args.businessPhones;
    if (args.emailAddresses) {
          contact.emailAddresses = args.emailAddresses.map(addr => ({
                  address: addr,
                  name: addr,
          }));
    }

  const endpoint = args.folderId
      ? `/me/contactFolders/${args.folderId}/contacts`
        : '/me/contacts';

  const result = await client.post(endpoint, contact);
    return {
          content: [{
                  type: 'text',
                  text: `Contact "${result.displayName}" created.\nID: ${result.id}`,
          }],
    };
}

async function updateContact(args, client) {
    const update = {};
    if (args.givenName) update.givenName = args.givenName;
    if (args.surname) update.surname = args.surname;
    if (args.companyName) update.companyName = args.companyName;
    if (args.jobTitle) update.jobTitle = args.jobTitle;
    if (args.mobilePhone) update.mobilePhone = args.mobilePhone;
    if (args.personalNotes) update.personalNotes = args.personalNotes;
    if (args.businessPhones) update.businessPhones = args.businessPhones;
    if (args.emailAddresses) {
          update.emailAddresses = args.emailAddresses.map(addr => ({
                  address: addr,
                  name: addr,
          }));
    }

  await client.patch(`/me/contacts/${args.contactId}`, update);
    return { content: [{ type: 'text', text: 'Contact updated successfully.' }] };
}

async function deleteContact(args, client) {
    await client.delete(`/me/contacts/${args.contactId}`);
    return { content: [{ type: 'text', text: 'Contact deleted.' }] };
}

async function listContactFolders(client) {
    const result = await client.get('/me/contactFolders?$select=id,displayName,parentFolderId');
    const folders = result.value || [];

  const formatted = folders.map(f => [
        `  Name: ${f.displayName}`,
        `  ID: ${f.id}`,
        '  ---',
      ].join('\n')).join('\n');

  return { content: [{ type: 'text', text: `Contact Folders (${folders.length}):\n\n${formatted}` }] };
}

async function createContactFolder(args, client) {
    const endpoint = args.parentFolderId
      ? `/me/contactFolders/${args.parentFolderId}/childFolders`
          : '/me/contactFolders';

  const result = await client.post(endpoint, { displayName: args.displayName });
    return {
          content: [{
                  type: 'text',
                  text: `Contact folder "${result.displayName}" created.\nID: ${result.id}`,
          }],
    };
}
