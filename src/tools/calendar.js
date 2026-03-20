/**
 * Calendar Tools
 * Tools for managing Outlook calendar events via Microsoft Graph API
 */

export const calendarTools = [
  {
        name: 'list-calendars',
        description: 'List all calendars in the user account.',
        inputSchema: { type: 'object', properties: {} },
  },
  {
        name: 'create-calendar',
        description: 'Create a new calendar.',
        inputSchema: {
                type: 'object',
                properties: {
                          name: { type: 'string', description: 'Name of the new calendar.' },
                          color: { type: 'string', enum: ['auto', 'lightBlue', 'lightGreen', 'lightOrange', 'lightGray', 'lightYellow', 'lightTeal', 'lightPink', 'lightBrown', 'lightRed', 'maxColor'], description: 'Calendar color.' },
                },
                required: ['name'],
        },
  },
  {
        name: 'list-events',
        description: 'List calendar events within a date range. Defaults to the next 7 days.',
        inputSchema: {
                type: 'object',
                properties: {
                          calendarId: { type: 'string', description: 'Calendar ID. Defaults to primary calendar.' },
                          startDateTime: { type: 'string', description: 'Start of time range in ISO 8601 format (e.g., "2025-01-01T00:00:00Z").' },
                          endDateTime: { type: 'string', description: 'End of time range in ISO 8601 format.' },
                          count: { type: 'number', description: 'Max events to return (default 20).' },
                },
        },
  },
  {
        name: 'create-event',
        description: 'Create a new calendar event with attendees, location, and recurrence options.',
        inputSchema: {
                type: 'object',
                properties: {
                          subject: { type: 'string', description: 'Event title/subject.' },
                          body: { type: 'string', description: 'Event description/body (HTML or text).' },
                          startDateTime: { type: 'string', description: 'Start time in ISO 8601 (e.g., "2025-03-20T10:00:00").' },
                          endDateTime: { type: 'string', description: 'End time in ISO 8601.' },
                          timeZone: { type: 'string', description: 'Time zone (e.g., "Eastern Standard Time", "UTC"). Defaults to UTC.' },
                          location: { type: 'string', description: 'Event location.' },
                          attendees: { type: 'array', items: { type: 'string' }, description: 'Array of attendee email addresses.' },
                          isOnlineMeeting: { type: 'boolean', description: 'Create as online meeting (Teams). Defaults to false.' },
                          reminderMinutes: { type: 'number', description: 'Reminder time in minutes before event.' },
                          calendarId: { type: 'string', description: 'Calendar ID to create event in. Defaults to primary.' },
                          isAllDay: { type: 'boolean', description: 'Mark as all-day event.' },
                          recurrence: {
                                      type: 'object',
                                      description: 'Recurrence pattern.',
                                      properties: {
                                                    type: { type: 'string', enum: ['daily', 'weekly', 'absoluteMonthly', 'relativeMonthly', 'absoluteYearly', 'relativeYearly'] },
                                                    interval: { type: 'number', description: 'Interval between occurrences.' },
                                                    daysOfWeek: { type: 'array', items: { type: 'string' }, description: 'Days of week for weekly recurrence.' },
                                                    endDate: { type: 'string', description: 'Recurrence end date (YYYY-MM-DD).' },
                                                    numberOfOccurrences: { type: 'number', description: 'Number of occurrences.' },
                                      },
                          },
                },
                required: ['subject', 'startDateTime', 'endDateTime'],
        },
  },
  {
        name: 'update-event',
        description: 'Update an existing calendar event.',
        inputSchema: {
                type: 'object',
                properties: {
                          eventId: { type: 'string', description: 'Event ID to update.' },
                          subject: { type: 'string', description: 'New subject.' },
                          body: { type: 'string', description: 'New body content.' },
                          startDateTime: { type: 'string', description: 'New start time.' },
                          endDateTime: { type: 'string', description: 'New end time.' },
                          timeZone: { type: 'string', description: 'Time zone.' },
                          location: { type: 'string', description: 'New location.' },
                },
                required: ['eventId'],
        },
  },
  {
        name: 'delete-event',
        description: 'Delete a calendar event.',
        inputSchema: {
                type: 'object',
                properties: {
                          eventId: { type: 'string', description: 'Event ID to delete.' },
                },
                required: ['eventId'],
        },
  },
  {
        name: 'accept-event',
        description: 'Accept a calendar event invitation.',
        inputSchema: {
                type: 'object',
                properties: {
                          eventId: { type: 'string', description: 'Event ID to accept.' },
                          comment: { type: 'string', description: 'Optional response comment.' },
                          sendResponse: { type: 'boolean', description: 'Send response to organizer. Defaults to true.' },
                },
                required: ['eventId'],
        },
  },
  {
        name: 'decline-event',
        description: 'Decline a calendar event invitation.',
        inputSchema: {
                type: 'object',
                properties: {
                          eventId: { type: 'string', description: 'Event ID to decline.' },
                          comment: { type: 'string', description: 'Optional response comment.' },
                          sendResponse: { type: 'boolean', description: 'Send response to organizer. Defaults to true.' },
                },
                required: ['eventId'],
        },
  },
  ];

export async function handleCalendarTool(name, args, client) {
    switch (name) {
      case 'list-calendars': return listCalendars(client);
      case 'create-calendar': return createCalendar(args, client);
      case 'list-events': return listEvents(args, client);
      case 'create-event': return createEvent(args, client);
      case 'update-event': return updateEvent(args, client);
      case 'delete-event': return deleteEvent(args, client);
      case 'accept-event': return respondEvent(args, client, 'accept');
      case 'decline-event': return respondEvent(args, client, 'decline');
      default:
              return { content: [{ type: 'text', text: `Unknown calendar tool: ${name}` }], isError: true };
    }
}

async function listCalendars(client) {
    const result = await client.get('/me/calendars?$select=id,name,color,isDefaultCalendar,canEdit');
    const calendars = result.value || [];

  const formatted = calendars.map(c => [
        `  Name: ${c.name}`,
        `  ID: ${c.id}`,
        `  Color: ${c.color}`,
        `  Default: ${c.isDefaultCalendar ? 'Yes' : 'No'}`,
        `  Editable: ${c.canEdit ? 'Yes' : 'No'}`,
        '  ---',
      ].join('\n')).join('\n');

  return { content: [{ type: 'text', text: `Calendars (${calendars.length}):\n\n${formatted}` }] };
}

async function createCalendar(args, client) {
    const body = { name: args.name };
    if (args.color) body.color = args.color;

  const result = await client.post('/me/calendars', body);
    return {
          content: [{
                  type: 'text',
                  text: `Calendar "${result.name}" created successfully.\nID: ${result.id}`,
          }],
    };
}

async function listEvents(args, client) {
    const now = new Date();
    const start = args.startDateTime || now.toISOString();
    const end = args.endDateTime || new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000).toISOString();
    const count = args.count || 20;

  const base = args.calendarId ? `/me/calendars/${args.calendarId}` : '/me';
    const endpoint = `${base}/calendarView?startDateTime=${encodeURIComponent(start)}&endDateTime=${encodeURIComponent(end)}&$top=${count}&$orderby=start/dateTime&$select=id,subject,start,end,location,organizer,isOnlineMeeting,attendees,recurrence,isCancelled`;

  const result = await client.get(endpoint);
    const events = result.value || [];

  if (events.length === 0) {
        return { content: [{ type: 'text', text: 'No events found in the specified time range.' }] };
  }

  const formatted = events.map(e => {
        const attendeeList = e.attendees?.map(a =>
                `${a.emailAddress.name || a.emailAddress.address} (${a.status?.response || 'none'})`
                                                  ).join(', ') || 'None';

                                   return [
                                           `Subject: ${e.subject}`,
                                           `ID: ${e.id}`,
                                           `Start: ${e.start?.dateTime} (${e.start?.timeZone})`,
                                           `End: ${e.end?.dateTime} (${e.end?.timeZone})`,
                                           `Location: ${e.location?.displayName || 'None'}`,
                                           `Online Meeting: ${e.isOnlineMeeting ? 'Yes' : 'No'}`,
                                           `Organizer: ${e.organizer?.emailAddress?.name || e.organizer?.emailAddress?.address || 'Unknown'}`,
                                           `Attendees: ${attendeeList}`,
                                           `Cancelled: ${e.isCancelled ? 'Yes' : 'No'}`,
                                           '---',
                                         ].join('\n');
  }).join('\n');

  return { content: [{ type: 'text', text: `Events (${events.length}):\n\n${formatted}` }] };
}

async function createEvent(args, client) {
    const tz = args.timeZone || 'UTC';
    const event = {
          subject: args.subject,
          start: { dateTime: args.startDateTime, timeZone: tz },
          end: { dateTime: args.endDateTime, timeZone: tz },
    };

  if (args.body) event.body = { contentType: 'html', content: args.body };
    if (args.location) event.location = { displayName: args.location };
    if (args.isAllDay) event.isAllDay = true;
    if (args.reminderMinutes !== undefined) event.reminderMinutesBeforeStart = args.reminderMinutes;
    if (args.isOnlineMeeting) {
          event.isOnlineMeeting = true;
          event.onlineMeetingProvider = 'teamsForBusiness';
    }

  if (args.attendees) {
        event.attendees = args.attendees.map(email => ({
                emailAddress: { address: email },
                type: 'required',
        }));
  }

  if (args.recurrence) {
        const r = args.recurrence;
        event.recurrence = {
                pattern: {
                          type: r.type,
                          interval: r.interval || 1,
                          ...(r.daysOfWeek && { daysOfWeek: r.daysOfWeek }),
                },
                range: r.endDate
                  ? { type: 'endDate', endDate: r.endDate, startDate: args.startDateTime.split('T')[0] }
                          : r.numberOfOccurrences
                  ? { type: 'numbered', numberOfOccurrences: r.numberOfOccurrences, startDate: args.startDateTime.split('T')[0] }
                          : { type: 'noEnd', startDate: args.startDateTime.split('T')[0] },
        };
  }

  const base = args.calendarId ? `/me/calendars/${args.calendarId}/events` : '/me/events';
    const result = await client.post(base, event);

  let text = `Event "${result.subject}" created successfully.\nID: ${result.id}`;
    if (result.onlineMeeting?.joinUrl) {
          text += `\nTeams Meeting URL: ${result.onlineMeeting.joinUrl}`;
    }

  return { content: [{ type: 'text', text }] };
}

async function updateEvent(args, client) {
    const update = {};
    if (args.subject) update.subject = args.subject;
    if (args.body) update.body = { contentType: 'html', content: args.body };
    if (args.location) update.location = { displayName: args.location };
    const tz = args.timeZone || 'UTC';
    if (args.startDateTime) update.start = { dateTime: args.startDateTime, timeZone: tz };
    if (args.endDateTime) update.end = { dateTime: args.endDateTime, timeZone: tz };

  await client.patch(`/me/events/${args.eventId}`, update);
    return { content: [{ type: 'text', text: 'Event updated successfully.' }] };
}

async function deleteEvent(args, client) {
    await client.delete(`/me/events/${args.eventId}`);
    return { content: [{ type: 'text', text: 'Event deleted successfully.' }] };
}

async function respondEvent(args, client, action) {
    await client.post(`/me/events/${args.eventId}/${action}`, {
          comment: args.comment || '',
          sendResponse: args.sendResponse !== false,
    });
    return { content: [{ type: 'text', text: `Event ${action}ed successfully.` }] };
}
