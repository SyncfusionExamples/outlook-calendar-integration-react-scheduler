// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <GetUserSnippet>
import { Client, GraphRequestOptions, PageCollection, PageIterator } from '@microsoft/microsoft-graph-client';
import { AuthCodeMSALBrowserAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';
import { endOfWeek, startOfWeek } from 'date-fns';
import { zonedTimeToUtc } from 'date-fns-tz';
import { User, Event } from '@microsoft/microsoft-graph-types';

let graphClient: Client | undefined = undefined;

function ensureClient(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
  if (!graphClient) {
    graphClient = Client.initWithMiddleware({

      authProvider: authProvider
    });
  }

  return graphClient;
}

export async function getUser(authProvider: AuthCodeMSALBrowserAuthenticationProvider): Promise<User> {
  ensureClient(authProvider);

  // Return the /me API endpoint result as a User object
  const user: User = await graphClient!.api('/me')
    // Only retrieve the specific fields needed
    .select('displayName,mail,mailboxSettings,userPrincipalName')
    .get();

  return user;
}
// </GetUserSnippet>

// <GetUserWeekCalendarSnippet>
export async function getUserWeekCalendar(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  timeZone: string, startDate: Date, endDate:Date): Promise<Event[]> {
  ensureClient(authProvider);

  // Generate startDateTime and endDateTime query params
  // to display a 7-day window
  const startDateTime = zonedTimeToUtc(startDate, timeZone).toISOString();
  const endDateTime = zonedTimeToUtc(endDate, timeZone).toISOString();


  // GET /me/calendarview?startDateTime=''&endDateTime=''
  // &$select=subject,organizer,start,end
  // &$orderby=start/dateTime
  // &$top=50
  var response: PageCollection = await graphClient!
    .api('/me/calendarview')
    .header('Prefer', `outlook.timezone="${timeZone}"`)
    .query({ startDateTime: startDateTime, endDateTime: endDateTime })
    .select('subject,organizer,start,end,recurrence')
    .orderby('start/dateTime')
    .top(25)
    .get();

  if (response["@odata.nextLink"]) {
    // Presence of the nextLink property indicates more results are available
    // Use a page iterator to get all results
    var events: Event[] = [];

    // Must include the time zone header in page
    // requests too
    var options: GraphRequestOptions = {
      headers: { 'Prefer': `outlook.timezone="${timeZone}"` }
    };

    var pageIterator = new PageIterator(graphClient!, response, (event) => {
      events.push(event);
      return true;
    }, options);

    await pageIterator.iterate();

    const schedulerEvents: any = events.map((event) => ({
      ...event,
      subject: event.subject,
      start: event.start?.dateTime,
      end: event.end?.dateTime,
      startTimezone: event.start?.timeZone,
      endTimezone: event.end?.timeZone,
      recurrence: event.recurrence
    }));
    return schedulerEvents;
  } else {
    const schedulerEvents = response.value.map((event) => ({
      ...event,
      subject: event.subject,
      start: event.start.dateTime,
      end: event.end.dateTime,
      startTimezone: event.start.timeZone,
      endTimezone: event.end.timeZone,
      recurrence: event.recurrence
    }));
    return schedulerEvents;
  }
}
// </GetUserWeekCalendarSnippet>



const buildRecurrence = (rule: string, event: any) => {
  const [freqPart, intervalPart, countPart] = rule.split(';');
  const frequency = freqPart.split('=')[1].toLowerCase();
  const interval = parseInt(intervalPart.split('=')[1]);
  const count = parseInt(countPart.split('=')[1]);

  return {
    pattern: {
      type: frequency,
      interval: interval,
    },
    range: {
      type: "numbered",
      startDate: event.start.toISOString().split('T')[0],
      numberOfOccurrences: count,
    },
  };
};

const timeZone = (tz: string | undefined) => (tz === null || tz === undefined) ? 'UTC' : tz;

export async function CreateEvent(event: any) {
  const event1 = {
    subject: `${event.subject}`,
    start: {
      dateTime: event.start.toISOString(),
      timeZone: timeZone(event.StartTimezone),
    },
    end: {
      dateTime: event.end.toISOString(),
      timeZone: timeZone(event.EndTimezone),
    },
    ...(event.RecurrenceRule && { recurrence: buildRecurrence(event.RecurrenceRule, event) }),
  };
  return await graphClient!.api("/me/events").post(event1);
}


export async function updateEvent(event: any) {

  if (!event || !event.id) {
    throw new Error('Event ID is required and must be valid.');
  }

  const event1 = {
    subject: `${event.subject}`,
    start: {
      dateTime: event.start.toISOString(),
      timeZone: timeZone(event.StartTimezone),
    },
    end: {
      dateTime: event.end.toISOString(),
      timeZone: timeZone(event.EndTimezone),
    },
    ...(event.RecurrenceRule && { recurrence: buildRecurrence(event.RecurrenceRule, event) }),
  };
  return await graphClient!.api(`/me/events/${event.id}`).patch(event1);
}

export async function deleteEvent(id: any) {
  return await graphClient!.api(`/me/events/${id}`).delete();
}


