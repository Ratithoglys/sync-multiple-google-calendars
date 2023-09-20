// Source : https://github.com/karbassi/sync-multiple-google-calendars
// Alter : Len

// ----------------------------------------------------------------------------
// CONFIGURE FROM HERE ON
// ----------------------------------------------------------------------------

// Calendars to merge from.
// Name       is just for show
// id         is the calendar id found in calendar settings
// prefix     is a string to prepend for all events of this calendar
// color_id   is the default color for each events of this calendar, don't specify a color to use the default color
//    0=default, 1=blue, 2=green, 3=purple, 4=red, 5=yellow, 6=orange, 
//    7=turquoise, 8=gray, 9=bold blue, 10=bold green, 11=bold red
const CALENDARS_TO_MERGE = {
  "General": {'id': "calendar-id@gmail.com", 'prefix': "", 'color_id': '2'},
  "Work": {'id': "calendar-id@gmail.com", 'prefix': "[WK]", 'color_id': '5'},
}

// The ID of the merged calendar
const CALENDAR_TO_MERGE_INTO = "shared-calendar-id@gmail.com";

// ----------------------------------------------------------------------------
// ONLY NECESSARY ON FIRST EXEC
// ----------------------------------------------------------------------------

// Number of days in the past and future to sync.
var SYNC_DAYS_IN_PAST = 10;
var SYNC_DAYS_IN_FUTURE = 15;

// Default title for events that don't have a title.
var DEFAULT_EVENT_TITLE = "That's a mystery ! 🤷🏻‍♂️";

// Sync only events marked as busy (or all). Default: false (all events)
var SKIP_NONBUSY_EVENTS = false;

// Do not sync declined events. Default: false (do sync)
var DO_NOT_SYNC_DECLINED = false

// ----------------------------------------------------------------------------
// MAIN LOGIC
// ----------------------------------------------------------------------------

// Base endpoint for the calendar API
const ENDPOINT_BASE = "https://www.googleapis.com/calendar/v3/calendars";

// Unique character to use in the title of the event to identify it as a clone.
// This is used to delete the old events.
// https://unicode-table.com/en/200B/
const SEARCH_CHARACTER = "\u200B";

function SyncCalendarsIntoOne() {
  // Execution semaphore - https://stackoverflow.com/questions/67066779/how-to-prevent-google-apps-script-trigger-if-a-function-is-already-running
  var isItRunning;
  isItRunning = CacheService.getScriptCache().put("itzRunning", "true", 90); //Keep this value in Cache for up to X minutes
  if (isItRunning) {
    // If this is true then another instance of this function is running which means that you dont want this instance of this function to run - so quit
    return;
  }

  checkProperties();

  // Start time is today at midnight - SYNC_DAYS_IN_PAST
  const startTime = new Date();
  startTime.setHours(0, 0, 0, 0);
  startTime.setDate(startTime.getDate() - SYNC_DAYS_IN_PAST);

  // End time is today at midnight + SYNC_DAYS_IN_FUTURE
  const endTime = new Date();
  endTime.setHours(0, 0, 0, 0);
  endTime.setDate(endTime.getDate() + SYNC_DAYS_IN_FUTURE + 1);

  // Delete any old events that have been already cloned over.
  //const deleteStartTime = new Date();
  //deleteStartTime.setFullYear(2000, 01, 01);
  //deleteStartTime.setHours(0, 0, 0, 0);

  deleteEvents(startTime, endTime);
  createEvents(startTime, endTime);

  // Remove execution semaphore
  CacheService.getScriptCache().remove("itzRunning");
}

// Delete any old events that have been already cloned over.
// This is basically a sync w/o finding and updating. Just deleted and recreate.
function deleteEvents(startTime, endTime) {
  const sharedCalendar = CalendarApp.getCalendarById(CALENDAR_TO_MERGE_INTO);

  // Find events with the search character in the title.
  // The `.filter` method is used since the getEvents method seems to return all events at the moment. It's a safety check.
  const events = sharedCalendar
    .getEvents(startTime, endTime, { search: SEARCH_CHARACTER })
    .filter((event) => event.getTitle().includes(SEARCH_CHARACTER));

  const requestBody = events.map((e, i) => ({
    method: "DELETE",
    endpoint: `${ENDPOINT_BASE}/${CALENDAR_TO_MERGE_INTO}/events/${e
      .getId()
      .replace("@google.com", "")}`,
  }));

  if (requestBody && requestBody.length) {
    const result = new BatchRequest({
      useFetchAll: true,
      batchPath: "batch/calendar/v3",
      requests: requestBody,
    });

    if (result.length !== requestBody.length) {
      console.log(result);
    }

    console.log(`${result.length} deleted events between ${startTime} and ${endTime}.`);
  } else {
    console.log("No events to delete.");
  }
}

function createEvents(startTime, endTime) {
  let requestBody = [];

  for (const [calendarName, calendarOption] of Object.entries(CALENDARS_TO_MERGE)) {
    const calendarId = calendarOption.id;
    const calendarPrefix = calendarOption.prefix;
    const calendarColorId = calendarOption.color_id || '0';
    const calendarToCopy = CalendarApp.getCalendarById(calendarId);

    if (!calendarToCopy) {
      console.log("Calendar not found: '%s'.", calendarId);
      continue;
    }

    // Find events
    const events = Calendar.Events.list(calendarId, {
      timeMin: startTime.toISOString(),
      timeMax: endTime.toISOString(),
      singleEvents: true,
      orderBy: "startTime",
    });

    // If nothing find, move to next calendar
    if (!(events.items && events.items.length > 0)) {
      continue;
    }

    events.items.forEach((event) => {
      // Don't copy "free" events.
      if (event.transparency && event.transparency === "transparent" && SKIP_NONBUSY_EVENTS) {
        return;
      }

      // Don't copy declined events.
      if (DO_NOT_SYNC_DECLINED && event.attendees && event.attendees.find((at) => at.self) && event.attendees.find((at) => at.self)['responseStatus'] == 'declined') {
        return
      }

      // If event.summary is undefined, empty, or null, set it to default title
      if (!event.summary || event.summary === "") {
        event.summary = DEFAULT_EVENT_TITLE;
      }

      requestBody.push({
        method: "POST",
        endpoint: `${ENDPOINT_BASE}/${CALENDAR_TO_MERGE_INTO}/events?conferenceDataVersion=1`,
        requestBody: {
          summary: `${SEARCH_CHARACTER}${calendarPrefix} ${event.summary}`,
          location: event.location,
          description: event.description,
          start: event.start,
          end: event.end,
          colorId: calendarColorId,
          conferenceData: event.conferenceData,
        },
      });
    });
  }

  if (requestBody && requestBody.length) {
    const result = new BatchRequest({
      batchPath: "batch/calendar/v3",
      requests: requestBody,
    });

    if (result.length !== requestBody.length) {
      console.log(result);
    }

    console.log(`${result.length} events created between ${startTime} and ${endTime}.`);
  } else {
    console.log("No events to create.");
  }
}

function checkProperties() {
  SYNC_DAYS_IN_PAST     = checkProperty('SYNC_DAYS_IN_PAST', SYNC_DAYS_IN_PAST);
  SYNC_DAYS_IN_FUTURE   = checkProperty('SYNC_DAYS_IN_FUTURE', SYNC_DAYS_IN_FUTURE);
  DEFAULT_EVENT_TITLE   = checkProperty('DEFAULT_EVENT_TITLE', DEFAULT_EVENT_TITLE);
  SKIP_NONBUSY_EVENTS   = (checkProperty('SKIP_NONBUSY_EVENTS', SKIP_NONBUSY_EVENTS) === 'true');
  DO_NOT_SYNC_DECLINED  = (checkProperty('DO_NOT_SYNC_DECLINED', DO_NOT_SYNC_DECLINED) === 'true');
}

function checkProperty(prop, val) {
  props = PropertiesService.getScriptProperties();

  if (props.getProperty(prop) == null) {
    Logger.log(`${prop} unavailable, setting to ${val}`)
    props.setProperty(prop, val)
  }

  return props.getProperty(prop);
}
