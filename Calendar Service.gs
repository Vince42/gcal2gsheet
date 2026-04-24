function resolveCalendars_() {
  return CONFIG.calendarNames.map((calendarName) => {
    const calendars = CalendarApp.getCalendarsByName(calendarName);
    if (!calendars || calendars.length === 0) {
      throw new Error(`Calendar not found by name: ${calendarName}`);
    }

    return {
      name: calendarName,
      id: calendars[0].getId(),
    };
  });
}

function fetchFullSnapshot_(ss, calendars, timeZone, scope) {
  const currentByKey = new Map();
  const nextSyncTokens = {};

  calendars.forEach((calendarInfo, index) => {
    setProgress_(
      ss,
      `Full import ${index + 1}/${calendars.length}: ${calendarInfo.name}...`
    );

    const response = fetchCalendarFull_(ss, calendarInfo, timeZone, scope);

    response.items.forEach((item) => {
      const converted = convertApiEvent_(calendarInfo, item, timeZone);
      if (converted && isManagedEventInScope_(converted, scope)) {
        currentByKey.set(converted.eventKey, converted);
      }
    });

    nextSyncTokens[calendarInfo.id] = response.nextSyncToken || '';
  });

  return {
    currentByKey,
    nextSyncTokens,
  };
}

function fetchIncrementalChanges_(ss, calendars, timeZone) {
  const deltaByKey = new Map();
  const nextSyncTokens = {};
  const syncTokens = loadSyncTokens_(calendars);

  calendars.forEach((calendarInfo, index) => {
    setProgress_(
      ss,
      `Incremental sync ${index + 1}/${calendars.length}: ${calendarInfo.name}...`
    );

    const tokenEntry = syncTokens.find((item) => item.calendarId === calendarInfo.id);
    const syncToken = tokenEntry ? tokenEntry.syncToken : '';
    const response = fetchCalendarIncremental_(ss, calendarInfo, timeZone, syncToken);

    response.items.forEach((item) => {
      const eventKey = buildEventKey_(calendarInfo.id, item.id);
      const converted = convertApiEvent_(calendarInfo, item, timeZone);

      if (converted) {
        deltaByKey.set(eventKey, converted);
      } else {
        deltaByKey.set(eventKey, null);
      }
    });

    nextSyncTokens[calendarInfo.id] = response.nextSyncToken || '';
  });

  return {
    deltaByKey,
    nextSyncTokens,
  };
}

function fetchCalendarFull_(ss, calendarInfo, timeZone, scope) {
  const items = [];
  let pageToken = null;
  let nextSyncToken = '';
  let page = 0;

  do {
    page += 1;
    const response = Calendar.Events.list(calendarInfo.id, {
      singleEvents: true,
      showDeleted: false,
      orderBy: 'startTime',
      timeZone,
      timeMin: scope.importStart.toISOString(),
      timeMax: scope.now.toISOString(),
      maxResults: 2500,
      pageToken,
    });

    const pageItems = response.items || [];
    pageItems.forEach((item) => items.push(item));
    pageToken = response.nextPageToken || null;
    nextSyncToken = response.nextSyncToken || nextSyncToken;

    setProgress_(
      ss,
      `Full import ${calendarInfo.name}: page ${page}, ${items.length} item(s)...`
    );
  } while (pageToken);

  return { items, nextSyncToken };
}

function fetchCalendarIncremental_(ss, calendarInfo, timeZone, syncToken) {
  const items = [];
  let pageToken = null;
  let nextSyncToken = '';
  let page = 0;

  do {
    page += 1;
    let response;

    try {
      response = Calendar.Events.list(calendarInfo.id, {
        singleEvents: true,
        showDeleted: true,
        timeZone,
        maxResults: 2500,
        pageToken,
        syncToken,
      });
    } catch (error) {
      const message = error && error.message ? String(error.message) : String(error);
      if (message.includes('410') || message.toLowerCase().includes('sync token')) {
        throw new Error(`Invalid sync token for ${calendarInfo.name}: ${message}`);
      }
      throw error;
    }

    const pageItems = response.items || [];
    pageItems.forEach((item) => items.push(item));
    pageToken = response.nextPageToken || null;
    nextSyncToken = response.nextSyncToken || nextSyncToken;

    setProgress_(
      ss,
      `Incremental sync ${calendarInfo.name}: page ${page}, ${items.length} change(s)...`
    );
  } while (pageToken);

  return { items, nextSyncToken };
}

function convertApiEvent_(calendarInfo, item, timeZone) {
  if (!item || !item.id) {
    return null;
  }

  if (item.status === 'cancelled') {
    return null;
  }

  if (!item.start || !item.end) {
    return null;
  }

  if (item.start.date || item.end.date) {
    return null;
  }

  if (!item.start.dateTime || !item.end.dateTime) {
    return null;
  }

  const start = new Date(item.start.dateTime);
  const end = new Date(item.end.dateTime);

  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
    return null;
  }

  if (end.getTime() < start.getTime()) {
    return null;
  }

  const dateOnly = toSheetDateOnly_(start, timeZone);
  const duration = (end.getTime() - start.getTime()) / 86400000;
  const eventKey = buildEventKey_(calendarInfo.id, item.id);

  const signature = buildImportedSignature_(
    {
      calendar: calendarInfo.name,
      title: item.summary || '',
      date: dateOnly,
      start,
      end,
      duration,
    },
    timeZone
  );

  return {
    eventKey,
    calendarId: calendarInfo.id,
    calendarName: calendarInfo.name,
    title: item.summary || '',
    date: dateOnly,
    start,
    end,
    duration,
    signature,
    values: [
      calendarInfo.name,
      item.summary || '',
      dateOnly,
      start,
      end,
      duration,
      '',
      '',
      '',
      '',
    ],
  };
}

function buildEventKey_(calendarId, eventId) {
  return `${calendarId}::${eventId}`;
}
