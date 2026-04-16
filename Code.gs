const CONFIG = Object.freeze({
  sheetName: 'Calendar',
  stateSheetName: '_calendar_state',
  tableName: 'Calendar',
  statusCell: 'L1',

  // Lower bound for managed imports: yyyy-mm-dd
  importStartDate: '2024-01-01',

  calendarNames: ['Event', 'dedc', 'EEC', 'CTG'],
  defaultCalendarName: 'Event',

  header: [
    'Calendar',
    'Event',
    'Date',
    'Start',
    'End',
    'Duration',
    'Customer',
    'Project',
    'InvoiceNumber',
    'InvoiceDate',
  ],

  stateHeader: ['EventKey', 'RowKind'],

  rowKind: {
    normal: 'NORMAL',
    changedCopy: 'CHANGED_COPY',
    unmanaged: 'UNMANAGED',
  },

  propertyPrefix: 'CALSYNC_TOKEN_',

  colors: {
    normal: '#000000',
    invoiced: '#7A1F1F',
    changed: '#1B5E20',
  },

  menu: {
    title: 'Calendar Import',
    item: 'Update calendar sheet',
  },

  toastTitle: 'Calendar import',
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CONFIG.menu.title)
    .addItem(CONFIG.menu.item, 'updateCalendarSheets')
    .addToUi();
}

function updateCalendarSheets() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  const timeZone = ss.getSpreadsheetTimeZone();
  const uiState = captureUiState_(ss);
  const scope = buildScope_(timeZone);

  try {
    setProgress_(ss, 'Resolving calendars...');
    const calendars = resolveCalendars_();

    setProgress_(ss, 'Preparing sheets...');
    const sheet = ensureCalendarSheet_(ss);
    const stateSheet = ensureStateSheet_(ss);
    ensureHeader_(sheet);
    ensureStateHeader_(stateSheet);
    ensureSheetFormatting_(sheet);
    ensureTable_(spreadsheetId, sheet);

    setProgress_(ss, 'Reading existing sheet state...');
    const existingState = readExistingState_(sheet, stateSheet, timeZone, scope);
    const syncTokens = loadSyncTokens_(calendars);

    let useIncremental =
      existingState.hasManagedRows &&
      syncTokens.length === calendars.length &&
      syncTokens.every((item) => !!item.syncToken);

    let fetchResult;
    if (useIncremental) {
      try {
        setProgress_(ss, 'Fetching incremental changes...');
        fetchResult = fetchIncrementalChanges_(ss, calendars, timeZone);
      } catch (error) {
        clearSyncTokens_(calendars);
        useIncremental = false;
        setProgress_(ss, 'Sync token invalid. Falling back to full import...');
        fetchResult = fetchFullSnapshot_(ss, calendars, timeZone, scope);
      }
    } else {
      setProgress_(ss, 'Performing initial full import...');
      fetchResult = fetchFullSnapshot_(ss, calendars, timeZone, scope);
    }

    const changedNotifications = [];
    let finalRows;

    setProgress_(ss, 'Rebuilding worksheet data...');
    if (useIncremental) {
      finalRows = rebuildFromIncremental_(
        existingState,
        fetchResult.deltaByKey,
        changedNotifications,
        timeZone,
        scope
      );
    } else {
      finalRows = rebuildFromFullSnapshot_(
        existingState,
        fetchResult.currentByKey,
        changedNotifications,
        timeZone,
        scope
      );
    }

    setProgress_(ss, 'Removing duplicates...');
    finalRows = removeManagedDuplicates_(finalRows);

    setProgress_(ss, `Writing ${finalRows.length} row(s)...`);
    writeVisibleBody_(sheet, finalRows);
    writeStateBody_(stateSheet, finalRows);
    applyNumberFormats_(sheet);
    applyRowColors_(sheet, finalRows);
    ensureTableRange_(spreadsheetId, sheet);

    saveSyncTokens_(fetchResult.nextSyncTokens);

    if (changedNotifications.length > 0) {
      setProgress_(ss, `Done. ${changedNotifications.length} changed invoiced event(s) detected.`);
      SpreadsheetApp.getUi().alert(
        'Changed invoiced events detected',
        buildChangedRowsMessage_(changedNotifications),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      ss.toast(
        `${changedNotifications.length} changed invoiced event(s) detected.`,
        CONFIG.toastTitle,
        8
      );
    } else {
      setProgress_(ss, 'Done.');
      ss.toast('Calendar import finished.', CONFIG.toastTitle, 5);
    }
  } finally {
    restoreUiState_(ss, uiState);
    lock.releaseLock();
  }
}

/* ============================================================================
 * Progress
 * ========================================================================== */

function setProgress_(ss, message) {
  ss.toast(message, CONFIG.toastTitle, 5);

  const sheet = ss.getSheetByName(CONFIG.sheetName);
  if (sheet) {
    sheet.getRange(CONFIG.statusCell).setValue(message);
  }

  SpreadsheetApp.flush();
}

/* ============================================================================
 * Scope
 * ========================================================================== */

function buildScope_() {
  const importStart = parseImportStartDate_(CONFIG.importStartDate);
  const now = new Date();

  return {
    importStart,
    now,
    importStartMillis: importStart.getTime(),
    nowMillis: now.getTime(),
  };
}

function parseImportStartDate_(value) {
  if (!value || !/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    throw new Error(`Invalid CONFIG.importStartDate: ${value}`);
  }

  return new Date(`${value}T00:00:00`);
}

function isManagedEventInScope_(eventObj, scope) {
  return (
    eventObj &&
    eventObj.start instanceof Date &&
    eventObj.end instanceof Date &&
    !Number.isNaN(eventObj.start.getTime()) &&
    !Number.isNaN(eventObj.end.getTime()) &&
    eventObj.start.getTime() >= scope.importStartMillis &&
    eventObj.end.getTime() <= scope.nowMillis
  );
}

function isExistingRowInScope_(rowValues, scope) {
  const start = rowValues[3];
  const end = rowValues[4];

  if (!(start instanceof Date) || Number.isNaN(start.getTime())) {
    return false;
  }
  if (!(end instanceof Date) || Number.isNaN(end.getTime())) {
    return false;
  }

  return (
    start.getTime() >= scope.importStartMillis &&
    end.getTime() <= scope.nowMillis
  );
}

/* ============================================================================
 * Calendar access
 * ========================================================================== */

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

/* ============================================================================
 * Existing state
 * ========================================================================== */

function readExistingState_(sheet, stateSheet, timeZone, scope) {
  const visibleLastRow = sheet.getLastRow();
  const stateLastRow = stateSheet.getLastRow();
  const visibleRowCount = Math.max(visibleLastRow - 1, 0);
  const stateRowCount = Math.max(stateLastRow - 1, 0);
  const rowCount = Math.max(visibleRowCount, stateRowCount);

  const result = {
    hasManagedRows: false,
    rowsByEventKey: new Map(),
    ignoredManagedRows: [],
    unmanagedRows: [],
  };

  if (rowCount === 0) {
    return result;
  }

  const visibleValues = visibleRowCount > 0
    ? sheet.getRange(2, 1, visibleRowCount, CONFIG.header.length).getValues()
    : [];

  const stateValues = stateRowCount > 0
    ? stateSheet.getRange(2, 1, stateRowCount, CONFIG.stateHeader.length).getValues()
    : [];

  for (let i = 0; i < rowCount; i += 1) {
    const rowValues = i < visibleValues.length
      ? visibleValues[i].slice()
      : new Array(CONFIG.header.length).fill('');

    const stateRow = i < stateValues.length
      ? stateValues[i]
      : ['', ''];

    const eventKey = toText_(stateRow[0]);
    const rowKind = toText_(stateRow[1]) || CONFIG.rowKind.unmanaged;
    const invoiceNumber = toText_(rowValues[8]);
    const signature = buildSheetRowSignature_(rowValues, timeZone);

    if (eventKey) {
      result.hasManagedRows = true;

      const row = {
        eventKey,
        rowKind,
        invoiceNumber,
        signature,
        values: rowValues,
      };

      if (isExistingRowInScope_(rowValues, scope)) {
        if (!result.rowsByEventKey.has(eventKey)) {
          result.rowsByEventKey.set(eventKey, []);
        }
        result.rowsByEventKey.get(eventKey).push(row);
      } else {
        result.ignoredManagedRows.push(row);
      }
    } else if (!isCompletelyBlankRow_(rowValues)) {
      result.unmanagedRows.push({
        syntheticKey: `__UNMANAGED__${i + 2}`,
        rowKind: CONFIG.rowKind.unmanaged,
        invoiceNumber,
        signature,
        values: rowValues,
      });
    }
  }

  return result;
}

/* ============================================================================
 * Rebuild logic
 * ========================================================================== */

function rebuildFromFullSnapshot_(existingState, currentByKey, changedNotifications, timeZone, scope) {
  const groups = [];
  const processedKeys = new Set();

  const sortedCurrent = Array.from(currentByKey.values()).sort(compareImportedEvents_);

  sortedCurrent.forEach((currentEvent) => {
    const existingRows = existingState.rowsByEventKey.get(currentEvent.eventKey) || [];
    const group = buildGroupForCurrentEvent_(existingRows, currentEvent, changedNotifications, timeZone);
    if (group) {
      groups.push(group);
    }
    processedKeys.add(currentEvent.eventKey);
  });

  existingState.rowsByEventKey.forEach((existingRows, eventKey) => {
    if (processedKeys.has(eventKey)) {
      return;
    }

    const group = buildGroupForMissingCurrent_(existingRows);
    if (group) {
      groups.push(group);
    }
  });

  existingState.ignoredManagedRows.forEach((row) => {
    groups.push(buildGroupFromRows_([cloneRowModel_(row)]));
  });

  existingState.unmanagedRows.forEach((row) => {
    groups.push(buildGroupFromRows_([cloneRowModel_(row)]));
  });

  groups.sort(compareGroups_);
  return flattenGroups_(groups);
}

function rebuildFromIncremental_(existingState, deltaByKey, changedNotifications, timeZone, scope) {
  const groupsByKey = new Map();

  existingState.rowsByEventKey.forEach((rows, eventKey) => {
    const cloned = rows.map((row) => cloneRowModel_(row));
    const group = buildGroupFromRows_(cloned);
    if (group) {
      groupsByKey.set(eventKey, group);
    }
  });

  existingState.ignoredManagedRows.forEach((row) => {
    groupsByKey.set(`__IGNORED__${row.eventKey}`, buildGroupFromRows_([cloneRowModel_(row)]));
  });

  existingState.unmanagedRows.forEach((row) => {
    groupsByKey.set(row.syntheticKey, buildGroupFromRows_([cloneRowModel_(row)]));
  });

  deltaByKey.forEach((currentEvent, eventKey) => {
    const existingRows = existingState.rowsByEventKey.get(eventKey) || [];

    if (currentEvent && !isManagedEventInScope_(currentEvent, scope)) {
      return;
    }

    let group = null;
    if (currentEvent) {
      group = buildGroupForCurrentEvent_(existingRows, currentEvent, changedNotifications, timeZone);
    } else {
      group = buildGroupForMissingCurrent_(existingRows);
    }

    if (group && group.rows.length > 0) {
      groupsByKey.set(eventKey, group);
    } else {
      groupsByKey.delete(eventKey);
    }
  });

  const groups = Array.from(groupsByKey.values()).filter(Boolean);
  groups.sort(compareGroups_);
  return flattenGroups_(groups);
}

function buildGroupForCurrentEvent_(existingRows, currentEvent, changedNotifications, timeZone) {
  const invoicedRows = existingRows
    .filter((row) => !!row.invoiceNumber)
    .map((row) => cloneRowModel_(row));

  const nonInvoicedRows = existingRows
    .filter((row) => !row.invoiceNumber)
    .map((row) => cloneRowModel_(row));

  if (nonInvoicedRows.length > 0) {
    const target = nonInvoicedRows[nonInvoicedRows.length - 1];
    const updated = buildUpdatedRowFromImport_(target, currentEvent);
    return buildGroupFromRows_(invoicedRows.concat([updated]));
  }

  if (invoicedRows.length > 0) {
    const latest = invoicedRows[invoicedRows.length - 1];

    if (latest.signature !== currentEvent.signature) {
      const changedRow = buildNewRowFromImport_(
        currentEvent,
        latest.values[6] || '',
        latest.values[7] || '',
        '',
        '',
        CONFIG.rowKind.changedCopy
      );

      changedNotifications.push({
        calendar: currentEvent.calendarName,
        title: currentEvent.title,
        date: formatDateCell_(currentEvent.date, timeZone, 'yyyy-MM-dd'),
        start: formatDateCell_(currentEvent.start, timeZone, 'HH:mm'),
        end: formatDateCell_(currentEvent.end, timeZone, 'HH:mm'),
      });

      return buildGroupFromRows_(invoicedRows.concat([changedRow]));
    }

    return buildGroupFromRows_(invoicedRows);
  }

  return buildGroupFromRows_([
    buildNewRowFromImport_(currentEvent, '', '', '', '', CONFIG.rowKind.normal),
  ]);
}

function buildGroupForMissingCurrent_(existingRows) {
  const invoicedRows = existingRows
    .filter((row) => !!row.invoiceNumber)
    .map((row) => cloneRowModel_(row));

  if (invoicedRows.length === 0) {
    return null;
  }

  return buildGroupFromRows_(invoicedRows);
}

function buildGroupFromRows_(rows) {
  if (!rows || rows.length === 0) {
    return null;
  }

  return {
    anchor: extractAnchorFromRow_(rows[0]),
    rows,
  };
}

function flattenGroups_(groups) {
  const rows = [];
  groups.forEach((group) => {
    group.rows.forEach((row) => rows.push(row));
  });
  return rows;
}

function compareGroups_(a, b) {
  const byDate = a.anchor.date - b.anchor.date;
  if (byDate !== 0) {
    return byDate;
  }

  const byStart = a.anchor.start - b.anchor.start;
  if (byStart !== 0) {
    return byStart;
  }

  const byEnd = a.anchor.end - b.anchor.end;
  if (byEnd !== 0) {
    return byEnd;
  }

  return a.anchor.text.localeCompare(b.anchor.text);
}

function compareImportedEvents_(a, b) {
  const byDate = a.date.getTime() - b.date.getTime();
  if (byDate !== 0) {
    return byDate;
  }

  const byStart = a.start.getTime() - b.start.getTime();
  if (byStart !== 0) {
    return byStart;
  }

  const byEnd = a.end.getTime() - b.end.getTime();
  if (byEnd !== 0) {
    return byEnd;
  }

  return a.eventKey.localeCompare(b.eventKey);
}

function buildUpdatedRowFromImport_(existingRow, currentEvent) {
  const values = currentEvent.values.slice();
  values[6] = existingRow.values[6] || '';
  values[7] = existingRow.values[7] || '';
  values[8] = existingRow.values[8] || '';
  values[9] = existingRow.values[9] || '';

  return {
    eventKey: currentEvent.eventKey,
    rowKind: CONFIG.rowKind.normal,
    invoiceNumber: toText_(values[8]),
    signature: currentEvent.signature,
    values,
  };
}

function buildNewRowFromImport_(currentEvent, customer, project, invoiceNumber, invoiceDate, rowKind) {
  const values = currentEvent.values.slice();
  values[6] = customer || '';
  values[7] = project || '';
  values[8] = invoiceNumber || '';
  values[9] = invoiceDate || '';

  return {
    eventKey: currentEvent.eventKey,
    rowKind,
    invoiceNumber: toText_(values[8]),
    signature: currentEvent.signature,
    values,
  };
}

function extractAnchorFromRow_(row) {
  const date = row.values[2] instanceof Date ? row.values[2].getTime() : Number.MAX_SAFE_INTEGER;
  const start = row.values[3] instanceof Date ? row.values[3].getTime() : Number.MAX_SAFE_INTEGER;
  const end = row.values[4] instanceof Date ? row.values[4].getTime() : Number.MAX_SAFE_INTEGER;

  return {
    date,
    start,
    end,
    text: `${toText_(row.values[0])}|${toText_(row.values[1])}|${toText_(row.eventKey || row.syntheticKey || '')}`,
  };
}

/* ============================================================================
 * Duplicate handling
 * ========================================================================== */

function removeManagedDuplicates_(rows) {
  const groups = new Map();

  rows.forEach((row, index) => {
    if (!isRowManagedInScopeForDuplicateCheck_(row)) {
      return;
    }

    const key = buildDuplicateKey_(row.values);
    if (!groups.has(key)) {
      groups.set(key, []);
    }

    groups.get(key).push({
      index,
      row,
      calendar: toText_(row.values[0]),
    });
  });

  const removeIndexes = new Set();

  groups.forEach((entries) => {
    if (entries.length < 2) {
      return;
    }

    const byCalendar = new Map();

    entries.forEach((entry) => {
      if (!byCalendar.has(entry.calendar)) {
        byCalendar.set(entry.calendar, []);
      }
      byCalendar.get(entry.calendar).push(entry);
    });

    byCalendar.forEach((calendarEntries) => {
      if (calendarEntries.length > 1) {
        calendarEntries.forEach((entry) => removeIndexes.add(entry.index));
      }
    });

    const survivorsAfterSameCalendarPurge = entries.filter(
      (entry) => !removeIndexes.has(entry.index)
    );

    if (survivorsAfterSameCalendarPurge.length < 2) {
      return;
    }

    const specificEntries = survivorsAfterSameCalendarPurge.filter(
      (entry) => entry.calendar !== CONFIG.defaultCalendarName
    );

    let survivors = survivorsAfterSameCalendarPurge;

    if (specificEntries.length > 0) {
      survivorsAfterSameCalendarPurge
        .filter((entry) => entry.calendar === CONFIG.defaultCalendarName)
        .forEach((entry) => removeIndexes.add(entry.index));

      survivors = specificEntries;
    }

    if (survivors.length < 2) {
      return;
    }

    survivors.sort((a, b) => compareDuplicateCandidates_(a.row, b.row));
    survivors.slice(1).forEach((entry) => removeIndexes.add(entry.index));
  });

  return rows.filter((row, index) => !removeIndexes.has(index));
}

function isRowManagedInScopeForDuplicateCheck_(row) {
  return !!row.eventKey;
}

function buildDuplicateKey_(rowValues) {
  return JSON.stringify({
    event: normalizeDuplicateText_(rowValues[1]),
    date: duplicateDateToken_(rowValues[2]),
    start: duplicateDateToken_(rowValues[3]),
    end: duplicateDateToken_(rowValues[4]),
  });
}

function normalizeDuplicateText_(value) {
  return toText_(value).trim();
}

function duplicateDateToken_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.getTime();
  }

  return toText_(value);
}

function compareDuplicateCandidates_(a, b) {
  const aPriority = getCalendarPriority_(toText_(a.values[0]));
  const bPriority = getCalendarPriority_(toText_(b.values[0]));
  if (aPriority !== bPriority) {
    return aPriority - bPriority;
  }

  const aInvoice = a.invoiceNumber ? 0 : 1;
  const bInvoice = b.invoiceNumber ? 0 : 1;
  if (aInvoice !== bInvoice) {
    return aInvoice - bInvoice;
  }

  const aChanged = a.rowKind === CONFIG.rowKind.changedCopy ? 1 : 0;
  const bChanged = b.rowKind === CONFIG.rowKind.changedCopy ? 1 : 0;
  if (aChanged !== bChanged) {
    return aChanged - bChanged;
  }

  return 0;
}

function getCalendarPriority_(calendarName) {
  const nonDefault = CONFIG.calendarNames.filter(
    (name) => name !== CONFIG.defaultCalendarName
  );
  const order = nonDefault.concat([CONFIG.defaultCalendarName]);
  const index = order.indexOf(calendarName);
  return index >= 0 ? index : Number.MAX_SAFE_INTEGER;
}

/* ============================================================================
 * Writing
 * ========================================================================== */

function writeVisibleBody_(sheet, rows) {
  const oldLastRow = sheet.getLastRow();
  const neededLastRow = Math.max(rows.length + 1, 1);

  if (sheet.getMaxRows() < neededLastRow) {
    sheet.insertRowsAfter(sheet.getMaxRows(), neededLastRow - sheet.getMaxRows());
  }

  const clearRows = Math.max(oldLastRow, neededLastRow) - 1;
  if (clearRows > 0) {
    const clearRange = sheet.getRange(2, 1, clearRows, CONFIG.header.length);
    clearRange.clearContent();
    clearRange.setFontColor(CONFIG.colors.normal);
  }

  if (rows.length > 0) {
    const values = rows.map((row) => row.values);
    sheet.getRange(2, 1, rows.length, CONFIG.header.length).setValues(values);
  }
}

function writeStateBody_(stateSheet, rows) {
  const oldLastRow = stateSheet.getLastRow();
  const neededLastRow = Math.max(rows.length + 1, 1);

  if (stateSheet.getMaxRows() < neededLastRow) {
    stateSheet.insertRowsAfter(stateSheet.getMaxRows(), neededLastRow - stateSheet.getMaxRows());
  }

  const clearRows = Math.max(oldLastRow, neededLastRow) - 1;
  if (clearRows > 0) {
    stateSheet.getRange(2, 1, clearRows, CONFIG.stateHeader.length).clearContent();
  }

  if (rows.length > 0) {
    const values = rows.map((row) => [
      row.eventKey || '',
      row.rowKind || CONFIG.rowKind.unmanaged,
    ]);
    stateSheet.getRange(2, 1, rows.length, CONFIG.stateHeader.length).setValues(values);
  }
}

function applyNumberFormats_(sheet) {
  const lastRow = sheet.getLastRow();
  const rowCount = Math.max(lastRow - 1, 0);

  if (rowCount === 0) {
    return;
  }

  sheet.getRange(2, 3, rowCount, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 4, rowCount, 1).setNumberFormat('hh:mm');
  sheet.getRange(2, 5, rowCount, 1).setNumberFormat('hh:mm');
  sheet.getRange(2, 6, rowCount, 1).setNumberFormat('hh:mm');
  sheet.getRange(2, 10, rowCount, 1).setNumberFormat('yyyy-mm-dd');
}

function applyRowColors_(sheet, rows) {
  if (rows.length === 0) {
    return;
  }

  const colors = rows.map((row) => {
    let color = CONFIG.colors.normal;

    if (row.invoiceNumber) {
      color = CONFIG.colors.invoiced;
    } else if (row.rowKind === CONFIG.rowKind.changedCopy) {
      color = CONFIG.colors.changed;
    }

    return new Array(CONFIG.header.length).fill(color);
  });

  sheet.getRange(2, 1, rows.length, CONFIG.header.length).setFontColors(colors);
}

/* ============================================================================
 * Sync tokens
 * ========================================================================== */

function loadSyncTokens_(calendars) {
  const props = PropertiesService.getDocumentProperties();

  return calendars.map((calendarInfo) => {
    return {
      calendarId: calendarInfo.id,
      syncToken: props.getProperty(CONFIG.propertyPrefix + calendarInfo.id) || '',
    };
  });
}

function saveSyncTokens_(tokensByCalendarId) {
  const props = PropertiesService.getDocumentProperties();
  const payload = {};

  Object.keys(tokensByCalendarId).forEach((calendarId) => {
    payload[CONFIG.propertyPrefix + calendarId] = tokensByCalendarId[calendarId] || '';
  });

  props.setProperties(payload, false);
}

function clearSyncTokens_(calendars) {
  const props = PropertiesService.getDocumentProperties();

  calendars.forEach((calendarInfo) => {
    props.deleteProperty(CONFIG.propertyPrefix + calendarInfo.id);
  });
}

/* ============================================================================
 * Sheet and table
 * ========================================================================== */

function ensureCalendarSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.sheetName);
  }

  return sheet;
}

function ensureStateSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.stateSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.stateSheetName);
  }

  if (!sheet.isSheetHidden()) {
    sheet.hideSheet();
  }

  return sheet;
}

function ensureHeader_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, CONFIG.header.length);
  const current = headerRange.getValues()[0];

  const needsWrite = CONFIG.header.some((value, index) => current[index] !== value);
  if (needsWrite) {
    headerRange.setValues([CONFIG.header]);
  }
}

function ensureStateHeader_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, CONFIG.stateHeader.length);
  const current = headerRange.getValues()[0];

  const needsWrite = CONFIG.stateHeader.some((value, index) => current[index] !== value);
  if (needsWrite) {
    headerRange.setValues([CONFIG.stateHeader]);
  }
}

function ensureSheetFormatting_(sheet) {
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(true);
}

function ensureTable_(spreadsheetId, sheet) {
  const spreadsheetModel = getSpreadsheetModel_(spreadsheetId);
  const sheetModel = (spreadsheetModel.sheets || []).find(
    (entry) => entry.properties && entry.properties.sheetId === sheet.getSheetId()
  );
  const tables = (sheetModel && sheetModel.tables) || [];

  if (tables.length === 0) {
    Sheets.Spreadsheets.batchUpdate(
      {
        requests: [
          {
            addTable: {
              table: {
                name: CONFIG.tableName,
                range: {
                  sheetId: sheet.getSheetId(),
                  startRowIndex: 0,
                  endRowIndex: Math.max(sheet.getLastRow(), 1),
                  startColumnIndex: 0,
                  endColumnIndex: CONFIG.header.length,
                },
              },
            },
          },
        ],
      },
      spreadsheetId
    );
    return;
  }

  const table = tables.find((entry) => entry.name === CONFIG.tableName) || tables[0];

  Sheets.Spreadsheets.batchUpdate(
    {
      requests: [
        {
          updateTable: {
            table: {
              tableId: table.tableId,
              name: CONFIG.tableName,
              range: {
                sheetId: sheet.getSheetId(),
                startRowIndex: 0,
                endRowIndex: Math.max(sheet.getLastRow(), 1),
                startColumnIndex: 0,
                endColumnIndex: CONFIG.header.length,
              },
            },
            fields: 'name,range',
          },
        },
      ],
    },
    spreadsheetId
  );
}

function ensureTableRange_(spreadsheetId, sheet) {
  const spreadsheetModel = getSpreadsheetModel_(spreadsheetId);
  const sheetModel = (spreadsheetModel.sheets || []).find(
    (entry) => entry.properties && entry.properties.sheetId === sheet.getSheetId()
  );
  const table = ((sheetModel && sheetModel.tables) || []).find(
    (entry) => entry.name === CONFIG.tableName
  );

  if (!table) {
    ensureTable_(spreadsheetId, sheet);
    return;
  }

  const desiredEndRow = Math.max(sheet.getLastRow(), 1);
  const currentRange = table.range || {};

  const unchanged =
    currentRange.sheetId === sheet.getSheetId() &&
    currentRange.startRowIndex === 0 &&
    currentRange.endRowIndex === desiredEndRow &&
    currentRange.startColumnIndex === 0 &&
    currentRange.endColumnIndex === CONFIG.header.length;

  if (unchanged) {
    return;
  }

  Sheets.Spreadsheets.batchUpdate(
    {
      requests: [
        {
          updateTable: {
            table: {
              tableId: table.tableId,
              name: table.name,
              range: {
                sheetId: sheet.getSheetId(),
                startRowIndex: 0,
                endRowIndex: desiredEndRow,
                startColumnIndex: 0,
                endColumnIndex: CONFIG.header.length,
              },
            },
            fields: 'range',
          },
        },
      ],
    },
    spreadsheetId
  );
}

function getSpreadsheetModel_(spreadsheetId) {
  return Sheets.Spreadsheets.get(spreadsheetId, {
    fields: 'sheets(properties(sheetId,title),tables(tableId,name,range))',
    includeGridData: false,
  });
}

/* ============================================================================
 * UI state
 * ========================================================================== */

function captureUiState_(ss) {
  const activeSheet = ss.getActiveSheet();
  const activeRange = ss.getActiveRange();

  return {
    sheetName: activeSheet ? activeSheet.getName() : null,
    a1Notation: activeRange ? activeRange.getA1Notation() : null,
  };
}

function restoreUiState_(ss, uiState) {
  if (!uiState || !uiState.sheetName) {
    return;
  }

  const sheet = ss.getSheetByName(uiState.sheetName);
  if (!sheet) {
    return;
  }

  ss.setActiveSheet(sheet, false);

  if (uiState.a1Notation) {
    try {
      sheet.setActiveRange(sheet.getRange(uiState.a1Notation));
    } catch (error) {
      // Ignore invalid range restoration.
    }
  }
}

/* ============================================================================
 * Helpers
 * ========================================================================== */

function buildImportedSignature_(eventObj, timeZone) {
  return JSON.stringify({
    calendar: eventObj.calendar,
    title: eventObj.title,
    date: formatDateCell_(eventObj.date, timeZone, 'yyyy-MM-dd'),
    start: formatDateCell_(eventObj.start, timeZone, "yyyy-MM-dd'T'HH:mm"),
    end: formatDateCell_(eventObj.end, timeZone, "yyyy-MM-dd'T'HH:mm"),
    duration: normalizeDuration_(eventObj.duration),
  });
}

function buildSheetRowSignature_(rowValues, timeZone) {
  return JSON.stringify({
    calendar: toText_(rowValues[0]),
    title: toText_(rowValues[1]),
    date: formatDateCell_(rowValues[2], timeZone, 'yyyy-MM-dd'),
    start: formatDateCell_(rowValues[3], timeZone, "yyyy-MM-dd'T'HH:mm"),
    end: formatDateCell_(rowValues[4], timeZone, "yyyy-MM-dd'T'HH:mm"),
    duration: normalizeDuration_(rowValues[5]),
  });
}

function normalizeDuration_(value) {
  if (typeof value === 'number') {
    return value.toFixed(10);
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return ((value.getTime() % 86400000) / 86400000).toFixed(10);
  }

  return toText_(value);
}

function toSheetDateOnly_(dateValue, timeZone) {
  const iso = Utilities.formatDate(dateValue, timeZone, "yyyy-MM-dd'T'12:00:00XXX");
  return new Date(iso);
}

function formatDateCell_(value, timeZone, pattern) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, timeZone, pattern);
  }

  return toText_(value);
}

function toText_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value);
}

function isCompletelyBlankRow_(rowValues) {
  return rowValues.every((value) => toText_(value) === '');
}

function cloneRowModel_(row) {
  return {
    eventKey: row.eventKey || '',
    syntheticKey: row.syntheticKey || '',
    rowKind: row.rowKind || CONFIG.rowKind.normal,
    invoiceNumber: row.invoiceNumber || '',
    signature: row.signature || '',
    values: row.values.slice(),
  };
}

function buildChangedRowsMessage_(items) {
  const lines = items.slice(0, 25).map((item) => {
    return `${item.calendar} | ${item.date} ${item.start}-${item.end} | ${item.title}`;
  });

  if (items.length > 25) {
    lines.push(`... and ${items.length - 25} more.`);
  }

  return lines.join('\n');
}
