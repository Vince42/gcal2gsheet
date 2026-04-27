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
      } else if (isExistingRowBeforeImportStart_(rowValues, scope)) {
        result.ignoredManagedRows.push(row);
      } else {
        // Preserve managed rows that are currently outside active scope (for example,
        // after manual date edits or invalid date cell states) to avoid data loss.
        result.ignoredManagedRows.push(row);
      }
    } else if (!isCompletelyBlankRow_(rowValues)) {
      const calendarName = toText_(rowValues[0]);
      const looksLikeImportedCalendarRow = CONFIG.calendarNames.includes(calendarName);
      const isFutureRow = isExistingRowAfterNow_(rowValues, scope);

      if (looksLikeImportedCalendarRow && isFutureRow) {
        continue;
      }

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
