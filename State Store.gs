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
  const props = getConfigPropertiesStore_();

  return calendars.map((calendarInfo) => {
    let syncToken = '';
    try {
      syncToken = props.getProperty(CONFIG.propertyPrefix + calendarInfo.id) || '';
    } catch (error) {
      if (!isPermissionDeniedError_(error)) {
        throw error;
      }
    }
    return {
      calendarId: calendarInfo.id,
      syncToken,
    };
  });
}

function saveSyncTokens_(tokensByCalendarId) {
  const props = getConfigPropertiesStore_();
  const payload = {};

  Object.keys(tokensByCalendarId).forEach((calendarId) => {
    payload[CONFIG.propertyPrefix + calendarId] = tokensByCalendarId[calendarId] || '';
  });

  try {
    props.setProperties(payload, false);
  } catch (error) {
    if (!isPermissionDeniedError_(error)) {
      throw error;
    }
  }
}

function clearSyncTokens_(calendars) {
  const props = getConfigPropertiesStore_();

  calendars.forEach((calendarInfo) => {
    try {
      props.deleteProperty(CONFIG.propertyPrefix + calendarInfo.id);
    } catch (error) {
      if (!isPermissionDeniedError_(error)) {
        throw error;
      }
    }
  });
}
