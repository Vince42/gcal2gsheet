function readExistingState_(sheet, timeZone, scope, invoiceStore, nonBillableStore) {
  const visibleRowCount = Math.max(sheet.getLastRow() - 1, 0);

  const result = {
    hasManagedRows: false,
    rowsByEventKey: new Map(),
    ignoredManagedRows: [],
    ignoredBeforeImportStartCount: 0,
    unmanagedRows: [],
  };

  if (visibleRowCount === 0) {
    return result;
  }

  const visibleValues = sheet.getRange(2, 1, visibleRowCount, CONFIG.header.length).getValues();

  for (let i = 0; i < visibleValues.length; i += 1) {
    const fullRowValues = visibleValues[i].slice();
    const eventKey = toText_(fullRowValues[0]);
    const rowValues = fullRowValues.slice(1);
    const rowKind = CONFIG.rowKind.normal;
    const invoiceNumber = getRegisterStatusMarkerForEventKey_(eventKey, invoiceStore, nonBillableStore);
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
        result.ignoredBeforeImportStartCount += 1;
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
      logStorageDebug_(
        'load-sync-tokens',
        `Ignored denied read for calendarId "${calendarInfo.id}": ${error}`
      );
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
    logStorageDebug_('save-sync-tokens', `Ignored denied write while saving sync tokens: ${error}`);
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
      logStorageDebug_(
        'clear-sync-tokens',
        `Ignored denied delete for calendarId "${calendarInfo.id}": ${error}`
      );
    }
  });
}

function getRegisterStatusMarkerForEventKey_(eventKey, invoiceStore, nonBillableStore) {
  if (!eventKey) {
    return '';
  }

  const invoice = invoiceStore && invoiceStore.byEventKey
    ? invoiceStore.byEventKey.get(eventKey)
    : null;
  if (invoice) {
    return invoice.invoiceNumber || 'INVOICED';
  }

  const nonBillable = nonBillableStore && nonBillableStore.byEventKey
    ? nonBillableStore.byEventKey.get(eventKey)
    : null;
  if (nonBillable) {
    return 'NON_BILLABLE';
  }

  return '';
}
