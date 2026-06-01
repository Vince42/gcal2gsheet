function filterCalendarForOpen() {
  filterCalendarByStatus_('Open');
}

function filterCalendarForInvoiced() {
  filterCalendarByStatus_('Invoiced');
}

function filterCalendarForNonBillable() {
  filterCalendarByStatus_('Non-billable');
}

function markVisibleCalendarRowsAsInvoiced() {
  refreshConfig_();
  markVisibleCalendarRows_(CONFIG.invoicingSheetName);
}

function markVisibleCalendarRowsAsNonBillable() {
  refreshConfig_();
  markVisibleCalendarRows_(CONFIG.nonBillableSheetName);
}

function filterCalendarByStatus_(statusValue) {
  refreshConfig_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  const managedSheets = ensureManagedWorkbookStructure_(ss, spreadsheetId);
  const sheet = managedSheets.sheet;

  const statusColumn = CONFIG.header.indexOf('Status') + 1;
  if (statusColumn <= 0) {
    throw new Error('Calendar Status column is not configured.');
  }

  clearManualCalendarStatusFilter_(sheet);

  try {
    const range = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), CONFIG.header.length);
    let filter = sheet.getFilter();
    if (!filter) {
      filter = range.createFilter();
    }

    const criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo(statusValue).build();
    filter.setColumnFilterCriteria(statusColumn, criteria);
  } catch (error) {
    applyManualCalendarStatusFilter_(sheet, statusValue, statusColumn);
  }

  showToastMessage_(ss, `Filtered Calendar for ${statusValue}.`, { severity: 'info' });
}

function clearManualCalendarStatusFilter_(sheet) {
  const rowCount = Math.max(sheet.getLastRow() - 1, 0);
  if (rowCount > 0) {
    sheet.showRows(2, rowCount);
  }
}

function applyManualCalendarStatusFilter_(sheet, statusValue, statusColumn) {
  const rowCount = Math.max(sheet.getLastRow() - 1, 0);
  if (rowCount === 0) {
    return;
  }

  const statuses = sheet.getRange(2, statusColumn, rowCount, 1).getDisplayValues();
  statuses.forEach((row, index) => {
    if (row[0] !== statusValue) {
      sheet.hideRows(index + 2);
    }
  });
}

function markVisibleCalendarRows_(targetSheetName) {
  refreshConfig_();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = ss.getId();
    const managedSheets = ensureManagedWorkbookStructure_(ss, spreadsheetId);
    const visibleRows = collectVisibleCalendarRows_(managedSheets.sheet, managedSheets.stateSheet);

    if (visibleRows.length === 0) {
      showToastMessage_(ss, 'No visible Calendar rows to mark.', { severity: 'info' });
      return;
    }

    let markedCount = 0;
    if (targetSheetName === CONFIG.invoicingSheetName) {
      markedCount = appendCalendarRowsToInvoicing_(
        visibleRows,
        managedSheets.invoicingSheet,
        managedSheets.invoicingStateSheet
      );
      removeRegisterRowsByEventKeys_(
        managedSheets.nonBillableSheet,
        managedSheets.nonBillableStateSheet,
        CONFIG.nonBillableHeader,
        CONFIG.nonBillableStateHeader,
        visibleRows.map((row) => row.eventKey)
      );
    } else if (targetSheetName === CONFIG.nonBillableSheetName) {
      markedCount = appendCalendarRowsToNonBillable_(
        visibleRows,
        managedSheets.nonBillableSheet,
        managedSheets.nonBillableStateSheet
      );
      removeRegisterRowsByEventKeys_(
        managedSheets.invoicingSheet,
        managedSheets.invoicingStateSheet,
        CONFIG.invoicingHeader,
        CONFIG.invoicingStateHeader,
        visibleRows.map((row) => row.eventKey)
      );
    } else {
      throw new Error(`Unsupported mark target: ${targetSheetName}`);
    }

    applyNumberFormats_(managedSheets.invoicingSheet, CONFIG.invoicingHeader);
    applyNumberFormats_(managedSheets.nonBillableSheet, CONFIG.nonBillableHeader);
    ensureTableRange_(spreadsheetId, managedSheets.invoicingSheet, CONFIG.invoicingTableName, CONFIG.invoicingHeader);
    ensureTableRange_(
      spreadsheetId,
      managedSheets.nonBillableSheet,
      CONFIG.nonBillableTableName,
      CONFIG.nonBillableHeader
    );

    SpreadsheetApp.flush();
    showToastMessage_(ss, `${markedCount} visible Calendar row(s) marked as ${targetSheetName}.`, {
      severity: 'info',
    });
  } finally {
    lock.releaseLock();
  }
}

function collectVisibleCalendarRows_(sheet, stateSheet) {
  const rowCount = Math.min(
    Math.max(sheet.getLastRow() - 1, 0),
    Math.max(stateSheet.getLastRow() - 1, 0)
  );
  if (rowCount === 0) {
    return [];
  }

  const values = sheet.getRange(2, 1, rowCount, CONFIG.header.length).getValues();
  const stateValues = stateSheet.getRange(2, 1, rowCount, CONFIG.stateHeader.length).getValues();
  const visibleRows = [];

  for (let index = 0; index < rowCount; index += 1) {
    const sheetRow = index + 2;
    if (sheet.isRowHiddenByFilter(sheetRow) || sheet.isRowHiddenByUser(sheetRow)) {
      continue;
    }

    const eventKey = toText_(stateValues[index][0]);
    if (!eventKey || isCompletelyBlankRow_(values[index])) {
      continue;
    }

    visibleRows.push({
      eventKey,
      values: values[index].slice(),
    });
  }

  return visibleRows;
}

function appendCalendarRowsToInvoicing_(calendarRows, invoicingSheet, invoicingStateSheet) {
  const invoiceStore = readInvoicingState_(invoicingSheet, invoicingStateSheet);
  const appendValues = [];
  const appendStateValues = [];

  calendarRows.forEach((row) => {
    if (invoiceStore.byEventKey.has(row.eventKey)) {
      return;
    }

    appendValues.push([
      row.values[0],
      row.values[1],
      row.values[2],
      row.values[3],
      row.values[4],
      row.values[5],
      '',
      '',
      '',
      '',
    ]);
    appendStateValues.push([row.eventKey]);
  });

  appendInvoicingRows_(invoicingSheet, invoicingStateSheet, appendValues, appendStateValues);
  return appendValues.length;
}

function appendCalendarRowsToNonBillable_(calendarRows, nonBillableSheet, nonBillableStateSheet) {
  const nonBillableStore = readNonBillableState_(nonBillableSheet, nonBillableStateSheet);
  const appendValues = [];
  const appendStateValues = [];

  calendarRows.forEach((row) => {
    if (nonBillableStore.byEventKey.has(row.eventKey)) {
      return;
    }

    appendValues.push([
      row.values[0],
      row.values[1],
      row.values[2],
      row.values[3],
      row.values[4],
      row.values[5],
      '',
    ]);
    appendStateValues.push([row.eventKey]);
  });

  appendNonBillableRows_(nonBillableSheet, nonBillableStateSheet, appendValues, appendStateValues);
  return appendValues.length;
}

function removeRegisterRowsByEventKeys_(sheet, stateSheet, header, stateHeader, eventKeys) {
  const keys = new Set((eventKeys || []).filter((key) => key));
  if (keys.size === 0) {
    return 0;
  }

  const rowCount = Math.max(stateSheet.getLastRow() - 1, 0);
  if (rowCount === 0) {
    return 0;
  }

  const stateValues = stateSheet.getRange(2, 1, rowCount, stateHeader.length).getValues();
  let removedCount = 0;

  for (let index = stateValues.length - 1; index >= 0; index -= 1) {
    const eventKey = toText_(stateValues[index][0]);
    if (!keys.has(eventKey)) {
      continue;
    }

    sheet.getRange(index + 2, 1, 1, header.length).deleteCells(SpreadsheetApp.Dimension.ROWS);
    stateSheet.getRange(index + 2, 1, 1, stateHeader.length).deleteCells(SpreadsheetApp.Dimension.ROWS);
    removedCount += 1;
  }

  return removedCount;
}
