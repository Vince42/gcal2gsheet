function filterCalendarForOpen() {
  filterCalendarByStatus_('Open');
}

function filterCalendarForInvoiced() {
  filterCalendarByStatus_('Invoicing');
}

function filterCalendarForNonBillable() {
  filterCalendarByStatus_('Non-billable');
}

function markSelectedCalendarRowsAsInvoiced() {
  refreshConfig_();
  markSelectedCalendarRows_(CONFIG.invoicingSheetName);
}

function markSelectedCalendarRowsAsNonBillable() {
  refreshConfig_();
  markSelectedCalendarRows_(CONFIG.nonBillableSheetName);
}

function filterCalendarByStatus_(statusValue) {
  refreshConfig_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  const managedSheets = ensureManagedWorkbookStructure_(ss, spreadsheetId);
  const sheet = managedSheets.sheet;

  const statusColumn = CONFIG.header.indexOf('State') + 1;
  if (statusColumn <= 0) {
    throw new Error('Calendar State column is not configured.');
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

function markSelectedCalendarRows_(targetSheetName) {
  refreshConfig_();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = ss.getId();
    setProgress_(ss, `Preparing to mark selected Calendar rows ${getMarkActionLabel_(targetSheetName)}...`);
    const managedSheets = ensureManagedWorkbookStructure_(ss, spreadsheetId);
    const selectedRows = collectSelectedCalendarRows_(ss, managedSheets.sheet);

    if (selectedRows.length === 0) {
      showToastMessage_(ss, 'No selected Calendar rows to mark.', { severity: 'info' });
      return;
    }

    showMarkProgress_(ss, targetSheetName, 0, selectedRows.length, 'Preparing selected rows');

    let markedCount = 0;
    let removedCount = 0;
    if (targetSheetName === CONFIG.invoicingSheetName) {
      markedCount = appendCalendarRowsToInvoicing_(
        selectedRows,
        managedSheets.invoicingSheet,
        (done, total) => showMarkProgress_(ss, targetSheetName, done, total, 'Preparing register rows')
      );
      showMarkProgress_(ss, targetSheetName, selectedRows.length, selectedRows.length, 'Removing moved rows');
      removedCount = removeRegisterRowsByEventKeys_(
        managedSheets.nonBillableSheet,
        CONFIG.nonBillableHeader,
        selectedRows.map((row) => row.eventKey),
        (done, total) => showMarkProgress_(ss, targetSheetName, done, total, 'Removing moved rows')
      );
    } else if (targetSheetName === CONFIG.nonBillableSheetName) {
      markedCount = appendCalendarRowsToNonBillable_(
        selectedRows,
        managedSheets.nonBillableSheet,
        (done, total) => showMarkProgress_(ss, targetSheetName, done, total, 'Preparing register rows')
      );
      showMarkProgress_(ss, targetSheetName, selectedRows.length, selectedRows.length, 'Removing moved rows');
      removedCount = removeRegisterRowsByEventKeys_(
        managedSheets.invoicingSheet,
        CONFIG.invoicingHeader,
        selectedRows.map((row) => row.eventKey),
        (done, total) => showMarkProgress_(ss, targetSheetName, done, total, 'Removing moved rows')
      );
    } else {
      throw new Error(`Unsupported mark target: ${targetSheetName}`);
    }

    showMarkProgress_(ss, targetSheetName, selectedRows.length, selectedRows.length, 'Formatting registers');
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
    const movedMessage = removedCount > 0 ? ` ${removedCount} row(s) removed from the other register.` : '';
    showToastMessage_(
      ss,
      `${markedCount} selected Calendar row(s) marked ${getMarkActionLabel_(targetSheetName)}.${movedMessage}`,
      { severity: 'info' }
    );
  } finally {
    lock.releaseLock();
  }
}

function collectSelectedCalendarRows_(ss, sheet) {
  const rowCount = Math.max(sheet.getLastRow() - 1, 0);
  if (rowCount === 0) {
    return [];
  }

  const selectedRowNumbers = collectSelectedCalendarRowNumbers_(ss, sheet, rowCount);
  if (selectedRowNumbers.length === 0) {
    return [];
  }

  const values = sheet.getRange(2, 1, rowCount, CONFIG.header.length).getValues();
  const selectedRows = [];

  selectedRowNumbers.forEach((sheetRow) => {
    const index = sheetRow - 2;
    const eventKey = toText_(values[index][0]);
    const rowValues = values[index].slice(1);
    if (!eventKey || isCompletelyBlankRow_(rowValues)) {
      return;
    }

    selectedRows.push({
      eventKey,
      values: rowValues,
    });
  });

  return selectedRows;
}

function collectSelectedCalendarRowNumbers_(ss, sheet, rowCount) {
  const ranges = getSelectedCalendarRanges_(ss, sheet);
  const rowNumbers = new Set();
  const firstDataRow = 2;
  const lastDataRow = rowCount + 1;
  const hiddenRowsByGridMetadata = collectHiddenRowsByGridMetadata_(ss, sheet, firstDataRow, lastDataRow);

  ranges.forEach((range) => {
    const startRow = Math.max(range.getRow(), firstDataRow);
    const endRow = Math.min(range.getLastRow(), lastDataRow);
    for (let row = startRow; row <= endRow; row += 1) {
      if (isCalendarSheetRowVisible_(sheet, row, hiddenRowsByGridMetadata)) {
        rowNumbers.add(row);
      }
    }
  });

  return Array.from(rowNumbers).sort((left, right) => left - right);
}

function isCalendarSheetRowVisible_(sheet, row, hiddenRowsByGridMetadata) {
  if (hiddenRowsByGridMetadata && hiddenRowsByGridMetadata.has(row)) {
    return false;
  }
  if (sheet.isRowHiddenByFilter(row)) {
    return false;
  }
  if (sheet.isRowHiddenByUser(row)) {
    return false;
  }
  return true;
}

function collectHiddenRowsByGridMetadata_(ss, sheet, firstRow, lastRow) {
  const hiddenRows = new Set();
  if (!ss || !sheet || firstRow > lastRow || typeof Sheets === 'undefined') {
    return hiddenRows;
  }

  try {
    const rangeA1 = `'${escapeSheetNameForFormula_(sheet.getName())}'!${firstRow}:${lastRow}`;
    const model = Sheets.Spreadsheets.get(ss.getId(), {
      ranges: [rangeA1],
      fields: 'sheets(properties(sheetId),data(startRow,rowMetadata(hiddenByFilter,hiddenByUser)))',
      includeGridData: true,
    });
    const sheetModel = (model.sheets || []).find(
      (entry) => entry.properties && entry.properties.sheetId === sheet.getSheetId()
    );
    const gridData = sheetModel && sheetModel.data && sheetModel.data[0];
    const rowMetadata = (gridData && gridData.rowMetadata) || [];
    const zeroBasedStartRow = Number(gridData && gridData.startRow) || firstRow - 1;

    rowMetadata.forEach((metadata, index) => {
      if (metadata && (metadata.hiddenByFilter || metadata.hiddenByUser)) {
        hiddenRows.add(zeroBasedStartRow + index + 1);
      }
    });
  } catch (error) {
    // Fall back to Apps Script row checks below if grid metadata is unavailable.
  }

  return hiddenRows;
}

function getSelectedCalendarRanges_(ss, sheet) {
  const rangeList = ss.getActiveRangeList ? ss.getActiveRangeList() : null;
  if (rangeList) {
    return rangeList.getRanges().filter((range) => range.getSheet().getSheetId() === sheet.getSheetId());
  }

  const range = ss.getActiveRange();
  if (range && range.getSheet().getSheetId() === sheet.getSheetId()) {
    return [range];
  }

  return [];
}

function getMarkActionLabel_(targetSheetName) {
  if (targetSheetName === CONFIG.invoicingSheetName) {
    return 'for Invoicing';
  }
  if (targetSheetName === CONFIG.nonBillableSheetName) {
    return 'as Non-Billable';
  }
  return `as ${targetSheetName}`;
}

function showMarkProgress_(ss, targetSheetName, done, total, stepLabel) {
  const normalizedDone = Math.min(Math.max(Number(done) || 0, 0), Math.max(Number(total) || 0, 0));
  const normalizedTotal = Math.max(Number(total) || 0, 0);
  const percentage = normalizedTotal > 0 ? Math.floor((normalizedDone / normalizedTotal) * 100) : 0;
  const stepText = stepLabel ? `${stepLabel} — ` : '';

  writeStatusCellMessage_(
    ss,
    `${stepText}Marking selected Calendar rows ${getMarkActionLabel_(targetSheetName)}: ${normalizedDone}/${normalizedTotal} (${percentage}%)`
  );
  SpreadsheetApp.flush();
}

function appendCalendarRowsToInvoicing_(calendarRows, invoicingSheet, progressCallback) {
  const invoiceStore = readInvoicingState_(invoicingSheet);
  const appendValues = [];
  calendarRows.forEach((row, index) => {
    if (!invoiceStore.byEventKey.has(row.eventKey)) {
      appendValues.push([
        row.eventKey,
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
    }

    reportMarkProgress_(progressCallback, index + 1, calendarRows.length);
  });

  appendInvoicingRows_(invoicingSheet, appendValues);
  return appendValues.length;
}

function appendCalendarRowsToNonBillable_(calendarRows, nonBillableSheet, progressCallback) {
  const nonBillableStore = readNonBillableState_(nonBillableSheet);
  const appendValues = [];
  calendarRows.forEach((row, index) => {
    if (!nonBillableStore.byEventKey.has(row.eventKey)) {
      appendValues.push([
        row.eventKey,
        row.values[0],
        row.values[1],
        row.values[2],
        row.values[3],
        row.values[4],
        row.values[5],
        '',
      ]);
    }

    reportMarkProgress_(progressCallback, index + 1, calendarRows.length);
  });

  appendNonBillableRows_(nonBillableSheet, appendValues);
  return appendValues.length;
}

function reportMarkProgress_(progressCallback, done, total) {
  if (!progressCallback) {
    return;
  }
  if (done === total || done % 10 === 0) {
    progressCallback(done, total);
  }
}

function removeRegisterRowsByEventKeys_(sheet, header, eventKeys, progressCallback) {
  const keys = new Set((eventKeys || []).filter((key) => key));
  if (keys.size === 0) {
    return 0;
  }

  const rowCount = Math.max(sheet.getLastRow() - 1, 0);
  if (rowCount === 0) {
    return 0;
  }

  const eventIdValues = sheet.getRange(2, 1, rowCount, 1).getValues();
  let removedCount = 0;

  for (let index = eventIdValues.length - 1; index >= 0; index -= 1) {
    const eventKey = toText_(eventIdValues[index][0]);
    if (keys.has(eventKey)) {
      sheet.getRange(index + 2, 1, 1, header.length).deleteCells(SpreadsheetApp.Dimension.ROWS);
      removedCount += 1;
    }

    reportMarkProgress_(progressCallback, eventIdValues.length - index, eventIdValues.length);
  }

  return removedCount;
}
