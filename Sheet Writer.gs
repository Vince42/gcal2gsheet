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
    const values = rows.map((row, index) => {
      const rowValues = row.values.slice();
      if (CONFIG.header[6] === 'Status') {
        rowValues[6] = buildStatusFormula_(index + 2);
      }
      return rowValues;
    });
    sheet.getRange(2, 1, rows.length, CONFIG.header.length).setValues(values);
  }
}

function buildStatusFormula_(rowNumber) {
  return `=IF(COUNTIF('${CONFIG.invoicingStateSheetName}'!$A:$A,'${CONFIG.stateSheetName}'!$A${rowNumber})>0,"Invoiced",IF(COUNTIF('${CONFIG.nonBillableStateSheetName}'!$A:$A,'${CONFIG.stateSheetName}'!$A${rowNumber})>0,"Non-billable",IF('${CONFIG.stateSheetName}'!$B${rowNumber}="${CONFIG.rowKind.changedCopy}","Changed","Open")))`;
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

function applyNumberFormats_(sheet, header) {
  const effectiveHeader = header || CONFIG.header;
  const lastRow = sheet.getLastRow();
  const rowCount = Math.max(lastRow - 1, 0);

  if (rowCount === 0) {
    return;
  }

  applyColumnNumberFormatByName_(sheet, effectiveHeader, 'Date', rowCount, 'yyyy-mm-dd');
  applyColumnNumberFormatByName_(sheet, effectiveHeader, 'Start', rowCount, 'hh:mm');
  applyColumnNumberFormatByName_(sheet, effectiveHeader, 'End', rowCount, 'hh:mm');
  applyColumnNumberFormatByName_(sheet, effectiveHeader, 'Duration', rowCount, 'hh:mm');
  applyColumnNumberFormatByName_(sheet, effectiveHeader, 'InvoiceDate', rowCount, 'yyyy-mm-dd');
}

function applyColumnNumberFormatByName_(sheet, header, columnName, rowCount, numberFormat) {
  const index = header.indexOf(columnName);
  if (index < 0) {
    return;
  }

  sheet.getRange(2, index + 1, rowCount, 1).setNumberFormat(numberFormat);
}

function applyRowColors_(sheet, rows) {
  if (rows.length === 0) {
    return;
  }

  const colors = rows.map(() => new Array(CONFIG.header.length).fill(CONFIG.colors.normal));
  sheet.getRange(2, 1, rows.length, CONFIG.header.length).setFontColors(colors);
}

function clearRetiredCalendarInvoiceColumns_(sheet) {
  const firstRetiredColumn = CONFIG.header.length + 1;
  const retiredColumnCount = Math.max(CONFIG.invoicingHeader.length - CONFIG.header.length, 0);
  if (retiredColumnCount === 0) {
    return;
  }

  const rowCount = Math.max(sheet.getLastRow(), 1);
  sheet.getRange(1, firstRetiredColumn, rowCount, retiredColumnCount).clearContent();
}
