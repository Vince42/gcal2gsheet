function validateCalendarRowsForWrite_(rows) {
  const failures = [];
  (rows || []).forEach((row, index) => {
    const sheetRow = index + 2;
    const rowValues = row && row.values ? row.values : [];
    const missing = [];

    if (!toText_(row && row.eventKey).trim()) {
      missing.push('ID');
    }
    collectMissingCalendarRequiredFields_(rowValues).forEach((fieldName) => missing.push(fieldName));

    if (missing.length > 0) {
      failures.push(`row ${sheetRow}: ${missing.join(', ')}`);
    }
  });

  if (failures.length > 0) {
    throw new Error(
      `Calendar import validation failed before writing. Required cells must not be empty or invalid: ${failures.slice(0, 20).join('; ')}${failures.length > 20 ? `; ... and ${failures.length - 20} more` : ''}. The sheet was left unchanged so the incomplete retrieval cannot overwrite existing data.`
    );
  }
}

function collectMissingCalendarRequiredFields_(rowValues) {
  const missing = [];
  const fieldChecks = [
    { name: 'Calendar', index: 0, valid: (value) => toText_(value).trim() !== '' },
    { name: 'Event', index: 1, valid: (value) => toText_(value).trim() !== '' },
    { name: 'Date', index: 2, valid: isValidDateCellValue_ },
    { name: 'Start', index: 3, valid: isValidDateCellValue_ },
    { name: 'End', index: 4, valid: isValidDateCellValue_ },
    { name: 'Duration', index: 5, valid: isValidDurationCellValue_ },
  ];

  fieldChecks.forEach((check) => {
    if (!check.valid(rowValues[check.index])) {
      missing.push(check.name);
    }
  });
  return missing;
}

function isValidDateCellValue_(value) {
  return Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime());
}

function isValidDurationCellValue_(value) {
  return typeof value === 'number' && Number.isFinite(value) && value >= 0;
}

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
    const values = rows.map((row) => [row.eventKey || ''].concat(row.values.slice()));
    sheet.getRange(2, 1, rows.length, CONFIG.header.length).setValues(values);
  }

  writeStateFormulas_(sheet, Math.max(rows.length, clearRows));
}

function writeStateFormulas_(sheet, rowCount) {
  const stateIndex = CONFIG.header.indexOf('State');
  if (stateIndex < 0 || rowCount <= 0) {
    return;
  }

  const formulas = [];
  for (let index = 0; index < rowCount; index += 1) {
    formulas.push([buildStateFormula_(index + 2)]);
  }
  sheet.getRange(2, stateIndex + 1, rowCount, 1).setFormulas(formulas);
}

function buildStateFormula_(rowNumber) {
  const invoicingSheetName = escapeSheetNameForFormula_(CONFIG.invoicingSheetName);
  const nonBillableSheetName = escapeSheetNameForFormula_(CONFIG.nonBillableSheetName);
  return `=IF(COUNTIF('${invoicingSheetName}'!$A:$A,$A${rowNumber})>0,"Invoicing",IF(COUNTIF('${nonBillableSheetName}'!$A:$A,$A${rowNumber})>0,"Non-billable","Open"))`;
}

function escapeSheetNameForFormula_(sheetName) {
  return String(sheetName || '').replace(/'/g, "''");
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
  const retiredColumnCount = getRetiredCalendarInvoiceColumnCount_();
  if (retiredColumnCount === 0) {
    return;
  }

  const rowCount = Math.max(sheet.getLastRow(), 1);
  sheet.getRange(1, firstRetiredColumn, rowCount, retiredColumnCount).clearContent();
}

function getRetiredCalendarInvoiceColumnCount_() {
  const legacyInvoiceTailColumnCount = Math.max(
    LEGACY_INVOICING_HEADER.length - (LEGACY_CALENDAR_HEADER.length - 1),
    0
  );
  const currentHeaderDelta = Math.max(CONFIG.invoicingHeader.length - CONFIG.header.length, 0);
  return Math.max(legacyInvoiceTailColumnCount, currentHeaderDelta);
}
