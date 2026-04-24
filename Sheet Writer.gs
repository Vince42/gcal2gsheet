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
