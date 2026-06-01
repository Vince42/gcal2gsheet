function readInvoicingState_(sheet, stateSheet) {
  const visibleLastRow = sheet.getLastRow();
  const stateLastRow = stateSheet.getLastRow();
  const visibleRowCount = Math.max(visibleLastRow - 1, 0);
  const stateRowCount = Math.max(stateLastRow - 1, 0);
  const rowCount = Math.max(visibleRowCount, stateRowCount);
  const byEventKey = new Map();

  if (rowCount === 0) {
    return { byEventKey };
  }

  const visibleValues = visibleRowCount > 0
    ? sheet.getRange(2, 1, visibleRowCount, CONFIG.invoicingHeader.length).getValues()
    : [];
  const stateValues = stateRowCount > 0
    ? stateSheet.getRange(2, 1, stateRowCount, CONFIG.invoicingStateHeader.length).getValues()
    : [];

  for (let i = 0; i < rowCount; i += 1) {
    const rowValues = i < visibleValues.length
      ? visibleValues[i].slice()
      : new Array(CONFIG.invoicingHeader.length).fill('');
    const stateRow = i < stateValues.length ? stateValues[i] : [''];
    const eventKey = toText_(stateRow[0]);

    if (!eventKey || isCompletelyBlankRow_(rowValues)) {
      continue;
    }

    byEventKey.set(eventKey, {
      eventKey,
      invoiceNumber: toText_(rowValues[8]),
      values: rowValues,
    });
  }

  return { byEventKey };
}

function migrateCalendarInvoicesToInvoicing_(calendarSheet, calendarStateSheet, invoicingSheet, invoicingStateSheet) {
  const invoiceStore = readInvoicingState_(invoicingSheet, invoicingStateSheet);
  const existingKeys = new Set(invoiceStore.byEventKey.keys());
  const legacyHeaderLength = CONFIG.invoicingHeader.length;
  const visibleRowCount = Math.max(calendarSheet.getLastRow() - 1, 0);
  const stateRowCount = Math.max(calendarStateSheet.getLastRow() - 1, 0);
  const rowCount = Math.min(visibleRowCount, stateRowCount);

  if (rowCount === 0) {
    return 0;
  }

  const visibleValues = calendarSheet.getRange(2, 1, rowCount, legacyHeaderLength).getValues();
  const stateValues = calendarStateSheet.getRange(2, 1, rowCount, CONFIG.stateHeader.length).getValues();
  const appendValues = [];
  const appendStateValues = [];

  for (let i = 0; i < rowCount; i += 1) {
    const rowValues = visibleValues[i].slice();
    const eventKey = toText_(stateValues[i][0]);
    const invoiceNumber = toText_(rowValues[8]).trim();

    if (!eventKey || !invoiceNumber || existingKeys.has(eventKey)) {
      continue;
    }

    appendValues.push(normalizeInvoicingValues_(rowValues));
    appendStateValues.push([eventKey]);
    existingKeys.add(eventKey);
  }

  appendInvoicingRows_(invoicingSheet, invoicingStateSheet, appendValues, appendStateValues);
  return appendValues.length;
}

function appendInvoicingRows_(invoicingSheet, invoicingStateSheet, values, stateValues) {
  if (!values || values.length === 0) {
    return;
  }

  const startRow = Math.max(invoicingSheet.getLastRow() + 1, 2);
  const neededLastRow = startRow + values.length - 1;
  if (invoicingSheet.getMaxRows() < neededLastRow) {
    invoicingSheet.insertRowsAfter(invoicingSheet.getMaxRows(), neededLastRow - invoicingSheet.getMaxRows());
  }
  if (invoicingStateSheet.getMaxRows() < neededLastRow) {
    invoicingStateSheet.insertRowsAfter(
      invoicingStateSheet.getMaxRows(),
      neededLastRow - invoicingStateSheet.getMaxRows()
    );
  }

  invoicingSheet.getRange(startRow, 1, values.length, CONFIG.invoicingHeader.length).setValues(values);
  invoicingStateSheet.getRange(startRow, 1, stateValues.length, CONFIG.invoicingStateHeader.length).setValues(stateValues);
}

function normalizeInvoicingValues_(rowValues) {
  const values = rowValues.slice(0, CONFIG.invoicingHeader.length);
  while (values.length < CONFIG.invoicingHeader.length) {
    values.push('');
  }
  return values;
}

function repairInvoicingStateFromCalendarRows_(calendarSheet, calendarStateSheet, invoicingSheet, invoicingStateSheet) {
  const calendarRowCount = Math.min(
    Math.max(calendarSheet.getLastRow() - 1, 0),
    Math.max(calendarStateSheet.getLastRow() - 1, 0)
  );
  const invoicingRowCount = Math.max(invoicingSheet.getLastRow() - 1, 0);

  if (calendarRowCount === 0 || invoicingRowCount === 0) {
    return 0;
  }

  const calendarValues = calendarSheet.getRange(2, 1, calendarRowCount, CONFIG.header.length).getValues();
  const calendarStateValues = calendarStateSheet.getRange(2, 1, calendarRowCount, CONFIG.stateHeader.length).getValues();
  const invoicingValues = invoicingSheet.getRange(2, 1, invoicingRowCount, CONFIG.invoicingHeader.length).getValues();
  const invoicingStateValues = invoicingStateSheet.getRange(
    2,
    1,
    invoicingRowCount,
    CONFIG.invoicingStateHeader.length
  ).getValues();
  const eventKeysByInvoiceMatchKey = new Map();

  calendarValues.forEach((rowValues, index) => {
    const eventKey = toText_(calendarStateValues[index][0]);
    if (eventKey) {
      eventKeysByInvoiceMatchKey.set(buildInvoiceMatchKey_(rowValues), eventKey);
    }
  });

  let repairedCount = 0;
  invoicingValues.forEach((rowValues, index) => {
    if (toText_(invoicingStateValues[index][0])) {
      return;
    }

    const eventKey = eventKeysByInvoiceMatchKey.get(buildInvoiceMatchKey_(rowValues));
    if (!eventKey) {
      return;
    }

    invoicingStateValues[index][0] = eventKey;
    repairedCount += 1;
  });

  if (repairedCount > 0) {
    invoicingStateSheet
      .getRange(2, 1, invoicingStateValues.length, CONFIG.invoicingStateHeader.length)
      .setValues(invoicingStateValues);
  }

  return repairedCount;
}

function buildInvoiceMatchKey_(rowValues) {
  return JSON.stringify({
    calendar: toText_(rowValues[0]).trim(),
    event: normalizeDuplicateText_(rowValues[1]),
    date: duplicateDateToken_(rowValues[2]),
    start: duplicateDateToken_(rowValues[3]),
    end: duplicateDateToken_(rowValues[4]),
  });
}

function repairInvoicingStateFromImportedEvents_(currentByKey, invoicingSheet, invoicingStateSheet) {
  const invoicingRowCount = Math.max(invoicingSheet.getLastRow() - 1, 0);
  if (invoicingRowCount === 0) {
    return 0;
  }

  const invoicingValues = invoicingSheet.getRange(2, 1, invoicingRowCount, CONFIG.invoicingHeader.length).getValues();
  const invoicingStateValues = invoicingStateSheet.getRange(
    2,
    1,
    invoicingRowCount,
    CONFIG.invoicingStateHeader.length
  ).getValues();
  const eventKeysByInvoiceMatchKey = new Map();

  currentByKey.forEach((eventObj, eventKey) => {
    eventKeysByInvoiceMatchKey.set(buildInvoiceMatchKey_(eventObj.values), eventKey);
  });

  let repairedCount = 0;
  invoicingValues.forEach((rowValues, index) => {
    if (toText_(invoicingStateValues[index][0])) {
      return;
    }

    const eventKey = eventKeysByInvoiceMatchKey.get(buildInvoiceMatchKey_(rowValues));
    if (!eventKey) {
      return;
    }

    invoicingStateValues[index][0] = eventKey;
    repairedCount += 1;
  });

  if (repairedCount > 0) {
    invoicingStateSheet
      .getRange(2, 1, invoicingStateValues.length, CONFIG.invoicingStateHeader.length)
      .setValues(invoicingStateValues);
  }

  return repairedCount;
}
