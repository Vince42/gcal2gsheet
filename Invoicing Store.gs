function readInvoicingState_(sheet) {
  const rowCount = Math.max(sheet.getLastRow() - 1, 0);
  const byEventKey = new Map();

  if (rowCount === 0) {
    return { byEventKey };
  }

  const values = sheet.getRange(2, 1, rowCount, CONFIG.invoicingHeader.length).getValues();

  values.forEach((rowValues) => {
    const eventKey = toText_(rowValues[0]);
    const invoiceValues = rowValues.slice(1);

    if (!eventKey || isCompletelyBlankRow_(invoiceValues)) {
      return;
    }

    byEventKey.set(eventKey, {
      eventKey,
      invoiceNumber: toText_(rowValues[9]),
      values: rowValues,
    });
  });

  return { byEventKey };
}

function migrateCalendarInvoicesToInvoicing_(calendarSheet, invoicingSheet) {
  const invoiceStore = readInvoicingState_(invoicingSheet);
  const existingKeys = new Set(invoiceStore.byEventKey.keys());
  const rowCount = Math.max(calendarSheet.getLastRow() - 1, 0);

  if (rowCount === 0) {
    return 0;
  }

  const readWidth = Math.max(CONFIG.invoicingHeader.length + 1, LEGACY_INVOICING_HEADER.length + 2);
  const calendarValues = calendarSheet.getRange(2, 1, rowCount, readWidth).getValues();
  const appendValues = [];

  calendarValues.forEach((rowValues) => {
    const eventKey = toText_(rowValues[0]);
    const invoicingValues = normalizeCalendarInvoiceValues_(rowValues);

    if (!eventKey || !hasInvoiceRegisterData_(invoicingValues) || existingKeys.has(eventKey)) {
      return;
    }

    appendValues.push(invoicingValues);
    existingKeys.add(eventKey);
  });

  appendInvoicingRows_(invoicingSheet, appendValues);
  return appendValues.length;
}

function appendInvoicingRows_(invoicingSheet, values) {
  if (!values || values.length === 0) {
    return;
  }

  const startRow = Math.max(invoicingSheet.getLastRow() + 1, 2);
  const neededLastRow = startRow + values.length - 1;
  if (invoicingSheet.getMaxRows() < neededLastRow) {
    invoicingSheet.insertRowsAfter(invoicingSheet.getMaxRows(), neededLastRow - invoicingSheet.getMaxRows());
  }

  invoicingSheet.getRange(startRow, 1, values.length, CONFIG.invoicingHeader.length).setValues(values);
}

function normalizeInvoicingValues_(rowValues) {
  const values = rowValues.slice(0, CONFIG.invoicingHeader.length);
  while (values.length < CONFIG.invoicingHeader.length) {
    values.push('');
  }
  return values;
}

function normalizeCalendarInvoiceValues_(rowValues) {
  if (isLegacyCalendarStatusValue_(rowValues[7])) {
    return normalizeInvoicingValues_([
      rowValues[0],
      rowValues[1],
      rowValues[2],
      rowValues[3],
      rowValues[4],
      rowValues[5],
      rowValues[6],
      rowValues[8],
      rowValues[9],
      rowValues[10],
      rowValues[11],
    ]);
  }

  return normalizeInvoicingValues_(rowValues);
}

function isLegacyCalendarStatusValue_(value) {
  const text = toText_(value).trim();
  return text === 'Open' || text === 'Invoiced' || text === 'Non-billable';
}

function hasInvoiceRegisterData_(rowValues) {
  return [7, 8, 9, 10].some((index) => toText_(rowValues[index]).trim() !== '');
}

function repairInvoicingStateFromCalendarRows_(calendarSheet, invoicingSheet) {
  const calendarRowCount = Math.max(calendarSheet.getLastRow() - 1, 0);
  const invoicingRowCount = Math.max(invoicingSheet.getLastRow() - 1, 0);

  if (calendarRowCount === 0 || invoicingRowCount === 0) {
    return 0;
  }

  const calendarValues = calendarSheet.getRange(2, 1, calendarRowCount, CONFIG.header.length).getValues();
  const invoicingValues = invoicingSheet.getRange(2, 1, invoicingRowCount, CONFIG.invoicingHeader.length).getValues();
  const eventKeysByInvoiceMatchKey = new Map();

  calendarValues.forEach((rowValues) => {
    const eventKey = toText_(rowValues[0]);
    if (eventKey) {
      eventKeysByInvoiceMatchKey.set(buildInvoiceMatchKey_(rowValues.slice(1)), eventKey);
    }
  });

  let repairedCount = 0;
  invoicingValues.forEach((rowValues, index) => {
    if (toText_(rowValues[0])) {
      return;
    }

    const eventKey = eventKeysByInvoiceMatchKey.get(buildInvoiceMatchKey_(rowValues.slice(1)));
    if (!eventKey) {
      return;
    }

    invoicingValues[index][0] = eventKey;
    repairedCount += 1;
  });

  if (repairedCount > 0) {
    invoicingSheet.getRange(2, 1, invoicingValues.length, CONFIG.invoicingHeader.length).setValues(invoicingValues);
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

function repairInvoicingStateFromImportedEvents_(currentByKey, invoicingSheet) {
  const invoicingRowCount = Math.max(invoicingSheet.getLastRow() - 1, 0);
  if (invoicingRowCount === 0) {
    return 0;
  }

  const invoicingValues = invoicingSheet.getRange(2, 1, invoicingRowCount, CONFIG.invoicingHeader.length).getValues();
  const eventKeysByInvoiceMatchKey = new Map();

  currentByKey.forEach((eventObj, eventKey) => {
    eventKeysByInvoiceMatchKey.set(buildInvoiceMatchKey_(eventObj.values), eventKey);
  });

  let repairedCount = 0;
  invoicingValues.forEach((rowValues, index) => {
    if (toText_(rowValues[0])) {
      return;
    }

    const eventKey = eventKeysByInvoiceMatchKey.get(buildInvoiceMatchKey_(rowValues.slice(1)));
    if (!eventKey) {
      return;
    }

    invoicingValues[index][0] = eventKey;
    repairedCount += 1;
  });

  if (repairedCount > 0) {
    invoicingSheet.getRange(2, 1, invoicingValues.length, CONFIG.invoicingHeader.length).setValues(invoicingValues);
  }

  return repairedCount;
}
