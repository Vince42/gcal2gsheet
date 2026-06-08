function readNonBillableState_(sheet) {
  const rowCount = Math.max(sheet.getLastRow() - 1, 0);
  const byEventKey = new Map();

  if (rowCount === 0) {
    return { byEventKey };
  }

  const values = sheet.getRange(2, 1, rowCount, CONFIG.nonBillableHeader.length).getValues();

  values.forEach((rowValues) => {
    const eventKey = toText_(rowValues[0]);
    const nonBillableValues = rowValues.slice(1);

    if (!eventKey || isCompletelyBlankRow_(nonBillableValues)) {
      return;
    }

    byEventKey.set(eventKey, {
      eventKey,
      reason: toText_(rowValues[7]),
      values: rowValues,
    });
  });

  return { byEventKey };
}

function repairNonBillableStateFromImportedEvents_(currentByKey, nonBillableSheet) {
  const rowCount = Math.max(nonBillableSheet.getLastRow() - 1, 0);
  if (rowCount === 0) {
    return 0;
  }

  const values = nonBillableSheet.getRange(2, 1, rowCount, CONFIG.nonBillableHeader.length).getValues();
  const eventKeysByMatchKey = new Map();

  currentByKey.forEach((eventObj, eventKey) => {
    eventKeysByMatchKey.set(buildEventRecordMatchKey_(eventObj.values), eventKey);
  });

  let repairedCount = 0;
  values.forEach((rowValues, index) => {
    if (toText_(rowValues[0])) {
      return;
    }

    const eventKey = eventKeysByMatchKey.get(buildEventRecordMatchKey_(rowValues.slice(1)));
    if (!eventKey) {
      return;
    }

    values[index][0] = eventKey;
    repairedCount += 1;
  });

  if (repairedCount > 0) {
    nonBillableSheet.getRange(2, 1, values.length, CONFIG.nonBillableHeader.length).setValues(values);
  }

  return repairedCount;
}

function applyRegisterStatusesToImportedEvents_(currentByKey, invoiceStore, nonBillableStore) {
  const invoicesByEventKey = invoiceStore && invoiceStore.byEventKey
    ? invoiceStore.byEventKey
    : new Map();
  const nonBillableByEventKey = nonBillableStore && nonBillableStore.byEventKey
    ? nonBillableStore.byEventKey
    : new Map();

  currentByKey.forEach((eventObj, eventKey) => {
    const invoice = invoicesByEventKey.get(eventKey);
    if (invoice) {
      eventObj.invoiceNumber = invoice.invoiceNumber || 'INVOICED';
      eventObj.values[6] = 'Invoicing';
      return;
    }

    const nonBillable = nonBillableByEventKey.get(eventKey);
    if (nonBillable) {
      eventObj.invoiceNumber = 'NON_BILLABLE';
      eventObj.values[6] = 'Non-billable';
      return;
    }

    eventObj.invoiceNumber = '';
    eventObj.values[6] = 'Open';
  });
}

function buildEventRecordMatchKey_(rowValues) {
  return JSON.stringify({
    calendar: toText_(rowValues[0]).trim(),
    event: normalizeDuplicateText_(rowValues[1]),
    date: duplicateDateToken_(rowValues[2]),
    start: duplicateDateToken_(rowValues[3]),
    end: duplicateDateToken_(rowValues[4]),
  });
}

function appendNonBillableRows_(nonBillableSheet, values) {
  if (!values || values.length === 0) {
    return;
  }

  const startRow = Math.max(nonBillableSheet.getLastRow() + 1, 2);
  const neededLastRow = startRow + values.length - 1;
  if (nonBillableSheet.getMaxRows() < neededLastRow) {
    nonBillableSheet.insertRowsAfter(nonBillableSheet.getMaxRows(), neededLastRow - nonBillableSheet.getMaxRows());
  }

  nonBillableSheet.getRange(startRow, 1, values.length, CONFIG.nonBillableHeader.length).setValues(values);
}
