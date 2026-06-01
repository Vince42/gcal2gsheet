function readNonBillableState_(sheet, stateSheet) {
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
    ? sheet.getRange(2, 1, visibleRowCount, CONFIG.nonBillableHeader.length).getValues()
    : [];
  const stateValues = stateRowCount > 0
    ? stateSheet.getRange(2, 1, stateRowCount, CONFIG.nonBillableStateHeader.length).getValues()
    : [];

  for (let i = 0; i < rowCount; i += 1) {
    const rowValues = i < visibleValues.length
      ? visibleValues[i].slice()
      : new Array(CONFIG.nonBillableHeader.length).fill('');
    const stateRow = i < stateValues.length ? stateValues[i] : [''];
    const eventKey = toText_(stateRow[0]);

    if (!eventKey || isCompletelyBlankRow_(rowValues)) {
      continue;
    }

    byEventKey.set(eventKey, {
      eventKey,
      reason: toText_(rowValues[6]),
      values: rowValues,
    });
  }

  return { byEventKey };
}

function repairNonBillableStateFromImportedEvents_(currentByKey, nonBillableSheet, nonBillableStateSheet) {
  const rowCount = Math.max(nonBillableSheet.getLastRow() - 1, 0);
  if (rowCount === 0) {
    return 0;
  }

  const values = nonBillableSheet.getRange(2, 1, rowCount, CONFIG.nonBillableHeader.length).getValues();
  const stateValues = nonBillableStateSheet
    .getRange(2, 1, rowCount, CONFIG.nonBillableStateHeader.length)
    .getValues();
  const eventKeysByMatchKey = new Map();

  currentByKey.forEach((eventObj, eventKey) => {
    eventKeysByMatchKey.set(buildEventRecordMatchKey_(eventObj.values), eventKey);
  });

  let repairedCount = 0;
  values.forEach((rowValues, index) => {
    if (toText_(stateValues[index][0])) {
      return;
    }

    const eventKey = eventKeysByMatchKey.get(buildEventRecordMatchKey_(rowValues));
    if (!eventKey) {
      return;
    }

    stateValues[index][0] = eventKey;
    repairedCount += 1;
  });

  if (repairedCount > 0) {
    nonBillableStateSheet
      .getRange(2, 1, stateValues.length, CONFIG.nonBillableStateHeader.length)
      .setValues(stateValues);
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
      eventObj.values[6] = 'Invoiced';
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

function appendNonBillableRows_(nonBillableSheet, nonBillableStateSheet, values, stateValues) {
  if (!values || values.length === 0) {
    return;
  }

  const startRow = Math.max(nonBillableSheet.getLastRow() + 1, 2);
  const neededLastRow = startRow + values.length - 1;
  if (nonBillableSheet.getMaxRows() < neededLastRow) {
    nonBillableSheet.insertRowsAfter(nonBillableSheet.getMaxRows(), neededLastRow - nonBillableSheet.getMaxRows());
  }
  if (nonBillableStateSheet.getMaxRows() < neededLastRow) {
    nonBillableStateSheet.insertRowsAfter(
      nonBillableStateSheet.getMaxRows(),
      neededLastRow - nonBillableStateSheet.getMaxRows()
    );
  }

  nonBillableSheet.getRange(startRow, 1, values.length, CONFIG.nonBillableHeader.length).setValues(values);
  nonBillableStateSheet
    .getRange(startRow, 1, stateValues.length, CONFIG.nonBillableStateHeader.length)
    .setValues(stateValues);
}
