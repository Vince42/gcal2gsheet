function buildImportedSignature_(eventObj, timeZone) {
  return JSON.stringify({
    calendar: eventObj.calendar,
    title: eventObj.title,
    date: formatDateCell_(eventObj.date, timeZone, 'yyyy-MM-dd'),
    start: formatDateCell_(eventObj.start, timeZone, "yyyy-MM-dd'T'HH:mm"),
    end: formatDateCell_(eventObj.end, timeZone, "yyyy-MM-dd'T'HH:mm"),
    duration: normalizeDuration_(eventObj.duration),
  });
}

function buildSheetRowSignature_(rowValues, timeZone) {
  return JSON.stringify({
    calendar: toText_(rowValues[0]),
    title: toText_(rowValues[1]),
    date: formatDateCell_(rowValues[2], timeZone, 'yyyy-MM-dd'),
    start: formatDateCell_(rowValues[3], timeZone, "yyyy-MM-dd'T'HH:mm"),
    end: formatDateCell_(rowValues[4], timeZone, "yyyy-MM-dd'T'HH:mm"),
    duration: normalizeDuration_(rowValues[5]),
  });
}

function normalizeDuration_(value) {
  if (typeof value === 'number') {
    return value.toFixed(10);
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return ((value.getTime() % 86400000) / 86400000).toFixed(10);
  }

  return toText_(value);
}

function toSheetDateOnly_(dateValue, timeZone) {
  const iso = Utilities.formatDate(dateValue, timeZone, "yyyy-MM-dd'T'12:00:00XXX");
  return new Date(iso);
}

function formatDateCell_(value, timeZone, pattern) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, timeZone, pattern);
  }

  return toText_(value);
}

function toText_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value);
}

function isCompletelyBlankRow_(rowValues) {
  return rowValues.every((value) => toText_(value) === '');
}

function cloneRowModel_(row) {
  return {
    eventKey: row.eventKey || '',
    syntheticKey: row.syntheticKey || '',
    rowKind: row.rowKind || CONFIG.rowKind.normal,
    invoiceNumber: row.invoiceNumber || '',
    signature: row.signature || '',
    values: row.values.slice(),
  };
}
