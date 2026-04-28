function removeManagedDuplicates_(rows, scope) {
  const groups = new Map();

  rows.forEach((row, index) => {
    if (!isRowManagedInScopeForDuplicateCheck_(row, scope)) {
      return;
    }

    const key = buildDuplicateKey_(row.values);
    if (!groups.has(key)) {
      groups.set(key, []);
    }

    groups.get(key).push({
      index,
      row,
      calendar: toText_(row.values[0]),
    });
  });

  const removeIndexes = new Set();

  groups.forEach((entries) => {
    if (entries.length < 2) {
      return;
    }

    const byCalendar = new Map();

    entries.forEach((entry) => {
      if (!byCalendar.has(entry.calendar)) {
        byCalendar.set(entry.calendar, []);
      }
      byCalendar.get(entry.calendar).push(entry);
    });

    byCalendar.forEach((calendarEntries) => {
      if (calendarEntries.length > 1) {
        calendarEntries.forEach((entry) => removeIndexes.add(entry.index));
      }
    });

    const survivorsAfterSameCalendarPurge = entries.filter(
      (entry) => !removeIndexes.has(entry.index)
    );

    if (survivorsAfterSameCalendarPurge.length < 2) {
      return;
    }

    const specificEntries = survivorsAfterSameCalendarPurge.filter(
      (entry) => entry.calendar !== CONFIG.defaultCalendarName
    );

    let survivors = survivorsAfterSameCalendarPurge;

    if (specificEntries.length > 0) {
      survivorsAfterSameCalendarPurge
        .filter((entry) => entry.calendar === CONFIG.defaultCalendarName)
        .forEach((entry) => removeIndexes.add(entry.index));

      survivors = specificEntries;
    }

    if (survivors.length < 2) {
      return;
    }

    survivors.sort((a, b) => compareDuplicateCandidates_(a.row, b.row));
    survivors.slice(1).forEach((entry) => removeIndexes.add(entry.index));
  });

  return rows.filter((row, index) => !removeIndexes.has(index));
}

function isRowManagedInScopeForDuplicateCheck_(row, scope) {
  return !!(row && row.eventKey);
}

function buildDuplicateKey_(rowValues) {
  return JSON.stringify({
    event: normalizeDuplicateText_(rowValues[1]),
    date: duplicateDateToken_(rowValues[2]),
    start: duplicateDateToken_(rowValues[3]),
    end: duplicateDateToken_(rowValues[4]),
  });
}

function normalizeDuplicateText_(value) {
  return toText_(value).trim();
}

function duplicateDateToken_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.getTime();
  }

  return toText_(value);
}

function compareDuplicateCandidates_(a, b) {
  const aPriority = getCalendarPriority_(toText_(a.values[0]));
  const bPriority = getCalendarPriority_(toText_(b.values[0]));
  if (aPriority !== bPriority) {
    return aPriority - bPriority;
  }

  const aInvoice = a.invoiceNumber ? 0 : 1;
  const bInvoice = b.invoiceNumber ? 0 : 1;
  if (aInvoice !== bInvoice) {
    return aInvoice - bInvoice;
  }

  const aChanged = a.rowKind === CONFIG.rowKind.changedCopy ? 1 : 0;
  const bChanged = b.rowKind === CONFIG.rowKind.changedCopy ? 1 : 0;
  if (aChanged !== bChanged) {
    return aChanged - bChanged;
  }

  return 0;
}

function getCalendarPriority_(calendarName) {
  const nonDefault = CONFIG.calendarNames.filter(
    (name) => name !== CONFIG.defaultCalendarName
  );
  const order = nonDefault.concat([CONFIG.defaultCalendarName]);
  const index = order.indexOf(calendarName);
  return index >= 0 ? index : Number.MAX_SAFE_INTEGER;
}
