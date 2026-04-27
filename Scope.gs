function buildScope_() {
  const importStart = parseImportStartDate_(CONFIG.importStartDate);
  const now = new Date();

  return {
    importStart,
    now,
    importStartMillis: importStart.getTime(),
    nowMillis: now.getTime(),
  };
}

function parseImportStartDate_(value) {
  if (typeof value !== 'string') {
    throw new Error(
      `Invalid CONFIG.importStartDate: "${value}". Use ISO date format YYYY-MM-DD (example: 2024-01-01).`
    );
  }

  const trimmed = value.trim();
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(trimmed);
  if (!match) {
    throw new Error(
      `Invalid CONFIG.importStartDate: "${value}". Use ISO date format YYYY-MM-DD (example: 2024-01-01).`
    );
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  const isRealDate = date.getUTCFullYear() === year
    && date.getUTCMonth() === month - 1
    && date.getUTCDate() === day;
  if (!isRealDate) {
    throw new Error(
      `Invalid CONFIG.importStartDate: "${value}" is not a real calendar date. Use YYYY-MM-DD (example: 2024-01-01).`
    );
  }

  return date;
}

function isManagedEventInScope_(eventObj, scope) {
  return (
    eventObj &&
    eventObj.start instanceof Date &&
    eventObj.end instanceof Date &&
    !Number.isNaN(eventObj.start.getTime()) &&
    !Number.isNaN(eventObj.end.getTime()) &&
    eventObj.start.getTime() >= scope.importStartMillis &&
    eventObj.end.getTime() <= scope.nowMillis
  );
}

function isExistingRowInScope_(rowValues, scope) {
  const start = rowValues[3];
  const end = rowValues[4];

  if (!(start instanceof Date) || Number.isNaN(start.getTime())) {
    return false;
  }
  if (!(end instanceof Date) || Number.isNaN(end.getTime())) {
    return false;
  }

  return (
    start.getTime() >= scope.importStartMillis &&
    end.getTime() <= scope.nowMillis
  );
}

function isExistingRowBeforeImportStart_(rowValues, scope) {
  const start = rowValues[3];

  if (!(start instanceof Date) || Number.isNaN(start.getTime())) {
    return false;
  }

  return start.getTime() < scope.importStartMillis;
}

function isExistingRowAfterNow_(rowValues, scope) {
  const start = rowValues[3];

  if (!(start instanceof Date) || Number.isNaN(start.getTime())) {
    return false;
  }

  return start.getTime() > scope.nowMillis;
}
