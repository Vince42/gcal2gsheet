function buildScope_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeZone = ss ? ss.getSpreadsheetTimeZone() : 'Etc/UTC';
  const importStart = parseImportStartDate_(CONFIG.importStartDate, timeZone);
  const now = new Date();

  return {
    importStart,
    now,
    importStartMillis: importStart.getTime(),
    nowMillis: now.getTime(),
  };
}

function parseImportStartDate_(value, timeZone) {
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
  const date = buildMidnightForTimeZone_(year, month, day, timeZone || 'Etc/UTC');
  const isRealDate = date instanceof Date && !Number.isNaN(date.getTime());
  if (!isRealDate) {
    throw new Error(
      `Invalid CONFIG.importStartDate: "${value}" is not a real calendar date. Use YYYY-MM-DD (example: 2024-01-01).`
    );
  }

  return date;
}

function buildMidnightForTimeZone_(year, month, day, timeZone) {
  let guessUtcMs = Date.UTC(year, month - 1, day, 0, 0, 0);
  for (let i = 0; i < 3; i += 1) {
    const offsetMinutes = parseTzOffsetMinutes_(Utilities.formatDate(new Date(guessUtcMs), timeZone, 'Z'));
    const nextGuessUtcMs = Date.UTC(year, month - 1, day, 0, 0, 0) - offsetMinutes * 60000;
    if (nextGuessUtcMs === guessUtcMs) {
      break;
    }
    guessUtcMs = nextGuessUtcMs;
  }
  return new Date(guessUtcMs);
}

function parseTzOffsetMinutes_(offset) {
  const text = String(offset || '').trim();
  const match = /^([+-])(\d{2})(\d{2})$/.exec(text);
  if (!match) {
    throw new Error(`Invalid timezone offset "${offset}".`);
  }
  const sign = match[1] === '-' ? -1 : 1;
  const hours = Number(match[2]);
  const minutes = Number(match[3]);
  return sign * (hours * 60 + minutes);
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
