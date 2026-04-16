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
  if (!value || !/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    throw new Error(`Invalid CONFIG.importStartDate: ${value}`);
  }

  return new Date(`${value}T00:00:00`);
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
