const LEGACY_CALENDAR_STATE_SHEET_NAME = '_calendar_state';
const LEGACY_INVOICING_STATE_SHEET_NAME = '_invoicing_state';
const LEGACY_NON_BILLABLE_STATE_SHEET_NAME = '_non_billable_state';

function ensureCalendarSheet_(ss) {
  return ensureNamedSheet_(ss, CONFIG.sheetName, false);
}

function ensureInvoicingSheet_(ss) {
  return ensureNamedSheet_(ss, CONFIG.invoicingSheetName, false);
}

function ensureNonBillableSheet_(ss) {
  return ensureNamedSheet_(ss, CONFIG.nonBillableSheetName, false);
}

function ensureNamedSheet_(ss, sheetName, hidden) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (hidden && !sheet.isSheetHidden()) {
    sheet.hideSheet();
  }

  return sheet;
}

function ensureHeader_(sheet, header, options) {
  const effectiveHeader = header || CONFIG.header;
  const headerRange = sheet.getRange(1, 1, 1, effectiveHeader.length);
  const current = headerRange.getValues()[0];

  const needsWrite = effectiveHeader.some((value, index) => current[index] !== value);
  if (!needsWrite) {
    return;
  }

  const allowOverwrite = !options || options.allowOverwrite !== false;
  if (!allowOverwrite && !isSheetBlankForManagedHeader_(sheet, current)) {
    throw new Error(
      `Sheet "${sheet.getName()}" already exists and does not have the expected managed header. Rename that sheet or configure a different managed sheet name before running the import.`
    );
  }

  headerRange.setValues([effectiveHeader]);
}

function isSheetBlankForManagedHeader_(sheet, currentHeader) {
  if (sheet.getLastRow() === 0) {
    return true;
  }

  if (sheet.getLastRow() > 1) {
    return false;
  }

  return currentHeader.every((value) => toText_(value) === '');
}

function ensureManagedWorkbookStructure_(ss, spreadsheetId) {
  const sheet = ensureCalendarSheet_(ss);
  const invoicingSheet = ensureInvoicingSheet_(ss);
  const nonBillableSheet = ensureNonBillableSheet_(ss);
  const legacyStateSheetNames = getLegacyStateSheetNameCandidates_();
  const legacyStateSheet = resolveLegacyStateSheet_(ss, legacyStateSheetNames.calendar);
  const legacyInvoicingStateSheet = resolveLegacyStateSheet_(ss, legacyStateSheetNames.invoicing);
  const legacyNonBillableStateSheet = resolveLegacyStateSheet_(ss, legacyStateSheetNames.nonBillable);

  migrateSheetToInlineIds_(sheet, CONFIG.header, LEGACY_CALENDAR_HEADER, legacyStateSheet);
  migrateSheetToInlineIds_(invoicingSheet, CONFIG.invoicingHeader, LEGACY_INVOICING_HEADER, legacyInvoicingStateSheet);
  migrateSheetToInlineIds_(
    nonBillableSheet,
    CONFIG.nonBillableHeader,
    LEGACY_NON_BILLABLE_HEADER,
    legacyNonBillableStateSheet
  );

  ensureHeader_(sheet);
  ensureHeader_(invoicingSheet, CONFIG.invoicingHeader, { allowOverwrite: false });
  ensureHeader_(nonBillableSheet, CONFIG.nonBillableHeader, { allowOverwrite: false });

  assertSheetHasExpectedColumns_(sheet, CONFIG.header);
  assertSheetHasExpectedColumns_(invoicingSheet, CONFIG.invoicingHeader);
  assertSheetHasExpectedColumns_(nonBillableSheet, CONFIG.nonBillableHeader);

  ensureSheetFormatting_(sheet);
  ensureSheetFormatting_(invoicingSheet);
  ensureSheetFormatting_(nonBillableSheet);
  ensureTable_(spreadsheetId, sheet);
  ensureCalendarStartDateFilter_(sheet);
  ensureTable_(spreadsheetId, invoicingSheet, CONFIG.invoicingTableName, CONFIG.invoicingHeader);
  ensureTable_(spreadsheetId, nonBillableSheet, CONFIG.nonBillableTableName, CONFIG.nonBillableHeader);
  deleteLegacyStateSheets_(
    ss,
    collectLegacyStateSheets_(
      ss,
      [].concat(
        legacyStateSheetNames.calendar,
        legacyStateSheetNames.invoicing,
        legacyStateSheetNames.nonBillable
      )
    )
  );

  return {
    sheet,
    invoicingSheet,
    nonBillableSheet,
  };
}



function resolveLegacyStateSheet_(ss, sheetNames) {
  const sheets = collectLegacyStateSheets_(ss, sheetNames || []);
  return sheets.length > 0 ? sheets[0] : null;
}

function collectLegacyStateSheets_(ss, sheetNames) {
  const seenSheetIds = new Set();
  const sheets = [];
  (sheetNames || []).forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || seenSheetIds.has(sheet.getSheetId())) {
      return;
    }
    seenSheetIds.add(sheet.getSheetId());
    sheets.push(sheet);
  });
  return sheets;
}

function migrateSheetToInlineIds_(sheet, targetHeader, legacyHeader, legacyStateSheet) {
  const currentFirstHeader = toText_(sheet.getRange(1, 1).getValue());
  if (currentFirstHeader === targetHeader[0]) {
    return;
  }
  if (currentFirstHeader && currentFirstHeader !== legacyHeader[0]) {
    return;
  }

  const rowCount = Math.max(sheet.getLastRow() - 1, 0);
  sheet.getRange(1, 1, 1, targetHeader.length).setValues([targetHeader]);
  if (rowCount === 0) {
    return;
  }

  const readWidth = targetHeader[0] === 'ID'
    ? Math.max(legacyHeader.length, LEGACY_INVOICING_HEADER.length + 1)
    : legacyHeader.length;
  const legacyValues = sheet.getRange(2, 1, rowCount, readWidth).getValues();
  const legacyStateRowCount = legacyStateSheet ? Math.max(legacyStateSheet.getLastRow() - 1, 0) : 0;
  const legacyStateValues = legacyStateRowCount > 0
    ? legacyStateSheet.getRange(2, 1, legacyStateRowCount, 1).getValues()
    : [];
  const migratedValues = legacyValues.map((row, index) => {
    const legacyId = index < legacyStateValues.length ? toText_(legacyStateValues[index][0]) : '';
    return [legacyId].concat(row);
  });

  const writeWidth = Math.max(targetHeader.length, migratedValues[0].length);
  sheet.getRange(2, 1, migratedValues.length, writeWidth).setValues(migratedValues);
}

function deleteLegacyStateSheets_(ss, stateSheets) {
  stateSheets.forEach((sheet) => {
    if (!sheet || ss.getSheets().length <= 1) {
      return;
    }
    ss.deleteSheet(sheet);
  });
}

function assertSheetHasExpectedColumns_(sheet, expectedHeader) {
  const current = sheet.getRange(1, 1, 1, expectedHeader.length).getValues()[0];
  const mismatchIndex = expectedHeader.findIndex((value, index) => current[index] !== value);
  if (mismatchIndex >= 0) {
    throw new Error(
      `Sheet "${sheet.getName()}" has invalid column ${mismatchIndex + 1}. Expected "${expectedHeader[mismatchIndex]}", got "${current[mismatchIndex]}".`
    );
  }
}

function ensureSheetFormatting_(sheet) {
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(true);
  ensureManagedIdColumnHidden_(sheet);
}

function ensureManagedIdColumnHidden_(sheet) {
  sheet.hideColumns(1);
}


function isCalendarTableMaintenanceRequest_(tableName, header) {
  return !tableName && !header;
}

function captureSheetFilterCriteria_(sheet, columnCount) {
  const filter = sheet.getFilter();
  if (!filter) {
    return { hadFilter: false, criteriaByColumn: [] };
  }

  const criteriaByColumn = [];
  for (let column = 1; column <= columnCount; column += 1) {
    criteriaByColumn[column] = filter.getColumnFilterCriteria(column);
  }
  return { hadFilter: true, criteriaByColumn };
}

function removeSheetFilterForTableUpdate_(sheet) {
  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }
}

function restoreSheetFilterCriteria_(sheet, filterSnapshot, rowCount, columnCount) {
  if (!filterSnapshot || !filterSnapshot.hadFilter) {
    return null;
  }

  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }

  const filter = sheet.getRange(1, 1, rowCount, columnCount).createFilter();
  (filterSnapshot.criteriaByColumn || []).forEach((criteria, column) => {
    if (criteria && column > 0) {
      filter.setColumnFilterCriteria(column, criteria);
    }
  });
  return filter;
}

function updateTableWithFilterSafeBatch_(spreadsheetId, sheet, header, requests) {
  const columnCount = header.length;
  const rowCount = Math.max(sheet.getLastRow(), 1);
  const filterSnapshot = captureSheetFilterCriteria_(sheet, columnCount);

  removeSheetFilterForTableUpdate_(sheet);
  try {
    Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
  } finally {
    restoreSheetFilterCriteria_(sheet, filterSnapshot, rowCount, columnCount);
  }
}

function assertManagedTableHasInlineIdColumn_(header) {
  const firstColumn = header && header.length > 0 ? header[0] : '';
  if (firstColumn !== 'ID' && firstColumn !== 'EventID') {
    throw new Error(`Managed table header must start with hidden ID/EventID column. Got "${firstColumn}".`);
  }
}

function ensureTable_(spreadsheetId, sheet, tableName, header) {
  const effectiveTableName = tableName || CONFIG.tableName;
  const effectiveHeader = header || CONFIG.header;
  assertManagedTableHasInlineIdColumn_(effectiveHeader);
  ensureManagedIdColumnHidden_(sheet);
  const spreadsheetModel = getSpreadsheetModel_(spreadsheetId);
  const sheetModel = (spreadsheetModel.sheets || []).find(
    (entry) => entry.properties && entry.properties.sheetId === sheet.getSheetId()
  );
  const tables = (sheetModel && sheetModel.tables) || [];

  if (tables.length === 0) {
    updateTableWithFilterSafeBatch_(spreadsheetId, sheet, effectiveHeader, [
      {
        addTable: {
          table: {
            name: effectiveTableName,
            range: {
              sheetId: sheet.getSheetId(),
              startRowIndex: 0,
              endRowIndex: Math.max(sheet.getLastRow(), 1),
              startColumnIndex: 0,
              endColumnIndex: effectiveHeader.length,
            },
          },
        },
      },
    ]);
    return;
  }

  const table = tables.find((entry) => entry.name === effectiveTableName) || tables[0];

  updateTableWithFilterSafeBatch_(spreadsheetId, sheet, effectiveHeader, [
    {
      updateTable: {
        table: {
          tableId: table.tableId,
          name: effectiveTableName,
          range: {
            sheetId: sheet.getSheetId(),
            startRowIndex: 0,
            endRowIndex: Math.max(sheet.getLastRow(), 1),
            startColumnIndex: 0,
            endColumnIndex: effectiveHeader.length,
          },
        },
        fields: 'name,range',
      },
    },
  ]);
}

function ensureTableRange_(spreadsheetId, sheet, tableName, header) {
  const shouldEnsureCalendarFilter = isCalendarTableMaintenanceRequest_(tableName, header);
  const effectiveTableName = tableName || CONFIG.tableName;
  const effectiveHeader = header || CONFIG.header;
  assertManagedTableHasInlineIdColumn_(effectiveHeader);
  ensureManagedIdColumnHidden_(sheet);
  const spreadsheetModel = getSpreadsheetModel_(spreadsheetId);
  const sheetModel = (spreadsheetModel.sheets || []).find(
    (entry) => entry.properties && entry.properties.sheetId === sheet.getSheetId()
  );
  const table = ((sheetModel && sheetModel.tables) || []).find(
    (entry) => entry.name === effectiveTableName
  );

  if (!table) {
    ensureTable_(spreadsheetId, sheet, effectiveTableName, effectiveHeader);
    if (shouldEnsureCalendarFilter) {
      ensureCalendarStartDateFilter_(sheet);
    }
    return;
  }

  const desiredEndRow = Math.max(sheet.getLastRow(), 1);
  const currentRange = table.range || {};

  const unchanged =
    currentRange.sheetId === sheet.getSheetId() &&
    currentRange.startRowIndex === 0 &&
    currentRange.endRowIndex === desiredEndRow &&
    currentRange.startColumnIndex === 0 &&
    currentRange.endColumnIndex === effectiveHeader.length;

  if (unchanged) {
    if (shouldEnsureCalendarFilter) {
      ensureCalendarStartDateFilter_(sheet);
    }
    return;
  }

  updateTableWithFilterSafeBatch_(spreadsheetId, sheet, effectiveHeader, [
    {
      updateTable: {
        table: {
          tableId: table.tableId,
          name: table.name,
          range: {
            sheetId: sheet.getSheetId(),
            startRowIndex: 0,
            endRowIndex: desiredEndRow,
            startColumnIndex: 0,
            endColumnIndex: effectiveHeader.length,
          },
        },
        fields: 'range',
      },
    },
  ]);

  if (shouldEnsureCalendarFilter) {
    ensureCalendarStartDateFilter_(sheet);
  }
}


function ensureCalendarStartDateFilter_(sheet) {
  const dateColumn = CONFIG.header.indexOf('Date') + 1;
  if (dateColumn <= 0) {
    throw new Error('Calendar Date column is not configured.');
  }

  const filter = ensureCalendarSheetFilter_(sheet);
  const criteria = buildCalendarStartDateFilterCriteria_(CONFIG.importStartDate, dateColumn);
  filter.setColumnFilterCriteria(dateColumn, criteria);
}

function ensureCalendarSheetFilter_(sheet) {
  const rowCount = Math.max(sheet.getLastRow(), 1);
  const columnCount = CONFIG.header.length;
  const existingFilter = sheet.getFilter();
  if (!existingFilter) {
    return sheet.getRange(1, 1, rowCount, columnCount).createFilter();
  }

  const filterRange = existingFilter.getRange();
  const hasExpectedRange =
    filterRange.getRow() === 1 &&
    filterRange.getColumn() === 1 &&
    filterRange.getNumRows() === rowCount &&
    filterRange.getNumColumns() === columnCount;
  if (hasExpectedRange) {
    return existingFilter;
  }

  const savedCriteria = [];
  for (let column = 1; column <= columnCount; column += 1) {
    savedCriteria[column] = existingFilter.getColumnFilterCriteria(column);
  }
  existingFilter.remove();

  const filter = sheet.getRange(1, 1, rowCount, columnCount).createFilter();
  savedCriteria.forEach((criteria, column) => {
    if (criteria && column > 0) {
      filter.setColumnFilterCriteria(column, criteria);
    }
  });
  return filter;
}

function buildCalendarStartDateFilterCriteria_(importStartDate, dateColumn) {
  const parts = parseImportStartDatePartsForFilter_(importStartDate);
  const columnLetter = columnIndexToLetter_(dateColumn);
  const formula = `=$${columnLetter}2>=DATE(${parts.year},${parts.month},${parts.day})`;
  return SpreadsheetApp.newFilterCriteria().whenFormulaSatisfied(formula).build();
}

function parseImportStartDatePartsForFilter_(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(toText_(value).trim());
  if (!match) {
    throw new Error(
      `Invalid CONFIG.importStartDate: "${value}". Use ISO date format YYYY-MM-DD (example: 2024-01-01).`
    );
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  const isRealDate =
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day;
  if (!isRealDate) {
    throw new Error(
      `Invalid CONFIG.importStartDate: "${value}" is not a real calendar date. Use YYYY-MM-DD (example: 2024-01-01).`
    );
  }

  return { year, month, day };
}

function columnIndexToLetter_(columnIndex) {
  if (!Number.isInteger(columnIndex) || columnIndex < 1) {
    throw new Error(`Invalid column index: ${columnIndex}.`);
  }

  let remaining = columnIndex;
  let letters = '';
  while (remaining > 0) {
    const remainder = (remaining - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return letters;
}

function getSpreadsheetModel_(spreadsheetId) {
  return Sheets.Spreadsheets.get(spreadsheetId, {
    fields: 'sheets(properties(sheetId,title),tables(tableId,name,range))',
    includeGridData: false,
  });
}
