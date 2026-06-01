function ensureCalendarSheet_(ss) {
  return ensureNamedSheet_(ss, CONFIG.sheetName, false);
}

function ensureInvoicingSheet_(ss) {
  return ensureNamedSheet_(ss, CONFIG.invoicingSheetName, false);
}

function ensureNonBillableSheet_(ss) {
  return ensureNamedSheet_(ss, CONFIG.nonBillableSheetName, false);
}

function ensureStateSheet_(ss) {
  return ensureNamedSheet_(ss, CONFIG.stateSheetName, true);
}

function ensureInvoicingStateSheet_(ss) {
  return ensureNamedSheet_(ss, CONFIG.invoicingStateSheetName, true);
}

function ensureNonBillableStateSheet_(ss) {
  return ensureNamedSheet_(ss, CONFIG.nonBillableStateSheetName, true);
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

function ensureStateHeader_(sheet) {
  ensureHeader_(sheet, CONFIG.stateHeader);
}

function ensureInvoicingStateHeader_(sheet) {
  ensureHeader_(sheet, CONFIG.invoicingStateHeader);
}

function ensureNonBillableStateHeader_(sheet) {
  ensureHeader_(sheet, CONFIG.nonBillableStateHeader);
}

function ensureManagedWorkbookStructure_(ss, spreadsheetId) {
  const sheet = ensureCalendarSheet_(ss);
  const stateSheet = ensureStateSheet_(ss);
  const invoicingSheet = ensureInvoicingSheet_(ss);
  const invoicingStateSheet = ensureInvoicingStateSheet_(ss);
  const nonBillableSheet = ensureNonBillableSheet_(ss);
  const nonBillableStateSheet = ensureNonBillableStateSheet_(ss);

  ensureHeader_(sheet);
  ensureStateHeader_(stateSheet);
  ensureHeader_(invoicingSheet, CONFIG.invoicingHeader, { allowOverwrite: false });
  ensureInvoicingStateHeader_(invoicingStateSheet);
  ensureHeader_(nonBillableSheet, CONFIG.nonBillableHeader, { allowOverwrite: false });
  ensureNonBillableStateHeader_(nonBillableStateSheet);

  assertSheetHasExpectedColumns_(sheet, CONFIG.header);
  assertSheetHasExpectedColumns_(stateSheet, CONFIG.stateHeader);
  assertSheetHasExpectedColumns_(invoicingSheet, CONFIG.invoicingHeader);
  assertSheetHasExpectedColumns_(invoicingStateSheet, CONFIG.invoicingStateHeader);
  assertSheetHasExpectedColumns_(nonBillableSheet, CONFIG.nonBillableHeader);
  assertSheetHasExpectedColumns_(nonBillableStateSheet, CONFIG.nonBillableStateHeader);

  ensureSheetFormatting_(sheet);
  ensureSheetFormatting_(invoicingSheet);
  ensureSheetFormatting_(nonBillableSheet);
  ensureTable_(spreadsheetId, sheet);
  ensureTable_(spreadsheetId, invoicingSheet, CONFIG.invoicingTableName, CONFIG.invoicingHeader);
  ensureTable_(spreadsheetId, nonBillableSheet, CONFIG.nonBillableTableName, CONFIG.nonBillableHeader);

  return {
    sheet,
    stateSheet,
    invoicingSheet,
    invoicingStateSheet,
    nonBillableSheet,
    nonBillableStateSheet,
  };
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
}

function ensureTable_(spreadsheetId, sheet, tableName, header) {
  const effectiveTableName = tableName || CONFIG.tableName;
  const effectiveHeader = header || CONFIG.header;
  const spreadsheetModel = getSpreadsheetModel_(spreadsheetId);
  const sheetModel = (spreadsheetModel.sheets || []).find(
    (entry) => entry.properties && entry.properties.sheetId === sheet.getSheetId()
  );
  const tables = (sheetModel && sheetModel.tables) || [];

  if (tables.length === 0) {
    Sheets.Spreadsheets.batchUpdate(
      {
        requests: [
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
        ],
      },
      spreadsheetId
    );
    return;
  }

  const table = tables.find((entry) => entry.name === effectiveTableName) || tables[0];

  Sheets.Spreadsheets.batchUpdate(
    {
      requests: [
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
      ],
    },
    spreadsheetId
  );
}

function ensureTableRange_(spreadsheetId, sheet, tableName, header) {
  const effectiveTableName = tableName || CONFIG.tableName;
  const effectiveHeader = header || CONFIG.header;
  const spreadsheetModel = getSpreadsheetModel_(spreadsheetId);
  const sheetModel = (spreadsheetModel.sheets || []).find(
    (entry) => entry.properties && entry.properties.sheetId === sheet.getSheetId()
  );
  const table = ((sheetModel && sheetModel.tables) || []).find(
    (entry) => entry.name === effectiveTableName
  );

  if (!table) {
    ensureTable_(spreadsheetId, sheet, effectiveTableName, effectiveHeader);
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
    return;
  }

  Sheets.Spreadsheets.batchUpdate(
    {
      requests: [
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
      ],
    },
    spreadsheetId
  );
}

function getSpreadsheetModel_(spreadsheetId) {
  return Sheets.Spreadsheets.get(spreadsheetId, {
    fields: 'sheets(properties(sheetId,title),tables(tableId,name,range))',
    includeGridData: false,
  });
}
