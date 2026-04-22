function ensureCalendarSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.sheetName);
  }

  return sheet;
}

function ensureStateSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.stateSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.stateSheetName);
  }

  if (!sheet.isSheetHidden()) {
    sheet.hideSheet();
  }

  return sheet;
}

function ensureHeader_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, CONFIG.header.length);
  const current = headerRange.getValues()[0];

  const needsWrite = CONFIG.header.some((value, index) => current[index] !== value);
  if (needsWrite) {
    headerRange.setValues([CONFIG.header]);
  }
}

function ensureStateHeader_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, CONFIG.stateHeader.length);
  const current = headerRange.getValues()[0];

  const needsWrite = CONFIG.stateHeader.some((value, index) => current[index] !== value);
  if (needsWrite) {
    headerRange.setValues([CONFIG.stateHeader]);
  }
}

function ensureSheetFormatting_(sheet) {
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(true);
}

function ensureTable_(spreadsheetId, sheet) {
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
                name: CONFIG.tableName,
                range: {
                  sheetId: sheet.getSheetId(),
                  startRowIndex: 0,
                  endRowIndex: Math.max(sheet.getLastRow(), 1),
                  startColumnIndex: 0,
                  endColumnIndex: CONFIG.header.length,
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

  const table = tables.find((entry) => entry.name === CONFIG.tableName) || tables[0];

  Sheets.Spreadsheets.batchUpdate(
    {
      requests: [
        {
          updateTable: {
            table: {
              tableId: table.tableId,
              name: CONFIG.tableName,
              range: {
                sheetId: sheet.getSheetId(),
                startRowIndex: 0,
                endRowIndex: Math.max(sheet.getLastRow(), 1),
                startColumnIndex: 0,
                endColumnIndex: CONFIG.header.length,
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

function ensureTableRange_(spreadsheetId, sheet) {
  const spreadsheetModel = getSpreadsheetModel_(spreadsheetId);
  const sheetModel = (spreadsheetModel.sheets || []).find(
    (entry) => entry.properties && entry.properties.sheetId === sheet.getSheetId()
  );
  const table = ((sheetModel && sheetModel.tables) || []).find(
    (entry) => entry.name === CONFIG.tableName
  );

  if (!table) {
    ensureTable_(spreadsheetId, sheet);
    return;
  }

  const desiredEndRow = Math.max(sheet.getLastRow(), 1);
  const currentRange = table.range || {};

  const unchanged =
    currentRange.sheetId === sheet.getSheetId() &&
    currentRange.startRowIndex === 0 &&
    currentRange.endRowIndex === desiredEndRow &&
    currentRange.startColumnIndex === 0 &&
    currentRange.endColumnIndex === CONFIG.header.length;

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
                endColumnIndex: CONFIG.header.length,
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
