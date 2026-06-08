#!/usr/bin/env node
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const context = {
  console,
  Logger: { log() {} },
  SpreadsheetApp: {
    newFilterCriteria() {
      return {
        whenFormulaSatisfied(formula) {
          this.formula = formula;
          return this;
        },
        build() {
          return { formula: this.formula };
        },
      };
    },
  },
};
vm.createContext(context);
['Helper.gs', 'Config.gs', 'Table Service.gs'].forEach((file) => {
  vm.runInContext(fs.readFileSync(file, 'utf8'), context, { filename: file });
});

const candidates = context.collectLegacyStateSheetNameCandidatesFromConfig_({
  stateSheetName: '_custom_calendar_state',
  invoicingStateSheetName: '_custom_invoicing_state',
  nonBillableStateSheetName: '_custom_non_billable_state',
});
assert.deepEqual(candidates.calendar, ['_custom_calendar_state', '_calendar_state']);
assert.deepEqual(candidates.invoicing, ['_custom_invoicing_state', '_invoicing_state']);
assert.deepEqual(candidates.nonBillable, ['_custom_non_billable_state', '_non_billable_state']);

context.rememberLegacyStateSheetNameCandidates_({
  stateSheetName: '_remembered_calendar_state',
  invoicingStateSheetName: '_remembered_invoicing_state',
  nonBillableStateSheetName: '_remembered_non_billable_state',
});
const remembered = context.getLegacyStateSheetNameCandidates_();
assert(remembered.calendar.includes('_remembered_calendar_state'));
assert(remembered.calendar.includes('_calendar_state'));
assert(remembered.invoicing.includes('_remembered_invoicing_state'));
assert(remembered.nonBillable.includes('_remembered_non_billable_state'));

function sheet(name, id) {
  return {
    name,
    id,
    getSheetId() { return id; },
  };
}
const customCalendarState = sheet('_custom_calendar_state', 1);
const defaultCalendarState = sheet('_calendar_state', 2);
const customInvoiceState = sheet('_custom_invoicing_state', 3);
const spreadsheet = {
  byName: new Map([
    ['_custom_calendar_state', customCalendarState],
    ['_calendar_state', defaultCalendarState],
    ['_custom_invoicing_state', customInvoiceState],
  ]),
  getSheetByName(name) {
    return this.byName.get(name) || null;
  },
};

assert.strictEqual(
  context.resolveLegacyStateSheet_(spreadsheet, ['_custom_calendar_state', '_calendar_state']),
  customCalendarState
);
assert.strictEqual(
  context.resolveLegacyStateSheet_(spreadsheet, ['_missing_state', '_calendar_state']),
  defaultCalendarState
);
assert.deepEqual(
  context.collectLegacyStateSheets_(spreadsheet, ['_custom_calendar_state', '_calendar_state', '_custom_calendar_state']).map((entry) => entry.name),
  ['_custom_calendar_state', '_calendar_state']
);

assert.deepEqual(context.parseImportStartDatePartsForFilter_('2024-01-01'), {
  year: 2024,
  month: 1,
  day: 1,
});
assert.throws(() => context.parseImportStartDatePartsForFilter_('2024-02-31'), /not a real calendar date/);
assert.equal(context.columnIndexToLetter_(1), 'A');
assert.equal(context.columnIndexToLetter_(4), 'D');
assert.equal(context.columnIndexToLetter_(28), 'AB');
assert.deepEqual(
  context.buildCalendarStartDateFilterCriteria_('2024-01-01', 4),
  { formula: '=$D2>=DATE(2024,1,1)' }
);

function mockRange(row, column, numRows, numColumns, onCreateFilter) {
  return {
    getRow() { return row; },
    getColumn() { return column; },
    getNumRows() { return numRows; },
    getNumColumns() { return numColumns; },
    createFilter() { return onCreateFilter(this); },
  };
}

function mockFilter(range, criteriaByColumn) {
  return {
    removed: false,
    applied: [],
    getRange() { return range; },
    getColumnFilterCriteria(column) { return criteriaByColumn[column] || null; },
    setColumnFilterCriteria(column, criteria) {
      criteriaByColumn[column] = criteria;
      this.applied.push({ column, criteria });
    },
    remove() { this.removed = true; },
  };
}

let createdFilter;
const resizedSheet = {
  existingFilter: null,
  getLastRow() { return 10; },
  getFilter() { return this.existingFilter; },
  getRange(row, column, numRows, numColumns) {
    return mockRange(row, column, numRows, numColumns, (range) => {
      createdFilter = mockFilter(range, {});
      this.existingFilter = createdFilter;
      return createdFilter;
    });
  },
};
const calendarHeader = context.mergeConfigWithDefaults_({}).header;
const oldRange = mockRange(1, 1, 5, calendarHeader.length, () => null);
const statusCriteria = { status: 'Open' };
const oldFilter = mockFilter(oldRange, { 8: statusCriteria });
resizedSheet.existingFilter = oldFilter;
const ensuredFilter = context.ensureCalendarSheetFilter_(resizedSheet);
assert.equal(oldFilter.removed, true);
assert.strictEqual(ensuredFilter.getColumnFilterCriteria(8), statusCriteria);
assert.equal(ensuredFilter.getRange().getNumRows(), 10);

context.ensureCalendarStartDateFilter_(resizedSheet);
assert.deepEqual(createdFilter.getColumnFilterCriteria(4), { formula: '=$D2>=DATE(2024,1,1)' });


function mockSheetFilterForTable(range, criteriaByColumn) {
  return mockFilter(range, criteriaByColumn || {});
}

let tableModel = {
  sheets: [
    {
      properties: { sheetId: 42, title: 'Calendar' },
      tables: [],
    },
  ],
};
let batchRequests = [];
let tableUpdateFilter;
let restoredFilter;
context.Sheets = {
  Spreadsheets: {
    get() {
      return tableModel;
    },
    batchUpdate(body) {
      assert.equal(tableUpdateFilter.removed, true);
      batchRequests.push(body.requests);
    },
  },
};

function mockTableSheet() {
  return {
    filter: null,
    hiddenColumns: [],
    getSheetId() { return 42; },
    getLastRow() { return 10; },
    hideColumns(column) { this.hiddenColumns.push(column); },
    getFilter() { return this.filter && !this.filter.removed ? this.filter : null; },
    getRange(row, column, numRows, numColumns) {
      return mockRange(row, column, numRows, numColumns, (range) => {
        restoredFilter = mockFilter(range, {});
        this.filter = restoredFilter;
        return restoredFilter;
      });
    },
  };
}

const tableSheet = mockTableSheet();
tableUpdateFilter = mockSheetFilterForTable(
  mockRange(1, 1, 10, calendarHeader.length, () => null),
  { 8: statusCriteria }
);
tableSheet.filter = tableUpdateFilter;
context.ensureTable_('spreadsheet-id', tableSheet);
assert.equal(batchRequests.length, 1);
assert(batchRequests[0][0].addTable);
assert.equal(restoredFilter.getColumnFilterCriteria(8), statusCriteria);
assert.equal(restoredFilter.getRange().getNumRows(), 10);

tableModel = {
  sheets: [
    {
      properties: { sheetId: 42, title: 'Calendar' },
      tables: [
        {
          tableId: 'tbl1',
          name: 'Calendar',
          range: {
            sheetId: 42,
            startRowIndex: 0,
            endRowIndex: 5,
            startColumnIndex: 0,
            endColumnIndex: calendarHeader.length,
          },
        },
      ],
    },
  ],
};
batchRequests = [];
restoredFilter = null;
tableUpdateFilter = mockSheetFilterForTable(
  mockRange(1, 1, 5, calendarHeader.length, () => null),
  { 8: statusCriteria }
);
tableSheet.filter = tableUpdateFilter;
context.ensureTableRange_('spreadsheet-id', tableSheet);
assert.equal(batchRequests.length, 1);
assert(batchRequests[0][0].updateTable);
assert.equal(restoredFilter.getColumnFilterCriteria(4).formula, '=$D2>=DATE(2024,1,1)');
assert.equal(restoredFilter.getColumnFilterCriteria(8), statusCriteria);

console.log('table-service-test: PASS');
