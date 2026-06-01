#!/usr/bin/env node
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const context = {
  console,
  Logger: { log() {} },
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

console.log('table-service-test: PASS');
