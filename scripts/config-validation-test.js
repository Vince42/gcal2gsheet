#!/usr/bin/env node
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const context = { console };
vm.createContext(context);
['Helper.gs', 'Config.gs'].forEach((file) => {
  vm.runInContext(fs.readFileSync(file, 'utf8'), context, { filename: file });
});

function mergedConfig(overrides) {
  const config = context.mergeConfigWithDefaults_(overrides);
  context.validateConfig_(config);
  return config;
}

const canonical = mergedConfig({});
assert.deepEqual(canonical.header, ['ID', 'Calendar', 'Event', 'Date', 'Start', 'End', 'Duration', 'Status']);
assert.deepEqual(canonical.invoicingHeader, ['EventID', 'Calendar', 'Event', 'Date', 'Start', 'End', 'Duration', 'Customer', 'Project', 'InvoiceNumber', 'InvoiceDate']);
assert.deepEqual(canonical.nonBillableHeader, ['EventID', 'Calendar', 'Event', 'Date', 'Start', 'End', 'Duration', 'Reason']);

const staleInvoiceCalendarHeader = mergedConfig({
  header: ['Calendar', 'Event', 'Date', 'Start', 'End', 'Duration', 'Customer', 'Project', 'InvoiceNumber', 'InvoiceDate'],
});
assert.deepEqual(staleInvoiceCalendarHeader.header, canonical.header);

const staleInlineInvoiceCalendarHeader = mergedConfig({
  header: ['EventID', 'Calendar', 'Event', 'Date', 'Start', 'End', 'Duration', 'Customer', 'Project', 'InvoiceNumber', 'InvoiceDate'],
});
assert.deepEqual(staleInlineInvoiceCalendarHeader.header, canonical.header);

const malformedHeaderValues = mergedConfig({
  header: ['bad'],
  invoicingHeader: ['bad'],
  nonBillableHeader: ['bad'],
});
assert.deepEqual(malformedHeaderValues.header, canonical.header);
assert.deepEqual(malformedHeaderValues.invoicingHeader, canonical.invoicingHeader);
assert.deepEqual(malformedHeaderValues.nonBillableHeader, canonical.nonBillableHeader);

const screenshotStyleLegacyConfig = {
  sheetName: 'Calendar',
  stateSheetName: '_custom_calendar_state',
  tableName: 'Calendar',
  statusCell: 'L1',
  importStartDate: '2024-01-01',
  calendarNames: ['Event', 'dedc', 'EEC', 'CTG'],
  defaultCalendarName: 'Event',
  header: ['Calendar', 'Event', 'Date', 'Start', 'End', 'Duration', 'Customer', 'Project', 'InvoiceNumber', 'InvoiceDate'],
  stateHeader: ['EventKey', 'RowKind'],
  rowKind: {
    normal: 'NORMAL',
    changedCopy: 'CHANGED_COPY',
    unmanaged: 'UNMANAGED',
  },
  propertyPrefix: 'CALSYNC_TOKEN_',
  configPropertyKey: 'CALSYNC_CONFIG_JSON',
  colors: {
    normal: '#000000',
    invoiced: '#7A1F1F',
    changed: '#1B5E20',
  },
  menu: {
    title: 'Calendar Import',
    item: 'Update calendar sheet',
  },
  toastTitle: 'Calendar import',
};
const healedLegacyConfig = context.mergeConfigWithDefaults_(screenshotStyleLegacyConfig);
context.validateConfigStrictWithSchema_(healedLegacyConfig, screenshotStyleLegacyConfig, { allowedKeys: [] });
assert.deepEqual(healedLegacyConfig.header, canonical.header);
assert.deepEqual(healedLegacyConfig.invoicingHeader, canonical.invoicingHeader);
assert.deepEqual(healedLegacyConfig.nonBillableHeader, canonical.nonBillableHeader);
assert.equal(healedLegacyConfig.statusCell, 'J1');
assert.equal(healedLegacyConfig.menu.title, 'Invoicing');
assert.equal(healedLegacyConfig.toastTitle, 'Invoicing');
assert.deepEqual(healedLegacyConfig.calendarNames, ['Event', 'dedc', 'EEC', 'CTG']);
assert.equal(healedLegacyConfig.defaultCalendarName, 'Event');

const customStatusCellConfig = mergedConfig({ statusCell: 'k1' });
assert.equal(customStatusCellConfig.statusCell, 'K1');

assert.throws(
  () => context.validateConfigStrictWithSchema_(
    context.mergeConfigWithDefaults_({ unexpectedKey: true }),
    { unexpectedKey: true },
    { allowedKeys: [] }
  ),
  /unknown key "unexpectedKey"/
);

console.log('config-validation-test: PASS');
