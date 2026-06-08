#!/usr/bin/env node
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const context = { console };
vm.createContext(context);
['Helper.gs', 'Config.gs', 'Sheet Writer.gs'].forEach((file) => {
  vm.runInContext(fs.readFileSync(file, 'utf8'), context, { filename: file });
});

const formula = context.buildStateFormula_(2);
assert(formula.includes('"Invoicing"'));
assert(!formula.includes('"Invoiced"'));

assert.doesNotThrow(() => context.validateCalendarRowsForWrite_([
  {
    eventKey: 'cal::event',
    values: ['Calendar', 'Event title', new Date('2026-01-01T00:00:00Z'), new Date('2026-01-01T10:00:00Z'), new Date('2026-01-01T11:00:00Z'), 1 / 24, 'Open'],
  },
]));

assert.throws(
  () => context.validateCalendarRowsForWrite_([
    { eventKey: 'cal::event', values: ['Calendar', 'Event title', '', '', '', '', 'Open'] },
  ]),
  /Date, Start, End, Duration/
);

console.log('sheet-writer-test: PASS');
