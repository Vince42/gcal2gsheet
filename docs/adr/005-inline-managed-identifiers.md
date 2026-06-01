# ADR-005: Inline Managed Identifiers

## Status

Accepted.

## Context

The original workbook design stored Calendar event identity in hidden parallel sheets:

- `_calendar_state`
- `_invoicing_state`
- `_non_billable_state`

Those sheets had to stay row-aligned with visible data sheets. Any row insertion, deletion, filtering workflow, table maintenance, or manual recovery could break that alignment and cause actions such as marking rows as Invoiced or Non-Billable to target the wrong event.

The managed Google Sheets table objects now provide a better place to keep row identity: the first column of each managed data table. The identifier can stay in the table range for formulas, repairs, and register cleanup while being hidden from users so it does not distract from normal invoicing work.

## Decision

Managed row identifiers MUST live in the data tables, not in separate state sheets.

The Calendar table MUST use a first column named `ID`. Its value is the stable Calendar event key.

The Invoicing and Non-Billable tables MUST use a first column named `EventID`. Its value is the Calendar `ID` referenced by that register row.

The first identifier column MUST be hidden whenever the managed workbook structure, sheet formatting, or table range is maintained.

The hidden state sheets `_calendar_state`, `_invoicing_state`, and `_non_billable_state` are obsolete. Workbook setup MAY read them during legacy migration, but after migration they MUST NOT remain part of the active data model.

The Calendar sheet MUST NOT persist `RowKind` as a worksheet column. Any row-kind behavior that remains necessary must stay internal to the import/rebuild model.

## Consequences

- Row identity travels with visible table rows, so selected-row actions can read the event ID from the selected Calendar row directly.
- Status formulas can reference register table `EventID` columns directly.
- Register cleanup can delete moved rows by matching first-column `EventID` values without maintaining parallel state sheets.
- Legacy migration must preserve user-entered invoice metadata before retired Calendar invoice columns are cleared.
- Tests and release gates should no longer require visible/state row alignment for the removed state sheets; they should instead verify inline identifier presence and hidden first-column behavior.

## Invariants

- Calendar column A is `ID` and is hidden.
- Invoicing column A is `EventID` and is hidden.
- Non-Billable column A is `EventID` and is hidden.
- Managed table ranges include the hidden identifier column.
- Legacy state sheets are migration inputs only and must be deleted after migration succeeds.
- User-entered register metadata in legacy Calendar invoice columns must be migrated even when `InvoiceNumber` is blank.
