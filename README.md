# Google Sheets Calendar Import for Invoicing Preparation

## Purpose

This Apps Script project imports timed Google Calendar events into a Google Sheets workbook for invoicing preparation.

The script builds and maintains one visible worksheet named `Calendar` and one hidden technical worksheet named `_calendar_state`.

The visible worksheet contains the invoicing preparation table.  
The hidden worksheet stores technical row identity and row type information.

The script is designed to:

- import timed events from multiple calendars
- exclude all-day events
- exclude cancelled events
- expand recurring events into normal event instances
- ignore future events
- preserve historical rows before a configured start date
- preserve invoiced rows
- detect changed invoiced events and create a new follow-up row
- remove duplicates according to calendar precedence rules
- keep all date and time values as real spreadsheet values
- preserve a native Google Sheets table object on the visible worksheet

---

## Calendars

Current calendar list:

- `Event`
- `dedc`
- `EEC`
- `CTG`

`Event` is treated as the default calendar.  
Specific calendars beat the default calendar during duplicate cleanup.

---

## Workbook structure

### Visible worksheet

The visible worksheet is always:

- `Calendar`

It contains exactly these columns in this order:

1. `Calendar`
2. `Event`
3. `Date`
4. `Start`
5. `End`
6. `Duration`
7. `Customer`
8. `Project`
9. `InvoiceNumber`
10. `InvoiceDate`

### Hidden worksheet

The technical worksheet is always:

- `_calendar_state`

It stores one technical row per visible data row:

- `EventKey`
- `RowKind`

This sheet must remain hidden and must not be edited manually.

---

## Managed scope

The script manages only events inside the active scope.

### Lower bound

The lower bound is configured in:

- `CONFIG.importStartDate`

Format:

- `yyyy-mm-dd`

### Upper bound

The upper bound is always the current timestamp at execution time.

### Effective rule

An event is managed only if:

- `Start >= importStartDate`
- `End <= now`

This means:

- future events must not be imported
- recurring events may be expanded by the Calendar API, but future instances are filtered out
- rows before the configured start date remain in the sheet and are preserved untouched

---

## Import rules

The script imports:

- timed events only
- recurring events as expanded timed instances

The script excludes:

- all-day events
- cancelled events
- future events
- events ending after the current timestamp

---

## Update rules

The script performs a full rebuild on every run.

That is intentional.

The rebuild strategy is:

1. read the current visible sheet
2. read the hidden technical state sheet
3. fetch all relevant calendar events inside the managed scope
4. normalize imported events
5. remove duplicates among imported events
6. merge imported events with existing rows
7. preserve invoicing-related user-entered columns
8. generate changed follow-up rows for changed invoiced events
9. remove duplicates again on the final row set
10. rewrite the visible sheet in bulk
11. rewrite the hidden state sheet in bulk
12. refresh number formats, row colors, and the table range

This avoids complicated incremental-sync edge cases and keeps behavior deterministic.

---

## Invoicing logic

### Uninvoiced rows

A row is considered uninvoiced if:

- `InvoiceNumber` is empty

Uninvoiced rows are updated silently in place when the matching calendar event changes.

### Invoiced rows

A row is considered invoiced if:

- `InvoiceNumber` is not empty

If an invoiced event changes:

- the original row remains unchanged
- a new follow-up row is created directly below the historical invoiced row
- the new row gets `RowKind = CHANGED_COPY`
- the user is notified after the run

### Invoice date

`InvoiceDate` currently has no control logic.

It is stored and preserved, but not used as a decision criterion.

---

## Duplicate handling

A duplicate is defined by identical values in:

- `Event`
- `Date`
- `Start`
- `End`

### Duplicate rules

#### Rule one: duplicates inside the same calendar

If the same duplicate appears multiple times in the same calendar:

- all rows of that calendar for that duplicate are removed

#### Rule two: specific calendars beat the default calendar

If at least one specific calendar row exists for a duplicate:

- all `Event` rows for that duplicate are removed

#### Rule three: multiple specific calendars

If duplicates still remain across multiple non-default calendars:

- only one survives according to configured calendar priority
- the rest are removed

Current priority is derived from:

- `dedc`
- `EEC`
- `CTG`
- `Event`

This logic is applied:

- once on the imported events
- once again on the final rebuilt rows

---

## Formatting rules

The visible worksheet is maintained with:

- hidden gridlines
- frozen header row
- real spreadsheet date values in `Date`
- real spreadsheet time values in `Start`
- real spreadsheet time values in `End`
- real spreadsheet duration values in `Duration`

Formats:

- `Date` → `yyyy-mm-dd`
- `Start` → `hh:mm`
- `End` → `hh:mm`
- `Duration` → `hh:mm`
- `InvoiceDate` → `yyyy-mm-dd`

### Row colors

- normal rows → black font
- invoiced rows → dark red font
- changed follow-up rows → dark green font

---

## Google Sheets table object

The visible worksheet contains one native Google Sheets table object:

- table name: `Calendar`

The script keeps the table range aligned with the current data body through the Sheets API.

The table itself is not used as the business logic store.  
It is treated as a workbook UI structure that must remain present and consistent.

---

## Progress reporting

During execution, progress is written to:

- toast notifications
- `Calendar!L1`

This gives the user visible runtime feedback without using modal dialogs during processing.

---

## Files

### `Code.gs`
Main orchestration entrypoints:

- `onOpen()`
- `updateCalendarSheets()`

### `Config.gs`
Constants and configuration validation.

### `Ui.gs`
Menu support, progress reporting, UI state restore.

### `Scope.gs`
Scope and time-bound logic.

### `Calendar Service.gs`
Calendar resolution and event fetching.

### `State Store.gs`
Read and write logic for `_calendar_state`.

### `Rebuild Engine.gs`
Main merge and rebuild logic.

### `Duplicate Engine.gs`
Duplicate detection and removal.

### `Sheet Writer.gs`
Visible sheet writing, number formats, row colors.

### `Table Service.gs`
Google Sheets table object maintenance.

### `Helper.gs`
Shared low-level helpers.

---

## Main callable function

The main callable function is:

- `updateCalendarSheets()`

This is the function intended for manual execution and menu invocation.

---

## Required services

Enable these Advanced Services in the Apps Script project:

- `Google Sheets API`
- `Google Calendar API`

Without them, table maintenance and Calendar API event listing will fail.

---

## Operational assumptions

- The workbook timezone is the source of truth for formatting.
- Calendar names are resolved by exact visible calendar name.
- The hidden state sheet must stay aligned row-for-row with the visible sheet.
- User edits are expected only in the visible business columns.
- The hidden technical sheet is internal implementation state.

---

## Known design decisions

### Full rebuild instead of incremental sync

The project currently uses a full rebuild on each run.

Reason:

- incremental sync with `syncToken` complicates scope changes
- incremental sync cannot be combined with several useful query parameters
- correctness is currently more important than sync-token optimization

### Hidden technical state sheet instead of developer metadata

The project uses `_calendar_state` instead of row-level developer metadata.

Reason:

- developer metadata storage limits were reached
- hidden sheet storage is easier to inspect and more scalable for this workload

---

## Safe editing rules for future changes

When changing this codebase:

- do not mix UI code with business logic
- keep date-scope rules inside `Scope.gs`
- keep duplicate logic inside `Duplicate Engine.gs`
- keep calendar fetch normalization inside `Calendar Service.gs`
- keep hidden state row alignment intact
- avoid row-by-row writes when a bulk write is possible
- keep all visible date and time cells as real spreadsheet values, never text
- do not introduce helper columns into the visible sheet
- do not remove the table object unless explicitly intended

---

## Future extension ideas

Possible future extensions:

- calendar selection dialog
- configurable duplicate priority
- logging sheet or debug mode
- test harness for pure logic modules
- optional incremental sync mode
- import of Google Calendar description field in addition to event title
- configurable status cell
- stricter validation of manual row edits

---

## Automated push to Apps Script (GitHub Actions)

This repository includes a workflow to push source changes directly to Google Apps Script on every push to `main` (and manually via `workflow_dispatch`).

Workflow file:

- `.github/workflows/apps-script-push.yml`
- `appsscript.json` (required by `clasp push`)

Security model:

- credentials are never committed to the repository
- local credential files are ignored by `.gitignore` (`.clasp.json`, `.clasprc.json`)
- CI receives credentials only via GitHub Secrets

Required GitHub repository secrets:

- `CLASPRC_JSON`  
  The full JSON contents of a valid `~/.clasprc.json` for `clasp`.
  Base64-encoded JSON is also accepted.

Configured target Script ID:

- `1XlO8Fb7sktGCrmdqbwtpgLarTw6HpoXUAA7Vv6oakA5OcMDmqHSTm0lC`

Optional safety recommendation:

- protect the `main` branch so only reviewed merges trigger production pushes

### Making CI publishing reliable (avoiding `PERMISSION_DENIED`)

You cannot guarantee "always" in OAuth systems (tokens can be revoked/expired, account access can be removed), but you can make failures rare and recoverable:

1. Use a dedicated deployment Google account (bot user), not a personal account.
2. Make that account **Editor** (or Owner) on the target Apps Script project.
3. Generate `~/.clasprc.json` while logged in as that deployment account and store it in the GitHub secret `CLASPRC_JSON`.
4. If publishing fails with permission errors, rotate `CLASPRC_JSON` from the same deployment account and re-run the workflow.
5. Keep the Apps Script API enabled for that Google Cloud project/account.

This workflow now performs explicit auth/script-access checks before push and prints remediation guidance when permission errors are detected.
