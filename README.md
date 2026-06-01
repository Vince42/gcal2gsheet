# Google Sheets Calendar Import for Invoicing Preparation

## Purpose

This Apps Script project imports timed Google Calendar events into a Google Sheets workbook for invoicing preparation.

The script builds and maintains managed worksheets named `Calendar`, `Invoicing`, and `Non-Billable`.

The `Calendar` worksheet contains the invoicing preparation table. Hidden first-column identifiers (`ID` / `EventID`) store technical row identity inside each managed table.

The script is designed to:

- import timed events from multiple calendars
- exclude all-day events
- exclude cancelled events
- expand recurring events into normal event instances
- ignore future events
- preserve historical rows before a configured start date
- preserve invoiced rows
- remove duplicates according to calendar precedence rules
- keep all date and time values as real spreadsheet values
- preserve a native Google Sheets table object on the visible worksheet

---

## Calendars

Calendars are configured in `ConfigJson` on the `Config` sheet:

- `calendarNames` lists the visible Google Calendar names to import.
- `defaultCalendarName` identifies the general/default calendar used for duplicate precedence.

Specific calendars beat the configured default calendar during duplicate cleanup.
Keep calendar names and priority order in configuration rather than documenting personal calendar names in this README.

---

## Workbook structure

### Visible worksheet

The visible worksheet is always:

- `Calendar`

It contains exactly these columns in this order:

1. `ID` (hidden)
2. `Calendar`
3. `Event`
4. `Date`
5. `Start`
6. `End`
7. `Duration`
8. `Status`

`Status` is a formula derived from the hidden `ID` column and the durable register sheets' hidden `EventID` columns. The normal states are:

- `Open` — no matching register row exists yet
- `Invoiced` — the event is listed in `Invoicing`
- `Non-billable` — the event is listed in `Non-Billable`

### Invoicing worksheet

The durable invoicing worksheet is always:

- `Invoicing`

It contains the invoice register table with these columns:

1. `EventID` (hidden)
2. `Calendar`
3. `Event`
4. `Date`
5. `Start`
6. `End`
7. `Duration`
8. `Customer`
9. `Project`
10. `InvoiceNumber`
11. `InvoiceDate`

### Non-billable worksheet

The durable non-billable register worksheet is always:

- `Non-Billable`

It contains event records that should not become invoices:

1. `EventID` (hidden)
2. `Calendar`
3. `Event`
4. `Date`
5. `Start`
6. `End`
7. `Duration`
8. `Reason`

### Hidden identity columns

The first column of every managed table is hidden and must not be edited manually. Legacy hidden state sheets (`_calendar_state`, `_invoicing_state`, and `_non_billable_state`) are migration inputs only and are deleted after successful workbook structure maintenance.

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

The script performs a full import of the configured Calendar scope and rebuilds the managed worksheet rows on every run.

That is intentional. A full import is slower than an incremental sync, but it gives the import a self-healing property: if a managed event row was accidentally removed from the visible sheet, the next update can recreate the calendar-derived row from Google Calendar as long as the event is still inside the managed scope.

The rebuild strategy is:

1. check that all managed sheets exist
2. check that each managed sheet has the expected columns
3. read the current managed sheets
4. read inline hidden `ID` / `EventID` values
5. fetch all calendar event data for the managed scope
6. normalize imported events
7. remove duplicates among imported events
8. merge imported events with existing rows
9. migrate legacy invoice data from old `Calendar` invoice columns into `Invoicing` when needed
10. mark imported rows as `Invoiced`, `Non-billable`, or `Open` by looking up their hidden `ID` in register `EventID` columns
11. remove duplicates again on eligible final rows
12. rewrite the managed `Calendar` sheet in bulk
13. refresh number formats, row colors, hidden ID columns, and the table ranges

This keeps row identity inside managed table rows and favors correctness over minimizing Calendar API fetch volume. If a managed sheet is missing, the script recreates it; if expected columns do not match after repair, the update stops instead of silently writing into the wrong shape.

---

## Status logic

### Open rows

A calendar row is `Open` if:

- its hidden `ID` does not exist in the Invoicing or Non-Billable `EventID` column

Open rows are updated silently in place when the matching calendar event changes.

### Invoiced rows

A calendar row is `Invoiced` if:

- its hidden `ID` exists in the `Invoicing` register's `EventID` column

### Non-billable rows

A calendar row is `Non-billable` if:

- its hidden `ID` exists in the `Non-Billable` register's `EventID` column

This name is preferred over `Non-Invoicing` because it describes the business decision: the event should not be billed.

If an invoiced or non-billable event changes in Google Calendar, the single `Calendar` import row is updated from Google Calendar and keeps its register-derived status. No duplicate follow-up row is created.

### Invoice date

`InvoiceDate` currently has no control logic.

It is stored and preserved in `Invoicing`, but not used as a decision criterion.

### Recovery after accidental row deletion

A later update can recreate missing calendar-derived rows because every update performs a full import of the managed scope. The recovery boundary is the data source:

- Calendar fields can be imported again if the event still exists and is still inside the managed scope.
- User-entered invoicing data (`Customer`, `Project`, `InvoiceNumber`, `InvoiceDate`) is preserved in `Invoicing`, keyed by the hidden `EventID` column.
- Non-billable decisions are preserved in `Non-Billable`, keyed by the hidden `EventID` column.
- If a register row is deleted from its managed register table, that business metadata cannot be reconstructed from Google Calendar. Restore it from Google Sheets version history, a backup, or another business record.

### User-entered information and adjacent columns

Only managed table columns are row-aware during rebuilds:

- Calendar import/review columns live in `Calendar`.
- Durable invoice business columns live in `Invoicing`.
- Durable non-billable decisions live in `Non-Billable`.
- Technical identity lives in hidden first-column `ID` / `EventID` values inside the managed tables.

Do not store per-event business information in columns adjacent to either managed table. The script rewrites and sorts managed rows in bulk, so adjacent cells outside managed tables are not guaranteed to stay aligned with the same event. If more per-event fields are required, add them deliberately to the managed schema and state-preservation logic rather than placing them beside the table.

---

## Architecture: separate status registers

The status model treats the Calendar import as a reproducible source view and business decisions as durable register records:

1. Keep `Calendar` as an import/review table that is always rebuilt from Google Calendar.
2. Keep `Invoicing` as the invoice register table, with one row per invoiced event.
3. Keep `Non-Billable` as the non-billable register table, with one row per event that should not be billed.
4. Store stable event identity (`EventKey`, built from calendar ID and event ID) in hidden first-column `ID` / `EventID` values.
5. Mark imported calendar rows as `Invoiced`, `Non-billable`, or `Open` by looking up their hidden `ID` in register `EventID` columns.
6. Preserve business metadata in register sheets, not in cells that can disappear during import-table repair.

Invoice register columns:

- `Calendar`
- `Event`
- `Date`
- `Start`
- `End`
- `Duration`
- `Customer`
- `Project`
- `InvoiceNumber`
- `InvoiceDate`

Non-billable register columns:

- `Calendar`
- `Event`
- `Date`
- `Start`
- `End`
- `Duration`
- `Reason`

With this model, a deleted import row is harmless: the full import recreates the calendar row, and the register sheets mark it with the correct status again by `EventKey`. If the calendar event later changes, the single Calendar import row updates from Google Calendar while its register-derived status remains intact.

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

- all rows from the configured default calendar for that duplicate are removed

#### Rule three: multiple specific calendars

If duplicates still remain across multiple non-default calendars:

- only one survives according to configured calendar priority
- the rest are removed

Priority is derived from the configured `calendarNames` order, with non-default calendars considered before `defaultCalendarName`.

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
- `InvoiceDate` in `Invoicing` → `yyyy-mm-dd`

### Row colors

Rows are no longer color-encoded by status. The `Status` column is the source of truth, and update runs reset the managed `Calendar` rows to the normal font color.

---

## Google Sheets table objects

The visible managed worksheets contain native Google Sheets table objects:

- `Calendar` table on the `Calendar` sheet
- `Invoicing` table on the `Invoicing` sheet
- `NonBillable` table on the `Non-Billable` sheet

The script keeps all managed table ranges aligned with their current data bodies through the Sheets API.

The table names are formula-compatible identifiers; for example, the `Non-Billable` sheet uses the `NonBillable` table name because Google Sheets table names cannot contain hyphens.

The tables themselves are not used as hidden business logic stores.
They are treated as workbook UI structures that must remain present and consistent.

---

## Menu actions

The custom menu contains:

- `Filter for` → `Open`, `Invoiced`, `Non-Billable`
- `Mark as` → `Invoiced`, `Non-Billable`

`Filter for` applies a `Status` filter on the `Calendar` sheet and leaves any existing date/start filter criteria in place. If the native table filter cannot be controlled directly, the script falls back to hiding non-matching rows while keeping date-filtered rows hidden.

`Mark as` reads the currently selected `Calendar` rows and appends those rows to the selected durable register. Selections may be adjacent or non-adjacent; if cells B5, C7, D7, and E10 are selected, rows 5, 7, and 10 are marked. After the register state is written, the `Status` formula changes those rows to the corresponding status.

---

## Progress reporting

During execution, progress is written to:

- toast notifications
- the configured status cell on the `Calendar` sheet (`CONFIG.statusCell`)

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
Read existing Calendar rows using inline hidden `ID` values.

### `Invoicing Store.gs`
Read, migrate, and look up invoice-register rows by hidden `EventID`.

### `Non Billable Store.gs`
Read, repair, and look up non-billable register rows by hidden `EventID`.

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
- Hidden `ID` / `EventID` columns must remain present in the managed table ranges and hidden from users.
- User invoice edits are expected in the visible `Invoicing` business columns.
- User non-billable decisions are expected in the visible `Non-Billable` register.
- Legacy hidden state sheets are migration inputs only and are deleted after successful migration.

---

## Known design decisions

### Full-scope import for self-healing reconciliation

The project fetches the full managed Calendar scope and rebuilds the managed worksheet rows on each run.

Reason:

- missing visible rows can be recreated from Google Calendar on the next update
- the visible `Calendar` / `Invoicing` / `Non-Billable` sheets must retain hidden inline `ID` / `EventID` values
- scope-affecting configuration changes require full-scope reasoning rather than stale incremental assumptions
- correctness and trust are more important than minimizing API fetches

### Hidden inline identifiers instead of developer metadata

The project uses hidden first-column `ID` / `EventID` values instead of row-level developer metadata.

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
- keep hidden inline row identity intact
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

Target Apps Script IDs are environment-specific and must not be documented with personal or production identifiers in this README. Store deployment targets in local `.clasp.json` files or protected CI secrets, never in committed documentation unless they are intentionally public test fixtures.

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

---

## Live testing on a real Google Sheet

Yes — you can test this project against a real Google Sheet using a dedicated test spreadsheet and Apps Script project.

### Security-first setup

- Use a **separate test Google account** or restricted test workspace where possible.
- Never commit `.clasp.json` or `.clasprc.json` (already gitignored).
- Keep OAuth/session credentials in your local environment only.
- Do not run live tests against production invoice sheets.

### Fast runtime verification

This repository includes a live check script:

- `bash scripts/live-sheet-check.sh`

What it does:

1. verifies `clasp` is installed
2. verifies `.clasp.json` exists locally
3. executes remote `onOpen()` via Apps Script Execution API path (`clasp run onOpen`)
4. optionally executes `updateCalendarSheets()` when you explicitly allow write testing

Write test opt-in:

- `RUN_UPDATE_CALENDAR_SHEETS=1 bash scripts/live-sheet-check.sh`

This gives you an actual runtime check in Google’s environment (not only static grep-based checks) while keeping credentials out of the repo.
