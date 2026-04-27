# Agents

This repository contains a modular Google Apps Script solution for importing Google Calendar events into Google Sheets for invoicing preparation.

The code is intentionally split by responsibility.  
Any automated coding agent must preserve that separation.

## Global working rules

- Keep the current modular structure.
- Prefer minimal, robust changes.
- Do not move logic across modules without a strong reason.
- Do not introduce visible helper columns into the `Calendar` sheet.
- Do not replace real spreadsheet date/time values with text.
- Do not remove the hidden `_calendar_state` sheet design unless explicitly requested.
- Do not remove the Google Sheets native table object unless explicitly requested.
- Preserve the main callable function `updateCalendarSheets()`.
- Preserve menu creation in `onOpen()`.
- `onOpen()` must always add the custom menu, even when configuration is invalid, so users retain recovery access to the configuration dialog.
- Keep bulk sheet writes preferred over row-by-row writes.
- Treat correctness as more important than clever optimization.
- Avoid speculative fixes. Trace the logic path first.

## Domain rules

### Event scope
An event is managed only if:

- `Start >= CONFIG.importStartDate`
- `End <= now`

Future events must never be written.

### Event types
Import only:

- timed events

Exclude:

- all-day events
- cancelled events

Recurring events may be expanded by the Calendar API, but only past instances inside the active scope may survive filtering.

### Historical rows before the start date
Rows before `CONFIG.importStartDate` must remain untouched.

They may stay in the visible sheet, but they must not be re-imported, updated, or otherwise changed by calendar refresh logic.

### Invoicing logic
- Empty `InvoiceNumber` means uninvoiced.
- Non-empty `InvoiceNumber` means invoiced.
- Uninvoiced rows may be updated in place.
- Changed invoiced rows must keep the historical row and create a new follow-up row directly below it.
- Follow-up rows use `RowKind = CHANGED_COPY`.

### Duplicate rules
A duplicate is defined by identical:

- `Event`
- `Date`
- `Start`
- `End`

Duplicate precedence:

1. duplicates within the same calendar → all rows of that calendar for that duplicate are removed
2. if any specific calendar exists, `Event` rows lose
3. if multiple non-default calendars remain, keep exactly one by configured priority

## Module responsibilities

### `Code.gs`
Orchestration only.
No low-level business logic should accumulate here.

### `Config.gs`
All constants and configuration validation.
New runtime flags belong here unless they are purely transient.

### `Ui.gs`
Only user feedback and UI state restoration.
Do not place business rules here.

### `Scope.gs`
All scope and date-boundary logic.
Any fix related to “before start date”, “future events”, “current timestamp”, or “now” belongs here first.

### `Calendar Service.gs`
Only calendar retrieval and normalization.
This is where event fetch bugs and recurrence handling bugs should be fixed.

### `State Store.gs`
Only hidden sheet read/write behavior.
Keep visible and hidden row alignment stable.

### `Rebuild Engine.gs`
Merging imported events with existing sheet rows.
Invoice preservation logic belongs here.

### `Duplicate Engine.gs`
All duplicate detection and resolution logic.
Do not spread duplicate behavior into other modules.

### `Sheet Writer.gs`
Formatting, color application, and visible sheet writing.
Do not put business decisions here.

### `Table Service.gs`
Only native Google Sheets table object maintenance.
Do not use this module as a business data store.

### `Helper.gs`
Shared pure helpers only.

## Recommended debugging workflow

When behavior is wrong:

1. decide whether the problem is fetch, scope, rebuild, duplicate, state alignment, or writing
2. inspect the responsible module first
3. make the smallest valid fix in that module
4. avoid touching unrelated modules
5. re-check whether the fix changes:
   - scope behavior
   - duplicate behavior
   - invoice preservation
   - row alignment with `_calendar_state`

## Preferred implementation style

- pure functions where possible
- explicit naming
- no fragile hacks
- no magical hidden side effects
- no silent format degradation
- no unnecessary API chatter

## Things to verify after every change

- no future events imported
- no all-day events imported
- recurring past events expand correctly
- rows before the configured start date stay untouched
- uninvoiced changed rows update in place
- invoiced changed rows produce a follow-up row
- duplicate cleanup still follows precedence rules
- visible and hidden sheets remain row-aligned
- table range still matches visible data
- date and time columns remain real spreadsheet values

## Mandatory release gates (must pass before commit)

No agent may commit unless every gate below is explicitly checked in the PR notes:

1. **Scope gate**
   - config scope-affecting fields (`importStartDate`, `calendarNames`, `defaultCalendarName`) correctly trigger sync-token invalidation
   - config save compares against persisted config state, not stale in-memory state
   - import-start boundary is built in spreadsheet timezone (not forced UTC midnight)

2. **Recovery gate**
   - `onOpen()` always creates menu even if config is invalid
   - configuration dialog remains reachable for in-sheet recovery

3. **Data safety gate**
   - invalid config handling does not silently reset unrelated fields to defaults
   - existing user sheets/tabs and user-owned named ranges are never clobbered by config initialization

4. **Security gate**
   - no credentials/tokens are committed
   - any credential-like file must be covered by `.gitignore`
   - CI auth must come from secrets/environment, never hardcoded repo files
   - HTML template model injection must be script-safe escaped (no raw unescaped `<?!=` JSON in `<script>` blocks)

5. **Quality gate**
   - run at least `git diff --check` and confirm a clean working tree after commit
   - run `bash scripts/preflight-review.sh` before commit/PR publication
   - include explicit file/line citations for behavioral changes in final report

## Agent roles and enforcement

- **implementation agent**: writes the smallest correct fix in the responsible module.
- **review agent**: blocks commit if any mandatory release gate is not satisfied.
- **security agent**: blocks commit if credentials handling violates Security gate.
- **release agent**: blocks PR publication unless checks and evidence are documented.
