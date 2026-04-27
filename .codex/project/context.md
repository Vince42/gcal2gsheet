# Project Context

## Runtime

- Platform: Google Apps Script (container-bound to Google Sheets)
- Main entrypoint: `updateCalendarSheets()`
- Menu entrypoint: `onOpen()`

## Architecture

- Modular files by responsibility (`Config.gs`, `Ui.gs`, `Scope.gs`, `Calendar Service.gs`, etc.)
- Hidden technical sheet `_calendar_state` keeps row alignment metadata
- Visible business sheet `Calendar` remains user-facing

## Critical constraints

- Keep date/time cells as native spreadsheet values (not text)
- Preserve alignment between visible rows and `_calendar_state`
- Keep duplicate and scope logic centralized in dedicated modules
- Never commit credentials; use secrets management only

## Deployment

- CI workflow pushes via `clasp` to configured Script ID
- Secrets expected in GitHub Actions (`CLASPRC_JSON`)
- Manifest file `appsscript.json` required for `clasp push`

## Operational lessons learned

- `google.script.run` should call public wrapper functions (avoid private/internal naming pitfalls).
- `PropertiesService` access requires correct manifest scopes (`script.storage`, `script.scriptapp`).
- For storage permission edge-cases, config/token persistence should gracefully fallback from `DocumentProperties` to `ScriptProperties`.
- `clasp` CI auth is sensitive to `invalid_grant`/`invalid_rapt`; token lifecycle handling must be documented and repeatable.
- Header/state schema must remain structurally fixed; downstream logic depends on positional columns.
- Managed rows outside active scope should be preserved to avoid data loss during rebuild.
- Scope-affecting config changes (importStartDate/calendarNames) must invalidate sync tokens to force full snapshot backfill.
- Unmanaged future rows must not be dropped solely due to calendar-name heuristics.
