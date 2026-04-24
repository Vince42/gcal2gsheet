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
