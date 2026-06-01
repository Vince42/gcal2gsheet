# Test Matrix

## Purpose

This matrix is the required behavioral coverage map for the ADR contracts. It defines the supported spreadsheet, configuration, migration, recovery, and import states that future pull requests must preserve.

A pull request does not need to run every row in this document unless it can affect every area. It must run, update, or explicitly justify the rows for each state family it can affect. Documentation-only changes may validate this file structurally.

## Coverage Rules

- Tests must assert behavior, not only source-text presence.
- Negative and recovery paths are required whenever configuration, recovery, logging, migration, or import behavior can be affected.
- `ConfigJson` remains the only business-configuration source of truth in every scenario.
- `SchemaRegistryJson`, `Validity`, and `Log` are output or metadata channels only; none may become business-configuration input.
- Credentials, tokens, secrets, and credential-like values must never be committed, written to visible sheets, or logged.
- User-owned sheets, tabs, named ranges, and unrelated cells must not be clobbered by initialization, migration, reset, logging, or import.

## State Families

| Family | Meaning |
|---|---|
| `SS` | Spreadsheet shape and managed-sheet availability states. |
| `CFG` | Configuration and schema validity states. |
| `MIG` | Layout initialization, migration, and reset states. |
| `REC` | Recovery entry point and diagnostic states. |
| `IMP` | Calendar import and business-domain states. |
| `SEC` | Credential and sensitive-data handling states. |

## Baseline Fixture

Unless a row says otherwise, tests start from this baseline:

- `Config` sheet exists with canonical `ConfigJson`, `SchemaRegistryJson`, and `Validity` locations.
- `ConfigJson` is valid and complete.
- `SchemaRegistryJson` is valid metadata.
- `Calendar` sheet exists with the native Google Sheets table object.
- Calendar uses a hidden `ID` column and register sheets use hidden `EventID` columns.
- `Log` exists with the current supported log layout.
- User-owned sheets and named ranges exist and must remain unchanged.
- Calendar fixture data includes past timed events, future timed events, all-day events, cancelled events, recurring past instances, duplicates, invoiced rows, uninvoiced rows, and rows before `importStartDate`.

## 1. Spreadsheet State Matrix

| ID | Supported state | Operation | Required assertions |
|---|---|---|---|
| SS-01 | Fresh spreadsheet with no managed sheets. | Open spreadsheet, initialize/recover, then validate. | Menu is created; required managed sheets can be initialized or recovered; no import runs until valid config exists; user-owned sheets/ranges are not touched. |
| SS-02 | Current valid managed layout. | `onOpen()`, `onEdit()` on config, and one import. | Menu exists; `Validity` reports valid config; import can run; tables, hidden ID/EventID columns, and log remain valid. |
| SS-03 | `Config` sheet missing. | `onOpen()`, recovery/reset. | Menu is created; recovery/reset is reachable; diagnostics are best-effort; no calendar import runs before valid config is restored. |
| SS-04 | `Calendar` sheet missing. | Import with valid config. | Calendar sheet/table can be created or repaired according to implementation policy; the hidden `ID` column is present; user-owned sheets are not touched. |
| SS-05 | Legacy `_calendar_state` sheet exists or is missing during inline-ID migration. | Import with existing visible `Calendar` rows. | Existing state keys are migrated when available; missing legacy state does not silently shift row identity; obsolete state sheets are deleted after migration. |
| SS-06 | `Log` sheet missing. | `onOpen()`, config validation, import error path. | Menu/recovery/config validation still work; `Log` may be recreated; failure to recreate it does not block recovery. |
| SS-07 | `Log` sheet malformed or unwritable. | `onOpen()`, reset/recovery, config validation. | Logging failure is contained; menu/recovery/config validation still work; validation failures remain visible through non-log diagnostics where possible. |
| SS-08 | `Calendar` contains rows before `importStartDate`. | Import. | Historical pre-start rows remain untouched and are not re-imported, updated, deleted, or state-mutated. |
| SS-09 | Visible `Calendar` rows have missing or stale inline IDs. | Import or preflight validation. | Import is blocked or inline IDs are repaired deterministically before writes; no row identity is silently shifted. |
| SS-10 | User-owned sheets, tabs, named ranges, and extra cells exist. | Initialization, migration, reset, and import. | Only managed ranges/sheets are changed; user-owned spreadsheet state is preserved. |
| SS-11 | Stale values exist in old managed config rows/cells. | Migration/reset. | Stale cells do not become `ConfigJson`, `SchemaRegistryJson`, runtime config, or log input; final layout validates. |
| SS-12 | Managed sheet write fails for diagnostics. | `onOpen()` with invalid config. | Menu creation still completes; recovery remains reachable; failure is surfaced by the safest available channel. |

## 2. Configuration State Matrix

| ID | Supported state | Operation | Required assertions |
|---|---|---|---|
| CFG-01 | Valid complete canonical `ConfigJson` and valid `SchemaRegistryJson`. | Validate and import. | Config validates deterministically; import may run; `Validity` is success output only. |
| CFG-02 | `ConfigJson` is malformed JSON. | `onOpen()`, `onEdit()`, import. | Menu/recovery remain available; `Validity` records failure when writable; import is blocked; no fallback config is used. |
| CFG-03 | `ConfigJson` is syntactically valid but missing required keys. | Validate and import. | Validation fails; import is blocked; unrelated fields are not reset to defaults during normal runtime. |
| CFG-04 | `ConfigJson` has invalid types or values. | Validate and import. | Validation fails with actionable diagnostics; import is blocked; invalid values are not coerced into runtime config except by documented validation rules. |
| CFG-05 | `ConfigJson` contains unknown keys after strict validation. | Validate. | Unknown keys are hard errors; no import runs; no unknown key is ignored as runtime config. |
| CFG-06 | `ConfigJson` uses legacy aliases instead of canonical keys. | Validate or migrate within explicit migration boundary. | Normal runtime rejects aliases; migration may convert only documented aliases inside migration-only code; aliases are not runtime config sources. |
| CFG-07 | `SchemaRegistryJson` is malformed JSON. | Validate config and open spreadsheet. | Schema failure is reported; schema does not become runtime config; menu/recovery remain available. |
| CFG-08 | `SchemaRegistryJson` is valid JSON but invalid metadata shape. | Validate config. | Validation fails or falls back only to built-in schema policy if explicitly documented; schema values do not supplement business config. |
| CFG-09 | `SchemaRegistryJson` allows or omits keys unexpectedly. | Validate config. | Canonical implementation schema remains authoritative; unexpected metadata does not permit invalid business config. |
| CFG-10 | `Validity` contains stale success or failure text. | Validate and import. | Runtime does not read `Validity`; current validation rewrites it when possible; stale text does not affect import decisions. |
| CFG-11 | Scope-affecting keys change: `importStartDate`, `calendarNames`, or `defaultCalendarName`. | Save config, then import. | Persisted config comparison detects the change; sync tokens are invalidated; import-start boundary uses spreadsheet timezone. |
| CFG-12 | Non-scope-affecting keys change. | Save config, then import. | Config is persisted and validated; sync tokens are not invalidated solely because of unrelated changes. |
| CFG-13 | `statusCell` is missing or invalid. | Validate and surface UI feedback. | Validation follows schema policy; display fallback may be used only for diagnostics/recovery and never as business config. |
| CFG-14 | Credential-like values appear in config fields. | Validate, log, and report errors. | Secrets are rejected where inappropriate or redacted in diagnostics; credentials are not logged or written to visible runtime outputs. |

## 3. Migration State Matrix

| ID | Supported state | Operation | Required assertions |
|---|---|---|---|
| MIG-01 | Current layout already matches target layout. | Run migration/ensure-layout. | No semantic changes; canonical values preserved; operation is idempotent. |
| MIG-02 | Fresh spreadsheet requires initial managed layout. | Initialize layout. | Required managed locations are created; no import runs until valid config is present; user-owned state is preserved. |
| MIG-03 | Old named-range or key-row config layout exists. | Explicit migration. | Known canonical values are copied into `ConfigJson`; managed layout area is cleaned; legacy sources are not used after migration. |
| MIG-04 | Old config rows contain stale values in cells now used by schema or validity. | Explicit migration. | Stale values are cleared or ignored; stale row values do not become `SchemaRegistryJson`, `Validity`, or runtime config. |
| MIG-05 | Old debug/log rows exist inside `Config`. | Explicit migration. | Old log rows do not corrupt config/schema; persistent future diagnostics go only to `Log`. |
| MIG-06 | `SchemaRegistryJson` row missing. | Explicit migration/reset. | Schema metadata location is initialized or repaired; business config remains from canonical `ConfigJson`. |
| MIG-07 | Existing `ConfigJson` is invalid before migration. | Explicit migration. | Layout may be repaired, but invalid config content remains invalid; no defaults or last-valid snapshots silently replace it. |
| MIG-08 | Extra legacy rows exist below current managed layout. | Explicit migration. | Managed stale rows are cleared according to documented bounds; user-owned rows outside managed bounds are preserved. |
| MIG-09 | Managed named ranges are missing, invalid, or point to old cells. | Explicit migration. | Managed ranges may be repaired; user-owned named ranges are not deleted or rewritten. |
| MIG-10 | Migration is interrupted or a write fails. | Retry migration/recovery. | Partial migration does not create a valid-looking but semantically corrupt layout; retry is deterministic; user data is preserved. |
| MIG-11 | Reset is invoked with invalid config and malformed schema. | Reset/recovery. | Reset is reachable without valid config/schema; exact reset effects are deterministic; no unrelated fields or user-owned state are clobbered. |
| MIG-12 | Migration completes. | Post-migration validate. | Final layout validates; migration outcome is logged if possible; logging failure does not invalidate migration or block recovery. |

## 4. Recovery State Matrix

| ID | Trigger | State | Required assertions |
|---|---|---|---|
| REC-01 | `onOpen()` | Valid config, valid schema, writable log. | Menu is created first; no blocking recovery alert is required; `Validity` may show success. |
| REC-02 | `onOpen()` | Malformed `ConfigJson`. | Menu is created; recovery/reset is reachable; `Validity` failure is written when possible; import remains blocked. |
| REC-03 | `onOpen()` | Missing required config keys or invalid values. | Menu is created; actionable diagnostics are produced; import remains blocked; no fallback config is used. |
| REC-04 | `onOpen()` | Malformed or invalid `SchemaRegistryJson`. | Menu is created; schema diagnostic is surfaced; schema does not become runtime config. |
| REC-05 | `onOpen()` | `Config` sheet missing. | Menu is created; reset/recovery can recreate or repair required layout; import remains blocked until valid config exists. |
| REC-06 | `onOpen()` | `Log` missing, malformed, or unwritable. | Menu is created; recovery remains available; logging failure does not hide config diagnostics. |
| REC-07 | `onOpen()` | `Validity` cannot be written. | Menu is created; recovery remains available; safest alternate diagnostic channel is used. |
| REC-08 | `onEdit()` | Edit on `Config` sheet with valid config. | Only relevant config edits trigger validation; `Validity` becomes success when writable. |
| REC-09 | `onEdit()` | Edit on `Config` sheet with invalid config. | Validation failure is surfaced; import remains blocked; unrelated fields are not reset. |
| REC-10 | `onEdit()` | Edit on non-`Config` sheet. | Config validation/recovery side effects are not triggered except by explicitly documented paths. |
| REC-11 | `updateCalendarSheets()` | Invalid config or schema. | Import is blocked before calendar fetch/write; diagnostics are best-effort; no fallback config is used. |
| REC-12 | Reset/recovery entry point | Invalid config, invalid schema, and broken log. | Reset/recovery still runs; exact managed layout effects are deterministic; user-owned state is preserved. |
| REC-13 | Any recovery path | Calendar API unavailable. | Recovery does not require calendar access; menu/reset/config diagnostics remain available. |

## 5. Import State Matrix

| ID | Supported import state | Operation | Required assertions |
|---|---|---|---|
| IMP-01 | Past timed events inside scope. | Import. | Events are written with real spreadsheet date/time values and hidden inline `ID` values. |
| IMP-02 | Future timed events. | Import. | Future events are never written. |
| IMP-03 | All-day events. | Import. | All-day events are never imported. |
| IMP-04 | Cancelled events. | Import. | Cancelled events are never imported. |
| IMP-05 | Recurring events with past and future instances. | Import. | Only past timed instances inside active scope survive filtering. |
| IMP-06 | Existing rows before `importStartDate`. | Import. | Rows remain untouched, including their hidden inline `ID` values. |
| IMP-07 | Existing uninvoiced row changed in source calendar. | Import. | Row updates in place; inline `ID` remains stable. |
| IMP-08 | Existing invoiced row changed in source calendar. | Import. | Historical row is preserved; directly following row is created with `RowKind = CHANGED_COPY`. |
| IMP-09 | Duplicate rows within the same calendar. | Import/rebuild duplicate cleanup. | All rows for that same-calendar duplicate are removed according to duplicate precedence. |
| IMP-10 | Duplicate between default `Event` calendar and a specific configured calendar. | Import/rebuild duplicate cleanup. | Specific calendar row wins; `Event` row loses. |
| IMP-11 | Duplicate across multiple non-default calendars. | Import/rebuild duplicate cleanup. | Exactly one row remains according to configured calendar priority. |
| IMP-12 | Empty `InvoiceNumber`. | Import changed event. | Row is treated as uninvoiced and may update in place. |
| IMP-13 | Non-empty `InvoiceNumber`. | Import changed event. | Row is treated as invoiced and must not be overwritten by changed source values. |
| IMP-14 | Calendar rows contain inline `ID` values before import. | Import. | Inline `ID` values remain present and travel with managed rows after write. |
| IMP-15 | Native table range exists before import. | Import. | Table range matches visible data after write. |
| IMP-16 | Date/time cells exist as spreadsheet values before import. | Import. | Date/time columns remain spreadsheet date/time values, not text. |
| IMP-17 | Calendar fixture has no changes since previous import. | Import twice. | Second import is idempotent except for documented diagnostics/log timestamps. |
| IMP-18 | Scope-affecting config changed since last import. | Save config and import. | Sync tokens are invalidated and import scope is rebuilt from persisted config. |
| IMP-19 | Calendar API fetch fails after valid config. | Import. | Failure is reported; partial writes do not corrupt inline row identity; recovery remains available. |
| IMP-20 | Log write fails during import diagnostics. | Import error path. | Logging failure does not mask import failure or corrupt config/recovery state. |

## 6. Security State Matrix

| ID | Supported state | Operation | Required assertions |
|---|---|---|---|
| SEC-01 | Repository contains no credential files. | Preflight/static check. | `.clasprc.json`, `.clasp.json`, tokens, and credential-like files are untracked or ignored. |
| SEC-02 | CI or live smoke authentication is required. | Run test setup. | Authentication comes from secrets/environment, never hardcoded repo files. |
| SEC-03 | Error contains credential-like text. | Log and surface diagnostics. | Diagnostic outputs redact or avoid the sensitive value. |
| SEC-04 | User enters a token-like value into visible config or calendar cells. | Validate/log. | Value is not propagated into logs; credential storage policy is enforced by validation or documentation. |
| SEC-05 | HTML template receives model data. | Static check or UI test. | Model injection is script-safe escaped; raw unescaped JSON injection is not introduced. |

## 7. Minimum Change-to-Test Mapping

| Change area | Required matrix rows |
|---|---|
| Pure config schema or validation | CFG-01 through CFG-10, CFG-13, CFG-14, REC-02 through REC-04, REC-11. |
| Scope-affecting config save behavior | CFG-11, CFG-12, IMP-06, IMP-18. |
| Config sheet layout, migration, or reset | SS-01, SS-03, SS-10, SS-11, MIG-01 through MIG-12, REC-05, REC-12. |
| `onOpen()`, menu, alerts, toasts, or recovery UX | REC-01 through REC-07, REC-12, REC-13, SS-03, SS-06, SS-07, SS-12. |
| `onEdit()` behavior | REC-08 through REC-10, CFG-01 through CFG-05. |
| Logging implementation or log layout | SS-06, SS-07, REC-06, IMP-20, SEC-03, plus affected config validation rows proving logging failure is non-blocking. |
| Calendar fetch, scope, or event normalization | IMP-01 through IMP-06, IMP-18, IMP-19. |
| Rebuild, invoice preservation, or duplicate logic | IMP-07 through IMP-14, IMP-17. |
| Sheet writing, table maintenance, formatting, or state storage | SS-04, SS-05, SS-08, SS-09, IMP-14 through IMP-16, IMP-19. |
| Security, credentials, CI auth, or HTML template data | SEC-01 through SEC-05. |
| Documentation-only contract changes | Structural validation of changed docs, plus review that the matrix remains consistent with ADR-001 through ADR-004. |

## 8. Live Smoke Subset

After applicable local scenario tests pass, live smoke coverage should include at least:

| ID | Live check | Required assertions |
|---|---|---|
| LIVE-01 | Open known fixture spreadsheet. | Menu is visible and recovery entry point is reachable. |
| LIVE-02 | Validate fixture `Config`. | `ConfigJson`, schema metadata, and `Validity` locations behave according to ADR-001. |
| LIVE-03 | Validate fixture `Log`. | Log sheet accepts structured diagnostics or fails without blocking recovery. |
| LIVE-04 | Run one import on fixture data. | Import-domain invariants in IMP-01 through IMP-16 hold for fixture coverage. |
| LIVE-05 | Temporarily make `ConfigJson` invalid, then open/recover. | Menu remains visible; diagnostics are visible; import is blocked; original valid config is restored from the test harness, not from runtime fallback. |
