# Project recovery plan

## Purpose

This plan assumes the project is starting today with full knowledge of the lessons learned during the last 20 pull requests. Its goal is to maximize stability, predictability, and development velocity while minimizing future rework.

The recovery plan is intentionally prescriptive. The project should stop re-litigating core architecture decisions unless a future requirement explicitly invalidates one of them.

---

## 1. Architectural decisions that are now fixed

### Decision 1: `ConfigJson` is the canonical application configuration

**Decision:** The application configuration is stored as one canonical JSON document under `Config!ConfigJson`.

**Reasoning:** Repeated movement between named ranges, dialog state, key rows, row overrides, legacy aliases, and fallback snapshots caused most of the recent rework. A single canonical payload eliminates precedence ambiguity.

**Expected impact:** Reduces config-related regressions by removing multi-source merge behavior and making validation deterministic.

### Decision 2: `SchemaRegistryJson` is metadata, not runtime configuration

**Decision:** `SchemaRegistryJson` may describe allowed config keys and validation metadata, but it is not business configuration. It must not override or supplement `ConfigJson` values.

**Reasoning:** Treating schema as another source of runtime data would recreate the same ambiguity that `ConfigJson` was meant to eliminate.

**Expected impact:** Keeps strict validation extensible without creating hidden runtime behavior.

### Decision 3: `Validity` is output-only

**Decision:** `Validity` is a diagnostic output cell only. Code may write it, but runtime behavior must never read configuration from it.

**Reasoning:** Diagnostic cells must not become configuration channels. This prevents accidental feedback loops where errors or logs influence runtime behavior.

**Expected impact:** Improves debuggability without increasing configuration complexity.

### Decision 4: `Log` is the only persistent debug-log sheet

**Decision:** Runtime diagnostics that need sheet persistence go to `Log`, not `Config`, managed data tables, or hidden ID/EventID columns.

**Reasoning:** Mixing debug logs into the config sheet polluted layout semantics and caused migration/validation problems.

**Expected impact:** Separates observability from configuration and reduces future layout collisions.

### Decision 5: Recovery must work without valid configuration

**Decision:** Menu creation, reset/recovery entry points, and minimal diagnostics must not depend on successful `ConfigJson` parsing.

**Reasoning:** The user must be able to recover from malformed configuration. Any recovery path that requires valid config is structurally broken.

**Expected impact:** Prevents lockout scenarios and reduces emergency repair PRs.

### Decision 6: Sheet layout migration is allowed; config-content fallback is not

**Decision:** The project may safely migrate spreadsheet layout, but it should not silently fall back from invalid `ConfigJson` to last-valid or default business configuration during normal runtime.

**Reasoning:** The user rejected fallback/versioning for `ConfigJson`, but spreadsheets still need deterministic layout repair when rows/cells move.

**Expected impact:** Preserves strictness while allowing safe upgrades from older sheet layouts.

### Decision 7: `onOpen()` always creates the menu

**Decision:** `onOpen()` must always attempt to create the menu before surfacing config failures.

**Reasoning:** Menu availability is a recovery invariant, not a convenience feature.

**Expected impact:** Prevents invalid config from blocking the user’s path to recovery.

### Decision 8: Calendar data invariants remain higher priority than UX polish

**Decision:** Future events, all-day events, rows before `importStartDate`, invoice-preservation behavior, duplicate precedence, and inline `ID`/`EventID` identity are non-negotiable domain invariants.

**Reasoning:** These are core business correctness rules. Config or UX work must not destabilize them.

**Expected impact:** Keeps future config/recovery work from accidentally damaging import correctness.

### Decision 9: Google Sheets native table and inline identifiers stay

**Decision:** The visible managed tables and hidden first-column `ID`/`EventID` identifiers remain part of the architecture unless explicitly replaced by a dedicated migration plan.

**Reasoning:** These mechanisms preserve user-facing data while keeping row identity inside each managed table. Removing them opportunistically would create data-loss risk.

**Expected impact:** Stabilizes the data model and prevents unrelated refactors from altering persistence semantics.

### Decision 10: Credentials and tokens never move into tracked files or visible sheets

**Decision:** Credentials/tokens must remain in secure Apps Script/property storage or environment/secret systems, never in tracked files or visible sheet cells.

**Reasoning:** Security constraints must remain independent of convenience debugging or migration needs.

**Expected impact:** Prevents accidental credential disclosure while preserving user trust.

---

## 2. Key invariants that must never be violated

### Configuration invariants

1. `ConfigJson` is the only business-configuration source of truth.
2. `SchemaRegistryJson` validates metadata only.
3. `Validity` is output-only.
4. Invalid `ConfigJson` blocks calendar import.
5. Invalid `ConfigJson` must not silently reset unrelated fields to defaults during normal runtime.
6. Unknown config keys are hard errors once strict validation is enabled.
7. Canonical key names are authoritative; legacy aliases must not be reintroduced without an explicit migration-only boundary.

### Recovery invariants

1. `onOpen()` always creates the menu, even when config is invalid.
2. Recovery/reset must be reachable without valid config.
3. Diagnostics must still work when config parsing fails.
4. A malformed `Log` sheet must not prevent menu creation or config recovery.
5. Recovery must not clobber user-owned tabs, named ranges, or calendar data.

### Calendar/import invariants

1. Future events are never written.
2. All-day and cancelled events are never imported.
3. Rows before `CONFIG.importStartDate` remain untouched.
4. Empty `InvoiceNumber` means uninvoiced.
5. Non-empty `InvoiceNumber` means invoiced.
6. Changed invoiced rows preserve the old row and create a `CHANGED_COPY` follow-up row.
7. Managed rows retain inline hidden `ID`/`EventID` values.
8. The native Google Sheets table range matches visible data.
9. Date and time columns remain real spreadsheet values, not text.
10. Duplicate precedence remains stable.

### Security invariants

1. No credentials or tokens in tracked files.
2. No credentials or tokens in visible sheets.
3. CI authentication comes from secrets/environment.
4. Logs redact or avoid credential-like values.
5. HTML model injection remains script-safe escaped.

---

## 3. Components to simplify, rewrite, or redesign

### Immediate simplification: `Config.gs`

**Problem:** `Config.gs` has accumulated constants, validation, sheet layout, schema registry, recovery/reset, property-store debug logic, and log-sheet persistence.

**Recommendation:** Split it into clearer modules:

- `Config.gs`: default config, schema, validation, pure config helpers.
- `Config Sheet.gs`: `Config` sheet layout, read/write, migration, reset.
- `Log.gs`: `Log` sheet creation, append, rotation, formatting.
- `Recovery.gs` or `Ui.gs`: recovery UX, alerts, toasts, menu-safe diagnostics.

**Reasoning:** The last 20 PRs show `Config.gs` as the main churn hotspot. Splitting responsibilities reduces blast radius.

**Expected impact:** Smaller PRs, clearer ownership, fewer cross-module regressions.

### Immediate redesign: config migration path

**Problem:** Layout rewrites have repeatedly left stale values in newly meaningful cells.

**Recommendation:** Create explicit migration functions with preconditions and postconditions:

1. detect current layout,
2. snapshot recoverable values,
3. clear managed layout area,
4. write target layout,
5. restore only approved values,
6. validate final layout,
7. log migration outcome.

**Reasoning:** Spreadsheet cells retain stale values unless explicitly cleared. A key rewrite alone is not a migration.

**Expected impact:** Prevents invalid schema/config states after upgrades.

### Immediate redesign: diagnostics independent of config

**Problem:** Diagnostics have failed when config initialization failed.

**Recommendation:** Provide a minimal `safeLog_()` / `safeValidity_()` path that avoids config parsing entirely.

**Reasoning:** Error reporting must not depend on the subsystem that is failing.

**Expected impact:** Faster diagnosis and less user confusion.

### Simplify: `StatusCell` and progress feedback

**Problem:** `StatusCell` behavior accumulated defaults, aliases, overrides, and validation changes.

**Recommendation:** Keep one canonical config key (`statusCell`) and one fallback for display only if config is invalid. Do not add row-level overrides.

**Reasoning:** Multiple override channels created ambiguity and churn.

**Expected impact:** Reduces config parsing and UI feedback complexity.

### Postpone rewrite: calendar fetch/rebuild/duplicate engines

**Problem:** These modules are business-critical but were not the dominant source of recent instability.

**Recommendation:** Do not refactor them during the config recovery phase except for tests or narrowly scoped bug fixes.

**Reasoning:** Refactoring stable business logic while config/recovery is unstable increases risk without addressing root friction.

**Expected impact:** Keeps recovery work focused and protects import correctness.

---

## 4. Target architecture for the next development phase

### Proposed module boundaries

```text
Code.gs
  - onOpen()
  - updateCalendarSheets()
  - high-level orchestration only

Config.gs
  - DEFAULT_CONFIG
  - schema definition / schema registry builder
  - pure validation
  - pure clone/freeze helpers

Config Sheet.gs
  - ensureConfigSheetLayout_()
  - readConfigJsonFromSheet_()
  - writeConfigJsonToSheet_()
  - migrateConfigSheetLayout_()
  - resetConfigSheet_()

Log.gs
  - ensureLogSheet_()
  - appendLog_()
  - rotateLog_()
  - safeLog_()

Recovery.gs or Ui.gs
  - menu-safe recovery entry points
  - alert/toast policy
  - validity-message policy

Scope.gs
Calendar Service.gs
State Store.gs
Rebuild Engine.gs
Duplicate Engine.gs
Sheet Writer.gs
Table Service.gs
Helper.gs
  - keep current domain responsibilities
```

### Runtime flow

```text
onOpen()
  -> create menu using hardcoded safe fallback labels
  -> validate config through Config Sheet + Config
  -> write Validity / Log if invalid
  -> never block menu creation

updateCalendarSheets()
  -> validate config strictly
  -> block import on invalid config
  -> build scope
  -> fetch calendars
  -> rebuild rows
  -> duplicate cleanup
  -> write managed sheets with inline IDs
  -> update table
  -> save sync tokens

onEdit(e)
  -> if Config sheet changed:
       validate ConfigJson and SchemaRegistryJson
       write Validity
       alert only for direct user edits
```

### Data ownership

| Data | Owner | Notes |
|---|---|---|
| Business config | `Config!ConfigJson` | Canonical source of truth |
| Config metadata | `Config!SchemaRegistryJson` | Validation metadata only |
| Config status | `Config!Validity` | Output-only diagnostics |
| Debug log | `Log` | Structured diagnostics only |
| Imported rows | `Calendar` | User-visible data |
| Row identity/state | Hidden `ID` / `EventID` columns | Inline managed row identity |
| Sync tokens | Apps Script properties | Never visible/tracked |

---

## 5. Technical debt: immediate vs postponed

### Address immediately

#### 1. Document the config contract

**Reasoning:** Without a written contract, every future PR will reinterpret config behavior.

**Expected impact:** Stops architectural drift.

#### 2. Split logging out of `Config.gs`

**Reasoning:** Logging is orthogonal to config validation and has already caused layout pollution.

**Expected impact:** Reduces churn in `Config.gs` and isolates diagnostics.

#### 3. Add scenario tests for config layout migration

**Reasoning:** Migration failures caused direct user-facing breakage.

**Expected impact:** Prevents stale-cell and invalid-schema regressions.

#### 4. Add scenario tests for invalid config recovery

**Reasoning:** The system must remain recoverable when config is broken.

**Expected impact:** Prevents menu lockout and blank diagnostics.

#### 5. Update `agents.md` release gates

**Reasoning:** The gate still references a configuration dialog even though the dialog was removed.

**Expected impact:** Aligns process constraints with actual architecture.

### Postpone

#### 1. Large calendar import refactors

**Reasoning:** Calendar import/rebuild/duplicate logic is business-critical and not the current instability center.

**Expected impact:** Avoids introducing unrelated data regressions.

#### 2. UI polish beyond recovery UX

**Reasoning:** Alerts/toasts/status cells should be stabilized functionally before visual polish.

**Expected impact:** Keeps effort focused on reliability.

#### 3. Advanced schema features

**Reasoning:** Complex schema validation can become another moving target.

**Expected impact:** Preserve strictness without over-engineering.

#### 4. New user-facing configuration editor

**Reasoning:** A new editor should only be built after the config contract is frozen and tested.

**Expected impact:** Avoids recreating the removed dialog’s complexity prematurely.

---

## 6. Pull request strategy to minimize regressions and rework

### PR size and scope rules

1. One architectural boundary per PR.
2. No PR should mix config migration, UI recovery, and calendar import behavior.
3. Every PR touching config must include migration/recovery test evidence.
4. Every PR touching logging must prove config validation still works when logging fails.
5. Every PR touching `onOpen()` must prove menu creation still happens with invalid config.
6. No opportunistic cleanup in business-critical import modules during config/recovery PRs.

### Recommended PR template additions

Each PR should answer:

1. Which invariant does this change rely on?
2. Which invariant could this change break?
3. What sheet layouts were tested?
4. What happens when `ConfigJson` is invalid?
5. What happens when `SchemaRegistryJson` is invalid?
6. What happens when `Log` is missing or malformed?
7. Did `onOpen()` still create the menu?
8. Did scope-affecting config changes invalidate sync tokens?

### Branching strategy

- Use small sequential PRs for recovery architecture.
- Avoid parallel PRs touching `Config.gs`, `Code.gs`, or `Ui.gs` until boundaries are split.
- Allow parallel work only in disjoint modules after contracts are documented.

### Review strategy

- Require an architecture review for any PR touching config source-of-truth semantics.
- Require a recovery review for any PR touching `onOpen()`, `onEdit()`, alerts, toasts, `Validity`, or reset.
- Require a data-safety review for any PR touching `Calendar`, hidden `ID`/`EventID` columns, table ranges, or sync tokens.

---

## 7. Testing strategy that would have prevented observed problems

### Test layer 1: pure validation tests

Test `ConfigJson` parsing and validation without spreadsheet mocks.

Required cases:

- valid default config,
- invalid JSON,
- unknown top-level key,
- missing required key if completeness is required,
- invalid `statusCell`,
- `defaultCalendarName` not in `calendarNames`,
- invalid `importStartDate`,
- invalid headers/state headers.

### Test layer 2: sheet-layout migration tests

Use a lightweight sheet abstraction or Apps Script mock.

Required cases:

- current layout remains unchanged,
- old named-range/key-row layout migrates safely,
- stale row 3 value does not become `SchemaRegistryJson`,
- old log rows inside `Config` do not corrupt config,
- missing schema row is initialized,
- invalid schema row is repaired only during explicit reset/migration,
- user-owned tabs/ranges are not touched.

### Test layer 3: recovery tests

Required cases:

- `onOpen()` creates menu with valid config,
- `onOpen()` creates menu with invalid `ConfigJson`,
- `onOpen()` writes `Validity` on invalid config,
- `onOpen()` logs useful diagnostics if `Log` exists,
- `onOpen()` still completes if `Log` cannot be written,
- `onEdit()` validates only edits on the `Config` sheet,
- reset function works without valid config.

### Test layer 4: import invariants

Required cases:

- future events excluded,
- all-day/cancelled events excluded,
- rows before `importStartDate` untouched,
- uninvoiced changed rows update in place,
- invoiced changed rows create `CHANGED_COPY`,
- duplicate precedence remains stable,
- inline `ID` / `EventID` identity preserved,
- date/time values remain spreadsheet dates/times.

### Test layer 5: live smoke tests

Run a small Apps Script live check only after local scenario tests pass.

Required live checks:

- `onOpen()` menu visible,
- `Config` layout valid,
- `Log` layout valid,
- one update run completes on known fixture sheet,
- invalid config produces visible diagnostics and no import.

### Why this would have prevented prior problems

- Missing helper functions would be caught by runtime-path tests.
- Blank `Validity` would be caught by invalid-config recovery tests.
- Useless `Log` rows would be caught by structured-log assertions.
- Legacy stale-cell migration bugs would be caught by layout migration tests.
- Menu lockout would be caught by `onOpen()` invalid-config tests.

---

## 8. Prioritized roadmap for the next 10 pull requests

### PR 1: Freeze and document the config contract

**Scope:** Add `docs/config-contract.md` and update README/agents gates.

**Reasoning:** No more implementation work should proceed without a stable written contract.

**Expected impact:** Stops requirements drift and gives reviewers a fixed reference.

### PR 2: Add config validation unit/scenario tests

**Scope:** Add tests for valid/invalid `ConfigJson`, unknown keys, invalid schema, and required fields.

**Reasoning:** Strict config behavior has been the largest regression source.

**Expected impact:** Prevents config parser regressions.

### PR 3: Add sheet-layout migration tests

**Scope:** Test old layouts, stale cells, missing schema rows, and reset behavior.

**Reasoning:** Spreadsheet migration bugs caused repeated user-visible failures.

**Expected impact:** Makes layout changes safe.

### PR 4: Extract `Log.gs`

**Scope:** Move log sheet creation/appending/rotation out of `Config.gs`.

**Reasoning:** Logging is not config validation.

**Expected impact:** Reduces `Config.gs` churn and isolates diagnostics.

### PR 5: Extract `Config Sheet.gs`

**Scope:** Move sheet layout/read/write/reset/migration helpers out of `Config.gs`.

**Reasoning:** This creates a clean boundary between pure config validation and spreadsheet I/O.

**Expected impact:** Makes validation easier to test and migration easier to reason about.

### PR 6: Implement safe recovery diagnostics

**Scope:** Add minimal safe diagnostic paths that do not depend on valid config or working log sheet.

**Reasoning:** Recovery must work exactly when config/logging is broken.

**Expected impact:** Prevents blank `Validity` and useless log failures.

### PR 7: Harden `onOpen()` and `onEdit()` with scenario tests

**Scope:** Keep menu creation first and validate all invalid-config branches.

**Reasoning:** `onOpen()` is the user’s recovery gateway.

**Expected impact:** Prevents menu lockout and alert/toast regressions.

### PR 8: Remove remaining legacy ambiguity

**Scope:** Remove or isolate legacy aliases/row override paths that are not part of the frozen contract.

**Reasoning:** Ambiguity keeps causing new precedence bugs.

**Expected impact:** Makes config behavior deterministic.

### PR 9: Add import invariant tests

**Scope:** Add tests around scope, duplicates, invoice preservation, state alignment, and date/time formats.

**Reasoning:** Once config stabilizes, protect business logic before further feature work.

**Expected impact:** Enables safer future development outside config.

### PR 10: Resume feature work only behind contracts

**Scope:** Pick one small user-facing improvement or import feature, with full invariant tests.

**Reasoning:** Feature delivery should resume only after the recovery foundation is stable.

**Expected impact:** Restores development velocity without reintroducing uncontrolled rework.

---

## Success criteria

The recovery phase is successful when:

1. `Config.gs` is no longer the dominant churn hotspot.
2. A malformed `ConfigJson` never prevents menu recovery.
3. A malformed `Log` sheet never prevents config recovery.
4. Config layout migrations are covered by scenario tests.
5. The release gates match the actual architecture.
6. New PRs touch fewer modules and include stronger behavioral evidence.
7. Import-domain invariants remain stable while config/recovery work proceeds.

---

## Final recommendation

The project should treat configuration/recovery as infrastructure, not as incidental UI glue. The next phase should prioritize freezing the contract, splitting responsibilities, and adding scenario tests before adding new features.

The expected result is fewer emergency fixes, less repeated churn in `Config.gs`, and faster development because contributors can make changes against stable, documented boundaries.
