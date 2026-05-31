# Post-mortem review: last 20 PRs and current in-flight PR

## Scope and evidence base

This review treats the current unmerged/latest commit (`78e01b4`, `Add schema registry, strict config validation, log sheet, and onEdit validation UX`) as the latest in-flight PR, and reviews it alongside the latest 20 merged PRs found in git history.

Commands used to produce the review:

```bash
git log --oneline --decorate -n 40
git log --merges --oneline --grep='Merge pull request' -n 25
git log --first-parent --oneline --decorate -n 60
git log --reverse --no-merges --pretty='%h %s' 698a059^..HEAD
git diff --shortstat <parent>..<merge>
git diff --numstat <parent>..<merge>
git show --shortstat --oneline HEAD --
rg -n "Config|Log|SchemaRegistry|StatusCell|named range|configuration|Validity" README.md Config.gs Code.gs Ui.gs agents.md
git status --short
```

## Executive summary

The project’s development friction is not mainly caused by individual bugs. It is caused by repeated architectural oscillation around one unstable subsystem: configuration and recovery UX.

Across the reviewed PRs, the same core concepts changed repeatedly:

- where configuration lives,
- how configuration is discovered,
- whether config values are read from JSON, named ranges, key/value rows, or fallback rows,
- whether invalid config should block, warn, toast, write status, or show alerts,
- where debug logs belong,
- whether recovery is through a dialog, menu, sheet state, or reset helper,
- how much backwards compatibility should be preserved.

The strongest evidence is file churn: across the last 20 merged PRs, `Config.gs` was touched in 17 of 20 PRs with 575 insertions and 472 deletions; `Ui.gs` and `Code.gs` were each touched in 10 of 20 PRs. That means the system’s configuration boundary was not stable enough to support incremental changes.

Current project rules correctly say the codebase should remain modular and that `Config.gs` owns constants/config validation, `Ui.gs` owns feedback/UI state, and `Code.gs` should remain orchestration-only. But the recent history shows these boundaries repeatedly blurred around validation, recovery, logging, and menu behavior.

---

## 1. Architectural decisions that changed multiple times

### 1.1 Configuration source of truth changed repeatedly

The most unstable architectural decision was: what is the source of truth for configuration?

Observed sequence from PR / commit history:

1. Named-range-based configuration and dialog flow.
   - PRs around #30–#33 focused on named range validation, cleanup, diagnostics, and save behavior.
   - Relevant commits included:
     - `0de4173 Clean invalid managed named ranges and add config save diagnostics`
     - `b9f7662 Sanitize all invalid named ranges on load with user notification`
     - `a8893d8 Restrict cleanup to managed ranges and harden onOpen/save diagnostics`
     - `8d248e7 Remove blocking startup/save alerts and keep non-blocking diagnostics`

2. Dialog removal and key-row config.
   - PR #34 / commit `ec2e436 Load config by key rows and remove named-range dependencies`.
   - This was a major architectural swing: it removed `ConfigDialog.html` and moved behavior into key/value rows.

3. Status-cell row overrides and hybrid fallback.
   - PRs #39–#41 repeatedly changed `StatusCell` behavior:
     - `fcdab8b Validate StatusCell A1 format and strict setting lookup`
     - `f256770 Derive default StatusCell from header width and add smoke test`
     - `04c64e9 Honor Pascal-case StatusCell in config I/O and validation`
     - `6b41e55 Honor StatusCell row override over ConfigJson fallback`

4. Config sheet by exact name.
   - PR #43 / commit `ff24783 Accept Config sheet by name and improve config diagnostics`.

5. Strict `ConfigJson` + `SchemaRegistryJson`.
   - Current HEAD `78e01b4` switched to strict parsing of `ConfigJson` plus validation against `SchemaRegistryJson`.

6. Reset helper scanning all cells for JSON.
   - Current code also contains reset/recovery behavior that scans the sheet for JSON and rebuilds the expected layout.

### Architectural conclusion

The config subsystem has not had a stable contract. It moved from:

```text
named ranges → dialog → key/value rows → row overrides → exact sheet name → strict JSON/schema → reset/search recovery
```

That is a large amount of architectural movement for one subsystem.

### 1.2 Recovery UX changed repeatedly

The recovery path also changed multiple times:

- `onOpen()` must always add the custom menu even when config is invalid, so users keep recovery access.
- Later release gates reinforced this: recovery requires `onOpen()` to always create the menu and keep recovery reachable.
- Recent PRs changed whether invalid config should:
  - show blocking alerts,
  - avoid blocking alerts,
  - write status messages,
  - use toast,
  - use `StatusCell`,
  - write to `Validity`,
  - or create reset helpers.

### Architectural conclusion

Recovery behavior became a cross-cutting concern across `Code.gs`, `Config.gs`, and `Ui.gs`, rather than a clean, documented state machine.

### 1.3 Debug logging location changed multiple times

Debug logging moved from ad-hoc/breadcrumb style to config-sheet persistence and then to a dedicated `Log` sheet.

Evidence from commit history:

- `479fa48 Bump dialog revision and add storage debug breadcrumbs`
- `712f793 Persist storage debug breadcrumbs in config sheet`
- `d46f4e6 Fix config debug logging persistence and sheet reuse`
- current HEAD moved to structured `Log` sheet.

### Architectural conclusion

The logging system was initially optimized for immediate troubleshooting, but it became entangled with config layout. That created later migration and parsing hazards.

### 1.4 Validation strictness changed repeatedly

Validation moved through several modes:

- permissive/fallback behavior,
- last-valid JSON recovery,
- row overrides,
- legacy aliases such as Pascal-case `StatusCell`,
- strict canonical `statusCell`,
- schema registry validation,
- reset recovery.

Strict unknown-key validation is sound in principle, but the project has also mixed strict validation with automatic default merging. That creates ambiguity unless the intended behavior is explicitly documented.

---

## 2. Files/modules modified disproportionately often

Aggregated across the last 20 merged PRs, based on `git diff --numstat <parent>..<merge>`:

| File | PRs touched | Insertions | Deletions | Total churn |
|---|---:|---:|---:|---:|
| `Config.gs` | 17 / 20 | 575 | 472 | 1047 |
| `Ui.gs` | 10 / 20 | 269 | 193 | 462 |
| `ConfigDialog.html` | 1 / 20 | 0 | 246 | 246 |
| `Code.gs` | 10 / 20 | 86 | 62 | 148 |
| `README.md` | 1 / 20 | 32 | 0 | 32 |
| `scripts/live-sheet-check.sh` | 1 / 20 | 30 | 0 | 30 |
| `scripts/smoke-test.sh` | 2 / 20 | 20 | 0 | 20 |
| `State Store.gs` | 1 / 20 | 9 | 0 | 9 |
| `.github/workflows/apps-script-push.yml` | 1 / 20 | 2 | 2 | 4 |
| `scripts/preflight-review.sh` | 1 / 20 | 3 | 0 | 3 |

### Interpretation

#### `Config.gs` is the instability hotspot

`Config.gs` owns constants and configuration validation according to project docs. But recent history shows it also accumulated:

- config sheet layout migration,
- schema registry validation,
- storage/property handling,
- debug logging,
- log sheet creation,
- reset/recovery helpers,
- config cell scanning,
- validity message writing.

That is more than constants and configuration validation.

#### `Code.gs` and `Ui.gs` were repeatedly pulled into config concerns

`Code.gs` is supposed to be orchestration only. But it now participates in config validation UX on open/edit, validity reporting, storage debug logging, and menu fallback.

`Ui.gs` is supposed to own UI feedback and state restoration only. But repeated PRs around status cells and progress feedback caused it to participate in configuration recovery semantics.

---

## 3. Requirements that appear unstable, ambiguous, or evolving

### 3.1 “ConfigJson is source of truth” was not consistently defined

The latest requirement says `ConfigJson` should be the only source of truth. Current implementation mostly follows that by reading `ConfigJson` and `SchemaRegistryJson`.

However, current behavior also includes:

- legacy layout normalization,
- scanning all cells for recoverable JSON,
- initializing schema if invalid,
- merging config with defaults before validation.

Those may be practical recovery mechanisms, but they complicate the statement “only source of truth.”

### 3.2 Config sheet layout was underspecified

The README only describes `Config.gs` as constants and configuration validation. It does not fully document the actual `Config` sheet contract:

- exact sheet name,
- required rows,
- row keys,
- whether row order matters,
- whether old keys may remain,
- expected monospace formatting,
- whether `SchemaRegistryJson` is user-editable,
- how reset works,
- when validation runs.

### 3.3 Logging requirements evolved from troubleshooting to architecture

Debug logging began as a narrow troubleshooting feature. It later became:

- persisted,
- part of config sheet,
- then moved to `Log`,
- then structured into five columns,
- then part of reset behavior.

The absence of an early logging contract caused the log format and location to become another moving target.

### 3.4 Recovery behavior was ambiguous

There were competing requirements:

- invalid config must not block menu creation,
- invalid config should show visible warnings,
- startup should avoid blocking alerts,
- later startup should show alert,
- recovery dialog should remain reachable,
- later dialog was removed,
- then reset helper was added.

The release gate still refers to a reachable configuration dialog, even though the dialog was removed in PR #34. That is a documentation/architecture mismatch.

---

## 4. Recurring implementation mistakes and regression patterns

### 4.1 Partial migrations broke existing sheets

Examples:

- Schema registry row was introduced but legacy row values remained in column B, causing old values to be parsed as schema JSON.
- Later a reset helper was added to search cells for JSON and rebuild layout.

Pattern:

```text
Layout changes were implemented as local key rewrites, not as explicit migrations with pre/post conditions.
```

Even if `ConfigJson` itself should not have fallback/versioning, sheet-layout migration still needs a deterministic upgrade path. The lack of one caused rework.

### 4.2 Undefined helper/function regressions

A reported runtime error was:

```text
readConfigSettingCell_ is not defined
```

That indicates code was introduced referencing a helper that did not exist.

Pattern:

```text
Static smoke tests searched for symbols but did not execute the Apps Script paths that would catch missing runtime references.
```

### 4.3 Diagnostics often failed exactly when needed

The user reported that `Validity` was empty and `Log` was not useful during failure.

Pattern:

```text
Error-reporting code depended on the same config initialization path that was failing.
```

Diagnostics must be designed to work when config is malformed, missing, or partially migrated.

### 4.4 Cross-module leakage

Examples:

- `Code.gs` handles `onOpen()` config validation, menu fallback, logging, toast, and alert behavior.
- `Config.gs` owns sheet migration, log sheet creation, reset, and debug persistence.

Pattern:

```text
The quickest fix was often placed wherever the failure appeared, not necessarily where the responsibility belonged.
```

This conflicts with the repo’s own modularity rules.

### 4.5 Static tests became assertion-by-string, not behavioral tests

`scripts/smoke-test.sh` checks for presence of function names and strings. It does not simulate:

- legacy Config layout,
- invalid `SchemaRegistryJson`,
- invalid `ConfigJson`,
- missing `Log`,
- `onOpen()` with invalid config,
- `onEdit()` behavior,
- reset recovery.

That explains why regressions like missing helpers, blank `Validity`, or broken schema-row migration survived.

---

## 5. Cases where local optimization created a larger system problem

### 5.1 Logging into `Config` was locally convenient but architecturally harmful

Early debug persistence into the config sheet made troubleshooting easy, but it polluted the config surface and made row layout harder to reason about. Later, logs had to be moved to `Log`.

### 5.2 Cleaning invalid named ranges was locally protective but globally risky

PRs #30–#33 show repeated work around named range cleanup:

- clean invalid managed ranges,
- sanitize all invalid ranges,
- restrict cleanup to managed ranges,
- remove blocking diagnostics.

This suggests a local optimization — automatically cleaning invalid ranges — risked touching user-owned spreadsheet state.

### 5.3 Removing the configuration dialog simplified code but weakened recovery

PR #34 removed `ConfigDialog.html` and deleted 666 lines across four files. That reduced code size, but it also conflicted with the documented recovery gate that still refers to a reachable configuration dialog.

Local simplification created a system-level mismatch: recovery requirements remained, but the recovery mechanism changed.

### 5.4 Supporting `StatusCell` aliases preserved compatibility but expanded ambiguity

PRs #39–#41 repeatedly handled status cell derivation, A1 validation, Pascal-case `StatusCell`, and row override precedence.

That compatibility work helped users temporarily, but it created multiple ways for a status target to be derived:

- default from header width,
- JSON `statusCell`,
- legacy `StatusCell`,
- row override,
- fallback behavior.

Later strict config work tried to collapse this back to one canonical key.

### 5.5 Strict schema registry improved correctness but introduced migration fragility

Rejecting unknown keys through a schema registry is sound in principle. But introducing `SchemaRegistryJson` into an existing sheet without fully isolating old row values caused validation failures. This is a classic case where a correctness improvement lacked a safe migration envelope.

---

## 6. Areas where constraints, assumptions, or invariants were insufficiently documented

### 6.1 Config sheet contract

Needs explicit documentation:

- sheet name must be exactly `Config`,
- required rows,
- whether row order matters,
- whether old rows may exist below current layout,
- valid `SchemaRegistryJson` shape,
- whether `SchemaRegistryJson` is user-editable,
- whether missing config keys are allowed,
- whether defaults are merged or strict completeness is required,
- reset function name and expected effect.

### 6.2 Recovery state machine

The intended behavior should be documented as a matrix:

| Trigger | Valid config | Invalid config | Missing Config sheet | Broken Log sheet |
|---|---|---|---|---|
| `onOpen()` | menu + no alert | menu + alert + Validity | menu? reset? | console fallback |
| `onEdit()` | Validity = VALID | alert + Validity | n/a | console fallback |
| `updateCalendarSheets()` | run | block | block | run? |

Currently that behavior is spread across `Code.gs` and `Config.gs`.

### 6.3 Migration policy

The user rejected fallback/versioning for `ConfigJson`, but sheet layout migration is still necessary. This distinction was not documented:

- No fallback for config content is one rule.
- Safe migration of sheet layout is a separate rule.

Those got conflated.

### 6.4 Ownership boundaries

`Config.gs` currently includes logging and reset behavior. That may be expedient, but the documented responsibility for `Config.gs` is only constants/config validation.

A more stable architecture would probably split:

- `Config.gs`: schema, parsing, validation
- `Config Sheet.gs`: layout/read/write/reset
- `Log.gs`: log sheet creation/writes
- `Recovery.gs` or `Ui.gs`: alerts/toasts/recovery UX

### 6.5 Tests required by release gates are too weak

The release gates require `git diff --check` and `bash scripts/preflight-review.sh`. But the recurring failures are behavioral and migration-related. A string-based smoke test cannot validate those.

---

## 7. Estimated effort distribution

This is an estimate based on PR titles, commit subjects, changed files, and churn. It is not time-tracking data.

| Category | Estimated share | Evidence |
|---|---:|---|
| New functionality | 15–20% | `Log` sheet, live sheet check workflow, schema registry, `onEdit`, reset helper, Node 24 workflow update |
| Refactoring / architectural reshaping | 30–35% | named ranges → key rows, dialog removal, config layout rewrites, schema registry introduction, logging relocation |
| Bug fixes / regression fixes | 35–40% | repeated PRs titled `fix-code-issue`, `fix-debugging-information-logging`, `fix-permission_denied`, `restore-setprogress_`, config diagnostics fixes |
| Reverting / undoing / repairing previous changes | 15–20% | removal of blocking alerts, removing dialog after adding/maintaining it, row override changes, correcting broad named-range cleanup, reset helper after broken migration |

## Important observation

The line-churn profile supports this estimate:

- `Config.gs`: 1047 changed lines across 17/20 PRs.
- `Ui.gs`: 462 changed lines across 10/20 PRs.
- `Code.gs`: 148 changed lines across 10/20 PRs.
- `ConfigDialog.html`: 246 deleted lines in one PR.

This is not a profile of steady feature delivery. It is a profile of repeated stabilization around the same subsystem.

---

## 8. Most likely root causes of development friction

### Root cause 1: No stable configuration contract

The project lacked a single, documented config contract early enough. The contract kept changing between named ranges, dialog state, key rows, JSON, row overrides, and schema registry.

Recommendation:

- Write a `docs/config-contract.md` or README section with:
  - exact sheet layout,
  - exact JSON schema,
  - validation timing,
  - migration/reset behavior,
  - examples of valid/invalid config,
  - ownership boundaries.

### Root cause 2: Recovery behavior was not modeled as a state machine

The system needs to work when config is invalid. That means recovery code must be independent of config parsing. Current code partly addresses this, but the recurring failures show it was added reactively.

Recommendation:

- Define recovery states:
  - `CONFIG_OK`
  - `CONFIG_INVALID_JSON`
  - `SCHEMA_INVALID`
  - `CONFIG_SCHEMA_MISMATCH`
  - `CONFIG_SHEET_MISSING`
  - `LOG_UNAVAILABLE`
- Ensure each state has deterministic UI/log/validity behavior.

### Root cause 3: Migrations were implicit and under-tested

Layout rewrites happened by updating row labels and preserving some values. That is fragile in spreadsheets because old values remain in cells and row meanings change.

Recommendation:

- Even without `ConfigJson` versioning, introduce explicit sheet layout migration functions:
  - detect old layout,
  - copy known values,
  - clear stale rows/cells,
  - write new layout,
  - validate final state,
  - log migration result.

This is not the same as config fallback/versioning.

### Root cause 4: `Config.gs` became a god module

`Config.gs` now owns validation, sheet layout, schema registry, properties-store debug, log sheet creation, reset, and recovery helpers.

Recommendation:

- Split responsibilities:
  - `Config.gs`: config object, schema, validation only
  - `Config Sheet.gs`: reading/writing/resetting the `Config` sheet
  - `Log.gs`: log sheet creation/writes
  - `Recovery.gs` or `Ui.gs`: alerts/toasts/recovery UX

### Root cause 5: Tests verified code presence, not behavior

The smoke tests caught some accidental removals, but they did not catch broken runtime paths.

Recommendation:

Add a local Apps Script simulation test harness or lightweight JS unit tests for:

- old config layout → new layout migration,
- invalid `SchemaRegistryJson`,
- invalid `ConfigJson`,
- unknown key rejection,
- missing `Config` sheet behavior,
- `onOpen()` always creates menu,
- `Validity` is written on failure,
- `Log` rows are structured.

### Root cause 6: Too many PRs optimized for immediate symptoms

Many PRs appear to fix the current visible problem:

- debug line not visible,
- menu missing,
- status cell wrong,
- named range invalid,
- permission denied,
- alert too blocking,
- config sheet not found.

Each local fix was reasonable in isolation, but collectively they changed architectural assumptions.

Recommendation:

Before touching config/recovery again, create a short architecture decision record (ADR):

- `ADR-001 Config source of truth`
- `ADR-002 Config sheet layout and migration`
- `ADR-003 Recovery UX and invalid-config behavior`
- `ADR-004 Logging contract`

---

## Specific recommendations for stabilizing the project

### Immediate

1. Freeze the config contract.
   - `ConfigJson` is canonical.
   - `SchemaRegistryJson` is metadata.
   - `Validity` is output-only.
   - No row overrides.
   - No legacy aliases.

2. Document the reset function.
   - Name, trigger, exact effects, whether it clears logs, whether it preserves old JSON.

3. Add behavioral tests for config migration.
   - Especially legacy rows and stale column B values.

4. Move logging out of `Config.gs`.
   - A dedicated `Log.gs` would reduce future config churn.

5. Update release gates.
   - Current gates mention a configuration dialog even though the dialog was removed.
   - Replace that with “configuration recovery remains reachable through menu/reset function.”

### Medium term

1. Create an architectural boundary around recovery.
   - Recovery must not depend on valid config.

2. Create a compatibility policy.
   - Decide how many old layouts are supported and for how long.
   - Do not let every PR invent a new fallback.

3. Replace string smoke tests with scenario tests.
   - The current failures are scenario failures, not syntax failures.

---

## Final judgment

The project’s slow progress and rework are primarily caused by architectural instability around configuration ownership, validation, recovery, and diagnostics.

The system repeatedly changed the same boundary:

```text
How does the app know its configuration, validate it, recover from invalid state, and tell the user what happened?
```

Until that boundary is documented, tested, and separated into clearer modules, future PRs will likely continue the same pattern: a local fix will solve today’s visible spreadsheet problem while creating tomorrow’s config migration, logging, or recovery regression.
