# ADR-004: Testing Contract

## Status

Accepted.

## Context

Previous checks often verified that code or strings existed, but not that runtime paths behaved correctly. Regressions recurred around config parsing, sheet layout migration, invalid-config recovery, logging, `onOpen()` menu creation, import invariants, and row alignment.

The project requires behavioral tests and release evidence that protect architectural contracts rather than implementation details.

## Decision

Future changes MUST be tested against the applicable architectural invariants, with emphasis on behavior under failure states.

Tests SHOULD be layered: pure configuration validation, sheet-layout migration, recovery behavior, import-domain invariants, and live smoke checks when local scenario tests pass. A pull request MUST include evidence for every layer it can affect.

`docs/test-matrix.md` is the authoritative coverage map for the supported state families and minimum change-to-test mapping. Future test suites MAY implement the matrix with any harness that preserves the required behavioral assertions.

Tests MUST verify negative and recovery paths, not only happy paths. Configuration, recovery, and logging changes MUST prove that invalid configuration, invalid schema metadata, missing sheets, and logging failures behave deterministically.

Release gates MUST match the actual architecture. They MUST require recovery through the current menu/reset path rather than a removed or hypothetical UI.

## Invariants

- Tests MUST protect `ConfigJson` as the only business-configuration source of truth.
- Tests MUST prove invalid `ConfigJson` blocks calendar import.
- Tests MUST prove `onOpen()` creates the menu when config is valid and when config is invalid.
- Tests MUST prove recovery/reset remains reachable without valid configuration.
- Tests MUST prove config validation still works when logging is unavailable or malformed.
- Tests MUST protect layout migration from stale cells, old rows, and user-owned sheet/range clobbering.
- Tests MUST protect import-domain invariants: no future events, no all-day or cancelled events, untouched rows before `importStartDate`, invoice preservation, duplicate precedence, inline `ID` / `EventID` identity, and real spreadsheet date/time values.
- Tests MUST avoid relying solely on source-text presence checks for behavior-critical paths.

## Allowed Future Changes

- The test harness may use Apps Script mocks, local JavaScript tests, live smoke scripts, or other tools that verify the required behaviors.
- Additional required scenarios may be added as contracts evolve.
- Live smoke tests may remain narrower than local tests if they validate the highest-risk integration paths.
- Test implementation may be refactored if coverage of required behaviors is preserved or improved.

## Explicitly Forbidden Changes

- Merging configuration, recovery, logging, or import-domain changes without evidence for affected invariants.
- Replacing behavioral tests with source-presence checks for critical runtime paths.
- Treating a passing happy-path import as sufficient evidence for config, recovery, or logging changes.
- Removing tests for invalid config, invalid schema, unavailable logs, menu creation, reset/recovery reachability, migration safety, or import-domain invariants without an approved replacement.
- Maintaining release gates that reference removed recovery mechanisms or ignore the current menu/reset recovery path.
