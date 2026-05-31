# ADR-002: Recovery Contract

## Status

Accepted.

## Context

Recovery behavior previously evolved reactively around invalid configuration, missing sheets, malformed schema metadata, logging failures, menu lockout, and removed UI paths. A recovery path that depends on valid configuration fails exactly when users need it most.

The project requires a recovery contract that is independent of normal runtime configuration and implementation-specific UI details.

## Decision

Recovery MUST remain reachable without valid `ConfigJson`, valid `SchemaRegistryJson`, or a writable `Log` sheet.

`onOpen()` MUST attempt to create the project menu before surfacing configuration failures. Menu creation is a recovery invariant, not optional UX polish.

Invalid configuration MUST prevent calendar import but MUST NOT prevent recovery entry points from being available. Implementations MUST provide a deterministic reset or recovery path reachable from the spreadsheet UI.

Recovery diagnostics SHOULD write useful status to `Validity` when possible and MAY write structured diagnostics to `Log`, but failure to write diagnostics MUST NOT block menu creation or reset/recovery behavior.

Sheet layout migration MAY repair known layout problems, but normal runtime MUST NOT silently substitute fallback business configuration for invalid `ConfigJson`.

## Invariants

- `onOpen()` MUST always attempt menu creation, including when configuration is invalid.
- Recovery/reset MUST be reachable without valid configuration.
- Recovery/reset MUST be reachable even if logging is unavailable.
- Invalid `ConfigJson` MUST block calendar import.
- Invalid configuration MUST produce deterministic user-visible diagnostics when the sheet can be written.
- Recovery code MUST NOT depend on successful runtime configuration parsing.
- Layout repair and business-configuration fallback are distinct; layout repair MUST NOT become runtime value fallback.
- Existing user-owned sheets, tabs, ranges, and unrelated fields MUST NOT be clobbered by recovery or initialization.

## Allowed Future Changes

- Recovery UI may change from menu item, dialog, sidebar, toast, alert, or another spreadsheet-native entry point if it remains reachable without valid configuration.
- Reset behavior may become more granular if its effects are documented and deterministic.
- Additional recovery states may be defined for malformed config, malformed schema, missing config sheet, unavailable log sheet, or layout mismatch.
- Explicit migration functions may be added to detect old layouts, copy known canonical values, clear stale layout artifacts, and validate final state.

## Explicitly Forbidden Changes

- Blocking menu creation on successful config parsing, schema parsing, calendar access, or log writing.
- Running calendar import with invalid `ConfigJson`.
- Silently replacing invalid `ConfigJson` with defaults, last-valid snapshots, schema values, row values, or UI state during normal runtime.
- Making reset/recovery require valid business configuration.
- Allowing logging failures to prevent recovery.
- Clobbering user-owned sheets, tabs, ranges, or unrelated configuration fields during initialization, migration, or reset.
