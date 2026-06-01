# ADR-003: Logging Contract

## Status

Accepted.

## Context

Debug logging moved between ad-hoc breadcrumbs, configuration persistence, and a dedicated log sheet. Mixing logs with configuration polluted layout semantics and created migration and validation hazards.

The project requires observability that is useful for recovery and troubleshooting without becoming another configuration channel or single point of failure.

## Decision

`Log` is the only persistent spreadsheet sheet for runtime debug diagnostics.

Persistent runtime diagnostics MUST NOT be written to `Config`, managed data tables, hidden `ID`/`EventID` columns, or configuration metadata locations. Logs MUST be treated as observability output, never as runtime input.

Logging MUST be best-effort. Failure to create, append, format, or rotate logs MUST NOT block menu creation, recovery/reset, configuration validation, or calendar import error handling.

Log entries SHOULD be structured enough to support troubleshooting across recovery and import paths, including event/category, severity or outcome, timestamp, and concise message or detail fields. The exact representation may evolve if the logging contract remains stable.

## Invariants

- `Log` is the only persistent debug-log sheet.
- Logs are output-only and MUST NOT affect runtime configuration or control flow.
- Logging failure MUST NOT block recovery.
- Logging failure MUST NOT convert an invalid configuration into a valid one or hide a validation failure.
- Configuration validation MUST still work when logging is unavailable or malformed.
- Logs MUST NOT be stored in `Config`, managed data tables, hidden `ID`/`EventID` columns, `ConfigJson`, `SchemaRegistryJson`, or `Validity`.
- Credentials, tokens, and secrets MUST NOT be logged.

## Allowed Future Changes

- The log schema may add structured fields, severity levels, categories, correlation identifiers, or retention metadata.
- Log rotation, truncation, or archival may be added if it does not affect configuration, recovery, or import correctness.
- In-memory or console diagnostics may supplement the `Log` sheet, provided persistent spreadsheet diagnostics still use `Log` only.
- Logging implementation may move between modules if the output-only and best-effort contract is preserved.

## Explicitly Forbidden Changes

- Writing persistent debug logs into configuration, managed data tables, hidden ID/EventID columns, schema, or validity locations.
- Reading logs as business configuration, validation metadata, migration input, or runtime control signals.
- Allowing malformed or unavailable logs to prevent `onOpen()` menu creation, reset/recovery, or config validation.
- Logging credentials, tokens, secrets, or other sensitive authentication material.
- Using logging side effects to repair or mutate business configuration.
