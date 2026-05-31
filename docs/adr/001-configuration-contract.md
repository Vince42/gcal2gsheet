# ADR-001: Configuration Contract

## Status

Accepted.

## Context

Configuration behavior has repeatedly shifted across named ranges, dialog state, key rows, row overrides, legacy aliases, schema metadata, diagnostic cells, and fallback snapshots. Those competing channels made runtime precedence ambiguous and caused repeated migration, validation, and recovery regressions.

The project requires a single, deterministic configuration contract that future implementations can satisfy without depending on a particular module layout or UI mechanism.

## Decision

The application configuration MUST be represented by one canonical JSON document stored at `Config!ConfigJson`.

`ConfigJson` is the only source of business configuration used for normal runtime behavior. All runtime configuration values MUST be read from that document after validation.

`SchemaRegistryJson` MAY describe allowed keys and validation metadata, but it MUST NOT provide, override, merge, or supplement runtime business configuration.

`Validity` is a diagnostic output only. Implementations MAY write validation status or failure details to it, but MUST NOT read business configuration or control decisions from it.

Strict validation is part of the contract. Invalid JSON, missing required fields, invalid values, and unknown keys MUST be treated as configuration errors once strict validation is enabled.

## Invariants

- `ConfigJson` is the only business-configuration source of truth.
- Invalid `ConfigJson` MUST block calendar import.
- Invalid `ConfigJson` MUST NOT silently reset unrelated fields to defaults during normal runtime.
- Unknown configuration keys MUST be hard errors once strict validation is enabled.
- Canonical key names are authoritative.
- `SchemaRegistryJson` is metadata only.
- `Validity` is output-only.
- Configuration validation MUST be deterministic for a given `ConfigJson` and schema definition.
- Credentials, tokens, and secrets MUST NOT be stored in tracked files or visible sheet cells.

## Allowed Future Changes

- The schema may add new canonical configuration keys when they are documented and validated.
- Validation rules may become stricter when accompanied by explicit migration or release notes.
- Spreadsheet layout may be migrated if the migration preserves the canonical `ConfigJson` semantics.
- A user-facing configuration editor may be added if it writes the canonical `ConfigJson` and does not create a second source of truth.
- `SchemaRegistryJson` may evolve as validation metadata if it remains non-authoritative for runtime business values.

## Explicitly Forbidden Changes

- Reintroducing named ranges, key rows, row overrides, dialog state, cached snapshots, or legacy aliases as runtime configuration sources.
- Falling back from invalid `ConfigJson` to last-valid, default, schema-derived, or UI-derived business configuration during normal runtime.
- Treating `SchemaRegistryJson` as business configuration.
- Reading configuration or runtime control signals from `Validity`.
- Accepting unknown configuration keys silently after strict validation is enabled.
- Storing credentials, tokens, or secrets in tracked files or visible sheet cells.
