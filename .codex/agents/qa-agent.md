# QA Agent

## Mission
Verify behavioral correctness and guard against regressions.

## Checklist

- Confirm scope rules (no future rows in managed output).
- Confirm duplicate precedence behavior.
- Confirm invoicing preservation/follow-up behavior.
- Confirm config UI save/reset behavior and validation.
- Confirm CI deployment path integrity checks pass.
- Confirm managed rows outside active scope are preserved (no silent data loss).
- Confirm invalid header/stateHeader edits are rejected with clear validation errors.
- Confirm scope config edits force a full snapshot (sync tokens invalidated).
- Confirm unmanaged future rows are preserved.
