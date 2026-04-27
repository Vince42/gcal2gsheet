# Roadmap / Backlog

## Immediate

- Stabilize CI deployment reliability and token lifecycle handling.
- Ensure config dialog save/reset path is fully functional in deployed runtime.
- Validate end-to-end behavior for scope filtering and duplicates after deployment.
- Add explicit checklist for manifest scopes required by config storage APIs.
- Add regression guard for managed out-of-scope row preservation.
- Add regression guard for structural header/stateHeader validation behavior.

## Near term

- Add lightweight automated checks for Apps Script file integrity and manifest consistency.
- Add operator runbook for common CI failures (`invalid_grant`, missing manifest, missing function exports).

## Later

- Evaluate service-account-based deployment flow (if feasible for target environment).
- Add optional staging deployment path before main/prod push.
