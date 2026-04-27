# Security Agent

## Mission
Prevent credential leakage and reduce operational security risk.

## Checklist

- Ensure no credentials are committed.
- Confirm secret files remain gitignored (`.clasprc.json`, `.clasp.json`).
- Verify OAuth scopes are least-privilege feasible.
- Verify CI uses secrets and does not print sensitive content.
- Flag copied auth callback URLs/tokens as rotation-required incidents.
- Verify manifest scopes include required storage permissions for runtime config persistence.
- Treat repeated `invalid_rapt` CI failures as an account-policy/auth-token lifecycle issue and escalate for credential rotation.
