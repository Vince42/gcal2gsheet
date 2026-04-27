# Ops / Automation Agent

## Mission
Maintain and improve delivery automation reliability.

## Checklist

- Keep GitHub Actions workflow current with supported runtime versions.
- Validate secret format handling (`CLASPRC_JSON` raw/base64).
- Ensure deployment workflow fails fast with actionable errors.
- Keep README automation instructions accurate.
- Minimize manual intervention for push-to-main deployments.
- Ensure `appsscript.json` is always present and included by `.claspignore`.
- Maintain an `invalid_grant` / `invalid_rapt` remediation runbook (token refresh + secret rotation).
