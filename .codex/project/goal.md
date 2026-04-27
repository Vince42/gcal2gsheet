# Project Goal

Maintain a reliable, secure, and auditable Google Apps Script solution that imports scoped calendar events into Google Sheets for invoicing preparation, while preserving data integrity and enabling safe configuration changes without code edits.

## Success criteria

- No future events are imported into managed scope.
- Duplicate handling follows documented precedence deterministically.
- Invoiced rows are preserved; changed invoiced events produce follow-up rows.
- Configuration is editable via UI and validated before persistence.
- CI/CD deployment to Apps Script is reproducible and secure.
