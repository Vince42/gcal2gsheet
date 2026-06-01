#!/usr/bin/env bash
set -euo pipefail

# Static smoke test for strict ConfigJson + schema registry.
rg -n "function readConfigStateFromSheet_\(" Config.gs >/dev/null
rg -n "SchemaRegistryJson" Config.gs >/dev/null
rg -n "validateSchemaRegistry_" Config.gs >/dev/null
rg -n "validateConfigStrictWithSchema_" Config.gs >/dev/null
rg -n "unknown key" Config.gs >/dev/null
rg -n "function resetConfigAndLogSheets_\\(" Config.gs >/dev/null
rg -n "function findBestConfigJsonInSheet_\\(" Config.gs >/dev/null

# Static smoke test for config validation UX/recovery.
rg -n "function onOpen\(" Code.gs >/dev/null
rg -n "ensureMenuVisible_\(ui\)" Code.gs >/dev/null
rg -n "ui\.alert\('Configuration validation failed'" Code.gs >/dev/null
rg -n "function onEdit\(e\)" Code.gs >/dev/null

# Static smoke test for structured Log sheet.
rg -n "function ensureLogSheet_\(" Config.gs >/dev/null
rg -n "Timestamp.*Level.*Component.*Event.*Message" Config.gs >/dev/null
rg -n "setFontFamily\('Courier New'\)" Config.gs >/dev/null

# Static smoke test for generic default config and self-healing full imports.
rg -n "calendarNames: \['Calendar'\]" Config.gs >/dev/null
rg -n "defaultCalendarName: 'Calendar'" Config.gs >/dev/null
rg -n "invoicingSheetName: 'Invoicing'" Config.gs >/dev/null
rg -n "'Status'" Config.gs >/dev/null
rg -n "nonBillableSheetName: 'Non-Billable'" Config.gs >/dev/null
rg -n "nonBillableTableName: 'NonBillable'" Config.gs >/dev/null
rg -n "function assertValidTableName_" Config.gs >/dev/null
rg -n "ensureManagedWorkbookStructure_" "Table Service.gs" >/dev/null
rg -n "assertSheetHasExpectedColumns_" "Table Service.gs" >/dev/null
rg -n "ensureNonBillableSheet_" "Table Service.gs" >/dev/null
rg -n "migrateCalendarInvoicesToInvoicing_" "Invoicing Store.gs" >/dev/null
rg -n "repairInvoicingStateFromImportedEvents_" "Invoicing Store.gs" >/dev/null
rg -n "readNonBillableState_" "Non Billable Store.gs" >/dev/null
rg -n "applyRegisterStatusesToImportedEvents_" "Non Billable Store.gs" >/dev/null
rg -n "Performing full import for self-healing reconciliation" Code.gs >/dev/null
if rg -n "fetchIncrementalChanges_\(ss, calendars, timeZone\)" Code.gs >/dev/null; then
  echo "smoke-test: FAIL: updateCalendarSheets should not use incremental fetches."
  exit 1
fi

# Static smoke test for import-start and invoice-preservation safeguards.
rg -n "ignoredBeforeImportStartCount" "State Store.gs" >/dev/null
rg -n "excluded from this update and left unchanged" Code.gs >/dev/null
rg -n "Registered event updates acknowledged" Code.gs >/dev/null
rg -n "!row\.invoiceNumber" "Duplicate Engine.gs" >/dev/null
rg -n "isExistingRowInScope_\(row\.values, scope\)" "Duplicate Engine.gs" >/dev/null

echo "smoke-test: PASS"
