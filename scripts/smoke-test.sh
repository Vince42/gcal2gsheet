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

echo "smoke-test: PASS"
