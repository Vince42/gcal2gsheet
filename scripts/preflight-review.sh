#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

echo "[preflight] running git diff whitespace check..."
git diff --check

echo "[preflight] checking for unsafe raw model injection in HTML templates..."
if rg -n "<\\?!=\\s*JSON\\.stringify\\(" --glob "*.html" . >/tmp/preflight_html_hits.txt; then
  cat /tmp/preflight_html_hits.txt
  echo "[preflight] FAIL: found raw unescaped template JSON injection pattern."
  exit 1
fi

echo "[preflight] checking that credential files are not tracked..."
TRACKED_CREDS="$(git ls-files | rg -n "^\\.clasprc\\.json$|^\\.clasp\\.json$" || true)"
if [[ -n "$TRACKED_CREDS" ]]; then
  echo "$TRACKED_CREDS"
  echo "[preflight] FAIL: tracked clasp credential/project files detected."
  exit 1
fi

echo "[preflight] checking for mandatory release-gate policy presence..."
if ! rg -n "Mandatory release gates|Scope gate|Security gate" agents.md >/dev/null; then
  echo "[preflight] FAIL: mandatory release-gate policy missing from agents.md."
  exit 1
fi

echo "[preflight] PASS"
