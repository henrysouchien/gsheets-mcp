#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../../.." && pwd)"
AI_ROOT="${AI_EXCEL_ADDIN_ROOT:-$ROOT/../AI-excel-addin}"
RISK_ROOT="${RISK_MODULE_ROOT:-$ROOT/../risk_module}"

jq -e '.mcpServers["gsheets-mcp"].args == ["serve"] and .mcpServers["gsheets-mcp"].env.GSHEETS_TOKEN_MODE == "broker"' \
  "$AI_ROOT/deploy/mcp.production.json" >/dev/null
jq -e '.mcpServers["gsheets-mcp"].args == ["serve"] and .mcpServers["gsheets-mcp"].env.GSHEETS_TOKEN_MODE == "broker"' \
  "$RISK_ROOT/config/research_gateway_mcp.local.json" >/dev/null
rg -q 'no-compatibility cutover' "$AI_ROOT/deploy/README.md"
rg -Uq 'as one maintenance-window[[:space:]]+unit' "$AI_ROOT/deploy/README.md"
