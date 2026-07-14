#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../../.." && pwd)"
AI_ROOT="${AI_EXCEL_ADDIN_ROOT:-$ROOT/../AI-excel-addin}"

test -d "$AI_ROOT/packages/agent-gateway"
cd "$AI_ROOT"
PYTHONPATH="packages/agent-gateway" python3 -m pytest -q \
  packages/agent-gateway/tests/test_mcp_client_catalog.py::test_generic_stdio_sheets_mutation_transport_loss_reconnects_without_replay \
  packages/agent-gateway/tests/test_mcp_client_catalog.py::test_generic_stdio_sheets_mutation_timeout_reconnects_without_replay \
  packages/agent-gateway/tests/test_mcp_client_catalog.py::test_sheets_requires_direct_structured_result_but_other_servers_keep_json_fallback
