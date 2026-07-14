#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../../.." && pwd)"

"$ROOT/venv/bin/gsheets-mcp" capabilities --json |
  "$ROOT/venv/bin/python" -c '
import json, sys
p = json.load(sys.stdin)
names = [tool["name"] for tool in p["tools"]]
assert len(names) == 9
assert all(name.startswith("gsheets_") for name in names)
assert not any(name.startswith("gsheet_") for name in names)
'

jq -e 'select(.pass == 2 and .corpus_id == "C003" and .classification == "inferred_and_acted" and .test_double == true)' \
  "$ROOT/agent_ergonomics_audit/audit/intent_inference_corpus.jsonl" >/dev/null

cd "$ROOT"
venv/bin/python -m pytest -q \
  tests/test_server.py::test_retired_tool_namespace_is_not_behaviorally_recognized
