#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../../.." && pwd)"
cd "$ROOT"
venv/bin/python scripts/mcp_lint.py --json | venv/bin/python -c 'import json,sys; assert json.load(sys.stdin) == []'
venv/bin/python -m pytest -q
rg -q 'python3 scripts/mcp_lint.py' .github/workflows/mcp-audit.yml
rg -q 'python3 -m pytest -q' .github/workflows/mcp-audit.yml
