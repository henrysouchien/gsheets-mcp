#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../../.." && pwd)"
cd "$ROOT"
venv/bin/python -m pytest -q tests/test_cli.py
test ! -e run_server.py
test ! -e src/server.py
test ! -e src/sheets_client.py
