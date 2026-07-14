#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../../.." && pwd)"
cd "$ROOT"
venv/bin/python -m pytest -q \
  tests/test_server.py::test_mode_specific_discovery_is_credential_free \
  tests/test_server.py::test_broker_direct_search_call_returns_mode_aware_capability_error \
  tests/test_server.py::test_title_reference_is_rejected_before_credentials \
  tests/test_cli.py::test_capabilities_json_is_credential_free_and_mode_specific \
  tests/test_cli.py::test_invalid_mode_fails_closed_with_stable_stderr_diagnostic
