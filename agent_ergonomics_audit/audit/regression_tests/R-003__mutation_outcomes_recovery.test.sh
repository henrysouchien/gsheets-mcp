#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../../.." && pwd)"
cd "$ROOT"
venv/bin/python -m pytest -q \
  tests/test_server.py::test_mutation_failure_is_structured_and_marks_uncertain_outcome \
  tests/test_server.py::test_copy_partial_error_exposes_destination_progress \
  tests/test_sheets_client.py::test_recalculate_uncertain_clear_enters_compensation_and_verifies \
  tests/test_sheets_client.py::test_recalculate_uncertain_clear_never_repeats_failed_compensation \
  tests/test_sheets_client.py -k 'copy_ or recalculate_ or mutation_failure_'
