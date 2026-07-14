#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../../.." && pwd)"
cd "$ROOT"
venv/bin/python -m pytest -q \
  tests/test_server.py::test_discovery_exposes_only_canonical_strict_input_schemas \
  tests/test_server.py::test_every_input_and_output_object_schema_is_closed_world \
  tests/test_server.py::test_output_schema_requires_wire_discriminators_and_operation \
  tests/test_server.py::test_success_is_direct_structured_content_and_not_json_inside_json \
  tests/test_server.py::test_unknown_fields_fail_before_dispatch_without_echoing_input \
  tests/test_server.py::test_real_stdio_initialize_list_and_validation_call \
  tests/test_server.py::test_real_stdio_success_is_direct_structured_content
