"""Unit tests for gsheets-mcp tool wrappers."""

import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src import server


def test_gsheet_copy_spreadsheet_error_shape(monkeypatch: pytest.MonkeyPatch) -> None:
    def fail_resolve(source: str) -> tuple[str, str]:
        raise ValueError(f"Spreadsheet not found: {source}")

    monkeypatch.setattr(server.sheets_client, "resolve_spreadsheet_id", fail_resolve)

    payload = json.loads(
        server.gsheet_copy_spreadsheet("Missing Comps", "[hank] Missing Comps")
    )

    assert payload == {
        "status": "error",
        "error": "Spreadsheet not found: Missing Comps",
        "operation": "gsheet_copy_spreadsheet",
    }
