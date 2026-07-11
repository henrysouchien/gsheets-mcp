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


def test_gsheet_copy_spreadsheet_surfaces_copy_semantics_warnings(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    warning = (
        "1 spreadsheet-level named range(s) NOT copied "
        "(tab-level copy cannot carry them): RosterRange"
    )
    monkeypatch.setattr(
        server.sheets_client, "resolve_spreadsheet_id", lambda source: ("src", "Src")
    )
    monkeypatch.setattr(server.sheets_client, "get_sheets_service", lambda: object())
    monkeypatch.setattr(
        server.sheets_client,
        "copy_spreadsheet",
        lambda *args, **kwargs: {
            "spreadsheet_id": "new",
            "title": "[hank] Copy",
            "url": "https://docs.google.com/spreadsheets/d/new",
            "copied_tabs": ["Tab1"],
            "warnings": [warning],
        },
    )

    payload = json.loads(server.gsheet_copy_spreadsheet("src", "[hank] Copy"))

    assert payload["status"] == "ok"
    assert payload["warnings"] == [warning]
