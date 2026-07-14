"""Offline unit tests for strict references and mutation recovery semantics."""

from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from gsheets_mcp import sheets_client


SPREADSHEET_ID = "1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789"


@pytest.mark.parametrize(
    "reference",
    [
        SPREADSHEET_ID,
        f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid=0",
        f"https://DOCS.GOOGLE.COM/spreadsheets/d/{SPREADSHEET_ID}/view",
        f"https://docs.google.com/spreadsheets/u/0/d/{SPREADSHEET_ID}/edit?usp=sharing",
        f"https://docs.google.com/spreadsheets/u/12/d/{SPREADSHEET_ID}/copy",
    ],
)
def test_parse_spreadsheet_reference_accepts_only_id_or_strict_url(
    reference: str,
) -> None:
    assert sheets_client.parse_spreadsheet_reference(reference) == SPREADSHEET_ID


@pytest.mark.parametrize(
    "reference",
    [
        "Quarterly Comps",
        "short-id",
        "https://example.com/spreadsheets/d/not-google",
        f"https://docs.google.com@evil.example/spreadsheets/d/{SPREADSHEET_ID}",
        f"https://docs.google.com:8443/spreadsheets/d/{SPREADSHEET_ID}",
        f"https://docs.google.com:bad/spreadsheets/d/{SPREADSHEET_ID}",
        f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/../d/other7890123456789012",
        f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export",
        f"https://docs.google.com/spreadsheets/u/x/d/{SPREADSHEET_ID}/edit",
    ],
)
def test_parse_spreadsheet_reference_rejects_titles_and_unsafe_urls(
    reference: str,
) -> None:
    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client.parse_spreadsheet_reference(reference)

    assert exc_info.value.code == "invalid_spreadsheet_reference"
    assert exc_info.value.details["outcome_state"] == "not_started"
    assert exc_info.value.details["retry_action"] == "correct_arguments"


def test_resolve_spreadsheet_verifies_normalized_id_with_sheets_only() -> None:
    service = MagicMock()
    service.spreadsheets.return_value.get.return_value.execute.return_value = {
        "spreadsheetId": SPREADSHEET_ID,
        "properties": {"title": "Validated"},
    }

    assert sheets_client.resolve_spreadsheet(
        service,
        f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit",
    ) == (SPREADSHEET_ID, "Validated")
    service.spreadsheets.return_value.get.assert_called_once_with(
        spreadsheetId=SPREADSHEET_ID,
        fields="spreadsheetId,properties(title)",
    )


def test_read_range_rejects_unknown_render_options_before_api_call() -> None:
    service = MagicMock()

    with pytest.raises(ValueError, match="Invalid value_render_option"):
        sheets_client.read_sheet_range(
            service,
            SPREADSHEET_ID,
            "Comps!A1:B2",
            value_render_option="RAW",
        )
    with pytest.raises(ValueError, match="Invalid date_time_render_option"):
        sheets_client.read_sheet_range(
            service,
            SPREADSHEET_ID,
            "Comps!A1:B2",
            date_time_render_option="ISO8601",
        )

    service.spreadsheets.assert_not_called()


def test_search_spreadsheets_returns_normalized_reusable_references() -> None:
    drive_service = MagicMock()
    drive_service.files.return_value.list.return_value.execute.return_value = {
        "files": [
            {
                "id": SPREADSHEET_ID,
                "name": "O'Reilly Comps",
                "modifiedTime": "2026-07-14T12:00:00Z",
                "webViewLink": f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}",
            }
        ]
    }

    result = sheets_client.search_spreadsheets(
        drive_service,
        query="O'Reilly",
        limit=5,
    )

    assert result == [
        {
            "spreadsheet": SPREADSHEET_ID,
            "title": "O'Reilly Comps",
            "url": f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}",
            "modified_at": "2026-07-14T12:00:00Z",
        }
    ]
    list_kwargs = drive_service.files.return_value.list.call_args.kwargs
    assert "O\\'Reilly" in list_kwargs["q"]
    assert sheets_client.GOOGLE_SHEET_MIME in list_kwargs["q"]
    assert "trashed = false" in list_kwargs["q"]
    assert list_kwargs["supportsAllDrives"] is True
    assert list_kwargs["includeItemsFromAllDrives"] is True
    assert list_kwargs["pageSize"] == 5


def _copy_service() -> MagicMock:
    service = MagicMock()
    spreadsheets = service.spreadsheets.return_value
    spreadsheets.get.return_value.execute.return_value = {
        "properties": {"locale": "en_US", "timeZone": "America/New_York"},
        "namedRanges": [{"name": "RosterRange"}],
        "sheets": [
            {
                "properties": {
                    "sheetId": 101,
                    "title": "Comps",
                    "index": 0,
                    "gridProperties": {"rowCount": 20, "columnCount": 8},
                }
            },
            {
                "properties": {
                    "sheetId": 102,
                    "title": "Assumptions",
                    "index": 1,
                    "gridProperties": {"rowCount": 10, "columnCount": 4},
                }
            },
        ],
    }
    spreadsheets.create.return_value.execute.return_value = {
        "spreadsheetId": "destination-spreadsheet-id",
        "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/destination-spreadsheet-id",
        "properties": {"locale": "en_GB", "timeZone": "UTC"},
        "sheets": [{"properties": {"sheetId": 900, "title": "Sheet1", "index": 0}}],
    }
    spreadsheets.batchUpdate.return_value.execute.return_value = {}
    return service


def test_copy_spreadsheet_returns_destination_and_structured_warnings() -> None:
    service = _copy_service()
    copied_properties = iter(
        [
            {"sheetId": 201, "title": "Copy of Comps", "index": 1},
            {"sheetId": 202, "title": "Copy of Assumptions", "index": 2},
        ]
    )
    service.spreadsheets.return_value.sheets.return_value.copyTo.return_value.execute.side_effect = (
        lambda: next(copied_properties)
    )

    result = sheets_client.copy_spreadsheet(
        service,
        SPREADSHEET_ID,
        "[hank] Working Copy",
    )

    assert result["spreadsheet"] == "destination-spreadsheet-id"
    assert result["tabs"] == ["Comps", "Assumptions"]
    assert [warning["code"] for warning in result["warnings"]] == [
        "named_ranges_not_copied",
        "locale_not_preserved",
        "time_zone_not_preserved",
    ]
    batch_body = service.spreadsheets.return_value.batchUpdate.call_args.kwargs["body"]
    assert batch_body["requests"][0] == {"deleteSheet": {"sheetId": 900}}
    assert [
        request["updateSheetProperties"]["properties"]["title"]
        for request in batch_body["requests"][1:]
    ] == ["Comps", "Assumptions"]


def test_copy_partial_failure_preserves_destination_progress() -> None:
    service = _copy_service()

    def copy_to(*, spreadsheetId, sheetId, body):
        del spreadsheetId, body
        request = MagicMock()
        if sheetId == 101:
            request.execute.return_value = {"sheetId": 201}
        else:
            request.execute.side_effect = ConnectionError("response lost")
        return request

    service.spreadsheets.return_value.sheets.return_value.copyTo.side_effect = copy_to

    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client.copy_spreadsheet(
            service,
            SPREADSHEET_ID,
            "[hank] Partial Copy",
        )

    error = exc_info.value
    assert error.code == "operation_partial"
    assert error.details["outcome_state"] == "partial"
    assert error.details["mutation_may_have_occurred"] is True
    assert error.details["retry_safe"] is False
    assert error.details["retry_automatic"] is False
    assert error.details["recovery"] == {
        "kind": "copy_progress",
        "destination_spreadsheet": "destination-spreadsheet-id",
        "destination_url": "https://docs.google.com/spreadsheets/d/destination-spreadsheet-id",
        "confirmed_tabs": ["Comps"],
        "active_tab": "Assumptions",
        "active_tab_state": "uncertain",
        "remaining_tabs": ["Assumptions"],
        "finalization_state": "not_started",
    }


def test_copy_source_failure_is_not_started_before_destination_creation() -> None:
    service = MagicMock()
    spreadsheets = service.spreadsheets.return_value
    spreadsheets.get.return_value.execute.side_effect = ConnectionError("offline")

    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client.copy_spreadsheet(service, SPREADSHEET_ID, "Copy")

    assert exc_info.value.details["outcome_state"] == "not_started"
    assert exc_info.value.details["mutation_may_have_occurred"] is False
    spreadsheets.create.assert_not_called()


def test_copy_create_response_loss_is_uncertain_without_fabricated_destination() -> (
    None
):
    service = _copy_service()
    spreadsheets = service.spreadsheets.return_value
    spreadsheets.create.return_value.execute.side_effect = ConnectionError(
        "response lost"
    )

    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client.copy_spreadsheet(service, SPREADSHEET_ID, "Copy")

    assert exc_info.value.details["outcome_state"] == "uncertain"
    assert exc_info.value.details["phase"] == "create_destination"
    assert exc_info.value.details["recovery"] is None


def test_copy_finalization_failure_preserves_confirmed_destination_progress() -> None:
    service = _copy_service()
    sheets_api = service.spreadsheets.return_value.sheets.return_value
    copied_properties = iter([{"sheetId": 201}, {"sheetId": 202}])
    sheets_api.copyTo.return_value.execute.side_effect = lambda: next(copied_properties)
    service.spreadsheets.return_value.batchUpdate.return_value.execute.side_effect = (
        ConnectionError("response lost")
    )

    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client.copy_spreadsheet(service, SPREADSHEET_ID, "Copy")

    error = exc_info.value
    assert error.details["outcome_state"] == "partial"
    assert error.details["phase"] == "finalize_destination"
    assert error.details["recovery"] == {
        "kind": "copy_progress",
        "destination_spreadsheet": "destination-spreadsheet-id",
        "destination_url": "https://docs.google.com/spreadsheets/d/destination-spreadsheet-id",
        "confirmed_tabs": ["Comps", "Assumptions"],
        "active_tab": None,
        "active_tab_state": "not_started",
        "remaining_tabs": [],
        "finalization_state": "uncertain",
    }


def test_recalculate_blank_range_is_a_verified_noop() -> None:
    service = MagicMock()
    values = service.spreadsheets.return_value.values.return_value
    values.get.return_value.execute.return_value = {"values": [["", None], []]}

    result = sheets_client.recalculate_sheet_range(
        service,
        SPREADSHEET_ID,
        "Comps!A1:B2",
    )

    assert result == {
        "range": "Comps!A1:B2",
        "cell_count": 0,
        "recovery_performed": False,
        "formulas_verified": True,
    }
    values.clear.assert_not_called()
    values.update.assert_not_called()


def test_recalculate_literal_range_is_rejected_without_mutation() -> None:
    service = MagicMock()
    values = service.spreadsheets.return_value.values.return_value
    values.get.return_value.execute.return_value = {
        "values": [["=SUM(A2:A3)", 42], ["", None]]
    }

    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client.recalculate_sheet_range(
            service,
            SPREADSHEET_ID,
            "Comps!A1:B2",
        )

    assert exc_info.value.code == "formula_range_required"
    assert exc_info.value.details["outcome_state"] == "unchanged"
    assert exc_info.value.details["mutation_may_have_occurred"] is False
    values.clear.assert_not_called()
    values.update.assert_not_called()


def test_recalculate_formula_range_clears_restores_and_verifies_exactly() -> None:
    service = MagicMock()
    values = service.spreadsheets.return_value.values.return_value
    formulas = [["=SUM(A2:A3)", ""], [None, "=CUSTOM(B2)"]]
    values.get.return_value.execute.side_effect = [
        {"values": formulas},
        {"values": formulas},
    ]
    values.clear.return_value.execute.return_value = {"clearedRange": "Comps!A1:B2"}
    values.update.return_value.execute.return_value = {
        "updatedRange": "Comps!A1:B2",
        "updatedCells": 4,
    }

    result = sheets_client.recalculate_sheet_range(
        service,
        SPREADSHEET_ID,
        "Comps!A1:B2",
    )

    assert result == {
        "range": "Comps!A1:B2",
        "cell_count": 2,
        "recovery_performed": False,
        "formulas_verified": True,
    }
    values.clear.assert_called_once_with(
        spreadsheetId=SPREADSHEET_ID,
        range="Comps!A1:B2",
        body={},
    )
    values.update.assert_called_once_with(
        spreadsheetId=SPREADSHEET_ID,
        range="Comps!A1:B2",
        valueInputOption="USER_ENTERED",
        body={"values": formulas},
    )
    assert values.get.call_count == 2


def test_recalculate_uses_one_exact_compensation_after_primary_write_failure() -> None:
    service = MagicMock()
    values = service.spreadsheets.return_value.values.return_value
    formulas = [["=CUSTOM(A1)"]]
    values.get.return_value.execute.side_effect = [
        {"values": formulas},
        {"values": formulas},
    ]
    values.clear.return_value.execute.return_value = {"clearedRange": "Comps!A1"}
    values.update.return_value.execute.side_effect = [
        ConnectionError("primary write response lost"),
        {"updatedRange": "Comps!A1", "updatedCells": 1},
    ]

    result = sheets_client.recalculate_sheet_range(
        service,
        SPREADSHEET_ID,
        "Comps!A1",
    )

    assert result == {
        "range": "Comps!A1",
        "cell_count": 1,
        "recovery_performed": True,
        "formulas_verified": True,
    }
    assert values.update.call_count == 2
    for call in values.update.call_args_list:
        assert call.kwargs["body"] == {"values": formulas}
        assert call.kwargs["range"] == "Comps!A1"


def test_recalculate_uncertain_clear_enters_compensation_and_verifies() -> None:
    service = MagicMock()
    values = service.spreadsheets.return_value.values.return_value
    formulas = [["=CUSTOM(A1)"]]
    values.get.return_value.execute.side_effect = [
        {"values": formulas},
        {"values": formulas},
    ]
    values.clear.return_value.execute.side_effect = ConnectionError(
        "clear response lost"
    )
    values.update.return_value.execute.return_value = {
        "updatedRange": "Comps!A1",
        "updatedCells": 1,
    }

    result = sheets_client.recalculate_sheet_range(
        service,
        SPREADSHEET_ID,
        "Comps!A1",
    )

    assert result["recovery_performed"] is True
    assert result["formulas_verified"] is True
    assert values.update.call_count == 1


def test_recalculate_uncertain_clear_never_repeats_failed_compensation() -> None:
    service = MagicMock()
    values = service.spreadsheets.return_value.values.return_value
    formulas = [["=CUSTOM(A1)"]]
    values.get.return_value.execute.return_value = {"values": formulas}
    values.clear.return_value.execute.side_effect = ConnectionError(
        "clear response lost"
    )
    values.update.return_value.execute.side_effect = ConnectionError(
        "restore response lost"
    )

    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client.recalculate_sheet_range(
            service,
            SPREADSHEET_ID,
            "Comps!A1",
        )

    error = exc_info.value
    assert error.code == "recalculation_recovery_failed"
    assert error.details["outcome_state"] == "partial"
    assert error.details["retry_safe"] is False
    assert error.details["recovery"]["compensation_attempted"] is True
    assert error.details["recovery"]["compensation_verified"] is False
    assert values.update.call_count == 1
    assert "CUSTOM" not in str(error)
    assert "CUSTOM" not in str(error.details)


def test_recalculate_proven_clear_rejection_never_runs_compensation() -> None:
    service = MagicMock()
    values = service.spreadsheets.return_value.values.return_value
    formulas = [["=CUSTOM(A1)"]]
    values.get.return_value.execute.return_value = {"values": formulas}
    values.clear.return_value.execute.side_effect = sheets_client.HttpError(
        SimpleNamespace(status=400, reason="Bad Request"),
        b"sensitive body",
    )

    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client.recalculate_sheet_range(service, SPREADSHEET_ID, "Comps!A1")

    assert exc_info.value.details["outcome_state"] == "unchanged"
    assert exc_info.value.details["mutation_may_have_occurred"] is False
    values.update.assert_not_called()


def test_recalculate_verification_mismatch_runs_one_exact_compensation() -> None:
    service = MagicMock()
    values = service.spreadsheets.return_value.values.return_value
    formulas = [["=CUSTOM(A1)"]]
    values.get.return_value.execute.side_effect = [
        {"values": formulas},
        {"values": [["=WRONG()"]]},
        {"values": formulas},
    ]
    values.clear.return_value.execute.return_value = {"clearedRange": "Comps!A1"}
    values.update.return_value.execute.return_value = {
        "updatedRange": "Comps!A1",
        "updatedCells": 1,
    }

    result = sheets_client.recalculate_sheet_range(
        service,
        SPREADSHEET_ID,
        "Comps!A1",
    )

    assert result["recovery_performed"] is True
    assert result["formulas_verified"] is True
    assert values.update.call_count == 2
    assert all(
        call.kwargs["body"] == {"values": formulas}
        for call in values.update.call_args_list
    )


def test_recalculate_failed_compensation_reports_partial_range_recovery() -> None:
    service = MagicMock()
    values = service.spreadsheets.return_value.values.return_value
    formulas = [["=CUSTOM(A1)"]]
    values.get.return_value.execute.return_value = {"values": formulas}
    values.clear.return_value.execute.return_value = {"clearedRange": "Comps!A1"}
    values.update.return_value.execute.side_effect = ConnectionError("write failed")

    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client.recalculate_sheet_range(
            service,
            SPREADSHEET_ID,
            "Comps!A1",
        )

    error = exc_info.value
    assert error.code == "recalculation_recovery_failed"
    assert error.details["outcome_state"] == "partial"
    assert error.details["retry_safe"] is False
    assert error.details["recovery"] == {
        "kind": "range_state",
        "spreadsheet": SPREADSHEET_ID,
        "range": "Comps!A1",
        "formula_cell_count": 1,
        "compensation_attempted": True,
        "compensation_verified": False,
    }
    assert values.update.call_count == 2
    assert "CUSTOM" not in str(error)
    assert "CUSTOM" not in str(error.details)


def test_mutation_failure_distinguishes_rejected_from_uncertain() -> None:
    rejected = sheets_client.HttpError(
        SimpleNamespace(status=400, reason="Bad Request"),
        b"bad request",
    )
    rejected_error = sheets_client.mutation_failure(
        rejected,
        phase="write_range",
        dispatched=True,
    )
    uncertain_error = sheets_client.mutation_failure(
        ConnectionError("response lost"),
        phase="write_range",
        dispatched=True,
    )

    assert rejected_error.details["outcome_state"] == "unchanged"
    assert rejected_error.details["retry_safe"] is True
    assert uncertain_error.details["outcome_state"] == "uncertain"
    assert uncertain_error.details["retry_safe"] is False
