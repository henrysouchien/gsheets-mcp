"""Unit tests for gsheets-mcp Sheets client helpers."""

import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src import sheets_client


def _http_404_error() -> Exception:
    return sheets_client.HttpError(
        SimpleNamespace(status=404, reason="Not Found"),
        b"Not Found",
    )


def _files_list_call_args(service_mock) -> dict:
    return service_mock.files.return_value.list.call_args.kwargs


def test_value_render_option_validation() -> None:
    sheets_service = MagicMock()

    with pytest.raises(ValueError) as exc_info:
        sheets_client.read_sheet_range(
            sheets_service,
            spreadsheet_id="sheet-123",
            range_a1="Sheet1!A1:A2",
            value_render_option="BAD_VALUE",
        )

    msg = str(exc_info.value)
    assert "Invalid value_render_option 'BAD_VALUE'" in msg
    assert "FORMULA" in msg
    assert "FORMATTED_VALUE" in msg
    assert "UNFORMATTED_VALUE" in msg


def test_date_time_render_option_validation() -> None:
    sheets_service = MagicMock()

    with pytest.raises(ValueError) as exc_info:
        sheets_client.read_sheet_range(
            sheets_service,
            spreadsheet_id="sheet-123",
            range_a1="Sheet1!A1:A2",
            date_time_render_option="BAD_DATE",
        )

    msg = str(exc_info.value)
    assert "Invalid date_time_render_option 'BAD_DATE'" in msg
    assert "FORMATTED_STRING" in msg
    assert "SERIAL_NUMBER" in msg


def test_resolve_spreadsheet_id_by_pattern(monkeypatch: pytest.MonkeyPatch) -> None:
    drive_service = MagicMock()
    drive_service.files.return_value.get.return_value.execute.return_value = {
        "id": "abc1234567890XYZ___123",
        "name": "My Sheet",
        "mimeType": sheets_client.GOOGLE_SHEET_MIME,
    }
    monkeypatch.setattr(sheets_client, "authenticate", lambda: drive_service)

    spreadsheet_id, title = sheets_client.resolve_spreadsheet_id(
        "abc1234567890XYZ___123"
    )

    assert spreadsheet_id == "abc1234567890XYZ___123"
    assert title == "My Sheet"
    drive_service.files.return_value.get.assert_called_once()
    drive_service.files.return_value.list.assert_not_called()


def test_resolve_spreadsheet_id_by_name(monkeypatch: pytest.MonkeyPatch) -> None:
    drive_service = MagicMock()
    drive_service.files.return_value.get.return_value.execute.side_effect = _http_404_error()
    drive_service.files.return_value.list.return_value.execute.return_value = {
        "files": [{"id": "sheet-1", "name": "Comps"}]
    }
    monkeypatch.setattr(sheets_client, "authenticate", lambda: drive_service)

    spreadsheet_id, title = sheets_client.resolve_spreadsheet_id("Comps")

    assert spreadsheet_id == "sheet-1"
    assert title == "Comps"
    list_kwargs = _files_list_call_args(drive_service)
    assert list_kwargs["supportsAllDrives"] is True
    assert list_kwargs["includeItemsFromAllDrives"] is True
    assert "name = 'Comps'" in list_kwargs["q"]
    assert sheets_client.GOOGLE_SHEET_MIME in list_kwargs["q"]


def test_resolve_spreadsheet_ambiguous(monkeypatch: pytest.MonkeyPatch) -> None:
    drive_service = MagicMock()

    def fake_get(*, fileId, **kwargs):
        req = MagicMock()
        if fileId in {"parent-1", "parent-2"}:
            req.execute.return_value = {"id": fileId, "name": f"Folder {fileId[-1]}"}
        else:
            req.execute.side_effect = _http_404_error()
        return req

    drive_service.files.return_value.get.side_effect = fake_get
    drive_service.files.return_value.list.return_value.execute.return_value = {
        "files": [
            {
                "id": "sheet-a",
                "name": "Comps",
                "modifiedTime": "2026-02-21T00:00:00Z",
                "webViewLink": "https://docs.google.com/spreadsheets/d/sheet-a",
                "parents": ["parent-1"],
            },
            {
                "id": "sheet-b",
                "name": "Comps",
                "modifiedTime": "2026-02-22T00:00:00Z",
                "webViewLink": "https://docs.google.com/spreadsheets/d/sheet-b",
                "parents": ["parent-2"],
            },
        ]
    }
    monkeypatch.setattr(sheets_client, "authenticate", lambda: drive_service)

    with pytest.raises(ValueError) as exc_info:
        sheets_client.resolve_spreadsheet_id("Comps")

    msg = str(exc_info.value)
    assert "Multiple spreadsheets found" in msg
    assert "sheet-a" in msg
    assert "sheet-b" in msg


def test_search_spreadsheets_filters_mime() -> None:
    drive_service = MagicMock()
    drive_service.files.return_value.list.return_value.execute.return_value = {
        "files": []
    }

    sheets_client.search_spreadsheets(drive_service, query="Comp", max_results=5)

    list_kwargs = _files_list_call_args(drive_service)
    assert sheets_client.GOOGLE_SHEET_MIME in list_kwargs["q"]
    assert "trashed = false" in list_kwargs["q"]
    assert list_kwargs["pageSize"] == 5


def test_search_spreadsheets_shared_drives() -> None:
    drive_service = MagicMock()
    drive_service.files.return_value.list.return_value.execute.return_value = {
        "files": []
    }

    sheets_client.search_spreadsheets(drive_service, query="Comp")

    list_kwargs = _files_list_call_args(drive_service)
    assert list_kwargs["supportsAllDrives"] is True
    assert list_kwargs["includeItemsFromAllDrives"] is True


def test_search_spreadsheets_escapes_apostrophes() -> None:
    drive_service = MagicMock()
    drive_service.files.return_value.list.return_value.execute.return_value = {
        "files": []
    }

    sheets_client.search_spreadsheets(drive_service, query="O'Reilly")

    list_kwargs = _files_list_call_args(drive_service)
    assert "O\\'Reilly" in list_kwargs["q"]


def test_copy_spreadsheet_copies_all_tabs_and_renames_in_order() -> None:
    sheets_service = MagicMock()
    spreadsheets_api = sheets_service.spreadsheets.return_value
    sheets_api = spreadsheets_api.sheets.return_value
    call_order: list[str] = []

    def fake_get(**kwargs):
        call_order.append("get")
        request = MagicMock()
        request.execute.return_value = {
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
            ]
        }
        return request

    def fake_create(**kwargs):
        call_order.append("create")
        request = MagicMock()
        request.execute.return_value = {
            "spreadsheetId": "new-sheet",
            "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/new-sheet",
            "sheets": [{"properties": {"sheetId": 900, "title": "Sheet1", "index": 0}}],
        }
        return request

    copied_props = iter(
        [
            {"sheetId": 201, "title": "Copy of Comps", "index": 1},
            {"sheetId": 202, "title": "Copy of Assumptions", "index": 2},
        ]
    )

    def fake_copy_to(*, spreadsheetId, sheetId, body):
        call_order.append(f"copyTo:{sheetId}")
        request = MagicMock()
        request.execute.return_value = next(copied_props)
        return request

    def fake_batch_update(**kwargs):
        call_order.append("batchUpdate")
        request = MagicMock()
        request.execute.return_value = {}
        return request

    spreadsheets_api.get.side_effect = fake_get
    spreadsheets_api.create.side_effect = fake_create
    sheets_api.copyTo.side_effect = fake_copy_to
    spreadsheets_api.batchUpdate.side_effect = fake_batch_update

    result = sheets_client.copy_spreadsheet(
        sheets_service,
        source_spreadsheet_id="source-sheet",
        new_title="[hank] HCM Comps",
    )

    assert call_order == ["get", "create", "copyTo:101", "copyTo:102", "batchUpdate"]
    spreadsheets_api.create.assert_called_once_with(
        body={"properties": {"title": "[hank] HCM Comps"}},
        fields="spreadsheetId,spreadsheetUrl,sheets(properties(sheetId,title,index))",
    )
    spreadsheets_api.batchUpdate.assert_called_once_with(
        spreadsheetId="new-sheet",
        body={
            "requests": [
                {"deleteSheet": {"sheetId": 900}},
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": 201,
                            "title": "Comps",
                            "index": 0,
                        },
                        "fields": "title,index",
                    }
                },
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": 202,
                            "title": "Assumptions",
                            "index": 1,
                        },
                        "fields": "title,index",
                    }
                },
            ]
        },
    )
    assert result == {
        "spreadsheet_id": "new-sheet",
        "title": "[hank] HCM Comps",
        "url": "https://docs.google.com/spreadsheets/d/new-sheet",
        "copied_tabs": ["Comps", "Assumptions"],
    }


def test_copy_spreadsheet_tabs_subset_preserves_source_order() -> None:
    sheets_service = MagicMock()
    spreadsheets_api = sheets_service.spreadsheets.return_value
    sheets_api = spreadsheets_api.sheets.return_value
    copied_sheet_ids: list[int] = []

    spreadsheets_api.get.return_value.execute.return_value = {
        "sheets": [
            {"properties": {"sheetId": 11, "title": "Comps", "index": 0}},
            {"properties": {"sheetId": 22, "title": "Summary", "index": 1}},
            {"properties": {"sheetId": 33, "title": "Raw Data", "index": 2}},
        ]
    }
    spreadsheets_api.create.return_value.execute.return_value = {
        "spreadsheetId": "new-sheet",
        "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/new-sheet",
        "sheets": [{"properties": {"sheetId": 90, "title": "Sheet1", "index": 0}}],
    }
    copied_props = iter(
        [
            {"sheetId": 111, "title": "Copy of Comps", "index": 1},
            {"sheetId": 222, "title": "Copy of Summary", "index": 2},
        ]
    )

    def fake_copy_to(*, spreadsheetId, sheetId, body):
        copied_sheet_ids.append(sheetId)
        request = MagicMock()
        request.execute.return_value = next(copied_props)
        return request

    sheets_api.copyTo.side_effect = fake_copy_to

    result = sheets_client.copy_spreadsheet(
        sheets_service,
        source_spreadsheet_id="source-sheet",
        new_title="[hank] HCM Comps",
        tabs=["Summary", "Comps"],
    )

    assert copied_sheet_ids == [11, 22]
    assert result["copied_tabs"] == ["Comps", "Summary"]
    batch_body = spreadsheets_api.batchUpdate.call_args.kwargs["body"]
    assert batch_body["requests"][1]["updateSheetProperties"]["properties"] == {
        "sheetId": 111,
        "title": "Comps",
        "index": 0,
    }
    assert batch_body["requests"][2]["updateSheetProperties"]["properties"] == {
        "sheetId": 222,
        "title": "Summary",
        "index": 1,
    }


def test_clear_range_calls_api() -> None:
    sheets_service = MagicMock()
    clear_execute = (
        sheets_service.spreadsheets.return_value.values.return_value.clear.return_value.execute
    )
    clear_execute.return_value = {"clearedRange": "Sheet1!A1:C10"}

    result = sheets_client.clear_sheet_range(
        sheets_service,
        spreadsheet_id="sheet-123",
        range_a1="Sheet1!A1:C10",
    )

    sheets_service.spreadsheets.return_value.values.return_value.clear.assert_called_once_with(
        spreadsheetId="sheet-123",
        range="Sheet1!A1:C10",
        body={},
    )
    assert result["clearedRange"] == "Sheet1!A1:C10"


def test_touch_sheet_range_reads_clears_and_rewrites_in_order() -> None:
    sheets_service = MagicMock()
    values_api = sheets_service.spreadsheets.return_value.values.return_value
    call_order: list[str] = []

    def _record_read():
        call_order.append("read")
        return {"values": [["=SF(\"AAPL\",\"income\",\"revenue\")", 42], ["foo", "bar"]]}

    def _record_clear():
        call_order.append("clear")
        return {"clearedRange": "Sheet1!A1:B2"}

    def _record_update():
        call_order.append("update")
        return {"updatedRange": "Sheet1!A1:B2", "updatedCells": 4}

    values_api.get.return_value.execute.side_effect = _record_read
    values_api.clear.return_value.execute.side_effect = _record_clear
    values_api.update.return_value.execute.side_effect = _record_update

    result = sheets_client.touch_sheet_range(
        sheets_service,
        spreadsheet_id="sheet-123",
        range_a1="Sheet1!A1:B2",
    )

    assert call_order == ["read", "clear", "update"]
    values_api.get.assert_called_once_with(
        spreadsheetId="sheet-123",
        range="Sheet1!A1:B2",
        valueRenderOption="FORMULA",
        dateTimeRenderOption="FORMATTED_STRING",
    )
    values_api.clear.assert_called_once_with(
        spreadsheetId="sheet-123",
        range="Sheet1!A1:B2",
        body={},
    )
    values_api.update.assert_called_once_with(
        spreadsheetId="sheet-123",
        range="Sheet1!A1:B2",
        valueInputOption="USER_ENTERED",
        body={"values": [["=SF(\"AAPL\",\"income\",\"revenue\")", 42], ["foo", "bar"]]},
    )
    assert result == {"touchedRange": "Sheet1!A1:B2", "touchedCells": 4}


def test_touch_sheet_range_empty_range_returns_without_writes() -> None:
    sheets_service = MagicMock()
    values_api = sheets_service.spreadsheets.return_value.values.return_value
    values_api.get.return_value.execute.return_value = {"values": []}

    result = sheets_client.touch_sheet_range(
        sheets_service,
        spreadsheet_id="sheet-123",
        range_a1="Sheet1!A1:B2",
    )

    values_api.get.assert_called_once_with(
        spreadsheetId="sheet-123",
        range="Sheet1!A1:B2",
        valueRenderOption="FORMULA",
        dateTimeRenderOption="FORMATTED_STRING",
    )
    values_api.clear.assert_not_called()
    values_api.update.assert_not_called()
    assert result == {"touchedRange": "Sheet1!A1:B2", "touchedCells": 0}
