"""Synchronous Google Sheets tool implementations behind the MCP boundary."""

from __future__ import annotations

from typing import Callable

from . import sheets_client
from .contracts import (
    AppendRowsInput,
    AppendRowsSuccess,
    ClearRangeInput,
    ClearRangeSuccess,
    CopySpreadsheetInput,
    CopySpreadsheetSuccess,
    CreateSpreadsheetInput,
    CreateSpreadsheetSuccess,
    ListTabsInput,
    ListTabsSuccess,
    ReadRangeInput,
    ReadRangeSuccess,
    RecalculateRangeInput,
    RecalculateRangeSuccess,
    SearchSpreadsheetsInput,
    SearchSpreadsheetsSuccess,
    StrictModel,
    WriteRangeInput,
    WriteRangeSuccess,
)


def _sheets_service():
    try:
        return sheets_client.get_sheets_service()
    except sheets_client.SheetsClientError:
        raise
    except Exception:
        raise sheets_client.SheetsClientError(
            "sheets_service_unavailable",
            "The Google Sheets client could not be initialized.",
            outcome_state="not_started",
            phase="initialize_client",
            mutation_may_have_occurred=False,
            retry_safe=True,
            retry_automatic=False,
            retry_action="retry_after_inspection",
        ) from None


def _resolved_sheets(spreadsheet: str):
    spreadsheet_id = sheets_client.parse_spreadsheet_reference(spreadsheet)
    service = _sheets_service()
    spreadsheet_id, title = sheets_client.resolve_spreadsheet(service, spreadsheet_id)
    return service, spreadsheet_id, title


def list_tabs(arguments: ListTabsInput) -> ListTabsSuccess:
    service, spreadsheet_id, title = _resolved_sheets(arguments.spreadsheet)
    return ListTabsSuccess(
        spreadsheet=spreadsheet_id,
        title=title,
        tabs=sheets_client.list_sheet_tabs(service, spreadsheet_id),
    )


def read_range(arguments: ReadRangeInput) -> ReadRangeSuccess:
    service, spreadsheet_id, _title = _resolved_sheets(arguments.spreadsheet)
    values = sheets_client.read_sheet_range(
        service,
        spreadsheet_id,
        arguments.range,
        value_render_option=arguments.value_render_option,
        date_time_render_option=arguments.date_time_render_option,
    )
    return ReadRangeSuccess(
        spreadsheet=spreadsheet_id,
        range=arguments.range,
        value_render_option=arguments.value_render_option,
        date_time_render_option=arguments.date_time_render_option,
        values=values,
    )


def write_range(arguments: WriteRangeInput) -> WriteRangeSuccess:
    service, spreadsheet_id, _title = _resolved_sheets(arguments.spreadsheet)
    try:
        result = sheets_client.update_sheet_range(
            service,
            spreadsheet_id,
            arguments.range,
            arguments.values,
            value_input_option="USER_ENTERED",
        )
    except Exception as exc:
        if isinstance(exc, sheets_client.SheetsClientError):
            raise
        raise sheets_client.mutation_failure(
            exc,
            phase="write_range",
            dispatched=True,
        ) from None
    return WriteRangeSuccess(
        spreadsheet=spreadsheet_id,
        range=result["range"],
        cell_count=result["cell_count"],
    )


def append_rows(arguments: AppendRowsInput) -> AppendRowsSuccess:
    service, spreadsheet_id, _title = _resolved_sheets(arguments.spreadsheet)
    try:
        result = sheets_client.append_sheet_rows(
            service,
            spreadsheet_id,
            arguments.range,
            arguments.values,
            value_input_option="USER_ENTERED",
            insert_data_option="INSERT_ROWS",
        )
    except Exception as exc:
        if isinstance(exc, sheets_client.SheetsClientError):
            raise
        raise sheets_client.mutation_failure(
            exc,
            phase="append_rows",
            dispatched=True,
        ) from None
    return AppendRowsSuccess(
        spreadsheet=spreadsheet_id,
        range=result["range"],
        cell_count=result["cell_count"],
    )


def create_spreadsheet(arguments: CreateSpreadsheetInput) -> CreateSpreadsheetSuccess:
    service = _sheets_service()
    try:
        result = sheets_client.create_spreadsheet(service, arguments.title)
    except Exception as exc:
        if isinstance(exc, sheets_client.SheetsClientError):
            raise
        raise sheets_client.mutation_failure(
            exc,
            phase="create_spreadsheet",
            dispatched=True,
        ) from None
    if not result["spreadsheet"]:
        raise sheets_client.SheetsClientError(
            "created_spreadsheet_identity_missing",
            "Google Sheets did not return the created spreadsheet ID.",
            outcome_state="uncertain",
            phase="create_spreadsheet",
            mutation_may_have_occurred=True,
            retry_safe=False,
            retry_automatic=False,
            retry_action="inspect_drive_before_retry",
        )
    return CreateSpreadsheetSuccess(**result)


def copy_spreadsheet(arguments: CopySpreadsheetInput) -> CopySpreadsheetSuccess:
    service, source_spreadsheet_id, _title = _resolved_sheets(arguments.spreadsheet)
    result = sheets_client.copy_spreadsheet(
        service,
        source_spreadsheet_id,
        arguments.title,
        tabs=arguments.tabs,
    )
    return CopySpreadsheetSuccess(**result)


def search_spreadsheets(
    arguments: SearchSpreadsheetsInput,
) -> SearchSpreadsheetsSuccess:
    if sheets_client.is_broker_mode():
        raise sheets_client.SheetsClientError(
            "capability_unavailable",
            "Spreadsheet search is not available in broker mode.",
            outcome_state="not_started",
            phase="capability_check",
            mutation_may_have_occurred=False,
            retry_safe=False,
            retry_automatic=False,
            retry_action="use_spreadsheet_id_or_url",
        )
    drive_service = sheets_client.authenticate()
    results = sheets_client.search_spreadsheets(
        drive_service,
        query=arguments.query,
        limit=arguments.limit,
    )
    return SearchSpreadsheetsSuccess(
        query=arguments.query,
        count=len(results),
        results=results,
    )


def clear_range(arguments: ClearRangeInput) -> ClearRangeSuccess:
    service, spreadsheet_id, _title = _resolved_sheets(arguments.spreadsheet)
    try:
        result = sheets_client.clear_sheet_range(
            service,
            spreadsheet_id,
            arguments.range,
        )
    except Exception as exc:
        if isinstance(exc, sheets_client.SheetsClientError):
            raise
        raise sheets_client.mutation_failure(
            exc,
            phase="clear_range",
            dispatched=True,
        ) from None
    return ClearRangeSuccess(
        spreadsheet=spreadsheet_id,
        range=result["range"],
    )


def recalculate_range(arguments: RecalculateRangeInput) -> RecalculateRangeSuccess:
    service, spreadsheet_id, _title = _resolved_sheets(arguments.spreadsheet)
    result = sheets_client.recalculate_sheet_range(
        service,
        spreadsheet_id,
        arguments.range,
    )
    return RecalculateRangeSuccess(
        spreadsheet=spreadsheet_id,
        **result,
    )


HANDLERS: dict[str, Callable[[StrictModel], StrictModel]] = {
    "gsheets_list_tabs": list_tabs,
    "gsheets_read_range": read_range,
    "gsheets_write_range": write_range,
    "gsheets_append_rows": append_rows,
    "gsheets_create_spreadsheet": create_spreadsheet,
    "gsheets_copy_spreadsheet": copy_spreadsheet,
    "gsheets_search_spreadsheets": search_spreadsheets,
    "gsheets_clear_range": clear_range,
    "gsheets_recalculate_range": recalculate_range,
}
