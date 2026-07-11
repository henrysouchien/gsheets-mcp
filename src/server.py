"""Standalone MCP server for Google Sheets operations."""

import json

from mcp.server.fastmcp import FastMCP

from . import sheets_client

mcp = FastMCP(
    "gsheets-mcp",
    instructions=(
        "Google Sheets tools for tab listing, range reads/writes, search, "
        "sheet creation/copying, range clearing, and range touch/recalc."
    ),
)


def _json_error(operation: str, error: Exception) -> str:
    """Return standardized JSON error payload for Sheets tools."""
    payload = {
        "status": "error",
        "error": str(error),
        "operation": operation,
    }
    if isinstance(error, sheets_client.SheetsClientError):
        payload["error_code"] = error.code
        payload.update(error.details)
    elif (
        sheets_client.is_broker_mode()
        and isinstance(error, sheets_client.HttpError)
        and getattr(getattr(error, "resp", None), "status", None) == 401
    ):
        payload["error"] = "Google Sheets authorization failed after one credential refresh."
        payload["error_code"] = "google_api_unauthorized"
    return json.dumps(payload)


def _validate_render_options(
    value_render_option: str,
    date_time_render_option: str,
) -> None:
    if value_render_option not in sheets_client.VALID_VALUE_RENDER_OPTIONS:
        raise ValueError(
            f"Invalid value_render_option '{value_render_option}'. "
            f"Must be one of: {sorted(sheets_client.VALID_VALUE_RENDER_OPTIONS)}"
        )
    if date_time_render_option not in sheets_client.VALID_DATETIME_RENDER_OPTIONS:
        raise ValueError(
            f"Invalid date_time_render_option '{date_time_render_option}'. "
            f"Must be one of: {sorted(sheets_client.VALID_DATETIME_RENDER_OPTIONS)}"
        )


@mcp.tool()
def gsheet_list_tabs(spreadsheet: str) -> str:
    """List tabs in a Google Sheets spreadsheet by URL or spreadsheet ID.

    In broker mode, pass a full Google Sheets URL or spreadsheet ID; titles
    require Drive scope and are not accepted. Local dev mode also accepts exact
    titles and can use gsheet_search. This is the safest first range-related call
    because it returns valid tab names.

    Sibling tools: use gsheet_read_range once you know the tab and A1 range.
    Use gsheet_update_range, gsheet_append_rows, gsheet_clear_range, or
    gsheet_touch_range only after confirming the target tab exists.

    Common mistake: spreadsheet can be a title or ID, but cell ranges are not
    accepted by this tool. Pass ranges to the range-specific tools.
    """
    try:
        spreadsheet_id, title = sheets_client.resolve_spreadsheet_id(spreadsheet)
        sheets_service = sheets_client.get_sheets_service()
        tabs = sheets_client.list_sheet_tabs(sheets_service, spreadsheet_id)
        return json.dumps(
            {
                "status": "ok",
                "spreadsheet_id": spreadsheet_id,
                "title": title,
                "tabs": tabs,
            }
        )
    except Exception as exc:
        return _json_error("gsheet_list_tabs", exc)


@mcp.tool()
def gsheet_read_range(
    spreadsheet: str,
    cell_range: str,
    value_render_option: str = "FORMATTED_VALUE",
    date_time_render_option: str = "FORMATTED_STRING",
) -> str:
    """Read values from a range in a Google Sheets spreadsheet.

    Pass a spreadsheet URL or ID and use gsheet_list_tabs to choose an existing
    tab before passing an A1-style range such as Sheet1!A1:D20. Local dev mode
    additionally supports titles and gsheet_search.

    Sibling tools: use gsheet_update_range to overwrite a range,
    gsheet_append_rows to add rows, gsheet_clear_range to remove values, and
    gsheet_touch_range to force recalculation without changing formulas.

    Common mistake: this returns values only. It does not expose formatting,
    formulas separately from rendered values, or Drive metadata.
    """
    try:
        _validate_render_options(value_render_option, date_time_render_option)
        spreadsheet_id, _ = sheets_client.resolve_spreadsheet_id(spreadsheet)
        sheets_service = sheets_client.get_sheets_service()
        values = sheets_client.read_sheet_range(
            sheets_service,
            spreadsheet_id,
            cell_range,
            value_render_option=value_render_option,
            date_time_render_option=date_time_render_option,
        )
        return json.dumps(
            {
                "status": "ok",
                "spreadsheet_id": spreadsheet_id,
                "range": cell_range,
                "value_render_option": value_render_option,
                "date_time_render_option": date_time_render_option,
                "values": values,
            }
        )
    except Exception as exc:
        return _json_error("gsheet_read_range", exc)


@mcp.tool()
def gsheet_update_range(spreadsheet: str, cell_range: str, values: list[list]) -> str:
    """Update a range in a Google Sheets spreadsheet using USER_ENTERED values.

    Pass a spreadsheet URL or ID, confirm its tab with gsheet_list_tabs, then
    inspect values before overwriting. Local dev mode additionally supports
    titles and gsheet_search. Values must be a two-dimensional row-major array.

    Sibling tools: use gsheet_append_rows when adding new rows below an
    existing table, gsheet_clear_range when removing values, and
    gsheet_touch_range when the goal is recalculation rather than replacement.

    Common mistake: this overwrites the target range. It does not append rows
    and it does not preserve a restore token; read the existing range first
    when the old values matter.
    """
    try:
        spreadsheet_id, _ = sheets_client.resolve_spreadsheet_id(spreadsheet)
        sheets_service = sheets_client.get_sheets_service()
        update_result = sheets_client.update_sheet_range(
            sheets_service,
            spreadsheet_id,
            cell_range,
            values,
            value_input_option="USER_ENTERED",
        )
        return json.dumps(
            {
                "status": "ok",
                "updatedRange": update_result.get("updatedRange", ""),
                "updatedCells": update_result.get("updatedCells", 0),
            }
        )
    except Exception as exc:
        return _json_error("gsheet_update_range", exc)


@mcp.tool()
def gsheet_append_rows(spreadsheet: str, cell_range: str, values: list[list]) -> str:
    """Append rows to a range in a Google Sheets spreadsheet.

    Pass a spreadsheet URL or ID, confirm its tab with gsheet_list_tabs, then
    read the table header/body to confirm the destination. Local dev mode also
    accepts titles and supports gsheet_search.

    Sibling tools: use gsheet_update_range for fixed-cell replacement and
    gsheet_clear_range for removing values. Use gsheet_list_tabs when an agent
    is unsure which tab owns the table.

    Common mistake: append placement follows Google Sheets table detection.
    For exact cell replacement, use gsheet_update_range instead.
    """
    try:
        spreadsheet_id, _ = sheets_client.resolve_spreadsheet_id(spreadsheet)
        sheets_service = sheets_client.get_sheets_service()
        append_result = sheets_client.append_sheet_rows(
            sheets_service,
            spreadsheet_id,
            cell_range,
            values,
            value_input_option="USER_ENTERED",
            insert_data_option="INSERT_ROWS",
        )
        return json.dumps(
            {
                "status": "ok",
                "updatedRange": append_result.get("updatedRange", ""),
                "updatedCells": append_result.get("updatedCells", 0),
            }
        )
    except Exception as exc:
        return _json_error("gsheet_append_rows", exc)


@mcp.tool()
def gsheet_create(title: str) -> str:
    """Create a new Google Sheets spreadsheet.

    Use this only when a new file is required. In local dev mode, gsheet_search
    can check whether a similarly titled spreadsheet already exists; broker
    mode intentionally has no Drive-wide search.

    Sibling tools: after creation, use gsheet_update_range or
    gsheet_append_rows to populate values and gsheet_list_tabs to inspect the
    default tab name.

    Common mistake: this creates a spreadsheet file, not a tab inside an
    existing spreadsheet.
    """
    try:
        sheets_service = sheets_client.get_sheets_service()
        spreadsheet_id, url = sheets_client.create_spreadsheet(sheets_service, title)
        return json.dumps(
            {
                "status": "ok",
                "spreadsheet_id": spreadsheet_id,
                "url": url,
            }
        )
    except Exception as exc:
        return _json_error("gsheet_create", exc)


@mcp.tool()
def gsheet_copy_spreadsheet(
    source: str,
    new_title: str,
    tabs: list[str] | None = None,
) -> str:
    """Copy all or selected tabs from one spreadsheet into a new spreadsheet.

    In broker mode, pass the source as a Google Sheets URL or spreadsheet ID.
    Local dev mode additionally accepts exact titles and supports gsheet_search.
    Use gsheet_list_tabs first when passing tabs so names match exactly.

    Sibling tools: after copying, use gsheet_update_range to swap inputs,
    gsheet_touch_range to force custom-function recalculation, and
    gsheet_read_range to inspect copied values or formulas.

    Output includes a warnings list when spreadsheet-level objects the
    tab-level copy cannot carry (named ranges, locale/time zone) would be
    lost; an empty list means no such gap was detected.

    Common mistake: this creates a new spreadsheet file. It does not duplicate
    a tab inside the source spreadsheet, and tabs must be tab names, not ranges.
    """
    try:
        source_spreadsheet_id, _ = sheets_client.resolve_spreadsheet_id(source)
        sheets_service = sheets_client.get_sheets_service()
        copy_result = sheets_client.copy_spreadsheet(
            sheets_service,
            source_spreadsheet_id,
            new_title,
            tabs=tabs,
        )
        return json.dumps(
            {
                "status": "ok",
                "spreadsheet_id": copy_result.get("spreadsheet_id", ""),
                "title": copy_result.get("title", new_title),
                "url": copy_result.get("url", ""),
                "copied_tabs": copy_result.get("copied_tabs", []),
                "warnings": copy_result.get("warnings", []),
            }
        )
    except Exception as exc:
        return _json_error("gsheet_copy_spreadsheet", exc)


@mcp.tool()
def gsheet_search(query: str, max_results: int = 10) -> str:
    """Search Google Drive for spreadsheets by name.

    Local dev mode can use this when the caller has a title or partial title;
    results include candidate IDs and names. Broker mode has spreadsheets scope
    only and returns requires_drive_scope; use a spreadsheet URL or ID instead.

    Sibling tools: pass the chosen result to gsheet_list_tabs before range
    operations, or directly to gsheet_read_range when the tab/range is already
    known.

    Common mistake: this searches spreadsheet files only. It does not search
    cell contents inside spreadsheets.
    """
    try:
        if sheets_client.is_broker_mode():
            raise sheets_client.SheetsClientError(
                "requires_drive_scope",
                "requires_drive_scope: gsheet_search is unavailable in broker mode; use a spreadsheet URL or ID.",
            )
        if max_results <= 0:
            raise ValueError("max_results must be > 0")
        drive_service = sheets_client.authenticate()
        files = sheets_client.search_spreadsheets(
            drive_service,
            query=query,
            max_results=max_results,
        )
        return json.dumps(
            {
                "status": "ok",
                "query": query,
                "results": files,
                "count": len(files),
            }
        )
    except Exception as exc:
        return _json_error("gsheet_search", exc)


@mcp.tool()
def gsheet_clear_range(spreadsheet: str, cell_range: str) -> str:
    """Clear all values in a range without deleting cells.

    Pass a spreadsheet URL or ID, confirm the tab with gsheet_list_tabs, then
    inspect values before clearing. Local dev mode additionally supports titles
    and gsheet_search. cell_range must be A1 notation such as Sheet1!A2:D20.

    Sibling tools: use gsheet_update_range to replace values,
    gsheet_append_rows to add rows, and gsheet_touch_range when recalculation
    is the goal. This tool removes values only; it does not delete rows,
    columns, tabs, or formatting.

    Common mistake: clear is not reversible through this MCP surface. Read the
    range first if the old values need to be preserved externally.
    """
    try:
        spreadsheet_id, _ = sheets_client.resolve_spreadsheet_id(spreadsheet)
        sheets_service = sheets_client.get_sheets_service()
        clear_result = sheets_client.clear_sheet_range(
            sheets_service,
            spreadsheet_id,
            cell_range,
        )
        return json.dumps(
            {
                "status": "ok",
                "spreadsheet_id": spreadsheet_id,
                "clearedRange": clear_result.get("clearedRange", cell_range),
            }
        )
    except Exception as exc:
        return _json_error("gsheet_clear_range", exc)


@mcp.tool()
def gsheet_touch_range(spreadsheet: str, cell_range: str) -> str:
    """Touch a range to force recalculation of custom functions.

    Pass a spreadsheet URL or ID, confirm its tab, then read the formula range.
    Local dev mode additionally supports titles and gsheet_search. The A1 range
    should target formulas that can be safely cleared and rewritten.

    Sibling tools: use gsheet_read_range for inspection, gsheet_update_range
    for intentional replacement, and gsheet_clear_range for value removal. This
    tool reads formulas, clears the range, then rewrites the formulas.

    Common mistake: do not use touch_range on plain input cells just to inspect
    values. It is a recalculation operation for formula/custom-function ranges.
    """
    try:
        spreadsheet_id, _ = sheets_client.resolve_spreadsheet_id(spreadsheet)
        sheets_service = sheets_client.get_sheets_service()
        touch_result = sheets_client.touch_sheet_range(
            sheets_service,
            spreadsheet_id,
            cell_range,
        )
        return json.dumps(
            {
                "status": "ok",
                "spreadsheet_id": spreadsheet_id,
                "touchedRange": touch_result.get("touchedRange", cell_range),
                "touchedCells": touch_result.get("touchedCells", 0),
            }
        )
    except Exception as exc:
        return _json_error("gsheet_touch_range", exc)


def main() -> None:
    """Run the MCP server."""
    mcp.run()


if __name__ == "__main__":
    main()
