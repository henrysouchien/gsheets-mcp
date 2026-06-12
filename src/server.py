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
    return json.dumps(
        {
            "status": "error",
            "error": str(error),
            "operation": operation,
        }
    )


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
    """List tabs in a Google Sheets spreadsheet by name or spreadsheet ID.

    Discovery: run gsheet_search first when you only know a spreadsheet title
    or partial name, then pass the returned spreadsheet_id or exact name here.
    This is the safest first call before any range read/write because it returns
    valid tab names.

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

    Discovery: run gsheet_search to resolve the spreadsheet and
    gsheet_list_tabs to choose an existing tab before passing an A1-style
    cell_range such as Sheet1!A1:D20. Use value_render_option and
    date_time_render_option only with Google Sheets API-supported values.

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

    Discovery: run gsheet_search and gsheet_list_tabs first to confirm the
    spreadsheet and tab, then gsheet_read_range to inspect the current values
    before overwriting. cell_range must be A1 notation and values must be a
    two-dimensional row-major array.

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

    Discovery: run gsheet_search and gsheet_list_tabs first, then
    gsheet_read_range on the table header/body to confirm the destination
    range. cell_range should point at the table or sheet area where Google
    Sheets should append rows, and values must be a two-dimensional row array.

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

    Discovery: run gsheet_search first if there may already be a spreadsheet
    with the requested title. Use this only when a new file is required.

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

    Discovery: run gsheet_search first when you only know a spreadsheet title
    or partial name, then pass the returned spreadsheet_id or exact name as
    source. Use gsheet_list_tabs first when passing tabs so the names match
    exactly.

    Sibling tools: after copying, use gsheet_update_range to swap inputs,
    gsheet_touch_range to force custom-function recalculation, and
    gsheet_read_range to inspect copied values or formulas.

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
            }
        )
    except Exception as exc:
        return _json_error("gsheet_copy_spreadsheet", exc)


@mcp.tool()
def gsheet_search(query: str, max_results: int = 10) -> str:
    """Search Google Drive for spreadsheets by name.

    Discovery: use this as the first step when the caller has a spreadsheet
    title, partial title, or human description instead of a spreadsheet_id.
    Results include candidate IDs and names for follow-up calls.

    Sibling tools: pass the chosen result to gsheet_list_tabs before range
    operations, or directly to gsheet_read_range when the tab/range is already
    known.

    Common mistake: this searches spreadsheet files only. It does not search
    cell contents inside spreadsheets.
    """
    try:
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

    Discovery: run gsheet_search and gsheet_list_tabs to confirm the file and
    tab, then gsheet_read_range to inspect the current values before clearing.
    cell_range must be A1 notation such as Sheet1!A2:D20.

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

    Discovery: run gsheet_search and gsheet_list_tabs first, then
    gsheet_read_range to confirm the formula range. cell_range must be A1
    notation and should target formulas that can be safely cleared and
    rewritten to trigger recalculation.

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
