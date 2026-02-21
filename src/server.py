"""Standalone MCP server for Google Sheets operations."""

import json

from mcp.server.fastmcp import FastMCP

from . import sheets_client

mcp = FastMCP(
    "gsheets-mcp",
    instructions=(
        "Google Sheets tools for tab listing, range reads/writes, search, "
        "sheet creation, and range clearing."
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
    """List tabs in a Google Sheets spreadsheet by name or spreadsheet ID."""
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
    """Read values from a range in a Google Sheets spreadsheet."""
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
    """Update a range in a Google Sheets spreadsheet using USER_ENTERED values."""
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
    """Append rows to a range in a Google Sheets spreadsheet."""
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
    """Create a new Google Sheets spreadsheet."""
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
def gsheet_search(query: str, max_results: int = 10) -> str:
    """Search Google Drive for spreadsheets by name."""
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
    """Clear all values in a range without deleting cells."""
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


def main() -> None:
    """Run the MCP server."""
    mcp.run()


if __name__ == "__main__":
    main()
