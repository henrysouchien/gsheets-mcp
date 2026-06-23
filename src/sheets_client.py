"""Google Sheets and Drive helpers for gsheets-mcp."""

import logging
import os
import pickle
import re
from pathlib import Path

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

GOOGLE_SHEET_MIME = "application/vnd.google-apps.spreadsheet"
SPREADSHEET_ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]{20,}$")

VALID_VALUE_RENDER_OPTIONS = {
    "FORMATTED_VALUE",
    "UNFORMATTED_VALUE",
    "FORMULA",
}
VALID_DATETIME_RENDER_OPTIONS = {
    "SERIAL_NUMBER",
    "FORMATTED_STRING",
}

BASE_DIR = Path(__file__).parent.parent
CREDENTIALS_FILE = Path(
    os.environ.get("GOOGLE_CREDENTIALS_FILE") or BASE_DIR / "drive_credentials.json"
)
TOKEN_FILE = Path(os.environ.get("GOOGLE_TOKEN_FILE") or BASE_DIR / "token.pickle")

_cached_creds = None
logger = logging.getLogger(__name__)


def _get_missing_scopes(creds) -> list[str]:
    """Return required scopes that are missing from credentials."""
    granted = set()
    if getattr(creds, "scopes", None):
        granted.update(creds.scopes)
    if getattr(creds, "granted_scopes", None):
        granted.update(creds.granted_scopes)
    return [scope for scope in SCOPES if scope not in granted]


def _get_credentials():
    """Load, refresh, or create OAuth credentials with required scopes."""
    global _cached_creds

    creds = _cached_creds
    if creds is None and TOKEN_FILE.exists():
        with open(TOKEN_FILE, "rb") as token_file:
            creds = pickle.load(token_file)

    missing_scopes = _get_missing_scopes(creds) if creds else []
    if missing_scopes:
        creds = None
        _cached_creds = None
        if TOKEN_FILE.exists():
            TOKEN_FILE.unlink()

    should_save_token = False
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            should_save_token = True
        else:
            if not CREDENTIALS_FILE.exists():
                raise FileNotFoundError(
                    f"Credentials file not found at {CREDENTIALS_FILE}. "
                    "Please copy your drive_credentials.json to the gsheets-mcp folder."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), SCOPES
            )
            creds = flow.run_local_server(port=0)
            should_save_token = True

    if should_save_token:
        with open(TOKEN_FILE, "wb") as token_file:
            pickle.dump(creds, token_file)

    _cached_creds = creds
    return creds


def authenticate():
    """Authenticate with Google Drive API and return a service object."""
    creds = _get_credentials()
    return build("drive", "v3", credentials=creds)


def get_sheets_service():
    """Authenticate with Google Sheets API and return a service object."""
    creds = _get_credentials()
    return build("sheets", "v4", credentials=creds)


def resolve_spreadsheet_id(name_or_id: str) -> tuple[str, str]:
    """Resolve a spreadsheet by ID or exact name and return (id, title)."""
    drive_service = authenticate()

    if SPREADSHEET_ID_PATTERN.match(name_or_id):
        try:
            file_info = drive_service.files().get(
                fileId=name_or_id,
                supportsAllDrives=True,
                fields="id, name, mimeType",
            ).execute()
            if file_info.get("mimeType") == GOOGLE_SHEET_MIME:
                return file_info["id"], file_info["name"]
        except HttpError as exc:
            if getattr(exc, "resp", None) is None or exc.resp.status != 404:
                raise

    escaped_name = name_or_id.replace("'", "\\'")
    query = (
        f"name = '{escaped_name}' and "
        f"mimeType = '{GOOGLE_SHEET_MIME}' and "
        "trashed = false"
    )
    results = drive_service.files().list(
        q=query,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        fields="files(id, name, modifiedTime, webViewLink, parents)",
    ).execute()
    files = results.get("files", [])

    if not files:
        raise ValueError(f"Spreadsheet not found: {name_or_id}")

    if len(files) == 1:
        file_info = files[0]
        return file_info["id"], file_info["name"]

    parent_ids: set[str] = set()
    for file_info in files:
        for parent_id in file_info.get("parents", []):
            parent_ids.add(parent_id)

    parent_names: dict[str, str] = {}
    for parent_id in parent_ids:
        try:
            parent_info = drive_service.files().get(
                fileId=parent_id,
                supportsAllDrives=True,
                fields="id, name",
            ).execute()
            parent_names[parent_id] = parent_info.get("name", "")
        except HttpError:
            parent_names[parent_id] = ""

    candidates = [
        "Multiple spreadsheets found. Use spreadsheet ID instead. Candidates:"
    ]
    for file_info in files:
        parent_name = ""
        parent_list = file_info.get("parents", [])
        if parent_list:
            parent_name = parent_names.get(parent_list[0], "")
        candidates.append(
            f"- id: {file_info.get('id', '')}, "
            f"name: {file_info.get('name', '')}, "
            f"modifiedTime: {file_info.get('modifiedTime', '')}, "
            f"webViewLink: {file_info.get('webViewLink', '')}, "
            f"parent: {parent_name}"
        )

    raise ValueError("\n".join(candidates))


def _list_sheet_properties(sheets_service, spreadsheet_id: str) -> list[dict]:
    """List tab properties in a spreadsheet, including sheet IDs."""
    spreadsheet = sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields=(
            "sheets(properties(sheetId,title,index,"
            "gridProperties(rowCount,columnCount)))"
        ),
    ).execute()

    tabs = []
    for sheet in spreadsheet.get("sheets", []):
        props = sheet.get("properties", {})
        grid = props.get("gridProperties", {})
        tabs.append(
            {
                "sheetId": props.get("sheetId", 0),
                "title": props.get("title", ""),
                "index": props.get("index", 0),
                "rowCount": grid.get("rowCount", 0),
                "columnCount": grid.get("columnCount", 0),
            }
        )
    return tabs


def list_sheet_tabs(sheets_service, spreadsheet_id: str) -> list[dict]:
    """List tabs in a spreadsheet."""
    tabs = []
    for tab in _list_sheet_properties(sheets_service, spreadsheet_id):
        tabs.append(
            {
                "title": tab["title"],
                "index": tab["index"],
                "rowCount": tab["rowCount"],
                "columnCount": tab["columnCount"],
            }
        )
    return tabs


def read_sheet_range(
    sheets_service,
    spreadsheet_id: str,
    range_a1: str,
    value_render_option: str = "FORMATTED_VALUE",
    date_time_render_option: str = "FORMATTED_STRING",
) -> list[list]:
    """Read values from a spreadsheet range with render option controls."""
    if value_render_option not in VALID_VALUE_RENDER_OPTIONS:
        raise ValueError(
            f"Invalid value_render_option '{value_render_option}'. "
            f"Must be one of: {sorted(VALID_VALUE_RENDER_OPTIONS)}"
        )
    if date_time_render_option not in VALID_DATETIME_RENDER_OPTIONS:
        raise ValueError(
            f"Invalid date_time_render_option '{date_time_render_option}'. "
            f"Must be one of: {sorted(VALID_DATETIME_RENDER_OPTIONS)}"
        )

    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_a1,
        valueRenderOption=value_render_option,
        dateTimeRenderOption=date_time_render_option,
    ).execute()
    return result.get("values", [])


def update_sheet_range(
    sheets_service,
    spreadsheet_id: str,
    range_a1: str,
    values: list[list],
    value_input_option: str = "USER_ENTERED",
) -> dict:
    """Update a spreadsheet range with values."""
    if not isinstance(values, list) or not values or any(
        not isinstance(row, list) for row in values
    ):
        raise ValueError("values must be a non-empty list of lists")

    result = sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_a1,
        valueInputOption=value_input_option,
        body={"values": values},
    ).execute()
    return {
        "updatedCells": result.get("updatedCells", 0),
        "updatedRange": result.get("updatedRange", ""),
    }


def append_sheet_rows(
    sheets_service,
    spreadsheet_id: str,
    range_a1: str,
    values: list[list],
    value_input_option: str = "USER_ENTERED",
    insert_data_option: str = "INSERT_ROWS",
) -> dict:
    """Append rows to a spreadsheet range."""
    if not isinstance(values, list) or not values or any(
        not isinstance(row, list) for row in values
    ):
        raise ValueError("values must be a non-empty list of lists")

    result = sheets_service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=range_a1,
        valueInputOption=value_input_option,
        insertDataOption=insert_data_option,
        body={"values": values},
    ).execute()
    updates = result.get("updates", {})
    return {
        "updatedCells": updates.get("updatedCells", 0),
        "updatedRange": updates.get("updatedRange", ""),
    }


def create_spreadsheet(sheets_service, title: str) -> tuple[str, str]:
    """Create a new spreadsheet and return (spreadsheet_id, spreadsheet_url)."""
    result = sheets_service.spreadsheets().create(
        body={"properties": {"title": title}},
        fields="spreadsheetId,spreadsheetUrl",
    ).execute()
    return result.get("spreadsheetId", ""), result.get("spreadsheetUrl", "")


def copy_spreadsheet(
    sheets_service,
    source_spreadsheet_id: str,
    new_title: str,
    tabs: list[str] | None = None,
) -> dict:
    """Copy selected tabs into a new spreadsheet with the original tab names."""
    if not new_title:
        raise ValueError("new_title must be non-empty")

    requested_tabs: set[str] | None = None
    if tabs is not None:
        if (
            not isinstance(tabs, list)
            or not tabs
            or any(not isinstance(tab, str) or not tab for tab in tabs)
        ):
            raise ValueError("tabs must be a non-empty list of tab names when provided")
        duplicate_tabs = sorted({tab for tab in tabs if tabs.count(tab) > 1})
        if duplicate_tabs:
            raise ValueError(
                f"tabs contains duplicate tab names: {', '.join(duplicate_tabs)}"
            )
        requested_tabs = set(tabs)

    source_tabs = _list_sheet_properties(sheets_service, source_spreadsheet_id)
    if not source_tabs:
        raise ValueError("Source spreadsheet has no tabs")

    if requested_tabs is None:
        selected_tabs = source_tabs
    else:
        selected_tabs = [tab for tab in source_tabs if tab["title"] in requested_tabs]
        selected_titles = {tab["title"] for tab in selected_tabs}
        missing_tabs = [tab for tab in tabs or [] if tab not in selected_titles]
        if missing_tabs:
            raise ValueError(f"Requested tabs not found: {', '.join(missing_tabs)}")

    create_result = sheets_service.spreadsheets().create(
        body={"properties": {"title": new_title}},
        fields="spreadsheetId,spreadsheetUrl,sheets(properties(sheetId,title,index))",
    ).execute()
    new_spreadsheet_id = create_result.get("spreadsheetId", "")
    spreadsheet_url = create_result.get("spreadsheetUrl", "")
    if not new_spreadsheet_id:
        raise ValueError("Sheets API did not return a new spreadsheet ID")

    default_sheet_id = None
    for sheet in create_result.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("title") == "Sheet1":
            default_sheet_id = props.get("sheetId")
            break
    if default_sheet_id is None:
        raise ValueError("Created spreadsheet default Sheet1 tab was not found")

    copied_tabs = []
    for source_tab in selected_tabs:
        copied_props = sheets_service.spreadsheets().sheets().copyTo(
            spreadsheetId=source_spreadsheet_id,
            sheetId=source_tab["sheetId"],
            body={"destinationSpreadsheetId": new_spreadsheet_id},
        ).execute()
        copied_sheet_id = copied_props.get("sheetId")
        if copied_sheet_id is None:
            raise ValueError(
                f"Sheets API did not return copied sheet ID for tab {source_tab['title']}"
            )
        copied_tabs.append(
            {
                "source_title": source_tab["title"],
                "copied_sheet_id": copied_sheet_id,
            }
        )

    batch_requests = [{"deleteSheet": {"sheetId": default_sheet_id}}]
    for index, copied_tab in enumerate(copied_tabs):
        batch_requests.append(
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": copied_tab["copied_sheet_id"],
                        "title": copied_tab["source_title"],
                        "index": index,
                    },
                    "fields": "title,index",
                }
            }
        )

    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=new_spreadsheet_id,
        body={"requests": batch_requests},
    ).execute()

    return {
        "spreadsheet_id": new_spreadsheet_id,
        "title": new_title,
        "url": spreadsheet_url,
        "copied_tabs": [copied_tab["source_title"] for copied_tab in copied_tabs],
    }


def search_spreadsheets(
    drive_service,
    query: str,
    max_results: int = 10,
) -> list[dict]:
    """Search Google Drive for spreadsheets by name."""
    escaped_query = query.replace("'", "\\'")
    search_query = (
        f"name contains '{escaped_query}' and "
        f"mimeType = '{GOOGLE_SHEET_MIME}' and "
        "trashed = false"
    )
    results = drive_service.files().list(
        q=search_query,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        fields="files(id, name, modifiedTime, webViewLink)",
        pageSize=max_results,
    ).execute()
    return results.get("files", [])


def clear_sheet_range(sheets_service, spreadsheet_id: str, range_a1: str) -> dict:
    """Clear all values in a spreadsheet range."""
    result = sheets_service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=range_a1,
        body={},
    ).execute()
    return {
        "clearedRange": result.get("clearedRange", range_a1),
    }


def touch_sheet_range(sheets_service, spreadsheet_id: str, range_a1: str) -> dict:
    """Force custom-function recalculation by clearing and rewriting a range."""
    original_values = read_sheet_range(
        sheets_service,
        spreadsheet_id,
        range_a1,
        value_render_option="FORMULA",
    )
    if not original_values:
        return {"touchedRange": range_a1, "touchedCells": 0}

    clear_sheet_range(sheets_service, spreadsheet_id, range_a1)
    try:
        updated = update_sheet_range(
            sheets_service,
            spreadsheet_id,
            range_a1,
            original_values,
            value_input_option="USER_ENTERED",
        )
    except Exception:
        logger.warning(
            "touch_sheet_range update failed after clear; spreadsheet_id=%s range=%s original_values=%r",
            spreadsheet_id,
            range_a1,
            original_values,
        )
        raise

    return {
        "touchedRange": updated.get("updatedRange", range_a1),
        "touchedCells": updated.get("updatedCells", 0),
    }
