"""Google Sheets and Drive helpers for gsheets-mcp."""

import logging
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
CREDENTIALS_FILE = BASE_DIR / "drive_credentials.json"
TOKEN_FILE = BASE_DIR / "token.pickle"

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


def list_sheet_tabs(sheets_service, spreadsheet_id: str) -> list[dict]:
    """List tabs in a spreadsheet."""
    spreadsheet = sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(title,index,gridProperties(rowCount,columnCount)))",
    ).execute()

    tabs = []
    for sheet in spreadsheet.get("sheets", []):
        props = sheet.get("properties", {})
        grid = props.get("gridProperties", {})
        tabs.append(
            {
                "title": props.get("title", ""),
                "index": props.get("index", 0),
                "rowCount": grid.get("rowCount", 0),
                "columnCount": grid.get("columnCount", 0),
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
