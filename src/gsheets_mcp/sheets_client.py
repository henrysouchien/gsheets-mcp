"""Google Sheets and Drive helpers for gsheets-mcp."""

import json
import math
import os
import pickle
import re
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib import error as urllib_error
from urllib import request as urllib_request
from urllib.parse import urljoin, urlparse

import google_auth_httplib2
from google.auth.credentials import Credentials as GoogleAuthCredentials
from google.auth.transport.requests import Request as GoogleAuthRequest
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError as HttpError
from googleapiclient.http import build_http

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

GOOGLE_SHEET_MIME = "application/vnd.google-apps.spreadsheet"
SPREADSHEET_ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]{20,}$")
# Real Sheets URLs may carry an account-qualifier segment (/u/<n>/) between
# /spreadsheets/ and /d/ when the browser session has multiple Google accounts.
# The suffix after the ID is anchored to the known UI endpoints so
# traversal-shaped paths (/d/<id>/../d/<other>) cannot pass validation.
SPREADSHEET_URL_PATH_PATTERN = re.compile(
    r"^/spreadsheets/(?:u/\d+/)?d/([A-Za-z0-9_-]+)"
    r"(?:/(?:edit|view|preview|copy|htmlview))?/?$"
)
BROKER_REFRESH_MARGIN_SECONDS = 60

VALID_VALUE_RENDER_OPTIONS = {
    "FORMATTED_VALUE",
    "UNFORMATTED_VALUE",
    "FORMULA",
}
VALID_DATETIME_RENDER_OPTIONS = {
    "SERIAL_NUMBER",
    "FORMATTED_STRING",
}

BASE_DIR = Path(__file__).resolve().parents[2]
CREDENTIALS_FILE = Path(
    os.environ.get("GOOGLE_CREDENTIALS_FILE") or BASE_DIR / "drive_credentials.json"
)
TOKEN_FILE = Path(os.environ.get("GOOGLE_TOKEN_FILE") or BASE_DIR / "token.pickle")

_cached_creds = None
_cached_broker_expires_at = 0


class SheetsClientError(RuntimeError):
    """Typed, user-safe error surfaced by MCP tools."""

    def __init__(self, code: str, message: str, **details) -> None:
        super().__init__(message)
        self.code = code
        self.details = details


def token_mode() -> str:
    """Return the validated immutable credential mode for this process."""
    raw_mode = os.environ.get("GSHEETS_TOKEN_MODE", "").strip().lower()
    if raw_mode in {"", "local"}:
        return "local"
    if raw_mode == "broker":
        return "broker"
    raise SheetsClientError(
        "invalid_configuration",
        "GSHEETS_TOKEN_MODE must be either 'local' or 'broker'.",
        outcome_state="not_started",
        phase="configuration",
        mutation_may_have_occurred=False,
        retry_safe=False,
        retry_action="correct_configuration",
    )


def is_broker_mode() -> bool:
    return token_mode() == "broker"


def invalidate_broker_credentials() -> None:
    """Discard only in-memory broker credentials after a Google 401."""
    global _cached_creds, _cached_broker_expires_at
    _cached_creds = None
    _cached_broker_expires_at = 0


def _broker_error(code: str, payload: dict | None = None) -> SheetsClientError:
    payload = payload or {}
    if code == "sheets_not_connected":
        return SheetsClientError(
            code,
            "Google Sheets is not connected; connect Google Sheets in Hank settings.",
            outcome_state="not_started",
            phase="broker_credentials",
            mutation_may_have_occurred=False,
            retry_safe=True,
            retry_automatic=False,
            retry_action="connect_google_sheets",
        )
    if code == "broker_rate_limited":
        raw_retry_after = payload.get("retry_after_s")
        retry_after = (
            raw_retry_after
            if not isinstance(raw_retry_after, bool)
            and isinstance(raw_retry_after, (int, float))
            and math.isfinite(raw_retry_after)
            and raw_retry_after >= 0
            else None
        )
        return SheetsClientError(
            code,
            "The Google Sheets token broker rate limited this call.",
            retry_after_s=retry_after,
            outcome_state="not_started",
            phase="broker_credentials",
            mutation_may_have_occurred=False,
            retry_safe=True,
            retry_automatic=False,
            retry_action="retry_later",
        )
    if code in {"auth_failed", "broker_session_expired"}:
        return SheetsClientError(
            "broker_session_expired",
            "The Google Sheets broker session expired and the child process must be replaced.",
            outcome_state="not_started",
            phase="broker_credentials",
            mutation_may_have_occurred=False,
            retry_safe=True,
            retry_action="replace_child_process",
        )
    return SheetsClientError(
        "sheets_unavailable",
        "Google Sheets is temporarily unavailable.",
        outcome_state="not_started",
        phase="broker_credentials",
        mutation_may_have_occurred=False,
        retry_safe=True,
        retry_automatic=False,
        retry_action="retry_later",
    )


def _fetch_broker_token() -> tuple[str, int]:
    base_url = os.environ.get("GSHEETS_BROKER_URL", "").strip()
    session_token = os.environ.get("GSHEETS_BROKER_SESSION_TOKEN", "").strip()
    if not base_url or not session_token:
        raise SheetsClientError(
            "sheets_unavailable",
            "Google Sheets broker configuration is unavailable.",
            outcome_state="not_started",
            phase="broker_configuration",
            mutation_may_have_occurred=False,
            retry_safe=False,
            retry_automatic=False,
            retry_action="correct_configuration",
        )
    endpoint = urljoin(
        base_url.rstrip("/") + "/", "api/internal/google/sheets-access-token"
    )
    req = urllib_request.Request(
        endpoint,
        data=b"{}",
        headers={
            "Authorization": f"Bearer {session_token}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib_request.urlopen(req, timeout=10) as response:
            payload = json.loads(response.read())
    except urllib_error.HTTPError as exc:
        try:
            payload = json.loads(exc.read())
        except (ValueError, TypeError):
            payload = {}
        if not isinstance(payload, dict):
            payload = {}
        raise _broker_error(
            str(payload.get("error") or "sheets_unavailable"), payload
        ) from None
    except (urllib_error.URLError, TimeoutError, OSError, ValueError):
        raise _broker_error("sheets_unavailable") from None

    token = payload.get("access_token") if isinstance(payload, dict) else None
    expires_at = payload.get("expires_at") if isinstance(payload, dict) else None
    if (
        not isinstance(token, str)
        or not token
        or isinstance(expires_at, bool)
        or not isinstance(expires_at, (int, float))
        or not math.isfinite(expires_at)
    ):
        raise _broker_error("sheets_unavailable")
    return token, int(expires_at)


class BrokerCredentials(GoogleAuthCredentials):
    """In-memory credentials refreshed exclusively through the token broker."""

    def __init__(self) -> None:
        super().__init__()
        self.expiry = None

    def refresh(self, request) -> None:
        del request  # The broker uses its own bounded HTTP request.
        global _cached_broker_expires_at
        token, expires_at = _fetch_broker_token()
        self.token = token
        # google-auth compares expiry with a naive UTC datetime.
        self.expiry = datetime.fromtimestamp(expires_at, tz=timezone.utc).replace(
            tzinfo=None
        )
        _cached_broker_expires_at = expires_at


def _get_missing_scopes(creds) -> list[str]:
    """Return required scopes that are missing from credentials."""
    granted = set()
    if getattr(creds, "scopes", None):
        granted.update(creds.scopes)
    if getattr(creds, "granted_scopes", None):
        granted.update(creds.granted_scopes)
    return [scope for scope in SCOPES if scope not in granted]


def _run_installed_app_flow():
    """Run local-dev interactive consent; never called by broker mode."""
    if os.environ.get("GSHEETS_HEADLESS") == "1":
        raise SheetsClientError(
            "interactive_consent_disabled",
            "Interactive Google consent is disabled in headless mode.",
            outcome_state="not_started",
            phase="local_credentials",
            mutation_may_have_occurred=False,
            retry_safe=False,
            retry_automatic=False,
            retry_action="complete_local_consent",
        )
    if not CREDENTIALS_FILE.exists():
        raise SheetsClientError(
            "local_credentials_missing",
            "Google OAuth desktop credentials are not configured.",
            outcome_state="not_started",
            phase="local_credentials",
            mutation_may_have_occurred=False,
            retry_safe=False,
            retry_automatic=False,
            retry_action="configure_google_credentials",
        )
    flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
    return flow.run_local_server(port=0)


def _get_credentials():
    """Load, refresh, or create OAuth credentials with required scopes."""
    global _cached_creds, _cached_broker_expires_at

    if is_broker_mode():
        if _cached_creds is None:
            _cached_creds = BrokerCredentials()
        if (
            _cached_broker_expires_at - int(time.time())
            <= BROKER_REFRESH_MARGIN_SECONDS
        ):
            _cached_creds.refresh(None)
        return _cached_creds

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
            creds.refresh(GoogleAuthRequest())
            should_save_token = True
        else:
            creds = _run_installed_app_flow()
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
    if is_broker_mode():
        authed_http = google_auth_httplib2.AuthorizedHttp(
            creds,
            http=build_http(),
            max_refresh_attempts=1,
        )
        return build("sheets", "v4", http=authed_http)
    return build("sheets", "v4", credentials=creds)


def parse_spreadsheet_reference(spreadsheet: str) -> str:
    """Parse a strict Sheets URL or ID without loading credentials."""
    value = spreadsheet.strip()
    parsed = urlparse(value)
    if parsed.scheme or parsed.netloc:
        path_match = SPREADSHEET_URL_PATH_PATTERN.fullmatch(parsed.path)
        try:
            port = parsed.port
        except ValueError:
            port = -1
        if (
            parsed.scheme != "https"
            or parsed.hostname != "docs.google.com"
            or parsed.username is not None
            or parsed.password is not None
            or port is not None
            or path_match is None
        ):
            raise SheetsClientError(
                "invalid_spreadsheet_reference",
                "Provide a Google Sheets spreadsheet ID or a strict docs.google.com Sheets URL.",
                outcome_state="not_started",
                phase="reference_validation",
                mutation_may_have_occurred=False,
                retry_safe=True,
                retry_action="correct_arguments",
            )
        value = path_match.group(1)

    if not SPREADSHEET_ID_PATTERN.fullmatch(value):
        raise SheetsClientError(
            "invalid_spreadsheet_reference",
            "Provide a Google Sheets spreadsheet ID or a strict docs.google.com Sheets URL; titles are not accepted.",
            outcome_state="not_started",
            phase="reference_validation",
            mutation_may_have_occurred=False,
            retry_safe=True,
            retry_action="correct_arguments",
        )
    return value


def resolve_spreadsheet(sheets_service, spreadsheet: str) -> tuple[str, str]:
    """Verify a Sheets URL/ID and return its normalized ID and title."""
    spreadsheet_id = parse_spreadsheet_reference(spreadsheet)
    try:
        metadata = (
            sheets_service.spreadsheets()
            .get(
                spreadsheetId=spreadsheet_id,
                fields="spreadsheetId,properties(title)",
            )
            .execute()
        )
    except Exception as exc:
        if isinstance(exc, SheetsClientError):
            raise
        raise mutation_failure(
            exc,
            phase="verify_spreadsheet",
            dispatched=False,
        ) from None
    return (
        metadata.get("spreadsheetId", spreadsheet_id),
        metadata.get("properties", {}).get("title", ""),
    )


SHEET_PROPERTIES_FIELDS = (
    "sheets(properties(sheetId,title,index,gridProperties(rowCount,columnCount)))"
)


def _list_sheet_properties(sheets_service, spreadsheet_id: str) -> list[dict]:
    """List tab properties in a spreadsheet, including sheet IDs."""
    spreadsheet = (
        sheets_service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            fields=SHEET_PROPERTIES_FIELDS,
        )
        .execute()
    )
    return _parse_sheet_properties(spreadsheet)


def _parse_sheet_properties(spreadsheet: dict) -> list[dict]:
    tabs = []
    for sheet in spreadsheet.get("sheets", []):
        props = sheet.get("properties", {})
        grid = props.get("gridProperties", {})
        tabs.append(
            {
                "sheet_id": props.get("sheetId", 0),
                "title": props.get("title", ""),
                "index": props.get("index", 0),
                "row_count": grid.get("rowCount", 0),
                "column_count": grid.get("columnCount", 0),
            }
        )
    return tabs


def list_sheet_tabs(sheets_service, spreadsheet_id: str) -> list[dict]:
    """List tabs in a spreadsheet."""
    tabs = []
    for tab in _list_sheet_properties(sheets_service, spreadsheet_id):
        tabs.append(
            {
                "sheet_id": tab["sheet_id"],
                "title": tab["title"],
                "index": tab["index"],
                "row_count": tab["row_count"],
                "column_count": tab["column_count"],
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

    result = (
        sheets_service.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=range_a1,
            valueRenderOption=value_render_option,
            dateTimeRenderOption=date_time_render_option,
        )
        .execute()
    )
    return result.get("values", [])


def update_sheet_range(
    sheets_service,
    spreadsheet_id: str,
    range_a1: str,
    values: list[list],
    value_input_option: str = "USER_ENTERED",
) -> dict:
    """Update a spreadsheet range with values."""
    if (
        not isinstance(values, list)
        or not values
        or any(not isinstance(row, list) for row in values)
    ):
        raise ValueError("values must be a non-empty list of lists")

    result = (
        sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=spreadsheet_id,
            range=range_a1,
            valueInputOption=value_input_option,
            body={"values": values},
        )
        .execute()
    )
    return {
        "cell_count": result.get("updatedCells", 0),
        "range": result.get("updatedRange", range_a1),
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
    if (
        not isinstance(values, list)
        or not values
        or any(not isinstance(row, list) for row in values)
    ):
        raise ValueError("values must be a non-empty list of lists")

    result = (
        sheets_service.spreadsheets()
        .values()
        .append(
            spreadsheetId=spreadsheet_id,
            range=range_a1,
            valueInputOption=value_input_option,
            insertDataOption=insert_data_option,
            body={"values": values},
        )
        .execute()
    )
    updates = result.get("updates", {})
    return {
        "cell_count": updates.get("updatedCells", 0),
        "range": updates.get("updatedRange", range_a1),
    }


def create_spreadsheet(sheets_service, title: str) -> dict:
    """Create a new spreadsheet and return normalized metadata."""
    result = (
        sheets_service.spreadsheets()
        .create(
            body={"properties": {"title": title}},
            fields="spreadsheetId,spreadsheetUrl",
        )
        .execute()
    )
    return {
        "spreadsheet": result.get("spreadsheetId", ""),
        "title": title,
        "url": result.get("spreadsheetUrl", ""),
    }


def _http_status(error: Exception) -> int | None:
    status = getattr(getattr(error, "resp", None), "status", None)
    return status if isinstance(status, int) else None


def mutation_failure(
    error: Exception,
    *,
    phase: str,
    dispatched: bool,
    partial: bool = False,
    recovery: dict | None = None,
) -> SheetsClientError:
    """Convert a mutation exception to a safe phase-aware error."""
    status = _http_status(error)
    if partial:
        state = "partial"
    elif not dispatched:
        state = "not_started"
    elif status is not None and 400 <= status < 500:
        state = "unchanged"
    else:
        state = "uncertain"

    if status == 401:
        code = "google_api_unauthorized"
        message = (
            "Google Sheets authorization failed after the bounded credential refresh."
        )
        action = "reconnect_google_sheets"
    elif status == 403:
        code = "google_api_forbidden"
        message = (
            "Google Sheets rejected the operation because access is not permitted."
        )
        action = "verify_permissions"
    elif status == 404:
        code = "spreadsheet_not_found"
        message = "The requested spreadsheet or range was not found."
        action = "verify_spreadsheet_and_range"
    elif status == 429:
        code = "google_rate_limited"
        message = "Google Sheets rate limited the operation."
        action = "retry_later_after_inspection"
    elif status is not None and 400 <= status < 500:
        code = "google_request_rejected"
        message = "Google Sheets rejected the operation without applying it."
        action = "correct_request"
    elif state == "not_started":
        code = "google_request_not_started"
        message = "The Google Sheets request failed before the requested operation was dispatched."
        action = "retry_after_inspection"
    elif state == "partial":
        code = "operation_partial"
        message = "The operation changed external state but did not finish."
        action = "inspect_partial_state"
    else:
        code = "operation_outcome_uncertain"
        message = (
            "The operation may have changed external state, but confirmation was lost."
        )
        action = "inspect_before_retry"

    return SheetsClientError(
        code,
        message,
        outcome_state=state,
        phase=phase,
        mutation_may_have_occurred=state in {"uncertain", "partial"},
        retry_safe=state in {"not_started", "unchanged"},
        retry_automatic=False,
        retry_action=action,
        recovery=recovery,
    )


def _copy_progress(
    *,
    spreadsheet: str,
    url: str,
    selected_tabs: list[dict],
    confirmed_tabs: list[str],
    active_tab: str | None,
    active_tab_state: str,
    finalization_state: str,
) -> dict:
    remaining = [
        tab["title"] for tab in selected_tabs if tab["title"] not in set(confirmed_tabs)
    ]
    return {
        "kind": "copy_progress",
        "destination_spreadsheet": spreadsheet,
        "destination_url": url,
        "confirmed_tabs": list(confirmed_tabs),
        "active_tab": active_tab,
        "active_tab_state": active_tab_state,
        "remaining_tabs": remaining,
        "finalization_state": finalization_state,
    }


def copy_spreadsheet(
    sheets_service,
    source_spreadsheet_id: str,
    title: str,
    tabs: list[str] | None = None,
) -> dict:
    """Copy tabs while retaining destination progress after creation."""
    requested_tabs = set(tabs) if tabs is not None else None
    try:
        source_meta = (
            sheets_service.spreadsheets()
            .get(
                spreadsheetId=source_spreadsheet_id,
                fields=f"properties(locale,timeZone),namedRanges(name),{SHEET_PROPERTIES_FIELDS}",
            )
            .execute()
        )
    except Exception as exc:
        raise mutation_failure(
            exc,
            phase="read_source_metadata",
            dispatched=False,
        ) from None

    source_tabs = _parse_sheet_properties(source_meta)
    if not source_tabs:
        raise SheetsClientError(
            "source_has_no_tabs",
            "The source spreadsheet has no tabs to copy.",
            outcome_state="not_started",
            phase="validate_source_tabs",
            mutation_may_have_occurred=False,
            retry_safe=False,
            retry_automatic=False,
            retry_action="choose_another_source",
        )

    if requested_tabs is None:
        selected_tabs = source_tabs
    else:
        selected_tabs = [tab for tab in source_tabs if tab["title"] in requested_tabs]
        selected_titles = {tab["title"] for tab in selected_tabs}
        missing_tabs = [tab for tab in tabs or [] if tab not in selected_titles]
        if missing_tabs:
            raise SheetsClientError(
                "tabs_not_found",
                "One or more requested tabs do not exist in the source spreadsheet.",
                outcome_state="not_started",
                phase="validate_source_tabs",
                mutation_may_have_occurred=False,
                retry_safe=True,
                retry_automatic=False,
                retry_action="use_listed_tab_titles",
            )

    try:
        create_result = (
            sheets_service.spreadsheets()
            .create(
                body={"properties": {"title": title}},
                fields="spreadsheetId,spreadsheetUrl,properties(locale,timeZone),sheets(properties(sheetId,title,index))",
            )
            .execute()
        )
    except Exception as exc:
        raise mutation_failure(
            exc,
            phase="create_destination",
            dispatched=True,
        ) from None

    new_spreadsheet_id = create_result.get("spreadsheetId", "")
    spreadsheet_url = create_result.get("spreadsheetUrl", "")
    if not new_spreadsheet_id:
        raise SheetsClientError(
            "destination_identity_missing",
            "Google Sheets created an unconfirmed destination without returning its ID.",
            outcome_state="uncertain",
            phase="create_destination",
            mutation_may_have_occurred=True,
            retry_safe=False,
            retry_automatic=False,
            retry_action="inspect_drive_before_retry",
        )

    created_sheets = create_result.get("sheets", [])
    default_sheet_id = None
    if len(created_sheets) == 1:
        default_sheet_id = created_sheets[0].get("properties", {}).get("sheetId")
    for sheet in created_sheets:
        props = sheet.get("properties", {})
        if props.get("title") == "Sheet1":
            default_sheet_id = props.get("sheetId")
            break
    if default_sheet_id is None:
        recovery = _copy_progress(
            spreadsheet=new_spreadsheet_id,
            url=spreadsheet_url,
            selected_tabs=selected_tabs,
            confirmed_tabs=[],
            active_tab=None,
            active_tab_state="not_started",
            finalization_state="not_started",
        )
        raise SheetsClientError(
            "copy_partial",
            "The destination was created but its default tab could not be identified.",
            outcome_state="partial",
            phase="inspect_destination",
            mutation_may_have_occurred=True,
            retry_safe=False,
            retry_automatic=False,
            retry_action="inspect_destination",
            recovery=recovery,
        )

    copied_tabs: list[dict] = []
    confirmed_titles: list[str] = []
    for source_tab in selected_tabs:
        active_title = source_tab["title"]
        try:
            copied_props = (
                sheets_service.spreadsheets()
                .sheets()
                .copyTo(
                    spreadsheetId=source_spreadsheet_id,
                    sheetId=source_tab["sheet_id"],
                    body={"destinationSpreadsheetId": new_spreadsheet_id},
                )
                .execute()
            )
        except Exception as exc:
            recovery = _copy_progress(
                spreadsheet=new_spreadsheet_id,
                url=spreadsheet_url,
                selected_tabs=selected_tabs,
                confirmed_tabs=confirmed_titles,
                active_tab=active_title,
                active_tab_state="uncertain",
                finalization_state="not_started",
            )
            raise mutation_failure(
                exc,
                phase="copy_tabs",
                dispatched=True,
                partial=True,
                recovery=recovery,
            ) from None
        copied_sheet_id = copied_props.get("sheetId")
        if copied_sheet_id is None:
            recovery = _copy_progress(
                spreadsheet=new_spreadsheet_id,
                url=spreadsheet_url,
                selected_tabs=selected_tabs,
                confirmed_tabs=confirmed_titles,
                active_tab=active_title,
                active_tab_state="uncertain",
                finalization_state="not_started",
            )
            raise SheetsClientError(
                "copy_partial",
                "A tab copy was not confirmed after the destination was created.",
                outcome_state="partial",
                phase="copy_tabs",
                mutation_may_have_occurred=True,
                retry_safe=False,
                retry_automatic=False,
                retry_action="inspect_destination",
                recovery=recovery,
            )
        copied_tabs.append(
            {
                "source_title": source_tab["title"],
                "copied_sheet_id": copied_sheet_id,
            }
        )
        confirmed_titles.append(active_title)

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

    try:
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=new_spreadsheet_id,
            body={"requests": batch_requests},
        ).execute()
    except Exception as exc:
        recovery = _copy_progress(
            spreadsheet=new_spreadsheet_id,
            url=spreadsheet_url,
            selected_tabs=selected_tabs,
            confirmed_tabs=confirmed_titles,
            active_tab=None,
            active_tab_state="not_started",
            finalization_state="uncertain",
        )
        raise mutation_failure(
            exc,
            phase="finalize_destination",
            dispatched=True,
            partial=True,
            recovery=recovery,
        ) from None

    warnings: list[dict] = []
    named_ranges = sorted(
        named_range.get("name", "")
        for named_range in source_meta.get("namedRanges", [])
    )
    if named_ranges:
        warnings.append(
            {
                "code": "named_ranges_not_copied",
                "message": (
                    f"{len(named_ranges)} spreadsheet-level named range(s) were not copied; "
                    "tab-level copy cannot carry them."
                ),
            }
        )
    source_props = source_meta.get("properties", {})
    copy_props = create_result.get("properties", {})
    for prop_key, label in (("locale", "locale"), ("timeZone", "time zone")):
        source_value = source_props.get(prop_key)
        copy_value = copy_props.get(prop_key)
        if source_value and copy_value and source_value != copy_value:
            warnings.append(
                {
                    "code": "locale_not_preserved"
                    if prop_key == "locale"
                    else "time_zone_not_preserved",
                    "message": (
                        f"The destination {label} differs from the source; "
                        "tab-level copy cannot preserve spreadsheet-level settings."
                    ),
                }
            )

    return {
        "spreadsheet": new_spreadsheet_id,
        "title": title,
        "url": spreadsheet_url,
        "tabs": [copied_tab["source_title"] for copied_tab in copied_tabs],
        "warnings": warnings,
    }


def search_spreadsheets(
    drive_service,
    query: str,
    limit: int = 10,
) -> list[dict]:
    """Search Google Drive for spreadsheets by name."""
    escaped_query = query.replace("\\", "\\\\").replace("'", "\\'")
    search_query = (
        f"name contains '{escaped_query}' and "
        f"mimeType = '{GOOGLE_SHEET_MIME}' and "
        "trashed = false"
    )
    results = (
        drive_service.files()
        .list(
            q=search_query,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            fields="files(id, name, modifiedTime, webViewLink)",
            pageSize=limit,
        )
        .execute()
    )
    return [
        {
            "spreadsheet": item.get("id", ""),
            "title": item.get("name", ""),
            "url": item.get("webViewLink", ""),
            "modified_at": item.get("modifiedTime"),
        }
        for item in results.get("files", [])
    ]


def clear_sheet_range(sheets_service, spreadsheet_id: str, range_a1: str) -> dict:
    """Clear all values in a spreadsheet range."""
    result = (
        sheets_service.spreadsheets()
        .values()
        .clear(
            spreadsheetId=spreadsheet_id,
            range=range_a1,
            body={},
        )
        .execute()
    )
    return {
        "range": result.get("clearedRange", range_a1),
    }


def _formula_counts(values: list[list]) -> tuple[int, int]:
    formula_count = 0
    literal_count = 0
    for row in values:
        for value in row:
            if value is None or value == "":
                continue
            if isinstance(value, str) and value.startswith("="):
                formula_count += 1
            else:
                literal_count += 1
    return formula_count, literal_count


def _formula_snapshot_matches(expected: list[list], actual: list[list]) -> bool:
    return actual == expected


def _range_recovery(
    spreadsheet_id: str,
    range_a1: str,
    formula_count: int,
    *,
    attempted: bool,
    verified: bool,
) -> dict:
    return {
        "kind": "range_state",
        "spreadsheet": spreadsheet_id,
        "range": range_a1,
        "formula_cell_count": formula_count,
        "compensation_attempted": attempted,
        "compensation_verified": verified,
    }


def recalculate_sheet_range(sheets_service, spreadsheet_id: str, range_a1: str) -> dict:
    """Request recalculation while preserving and verifying formula cells."""
    try:
        snapshot = read_sheet_range(
            sheets_service,
            spreadsheet_id,
            range_a1,
            value_render_option="FORMULA",
        )
    except Exception as exc:
        if isinstance(exc, SheetsClientError):
            raise
        raise SheetsClientError(
            "recalculation_read_failed",
            "The formula range could not be read before recalculation.",
            outcome_state="not_started",
            phase="read_formulas",
            mutation_may_have_occurred=False,
            retry_safe=True,
            retry_automatic=False,
            retry_action="retry_read",
        ) from None

    formula_count, literal_count = _formula_counts(snapshot)
    if literal_count:
        raise SheetsClientError(
            "formula_range_required",
            "Recalculation requires a range containing only formulas and blank cells.",
            outcome_state="unchanged",
            phase="validate_formula_range",
            mutation_may_have_occurred=False,
            retry_safe=True,
            retry_automatic=False,
            retry_action="choose_formula_only_range",
        )
    if formula_count == 0:
        return {
            "range": range_a1,
            "cell_count": 0,
            "recovery_performed": False,
            "formulas_verified": True,
        }

    clear_may_have_occurred = False
    compensation_path = False
    try:
        clear_sheet_range(sheets_service, spreadsheet_id, range_a1)
        clear_may_have_occurred = True
    except Exception as exc:
        status = _http_status(exc)
        if status is not None and 400 <= status < 500:
            raise mutation_failure(
                exc,
                phase="clear_formula_range",
                dispatched=True,
            ) from None
        clear_may_have_occurred = True
        compensation_path = True

    primary_error: Exception | None = None
    if clear_may_have_occurred:
        try:
            updated = update_sheet_range(
                sheets_service,
                spreadsheet_id,
                range_a1,
                snapshot,
                value_input_option="USER_ENTERED",
            )
            verified = read_sheet_range(
                sheets_service,
                spreadsheet_id,
                range_a1,
                value_render_option="FORMULA",
            )
            if _formula_snapshot_matches(snapshot, verified):
                return {
                    "range": updated.get("range", range_a1),
                    "cell_count": formula_count,
                    "recovery_performed": compensation_path,
                    "formulas_verified": True,
                }
            primary_error = RuntimeError("formula verification mismatch")
        except Exception as exc:
            primary_error = exc

    if compensation_path:
        # An uncertain clear makes the first exact restore the sole bounded
        # compensation attempt. A second write would risk replaying a mutation
        # whose first outcome is itself unknown.
        recovery = _range_recovery(
            spreadsheet_id,
            range_a1,
            formula_count,
            attempted=True,
            verified=False,
        )
        raise SheetsClientError(
            "recalculation_recovery_failed",
            "The formula range may be partially changed and automatic recovery could not be verified.",
            outcome_state="partial",
            phase="restore_and_verify",
            mutation_may_have_occurred=True,
            retry_safe=False,
            retry_automatic=False,
            retry_action="inspect_formula_range",
            recovery=recovery,
        ) from primary_error

    try:
        restored = update_sheet_range(
            sheets_service,
            spreadsheet_id,
            range_a1,
            snapshot,
            value_input_option="USER_ENTERED",
        )
        verified = read_sheet_range(
            sheets_service,
            spreadsheet_id,
            range_a1,
            value_render_option="FORMULA",
        )
        if _formula_snapshot_matches(snapshot, verified):
            return {
                "range": restored.get("range", range_a1),
                "cell_count": formula_count,
                "recovery_performed": True,
                "formulas_verified": True,
            }
    except Exception:
        pass

    recovery = _range_recovery(
        spreadsheet_id,
        range_a1,
        formula_count,
        attempted=True,
        verified=False,
    )
    raise SheetsClientError(
        "recalculation_recovery_failed",
        "The formula range may be partially changed and automatic recovery could not be verified.",
        outcome_state="partial" if clear_may_have_occurred else "uncertain",
        phase="restore_and_verify",
        mutation_may_have_occurred=True,
        retry_safe=False,
        retry_automatic=False,
        retry_action="inspect_formula_range",
        recovery=recovery,
    ) from primary_error
