"""Offline acceptance tests for per-user broker credential mode."""

import asyncio
import io
import json
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock
from urllib.error import HTTPError, URLError

import google_auth_httplib2
import httplib2
import pytest
from googleapiclient.discovery import build

from src import server, sheets_client


class _Response:
    def __init__(self, payload: dict):
        self.body = json.dumps(payload).encode()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def read(self):
        return self.body


@pytest.fixture(autouse=True)
def _broker_env(monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("GSHEETS_TOKEN_MODE", "broker")
    monkeypatch.setenv("GSHEETS_BROKER_URL", "http://broker.test/base")
    monkeypatch.setenv("GSHEETS_BROKER_SESSION_TOKEN", "session-placeholder")
    sheets_client.invalidate_broker_credentials()
    yield
    sheets_client.invalidate_broker_credentials()


def _http_error(status: int, payload: dict) -> HTTPError:
    return HTTPError(
        "http://broker.test",
        status,
        "error",
        {},
        io.BytesIO(json.dumps(payload).encode()),
    )


def test_broker_success_builds_credentials_and_caches(monkeypatch):
    calls = []

    def urlopen(request, timeout):
        calls.append(request)
        return _Response({"access_token": "redacted-token", "expires_at": int(time.time()) + 600})

    monkeypatch.setattr(sheets_client.urllib_request, "urlopen", urlopen)
    first = sheets_client._get_credentials()
    second = sheets_client._get_credentials()

    assert first is second
    assert first.token == "redacted-token"
    assert len(calls) == 1
    assert calls[0].full_url == "http://broker.test/base/api/internal/google/sheets-access-token"
    assert calls[0].get_method() == "POST"


def test_broker_refetches_near_expiry(monkeypatch):
    tokens = iter(["first-token", "second-token"])
    calls = 0

    def urlopen(request, timeout):
        nonlocal calls
        calls += 1
        return _Response({"access_token": next(tokens), "expires_at": int(time.time()) + 30})

    monkeypatch.setattr(sheets_client.urllib_request, "urlopen", urlopen)
    first = sheets_client._get_credentials()
    assert first.token == "first-token"
    second = sheets_client._get_credentials()
    assert second.token == "second-token"
    assert first is second
    assert calls == 2


def test_broker_service_pins_one_request_refresh_attempt(monkeypatch):
    underlying_http = MagicMock()
    captured = {}
    monkeypatch.setattr(sheets_client, "build_http", lambda: underlying_http)
    monkeypatch.setattr(
        sheets_client,
        "_fetch_broker_token",
        lambda: ("memory-only", int(time.time()) + 600),
    )

    def fake_build(api, version, **kwargs):
        captured.update(api=api, version=version, **kwargs)
        return object()

    monkeypatch.setattr(sheets_client, "build", fake_build)
    sheets_client.get_sheets_service()

    assert captured["api"] == "sheets"
    assert captured["version"] == "v4"
    assert "credentials" not in captured
    assert isinstance(captured["http"], google_auth_httplib2.AuthorizedHttp)
    assert captured["http"].http is underlying_http
    assert captured["http"]._max_refresh_attempts == 1


@pytest.mark.parametrize(
    ("status", "payload", "code"),
    [
        (404, {"error": "sheets_not_connected"}, "sheets_not_connected"),
        (401, {"error": "auth_failed"}, "broker_session_expired"),
        (401, {"error": "broker_session_expired"}, "broker_session_expired"),
    ],
)
def test_broker_typed_errors(monkeypatch, status, payload, code):
    monkeypatch.setattr(
        sheets_client.urllib_request,
        "urlopen",
        lambda *args, **kwargs: (_ for _ in ()).throw(_http_error(status, payload)),
    )
    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client._get_credentials()
    assert exc_info.value.code == code


def test_broker_rate_limit_is_terminal_and_surfaces_retry_after(monkeypatch):
    calls = 0

    def urlopen(*args, **kwargs):
        nonlocal calls
        calls += 1
        raise _http_error(429, {"error": "broker_rate_limited", "retry_after_s": 17})

    monkeypatch.setattr(sheets_client.urllib_request, "urlopen", urlopen)
    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client._get_credentials()
    assert exc_info.value.code == "broker_rate_limited"
    assert exc_info.value.details == {"retry_after_s": 17}
    assert "retry_after_s=17" in str(exc_info.value)
    assert calls == 1


def test_broker_connection_error_is_typed_unavailable(monkeypatch):
    monkeypatch.setattr(
        sheets_client.urllib_request,
        "urlopen",
        lambda *args, **kwargs: (_ for _ in ()).throw(URLError("offline")),
    )
    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client._get_credentials()
    assert exc_info.value.code == "sheets_unavailable"
    assert "offline" not in str(exc_info.value)


class _InjectingHttp:
    def __init__(self, responder):
        self.responder = responder
        self.requests = []

    def request(self, uri, method="GET", body=None, headers=None, **kwargs):
        self.requests.append((method, uri, body))
        status, payload = self.responder(method, uri, body)
        response = httplib2.Response(
            {"status": str(status), "reason": "Unauthorized" if status == 401 else "OK"}
        )
        response.reason = "Unauthorized" if status == 401 else "OK"
        return response, json.dumps(payload).encode()


def _real_service(http):
    creds = sheets_client.BrokerCredentials()
    creds.token = "initial-token"
    creds.expiry = datetime.now(timezone.utc).replace(tzinfo=None) + timedelta(minutes=10)
    authed_http = google_auth_httplib2.AuthorizedHttp(
        creds, http=http, max_refresh_attempts=1
    )
    return build("sheets", "v4", http=authed_http, static_discovery=True)


def _install_broker_refresh(monkeypatch, *, error=None):
    calls = []

    def fetch():
        calls.append(1)
        if error is not None:
            raise error
        return "refreshed-token", int(time.time()) + 600

    monkeypatch.setattr(sheets_client, "_fetch_broker_token", fetch)
    return calls


def test_copy_401_retries_only_copy_request(monkeypatch):
    copy_attempts = 0

    def respond(method, uri, body):
        nonlocal copy_attempts
        if method == "GET":
            if "namedRanges" in uri:
                return 200, {"properties": {}}
            return 200, {"sheets": [{"properties": {"sheetId": 101, "title": "Data", "index": 0}}]}
        if method == "POST" and "/v4/spreadsheets?" in uri:
            return 200, {"spreadsheetId": "destination", "spreadsheetUrl": "url", "sheets": [{"properties": {"sheetId": 900, "title": "Sheet1"}}]}
        if ":copyTo" in uri:
            copy_attempts += 1
            return (401, {}) if copy_attempts == 1 else (200, {"sheetId": 201})
        if ":batchUpdate" in uri:
            return 200, {}
        raise AssertionError((method, uri, body))

    http = _InjectingHttp(respond)
    service = _real_service(http)
    refreshes = _install_broker_refresh(monkeypatch)
    monkeypatch.setattr(sheets_client, "resolve_spreadsheet_id", lambda source: ("source", "Source"))
    monkeypatch.setattr(sheets_client, "get_sheets_service", lambda: service)

    payload = json.loads(server.gsheet_copy_spreadsheet("source", "Copy"))
    assert payload["status"] == "ok", payload
    assert sum(method == "POST" and "/v4/spreadsheets?" in uri for method, uri, _ in http.requests) == 1
    assert copy_attempts == 2
    assert len(refreshes) == 1


def test_touch_401_retries_only_update_request(monkeypatch):
    update_attempts = 0
    formulas = [["=CUSTOM()"]]

    def respond(method, uri, body):
        nonlocal update_attempts
        if method == "GET" and "/values/" in uri:
            return 200, {"values": formulas}
        if method == "POST" and uri.endswith(":clear?alt=json"):
            return 200, {"clearedRange": "Sheet1!A1"}
        if method == "PUT" and "/values/" in uri:
            update_attempts += 1
            return (401, {}) if update_attempts == 1 else (200, {"updatedRange": "Sheet1!A1", "updatedCells": 1})
        raise AssertionError((method, uri, body))

    http = _InjectingHttp(respond)
    service = _real_service(http)
    refreshes = _install_broker_refresh(monkeypatch)
    monkeypatch.setattr(sheets_client, "resolve_spreadsheet_id", lambda source: ("sheet", "Sheet"))
    monkeypatch.setattr(sheets_client, "get_sheets_service", lambda: service)

    payload = json.loads(server.gsheet_touch_range("sheet", "Sheet1!A1"))
    assert payload["status"] == "ok"
    assert payload["touchedCells"] == 1
    assert sum(method == "GET" for method, _, _ in http.requests) == 1
    assert sum(uri.endswith(":clear?alt=json") for _, uri, _ in http.requests) == 1
    assert update_attempts == 2
    assert len(refreshes) == 1


def test_second_google_401_is_typed_and_does_not_loop(monkeypatch):
    http = _InjectingHttp(lambda method, uri, body: (401, {}))
    service = _real_service(http)
    refreshes = _install_broker_refresh(monkeypatch)
    monkeypatch.setattr(sheets_client, "get_sheets_service", lambda: service)

    payload = json.loads(server.gsheet_create("New Sheet"))
    assert payload["error_code"] == "google_api_unauthorized"
    assert len(refreshes) == 1
    assert len(http.requests) == 2


def test_mid_tool_broker_rate_limit_is_typed_and_terminal(monkeypatch):
    http = _InjectingHttp(lambda method, uri, body: (401, {}))
    service = _real_service(http)
    refreshes = _install_broker_refresh(
        monkeypatch,
        error=sheets_client.SheetsClientError(
            "broker_rate_limited", "rate limited", retry_after_s=17
        ),
    )
    monkeypatch.setattr(sheets_client, "get_sheets_service", lambda: service)

    payload = json.loads(server.gsheet_create("New Sheet"))
    assert payload["error_code"] == "broker_rate_limited"
    assert payload["retry_after_s"] == 17
    assert len(refreshes) == 1
    assert len(http.requests) == 1


def test_broker_full_tool_call_writes_no_token_file(monkeypatch, tmp_path):
    token_path = tmp_path / "token.pickle"
    credentials_path = tmp_path / "drive_credentials.json"
    monkeypatch.setattr(sheets_client, "TOKEN_FILE", token_path)
    monkeypatch.setattr(sheets_client, "CREDENTIALS_FILE", credentials_path)
    monkeypatch.setattr(
        sheets_client.urllib_request,
        "urlopen",
        lambda *args, **kwargs: _Response(
            {"access_token": "memory-only", "expires_at": int(time.time()) + 600}
        ),
    )
    service = MagicMock()
    service.spreadsheets.return_value.create.return_value.execute.return_value = {
        "spreadsheetId": "new-sheet",
        "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/new-sheet",
    }
    monkeypatch.setattr(sheets_client, "build", lambda *args, **kwargs: service)

    assert json.loads(server.gsheet_create("New Sheet"))["status"] == "ok"
    assert not token_path.exists()
    assert not credentials_path.exists()
    assert list(tmp_path.iterdir()) == []


def test_broker_headless_consent_path_hard_fails(monkeypatch, tmp_path):
    monkeypatch.setenv("GSHEETS_HEADLESS", "1")
    monkeypatch.setattr(sheets_client, "CREDENTIALS_FILE", tmp_path / "client.json")
    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client._run_installed_app_flow()
    assert exc_info.value.code == "interactive_consent_disabled"


@pytest.mark.parametrize(
    ("address", "expected"),
    [
        ("abc1234567890XYZ___123", "abc1234567890XYZ___123"),
        ("https://docs.google.com/spreadsheets/d/abc1234567890XYZ___123/edit#gid=0", "abc1234567890XYZ___123"),
        ("https://DOCS.GOOGLE.COM/spreadsheets/d/abc1234567890XYZ___123/edit", "abc1234567890XYZ___123"),
        ("https://docs.google.com/spreadsheets/u/0/d/abc1234567890XYZ___123/edit#gid=0", "abc1234567890XYZ___123"),
        ("https://docs.google.com/spreadsheets/u/12/d/abc1234567890XYZ___123/edit?usp=sharing", "abc1234567890XYZ___123"),
    ],
)
def test_broker_resolve_id_or_url_validates_with_sheets(monkeypatch, address, expected):
    service = MagicMock()
    service.spreadsheets.return_value.get.return_value.execute.return_value = {
        "spreadsheetId": expected,
        "properties": {"title": "Validated"},
    }
    monkeypatch.setattr(sheets_client, "get_sheets_service", lambda: service)
    assert sheets_client.resolve_spreadsheet_id(address) == (expected, "Validated")
    service.spreadsheets.return_value.get.assert_called_once_with(
        spreadsheetId=expected,
        fields="spreadsheetId,properties(title)",
    )


def test_broker_resolve_title_and_malformed_url_are_typed():
    with pytest.raises(sheets_client.SheetsClientError) as title_error:
        sheets_client.resolve_spreadsheet_id("Quarterly Comps")
    assert title_error.value.code == "title_resolution_requires_drive_scope"

    with pytest.raises(sheets_client.SheetsClientError) as url_error:
        sheets_client.resolve_spreadsheet_id("https://example.com/spreadsheets/d/not-google")
    assert url_error.value.code == "invalid_spreadsheet_url"

    for address in (
        "https://docs.google.com@evil.com/spreadsheets/d/abc1234567890XYZ___123",
        "https://docs.google.com:8443/spreadsheets/d/abc1234567890XYZ___123",
    ):
        with pytest.raises(sheets_client.SheetsClientError) as unsafe_url_error:
            sheets_client.resolve_spreadsheet_id(address)
        assert unsafe_url_error.value.code == "invalid_spreadsheet_url"

    with pytest.raises(sheets_client.SheetsClientError) as qualifier_error:
        sheets_client.resolve_spreadsheet_id(
            "https://docs.google.com/spreadsheets/u/x/d/abc1234567890XYZ___123/edit"
        )
    assert qualifier_error.value.code == "invalid_spreadsheet_url"


def test_gsheet_search_broker_mode_requires_drive_scope():
    payload = json.loads(server.gsheet_search("Comps"))
    assert payload["error_code"] == "requires_drive_scope"


def test_gsheet_search_dev_mode_regression(monkeypatch):
    monkeypatch.delenv("GSHEETS_TOKEN_MODE")
    monkeypatch.setattr(sheets_client, "authenticate", lambda: object())
    monkeypatch.setattr(sheets_client, "search_spreadsheets", lambda *args, **kwargs: [{"id": "sheet-1"}])
    payload = json.loads(server.gsheet_search("Comps"))
    assert payload["status"] == "ok"
    assert payload["results"] == [{"id": "sheet-1"}]


def test_tools_list_needs_no_broker_or_local_credentials(monkeypatch, tmp_path):
    monkeypatch.delenv("GSHEETS_BROKER_URL")
    monkeypatch.delenv("GSHEETS_BROKER_SESSION_TOKEN")
    monkeypatch.setattr(sheets_client, "TOKEN_FILE", tmp_path / "absent-token")
    monkeypatch.setattr(sheets_client, "CREDENTIALS_FILE", tmp_path / "absent-client")
    tools = asyncio.run(server.mcp.list_tools())
    assert "gsheet_list_tabs" in {tool.name for tool in tools}
    assert "gsheet_search" in {tool.name for tool in tools}
