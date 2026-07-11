"""Offline acceptance tests for per-user broker credential mode."""

import asyncio
import io
import json
import time
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock
from urllib.error import HTTPError, URLError

import pytest

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
    assert sheets_client._get_credentials().token == "first-token"
    assert sheets_client._get_credentials().token == "second-token"
    assert calls == 2


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


def test_google_401_refetches_once_per_tool_call(monkeypatch):
    broker_calls = 0
    api_calls = 0

    def urlopen(*args, **kwargs):
        nonlocal broker_calls
        broker_calls += 1
        return _Response({"access_token": f"token-{broker_calls}", "expires_at": int(time.time()) + 600})

    def build_service():
        nonlocal api_calls
        sheets_client._get_credentials()
        service = MagicMock()

        def execute():
            nonlocal api_calls
            api_calls += 1
            if api_calls == 1:
                raise sheets_client.HttpError(SimpleNamespace(status=401, reason="Unauthorized"), b"")
            return {"spreadsheetId": "new-sheet", "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/new-sheet"}

        service.spreadsheets.return_value.create.return_value.execute.side_effect = execute
        return service

    monkeypatch.setattr(sheets_client.urllib_request, "urlopen", urlopen)
    monkeypatch.setattr(sheets_client, "get_sheets_service", build_service)
    payload = json.loads(server.gsheet_create("New Sheet"))
    assert payload["status"] == "ok"
    assert broker_calls == 2
    assert api_calls == 2


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
