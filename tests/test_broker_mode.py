"""Offline acceptance tests for per-user broker credential mode."""

from __future__ import annotations

import io
import json
import time
from datetime import datetime, timedelta, timezone
from unittest.mock import MagicMock
from urllib.error import HTTPError, URLError

import google_auth_httplib2
import httplib2
import pytest
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from gsheets_mcp import sheets_client, tools
from gsheets_mcp.contracts import CreateSpreadsheetInput


class _Response:
    def __init__(self, payload: dict):
        self.body = json.dumps(payload).encode()

    def __enter__(self):
        return self

    def __exit__(self, *_args):
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


def test_broker_success_builds_memory_credentials_and_caches(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    requests = []

    def urlopen(request, timeout):
        requests.append((request, timeout))
        return _Response(
            {
                "access_token": "redacted-token",
                "expires_at": int(time.time()) + 600,
            }
        )

    monkeypatch.setattr(sheets_client.urllib_request, "urlopen", urlopen)

    first = sheets_client._get_credentials()
    second = sheets_client._get_credentials()

    assert first is second
    assert first.token == "redacted-token"
    assert len(requests) == 1
    request, timeout = requests[0]
    assert request.full_url == (
        "http://broker.test/base/api/internal/google/sheets-access-token"
    )
    assert request.get_method() == "POST"
    assert request.headers["Authorization"] == "Bearer session-placeholder"
    assert timeout == 10


def test_broker_refetches_near_expiry(monkeypatch: pytest.MonkeyPatch) -> None:
    tokens = iter(["first-token", "second-token"])
    calls = 0

    def urlopen(_request, timeout):
        nonlocal calls
        assert timeout == 10
        calls += 1
        return _Response(
            {
                "access_token": next(tokens),
                "expires_at": int(time.time()) + 30,
            }
        )

    monkeypatch.setattr(sheets_client.urllib_request, "urlopen", urlopen)

    first = sheets_client._get_credentials()
    assert first.token == "first-token"
    second = sheets_client._get_credentials()

    assert first is second
    assert second.token == "second-token"
    assert calls == 2


def test_broker_service_pins_one_request_refresh_attempt(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
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
        (
            401,
            {"error": "broker_session_expired"},
            "broker_session_expired",
        ),
    ],
)
def test_broker_maps_auth_errors_to_typed_safe_codes(
    monkeypatch: pytest.MonkeyPatch,
    status: int,
    payload: dict,
    code: str,
) -> None:
    monkeypatch.setattr(
        sheets_client.urllib_request,
        "urlopen",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(_http_error(status, payload)),
    )

    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client._get_credentials()

    assert exc_info.value.code == code


def test_broker_rate_limit_is_terminal_and_surfaces_retry_after(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls = 0

    def urlopen(*_args, **_kwargs):
        nonlocal calls
        calls += 1
        raise _http_error(
            429,
            {"error": "broker_rate_limited", "retry_after_s": 17},
        )

    monkeypatch.setattr(sheets_client.urllib_request, "urlopen", urlopen)

    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client._get_credentials()

    assert exc_info.value.code == "broker_rate_limited"
    assert exc_info.value.details["retry_after_s"] == 17
    assert exc_info.value.details["outcome_state"] == "not_started"
    assert exc_info.value.details["mutation_may_have_occurred"] is False
    assert exc_info.value.details["retry_safe"] is True
    assert exc_info.value.details["retry_automatic"] is False
    assert (
        str(exc_info.value) == "The Google Sheets token broker rate limited this call."
    )
    assert calls == 1


def test_broker_connection_failure_is_sanitized(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        sheets_client.urllib_request,
        "urlopen",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            URLError("internal host and secret")
        ),
    )

    with pytest.raises(sheets_client.SheetsClientError) as exc_info:
        sheets_client._get_credentials()

    assert exc_info.value.code == "sheets_unavailable"
    assert "internal host" not in str(exc_info.value)
    assert "secret" not in str(exc_info.value)


def test_broker_tool_call_never_reads_or_writes_local_oauth_files(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    token_path = tmp_path / "token.pickle"
    credentials_path = tmp_path / "drive_credentials.json"
    monkeypatch.setattr(sheets_client, "TOKEN_FILE", token_path)
    monkeypatch.setattr(sheets_client, "CREDENTIALS_FILE", credentials_path)
    monkeypatch.setattr(
        sheets_client.urllib_request,
        "urlopen",
        lambda *_args, **_kwargs: _Response(
            {
                "access_token": "memory-only",
                "expires_at": int(time.time()) + 600,
            }
        ),
    )
    service = MagicMock()
    service.spreadsheets.return_value.create.return_value.execute.return_value = {
        "spreadsheetId": "new-spreadsheet-id",
        "spreadsheetUrl": ("https://docs.google.com/spreadsheets/d/new-spreadsheet-id"),
    }
    monkeypatch.setattr(sheets_client, "build", lambda *_args, **_kwargs: service)

    result = tools.create_spreadsheet(CreateSpreadsheetInput(title="New Sheet"))

    assert result.status == "ok"
    assert result.spreadsheet == "new-spreadsheet-id"
    assert not token_path.exists()
    assert not credentials_path.exists()
    assert list(tmp_path.iterdir()) == []


class _InjectingHttp:
    def __init__(self, responder):
        self.responder = responder
        self.requests = []

    def request(self, uri, method="GET", body=None, headers=None, **_kwargs):
        self.requests.append((method, uri, body, headers))
        status, payload = self.responder(method, uri, body)
        response = httplib2.Response(
            {
                "status": str(status),
                "reason": "Unauthorized" if status == 401 else "OK",
            }
        )
        response.reason = "Unauthorized" if status == 401 else "OK"
        return response, json.dumps(payload).encode()


def _real_service(http):
    credentials = sheets_client.BrokerCredentials()
    credentials.token = "initial-token"
    credentials.expiry = datetime.now(timezone.utc).replace(tzinfo=None) + timedelta(
        minutes=10
    )
    authorized_http = google_auth_httplib2.AuthorizedHttp(
        credentials,
        http=http,
        max_refresh_attempts=1,
    )
    return build("sheets", "v4", http=authorized_http, static_discovery=True)


def _install_broker_refresh(monkeypatch: pytest.MonkeyPatch, *, error=None):
    calls = []

    def fetch():
        calls.append(1)
        if error is not None:
            raise error
        return "refreshed-token", int(time.time()) + 600

    monkeypatch.setattr(sheets_client, "_fetch_broker_token", fetch)
    return calls


def test_google_401_refreshes_once_and_never_loops(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    http = _InjectingHttp(lambda _method, _uri, _body: (401, {}))
    service = _real_service(http)
    refreshes = _install_broker_refresh(monkeypatch)

    with pytest.raises(HttpError):
        sheets_client.create_spreadsheet(service, "New Sheet")

    assert len(refreshes) == 1
    assert len(http.requests) == 2


def test_copy_401_retries_only_the_inflight_google_request(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    copy_attempts = 0

    def respond(method, uri, _body):
        nonlocal copy_attempts
        if method == "GET":
            return 200, {
                "properties": {},
                "sheets": [
                    {
                        "properties": {
                            "sheetId": 101,
                            "title": "Data",
                            "index": 0,
                        }
                    }
                ],
            }
        if method == "POST" and "/v4/spreadsheets?" in uri:
            return 200, {
                "spreadsheetId": "destination-spreadsheet-id",
                "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/destination-spreadsheet-id",
                "sheets": [
                    {
                        "properties": {
                            "sheetId": 900,
                            "title": "Sheet1",
                        }
                    }
                ],
            }
        if ":copyTo" in uri:
            copy_attempts += 1
            if copy_attempts == 1:
                return 401, {}
            return 200, {"sheetId": 201}
        if ":batchUpdate" in uri:
            return 200, {}
        raise AssertionError((method, uri))

    http = _InjectingHttp(respond)
    service = _real_service(http)
    refreshes = _install_broker_refresh(monkeypatch)

    result = sheets_client.copy_spreadsheet(
        service,
        "source-spreadsheet-id",
        "Copy",
    )

    assert result["spreadsheet"] == "destination-spreadsheet-id"
    assert (
        sum(
            method == "POST" and "/v4/spreadsheets?" in uri
            for method, uri, _body, _headers in http.requests
        )
        == 1
    )
    assert copy_attempts == 2
    assert len(refreshes) == 1


def test_broker_mode_never_enters_interactive_consent(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        sheets_client,
        "_run_installed_app_flow",
        lambda: (_ for _ in ()).throw(
            AssertionError("broker mode must never launch browser consent")
        ),
    )
    monkeypatch.setattr(
        sheets_client,
        "_fetch_broker_token",
        lambda: ("memory-only", int(time.time()) + 600),
    )

    credentials = sheets_client._get_credentials()

    assert isinstance(credentials, sheets_client.BrokerCredentials)
