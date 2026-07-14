"""Contract and wire-level tests for the full-cutover MCP surface."""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any

import anyio
import jsonschema
import pytest
from mcp import ClientSession, types
from mcp.client.stdio import StdioServerParameters, stdio_client

from gsheets_mcp import __version__
from gsheets_mcp import server as server_module
from gsheets_mcp import sheets_client
from gsheets_mcp.contracts import ListTabsSuccess, Tab


SPREADSHEET_ID = "1ZyXwVuTsRqPoNmLkJiHgFeDcBa9876543210"
LOCAL_TOOLS = (
    "gsheets_list_tabs",
    "gsheets_read_range",
    "gsheets_write_range",
    "gsheets_append_rows",
    "gsheets_create_spreadsheet",
    "gsheets_copy_spreadsheet",
    "gsheets_search_spreadsheets",
    "gsheets_clear_range",
    "gsheets_recalculate_range",
)
BROKER_TOOLS = tuple(
    name for name in LOCAL_TOOLS if name != "gsheets_search_spreadsheets"
)
EXPECTED_INPUT_FIELDS = {
    "gsheets_list_tabs": {"spreadsheet"},
    "gsheets_read_range": {
        "spreadsheet",
        "range",
        "value_render_option",
        "date_time_render_option",
    },
    "gsheets_write_range": {"spreadsheet", "range", "values"},
    "gsheets_append_rows": {"spreadsheet", "range", "values"},
    "gsheets_create_spreadsheet": {"title"},
    "gsheets_copy_spreadsheet": {"spreadsheet", "title", "tabs"},
    "gsheets_search_spreadsheets": {"query", "limit"},
    "gsheets_clear_range": {"spreadsheet", "range"},
    "gsheets_recalculate_range": {"spreadsheet", "range"},
}
EXPECTED_ANNOTATIONS = {
    "gsheets_list_tabs": (True, False, True, True),
    "gsheets_read_range": (True, False, True, True),
    "gsheets_write_range": (False, True, False, True),
    "gsheets_append_rows": (False, False, False, True),
    "gsheets_create_spreadsheet": (False, False, False, True),
    "gsheets_copy_spreadsheet": (False, False, False, True),
    "gsheets_search_spreadsheets": (True, False, True, True),
    "gsheets_clear_range": (False, True, True, True),
    "gsheets_recalculate_range": (False, True, False, True),
}


def _run(awaitable_factory):
    async def runner():
        return await awaitable_factory()

    return anyio.run(runner)


def _list_tools(server) -> list[types.Tool]:
    async def request():
        response = await server.request_handlers[types.ListToolsRequest](
            types.ListToolsRequest()
        )
        return response.root.tools

    return _run(request)


def _call_tool(server, name: str, arguments: dict[str, Any]) -> types.CallToolResult:
    async def request():
        response = await server.request_handlers[types.CallToolRequest](
            types.CallToolRequest(
                params=types.CallToolRequestParams(name=name, arguments=arguments)
            )
        )
        return response.root

    return _run(request)


def test_mode_specific_discovery_is_credential_free(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def credentials_were_loaded(*_args, **_kwargs):
        raise AssertionError("tool discovery must not load credentials")

    monkeypatch.setattr(sheets_client, "_get_credentials", credentials_were_loaded)
    monkeypatch.setattr(sheets_client, "authenticate", credentials_were_loaded)
    monkeypatch.setattr(sheets_client, "get_sheets_service", credentials_were_loaded)

    local = _list_tools(server_module.create_server("local"))
    broker = _list_tools(server_module.create_server("broker"))

    assert tuple(tool.name for tool in local) == LOCAL_TOOLS
    assert tuple(tool.name for tool in broker) == BROKER_TOOLS
    assert len(local) == 9
    assert len(broker) == 8


def test_discovery_exposes_only_canonical_strict_input_schemas() -> None:
    tools = _list_tools(server_module.create_server("local"))

    assert {tool.name for tool in tools} == set(LOCAL_TOOLS)
    assert not any(tool.name.startswith("gsheet_") for tool in tools)
    for tool in tools:
        schema = tool.inputSchema
        assert schema["type"] == "object"
        assert schema["additionalProperties"] is False
        assert set(schema["properties"]) == EXPECTED_INPUT_FIELDS[tool.name]
        assert tool.outputSchema is not None
        assert tool.outputSchema.get("type") != "string"
        assert tool.annotations is not None
        assert (
            tool.annotations.readOnlyHint,
            tool.annotations.destructiveHint,
            tool.annotations.idempotentHint,
            tool.annotations.openWorldHint,
        ) == EXPECTED_ANNOTATIONS[tool.name]
        assert tool.meta["gsheets/contractVersion"] == "1.0"

    write = next(tool for tool in tools if tool.name == "gsheets_write_range")
    assert "spreadsheet_id" not in write.inputSchema["properties"]
    assert "range_a1" not in write.inputSchema["properties"]
    assert "data" not in write.inputSchema["properties"]


def _assert_recursive_objects_are_closed(schema: object) -> None:
    if isinstance(schema, dict):
        if schema.get("type") == "object":
            assert schema.get("additionalProperties") is False
        for value in schema.values():
            _assert_recursive_objects_are_closed(value)
    elif isinstance(schema, list):
        for value in schema:
            _assert_recursive_objects_are_closed(value)


def test_every_input_and_output_object_schema_is_closed_world() -> None:
    for tool in _list_tools(server_module.create_server("local")):
        _assert_recursive_objects_are_closed(tool.inputSchema)
        _assert_recursive_objects_are_closed(tool.outputSchema)


def test_output_schema_requires_wire_discriminators_and_operation() -> None:
    definition = next(
        tool
        for tool in _list_tools(server_module.create_server("local"))
        if tool.name == "gsheets_list_tabs"
    )
    success = {
        "status": "ok",
        "operation": "gsheets_list_tabs",
        "spreadsheet": SPREADSHEET_ID,
        "title": "Operating Comps",
        "tabs": [],
    }
    error = _call_tool(
        server_module.create_server("local"),
        "gsheets_list_tabs",
        {},
    ).structuredContent

    for payload in (success, error):
        jsonschema.validate(payload, definition.outputSchema)
        for required_key in ("status", "operation"):
            missing = {key: value for key, value in payload.items() if key != required_key}
            with pytest.raises(jsonschema.ValidationError):
                jsonschema.validate(missing, definition.outputSchema)


@pytest.mark.parametrize(
    ("tool_name", "arguments"),
    [
        (
            "gsheets_read_range",
            {
                "spreadsheet": SPREADSHEET_ID,
                "range": "Comps!A1",
                "value_render_option": "RAW",
            },
        ),
        (
            "gsheets_read_range",
            {"spreadsheet": 123, "range": "Comps!A1"},
        ),
        ("gsheets_search_spreadsheets", {"query": "comps", "limit": 0}),
        ("gsheets_search_spreadsheets", {"query": "comps", "limit": 101}),
        ("gsheets_search_spreadsheets", {"query": "comps", "limit": "10"}),
        (
            "gsheets_write_range",
            {"spreadsheet": SPREADSHEET_ID, "range": "A1", "values": []},
        ),
        (
            "gsheets_write_range",
            {"spreadsheet": SPREADSHEET_ID, "range": "A1", "values": [[]]},
        ),
        (
            "gsheets_copy_spreadsheet",
            {
                "spreadsheet": SPREADSHEET_ID,
                "title": "Copy",
                "tabs": ["Comps", "Comps"],
            },
        ),
        ("gsheets_create_spreadsheet", {"title": "   "}),
        ("gsheets_search_spreadsheets", {"query": "\t"}),
        (
            "gsheets_read_range",
            {"spreadsheet": SPREADSHEET_ID, "range": "  "},
        ),
        (
            "gsheets_copy_spreadsheet",
            {
                "spreadsheet": SPREADSHEET_ID,
                "title": "Copy",
                "tabs": ["   "],
            },
        ),
    ],
)
def test_strict_types_enums_bounds_and_nested_constraints_reject_before_dispatch(
    tool_name: str,
    arguments: dict[str, Any],
) -> None:
    result = _call_tool(server_module.create_server("local"), tool_name, arguments)

    assert result.isError is True
    assert result.structuredContent["error"]["code"] == "invalid_arguments"
    assert result.structuredContent["error"]["outcome"]["state"] == "not_started"


def test_fixed_examples_validate_against_every_input_schema() -> None:
    for spec in server_module.all_tool_specs():
        definition = spec.definition()
        example = spec.example_arguments.model_dump(mode="json")
        jsonschema.validate(example, definition.inputSchema)
        assert spec.input_model.model_validate(example) == spec.example_arguments


def test_success_is_direct_structured_content_and_not_json_inside_json(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def fake_list_tabs(_arguments):
        return ListTabsSuccess(
            spreadsheet=SPREADSHEET_ID,
            title="Operating Comps",
            tabs=[
                Tab(
                    sheet_id=7,
                    title="Comps",
                    index=0,
                    row_count=200,
                    column_count=20,
                )
            ],
        )

    monkeypatch.setitem(
        server_module.HANDLERS,
        "gsheets_list_tabs",
        fake_list_tabs,
    )
    server = server_module.create_server("local")
    result = _call_tool(
        server,
        "gsheets_list_tabs",
        {"spreadsheet": SPREADSHEET_ID},
    )

    assert result.isError is False
    assert result.structuredContent == {
        "status": "ok",
        "operation": "gsheets_list_tabs",
        "spreadsheet": SPREADSHEET_ID,
        "title": "Operating Comps",
        "tabs": [
            {
                "sheet_id": 7,
                "title": "Comps",
                "index": 0,
                "row_count": 200,
                "column_count": 20,
            }
        ],
    }
    text_payload = json.loads(result.content[0].text)
    assert isinstance(text_payload, dict)
    assert text_payload == result.structuredContent

    definition = next(
        tool for tool in _list_tools(server) if tool.name == "gsheets_list_tabs"
    )
    jsonschema.validate(result.structuredContent, definition.outputSchema)


def test_unknown_fields_fail_before_dispatch_without_echoing_input(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def should_not_run(_arguments):
        raise AssertionError("invalid input reached the tool handler")

    monkeypatch.setitem(
        server_module.HANDLERS,
        "gsheets_list_tabs",
        should_not_run,
    )
    supplied_value = "2ThisMustNeverBeEchoedBack987654321"
    result = _call_tool(
        server_module.create_server("local"),
        "gsheets_list_tabs",
        {
            "spreadsheet": supplied_value,
            "spreadsheet_id": "legacy-field-must-be-rejected",
        },
    )

    assert result.isError is True
    payload = result.structuredContent
    assert payload["status"] == "error"
    assert payload["operation"] == "gsheets_list_tabs"
    assert payload["error"]["code"] == "invalid_arguments"
    assert payload["error"]["outcome"] == {
        "state": "not_started",
        "phase": "validation",
        "mutation_may_have_occurred": False,
    }
    validation = payload["error"]["validation"]
    assert validation["allowed_fields"] == ["spreadsheet"]
    assert validation["required_fields"] == ["spreadsheet"]
    assert validation["issues"] == [
        {
            "path": "spreadsheet_id",
            "code": "unknown_field",
            "message": "Remove this field.",
        }
    ]
    serialized = json.dumps(payload, sort_keys=True)
    assert supplied_value not in serialized
    assert "legacy-field-must-be-rejected" not in serialized


def test_json_encoded_values_are_rejected_instead_of_compat_parsed(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def should_not_run(_arguments):
        raise AssertionError("JSON-string compatibility parsing is forbidden")

    monkeypatch.setitem(
        server_module.HANDLERS,
        "gsheets_write_range",
        should_not_run,
    )
    result = _call_tool(
        server_module.create_server("local"),
        "gsheets_write_range",
        {
            "spreadsheet": SPREADSHEET_ID,
            "range": "Comps!A1:B2",
            "values": '[["Ticker","PCTY"]]',
        },
    )

    assert result.isError is True
    payload = result.structuredContent
    assert payload["error"]["code"] == "invalid_arguments"
    assert payload["error"]["validation"]["issues"] == [
        {
            "path": "values",
            "code": "invalid_type",
            "message": "Use the type declared by the tool schema.",
        }
    ]


def test_title_reference_is_rejected_before_credentials(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def credentials_were_loaded(*_args, **_kwargs):
        raise AssertionError("reference validation must precede authentication")

    monkeypatch.setattr(
        sheets_client,
        "get_sheets_service",
        credentials_were_loaded,
    )
    result = _call_tool(
        server_module.create_server("local"),
        "gsheets_list_tabs",
        {"spreadsheet": "Quarterly Comps"},
    )

    assert result.isError is True
    assert result.structuredContent["error"]["code"] == (
        "invalid_spreadsheet_reference"
    )
    assert result.structuredContent["error"]["outcome"] == {
        "state": "not_started",
        "phase": "reference_validation",
        "mutation_may_have_occurred": False,
    }


def test_mutation_failure_is_structured_and_marks_uncertain_outcome(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def fail_after_dispatch(_arguments):
        raise sheets_client.SheetsClientError(
            "operation_outcome_uncertain",
            "The operation may have changed external state, but confirmation was lost.",
            outcome_state="uncertain",
            phase="write_range",
            mutation_may_have_occurred=True,
            retry_safe=False,
            retry_automatic=False,
            retry_action="inspect_before_retry",
        )

    monkeypatch.setitem(
        server_module.HANDLERS,
        "gsheets_write_range",
        fail_after_dispatch,
    )
    server = server_module.create_server("local")
    result = _call_tool(
        server,
        "gsheets_write_range",
        {
            "spreadsheet": SPREADSHEET_ID,
            "range": "Comps!A1",
            "values": [["PCTY"]],
        },
    )

    assert result.isError is True
    payload = result.structuredContent
    assert payload["status"] == "error"
    assert payload["operation"] == "gsheets_write_range"
    assert payload["error"]["code"] == "operation_outcome_uncertain"
    assert payload["error"]["outcome"] == {
        "state": "uncertain",
        "phase": "write_range",
        "mutation_may_have_occurred": True,
    }
    assert payload["error"]["retry"] == {
        "safe": False,
        "automatic": False,
        "action": "inspect_before_retry",
        "retry_after_seconds": None,
    }
    definition = next(
        tool for tool in _list_tools(server) if tool.name == "gsheets_write_range"
    )
    jsonschema.validate(payload, definition.outputSchema)


def test_copy_partial_error_exposes_destination_progress(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    recovery = {
        "kind": "copy_progress",
        "destination_spreadsheet": "destination-spreadsheet-id",
        "destination_url": (
            "https://docs.google.com/spreadsheets/d/destination-spreadsheet-id"
        ),
        "confirmed_tabs": ["Comps"],
        "active_tab": "Assumptions",
        "active_tab_state": "uncertain",
        "remaining_tabs": ["Assumptions"],
        "finalization_state": "not_started",
    }

    def fail_partially(_arguments):
        raise sheets_client.SheetsClientError(
            "operation_partial",
            "The operation changed external state but did not finish.",
            outcome_state="partial",
            phase="copy_tabs",
            mutation_may_have_occurred=True,
            retry_safe=False,
            retry_automatic=False,
            retry_action="inspect_partial_state",
            recovery=recovery,
        )

    monkeypatch.setitem(
        server_module.HANDLERS,
        "gsheets_copy_spreadsheet",
        fail_partially,
    )
    server = server_module.create_server("local")
    result = _call_tool(
        server,
        "gsheets_copy_spreadsheet",
        {
            "spreadsheet": SPREADSHEET_ID,
            "title": "[hank] Working Copy",
            "tabs": ["Comps", "Assumptions"],
        },
    )

    assert result.isError is True
    assert result.structuredContent["error"]["outcome"]["state"] == "partial"
    assert result.structuredContent["error"]["retry"]["safe"] is False
    assert result.structuredContent["error"]["recovery"] == recovery
    definition = next(
        tool for tool in _list_tools(server) if tool.name == "gsheets_copy_spreadsheet"
    )
    jsonschema.validate(result.structuredContent, definition.outputSchema)


def test_unknown_tool_returns_typed_helpful_error() -> None:
    result = _call_tool(
        server_module.create_server("broker"),
        "gsheets_search_spreadsheet",
        {"query": "comps"},
    )

    assert result.isError is True
    assert result.structuredContent["error"]["code"] == "unknown_tool"
    assert "gsheets_search_spreadsheets" in result.structuredContent["error"]["message"]
    assert "unavailable in this mode" in result.structuredContent["error"]["message"]
    assert "gsheets_search_spreadsheets" not in {
        tool.name for tool in _list_tools(server_module.create_server("broker"))
    }


def test_retired_tool_namespace_is_not_behaviorally_recognized() -> None:
    result = _call_tool(
        server_module.create_server("local"),
        "gsheet_read_range",
        {"spreadsheet": SPREADSHEET_ID, "range": "Comps!A1"},
    )

    assert result.isError is True
    assert result.structuredContent["error"]["code"] == "unknown_tool"
    assert result.structuredContent["error"]["message"] == (
        "Unknown Google Sheets tool."
    )
    assert "gsheets_read_range" not in result.content[0].text


def test_broker_direct_search_call_returns_mode_aware_capability_error() -> None:
    result = _call_tool(
        server_module.create_server("broker"),
        "gsheets_search_spreadsheets",
        {"query": "comps"},
    )

    assert result.isError is True
    assert result.structuredContent["status"] == "error"
    assert result.structuredContent["operation"] == "gsheets_search_spreadsheets"
    assert result.structuredContent["error"]["code"] == "capability_unavailable"
    assert result.structuredContent["error"]["outcome"]["state"] == "not_started"
    assert result.structuredContent["error"]["retry"] == {
        "safe": False,
        "automatic": False,
        "action": "use_spreadsheet_id_or_url",
        "retry_after_seconds": None,
    }


def test_real_stdio_initialize_list_and_validation_call(tmp_path: Path) -> None:
    """Exercise actual JSON-RPC framing without touching Google credentials."""

    repo_root = Path(__file__).resolve().parents[1]
    src_path = repo_root / "src"
    stderr_path = tmp_path / "server-stderr.log"

    async def scenario():
        environment = dict(os.environ)
        environment.update(
            {
                "PYTHONPATH": str(src_path),
                "GSHEETS_TOKEN_MODE": "broker",
            }
        )
        parameters = StdioServerParameters(
            command=sys.executable,
            args=["-m", "gsheets_mcp", "serve"],
            cwd=repo_root,
            env=environment,
        )
        with stderr_path.open("w+") as errlog:
            with anyio.fail_after(15):
                async with stdio_client(parameters, errlog=errlog) as (
                    read_stream,
                    write_stream,
                ):
                    async with ClientSession(read_stream, write_stream) as session:
                        initialized = await session.initialize()
                        listed = await session.list_tools()
                        called = await session.call_tool(
                            "gsheets_read_range",
                            {
                                "spreadsheet": SPREADSHEET_ID,
                                "range": "Comps!A1:B2",
                                "worksheet_name": "legacy-field",
                            },
                        )
        return initialized, listed, called

    initialized, listed, called = anyio.run(scenario)

    assert initialized.serverInfo.name == "gsheets-mcp"
    assert initialized.serverInfo.version == __version__
    assert initialized.capabilities.experimental == {
        "gsheets": {"contract_version": "1.0"}
    }
    assert tuple(tool.name for tool in listed.tools) == BROKER_TOOLS
    assert called.isError is True
    assert called.structuredContent["status"] == "error"
    assert called.structuredContent["operation"] == "gsheets_read_range"
    assert called.structuredContent["error"]["code"] == "invalid_arguments"
    assert called.structuredContent["error"]["validation"]["issues"] == [
        {
            "path": "worksheet_name",
            "code": "unknown_field",
            "message": "Remove this field.",
        }
    ]
    assert isinstance(json.loads(called.content[0].text), dict)


def test_real_stdio_success_is_direct_structured_content(tmp_path: Path) -> None:
    """Run the real wire server with an offline test handler and verify success."""

    repo_root = Path(__file__).resolve().parents[1]
    src_path = repo_root / "src"
    stderr_path = tmp_path / "server-success-stderr.log"
    child_script = f"""
import anyio
from gsheets_mcp import server
from gsheets_mcp.contracts import ListTabsSuccess, Tab

def offline_list_tabs(_arguments):
    return ListTabsSuccess(
        spreadsheet={SPREADSHEET_ID!r},
        title="Offline Fixture",
        tabs=[Tab(sheet_id=1, title="Comps", index=0, row_count=20, column_count=8)],
    )

server.HANDLERS["gsheets_list_tabs"] = offline_list_tabs
anyio.run(server.run_stdio, "broker")
"""

    async def scenario():
        environment = dict(os.environ)
        environment.update(
            {"PYTHONPATH": str(src_path), "GSHEETS_TOKEN_MODE": "broker"}
        )
        parameters = StdioServerParameters(
            command=sys.executable,
            args=["-c", child_script],
            cwd=repo_root,
            env=environment,
        )
        with stderr_path.open("w+") as errlog:
            with anyio.fail_after(15):
                async with stdio_client(parameters, errlog=errlog) as (
                    read_stream,
                    write_stream,
                ):
                    async with ClientSession(read_stream, write_stream) as session:
                        await session.initialize()
                        return await session.call_tool(
                            "gsheets_list_tabs",
                            {"spreadsheet": SPREADSHEET_ID},
                        )

    result = anyio.run(scenario)

    assert result.isError is False
    assert result.structuredContent == {
        "status": "ok",
        "operation": "gsheets_list_tabs",
        "spreadsheet": SPREADSHEET_ID,
        "title": "Offline Fixture",
        "tabs": [
            {
                "sheet_id": 1,
                "title": "Comps",
                "index": 0,
                "row_count": 20,
                "column_count": 8,
            }
        ],
    }
    assert json.loads(result.content[0].text) == result.structuredContent
