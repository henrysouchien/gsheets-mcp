"""Low-level MCP server with strict, typed Google Sheets contracts."""

from __future__ import annotations

import difflib
import json
import logging
import re
import uuid
from dataclasses import dataclass
from typing import Any, Callable

import anyio
from mcp import types
from mcp.server.lowlevel import NotificationOptions, Server
from mcp.server.stdio import stdio_server
from pydantic import BaseModel, TypeAdapter, ValidationError

from . import CONTRACT_VERSION, __version__, sheets_client
from .contracts import (
    ERROR_MODELS,
    INPUT_MODELS,
    RESULT_ADAPTERS,
    ErrorEnvelope,
    OperationOutcome,
    RetryInstruction,
    StrictModel,
    ValidationDetails,
    ValidationIssue,
)
from .tools import HANDLERS

logger = logging.getLogger(__name__)

_SAFE_FIELD_RE = re.compile(r"[^A-Za-z0-9_.\[\]-]+")
_DUMMY_SPREADSHEET = "1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789"
_MAX_VALIDATION_ISSUES = 20
_MAX_VALIDATION_PATH_LENGTH = 256


@dataclass(frozen=True)
class ToolSpec:
    """Single source of truth for discovery, validation, dispatch, and output."""

    name: str
    title: str
    description: str
    input_model: type[StrictModel]
    result_adapter: TypeAdapter
    error_model: type[ErrorEnvelope]
    handler: Callable[[StrictModel], StrictModel]
    example_arguments: StrictModel
    annotations: types.ToolAnnotations
    mutation_class: str

    def definition(self) -> types.Tool:
        return types.Tool(
            name=self.name,
            title=self.title,
            description=self.description,
            inputSchema=self.input_model.model_json_schema(),
            # Results are serialized model instances, so defaulted discriminator
            # fields such as status and operation are always present on the wire.
            # Serialization mode keeps the advertised schema exact by requiring
            # those fields instead of merely documenting their defaults.
            outputSchema=self.result_adapter.json_schema(mode="serialization"),
            annotations=self.annotations,
            _meta={
                "gsheets/contractVersion": CONTRACT_VERSION,
                "gsheets/mutationClass": self.mutation_class,
            },
        )

    @property
    def read_only(self) -> bool:
        return self.annotations.readOnlyHint is True


def _annotations(
    *,
    title: str,
    read_only: bool,
    destructive: bool,
    idempotent: bool,
) -> types.ToolAnnotations:
    return types.ToolAnnotations(
        title=title,
        readOnlyHint=read_only,
        destructiveHint=destructive,
        idempotentHint=idempotent,
        openWorldHint=True,
    )


def _spec(
    name: str,
    *,
    title: str,
    description: str,
    example: dict[str, Any],
    read_only: bool,
    destructive: bool,
    idempotent: bool,
    mutation_class: str,
) -> ToolSpec:
    input_model = INPUT_MODELS[name]
    return ToolSpec(
        name=name,
        title=title,
        description=description,
        input_model=input_model,
        result_adapter=RESULT_ADAPTERS[name],
        error_model=ERROR_MODELS[name],
        handler=HANDLERS[name],
        example_arguments=input_model.model_validate(example),
        annotations=_annotations(
            title=title,
            read_only=read_only,
            destructive=destructive,
            idempotent=idempotent,
        ),
        mutation_class=mutation_class,
    )


def all_tool_specs() -> tuple[ToolSpec, ...]:
    """Return every tool in stable display order without loading credentials."""
    return (
        _spec(
            "gsheets_list_tabs",
            title="List spreadsheet tabs",
            description=(
                "List tab titles and grid dimensions for a spreadsheet URL or ID; "
                "titles are not valid spreadsheet references. Use this before "
                "gsheets_read_range "
                "when the tab is unknown. The returned spreadsheet value is a normalized "
                "ID reusable by every sibling tool. "
                'Example: {"spreadsheet": "<spreadsheet-id-or-url>"}.'
            ),
            example={"spreadsheet": _DUMMY_SPREADSHEET},
            read_only=True,
            destructive=False,
            idempotent=True,
            mutation_class="read",
        ),
        _spec(
            "gsheets_read_range",
            title="Read a spreadsheet range",
            description=(
                "Read an A1 range from a spreadsheet URL or ID. Use gsheets_list_tabs first "
                "when the tab is unknown; pass the tab inside range, not as sheet_name. "
                "FORMULA returns formula text; formatted and unformatted modes return "
                "rendered values. Use gsheets_write_range or gsheets_append_rows to mutate. Example: "
                '{"spreadsheet": "<spreadsheet-id-or-url>", "range": "Sheet1!A1:D20"}.'
            ),
            example={"spreadsheet": _DUMMY_SPREADSHEET, "range": "Sheet1!A1:D20"},
            read_only=True,
            destructive=False,
            idempotent=True,
            mutation_class="read",
        ),
        _spec(
            "gsheets_write_range",
            title="Write a spreadsheet range",
            description=(
                "Overwrite an A1 range with USER_ENTERED values. This is destructive "
                "and approval-gated by the gateway; inspect the target first. A lost "
                "transport response has an uncertain outcome and must not be retried "
                "blindly. Use gsheets_append_rows to add rows instead; values must be a JSON "
                'matrix, not a string. Example: {"spreadsheet": "<id-or-url>", '
                '"range": "Sheet1!A1:B2", "values": [["Ticker", "PCTY"]]}.'
            ),
            example={
                "spreadsheet": _DUMMY_SPREADSHEET,
                "range": "Sheet1!A1:B2",
                "values": [["Ticker", "PCTY"], ["Revenue", 100.0]],
            },
            read_only=False,
            destructive=True,
            idempotent=False,
            mutation_class="overwrite",
        ),
        _spec(
            "gsheets_append_rows",
            title="Append spreadsheet rows",
            description=(
                "Append row-major values using Google Sheets table detection. This is "
                "non-idempotent: never repeat it after an uncertain response. Use "
                "gsheets_write_range for exact-cell replacement and read the table first. Values "
                'must be a JSON matrix, not a string. Example: {"spreadsheet": '
                '"<id-or-url>", "range": "Sheet1!A:B", "values": [["PCTY", 100]]}.'
            ),
            example={
                "spreadsheet": _DUMMY_SPREADSHEET,
                "range": "Sheet1!A:B",
                "values": [["PCTY", 100.0]],
            },
            read_only=False,
            destructive=False,
            idempotent=False,
            mutation_class="append",
        ),
        _spec(
            "gsheets_create_spreadsheet",
            title="Create a spreadsheet",
            description=(
                "Create a new spreadsheet file, not a tab. This is non-idempotent; "
                "retain the returned spreadsheet ID and never repeat after an uncertain "
                "response without checking Drive. Use gsheets_copy_spreadsheet when starting "
                'from existing tabs. Example: {"title": "[hank] Operating Comps - 2026-07-14"}.'
            ),
            example={"title": "[hank] Operating Comps - 2026-07-14"},
            read_only=False,
            destructive=False,
            idempotent=False,
            mutation_class="create",
        ),
        _spec(
            "gsheets_copy_spreadsheet",
            title="Copy a spreadsheet",
            description=(
                "Copy all or selected source tabs into a new spreadsheet file. The "
                "spreadsheet argument is the source and title names the destination. "
                "This is non-idempotent. Partial errors retain the destination ID and "
                "confirmed progress; inspect that destination instead of rerunning. Use "
                'gsheets_create_spreadsheet for a blank file. Example: {"spreadsheet": '
                '"<source-id-or-url>", "title": "Working Copy", "tabs": ["Comps"]}.'
            ),
            example={
                "spreadsheet": _DUMMY_SPREADSHEET,
                "title": "[hank] Working Copy - 2026-07-14",
                "tabs": ["Assumptions", "Comps"],
            },
            read_only=False,
            destructive=False,
            idempotent=False,
            mutation_class="multi_stage_create",
        ),
        _spec(
            "gsheets_search_spreadsheets",
            title="Search spreadsheets",
            description=(
                "Search local-mode Google Drive spreadsheet titles and return normalized "
                "IDs reusable as spreadsheet arguments. This searches files, not cell "
                "contents. Select results[n].spreadsheet, then use gsheets_list_tabs or "
                "gsheets_read_range; "
                "do not pass the title to sibling tools. It is absent from broker mode. "
                'Example: {"query": "operating comps", "limit": 10}.'
            ),
            example={"query": "operating comps", "limit": 10},
            read_only=True,
            destructive=False,
            idempotent=True,
            mutation_class="read",
        ),
        _spec(
            "gsheets_clear_range",
            title="Clear a spreadsheet range",
            description=(
                "Clear values from an A1 range without deleting cells or formatting. "
                "This is destructive and approval-gated; read the target first. Repeating "
                "a confirmed clear is idempotent, but an uncertain response still requires "
                "inspection. Use gsheets_write_range to replace rather than clear. Example: "
                '{"spreadsheet": "<id-or-url>", "range": "Sheet1!A2:D20"}.'
            ),
            example={"spreadsheet": _DUMMY_SPREADSHEET, "range": "Sheet1!A2:D20"},
            read_only=False,
            destructive=True,
            idempotent=True,
            mutation_class="clear",
        ),
        _spec(
            "gsheets_recalculate_range",
            title="Recalculate a formula range",
            description=(
                "Request custom-function recalculation by restoring a formula-only A1 "
                "range. Blank cells are allowed; any literal value rejects the call "
                "before clearing. The operation verifies formulas and performs at most "
                "one exact compensation write; it does not claim values are fresh. Use "
                "gsheets_read_range with FORMULA to inspect first. It is destructive and "
                'non-idempotent. Example: {"spreadsheet": "<id-or-url>", '
                '"range": "Comps!B2:H20"}.'
            ),
            example={"spreadsheet": _DUMMY_SPREADSHEET, "range": "Comps!B2:H20"},
            read_only=False,
            destructive=True,
            idempotent=False,
            mutation_class="recalculate",
        ),
    )


def tool_specs(mode: str | None = None) -> tuple[ToolSpec, ...]:
    """Return the credential-free, mode-specific registry."""
    selected_mode = sheets_client.token_mode() if mode is None else mode
    if selected_mode not in {"local", "broker"}:
        raise sheets_client.SheetsClientError(
            "invalid_configuration",
            "GSHEETS_TOKEN_MODE must be either 'local' or 'broker'.",
        )
    specs = all_tool_specs()
    if selected_mode == "broker":
        specs = tuple(
            spec for spec in specs if spec.name != "gsheets_search_spreadsheets"
        )
    return specs


def capabilities(mode: str | None = None) -> dict[str, Any]:
    """Return non-secret capability JSON derived from the live registry."""
    selected_mode = sheets_client.token_mode() if mode is None else mode
    specs = tool_specs(selected_mode)
    return {
        "package": "gsheets-mcp",
        "package_version": __version__,
        "contract_version": CONTRACT_VERSION,
        "mode": selected_mode,
        "tool_count": len(specs),
        "tools": [
            spec.definition().model_dump(mode="json", by_alias=True) for spec in specs
        ],
        "environment": {
            "GSHEETS_TOKEN_MODE": {
                "required": False,
                "values": ["local", "broker"],
                "default": "local",
            },
            "local": {
                "GOOGLE_CREDENTIALS_FILE": "optional path override",
                "GOOGLE_TOKEN_FILE": "optional path override",
                "GSHEETS_HEADLESS": "set to 1 to disable interactive consent",
            },
            "broker": {
                "GSHEETS_BROKER_URL": "required at call time",
                "GSHEETS_BROKER_SESSION_TOKEN": "required at call time",
            },
        },
    }


def _safe_path(location: tuple[Any, ...]) -> str:
    parts: list[str] = []
    for value in location:
        if isinstance(value, int):
            parts.append(f"[{value}]")
            continue
        cleaned = _SAFE_FIELD_RE.sub("_", str(value))[:64] or "field"
        if parts and not cleaned.startswith("["):
            parts.append(".")
        parts.append(cleaned)
    return ("".join(parts) or "arguments")[:_MAX_VALIDATION_PATH_LENGTH]


def _validation_issue(error: dict[str, Any]) -> ValidationIssue:
    error_type = str(error.get("type") or "")
    if error_type == "extra_forbidden":
        code = "unknown_field"
        message = "Remove this field."
    elif error_type == "missing":
        code = "required_field"
        message = "Add this required field."
    elif "type" in error_type or error_type.endswith("_parsing"):
        code = "invalid_type"
        message = "Use the type declared by the tool schema."
    else:
        code = "invalid_value"
        message = "Use a value allowed by the tool schema."
    return ValidationIssue(
        path=_safe_path(tuple(error.get("loc") or ())),
        code=code,
        message=message,
    )


def _validation_error(spec: ToolSpec, error: ValidationError) -> ErrorEnvelope:
    issues = [
        _validation_issue(item)
        for item in error.errors(
            include_input=False,
            include_url=False,
            include_context=False,
        )[:_MAX_VALIDATION_ISSUES]
    ]
    allowed_fields = list(spec.input_model.model_fields)
    required_fields = [
        name
        for name, field in spec.input_model.model_fields.items()
        if field.is_required()
    ]
    return spec.error_model(
        operation=spec.name,
        error={
            "code": "invalid_arguments",
            "message": "Arguments do not match the tool contract.",
            "outcome": OperationOutcome(
                state="not_started",
                phase="validation",
                mutation_may_have_occurred=False,
            ),
            "retry": RetryInstruction(
                safe=True,
                automatic=False,
                action="correct_arguments",
            ),
            "validation": ValidationDetails[spec.input_model](
                issues=issues,
                allowed_fields=allowed_fields,
                required_fields=required_fields,
                example_arguments=spec.example_arguments,
            ),
        },
    )


def _safe_http_error(error: Exception) -> tuple[str, str, str]:
    status = getattr(getattr(error, "resp", None), "status", None)
    if status == 401:
        return (
            "google_api_unauthorized",
            "Google Sheets authorization failed after the bounded credential refresh.",
            "reconnect_google_sheets",
        )
    if status == 403:
        return (
            "google_api_forbidden",
            "Google Sheets rejected the request because access is not permitted.",
            "verify_permissions",
        )
    if status == 404:
        return (
            "spreadsheet_not_found",
            "The requested spreadsheet or range was not found.",
            "verify_spreadsheet_and_range",
        )
    if status == 429:
        return (
            "google_rate_limited",
            "Google Sheets rate limited the request.",
            "retry_later",
        )
    return (
        "google_api_error",
        "Google Sheets could not complete the request.",
        "retry_after_inspection",
    )


def _application_error(spec: ToolSpec, error: Exception) -> ErrorEnvelope:
    incident_id: str | None = None
    recovery = None
    retry_after = None

    if isinstance(error, sheets_client.SheetsClientError):
        details = error.details
        code = error.code
        message = str(error)
        state = details.get("outcome_state", "not_started")
        phase = details.get("phase", "prepare_request")
        mutation_may_have_occurred = bool(
            details.get("mutation_may_have_occurred", False)
        )
        retry_safe = bool(
            details.get("retry_safe", state in {"not_started", "unchanged"})
        )
        retry_automatic = bool(
            details.get(
                "retry_automatic",
                spec.read_only
                and code == "broker_session_expired"
                and state == "not_started",
            )
        )
        retry_action = str(details.get("retry_action") or "inspect_error")
        retry_after = details.get("retry_after_s")
        recovery = details.get("recovery")
    elif isinstance(error, sheets_client.HttpError):
        code, message, retry_action = _safe_http_error(error)
        state = "unchanged"
        phase = "google_request"
        mutation_may_have_occurred = False
        retry_safe = spec.read_only
        retry_automatic = False
    else:
        incident_id = uuid.uuid4().hex
        logger.error(
            "Unexpected Sheets tool failure; operation=%s incident_id=%s exception_type=%s",
            spec.name,
            incident_id,
            type(error).__name__,
        )
        code = "internal_error"
        message = "The tool encountered an unexpected internal error."
        state = "not_started" if spec.read_only else "uncertain"
        phase = "execute"
        mutation_may_have_occurred = not spec.read_only
        retry_safe = spec.read_only
        retry_automatic = False
        retry_action = "report_incident"

    return spec.error_model(
        operation=spec.name,
        error={
            "code": code,
            "message": message,
            "outcome": {
                "state": state,
                "phase": phase,
                "mutation_may_have_occurred": mutation_may_have_occurred,
            },
            "retry": {
                "safe": retry_safe,
                "automatic": retry_automatic,
                "action": retry_action,
                "retry_after_seconds": retry_after,
            },
            "recovery": recovery,
            "incident_id": incident_id,
        },
    )


def _call_result(
    spec: ToolSpec, payload: BaseModel, *, is_error: bool
) -> types.CallToolResult:
    validated = spec.result_adapter.validate_python(payload)
    structured = validated.model_dump(mode="json")
    return types.CallToolResult(
        content=[
            types.TextContent(
                type="text",
                text=json.dumps(structured, sort_keys=True, separators=(",", ":")),
            )
        ],
        structuredContent=structured,
        isError=is_error,
    )


def _unknown_tool_result(
    name: str,
    available: tuple[str, ...],
    unavailable: tuple[str, ...] = (),
) -> types.CallToolResult:
    safe_name = _SAFE_FIELD_RE.sub("_", name)[:80] or "unknown"
    # The v1 cutover deliberately does not recognize retired namespaces. Only
    # help callers that have already selected the canonical ``gsheets_``
    # vocabulary and made a typo within it.
    match = (
        difflib.get_close_matches(
            safe_name,
            (*available, *unavailable),
            n=1,
            cutoff=0.55,
        )
        if safe_name.startswith("gsheets_")
        else []
    )
    message = "Unknown Google Sheets tool."
    if match:
        if match[0] in unavailable:
            message += (
                f" The closest canonical tool, {match[0]}, is unavailable in this mode."
            )
        else:
            message += f" Did you mean {match[0]}?"
    structured = {
        "status": "error",
        "operation": safe_name,
        "error": {
            "code": "unknown_tool",
            "message": message,
            "outcome": {
                "state": "not_started",
                "phase": "dispatch",
                "mutation_may_have_occurred": False,
            },
            "retry": {
                "safe": True,
                "automatic": False,
                "action": "choose_advertised_tool",
                "retry_after_seconds": None,
            },
            "validation": None,
            "recovery": None,
            "incident_id": None,
        },
    }
    return types.CallToolResult(
        content=[
            types.TextContent(type="text", text=json.dumps(structured, sort_keys=True))
        ],
        structuredContent=structured,
        isError=True,
    )


def create_server(mode: str | None = None) -> Server:
    """Create one immutable mode-specific MCP server."""
    selected_mode = sheets_client.token_mode() if mode is None else mode
    specs = tool_specs(selected_mode)
    by_name = {spec.name: spec for spec in specs}
    unavailable_by_name = {
        spec.name: spec for spec in all_tool_specs() if spec.name not in by_name
    }
    server = Server(
        "gsheets-mcp",
        version=__version__,
        instructions=(
            "Strict Google Sheets tools. Pass spreadsheet IDs or docs.google.com URLs; "
            "use range for A1 notation. Results and errors are direct typed objects."
        ),
    )

    @server.list_tools()
    async def list_tools() -> list[types.Tool]:
        return [spec.definition() for spec in specs]

    @server.call_tool(validate_input=False)
    async def call_tool(name: str, arguments: dict[str, Any]) -> types.CallToolResult:
        spec = by_name.get(name)
        if spec is None:
            unavailable_spec = unavailable_by_name.get(name)
            if unavailable_spec is not None:
                error = sheets_client.SheetsClientError(
                    "capability_unavailable",
                    "Spreadsheet title search is unavailable in broker mode; use a spreadsheet ID or URL.",
                    outcome_state="not_started",
                    phase="capability_check",
                    mutation_may_have_occurred=False,
                    retry_safe=False,
                    retry_automatic=False,
                    retry_action="use_spreadsheet_id_or_url",
                )
                return _call_result(
                    unavailable_spec,
                    _application_error(unavailable_spec, error),
                    is_error=True,
                )
            return _unknown_tool_result(
                name,
                tuple(by_name),
                tuple(unavailable_by_name),
            )
        try:
            validated_arguments = spec.input_model.model_validate(arguments)
        except ValidationError as exc:
            return _call_result(spec, _validation_error(spec, exc), is_error=True)
        try:
            payload = await anyio.to_thread.run_sync(spec.handler, validated_arguments)
        except Exception as exc:
            return _call_result(spec, _application_error(spec, exc), is_error=True)
        try:
            return _call_result(spec, payload, is_error=False)
        except Exception as exc:
            return _call_result(spec, _application_error(spec, exc), is_error=True)

    return server


async def run_stdio(mode: str | None = None) -> None:
    """Run the strict server over stdio."""
    server = create_server(mode)
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(
                notification_options=NotificationOptions(),
                experimental_capabilities={
                    "gsheets": {"contract_version": CONTRACT_VERSION},
                },
            ),
        )
