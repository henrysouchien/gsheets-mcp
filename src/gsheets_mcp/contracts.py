"""Strict input, result, and error contracts for the Google Sheets MCP."""

from __future__ import annotations

from typing import Annotated, Generic, Literal, TypeVar

from pydantic import (
    AllowInfNan,
    BaseModel,
    ConfigDict,
    Field,
    StrictBool,
    StrictFloat,
    StrictInt,
    StrictStr,
    Strict,
    TypeAdapter,
    field_validator,
)


class StrictModel(BaseModel):
    """Base model for wire objects that must reject unknown fields."""

    model_config = ConfigDict(
        extra="forbid",
        strict=True,
        validate_default=True,
        json_schema_serialization_defaults_required=True,
    )


Spreadsheet = Annotated[
    StrictStr,
    Field(
        min_length=1,
        max_length=2048,
        pattern=r"\S",
        description="Google Sheets spreadsheet ID or strict docs.google.com Sheets URL.",
        examples=["1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789"],
    ),
]
RangeA1 = Annotated[
    StrictStr,
    Field(
        min_length=1,
        max_length=1024,
        pattern=r"\S",
        description="A1 notation including the tab when ambiguity is possible.",
        examples=["Sheet1!A1:D20"],
    ),
]
Title = Annotated[
    StrictStr,
    Field(
        min_length=1,
        max_length=200,
        pattern=r"\S",
        description="Title for the new spreadsheet.",
        examples=["[hank] Operating Comps - 2026-07-14"],
    ),
]
FiniteStrictFloat = Annotated[float, Strict(), AllowInfNan(False)]
CellValue = StrictStr | StrictInt | FiniteStrictFloat | StrictBool | None
ValueRow = Annotated[
    list[CellValue],
    Field(min_length=1, description="One non-empty row of JSON cell values."),
]
ValueMatrix = Annotated[
    list[ValueRow],
    Field(
        min_length=1,
        description="Non-empty two-dimensional row-major cell values.",
        examples=[[["Ticker", "PCTY"], ["Revenue", 100.0]]],
    ),
]
ValueRenderOption = Literal["FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"]
DateTimeRenderOption = Literal["FORMATTED_STRING", "SERIAL_NUMBER"]


class ListTabsInput(StrictModel):
    spreadsheet: Spreadsheet


class ReadRangeInput(StrictModel):
    spreadsheet: Spreadsheet
    range: RangeA1
    value_render_option: ValueRenderOption = Field(
        default="FORMATTED_VALUE",
        description="How the Sheets API renders cell values.",
    )
    date_time_render_option: DateTimeRenderOption = Field(
        default="FORMATTED_STRING",
        description="How dates, times, and durations are rendered.",
    )


class WriteRangeInput(StrictModel):
    spreadsheet: Spreadsheet
    range: RangeA1
    values: ValueMatrix


class AppendRowsInput(StrictModel):
    spreadsheet: Spreadsheet
    range: RangeA1
    values: ValueMatrix


class CreateSpreadsheetInput(StrictModel):
    title: Title


class CopySpreadsheetInput(StrictModel):
    spreadsheet: Spreadsheet
    title: Title
    tabs: (
        Annotated[
            list[
                Annotated[
                    StrictStr,
                    Field(min_length=1, max_length=200, pattern=r"\S"),
                ]
            ],
            Field(
                min_length=1,
                description="Optional exact source tab titles to copy, in source order.",
                examples=[["Assumptions", "Comps"]],
            ),
        ]
        | None
    ) = None

    @field_validator("tabs")
    @classmethod
    def tabs_must_be_unique(cls, value: list[str] | None) -> list[str] | None:
        if value is not None and len(set(value)) != len(value):
            raise ValueError("tab titles must be unique")
        return value


class SearchSpreadsheetsInput(StrictModel):
    query: Annotated[
        StrictStr,
        Field(
            min_length=1,
            max_length=200,
            pattern=r"\S",
            description="Case-insensitive spreadsheet-title search text.",
            examples=["operating comps"],
        ),
    ]
    limit: Annotated[
        StrictInt,
        Field(ge=1, le=100, description="Maximum number of matches to return."),
    ] = 10


class ClearRangeInput(StrictModel):
    spreadsheet: Spreadsheet
    range: RangeA1


class RecalculateRangeInput(StrictModel):
    spreadsheet: Spreadsheet
    range: RangeA1


class Tab(StrictModel):
    sheet_id: StrictInt
    title: StrictStr
    index: StrictInt
    row_count: StrictInt
    column_count: StrictInt


class SpreadsheetMatch(StrictModel):
    spreadsheet: StrictStr
    title: StrictStr
    url: StrictStr
    modified_at: StrictStr | None = None


class CopyWarning(StrictModel):
    code: Literal[
        "named_ranges_not_copied",
        "locale_not_preserved",
        "time_zone_not_preserved",
    ]
    message: StrictStr


class ListTabsSuccess(StrictModel):
    status: Literal["ok"] = "ok"
    operation: Literal["gsheets_list_tabs"] = "gsheets_list_tabs"
    spreadsheet: StrictStr
    title: StrictStr
    tabs: list[Tab]


class ReadRangeSuccess(StrictModel):
    status: Literal["ok"] = "ok"
    operation: Literal["gsheets_read_range"] = "gsheets_read_range"
    spreadsheet: StrictStr
    range: StrictStr
    value_render_option: ValueRenderOption
    date_time_render_option: DateTimeRenderOption
    values: list[list[CellValue]]


class WriteRangeSuccess(StrictModel):
    status: Literal["ok"] = "ok"
    operation: Literal["gsheets_write_range"] = "gsheets_write_range"
    spreadsheet: StrictStr
    range: StrictStr
    cell_count: StrictInt


class AppendRowsSuccess(StrictModel):
    status: Literal["ok"] = "ok"
    operation: Literal["gsheets_append_rows"] = "gsheets_append_rows"
    spreadsheet: StrictStr
    range: StrictStr
    cell_count: StrictInt


class CreateSpreadsheetSuccess(StrictModel):
    status: Literal["ok"] = "ok"
    operation: Literal["gsheets_create_spreadsheet"] = "gsheets_create_spreadsheet"
    spreadsheet: StrictStr
    title: StrictStr
    url: StrictStr


class CopySpreadsheetSuccess(StrictModel):
    status: Literal["ok"] = "ok"
    operation: Literal["gsheets_copy_spreadsheet"] = "gsheets_copy_spreadsheet"
    spreadsheet: StrictStr
    title: StrictStr
    url: StrictStr
    tabs: list[StrictStr]
    warnings: list[CopyWarning]


class SearchSpreadsheetsSuccess(StrictModel):
    status: Literal["ok"] = "ok"
    operation: Literal["gsheets_search_spreadsheets"] = "gsheets_search_spreadsheets"
    query: StrictStr
    count: StrictInt
    results: list[SpreadsheetMatch]


class ClearRangeSuccess(StrictModel):
    status: Literal["ok"] = "ok"
    operation: Literal["gsheets_clear_range"] = "gsheets_clear_range"
    spreadsheet: StrictStr
    range: StrictStr


class RecalculateRangeSuccess(StrictModel):
    status: Literal["ok"] = "ok"
    operation: Literal["gsheets_recalculate_range"] = "gsheets_recalculate_range"
    spreadsheet: StrictStr
    range: StrictStr
    cell_count: StrictInt
    recovery_performed: StrictBool
    formulas_verified: StrictBool


OutcomeState = Literal["not_started", "unchanged", "uncertain", "partial", "restored"]


class OperationOutcome(StrictModel):
    state: OutcomeState
    phase: StrictStr
    mutation_may_have_occurred: StrictBool


class RetryInstruction(StrictModel):
    safe: StrictBool
    automatic: StrictBool
    action: StrictStr
    retry_after_seconds: StrictFloat | StrictInt | None = None


class ValidationIssue(StrictModel):
    path: StrictStr
    code: Literal["unknown_field", "required_field", "invalid_type", "invalid_value"]
    message: StrictStr


InputT = TypeVar("InputT", bound=StrictModel)


class ValidationDetails(StrictModel, Generic[InputT]):
    issues: list[ValidationIssue]
    allowed_fields: list[StrictStr]
    required_fields: list[StrictStr]
    example_arguments: InputT


class CopyProgress(StrictModel):
    kind: Literal["copy_progress"] = "copy_progress"
    destination_spreadsheet: StrictStr
    destination_url: StrictStr
    confirmed_tabs: list[StrictStr]
    active_tab: StrictStr | None = None
    active_tab_state: Literal["not_started", "uncertain", "confirmed"]
    remaining_tabs: list[StrictStr]
    finalization_state: Literal["not_started", "uncertain", "confirmed"]


class RangeRecovery(StrictModel):
    kind: Literal["range_state"] = "range_state"
    spreadsheet: StrictStr
    range: StrictStr
    formula_cell_count: StrictInt
    compensation_attempted: StrictBool
    compensation_verified: StrictBool


Recovery = Annotated[CopyProgress | RangeRecovery, Field(discriminator="kind")]


class ErrorDetail(StrictModel, Generic[InputT]):
    code: StrictStr
    message: StrictStr
    outcome: OperationOutcome
    retry: RetryInstruction
    validation: ValidationDetails[InputT] | None = None
    recovery: Recovery | None = None
    incident_id: StrictStr | None = None


class ErrorEnvelope(StrictModel, Generic[InputT]):
    status: Literal["error"] = "error"
    operation: StrictStr
    error: ErrorDetail[InputT]


class ListTabsError(ErrorEnvelope[ListTabsInput]):
    operation: Literal["gsheets_list_tabs"] = "gsheets_list_tabs"


class ReadRangeError(ErrorEnvelope[ReadRangeInput]):
    operation: Literal["gsheets_read_range"] = "gsheets_read_range"


class WriteRangeError(ErrorEnvelope[WriteRangeInput]):
    operation: Literal["gsheets_write_range"] = "gsheets_write_range"


class AppendRowsError(ErrorEnvelope[AppendRowsInput]):
    operation: Literal["gsheets_append_rows"] = "gsheets_append_rows"


class CreateSpreadsheetError(ErrorEnvelope[CreateSpreadsheetInput]):
    operation: Literal["gsheets_create_spreadsheet"] = "gsheets_create_spreadsheet"


class CopySpreadsheetError(ErrorEnvelope[CopySpreadsheetInput]):
    operation: Literal["gsheets_copy_spreadsheet"] = "gsheets_copy_spreadsheet"


class SearchSpreadsheetsError(ErrorEnvelope[SearchSpreadsheetsInput]):
    operation: Literal["gsheets_search_spreadsheets"] = "gsheets_search_spreadsheets"


class ClearRangeError(ErrorEnvelope[ClearRangeInput]):
    operation: Literal["gsheets_clear_range"] = "gsheets_clear_range"


class RecalculateRangeError(ErrorEnvelope[RecalculateRangeInput]):
    operation: Literal["gsheets_recalculate_range"] = "gsheets_recalculate_range"


def _result_adapter(
    success: type[StrictModel], error: type[StrictModel]
) -> TypeAdapter:
    union = Annotated[success | error, Field(discriminator="status")]
    return TypeAdapter(union)


RESULT_ADAPTERS: dict[str, TypeAdapter] = {
    "gsheets_list_tabs": _result_adapter(ListTabsSuccess, ListTabsError),
    "gsheets_read_range": _result_adapter(ReadRangeSuccess, ReadRangeError),
    "gsheets_write_range": _result_adapter(WriteRangeSuccess, WriteRangeError),
    "gsheets_append_rows": _result_adapter(AppendRowsSuccess, AppendRowsError),
    "gsheets_create_spreadsheet": _result_adapter(
        CreateSpreadsheetSuccess, CreateSpreadsheetError
    ),
    "gsheets_copy_spreadsheet": _result_adapter(
        CopySpreadsheetSuccess, CopySpreadsheetError
    ),
    "gsheets_search_spreadsheets": _result_adapter(
        SearchSpreadsheetsSuccess, SearchSpreadsheetsError
    ),
    "gsheets_clear_range": _result_adapter(ClearRangeSuccess, ClearRangeError),
    "gsheets_recalculate_range": _result_adapter(
        RecalculateRangeSuccess, RecalculateRangeError
    ),
}


INPUT_MODELS: dict[str, type[StrictModel]] = {
    "gsheets_list_tabs": ListTabsInput,
    "gsheets_read_range": ReadRangeInput,
    "gsheets_write_range": WriteRangeInput,
    "gsheets_append_rows": AppendRowsInput,
    "gsheets_create_spreadsheet": CreateSpreadsheetInput,
    "gsheets_copy_spreadsheet": CopySpreadsheetInput,
    "gsheets_search_spreadsheets": SearchSpreadsheetsInput,
    "gsheets_clear_range": ClearRangeInput,
    "gsheets_recalculate_range": RecalculateRangeInput,
}


ERROR_MODELS: dict[str, type[ErrorEnvelope]] = {
    "gsheets_list_tabs": ListTabsError,
    "gsheets_read_range": ReadRangeError,
    "gsheets_write_range": WriteRangeError,
    "gsheets_append_rows": AppendRowsError,
    "gsheets_create_spreadsheet": CreateSpreadsheetError,
    "gsheets_copy_spreadsheet": CopySpreadsheetError,
    "gsheets_search_spreadsheets": SearchSpreadsheetsError,
    "gsheets_clear_range": ClearRangeError,
    "gsheets_recalculate_range": RecalculateRangeError,
}
