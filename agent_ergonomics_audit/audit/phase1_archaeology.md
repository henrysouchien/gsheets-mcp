# Phase 1 Archaeology

## Baseline

- Target HEAD: `b72d8a3420fb417d7842b83ce87dd4961d74d10f`
- Installed MCP SDK: `mcp 1.26.0`
- Verified command: `venv/bin/python -m pytest -q` — 41 passed
- Verified command: `venv/bin/python scripts/mcp_lint.py` — exit 0, no findings
- Bare `pytest` is not a reliable command in the shared workspace: the generic
  package name `src` can resolve to another adjacent editable project. The package
  namespace is therefore part of the cutover, not a cosmetic cleanup.

## Architecture

`src.server` creates one FastMCP stdio server at import time and registers nine
tools. Each wrapper resolves a spreadsheet, obtains a Google service, calls a
helper in `src.sheets_client`, and serializes either success or failure to a JSON
string. Local mode loads OAuth credentials from disk and has Drive scope. Broker
mode receives an opaque gateway session token, exchanges it for a short-lived
Sheets access token, keeps that token in memory, and limits Google auth refresh to
one failed-request retry.

The installed `gsheets-mcp` console script and `run_server.py` both start the MCP
protocol immediately. There is no CLI help/version/capability command, MCP resource,
or prompt. The checked-in README documents only the source launcher.

## Current tool contract

| Tool | Input keys | Encoded success fields | Mutation |
|---|---|---|---|
| `gsheet_list_tabs` | `spreadsheet` | `status`, `spreadsheet_id`, `title`, `tabs` | no |
| `gsheet_read_range` | `spreadsheet`, `cell_range`, `value_render_option`, `date_time_render_option` | `status`, `spreadsheet_id`, `range`, render options, `values` | no |
| `gsheet_update_range` | `spreadsheet`, `cell_range`, `values` | `status`, `updatedRange`, `updatedCells` | overwrite |
| `gsheet_append_rows` | `spreadsheet`, `cell_range`, `values` | `status`, `updatedRange`, `updatedCells` | append; non-idempotent |
| `gsheet_create` | `title` | `status`, `spreadsheet_id`, `url` | create; non-idempotent |
| `gsheet_copy_spreadsheet` | `source`, `new_title`, `tabs` | `status`, `spreadsheet_id`, `title`, `url`, `copied_tabs`, `warnings` | multi-stage create; non-idempotent |
| `gsheet_search` | `query`, `max_results` | `status`, `query`, `results`, `count` | no; local mode only |
| `gsheet_clear_range` | `spreadsheet`, `cell_range` | `status`, `spreadsheet_id`, `clearedRange` | destructive; idempotent |
| `gsheet_touch_range` | `spreadsheet`, `cell_range` | `status`, `spreadsheet_id`, `touchedRange`, `touchedCells` | read→clear→rewrite; partial-failure risk |

The actual MCP output schema for all nine tools is only
`{type: object, properties: {result: {type: string}}, required: [result]}`.
The useful fields exist inside that string and are invisible to schema-aware
clients. Tool annotations, titles, metadata, property descriptions, examples,
string bounds, enum constraints, and `additionalProperties: false` are absent.
Unknown fields are silently discarded when the canonical required fields are also
present.

## Root causes and failure modes

1. **The contract teaches the wrong next call.** `spreadsheet` becomes
   `spreadsheet_id`; `cell_range` becomes `range` or camelCase variants; copy's
   `new_title` becomes `title`. The observed local corpus contains 21 wrong-key
   validation failures across 18 sessions. All recovered after a failed turn, at
   roughly 16 seconds of avoidable latency per failure on average.
2. **Validation failures bypass the application error helper.** FastMCP/Pydantic
   validates before the wrapper runs, producing raw framework errors that repeat
   the caller's input fragment and do not show a corrected call.
3. **Business failures look successful.** `_json_error` returns an ordinary string,
   so MCP emits `isError=false`. The gateway therefore parses most business errors
   as successful tool data.
4. **Broker capabilities are overstated.** `gsheet_search` is always present in
   `tools/list`, although broker mode cannot execute it and has no Drive scope.
5. **Touch can destroy the target temporarily and leak it permanently.** A failed
   rewrite leaves the range cleared, and the warning log records all original
   formulas/literal values. FORMULA rendering also returns ordinary values for
   non-formula cells, so mixed ranges are unsafe.
6. **Copy hides partial creation.** Failure after destination creation can leave an
   orphan or partially copied spreadsheet, while the error loses its destination ID
   and completed-tab progress. Replaying the whole tool can create another copy.
7. **Non-idempotent outcomes are underspecified.** Create and append have no
   dedupe key or explicit indeterminate-outcome marker. The existing bounded
   Google-auth retry is request-level and must not become a whole-tool retry.
8. **Mode documentation is false in local mode.** The tools say a URL is accepted,
   but local resolution treats a URL as an exact Drive title; URL extraction exists
   only inside the broker branch.
9. **The quality gates do not protect the live contract.** The linter misses
   append/copy/clear/touch mutation semantics and treats any try/except JSON envelope
   as sufficient. CI runs the linter only, omits pytest, and does not watch the
   client/helper/tests/package files that define the contract.

## Security and behavior invariants

The cutover must preserve:

- Per-user broker process isolation and positive authenticated user identity.
- In-memory-only broker access tokens, 60-second proactive refresh margin, and one
  failed-request Google 401 refresh rather than whole-operation replay.
- Strict `docs.google.com` URL validation, URL/ID verification through Sheets, and
  no broker title resolution or Drive search.
- Credential-free `tools/list` and capability discovery.
- User-safe auth codes including `sheets_not_connected`, `broker_rate_limited`,
  `broker_session_expired`, `sheets_unavailable`, `interactive_consent_disabled`,
  `invalid_spreadsheet_url`, and `google_api_unauthorized`.
- Local Drive-readonly plus Sheets scopes, shared-drive discovery, copy tab ordering,
  subset validation, and explicit named-range/locale/time-zone copy warnings.
- Gateway approval/redaction/citation ordering and no blind replay of create, copy,
  append, or recalculation operations.

## Cutover implications

The source package must move from generic `src` to a real `gsheets_mcp` namespace,
and the redundant source launcher should be replaced in every tracked first-party
configuration by one installed CLI contract. Tool inputs and results should use one
snake_case vocabulary, errors must set MCP `isError=true` while retaining structured
codes/recovery data, search must be registered only when executable, and both server
and gateway tests must pin the new schemas and structured-result behavior. No alias,
duplicate response key, JSON-string parser, or compatibility package is part of the
target design.
