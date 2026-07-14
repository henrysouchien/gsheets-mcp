# Phase 4 Contract Triangulation

Date: 2026-07-14

## Inputs

Three independent fresh-eyes reviews challenged the audit from different
directions:

1. **Vocabulary and intent:** optimize for the keys and verbs an agent guesses
   before reading source, and make every returned reference directly reusable.
2. **MCP and mutation safety:** use a boundary that can reject unknown fields,
   sanitize validation failures, emit real structured MCP errors, and describe
   mutation outcomes without implying a retry is safe.
3. **Cutover scope:** keep the release atomic across the server, gateway, active
   configuration, citations, policy, generated catalogs, tests, and runbooks;
   exclude deployment and unrelated security migrations.

All three agreed on native structured output, strict runtime schemas,
credential-free capability discovery, broker-mode search removal, a real package
namespace, removal of `run_server.py`, mutation-aware retry control, and no
compatibility layer.

## Canonical tool family

The cutover uses the following names and no aliases:

| Current tool | Cutover tool | Canonical inputs | Direct success fields |
|---|---|---|---|
| `gsheet_list_tabs` | `gsheets_list_tabs` | `spreadsheet` | `spreadsheet`, `title`, `tabs` |
| `gsheet_read_range` | `gsheets_read_range` | `spreadsheet`, `range`, render options | `spreadsheet`, `range`, render options, `values` |
| `gsheet_update_range` | `gsheets_write_range` | `spreadsheet`, `range`, `values` | `spreadsheet`, `range`, `cell_count` |
| `gsheet_append_rows` | `gsheets_append_rows` | `spreadsheet`, `range`, `values` | `spreadsheet`, `range`, `cell_count` |
| `gsheet_create` | `gsheets_create_spreadsheet` | `title` | `spreadsheet`, `title`, `url` |
| `gsheet_copy_spreadsheet` | `gsheets_copy_spreadsheet` | `spreadsheet`, `title`, `tabs?` | `spreadsheet`, `title`, `url`, `tabs`, `warnings` |
| `gsheet_search` | `gsheets_search_spreadsheets` | `query`, `limit?` | `query`, `results`, `count` |
| `gsheet_clear_range` | `gsheets_clear_range` | `spreadsheet`, `range` | `spreadsheet`, `range` |
| `gsheet_touch_range` | `gsheets_recalculate_range` | `spreadsheet`, `range` | `spreadsheet`, `range`, `cell_count`, recovery/verification flags |

`spreadsheet` accepts a strict Google Sheets URL or ID and returns the normalized
ID. `range` is A1 notation. Common tools no longer accept titles in either mode;
local title discovery is an explicit search-then-select flow. Search results use
`{spreadsheet, title, url, modified_at}`, so the selected value flows unchanged
into every other tool.

The vocabulary review preferred these short domain terms over
`spreadsheet_ref`/`range_a1`. The latter are explicit but were not the observed
first guesses. The corpus repeatedly showed agents guessing `range`, while the
current outputs already teach it. The chosen names therefore minimize both
inference cost and output-to-input translation.

## Canonical wire shape

Success and error payloads are direct typed objects. They are not nested beneath
a JSON-string `result`, and useful success fields are not nested beneath another
result object.

```json
{
  "status": "ok",
  "operation": "gsheets_read_range",
  "spreadsheet": "resolved-id",
  "range": "Sheet1!A1:D20",
  "values": []
}
```

```json
{
  "status": "error",
  "operation": "gsheets_read_range",
  "error": {
    "code": "invalid_arguments",
    "message": "Arguments do not match the tool contract.",
    "outcome": {
      "state": "not_started",
      "phase": "validation",
      "mutation_may_have_occurred": false
    },
    "retry": {
      "safe": true,
      "automatic": false,
      "action": "correct_arguments",
      "retry_after_seconds": null
    },
    "validation": {
      "issues": [{"path": "dry_run", "code": "unknown_field", "message": "Remove this field."}],
      "allowed_fields": ["spreadsheet", "range"],
      "required_fields": ["spreadsheet", "range"],
      "example_arguments": {"spreadsheet": "<spreadsheet-id-or-url>", "range": "Sheet1!A1:D20"}
    },
    "recovery": null
  }
}
```

Both payloads are returned as `CallToolResult` with equivalent text content and
structured content. Errors set `isError=true`. Each advertised output schema is a
discriminated `oneOf` over that tool's success model and the common error model.
All object models reject extra fields. Raw Pydantic errors, caller values, Google
response bodies, broker payloads, formulas, and cell contents are never included
in validation errors or logs.

## Mode and safety decisions

- Broker mode exposes exactly eight tools; search is absent from `tools/list`.
- Local mode exposes the same eight plus `gsheets_search_spreadsheets`.
- Mode parsing is explicit: unset/`local` and `broker` are valid; other values
  fail closed before serving.
- Discovery and `gsheets-mcp capabilities --json` do not load credentials.
- Google 401 handling remains bounded to one exact HTTP-request refresh. The
  gateway may automatically replay only read-only tools whose structured error
  explicitly says the operation did not start and automatic retry is safe.
- Create, append, copy, write, clear, and recalculation are never blindly replayed.
- Copy retains destination ID, confirmed tab progress, active stage, and remaining
  work after destination creation.
- Recalculation rejects mixed literal/formula ranges before clearing, never logs
  the snapshot, performs at most one exact-snapshot compensation write, and
  verifies formulas before claiming recovery.

## Explicit exclusions

- No legacy tool-name alias, argument alias, duplicate output key, JSON-string
  result parser for Sheets, `src` compatibility package, or default-serve CLI shim.
- No unbacked `dry_run`, `idempotency_key`, or `if_not_exists` parameter.
- No whole-operation retry for a mutation.
- No pickle-to-JSON token-cache migration in this contract release. That is a
  separate local-auth security change and, when done, must not load pickle as a
  compatibility path.
- No deployment, gateway restart, OAuth re-consent, installed user-config edit,
  or push in this work unit.
- The gateway's generic text/JSON fallback remains for unrelated MCP servers; the
  Sheets-specific path and tests use only the new structured contract.
