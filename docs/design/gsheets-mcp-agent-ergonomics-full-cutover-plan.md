# Google Sheets MCP Agent-Ergonomics Full-Cutover Plan

Status: Approved — independent overall, MCP-wire, and gateway reviews returned
literal `PASS` after R1 corrections

Date: 2026-07-14

Repositories in scope:

- `/Users/henrychien/Documents/Jupyter/gsheets-mcp`
- `/Users/henrychien/Documents/Jupyter/AI-excel-addin`
- `/Users/henrychien/Documents/Jupyter/risk_module` (tracked launch config only)

## 1. Outcome

Replace the current Google Sheets MCP contract with one strict, composable,
agent-native interface. The release removes the observed argument translation
trap, JSON-inside-a-string results, false-success errors, silently ignored safety
fields, broker capability overstatement, and unsafe retry/recalculation behavior.

This is an atomic cutover. It does not preserve or recognize old tool names,
argument names, result keys, import paths, or launch commands. Every tracked
first-party consumer moves in the same work unit. Deployment and restart are not
part of this implementation authorization.

## 2. Non-negotiable gates

1. No tool aliases, Pydantic aliases, duplicate output keys, legacy JSON-string
   result parser for Sheets, `src` compatibility package, `run_server.py`, or bare
   CLI default that silently starts stdio.
2. Unknown fields fail before credentials or Google calls. This includes plausible
   safety fields such as `dry_run`, `idempotency_key`, and `if_not_exists`.
3. Validation output may name field paths and fixed examples but must never echo
   submitted values.
4. Every operational failure sets MCP `isError=true` and carries typed structured
   code, outcome, retry, and recovery data.
5. Only a read-only tool explicitly marked `retry.safe=true`,
   `retry.automatic=true`, and `outcome.state=not_started` may be automatically
   replayed by the gateway.
6. Broker discovery remains credential-free and exposes exactly eight executable
   tools. Local discovery exposes those eight plus search.
7. Approval/redaction/citation authority stays in the gateway; MCP annotations are
   hints, not authorization.
8. Existing broker safety invariants remain: per-user process identity, in-memory
   broker tokens, 60-second proactive refresh, one exact-request Google 401 refresh,
   strict URL/ID verification, and no Drive search in broker mode.
9. No edits to any `-dist` repository. Preserve all unrelated dirty files.
10. Do not push, deploy, restart the gateway, mutate installed MCP config, or run
    OAuth consent. Prepare a coordinated maintenance-window handoff only.
11. Pin the implementation to `mcp>=1.26.0,<2` and `pydantic>=2.12,<3`; older
    SDK/Pydantic behavior does not satisfy this contract.

## 3. Verified baseline and root causes

- Target: `b72d8a3420fb417d7842b83ce87dd4961d74d10f` on current `main`.
- `venv/bin/python -m pytest -q`: 41 passed.
- `venv/bin/python scripts/mcp_lint.py`: exit 0.
- Bare `pytest` can import an adjacent editable project's generic `src` package;
  the namespace is nondeterministic outside the project venv command.
- MCP 1.26 FastMCP generates argument models with extra fields ignored and wraps
  Pydantic validation using raw string formatting. Its current output schema for
  every tool is only `{result: string}`.
- The audit found 21 wrong-argument failures across 18 sessions. All recovered,
  wasting about 16 seconds each. Current results teach keys different from the
  next tool's inputs.
- `gsheet_search` is advertised in broker mode even though it cannot execute.
- Copy can leave an unidentified orphan/partial destination; touch can clear data
  and log the entire snapshot; gateway broker-session recovery can replay a whole
  operation without mutation-phase proof.

## 4. Canonical contract

### 4.1 Tool names and inputs

| Tool | Required inputs | Optional inputs | Broker | Local |
|---|---|---|---:|---:|
| `gsheets_list_tabs` | `spreadsheet` | — | yes | yes |
| `gsheets_read_range` | `spreadsheet`, `range` | `value_render_option`, `date_time_render_option` | yes | yes |
| `gsheets_write_range` | `spreadsheet`, `range`, `values` | — | yes | yes |
| `gsheets_append_rows` | `spreadsheet`, `range`, `values` | — | yes | yes |
| `gsheets_create_spreadsheet` | `title` | — | yes | yes |
| `gsheets_copy_spreadsheet` | `spreadsheet`, `title` | `tabs` | yes | yes |
| `gsheets_search_spreadsheets` | `query` | `limit` (default 10, range 1..100) | no | yes |
| `gsheets_clear_range` | `spreadsheet`, `range` | — | yes | yes |
| `gsheets_recalculate_range` | `spreadsheet`, `range` | — | yes | yes |

`spreadsheet` means a strict `https://docs.google.com/spreadsheets/...` URL or a
spreadsheet ID. The returned value is always the normalized ID. Common tools do
not resolve titles in either mode. Local title lookup is
`gsheets_search_spreadsheets` followed by selection of
`results[n].spreadsheet`. `range` is A1 notation. `values` is a non-empty
two-dimensional JSON array; accepted cell scalars are string, finite number,
boolean, and null. JSON-stringified arrays and unknown keys are rejected.

Render options are schema enums:

- `value_render_option`: `FORMATTED_VALUE`, `UNFORMATTED_VALUE`, `FORMULA`
- `date_time_render_option`: `FORMATTED_STRING`, `SERIAL_NUMBER`

### 4.2 Direct structured successes

Every success includes `status: "ok"` and its literal `operation`. Useful fields
are top-level, snake_case, and typed:

- list: `spreadsheet`, `title`,
  `tabs[{sheet_id,title,index,row_count,column_count}]`
- read: `spreadsheet`, `range`, both render options, `values`
- write/append: `spreadsheet`, actual affected `range`, `cell_count`
- create: `spreadsheet`, `title`, `url`
- copy: destination `spreadsheet`, `title`, `url`, copied `tabs`, and
  `warnings[{code,message}]`
- search: `query`, `count`,
  `results[{spreadsheet,title,url,modified_at}]`
- clear: `spreadsheet`, actual cleared `range`
- recalculate: `spreadsheet`, `range`, `cell_count`, `recovery_performed`, and
  `formulas_verified`

There is no nested JSON string, camelCase mirror, old output key, or generic
`result` wrapper.

### 4.3 Structured errors

Every failure is a `CallToolResult` with equivalent JSON text and structured
content, `isError=true`, and this shape:

```json
{
  "status": "error",
  "operation": "gsheets_copy_spreadsheet",
  "error": {
    "code": "copy_partial",
    "message": "The destination was created but the copy did not finish.",
    "outcome": {
      "state": "partial",
      "phase": "copy_tabs",
      "mutation_may_have_occurred": true
    },
    "retry": {
      "safe": false,
      "automatic": false,
      "action": "inspect_destination",
      "retry_after_seconds": null
    },
    "validation": null,
    "recovery": {
      "kind": "copy_progress",
      "destination_spreadsheet": "...",
      "destination_url": "...",
      "confirmed_tabs": ["Data"],
      "active_tab": "Summary",
      "active_tab_state": "uncertain",
      "remaining_tabs": ["Assumptions"],
      "finalization_state": "not_started"
    }
  }
}
```

Allowed outcome states are `not_started`, `unchanged`, `uncertain`, `partial`,
and `restored`. Validation issues are built only from Pydantic error locations and
fixed type mappings using `errors(include_input=False, include_url=False,
include_context=False)`. Unknown field names are printable-character sanitized
and length limited. Arbitrary exception strings, HttpError bodies, broker payloads,
request bodies, formulas, and cell values never enter responses or logs.

An unexpected internal error may receive a generated incident ID that is logged
without arguments or payload data; other errors stay deterministic.

### 4.4 Annotations

| Class | readOnlyHint | destructiveHint | idempotentHint | openWorldHint |
|---|---:|---:|---:|---:|
| list/read/search | true | false | true | true |
| write | false | true | false | true |
| append/create/copy | false | false | false | true |
| clear | false | true | true | true |
| recalculate | false | true | false | true |

Write is marked non-idempotent conservatively because a transport failure can
leave its outcome uncertain. Gateway policy remains the approval source of truth.

## 5. Server architecture

### 5.1 Package and CLI

Move the implementation to a real src-layout package:

```text
src/gsheets_mcp/
  __init__.py
  cli.py
  contracts.py
  server.py
  sheets_client.py
  tools.py
```

Delete `src/__init__.py`, `src/server.py`, `src/sheets_client.py`, and
`run_server.py`; do not create forwarding modules. Update `pyproject.toml` to
version `1.0.0`, build `src/gsheets_mcp`, declare `mcp>=1.26.0,<2` and
`pydantic>=2.12,<3`, and expose `gsheets-mcp = gsheets_mcp.cli:main`.

The wire contract version is exactly `1.0`. Define package and contract versions
once in `gsheets_mcp.__init__`; initialize reports `serverInfo.version=1.0.0` and
an experimental `gsheets` capability carrying `contract_version=1.0`. Tool
metadata and capabilities output carry the same contract version. Tests pin all
three views so package and wire versions cannot drift.

CLI contract:

- `gsheets-mcp` and `gsheets-mcp --help`: concise help, exit 0, no stdio server.
- `gsheets-mcp --version`: version only on stdout.
- `gsheets-mcp serve`: run stdio MCP; stdout remains protocol-only.
- `gsheets-mcp capabilities`: short human summary from the live registry.
- `gsheets-mcp capabilities --json`: stable JSON with package/contract version,
  mode, tools, schemas, annotations, and non-secret environment requirements.

No command loads credentials except actual tool execution. Invalid
`GSHEETS_TOKEN_MODE` fails closed for `serve` and capabilities. Keep the existing
pickle token cache unchanged in this release; do not add a legacy migration or
touch the ignored credential/token files.

### 5.2 Typed low-level boundary

Use `mcp.server.lowlevel.Server`, not FastMCP decorators. `contracts.py` defines:

- strict Pydantic input models (`ConfigDict(extra="forbid", strict=True)`),
- strict nested success/error/progress/warning models,
- the common validation/outcome/retry/recovery models,
- per-tool discriminated output adapters.

`server.py` defines an immutable `ToolSpec` registry containing name, title,
description, input model, per-tool `result_adapter`, handler, fixed example
arguments, annotations, and optional metadata. Each `result_adapter` is the
status-discriminated union of that tool's success object and the common error
object. Its JSON schema is the exact advertised `Tool.outputSchema`, and the same
adapter validates both branches before any `CallToolResult` is constructed. The
same registry drives `tools/list`, safe validation, dispatch, output validation,
and CLI capabilities. Register
`call_tool(validate_input=False)` and perform custom validation so raw SDK errors
cannot escape. Validate every success/error envelope before constructing
`CallToolResult`, because direct results bypass low-level SDK output validation.

Tool descriptions include the workflow, destructive/non-idempotent warning,
sibling tool, common mistake, and canonical example. Schemas include property
descriptions, enums/bounds, required fields, and recursive
`additionalProperties:false`.

### 5.3 Sheets helpers and reference semantics

Refactor `sheets_client.py` without changing the broker token boundary:

- Parse an ID or strict URL identically before local/broker execution.
- Verify the normalized ID through Sheets metadata and return `(id, title)`.
- Remove implicit exact-title resolution from common operations.
- Keep Drive authentication only for local search.
- Normalize every Google camelCase response at the helper boundary.
- Classify mutation exceptions using the active phase and whether a mutating
  request was dispatched. Definite 4xx rejection is `unchanged`; transport/5xx
  after dispatch is `uncertain`; pre-dispatch failures are `not_started`.
- Preserve broker typed auth codes and the one exact-request refresh limit.

### 5.4 Copy state machine

Before creation, validate the title/tabs and read source metadata. After the
create response, immediately retain destination ID, URL, confirmed tabs, active
tab, remaining tabs, and finalization stage in a progress object. Any later error
raises a typed partial failure containing that object. Never rerun or auto-delete
the destination: Sheets-only scope cannot prove cleanup and a second run would
create another file. Warnings use stable codes (`named_ranges_not_copied`,
`locale_not_preserved`, `time_zone_not_preserved`) plus messages.

### 5.5 Recalculation state machine

1. Read the requested range with `FORMULA`.
2. Count cells before any clear. A blank-only range is a successful no-op. Any
   non-empty literal — whether the range is all literals or mixed with formulas —
   is a typed `unchanged` precondition error. Only formulas plus blanks may proceed.
3. Keep the exact formula matrix in memory only.
4. Clear, rewrite the exact matrix with `USER_ENTERED`, then read `FORMULA` and
   verify equality.
5. If clear/rewrite/verification fails after clear or write may have occurred,
   perform at most one exact snapshot write and one verification read. Never run
   compensation after a proven `not_started` or pre-clear `unchanged` failure.
6. If formulas match, return success with `recovery_performed=true` when the
   compensation path ran. If they do not, return a partial/uncertain error with
   identifiers and counts only.

Remove the current snapshot warning log. Do not claim that custom-function values
are fresh; claim only that recalculation was requested and formulas were restored
and verified.

## 6. Gateway and consumer cutover

### 6.1 Result/error handling

In `packages/agent-gateway/agent_gateway/mcp_client.py`:

- Add a structured Sheets error extractor for
  `structuredContent.error.{code,outcome,retry,recovery}`.
- Replace Sheets broker recovery based only on a top-level/string `error_code`.
- Always replace/drain an expired per-user child for future calls.
- Replay the call once only if gateway policy classifies it read-only and the
  error says safe + automatic + not_started.
- Return gateway errors with the MCP error code and structured outcome/retry/
  recovery details intact. Do not reduce them to message text.
- Never use the generic stdio reconnect replay for a Sheets mutation. Gate that
  path using the policy tool class.
- If a timeout, EOF, or child death occurs after a Sheets mutation was dispatched,
  synthesize a sanitized gateway error with `outcome.state="uncertain"`,
  `mutation_may_have_occurred=true`, and retry `safe=false`/`automatic=false`.
  Reconnect or replace the child for future calls but do not replay the operation.
  A spawn/config failure before dispatch is `not_started`. Read-side transport
  errors are also typed and sanitized, and may replay only under the explicit
  read-only rule above.
- Keep the generic text/JSON result fallback for unrelated MCPs; direct Sheets
  tests use only structured content.

Add focused helpers rather than coupling generic MCP parsing to Sheets field
names. Ensure logs/diagnostics include codes/stages/counts only, not cell values.

### 6.2 Policy and closed world

Update `api/agent/shared/server_policy_catalog.py` to the canonical eight broker
tools (two reads, six external writes) and new redaction keys. Add a
`strict_runtime_tool_set`/equivalent flag to `ServerPolicy`; when set, the policy
owner invariant hides and diagnoses any runtime tool not in that server's known
read/write/predicate sets. Set it only for `gsheets-mcp` in this work. Search is
valid locally but must never leak through the broker gateway definition.

Retain approval-before-dispatch and citation-before-sanitize ordering. MCP
annotations do not bypass approval. Update all drift and approval tests.

### 6.3 Guidance, citations, and generated catalogs

- Replace the temporary old-argument warning blocks in
  `mcp_client_catalog.py` with concise canonical workflow guidance; the input
  schema itself carries property descriptions.
- Update `api/agent/shared/citation_vendor_sources.py` and
  `api/fms/core/sources.py` from `gsheet_read_range` to
  `gsheets_read_range`.
- Change `tests/fms/test_vendor_handles.py` to use the direct dict success payload,
  `range`, and the new endpoint. Assert structured errors do not mint handles and
  direct result hashes change only with meaningful value changes.
- Update active profile/approval tests that name Sheets tools.
- Extend the context-manifest owner with a tested selective refresh/merge command:
  `--refresh-mcp-fixture --refresh-mcp-server gsheets-mcp --mcp-config PATH`.
  It must instantiate a manager limited to the named server, capture its live
  definitions without credentials, replace only that server in the existing full
  fixture, preserve every other server, fail if the requested server is absent,
  and record selective-live provenance. This avoids the currently degraded
  `~/.claude.json` capture. Do not use `--allow-degraded` and do not hand-edit the
  fixture. Use `../risk_module/config/research_gateway_mcp.local.json` after its
  launcher is updated.
- Run the selective refresh, regenerate `docs/agent-context/` with `--write`, and
  verify with `--check`.
- Mandatory active tool-surface refresh:
  `PYTHONPATH=api python3 -m agent.tool_surface_allocation --write`, followed by
  `PYTHONPATH=api python3 -m agent.tool_surface_allocation --check`.
- Update the current `docs/audits/_inventory/server_inventory.jsonl` gsheets row
  and `inventory_summary.md` to the console `serve` entry point, eight broker
  tools/nine local tools, the new package source, and current verification date.
  These are active inventory inputs, not immutable audit history.
- Do not rewrite completed specs, archives, session logs, historical synthesis
  outputs, or `.bak` memory files.

### 6.4 Active workflow instructions

Update the first-party comp-sheet workflow atomically:

- `api/memory/workspace/notes/skills/comp-sheet.md`: frontmatter `mcp_tools` and
  every workflow/example/warning reference.
- `api/memory/workspace/knowledge-sources.md`: copy, create, write, and warning
  instructions.
- `docs/design/comps-build-artifact-lane-spec.md`: active matrices, declared tool
  list, build/refresh flows, formula reads, writes, copy, and recalculation prose.

Add a regression test that scans active skill/prompt sources for the canonical
names and rejects the removed Sheets names. Historical completed/archived records
are outside that assertion.

## 7. Tracked config and operational documentation

Update:

- `AI-excel-addin/deploy/mcp.production.json` to command
  `/var/www/gsheets-mcp/venv/bin/gsheets-mcp` with `args: ["serve"]`.
- `risk_module/config/research_gateway_mcp.local.json` to the local installed
  console script with `args: ["serve"]`.
- `AI-excel-addin/scripts/deploy_gsheets_mcp.sh` import smoke to
  `import gsheets_mcp`, plus CLI version/capability smoke. Do not execute it.
- `gsheets-mcp/README.md` for the package/CLI, canonical tools, direct structured
  results/errors, mode matrix, local OAuth, and no-title common addressing.
- `AI-excel-addin/scripts/check_prod_deploy_prereqs.py` to require
  `src/gsheets_mcp/__init__.py` (and its tests).
- `AI-excel-addin/tests/scripts/test_check_prod_deploy_prereqs.py` and
  `tests/scripts/test_deploy_google_mcp_helpers.py` for the new package/import/
  launcher smoke contract.
- Active deployment/runbook references to the canonical launch and eight-tool
  broker capability set.
- Append a dated supersession note to
  `docs/design/gsheets-per-user-oauth-spec.md`; preserve historical acceptance
  claims rather than rewriting them.

Do not modify installed systemd units, `/etc` env files, live gateway config,
credentials, or processes. The authoritative gateway state paths remain
`/mnt/hank-data/agent_gateway/` as required by repository instructions.

## 8. Quality gates

### 8.1 gsheets-mcp tests

Add or rewrite tests for:

- all input/output JSON schemas and annotations;
- recursive unknown-field rejection, strict types, enums, bounds, and safe examples;
- validation messages with no submitted values;
- real stdio initialize, mode-specific `tools/list`, success, and `isError=true`
  calls;
- credential-free list/capabilities;
- URL/ID parity and common-tool title rejection;
- local search normalization and broker search absence;
- every broker auth code and bounded 401 refresh;
- each mutation outcome class;
- copy progress before/after create/copy/finalize stages;
- formula-only recalculation, no-op, compensation success, compensation failure,
  verification mismatch, and no snapshot leakage;
- CLI help/version/capabilities/serve discipline;
- installed package import from a neutral working directory;
- active runtime/schema absence of old tool/field/import/launcher names.

Replace or extend `scripts/mcp_lint.py` so it checks the live registry/schema rather
than declaring success because wrappers contain try/except. Keep stable `--json`
diagnostics. Update `.pre-commit-config.yaml` and `.github/workflows/mcp-audit.yml`
so changes to package helpers, contracts, tests, `pyproject.toml`, configs, and
workflow files run both pytest and lint.

Required commands:

```bash
venv/bin/pip install -e .
venv/bin/python -m pytest -q
venv/bin/python scripts/mcp_lint.py
venv/bin/gsheets-mcp --help
venv/bin/gsheets-mcp --version
GSHEETS_TOKEN_MODE=broker venv/bin/gsheets-mcp capabilities --json
(cd /tmp && /Users/henrychien/Documents/Jupyter/gsheets-mcp/venv/bin/python -c 'import gsheets_mcp')
```

### 8.2 AI-excel-addin tests

Run the focused suites covering:

- MCP result/error normalization and per-user respawn/retry;
- gateway-synthesized post-dispatch mutation transport uncertainty and
  pre-dispatch not-started errors, for both per-user and generic stdio paths;
- policy owner closed-world behavior;
- catalog enrichment and property descriptions;
- server policy drift and approval classification;
- research-producer/profile tool lists;
- direct-dict Sheets vendor handles and failed-result non-minting;
- deploy config/helper assertions;
- context manifest generation/check.
- active comp-sheet instruction legacy-name guard;
- mandatory tool-surface allocation write/check and current MCP inventory rows.

Then run the repo's practical broader gateway/FMS test group if focused tests pass.
No `schema/` file is planned. If implementation touches `schema/` or changes a
schema-consumed model path, run the mandatory guardrail before finishing:

```bash
PYTHONHASHSEED=0 python3 schema/smoke_accuracy_guardrail.py
```

Never update its baseline without explicit user authorization.

### 8.3 Config and audit validation

- Parse both JSON configs with `jq`/Python.
- Run the ergonomics workspace strict validator.
- Re-run the 45-case intent corpus against the new server, then add canonical
  successful composition cases and fresh-agent simulations.
- Re-score all 23 original surface classes using cutover equivalents, render pass
  2 scorecard/heatmap/uplift/regression artifacts, and require no safety or
  regression-resistance decrease.

## 9. Implementation order

1. Recheck branch/status in all three repos and record unrelated dirty paths.
2. Implement the package namespace, contracts, registry, CLI, and wire boundary.
3. Refactor reference/search helpers and implement typed mutation state machines.
4. Rewrite server tests/lint/CI/docs and pass the full gsheets-mcp suite.
5. Update gateway structured errors, retry gating, policy closed world, catalog,
   citations, active comp-sheet instructions, and focused tests.
6. Update production/local tracked configs, deploy/prerequisite helpers, README,
   supersession note, current inventory, and config tests.
7. Selectively refresh the Sheets context fixture, then regenerate all
   generator-owned context/tool-surface artifacts and run their checks.
8. Run cross-repo verification, intent replay, fresh-agent simulations, and
   ergonomics re-score.
9. Obtain independent implementation review. Fix every blocking finding and rerun
   affected tests until literal `PASS`.
10. Commit only scoped paths on each repo's current `main`; do not push. Record all
    commit SHAs and a coordinated deployment/rollback checklist.

## 10. Atomic deployment handoff (instructions only)

The later authorized maintenance window must:

1. Confirm all three tracked cutover commits are present and gateway jobs are
   drained/paused.
2. Install the new gsheets package and deploy the gateway/config source together.
3. Restart/recreate gateway and per-user child processes so cached schemas cannot
   retain the old contract.
4. Verify broker `tools/list` is exactly eight canonical tools without fetching a
   token.
5. Smoke one authenticated read, one approval-gated reversible write to a test
   sheet, a structured validation error, and a broker-session read retry.
6. Confirm create/copy/append/recalculation are never automatically replayed.
7. Resume traffic only after catalog/policy/citation checks pass.

Rollback is also atomic: restore the prior gsheets package, gateway source, and
both configs together, restart/recreate children, and verify the prior catalog.
Do not mix old and new versions. External spreadsheets created or partially
changed during smoke/failed operations are not undone by code rollback and must be
inspected using the returned progress identifiers.
