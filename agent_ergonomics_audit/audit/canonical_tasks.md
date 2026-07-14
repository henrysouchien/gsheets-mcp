# Canonical Tasks for gsheets-mcp — Post-Cutover

These tasks are safe to run without Google credentials. A task that needs a
successful Sheets payload must use an isolated in-process handler and label the
result as a test double; no task may claim a live Google success.

## Task 01: discover-the-contract

**Statement.** Starting with no repo knowledge, discover how to run the server
and obtain its machine-readable local tool contract.

**Tags.** read-only / credential-free / CLI / first-try

**Expected outcome.** Bare `gsheets-mcp` prints help and exits 0; capabilities
JSON reports package `1.0.0`, contract `1.0`, mode `local`, and nine tools.

**Post-pass target.** 2 invocations, no external docs.

## Task 02: compare-local-and-broker-discovery

**Statement.** Determine which tool is intentionally unavailable in broker mode
without loading credentials.

**Tags.** read-only / credential-free / mode-parity

**Expected outcome.** Local advertises nine tools; broker advertises eight and
omits only `gsheets_search_spreadsheets`.

**Post-pass target.** 2 invocations.

## Task 03: recover-from-wrong-range-arguments

**Statement.** Call the canonical read tool with legacy/wrong keys and determine
the corrected argument object without exposing supplied values.

**Tags.** read-only / MCP / error-pedagogy / credential-free

**Expected outcome.** `isError=true`, code `invalid_arguments`, outcome
`not_started`, issue paths, allowed/required fields, and canonical example.

**Post-pass target.** 1 call to learn the correction.

## Task 04: reject-fake-mutation-safety

**Statement.** Verify that adding `dry_run: true` to a write cannot silently
perform the write.

**Tags.** mutating-intent / MCP / safety / credential-free

**Expected outcome.** Closed-world validation rejects `dry_run` before dispatch;
the structured outcome says no mutation may have occurred.

**Post-pass target.** 1 call.

## Task 05: compose-results-without-key-translation

**Statement.** Prove that normalized results can feed sibling tools: search →
list, list → read, read → write, create → write, and copy → read.

**Tags.** MCP / composition / isolated-test-double

**Expected outcome.** Cases C001–C005 in the intent corpus return direct
`status: "ok"` objects through the real MCP input/result boundary. Every row is
marked `test_double: true` and `live_google_call: false`.

**Post-pass target.** 1 inspection plus optional replay; no live Google call.

## Task 06: verify-recalculation-recovery

**Statement.** Check the formula-only precondition, exact restore, single
compensation, verification, and partial recovery behavior without touching a
real spreadsheet.

**Tags.** mutating-intent / safety / unit-test / credential-free

**Expected outcome.** The targeted injected tests pass, and no values appear in
logs or error payloads.

**Post-pass target.** 1 targeted test command.

## Task 07: verify-package-boundary

**Statement.** Confirm that the installed package imports from a neutral working
directory and retired generic `src`/`run_server.py` launch paths are absent.

**Tags.** package / determinism / credential-free

**Expected outcome.** The targeted CLI/package tests pass.

**Post-pass target.** 1 targeted test command.
