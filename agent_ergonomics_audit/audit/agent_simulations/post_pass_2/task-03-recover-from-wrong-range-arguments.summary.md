# Task 03 — recover-from-wrong-range-arguments

**Result:** PASS
**Round-trips:** 1 (target: 1)
**First-try success:** yes

A real MCP stdio session called `gsheets_read_range` with the legacy keys `spreadsheet_id`, `sheet_name`, and `range_name`. The expected tool-level rejection arrived in one call with `isError: true`, code `invalid_arguments`, and outcome `{state: not_started, phase: validation, mutation_may_have_occurred: false}`.

The response listed all five issue paths, the allowed and required fields, and a canonical example object. None of the supplied canary values appeared in the response. The MCP transport/client process exited successfully; the structured tool error was the expected task result.
