# Task 04 — reject-fake-mutation-safety

**Result:** PASS
**Round-trips:** 1 (target: 1)
**First-try success:** yes

A real MCP stdio call supplied otherwise canonical write arguments plus `dry_run: true`. Closed-world validation rejected `dry_run` as an unknown field before dispatch. The response was `isError: true`, code `invalid_arguments`, with outcome `{state: not_started, phase: validation, mutation_may_have_occurred: false}`.

No write handler, credential loader, OAuth flow, or Google API was reached. The MCP transport/client process exited successfully; the structured tool error was the expected safety result.
