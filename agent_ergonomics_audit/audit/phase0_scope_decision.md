# Pass 1 Scope Decision

**Mode.** `full`

**Target.** `/Users/henrychien/Documents/Jupyter/gsheets-mcp`

**Workspace.** `/Users/henrychien/Documents/Jupyter/gsheets-mcp/agent_ergonomics_audit/` — in-tree

**Target branch.** Existing `main`; no branch creation or switching

**Triangulation appetite.** `peer-agent`

**CASS appetite.** `quick`, using already-mined local session evidence

**Date.** `2026-07-14`

**Target kind.** MCP server, with a thin installed console entry point

**Canonical entry point.** `gsheets-mcp = src.server:main`

**Transport.** stdio via FastMCP

## Canonical user jobs

The audit covers discovery and execution of the complete Sheets workflow: list tabs,
read a range, update a range, append rows, clear a range, verify/touch a range,
create a spreadsheet, and copy a spreadsheet. Search is assessed as a
capability-gated surface because broker mode cannot execute it.

## Cutover policy

This is a full contract cutover, not a compatibility migration. The resulting MCP
surface will advertise and accept one canonical argument vocabulary, return one
canonical structured result vocabulary, and use proper MCP error semantics.
Legacy aliases, duplicate output keys, JSON-string result envelopes, and error-as-
success behavior are not retained merely to preserve compatibility. Every in-scope
consumer and contract test must move atomically with the server contract.

## Must-not-touch

- Do not change campaign-owned Codex skill files; this audit consumes the skills
  but does not modify them.
- Do not push, deploy, restart, or rematerialize a running gateway/server without
  explicit user approval.
- Do not edit generated `-dist` deployment repositories; change source owners only.

## Deprecation policies

- No deprecation or compatibility period. This pass removes the legacy argument,
  output, and error contracts and cuts every first-party consumer over atomically.
- Do not preserve an alias, duplicate response field, or parsing shim solely to
  keep the old contract working.

## Out-of-scope feature work

- New spreadsheet business capabilities unrelated to agent ergonomics (formatting,
  charting, Drive-wide broker search, or general macro expansion) are out of scope.
- Deployment and live process restart are separate, approval-gated operations after
  the source cutover is committed and verified.

## Non-negotiable safety constraints

- Preserve per-user OAuth/broker isolation, token scoping, and broker-mode title
  restrictions.
- Preserve approval, redaction, citation, and idempotency policy ordering across
  the gateway boundary.
- Do not implement failure-driven retries or automatic argument remapping for
  mutating operations. Create, copy, append, and touch must not be blindly retried.
- Do not log spreadsheet cell values on failed verification or recovery paths.
- Do not expose tools that are guaranteed to fail in the active capability mode.
- Do not push, deploy, restart, or rematerialize a running gateway/server without
  explicit user approval.

## Execution constraints and fallbacks

- Use the repository's existing Python environment and toolchain; do not install
  dependencies without approval.
- The host does not provide `flock` (macOS provides `lockf`). Audit and source
  writes are serialized, with coordination used to prevent overlapping edits.
- Optional audit skills are unavailable except Agent Mail, so the audit uses the
  skill package's documented inline fallbacks.
- The optional CASS helper is unavailable. Phase 3 uses the already-mined local
  session corpus: 21 wrong-argument validation failures across 18 sessions, all
  corrected in-session, with roughly 16 seconds of avoidable retry latency on
  average.
- No campaign-owned skill files are changed as part of this work.

## Discovery correction

The generic discovery script reported `requires` and `build-backend` as binaries
because it scanned PEP 621/build-system keys. They are false positives. The only
installed command declared by the project is `gsheets-mcp`; its default behavior is
to run the MCP stdio server rather than a conventional command/subcommand CLI.
