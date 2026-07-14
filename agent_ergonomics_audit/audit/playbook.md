# Agent Ergonomics Playbook — Pass 1

This pass is a full cutover, not a compatibility migration. The recommendations
form one release unit; their priority order describes implementation risk, not
permission to ship a partial contract.

## 1. One composable vocabulary

Rename the tool family to the canonical `gsheets_*` names and use
`spreadsheet`, `range`, `title`, and `limit` consistently. Return normalized IDs
and affected ranges under the same keys so an agent can pass output to the next
tool without translation. No old name is accepted or mirrored.

## 2. One typed MCP boundary

Replace FastMCP's extra-ignore/generated-wrapper behavior with a low-level MCP
server backed by a typed `ToolSpec` registry. The registry owns names, schemas,
annotations, examples, handlers, output models, and CLI capability discovery.
Unknown fields fail before credentials or Google calls.

## 3. Honest errors and mutation states

Every failure is a real MCP error and carries a safe structured code, outcome,
retry decision, and recovery data. Validation names fields but never values.
Create/append/write/clear timeouts are allowed to be `uncertain`; the contract
must not pretend a retry is harmless.

## 4. Safe copy and recalculation

Copy retains its newly created destination and progress after every subsequent
stage. Recalculation is formula-only, keeps its snapshot in memory, never logs
cell data, performs at most one exact compensation write, and verifies the
formulas before reporting restoration.

## 5. Truthful mode capabilities

Broker mode advertises the executable eight-tool set. Local mode adds search.
Common tools accept strict Sheets URLs or IDs in both modes; local title use is
an explicit search/select workflow. Invalid mode configuration fails closed.

## 6. An explicit package and CLI

Install `gsheets_mcp`, delete the importable generic `src` package and
`run_server.py`, and make `gsheets-mcp serve` the only server launcher. Provide
help, version, and credential-free `capabilities --json` from the live registry.

## 7. Gateway-aware retry and policy

The gateway preserves structured error fields, respawns expired broker children
for future work, and automatically repeats only read-only calls explicitly marked
safe and not started. Sheets policy is closed-world; an unclassified runtime tool
is hidden and diagnosed. Approval policy remains authoritative over MCP hints.

## 8. Contract gates and atomic release

Pin schemas and results over a real stdio MCP connection; test CLI/package
installation, mutation phases, gateway retries, citation hashes, policies, and
mode-specific catalogs. Refresh generated context artifacts through their owner.
Update tracked launch configs now, but leave deploy/restart/push for a separately
authorized maintenance window. Roll back server and consumers together.
