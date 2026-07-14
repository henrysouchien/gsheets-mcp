# Agent-Ergonomics Full-Cutover Handoff

## Outcome

The post-cutover runtime is measurably more agent-ergonomic across all 23 stable
comparison surfaces. Pass 2 scores the canonical replacements while retaining
Pass 1 IDs solely for an apples-to-apples diff; it does not imply compatibility
aliases.

| Metric | Pass 1 | Pass 2 | Change |
|---|---:|---:|---:|
| Median weighted surface score | 450 | 881 | +431 |
| Median matched-surface uplift | — | — | +445 |
| Surfaces below weighted 750 | 23 | 0 | -23 |
| Weighted-score regressions | — | 0 | none |
| Individual dimension regressions | — | 0 / 253 | none |

No safety or regression-resistance dimension decreased.

## Delivered contract

- One canonical `gsheets_*` vocabulary; no runtime aliases for `gsheet_*`.
- One strict typed registry for discovery, validation, annotations, dispatch,
  output schemas, and CLI capabilities.
- Direct structured success and `isError=true` structured failures.
- Closed-world inputs that reject unknown safety-shaped arguments before
  credentials or mutation.
- Local nine-tool and broker eight-tool discovery; broker search calls teach the
  ID/URL recovery path without advertising or restoring search.
- Explicit mutation outcome states and retry policy; no blind mutation replay.
- Copy destination/progress recovery and compensating formula recalculation.
- Explicit package/CLI (`serve`, `capabilities`, `--version`) with the generic
  `src` and `run_server.py` launchers retired.
- Live registry lint, pre-commit checks, CI, and wire/package/safety tests.

Fresh-eyes review tightened three final invariants: every advertised output
schema requires serialized `status` and `operation`, uncertain-clear recovery
can issue at most one compensation write, and fuzzy suggestions operate only
inside the canonical `gsheets_*` namespace. Retired names receive a generic
unknown-tool response, not a compatibility hint.

## Landed provenance

- `gsheets-mcp`: `ea10997e540d28a5562cb346b9607ea0791376eb`
- `AI-excel-addin`: `e9499747ae281ffff71c0ad76c1f002ce7de49bd`
- `risk_module`: `a0588b7979acae5b78897f4089d7c497021b8a6b`

These are three distinct commits. Cross-repository files in
`applied_changes.jsonl` carry their owning repository SHA and are never
misrepresented as members of the gsheets runtime commit.

## Verification completed

- `venv/bin/python -m pytest -q`: **78 passed**.
- `venv/bin/python scripts/mcp_lint.py --json`: **0 issues**.
- Local capabilities repeated twice: identical SHA-256
  `e7f46efc0dfec1195776ad430e59c4d8bf184ed3cadd7a5ba253470fbf51e9d9`.
- Broker capabilities: exactly eight canonical tools; search absent.
- Stable 45-case intent replay: **39 useful hints, 0 useless errors, 0 silent
  failures, 6 intentionally skipped live calls**.
- Canonical composition C001–C005: **5/5 successful** through isolated MCP
  handlers, each explicitly marked non-live.
- Pass 2 scorecard: 23/23 scored with no dimension decrease.
- Fresh-context Phase 9 simulation: **7/7 tasks passed**, 10 transcript rows,
  median **1 round-trip**, **0 stuck**; the one optional replay corrected a
  simulator-only assertion while every MCP call had already succeeded.
- Two independent final reviews (runtime and cross-repository cutover):
  **PASS**.
- Eight recommendation-specific audit regression scripts: **all green**.
- Scorecard validator, strict artifact validator, and pass-consistency
  validator: **PASS** with zero failed records.

## Honest limits

- No live Google OAuth, token-broker, Drive, or Sheets request was made. The
  audit does not claim live success.
- Append, create, copy, and recalculation are inherently non-idempotent. The
  contract makes uncertain outcomes non-retryable; it does not invent a dedupe
  service.
- `dry_run`, `idempotency_key`, and `if_not_exists` are intentionally rejected,
  not emulated.
- Local token-cache pickle migration was explicitly outside this cutover.
- CLI subcommand typos still receive argparse's valid-choice list rather than
  an edit-distance correction; MCP tool/argument errors are more pedagogical.
- Deployment, gateway restart, OAuth exercise, and live smoke testing require a
  separately authorized maintenance window.

## Completion state

All eight ranked recommendations are marked applied against real landed SHAs;
the manifest, post-score evidence, fresh-agent transcripts, review record, and
regression mappings are complete. The remaining operational step is a separate,
authorized maintenance window that deploys all three commits as one cutover,
refreshes gateway schemas, and performs live OAuth/broker smoke tests.

No push, deployment, restart, or live credential action was part of this audit.
