# Phase 3 Intent Stress Results — Pass 2

The original 45 stable cases were replayed against the canonical `gsheets_*`
tool family and strict input models. The replay invoked only schema failures,
reference failures, mode failures, and unknown-tool paths that were guaranteed
not to load credentials or touch Google. Six valid calls were skipped. Five
additional output-to-input compositions ran through the real MCP validation and
result boundary with isolated handlers; they are explicitly marked
`test_double: true` and are not live Google successes.

## Stable-corpus outcomes

| Classification | Pass 1 | Pass 2 | Interpretation |
|---|---:|---:|---|
| `useful_hint` | 22 | 39 | Typed errors now include safe issue paths, allowed and required fields, canonical examples, or mode recovery. |
| `useless_error` | 1 | 0 | The old unknown-tool dead end now teaches the canonical tool or capability path. |
| `silent_fail` | 8 | 0 | No replayed business/input-domain failure returned a successful MCP result. |
| `skipped` | 14 | 6 | These calls are now valid enough to reach Google, so the credential-free audit does not execute them. |

Transitions across the same 45 IDs:

- 8 `silent_fail` → `useful_hint`
- 1 `useless_error` → `useful_hint`
- 11 formerly skipped unsafe/late-validation cases → credential-free `useful_hint`
- 19 `useful_hint` → stronger structured `useful_hint`
- 3 baseline skips remained skips
- 3 former validation-hint cases became valid canonical compositions and were
  therefore skipped live; C001–C005 exercise those composition properties with
  isolated handlers.

## Canonical composition cases

| Case | Composition | Result | Execution boundary |
|---|---|---|---|
| C001 | search → list tabs | direct `results[n].spreadsheet` reuse | isolated MCP handler |
| C002 | list tabs → read range | direct `spreadsheet` reuse plus returned tab title | isolated MCP handler |
| C003 | read range → write range | direct `spreadsheet`, `range`, and `values` reuse | isolated MCP handler |
| C004 | create spreadsheet → write range | direct destination `spreadsheet` reuse | isolated MCP handler |
| C005 | copy spreadsheet → read range | copy result unambiguously exposes destination `spreadsheet` | isolated MCP handler |

All five returned direct `status: "ok"` objects that passed the live result
adapter. No JSON-inside-string compatibility layer or legacy alias was used.

## Credential boundary

No OAuth, token broker, Drive, or Sheets API request was made by this replay.
S019 (valid recalculation) remains skipped because it mutates external data;
its formula-only, compensation, and recovery paths are covered by injected unit
tests. S020 (valid strict Sheets URL) remains skipped because a real read would
contact Google.
