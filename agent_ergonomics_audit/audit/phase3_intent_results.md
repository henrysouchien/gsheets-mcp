# Phase 3 Intent Stress Results

The corpus contains 45 stable cases: 25 naive and 20 source-aware. Mutating calls
were executed only when malformed required arguments guaranteed validation before
the tool body. Any call that could reach Google was skipped.

## Outcomes

| Classification | Count | Interpretation |
|---|---:|---|
| `useful_hint` | 22 | Pydantic names a missing field, but emits framework text, echoes input fragments, and gives no canonical example. |
| `useless_error` | 1 | A plausible tool-name typo receives only `Unknown tool`. |
| `silent_fail` | 8 | Business/input-domain failures return `isError=false` and a JSON string nested under `structuredContent.result`. |
| `skipped` | 14 | Valid-enough calls could read/write external state, or exercise known unsafe partial-failure paths. |

The current surface does not infer legacy/wrong intent, and the full-cutover policy
does not ask it to accept aliases. The target improvement is therefore strict
rejection with safe structured pedagogy: unknown keys must be named, canonical keys
and a paste-ready example must be returned, and no input values may be echoed.

The highest-risk corpus cases are unknown safety-shaped fields. `dry_run`,
`idempotency_key`, and `if_not_exists` are currently ignored when canonical required
fields are also present. A caller can therefore believe a mutation is previewed or
deduplicated while the server performs it. `additionalProperties: false` must be
enforced at runtime, not merely advertised.
