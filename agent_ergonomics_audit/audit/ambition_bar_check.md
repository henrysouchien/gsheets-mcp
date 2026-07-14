# Ambition Bar Check — Pass 1 Design

## Original issue

Agents repeatedly failed Google Sheets calls because outputs taught
`spreadsheet_id`, `range`, and `title` while inputs expected `spreadsheet`,
`cell_range`, `source`, and `new_title`. The server also encoded useful results
inside a string and reported operational errors as successful MCP calls.

## Local fix considered and rejected

Adding prose reminders or accepting aliases would reduce some validation failures
but preserve the split vocabulary, silent unknown-field behavior, nested strings,
false-success errors, unsafe mutation retry, and unexecutable broker search. It
would also create a second contract the system would need to support.

## Systemic redesign selected

The proposed cutover changes the source of truth rather than decorating the old
surface:

1. One typed registry drives runtime validation, schemas, annotations, dispatch,
   and CLI capability output.
2. Tool and field names compose directly across the entire family.
3. Structured error outcomes drive gateway respawn/retry decisions.
4. Copy and recalculation expose and compensate partial state.
5. Package, CLI, policy, citations, generated catalogs, active configs, tests, and
   release procedure move together.

The design therefore changes more than five meaningful workflow elements across
MCP API, CLI, packaging, safety, gateway policy, citations, CI, and operations. It
meets the ambition bar as a plan. The audit pass remains incomplete until those
changes are applied, independently reviewed, verified, re-scored, and simulated;
`ambition_bar_met` stays false in the manifest until then.

## Post-cutover result

The design is now landed as eight substantive recommendation stacks across
three intentionally atomic repository commits:

- `gsheets-mcp` — `ea10997e540d28a5562cb346b9607ea0791376eb`
- `AI-excel-addin` — `e9499747ae281ffff71c0ad76c1f002ce7de49bd`
- `risk_module` — `a0588b7979acae5b78897f4089d7c497021b8a6b`

The implementation touches all 11 scoring dimensions, applies all eight
recommendations, and pins each with an audit regression. It deliberately uses
three cross-repository cutover commits instead of one commit per recommendation:
the no-compatibility release unit must be reviewable and reversible as one
commit per repository, without mixed consumer/server versions.

### Required surface types

- Machine-readable capabilities/JSON: added as
  `gsheets-mcp capabilities --json`, derived from the live registry.
- Structured robot output: the MCP boundary and capabilities command emit
  direct typed JSON with closed schemas.
- Error rewrite: validation, capability, auth, and mutation failures expose a
  typed outcome, retry policy, correction, and recovery object.
- Intent inference: fuzzy suggestions operate only inside the canonical
  `gsheets_*` namespace. Retired `gsheet_*` names intentionally receive no
  migration suggestion or behavioral recognition.
- Mega-command: deliberately not added. This MCP exposes narrow Sheets
  operations; combining external mutations into a mega-command would weaken
  the explicit outcome and no-replay guarantees.

The verbatim `That's it??` self-prompt was run, followed by another apply/review
round. That round found and closed three non-trivial gaps: required serialized
wire discriminators, a one-write ceiling after an uncertain clear, and exclusion
of retired names from fuzzy suggestions. Two independent final reviews passed.

### Gate decision

- Substantive landed recommendation stacks: **8**
- Atomic repository commits: **3**
- Dimensions touched: **11 / 11**
- Median matched-surface uplift: **+445 points**
- Surfaces below weighted 750: **0 / 23**
- Regressions: **0 weighted; 0 / 253 individual dimensions**
- Fresh-context simulation: **7 / 7 tasks passed; median 1 round-trip; 0 stuck**

The soft non-trivial-CLI target of ten separate changes was not reached, but
the skill's required self-prompt and second apply round were completed, all
ranked recommendations landed, and measured uplift is well beyond the gate.
`ambition_bar_met` is therefore true under the skill's documented post-self-
prompt completion path. No backward-compatible alias or legacy behavior was
added to inflate the count.
