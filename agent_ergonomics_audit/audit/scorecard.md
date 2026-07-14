# Agent Ergonomics Scorecard

Generated: 2026-07-14T16:03:39Z
Source: `agent_ergonomics_audit/audit/agent_surfaces.jsonl` (pass 2)

Historical surface IDs are retained only for Pass 1 ↔ Pass 2 comparison. Each
record's notes identify the full-cutover equivalent; the old runtime symbols are
not compatibility aliases.

Post-score fresh-eyes review additionally pinned three cross-cut invariants:
wire schemas require serialized `status`/`operation` discriminators, an
uncertain clear permits at most one compensation write, and retired
`gsheet_*` names are excluded from fuzzy suggestions. Scores were unchanged;
the stronger evidence is recorded in the cumulative surface rows and Phase 7
review artifact.


## Per-surface scores

| surface_id | weighted | intu | ergo | ease | parse | error | intent | safe | det | self | comp | regr |
|------------|----------|------|------|------|-------|-------|--------|------|-----|------|------|------|
| mcp_server__metadata | 895 | 900 | 850 | 900 | 950 | 850 | 750 | 1000 | 900 | 900 | 900 | 950 |
| launcher__installed | 863 | 950 | 850 | 900 | 950 | 650 | 500 | 1000 | 950 | 900 | 950 | 900 |
| launcher__source | 863 | 950 | 850 | 900 | 950 | 650 | 500 | 1000 | 950 | 900 | 950 | 900 |
| verb__mcp_lint | 863 | 850 | 850 | 850 | 950 | 750 | 500 | 1000 | 950 | 900 | 950 | 950 |
| mcp_tool__gsheet_list_tabs | 913 | 900 | 850 | 950 | 950 | 900 | 750 | 1000 | 900 | 950 | 950 | 950 |
| mcp_tool__gsheet_read_range | 913 | 900 | 850 | 950 | 950 | 900 | 750 | 1000 | 900 | 950 | 950 | 950 |
| mcp_tool__gsheet_update_range | 881 | 900 | 850 | 950 | 950 | 900 | 750 | 800 | 750 | 950 | 950 | 950 |
| mcp_tool__gsheet_append_rows | 854 | 900 | 850 | 950 | 950 | 900 | 750 | 750 | 500 | 950 | 950 | 950 |
| mcp_tool__gsheet_create | 859 | 900 | 900 | 950 | 950 | 900 | 750 | 750 | 500 | 950 | 950 | 950 |
| mcp_tool__gsheet_copy_spreadsheet | 886 | 900 | 850 | 950 | 950 | 950 | 750 | 900 | 650 | 950 | 950 | 950 |
| mcp_tool__gsheet_search | 895 | 850 | 850 | 950 | 950 | 850 | 700 | 1000 | 850 | 950 | 950 | 950 |
| mcp_tool__gsheet_clear_range | 881 | 900 | 850 | 950 | 950 | 900 | 750 | 800 | 750 | 950 | 950 | 950 |
| mcp_tool__gsheet_touch_range | 895 | 900 | 850 | 950 | 950 | 950 | 750 | 900 | 750 | 950 | 950 | 950 |
| env__GOOGLE_CREDENTIALS_FILE | 818 | 800 | 750 | 750 | 1000 | 750 | 500 | 1000 | 800 | 900 | 850 | 900 |
| env__GOOGLE_TOKEN_FILE | 804 | 800 | 750 | 750 | 1000 | 750 | 500 | 1000 | 700 | 900 | 800 | 900 |
| env__GSHEETS_TOKEN_MODE | 927 | 950 | 850 | 950 | 1000 | 900 | 750 | 1000 | 950 | 950 | 950 | 950 |
| env__GSHEETS_BROKER_URL | 863 | 850 | 800 | 900 | 1000 | 850 | 500 | 1000 | 800 | 950 | 900 | 950 |
| env__GSHEETS_BROKER_SESSION_TOKEN | 872 | 850 | 800 | 900 | 1000 | 900 | 500 | 1000 | 800 | 950 | 950 | 950 |
| env__GSHEETS_HEADLESS | 872 | 850 | 800 | 900 | 1000 | 850 | 500 | 1000 | 850 | 950 | 950 | 950 |
| auth__credential_modes | 877 | 850 | 800 | 900 | 950 | 900 | 650 | 950 | 850 | 950 | 900 | 950 |
| error__argument_validation | 954 | 1000 | 1000 | 1000 | 950 | 950 | 750 | 1000 | 1000 | 950 | 950 | 950 |
| error__read_auth_failure | 936 | 1000 | 1000 | 1000 | 950 | 900 | 650 | 1000 | 1000 | 900 | 950 | 950 |
| error__mutation_partial_failure | 936 | 1000 | 1000 | 1000 | 950 | 950 | 650 | 900 | 1000 | 950 | 950 | 950 |

## Distribution histogram

### Weighted score distribution (per surface)

```
   0- 99 │  (0)
 100-199 │  (0)
 200-299 │  (0)
 300-399 │  (0)
 400-499 │  (0)
 500-599 │  (0)
 600-699 │  (0)
 700-799 │  (0)
 800-899 │ █████████████████ (17)
 900-999 │ ██████ (6)
1000     │  (0)
```

## Below-Polish-Bar surfaces (weighted < 750)
