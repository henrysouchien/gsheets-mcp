# Agent Ergonomics Scorecard

Generated: 2026-07-14T15:22:13Z
Source: `audit/agent_surfaces.jsonl` (pass 1)

## Per-surface scores

| surface_id | weighted | intu | ergo | ease | parse | error | intent | safe | det | self | comp | regr |
|------------|----------|------|------|------|-------|-------|--------|------|-----|------|------|------|
| mcp_server__metadata | 450 | 500 | 500 | 750 | 300 | 300 | 0 | 1000 | 700 | 500 | 400 | 0 |
| launcher__installed | 286 | 250 | 500 | 0 | 400 | 0 | 0 | 1000 | 600 | 0 | 400 | 0 |
| launcher__source | 250 | 0 | 250 | 0 | 500 | 0 | 0 | 1000 | 500 | 0 | 500 | 0 |
| verb__mcp_lint | 477 | 500 | 500 | 500 | 500 | 250 | 250 | 1000 | 750 | 250 | 750 | 0 |
| mcp_tool__gsheet_list_tabs | 504 | 700 | 700 | 750 | 350 | 350 | 0 | 1000 | 650 | 700 | 350 | 0 |
| mcp_tool__gsheet_read_range | 481 | 600 | 650 | 700 | 350 | 400 | 0 | 1000 | 650 | 650 | 300 | 0 |
| mcp_tool__gsheet_update_range | 368 | 650 | 500 | 700 | 300 | 450 | 0 | 0 | 500 | 650 | 300 | 0 |
| mcp_tool__gsheet_append_rows | 336 | 700 | 500 | 700 | 300 | 450 | 0 | 100 | 0 | 650 | 300 | 0 |
| mcp_tool__gsheet_create | 350 | 700 | 700 | 700 | 300 | 300 | 0 | 200 | 0 | 600 | 350 | 0 |
| mcp_tool__gsheet_copy_spreadsheet | 345 | 650 | 600 | 750 | 300 | 400 | 0 | 100 | 0 | 700 | 300 | 0 |
| mcp_tool__gsheet_search | 395 | 300 | 250 | 650 | 350 | 550 | 0 | 1000 | 300 | 600 | 350 | 0 |
| mcp_tool__gsheet_clear_range | 377 | 650 | 500 | 750 | 300 | 350 | 0 | 0 | 600 | 700 | 300 | 0 |
| mcp_tool__gsheet_touch_range | 300 | 500 | 500 | 750 | 300 | 150 | 0 | 0 | 100 | 700 | 300 | 0 |
| env__GOOGLE_CREDENTIALS_FILE | 477 | 500 | 750 | 0 | 1000 | 500 | 0 | 1000 | 500 | 0 | 500 | 500 |
| env__GOOGLE_TOKEN_FILE | 431 | 500 | 750 | 0 | 1000 | 250 | 0 | 1000 | 500 | 0 | 250 | 500 |
| env__GSHEETS_TOKEN_MODE | 477 | 750 | 750 | 0 | 1000 | 0 | 0 | 1000 | 750 | 0 | 500 | 500 |
| env__GSHEETS_BROKER_URL | 500 | 750 | 750 | 0 | 1000 | 250 | 0 | 1000 | 500 | 0 | 500 | 750 |
| env__GSHEETS_BROKER_SESSION_TOKEN | 477 | 750 | 750 | 0 | 1000 | 500 | 0 | 1000 | 250 | 0 | 750 | 250 |
| env__GSHEETS_HEADLESS | 522 | 750 | 750 | 0 | 1000 | 250 | 0 | 1000 | 750 | 0 | 500 | 750 |
| auth__credential_modes | 536 | 600 | 500 | 250 | 500 | 650 | 250 | 850 | 700 | 250 | 700 | 650 |
| error__argument_validation | 590 | 1000 | 1000 | 1000 | 250 | 500 | 0 | 1000 | 1000 | 250 | 500 | 0 |
| error__read_auth_failure | 590 | 1000 | 1000 | 1000 | 250 | 500 | 250 | 1000 | 1000 | 250 | 0 | 250 |
| error__mutation_partial_failure | 431 | 1000 | 1000 | 1000 | 250 | 250 | 0 | 0 | 1000 | 250 | 0 | 0 |

## Distribution

```text
   0- 99 |  (0)
 100-199 |  (0)
 200-299 | ## (2)
 300-399 | ####### (7)
 400-499 | ######## (8)
 500-599 | ###### (6)
 600-699 |  (0)
 700-799 |  (0)
 800-899 |  (0)
 900-999 |  (0)
1000     |  (0)
```

All 23 scored surfaces are below the 750 Polish Bar. The first implementation
pass targets the shared boundary and workflow defects that depress the full
distribution rather than polishing individual docstrings in isolation.
