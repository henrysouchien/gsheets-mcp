# gsheets-mcp

A strict, agent-friendly MCP server for Google Sheets. Version `1.0.0` exposes
wire contract `1.0` through one explicit CLI and direct structured results.

## Tool contract

| Tool | Inputs | Broker | Local |
|---|---|---:|---:|
| `gsheets_list_tabs` | `spreadsheet` | yes | yes |
| `gsheets_read_range` | `spreadsheet`, `range`, optional render options | yes | yes |
| `gsheets_write_range` | `spreadsheet`, `range`, `values` | yes | yes |
| `gsheets_append_rows` | `spreadsheet`, `range`, `values` | yes | yes |
| `gsheets_create_spreadsheet` | `title` | yes | yes |
| `gsheets_copy_spreadsheet` | source `spreadsheet`, destination `title`, optional `tabs` | yes | yes |
| `gsheets_search_spreadsheets` | `query`, optional `limit` | no | yes |
| `gsheets_clear_range` | `spreadsheet`, `range` | yes | yes |
| `gsheets_recalculate_range` | `spreadsheet`, `range` | yes | yes |

`spreadsheet` accepts only a spreadsheet ID or a strict
`https://docs.google.com/spreadsheets/...` URL. Common tools never treat a title
as an identifier. In local mode, search by title first and pass the selected
`results[n].spreadsheet` ID to the next tool. `range` is A1 notation, preferably
including the tab (for example, `Comps!A1:H20`).

Inputs are closed-world, strictly typed JSON. Unknown fields—including plausible
fields such as `dry_run` and `idempotency_key`—are rejected. `values` must be a
non-empty two-dimensional JSON array, not a JSON-encoded string.

## Results and failures

Successful calls return a direct object rather than JSON inside a string:

```json
{
  "status": "ok",
  "operation": "gsheets_read_range",
  "spreadsheet": "1AbCdEf...",
  "range": "Comps!A1:B2",
  "value_render_option": "FORMATTED_VALUE",
  "date_time_render_option": "FORMATTED_STRING",
  "values": [["Ticker", "PCTY"], ["Revenue", 100]]
}
```

Operational and validation failures set MCP `isError=true`. The structured
content carries a stable code plus explicit outcome and retry guidance:

```json
{
  "status": "error",
  "operation": "gsheets_write_range",
  "error": {
    "code": "operation_outcome_uncertain",
    "message": "The operation may have changed external state, but confirmation was lost.",
    "outcome": {
      "state": "uncertain",
      "phase": "write_range",
      "mutation_may_have_occurred": true
    },
    "retry": {
      "safe": false,
      "automatic": false,
      "action": "inspect_before_retry",
      "retry_after_seconds": null
    },
    "validation": null,
    "recovery": null,
    "incident_id": null
  }
}
```

Never blindly repeat append, create, copy, write, or recalculate after an
uncertain result. Copy failures can include destination progress. Recalculation
accepts only formulas plus blank cells, restores the exact formula matrix, and
verifies formulas; it does not claim that custom-function values are fresh.

## Install and inspect

Prerequisites are Python 3.10+ and a Google Cloud OAuth desktop client with the
Sheets and Drive APIs enabled.

```bash
python3 -m venv venv
source venv/bin/activate
python -m pip install -e .

gsheets-mcp --help
gsheets-mcp --version
gsheets-mcp capabilities --json
```

The bare command prints help and exits. Only `gsheets-mcp serve` starts the stdio
protocol server, and stdout is reserved for MCP frames while it runs. Discovery
and capabilities inspection do not load credentials.

## Local OAuth mode

Local mode is the default. Save the OAuth desktop client JSON as
`drive_credentials.json` in the repository root, or set
`GOOGLE_CREDENTIALS_FILE`. The token cache defaults to `token.pickle`; override
it with `GOOGLE_TOKEN_FILE`. Both files are ignored by Git.

Run interactive consent once from a trusted local terminal:

```bash
source venv/bin/activate
python -c 'from gsheets_mcp.sheets_client import authenticate; authenticate()'
```

Set `GSHEETS_HEADLESS=1` in non-interactive environments so missing consent fails
closed. A local MCP configuration uses the installed console script explicitly:

```json
{
  "mcpServers": {
    "gsheets-mcp": {
      "type": "stdio",
      "command": "/absolute/path/to/gsheets-mcp/venv/bin/gsheets-mcp",
      "args": ["serve"]
    }
  }
}
```

## Broker mode

Set `GSHEETS_TOKEN_MODE=broker` when a trusted token broker supplies per-user,
in-memory credentials. Calls then require `GSHEETS_BROKER_URL` and
`GSHEETS_BROKER_SESSION_TOKEN`. Broker discovery exposes exactly eight tools;
Drive title search is intentionally absent. Do not put broker tokens in tracked
configuration or command arguments.

```bash
GSHEETS_TOKEN_MODE=broker gsheets-mcp capabilities --json
GSHEETS_TOKEN_MODE=broker gsheets-mcp serve
```

An invalid mode fails closed. Broker credentials refresh proactively near expiry,
and Google 401 handling is bounded to one exact-request credential refresh.

## Development

```bash
source venv/bin/activate
python -m pip install -e '.[dev]'
python -m pytest -q
python scripts/mcp_lint.py
```

The lint command imports the live registry without credentials and verifies the
mode-specific tool sets, strict schemas, annotations, versions, dependency pins,
and absence of retired runtime paths.

## License

MIT
