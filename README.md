# gsheets-mcp

MCP server for Google Sheets operations.

## Tools

| Tool | Description |
|------|-------------|
| `gsheet_list_tabs` | List tabs in a spreadsheet by name or ID |
| `gsheet_read_range` | Read values from a range |
| `gsheet_update_range` | Update a range with user-entered values |
| `gsheet_append_rows` | Append rows to a range |
| `gsheet_create` | Create a new spreadsheet |
| `gsheet_search` | Search Drive for spreadsheets by name |
| `gsheet_clear_range` | Clear all values in a range |
| `gsheet_touch_range` | Rewrite formulas in a range to trigger recalculation |

## Setup

### Prerequisites
- Python 3.10+
- Google account with Sheets and Drive API access
- OAuth desktop client credentials from Google Cloud Console

### Installation
```bash
git clone https://github.com/<your-user>/gsheets-mcp.git
cd gsheets-mcp
python3 -m venv venv
source venv/bin/activate
pip install -e .
```

### Authentication
1. Save your OAuth desktop credentials file as `drive_credentials.json` in the repository root.
2. Run an auth bootstrap once:
```bash
source venv/bin/activate
python -c "from src import sheets_client; sheets_client.authenticate()"
```
3. Complete browser consent. A local `token.pickle` cache will be created.

### Claude Code Configuration
Add to `~/.claude.json`:

```json
{
  "mcpServers": {
    "gsheets-mcp": {
      "type": "stdio",
      "command": "/path/to/gsheets-mcp/venv/bin/python",
      "args": ["/path/to/gsheets-mcp/run_server.py"]
    }
  }
}
```

## Development

```bash
source venv/bin/activate
pytest
python run_server.py
```

## License
MIT
