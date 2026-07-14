# Task 01 — discover-the-contract

**Result:** PASS
**Round-trips:** 2 (target: 2)
**First-try success:** yes

The bare `gsheets-mcp` invocation printed useful help and exited 0. Its help exposed `serve` and the credential-free `capabilities` command. The second invocation, `gsheets-mcp capabilities --json`, returned package `gsheets-mcp`, package version `1.0.0`, contract version `1.0`, mode `local`, and `tool_count: 9`.

The installed executable was invoked by its bare command name with the repository virtual environment placed on a scrubbed `PATH`. No external documentation, Google credential, OAuth flow, or Google API was used.
