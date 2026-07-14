# Task 02 — compare-local-and-broker-discovery

**Result:** PASS
**Round-trips:** 2 (target: 2)
**First-try success:** yes

Credential-free capabilities discovery reported nine local tools and eight broker tools. The local-only difference was exactly `gsheets_search_spreadsheets`; the other eight names and their order matched.

Broker mode was selected only through `GSHEETS_TOKEN_MODE=broker`. No broker URL or session token was supplied or loaded. `HOME=/var/empty` and explicit nonexistent Google credential/token paths prevented ambient credential use.
