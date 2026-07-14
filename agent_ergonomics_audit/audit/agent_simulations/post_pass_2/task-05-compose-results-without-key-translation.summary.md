# Task 05 — compose-results-without-key-translation

**Result:** PASS
**Round-trips:** 2 (one replay plus one optional corrected replay)
**First-try success:** partial at the simulator verifier; every MCP call itself succeeded

The isolated in-process replay exercised C001–C005 through the real MCP `CallToolRequest` input/result boundary:

- C001 search → list
- C002 list → read
- C003 read → write
- C004 create → write
- C005 copy → read

Every upstream and downstream result was a direct `status: "ok"` object. The first replay's local assertion compared propagated fields against the downstream *result*; that was wrong for nested search output and for write results that intentionally do not echo values. The optional replay corrected the verifier to compare the source result keys with the composed downstream *input* keys, and all five cases passed without key translation.

This is explicitly **non-live test-double evidence**: every row records `test_double: true` and `live_google_call: false`. Credential and service constructors were replaced with guards that raise immediately if touched.
