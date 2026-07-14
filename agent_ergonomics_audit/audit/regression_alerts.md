# Pass 2 Regression Alerts

Compared with Pass 1 across the same 23 stable comparison handles. Retired
`gsheet_*` names are represented only by their historical `surface_id`; each
Pass 2 row identifies its canonical full-cutover equivalent.

## Hard stops (weighted drop > 50 points)

None.

## Dimension warnings (single-dimension drop > 25 points)

None.

## Strict no-decrease check

- Weighted score decreases: **0 of 23 surfaces**.
- Individual dimension decreases: **0 of 253 comparisons**.
- Minimum weighted uplift: **+341 points**.
- Median weighted uplift: **+445 points**.
- Maximum weighted uplift: **+613 points**.

The check compares every Pass 2 dimension with the matching Pass 1 dimension;
no safety or regression-resistance score decreased.

## Post-score review hardening

Fresh-eyes review found and closed three gaps without changing the scored
contract: required wire discriminators, a one-write ceiling for uncertain-clear
compensation, and canonical-only fuzzy suggestions. The corresponding R-001,
R-002, and R-003 audit regressions pin those invariants. Two independent final
reviews returned PASS.
