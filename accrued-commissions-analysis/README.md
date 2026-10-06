# Accrued Commissions / Unclaimed Property Analysis

As of 10/6/2026. Source: `United_Cash_Requirements.xlsx` (not committed; contains bank detail).

## What the accrued commission balance is

Each entity tab keeps a "Previous Month's Accrued Commissions" block: checks that were deposited
(mostly from title companies, landlords and other brokers) but have not yet been matched to a
transaction and booked in Sage. The accrued balance is the sum of the rows with no "x" in the
"accounted for in Sage" column. The current-month "Deposited checks" block works the same way.

Every entity's extracted total ties exactly to the tab's "outstanding amount" cell.

| Entity | Prior-month accrued | Items | Oldest | Current-month accrued | Total |
|---|---:|---:|---|---:|---:|
| Dallas | 59,991.62 | 50 | 03/03/2023 | 15,586.87 | 75,578.49 |
| Houston | 76,889.50 | 40 | 01/31/2022 | 10,071.93 | 86,961.43 |
| TUR | 55,528.45 | 50 | 01/25/2023 | 20,054.20 | 75,582.65 |
| Chicago | 0.00 | 0 | - | 18,345.00 | 18,345.00 |
| Philly | 29,708.78 | 34 | 09/06/2023 | 49,091.50 | 78,800.28 |
| Gallery | 4,001.11 | 4 | 05/14/2026 | 0.00 | 4,001.11 |
| Leading Edge | 645.00 | 3 | 04/01/2026 | 31,392.50 | 32,037.50 |
| DC | 0.00 | 0 | - | 0.00 | 0.00 |
| **Total** | **226,764.46** | **181** | | **144,542.00** | **371,306.46** |

All current-month items are September/October 2026 deposits.

## Aging of the prior-month balance (from deposit date)

| Entity | 0-1 yrs | 1-2 yrs | 2-3 yrs | 3-5 yrs |
|---|---:|---:|---:|---:|
| Dallas | 39,767.62 | 6,244.96 | 7,179.74 | 6,799.30 |
| Houston | 26,416.50 | 35,650.90 | 9,185.50 | 5,636.60 |
| TUR | 34,473.00 | 5,539.51 | 13,657.94 | 1,858.00 |
| Philly | 8,931.12 | 10,077.25 | 10,281.06 | 419.35 |
| Gallery | 4,001.11 | - | - | - |
| Leading Edge | 645.00 | - | - | - |
| **Total** | **114,234.35** | **57,512.62** | **40,304.24** | **14,713.25** |

## Exposure bounds (upper bounds, before the fee/agent split is known)

| Test | Dormant today | Dormant at each state's next report cutoff |
|---|---:|---:|
| Short period (wages / compensation for personal services) | 102,452.86 | 115,241.60 |
| Long period (general catch-all) | 14,713.25 | 18,856.25 |

These are the amounts of unbooked deposits older than the state period, assuming 100% of each
check is owed to someone other than the company. The amount actually reportable is only the agent
share plus anything that belongs back to the payer. The company's own fee / E&O / company-dollar
share is revenue, not unclaimed property. The source workbook has E&O / FEE / Commission columns
for this split and they are empty for every item.

## Files

- `Accrued_Commissions_Unclaimed_Property_Analysis.xlsx` - Summary, Detail (181 items with yellow
  fill-in columns for owner, owner state and the fee/agent/refund split), Current Month, State
  Rules, Gaps & Next Steps, Method. Formulas re-age when the as-of date or a state period changes.
- `build_workbook.py` - rebuilds the workbook from the two CSV extracts.
- `source-extract/` - the rows extracted from the cash-requirements workbook.

## Caveat

The State Rules tab is built from search-engine summaries of the statutes and from state treasury
guidance; the statute pages themselves were blocked by the session's network policy. It is a
starting point for counsel, not a filing position.
