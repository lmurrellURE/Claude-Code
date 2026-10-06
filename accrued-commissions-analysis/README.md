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

## Classification of the prior-month items

Fee amounts come from the user, the five independent contractor agreements, the ICA Fee Structure
Cheat Sheet (current $595/$995 + $49 schedule, TUR legacy plan, Leading Edge price-band schedule)
and the JV ICA Transaction Fee Summary (older $495/$895 + $45/$49 schedule). They are inputs on
Summary rows 4-7 (all entities, Philly-only, TUR-only, Leading Edge-only). Small-amount threshold
$250 (adjustable). Agents are independent contractors. Many older checks cannot be tied to an
agent. Office states: Dallas, Houston, TUR = TX; Chicago = IL; Philly = PA; Leading Edge = AL;
Gallery = FL. Per the Entity Management Database, URE Dallas, URE Houston, Quick Close Properties
(TUR), URE Chicago and URE Philadelphia are all TEXAS LLCs (Chicago and Philadelphia are only
registered as foreign LLCs in IL and PA); RaySon Partners d/b/a URE Gallery is Florida; Leading
Edge Realty is Alabama. The formation state is each entity's default governing state under the
second priority rule; an item switches to the agent's address state once that is entered.

| Class (Detail col N) | Items | $ |
|---|---:|---:|
| Fee amount (exact match) - presumed company revenue | 48 | 20,825.00 |
| Small amount (at or under $250) - likely fee, review | 15 | 2,058.33 |
| Memo-flagged (not ours / unknown / duplicate / hold) - resolve with payer | 16 | 30,640.63 |
| Larger amount - likely includes an agent commission split | 102 | 173,240.50 |

## Exposure (upper bounds, larger-amount and memo-flagged items, governing state = formation state until an agent address is entered)

| Entity | Default governing state / likely test | Dormant today | Dormant at next report cutoff |
|---|---|---:|---:|
| Dallas | TX general, 3 yrs | 6,707.80 | 8,042.30 |
| Houston | TX general, 3 yrs | 5,636.60 | 6,011.60 |
| TUR | TX general, 3 yrs | 644.00 | 2,373.50 |
| Philly | TX general, 3 yrs (PA 2 yrs only for items with a PA agent address) | 419.35 | 4,829.35 |
| Chicago | TX general, 3 yrs (IL 1 yr only for items with an IL agent address) | 0.00 | 0.00 |
| Gallery | FL 3 yrs (717.1035, Florida holder, unknown owner address; 5 yrs under 717.102 if a FL address is known) | 0.00 | 0.00 |
| Leading Edge | AL compensation, 1 yr | 0.00 | 0.00 |
| **Total** | | **13,407.75** | **21,256.75** |

If Texas were instead held to its 1-year wage rule the same population would be 94,496.40 dormant
today and 106,204.14 at the next cutoff (Summary cols T and V: Dallas 17,639.54 / 21,952.66, Houston
47,791.00 / 48,641.00, TUR 13,942.20 / 15,339.70, Philly 15,123.66 / 20,270.78). Verified 10/6/2026
from the statute text: the Texas 1-year rule (Prop. Code 72.1015) uses the Labor Code 61.001
definition of wages, which is compensation owed by an employer to an employee and expressly excludes
independent contractors, and the Comptroller's own property-code table (Pub. 96-478, rev. March 2026)
lists commissions (code MS02) at 3 years and wages (MS01) at 1 year. The 3-year test is therefore the
Texas position; counsel need only confirm that the agents' contractor status holds.
Philly's figure falls from 8,566.41 to 419.35 because its unknown-agent items are now tested under
Texas's 3-year period rather than Pennsylvania's 2-year period; entering a PA address for an item
restores the 2-year test for that item.

Even these are upper bounds: within a larger check only the agent's share is reportable, the
company's split is revenue, and every ICA gives the Broker offset rights for amounts the agent
owes. The Philly ICA also forfeits the agent's share on files still incomplete 60 days after
closing and retains 100% where a license is not reactivated within 30 days. Verified 10/6/2026:
Texas (74.308), Illinois (15-610), Florida (717.129) and Alabama (35-12-88) each provide that the
expiration of a period set by contract does not prevent property from being presumed abandoned, and
Texas 74.309 bars taking funds into income by private agreement to circumvent the process.
Pennsylvania's 1301.16 mentions only statute and court order, so the Philly clause is a contract-law
question there; most Philly items default to Texas in any case. Treat a forfeiture as the company's
money only with counsel's support and documented facts per item.

## Files

- `HANDOFF.md` - everything a new session needs to continue (files, established facts, figures,
  statute pages to verify). `NEXT_SESSION_PROMPT.md` - the prompt to paste into that session.

- `Accrued_Commissions_Unclaimed_Property_SOP.md` / `.docx` - the standing SOP (ongoing application of
  deposits, quarterly unclaimed property review, treatment rules, state calendar, accounting, initial
  remediation, appendix with the verified citations). Snapshot of the editable Claude Doc.
- `Accrued_Commissions_CFO_Controller_Briefing.docx` - one-page briefing for the CFO and Controller
  (10/6/2026): the decision requested, the balance, exposure under the verified rules with the Texas
  1-year sensitivity, cash and P&L impact, the four sign-offs, the timeline to the 2027 filings, and
  what is still unverified. Snapshot of the editable Claude Doc.

- `Accrued_Commissions_Unclaimed_Property_Analysis.xlsx` - Summary (fee amounts and small-amount
  threshold are inputs in rows 4-5), Detail (181 items with yellow fill-in columns for owner, owner
  state and the fee/agent/refund split), Current Month, State Rules (with the likely test per state
  for independent-contractor commissions), Gaps & Next Steps, Method. Formulas re-age and reclassify
  when the as-of date, fee amounts, threshold or a state period changes.
- `build_workbook.py` - rebuilds the workbook from the two CSV extracts.
- `source-extract/` - the rows extracted from the cash-requirements workbook.

## Verification of the state rules (10/6/2026)

Every rule on the State Rules tab for Texas, Illinois, Pennsylvania, Florida and Alabama was read
from the statute text and the state treasury holder pages on 10/6/2026; the URLs actually read are
in the tab's Sources column and the Verification status column says what each one confirmed.

Confirmed as drafted: TX 1-year wages / 3-year general, July 1 report, March 1 cutoff, due diligence
over $250 at least 60 days before delivery, VDA under 74.707 with a 10-year look-back; IL 1-year
compensation / 3-year catch-all, $50 due diligence 60 days to one year before filing, VDA; PA 2-year
wages and other compensation / 3-year general, April 15 report, December 31 cutoff, $50 due diligence
60-120 days before April 15, VDA with 10-year look-back; FL 1-year wages, May 1 report, December 31
cutoff, $50 due diligence 60-120 days before filing, VDA; AL 1-year compensation / 3-year other,
November 1 report, June 30 cutoff, $50 due diligence 60 days before filing, VDA.

Corrected: Illinois business holders file by May 1 for the calendar year (the prior draft carried the
November 1 / June 30 cycle, which applies only to financial organizations, insurers and governments),
so the next Illinois cutoff is 12/31/2026, not 6/30/2027. Florida's operative long period for a
Florida-organized holder whose owner address is unknown is 3 years under 717.1035, not the 5-year
general rule in 717.102; the Gallery row now uses 3 years. The Pennsylvania commissions rule is
72 P.S. 1301.10(2), not 1301.5. The Texas, Pennsylvania and Florida VDA descriptions now follow the
treasury pages (Texas is not limited to first-time filers on anything the Comptroller publishes;
Pennsylvania's program is current and waives penalties and interest; Florida's look-back is ten
report years, not five). None of these corrections changes a dollar figure above: Chicago and Gallery
have no aged items.

Still open: no court decision or Comptroller ruling specifically on broker-to-agent commissions was
found, so the Texas 3-year position rests on the statutory definitions and the Comptroller's MS02
table; whether a court would enforce the Philly ICA 60-day forfeiture against an escheat claim;
deposit date is a proxy for the payable date; Missouri and DC were not researched (no holder is formed
in Missouri and the DC balance is $0); and Texas VDA eligibility for an entity that has filed before is
not stated on the Comptroller's pages. Nothing here is legal advice.
