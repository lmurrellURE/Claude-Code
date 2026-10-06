# Handoff: Accrued Commissions / Unclaimed Property Analysis

Written 10/6/2026 at the end of session https://claude.ai/code/session_01HGgbDQiKPymLDjGkpqtC9H.
Read this whole file before doing anything. Everything the next session needs is either in this
folder or recorded here. The uploads from the prior session (cash workbook, agreements, fee
documents, entity database) are NOT available in a new session; the facts taken from them are below.

## 1. What exists in this folder (branch `claude/accrued-commissions-analysis-mmu04n`)

| File | What it is |
| --- | --- |
| `Accrued_Commissions_Unclaimed_Property_Analysis.xlsx` | The deliverable. Tabs: Summary, Detail (181 items), Current Month, Gaps & Next Steps, ICA Terms, Method, State Rules. All formulas; recalculated with LibreOffice; zero errors. |
| `build_workbook.py` | Rebuilds the workbook from the two CSVs. Edit this, not the xlsx. Run `python3 build_workbook.py <out.xlsx>` from the folder containing the CSVs, then recalc with the xlsx skill's `recalc.py`. |
| `source-extract/accrued_rows.csv` | Every row of every entity's "Previous Month's Accrued Commissions" block, booked or not. |
| `source-extract/current_month_unbooked.csv` | Current-month unbooked deposits (all Sept/Oct 2026). |
| `README.md` | Memo-style summary with the current figures. Keep its numbers in sync with the workbook. |
| `Accrued_Commissions_Unclaimed_Property_SOP.md` / `.docx` | Snapshots of the SOP. The editable master is the Claude Doc: https://claude.ai/code/artifact/047dd24f-815c-4fdd-8c8c-414c9201c552. The SOP is a STANDING POLICY with no figures in it (for the team and for auditors); it is not updated each quarter. Point-in-time figures and the backlog plan live in the README (which serves as the 10/6/2026 findings memo) and the workbook. |

**How the user wants the process to work (stated 10/6/2026, after reviewing the SOP draft):**
- Deposits are applied in the ordinary course with NO deadline. Accrued commissions often sit a long time and get applied much later; that is normal and not a control failure. Do not reintroduce 30-day or 90-day matching rules, aging KPIs, or "clear the backlog in a quarter" targets.
- The quarterly review is about items that have hit the dormancy THRESHOLD (or will hit it before the next state cutoff) and getting those officially off the books: pay the agent, refund the payer, remit to the state, or retain with counsel support. Everything younger is left alone.
- The SOP stays evergreen. Each quarter produces a short findings memo with that quarter's figures; the SOP itself is only changed when the policy changes (e.g. a verified statute citation in the appendix).

The user also keeps copies at Desktop > Claude Code > Unclaimed Property Analysis. Nothing in a cloud session can write there; they download from GitHub or the chat file cards.

## 2. The task that remains: verify the state rules against primary sources

The prior session's network policy blocked every statute and state-treasury site, so the State Rules
tab (in `build_workbook.py`, the `STATE_ROWS` list and `LIKELY` dict), the README, and the SOP
appendix rest on search-engine summaries and secondary sources. The user has now set the
environment to full network access. The job is to open the primary pages, confirm or correct each
rule, and record the citation and URL actually read.

Verify, in this order (highest dollar exposure first):

1. **Texas** (five of seven holder entities; ~85% of the aged balance)
   - Prop. Code 72.101 (3-year general) and 72.1015 (1-year wages, wages per Labor Code 61.001): https://statutes.capitol.texas.gov/docs/PR/htm/PR.72.htm
   - Labor Code 61.001 definitions (employee excludes independent contractor): https://statutes.capitol.texas.gov/Docs/LA/htm/LA.61.htm
   - Prop. Code 74.101 (March 1 / July 1), 74.1011 (notice, >$250, 60 days), 74.301, penalty/interest 74.705-74.709, anti-limitation provision if present: https://statutes.capitol.texas.gov/Docs/PR/htm/PR.74.htm
   - Comptroller holder reporting and VDA: https://comptroller.texas.gov/programs/claim-it/holders/ and https://www.comptroller.texas.gov/taxes/publications/96-576.php
   - Confirm whether the Comptroller publishes a dormancy for NAUPA code MS02 (commissions) that differs from the 3-year general rule.
2. **Pennsylvania** (Philly; 2-year wages/commissions claimed from the Treasury matrix)
   - Dormancy matrix: https://www.patreasury.gov/pdf/unclaimed-property/Dormancy-Matrix.pdf
   - Holder page (April 15 deadline, December 31 cutoff, due diligence 60-120 days, $50): https://www.patreasury.gov/unclaimed-property/holders/
   - Statute 72 P.S. 1301.1 et seq. (any public text; the Treasury site links it). Confirm the section that sets 2 years for wages/commissions and the anti-limitation section.
3. **Illinois** (Chicago; URE Chicago is a Texas LLC, so Illinois only governs items with an IL agent address)
   - 765 ILCS 1026/15-201 (1 year for wages, commissions, other compensation for personal services; 3-year catch-all): https://www.ilga.gov/legislation/ilcs/documents/076510260K15-201.htm
   - 15-403 report deadline for business associations (the one unconfirmed Illinois fact: Nov 1 / June 30 vs. May 1 / Dec 31): https://www.ilga.gov/legislation/ilcs/documents/076510260K15-403.htm
   - 15-501 due diligence ($50, 60-365 days) and the anti-limitation section (15-610 or nearby).
   - Treasurer holder FAQ: https://icash.illinoistreasurer.gov/app/faq
4. **Florida** (Gallery = RaySon Partners, LLC, Florida)
   - 717.102 (5-year general), 717.115 (1-year wages), 717.117 (report before May 1, $50 listing, due diligence), 717.129 or nearby anti-limitation: https://www.flsenate.gov/Laws/Statutes/2024/Chapter717/All
   - Holder page and VDA: https://www.fltreasurehunt.gov/Holder.jsp
5. **Alabama** (Leading Edge Realty)
   - 35-12-72 (1 year wages/other compensation; 3 years other), 35-12-76 (Nov 1 report, 12 months to July 1), 35-12-77 (due diligence), anti-limitation section: https://treasury.alabama.gov/wp-content/uploads/2024/07/UCP-law.pdf
6. **Missouri** only matters if any holder were formed there; none is. Leave as a note.

For each state, record in `STATE_ROWS`: the exact period, the section number, the report due date and cutoff, the due-diligence threshold and timing, VDA availability, and the anti-limitation provision text (this decides whether the Philly 60-day forfeiture and TUR forfeiture clauses can be relied on). Change the `Verification status` text from "not verified" to "verified from <URL> on <date>" and put the URL in `Sources`.

Then: rebuild the workbook, recalc, read back the Summary, update README figures if any period changed, update the SOP appendix (Claude Doc via the docs tools, then re-export to the .md/.docx snapshots), commit, push to the same branch, and send the user the workbook.

## 3. Facts established with the user (do not re-ask)

- Offices and states: Dallas, Houston, TUR = Texas; Chicago = Illinois; Philly = Pennsylvania; Leading Edge = Alabama; Gallery = Florida. DC tab exists with $0.
- Holder entities and domicile (from the Entity Management Database): URE Dallas, L.L.C. (TX, formed 12/15/2010); URE Houston, L.L.C. (TX, 6/3/2011); Quick Close Properties, LLC dba Texas United Realty (TX); URE Chicago, L.L.C. (TX, 8/10/2011; foreign-registered in IL); URE Philadelphia, L.L.C. (TX, 10/6/2011; registered in PA, MD, NJ); RaySon Partners, LLC d/b/a URE Gallery (FL); Leading Edge Realty (AL, 55% JV); URE Washington DC, L.L.C. (TX). Consequence: unknown-agent items at Chicago and Philly default to Texas under the second priority rule.
- Agents are independent contractors (confirmed). Nothing had been done with the unapplied deposits before this analysis; no prior write-offs.
- Fee check amounts (user): $644, $1,044, $69, and "other small amounts". Philly also $495, $595, $295, $995, $49. From the ICA fee schedules: $595/$995 transaction fee per side + $49 E&O ($644 = 595+49, $1,044 = 995+49), $69 = $65 monthly + $4 card fee, $110 enhanced E&O (Philly personal sales). Older schedule (JV ICA Transaction Fee Summary): $495/$895 + $45/$49 E&O. TUR legacy plan: $150/$250/$350/$450 by commission size, $60 lease, $179 annual. Leading Edge: sale fee by price band $100/$395/$695/$995/$1,295/$1,595/$1,895/$2,195/$2,495/$2,795, E&O $0 or $35, fee after cap $150. All are inputs on Summary rows 4-7.
- Small-amount threshold $250 is the analyst's assumption (Summary B8).
- The user may not be able to find the Disbursement Authorization or Closing Disclosure for many old items. Agreed estimation rule: agent share = check less the standard fee (lower tier when price unknown), documented per item.
- ICA clauses that matter (details on the ICA Terms tab): Houston TXR-2301 para 16.C makes the agent's fee payable when the Broker receives its fee; Dallas/Chicago/Philly pay promptly after receipt subject to a complete file; every ICA has offset/hold rights; Philly ICA para 31 forfeits the agent share on files incomplete 60 days after closing and keeps 100% on a license lapse not cured in 30 days; TUR forfeits on certain breaches; Chicago ICA chooses Texas law.

## 4. Current figures (as of 10/6/2026, in the workbook)

- Prior-month accrued (unbooked): 226,764.46 across 181 items. Current-month: 144,542.00, all Sept/Oct 2026.
- Classes: fee amounts 48 items / 20,825.00; small 15 / 2,058.33; memo-flagged 16 / 30,640.63; larger (likely commission) 102 / 173,240.50.
- Likely-test exposure (larger + memo-flagged, governing state = formation state unless an owner state is entered): dormant today 13,407.75; at next cutoff 21,256.75. Dallas 6,707.80; Houston 5,636.60; TUR 644.00; Philly 419.35.
- If Texas were held to its 1-year wage rule the Texas figures would be far higher (Summary cols T-U). This is the single most valuable legal question to settle.

## 5. Conventions to keep

- Do not fabricate rules. If a page cannot be opened, say so and leave the verification status honest.
- Fonts Arial; formulas not hardcodes; recalc before delivering; zero formula errors.
- Commit messages end with the Co-Authored-By and Claude-Session lines the harness provides.
- Push only to `claude/accrued-commissions-analysis-mmu04n` unless the user says otherwise. Do not open a PR unless asked.
- The user's Desktop folder cannot be written from the cloud; give GitHub links (blob for .md, raw for .xlsx/.docx).
- Never commit the Entity Management Database or the cash workbook; they contain EINs, account numbers and a stored password.
