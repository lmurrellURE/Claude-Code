> Executed on 10/6/2026. Decisions made after this prompt ran (owner, Texas 3-year rule, forfeiture clauses, revenue timing, one annual cycle, no sensitivity columns on the one-pager) are recorded in HANDOFF.md; this file is kept as history.

# Prompt for the next session

Copy everything below the line into a new session started on the lmurrellURE/Claude-Code repository.

---

Continue the accrued commissions / unclaimed property analysis from the previous session.

First, check out the branch `claude/accrued-commissions-analysis-mmu04n` (fetch it from origin if it is not local) and read `accrued-commissions-analysis/HANDOFF.md` in full before doing anything else. It lists the files, the facts already established with me, the current figures, and the exact statute pages to verify. Do not re-ask anything that file answers.

The environment now has full network access. Your job is section 2 of the handoff: open the primary statute and state-treasury pages for Texas, Pennsylvania, Illinois, Florida and Alabama, confirm or correct every rule on the State Rules tab (periods, section numbers, report cutoff and due date, due-diligence threshold and timing, VDA availability, and each state's anti-limitation provision), and record the URL actually read and the date in the Verification status and Sources columns. Texas first; it carries most of the exposure. Pay particular attention to whether Texas treats independent-contractor commissions under the 1-year wage rule or the 3-year general rule, and to whether any state's anti-limitation provision would override the Philly 60-day forfeiture clause.

Then rebuild the workbook with `build_workbook.py`, recalculate it, check that the Summary still ties to 226,764.46, update the README figures if any period changed, update only the appendix of the SOP in the Claude Doc at https://claude.ai/code/artifact/047dd24f-815c-4fdd-8c8c-414c9201c552 with the verified citations (the SOP is a standing policy with no figures; leave the rest of it alone) and re-export the .md and .docx snapshots, commit, push to the same branch, and send me the workbook. Treat the README as the 10/6/2026 findings memo; update its figures if the verified rules move them.

After the verification is done, write a one-page briefing for my CFO and Controller so they can review the plan and its impact. Make it a Claude Doc, and also save a .docx snapshot to the same repository folder. One page means one page: lead with the decision we are asking them to make, then the balance and what it is made of, the exposure under the verified rules by entity and state (with the 1-year versus 3-year Texas sensitivity), what the company keeps and books to revenue, the cash and P&L impact of the remediation plan, the four decisions that need their sign-off (Texas period, Philly forfeiture clause, Texas VDA, estimation rule), the timeline to the 2027 filings, and what is still unverified. Numbers from the rebuilt workbook only; no new assumptions. Use one small table for the entity figures and plain sentences for the rest.

Do not fabricate anything. If a page will not open, say so and leave that rule marked unverified. Report at the end: which rules were confirmed, which changed and by how much the exposure moved, and which are still unverified and why.
