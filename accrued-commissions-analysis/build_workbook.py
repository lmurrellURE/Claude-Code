"""Build the Accrued Commissions / Unclaimed Property analysis workbook.

Inputs : accrued_rows.csv, current_month_unbooked.csv (extracted from United_Cash_Requirements.xlsx)
Output : Accrued_Commissions_Unclaimed_Property_Analysis.xlsx
"""
import csv, datetime, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

OUT = sys.argv[1] if len(sys.argv) > 1 else "Accrued_Commissions_Unclaimed_Property_Analysis.xlsx"
AS_OF = datetime.date(2026, 10, 6)

F = "Arial"
f_norm = Font(name=F, size=10)
f_bold = Font(name=F, size=10, bold=True)
f_title = Font(name=F, size=14, bold=True)
f_input = Font(name=F, size=10, color="0000FF")          # blue = hardcoded input
f_link = Font(name=F, size=10, color="008000")           # green = cross-sheet link
f_hdr = Font(name=F, size=10, bold=True, color="FFFFFF")
fill_hdr = PatternFill("solid", fgColor="1F3864")
fill_in = PatternFill("solid", fgColor="FFFF00")          # yellow = user fills in
fill_sub = PatternFill("solid", fgColor="D9E1F2")
fill_warn = PatternFill("solid", fgColor="FCE4D6")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
CUR = '$#,##0.00;($#,##0.00);-'
DATE = 'mm/dd/yyyy'

def hdr(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row, i, lab); c.font = f_hdr; c.fill = fill_hdr
        c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = box
    if widths:
        for i, w in enumerate(widths, 1): ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[row].height = 42

def put(ws, ref, val, font=f_norm, fmt=None, fill=None, wrap=False):
    c = ws[ref]; c.value = val; c.font = font
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if wrap: c.alignment = Alignment(wrap_text=True, vertical="top")
    return c

# ---------------------------------------------------------------- data
rows = [r for r in csv.DictReader(open("accrued_rows.csv")) if r["booked"] == "False" and r["amount"]]
for r in rows:
    r["amt"] = float(r["amount"])
    r["d"] = datetime.date.fromisoformat(r["date"]) if r["date"] else None
cur = list(csv.DictReader(open("current_month_unbooked.csv")))

ENTITIES = [  # entity, formation state (default governing state), office state, note
    ("Dallas", "TX", "TX", "URE Dallas, L.L.C. - Texas LLC (TX file 32043204570, formed 12/15/2010) per Entity Management Database. Office: Texas."),
    ("Houston", "TX", "TX", "URE Houston, L.L.C. - Texas LLC (TX 0801434954, formed 6/3/2011) per Entity Management Database. Office: Texas."),
    ("TUR", "TX", "TX", "Quick Close Properties, LLC dba Texas United Realty - Texas LLC (TX 0800329212) per Entity Management Database. Office: Texas."),
    ("Chicago", "TX", "IL", "URE Chicago, L.L.C. - TEXAS LLC (TX 0801464335, formed 8/10/2011), registered as a foreign LLC in Illinois (3661385) per Entity Management Database. Office: Illinois. Unknown-owner items therefore default to TEXAS under the second priority rule; items with an Illinois agent address go to Illinois."),
    ("Philly", "TX", "PA", "URE Philadelphia, L.L.C. - TEXAS LLC (TX 0801490897, formed 10/6/2011), registered in PA, MD and NJ per Entity Management Database. Office: Pennsylvania. Unknown-owner items therefore default to TEXAS under the second priority rule; items with a Pennsylvania agent address go to Pennsylvania."),
    ("Gallery", "FL", "FL", "RaySon Partners, LLC d/b/a URE Gallery - Florida LLC (FL L22000422171) per Entity Management Database. Office: Florida."),
    ("Leading Edge", "AL", "AL", "Leading Edge Realty - Alabama (AL 000-273-581) per Entity Management Database (55% JV). Office: Alabama."),
    ("DC", "TX", "DC", "URE Washington DC, L.L.C. - Texas LLC (TX 801575176, formed 3/30/2012) per Entity Management Database. $0 accrued balance; DC rules not researched."),
]
STATE_ROWS = [
    # state, name, short_period_years, short_basis, long_period_years, long_basis, report_due, cutoff, next_cutoff, due_diligence, vda, verification, sources
    ("TX", "Texas", 1, "Unclaimed 'wages' - Tex. Prop. Code 72.1015 (1 year). 'Wages' takes its meaning from Labor Code 61.001, which covers compensation owed by an EMPLOYER to an EMPLOYEE and excludes independent contractors. Independent-contractor agent commissions therefore likely fall under the general rule, not the 1-year rule - confirm with counsel / Comptroller.",
     3, "General personal property - Tex. Prop. Code 72.101(a): presumed abandoned after more than 3 years with owner location unknown and no claim/act of ownership.",
     "July 1 (Tex. Prop. Code 74.101)", "March 1 (property held on March 1 is reported the following July 1)", datetime.date(2027, 3, 1),
     "Notice to owner by mail/e-mail at least 60 days before delivery for property over $250 (Tex. Prop. Code 74.1011).",
     "Comptroller offers a Voluntary Disclosure Agreement with waiver of penalty and interest; not available once an investigation has started or (per secondary sources) to holders who have filed Texas reports before.",
     "Dormancy periods, report date, and due-diligence rule confirmed from search summaries of the statute text; statute pages themselves were blocked by this session's network policy. No business-to-business exemption was found for Texas.",
     "https://statutes.capitol.texas.gov/docs/PR/htm/PR.72.htm ; https://statutes.capitol.texas.gov/Docs/PR/htm/PR.74.htm ; https://texas.public.law/statutes/tex._prop._code_section_72.1015 ; https://texas.public.law/statutes/tex._labor_code_section_61.001 ; https://www.comptroller.texas.gov/taxes/publications/96-576.php"),
    ("IL", "Illinois", 1, "765 ILCS 1026/15-201: 'wages, commissions, bonuses, or reimbursements to which an employee is entitled, or other compensation for personal services' - 1 year after payable. The 'other compensation for personal services' language is broader than Texas and plausibly reaches independent-contractor commissions.",
     3, "765 ILCS 1026/15-201 catch-all: all other property 3 years after the obligation to pay arises (per secondary sources).",
     "November 1 for financial, insurance and government holders (confirmed). Deadline for ordinary business holders NOT confirmed from the statute in this session - the prior Illinois act used May 1 with a December 31 cutoff. CONFIRM.", "June 30 for the November 1 filers; business-holder cutoff NOT confirmed.", datetime.date(2027, 6, 30),
     "Due-diligence letter by first-class mail for property over $50, sent 60 to 365 days before the report (per Illinois Treasurer guidance summary).",
     "Not researched in this session.",
     "Illinois eliminated its business-to-business exemption when it adopted RUUPA (effective 1/1/2018). Chicago's prior-month accrued balance is $0, so Illinois exposure today is only the current-month unbooked deposits.",
     "https://www.ilga.gov/legislation/ilcs/documents/076510260K15-201.htm ; https://icash.illinoistreasurer.gov/app/faq ; https://www.duanemorris.com/articles/illinois_slides_backward_unclaimed_property_law_0917.html"),
    ("PA", "Pennsylvania", 2, "Wages / commissions - 2 years per the PA Treasury Dormancy Matrix as summarized in search results (cited there as 72 P.S. 1301.5). NOT verified from the statute text in this session.",
     3, "General dormancy period 3 years for most property types (72 P.S. 1301.1 et seq.) per secondary sources.",
     "April 15", "December 31 (per secondary sources; NOT verified from statute in this session)", datetime.date(2026, 12, 31),
     "Written notice by first-class mail 60 to 120 days before the April 15 deadline for property of $50 or more with a usable address.",
     "PA's formal amnesty program ended 10/31/2010. Interest of 12% per year from the date the property was due is cited by secondary sources. Ask PA Treasury about current voluntary compliance terms.",
     "All PA facts come from secondary summaries (PA Treasury Dormancy Matrix as quoted, Sovos, UPPO). Statute and Treasury pages were blocked in this session.",
     "https://www.patreasury.gov/pdf/unclaimed-property/Dormancy-Matrix.pdf ; https://www.patreasury.gov/unclaimed-property/holders/ ; https://sovos.com/tax-reporting/unclaimed-property-laws-by-state/pennsylvania/"),
    ("FL", "Florida", 1, "Fla. Stat. 717.115: unpaid WAGES owing in the ordinary course of business unclaimed for more than 1 year after becoming payable. Whether independent-contractor commissions are 'wages' here is not settled in the sources found - if not, the 5-year general rule applies.",
     5, "Fla. Stat. 717.102(1): general 5-year dormancy period for other intangible property.",
     "Before May 1, covering the preceding calendar year (Fla. Stat. 717.117)", "December 31", datetime.date(2026, 12, 31),
     "Report must list owners with property of $50 or more; due-diligence notice requirement in 717.117 (threshold and timing not verified from statute text in this session).",
     "Florida has a Voluntary Disclosure program: no penalties, 5-year look-back instead of 10, 2-year audit window after acceptance (per secondary summary of Florida DFS program).",
     "Dormancy periods and report date confirmed from search summaries of the statute; statute pages blocked. Gallery = Florida is an ASSUMPTION.",
     "https://www.flsenate.gov/Laws/Statutes/2024/717.115 ; https://www.flsenate.gov/Laws/Statutes/2024/717.102 ; https://flsenate.gov/Laws/Statutes/2024/717.117 ; https://www.fltreasurehunt.gov/Holder.jsp"),
    ("AL", "Alabama", 1, "Ala. Code 35-12-72: 'wages or other compensation for personal services' presumed abandoned 1 year after the compensation becomes payable. 'Other compensation for personal services' plausibly reaches independent-contractor commissions.",
     3, "Ala. Code 35-12-72: all other property 3 years after the owner's right to demand it or the obligation to pay arises, whichever is first.",
     "November 1, covering the 12 months preceding July 1", "June 30", datetime.date(2027, 6, 30),
     "Aggregate/reporting threshold $50 per secondary sources; due-diligence notice rule (35-12-77) not verified in this session.",
     "Not researched in this session.",
     "Dormancy periods confirmed from search summaries of the Code; Code pages blocked. Leading Edge = Alabama is an ASSUMPTION.",
     "https://law.justia.com/codes/alabama/2023/title-35/chapter-12/article-2/article-2a/section-35-12-72 ; https://treasury.alabama.gov/alabama-treasury-law/unclaimed-property-law/ ; https://unclaimed.org/reporting/alabama/"),
    ("MO", "Missouri (HQ / possible state of formation)", 3, "RSMo 447.536 as amended by HB 1075 (2014): payroll/wages dormancy reduced from 5 to 3 years effective 1/1/2015.",
     5, "RSMo 447.536: general 5-year dormancy period for most property types.",
     "November 1", "June 30", datetime.date(2027, 6, 30),
     "Due diligence letter required when the owner is due $50 or more (Missouri Treasurer guidance).",
     "Not researched in this session.",
     "Missouri matters only under the SECOND priority rule (Texas v. New Jersey, 379 U.S. 674 (1965)): property whose owner address is unknown goes to the holder's state of incorporation/formation - which is the entity's formation state, NOT its headquarters city. Missouri also has a business-to-business exclusion (HB 1075, 2014) for checks/credits owed to a business entity in the ordinary course of business.",
     "https://revisor.mo.gov/main/OneSection.aspx?section=447.536 ; https://treasurer.mo.gov/UCP/ReportingUnclaimedProperty.aspx ; https://ryan.com/about-ryan/news-and-insights/2014/missouri-legislation-adds-unclaimed-property-business-to-business-exclusions-and-other-changes/"),
    ("DC", "District of Columbia", None, "Not researched - DC accrued balance is $0.", None, "Not researched.", "", "", None, "", "", "Not researched.", ""),
]

wb = Workbook()

# ================================================================ State Rules
sr = wb.active; sr.title = "State Rules"
put(sr, "A1", "Unclaimed property rules by state - dormancy periods applied in this analysis", f_title)
put(sr, "A2", "Short period = the period for wages / compensation for personal services. Long period = the general catch-all period. Which one applies to independent-contractor agent commissions is a legal question that differs by state (see basis columns). Verification column says what could and could not be confirmed from primary sources in this session.", wrap=True)
sr.merge_cells("A2:N2"); sr.row_dimensions[2].height = 45
labels = ["State", "Name", "Short period (yrs)", "Short period basis", "Long period (yrs)", "Long period basis", "Report due", "Report cutoff ('as of')", "Next cutoff date", "Due diligence to owner", "Voluntary disclosure / penalties", "Verification status", "Likely test for independent-contractor commissions (SHORT/LONG)", "Why", "Sources"]
hdr(sr, 4, labels, [7, 18, 9, 55, 9, 45, 28, 28, 12, 40, 45, 45, 12, 55, 70])
LIKELY = {'TX': ('LONG', "Agents are independent contractors (confirmed by user). Texas' 1-year rule covers 'wages' as defined in Labor Code 61.001, which is limited to employer-to-employee compensation and expressly excludes independent contractors. The 3-year general rule is therefore the likely test. Confirm with counsel; some advisors report commissions under NAUPA code MS02 conservatively on the shorter period."), 'IL': ('SHORT', "Illinois' 1-year rule reaches 'other compensation for personal services', not only employee wages, so independent-contractor commissions plausibly fall under it. 3-year general rule is the fallback."), 'PA': ('SHORT', "PA Treasury's dormancy matrix lists wages/commissions at 2 years without an employee limitation in the sources found. 3-year general rule is the fallback. Not verified from statute text."), 'FL': ('LONG', "Fla. Stat. 717.115 covers 'unpaid wages'. Nothing found extends it to independent-contractor commissions, so the 5-year general rule is the likely test. Confirm with counsel."), 'AL': ('SHORT', "Ala. Code 35-12-72 covers 'wages or other compensation for personal services', so independent-contractor commissions plausibly fall under the 1-year rule. 3-year general rule is the fallback."), 'MO': ('LONG', "Missouri's 3-year rule is for payroll checks; commissions to non-employees would fall under the 5-year general rule. Applies only if an LLC is formed in Missouri and the owner's address is unknown."), 'DC': ('', 'Not researched.')}
for i, s in enumerate(STATE_ROWS, 5):
    vals = list(s)
    vals = vals[:12] + [LIKELY[vals[0]][0], LIKELY[vals[0]][1]] + vals[12:]
    for j, v in enumerate(vals, 1):
        c = sr.cell(i, j, v); c.font = f_input if j in (3, 5, 9, 13) else f_norm
        c.alignment = Alignment(wrap_text=True, vertical="top"); c.border = box
        if j == 9 and v: c.number_format = DATE
    sr.row_dimensions[i].height = 150
SR_FIRST, SR_LAST = 5, 4 + len(STATE_ROWS)
put(sr, f"A{SR_LAST+2}", "Priority rules (Texas v. New Jersey, 379 U.S. 674 (1965)): (1) the state of the owner's last-known address on the holder's books; (2) if no address, the holder's state of incorporation/formation. The user confirmed that for many older checks the agent is unknown: those items fall under rule (2) and go to the state where each LLC was formed. Per the Entity Management Database that is TEXAS for URE Dallas, URE Houston, Quick Close Properties (TUR), URE Chicago and URE Philadelphia (the last two are Texas LLCs registered as foreign entities in IL and PA), FLORIDA for RaySon Partners d/b/a URE Gallery and ALABAMA for Leading Edge Realty. Not the office's state, and not Kansas City. Private retention of amounts owed to others ('private escheat') is not permitted by any state; an unknown owner does not make the money the company's.", wrap=True)
sr.merge_cells(f"A{SR_LAST+2}:O{SR_LAST+2}"); sr.row_dimensions[SR_LAST+2].height = 45
put(sr, f"A{SR_LAST+3}", "Blue = hardcoded inputs you can change. All dormancy periods are expressed in years and are applied from the DEPOSIT DATE in the cash-requirements workbook as a proxy for the date the amount became payable - the true trigger is when the agent's share became payable, which may be earlier (closing date) or later.", wrap=True)
sr.merge_cells(f"A{SR_LAST+3}:O{SR_LAST+3}"); sr.row_dimensions[SR_LAST+3].height = 40

# ================================================================ Summary
sm = wb.create_sheet("Summary", 0)
put(sm, "A1", "Accrued Commissions - Unclaimed Property Exposure by Entity", f_title)
put(sm, "A2", "Source: United_Cash_Requirements.xlsx, each entity tab, 'Previous Month's Accrued Commissions' block (deposited checks not yet booked in Sage) plus the current-month 'Deposited checks' block. Extracted 10/6/2026; every entity total ties to the tab's 'outstanding amount' cell. Legend: blue = input, black = formula, green = link to another sheet, yellow fill = confirm or fill in.", wrap=True)
sm.merge_cells("A2:AC2"); sm.row_dimensions[2].height = 32
put(sm, "A3", "As-of date", f_bold); put(sm, "B3", AS_OF, f_input, DATE, fill_in)
sm["B3"].comment = Comment("Input. Change to re-age every item. 10/6/2026 = date the workbook was provided.", "Analysis")
put(sm, "A4", "Fee check amounts - all entities (inputs)", f_bold)
for j, v in enumerate([644, 1044, 69, 595, 995, 49, 110, 495, 895, 544, 944, 45, 75, 100, 125, 50]):
    put(sm, f"{L(3+j)}4", v, f_input, CUR, fill_in)
sm["C4"].comment = Comment("Sources: user 10/6/2026 ($644, $1,044, $69); ICA Fee Structure Cheat Sheet (current: $595/$995 sale fee per side, $49 E&O, $110 enhanced E&O Philly, lease/referral minimums $75/$100/$125, Dallas after-cap fee $50); JV ICA Transaction Fee Summary (older schedule: $495/$895 sale fee, $45 E&O). $644 = 595+49, $1,044 = 995+49, $544 = 495+49, $944 = 895+49, $69 = $65 dues + $4. Remove any amount you do not want treated as a fee. Blank cells are ignored.", "Analysis")
put(sm, "A5", "Additional fee amounts - Philly only (inputs)", f_bold)
for j, v in enumerate([295] + [None]*15):
    put(sm, f"{L(3+j)}5", v, f_input, CUR, fill_in)
sm["C5"].comment = Comment("Per user 10/6/2026: Philly fee amounts are $495, $595, $295, $995 and $49. All but $295 are now in the all-entities row. $295 also appears as the DC referral/lease maximum in the JV fee summary.", "Analysis")
put(sm, "A6", "Additional fee amounts - TUR only (inputs)", f_bold)
for j, v in enumerate([150, 250, 350, 450, 60, 179] + [None]*10):
    put(sm, f"{L(3+j)}6", v, f_input, CUR, fill_in)
sm["C6"].comment = Comment("TUR legacy Transaction Fee Plan (sign-up docs and cheat sheet): $150 / $250 / $350 / $450 by commission size, $60 residential lease (cheat sheet says $100 or 10%), $179 annual fee. TUR agents on the United plan pay $595/$995 + $49, already in row 4.", "Analysis")
put(sm, "A7", "Additional fee amounts - Leading Edge only (inputs)", f_bold)
for j, v in enumerate([100, 395, 695, 995, 1295, 1595, 1895, 2195, 2495, 2795, 35, 150] + [None]*4):
    put(sm, f"{L(3+j)}7", v, f_input, CUR, fill_in)
sm["C7"].comment = Comment("Leading Edge fee schedule from the cheat sheet: sale fee by price band $100 / $395 / $695 / $995 / $1,295 / $1,595 / $1,895 / $2,195 / $2,495 / $2,795; E&O $0 (or $35 in holding company); fee after cap $150. All three aged Leading Edge items ($395, $150, $100) match this schedule exactly.", "Analysis")
put(sm, "A8", "Small-amount threshold", f_bold); put(sm, "B8", 250, f_input, CUR, fill_in)
sm["B8"].comment = Comment("ASSUMPTION: user said fee checks are 'sometimes other small amounts' without a figure. Items at or below this threshold that are not an exact fee amount are classified 'Small amount - likely fee, review'. Change it to see the effect.", "Analysis")

labels = ["Entity", "Default governing state = LLC formation state (owner-address state overrides per item in Detail col U)", "Prior-month accrued (unbooked deposits)", "Current-month accrued (Sept/Oct 2026 deposits)", "Total accrued per workbook", "Items (count)",
          "0-1 yrs", "1-2 yrs", "2-3 yrs", "3-5 yrs", "5+ yrs",
          "$ at fee amounts (presumed company revenue)", "$ small amounts (likely fee, review)", "$ memo-flagged (not ours / unknown / hold)", "$ larger amounts (likely includes agent commission)",
          "Short period (yrs)", "Long period (yrs)",
          "Dormant today - SHORT test - ALL items", "Dormant today - LONG test - ALL items",
          "Dormant today - SHORT test - likely-commission + memo-flagged only", "Dormant today - LONG test - likely-commission + memo-flagged only",
          "Dormant at next report cutoff - SHORT - likely-commission + memo-flagged", "Dormant at next report cutoff - LONG - likely-commission + memo-flagged",
          "Likely test given independent-contractor agents", "Dormant today - LIKELY test - likely-commission + memo-flagged", "Dormant at next report cutoff - LIKELY test - likely-commission + memo-flagged",
          "Oldest item", "Office state", "Entity note (legal entity, formation, office)"]
hdr(sm, 10, labels, [14, 13, 16, 16, 16, 8, 13, 13, 13, 13, 11, 15, 15, 15, 16, 8, 8, 15, 15, 17, 17, 17, 17, 10, 17, 17, 12, 8, 70])
sm.row_dimensions[10].height = 70
R0 = 11
for i, (ent, st, office, note) in enumerate(ENTITIES):
    r = R0 + i
    put(sm, f"A{r}", ent, f_bold)
    put(sm, f"B{r}", st, f_input, fill=fill_in)
    put(sm, f"AB{r}", office, f_input)
    put(sm, f"C{r}", f'=SUMIFS(Detail!$H:$H,Detail!$A:$A,$A{r})', f_link, CUR)
    put(sm, f"D{r}", f"=SUMIFS('Current Month'!$F:$F,'Current Month'!$A:$A,$A{r})", f_link, CUR)
    put(sm, f"E{r}", f"=C{r}+D{r}", fmt=CUR)
    put(sm, f"F{r}", f'=COUNTIFS(Detail!$A:$A,$A{r},Detail!$H:$H,"<>")')
    for j, bk in enumerate(["0-1 yrs", "1-2 yrs", "2-3 yrs", "3-5 yrs", "5+ yrs"]):
        put(sm, f"{L(7+j)}{r}", f'=SUMIFS(Detail!$H:$H,Detail!$A:$A,$A{r},Detail!$L:$L,"{bk}")', f_link, CUR)
    put(sm, f"L{r}", f'=SUMIFS(Detail!$H:$H,Detail!$A:$A,$A{r},Detail!$N:$N,"Fee amount*")', f_link, CUR)
    put(sm, f"M{r}", f'=SUMIFS(Detail!$H:$H,Detail!$A:$A,$A{r},Detail!$N:$N,"Small amount*")', f_link, CUR)
    put(sm, f"N{r}", f'=SUMIFS(Detail!$H:$H,Detail!$A:$A,$A{r},Detail!$N:$N,"Memo flag*")', f_link, CUR)
    put(sm, f"O{r}", f'=SUMIFS(Detail!$H:$H,Detail!$A:$A,$A{r},Detail!$N:$N,"Larger amount*")', f_link, CUR)
    put(sm, f"P{r}", f"=IFERROR(INDEX('State Rules'!$C${SR_FIRST}:$C${SR_LAST},MATCH($B{r},'State Rules'!$A${SR_FIRST}:$A${SR_LAST},0)),\"\")", f_link)
    put(sm, f"Q{r}", f"=IFERROR(INDEX('State Rules'!$E${SR_FIRST}:$E${SR_LAST},MATCH($B{r},'State Rules'!$A${SR_FIRST}:$A${SR_LAST},0)),\"\")", f_link)
    put(sm, f"R{r}", f'=SUMIFS(Detail!$H:$H,Detail!$A:$A,$A{r},Detail!$P:$P,"<="&$B$3)', f_link, CUR)
    put(sm, f"S{r}", f'=SUMIFS(Detail!$H:$H,Detail!$A:$A,$A{r},Detail!$Q:$Q,"<="&$B$3)', f_link, CUR)
    put(sm, f"T{r}", f'=SUMIFS(Detail!$AA:$AA,Detail!$A:$A,$A{r},Detail!$P:$P,"<="&$B$3)', f_link, CUR)
    put(sm, f"U{r}", f'=SUMIFS(Detail!$AA:$AA,Detail!$A:$A,$A{r},Detail!$Q:$Q,"<="&$B$3)', f_link, CUR)
    put(sm, f"V{r}", f'=SUMIFS(Detail!$AA:$AA,Detail!$A:$A,$A{r},Detail!$R:$R,"Yes")', f_link, CUR)
    put(sm, f"W{r}", f'=SUMIFS(Detail!$AA:$AA,Detail!$A:$A,$A{r},Detail!$S:$S,"Yes")', f_link, CUR)
    put(sm, f"X{r}", f"=IFERROR(INDEX('State Rules'!$M${SR_FIRST}:$M${SR_LAST},MATCH($B{r},'State Rules'!$A${SR_FIRST}:$A${SR_LAST},0)),\"\")", f_link)
    put(sm, f"Y{r}", f'=SUMIFS(Detail!$AA:$AA,Detail!$A:$A,$A{r},Detail!$AC:$AC,"<="&$B$3)', f_link, CUR)
    put(sm, f"Z{r}", f'=SUMIFS(Detail!$AA:$AA,Detail!$A:$A,$A{r},Detail!$AD:$AD,"Yes")', f_link, CUR)
    put(sm, f"AA{r}", f'=IF(F{r}=0,"",_xlfn.MINIFS(Detail!$E$7:$E$400,Detail!$A$7:$A$400,$A{r}))', f_link, DATE)
    put(sm, f"AC{r}", note, wrap=True)
    for c in range(1, 30): sm.cell(r, c).border = box
RT = R0 + len(ENTITIES)
put(sm, f"A{RT}", "TOTAL", f_bold)
for col in ["C","D","E","G","H","I","J","K","L","M","N","O","R","S","T","U","V","W","Y","Z"]:
    put(sm, f"{col}{RT}", f"=SUM({col}{R0}:{col}{RT-1})", f_bold, CUR)
put(sm, f"F{RT}", f"=SUM(F{R0}:F{RT-1})", f_bold)
for c in range(1, 30): sm.cell(RT, c).border = box; sm.cell(RT, c).fill = fill_sub
sm.freeze_panes = "C11"

n = RT + 2
notes = [
    ("How to read this", f_bold),
    ("Classification (Detail col N) uses the fee amounts in rows 4-7 and the threshold in B8: an exact fee amount = presumed company revenue (keep; book to revenue, not unclaimed property). At or under the threshold = likely fee, review. Memo-flagged = memo says not ours / unknown agent / duplicate / legal hold (owed back to the payer or needs resolution). Everything else = larger amount that likely includes an agent's commission split - this is the unclaimed-property population.", f_norm),
    ("Columns R-S apply the dormancy tests to ALL unbooked items (upper bound). Columns T-W apply them only to the likely-commission and memo-flagged items. Even the T-W figures are upper bounds: within a larger check, only the agent's share (typically the check less the company's fee / split) is reportable; the company's share is revenue.", f_norm),
    ("SHORT test = the state's period for wages / compensation for personal services. LONG test = the general catch-all period. Column X picks the likely test for each entity's default governing state now that agents are confirmed to be independent contractors (Texas and Florida: LONG; Illinois, Pennsylvania, Alabama: SHORT), and columns Y-Z apply it item by item, so an item whose owner state has been filled in uses its own state's test. See State Rules cols M-N for the reasoning; it is a legal judgment for counsel to confirm.", f_norm),
    ("Unknown agent = second priority rule. Per the Entity Management Database, URE Dallas, URE Houston, Quick Close Properties (TUR), URE Chicago and URE Philadelphia are all TEXAS LLCs; Gallery (RaySon Partners) is Florida; Leading Edge Realty is Alabama. Column B holds that formation state as each entity's DEFAULT governing state, so unknown-agent items at Chicago and Philly are tested under TEXAS rules (3-year general period for contractor commissions), not Illinois or Pennsylvania. When an agent's last-known address state is entered in Detail col U, that item switches to the address state's rules (first priority rule). Office state is shown in col AB for reference.", f_norm),
    ("Current-month accrued (column D) is all September/October 2026 deposits in normal matching workflow. It is included so the total ties to the cash-requirements Summary, but none of it is near any dormancy period.", f_norm),
    ("Nothing here is legal advice. Dormancy periods, report dates and due-diligence rules on the State Rules tab were sourced from statute summaries and state guidance found by web search; the statute pages themselves could not be opened from this session. Have counsel or your unclaimed-property advisor confirm before filing.", f_norm),
]
for i, (t, f) in enumerate(notes):
    put(sm, f"A{n+i}", t, f, wrap=True); sm.merge_cells(f"A{n+i}:P{n+i}"); sm.row_dimensions[n+i].height = 16 if f is f_bold else 52

# ================================================================ Detail
dt = wb.create_sheet("Detail", 1)
put(dt, "A1", "Unbooked deposited checks - item detail (prior-month accrued commissions block of each entity tab)", f_title)
put(dt, "A2", "Columns A-J are copied from the source workbook. K-S and Z-AA are formulas. T-Y (yellow) are for you to fill in from Paperless Pipeline / Sage: once the agent share is entered, column Z shows the dollar amount that is both dormant (long test) and owed to someone else. Column AA is the heuristic stand-in until then.", wrap=True)
dt.merge_cells("A2:AD2"); dt.row_dimensions[2].height = 32
put(dt, "A3", "Classification (col N) is driven by Summary rows 4-7 (fee amounts, with Philly-only, TUR-only and Leading Edge-only lists) and Summary B8 (small-amount threshold). Col M flags memos that say not ours / unknown agent / duplicate / legal hold / '?'. Classifications are hints, not conclusions.", wrap=True)
dt.merge_cells("A3:AD3"); dt.row_dimensions[3].height = 32
put(dt, "A4", "As-of date (linked)"); put(dt, "B4", "=Summary!$B$3", f_link, DATE)
labels = ["Entity", "Governing state (owner state if filled in col U, else LLC formation state)", "Bank / block", "Source row", "Deposit date", "Check #", "Payer / check writer", "Amount", "Memo (source)", "Booked in Sage?",
          "Age (yrs)", "Age bucket", "Memo flag?", "Classification",
          "Short period (yrs)", "Dormant date - SHORT", "Dormant date - LONG", "Dormant at next cutoff - SHORT?", "Dormant at next cutoff - LONG?",
          "Agent / owner name (fill in)", "Owner last-known state (fill in)", "Brokerage fee / E&O share $ (fill in)", "Agent share $ (fill in)", "Refund to payer $ (fill in)", "Resolution / status (fill in)",
          "Reportable if dormant (LONG test) = agent + refund share (from fill-in)", "Heuristic reportable population $ (larger-amount + memo-flagged items)",
          "Likely test for this item's governing state", "Dormant date - LIKELY test", "Dormant at next cutoff - LIKELY test?"]
hdr(dt, 6, labels, [13, 11, 13, 8, 11, 14, 32, 12, 55, 8, 8, 9, 8, 38, 8, 12, 12, 11, 11, 22, 10, 13, 13, 13, 28, 16, 16, 10, 12, 11])
dt.row_dimensions[6].height = 56
FLAGS = ["not for this", "not in pp", "not sure", "unsure", "duplicate", "couldn't locate", "cant locate", "can't locate", "fraud", "on hold", "law enforcement", "unidentified", "possibly", "believe this transaction", "shortage", "supposed to be", "review check stub", "no memo", "?"]
def memo_flag(r):
    memo = (r["memo"] or "").lower()
    return "Yes" if any(k in memo for k in FLAGS) else "No"
ent_order = {e[0]: i for i, e in enumerate(ENTITIES)}
rows.sort(key=lambda r: (ent_order.get(r["entity"], 99), r["d"] or datetime.date(1900, 1, 1)))
D0 = 7
for i, r in enumerate(rows):
    rr = D0 + i
    put(dt, f"A{rr}", r["entity"])
    put(dt, f"B{rr}", f"=IF(U{rr}<>\"\",U{rr},IFERROR(INDEX(Summary!$B${R0}:$B${RT-1},MATCH($A{rr},Summary!$A${R0}:$A${RT-1},0)),\"\"))", f_link)
    put(dt, f"C{rr}", r["bank"]); put(dt, f"D{rr}", int(r["row"]))
    put(dt, f"E{rr}", r["d"], fmt=DATE); put(dt, f"F{rr}", r["check"])
    put(dt, f"G{rr}", r["payer"]); put(dt, f"H{rr}", r["amt"], fmt=CUR)
    put(dt, f"I{rr}", r["memo"], wrap=False); put(dt, f"J{rr}", "No")
    put(dt, f"K{rr}", f"=IF(E{rr}=\"\",\"\",($B$4-E{rr})/365.25)", fmt="0.00")
    put(dt, f"L{rr}", f'=IF(K{rr}="","no date",IF(K{rr}<1,"0-1 yrs",IF(K{rr}<2,"1-2 yrs",IF(K{rr}<3,"2-3 yrs",IF(K{rr}<5,"3-5 yrs","5+ yrs")))))')
    mf = memo_flag(r)
    put(dt, f"M{rr}", mf, fill=fill_warn if mf == "Yes" else None)
    put(dt, f"N{rr}", f'=IF(M{rr}="Yes","Memo flag - not ours / unknown / hold: resolve",IF(OR(COUNTIF(Summary!$C$4:$R$4,H{rr})>0,AND(A{rr}="Philly",COUNTIF(Summary!$C$5:$R$5,H{rr})>0),AND(A{rr}="TUR",COUNTIF(Summary!$C$6:$R$6,H{rr})>0),AND(A{rr}="Leading Edge",COUNTIF(Summary!$C$7:$R$7,H{rr})>0)),"Fee amount - presumed company revenue",IF(H{rr}<=Summary!$B$8,"Small amount - likely fee, review","Larger amount - likely includes agent commission")))', f_link)
    put(dt, f"O{rr}", f"=IFERROR(INDEX('State Rules'!$C${SR_FIRST}:$C${SR_LAST},MATCH($B{rr},'State Rules'!$A${SR_FIRST}:$A${SR_LAST},0)),\"\")", f_link)
    put(dt, f"P{rr}", f'=IF(OR(O{rr}="",E{rr}=""),"",EDATE(E{rr},12*O{rr}))', fmt=DATE)
    put(dt, f"Q{rr}", f"=IF(E{rr}=\"\",\"\",IFERROR(EDATE(E{rr},12*INDEX('State Rules'!$E${SR_FIRST}:$E${SR_LAST},MATCH($B{rr},'State Rules'!$A${SR_FIRST}:$A${SR_LAST},0))),\"\"))", fmt=DATE)
    nxt = f"INDEX('State Rules'!$I${SR_FIRST}:$I${SR_LAST},MATCH($B{rr},'State Rules'!$A${SR_FIRST}:$A${SR_LAST},0))"
    put(dt, f"R{rr}", f'=IF(P{rr}="","",IFERROR(IF(P{rr}<={nxt},"Yes","No"),""))')
    put(dt, f"S{rr}", f'=IF(Q{rr}="","",IFERROR(IF(Q{rr}<={nxt},"Yes","No"),""))')
    for col in "TUVWXY": put(dt, f"{col}{rr}", None, f_input, fill=fill_in)
    for col in "VWX": dt[f"{col}{rr}"].number_format = CUR
    put(dt, f"Z{rr}", f'=IF(Q{rr}="","",IF(Q{rr}<=$B$4,N(W{rr})+N(X{rr}),0))', fmt=CUR)
    put(dt, f"AA{rr}", f'=IF(OR(LEFT(N{rr},6)="Larger",LEFT(N{rr},4)="Memo"),H{rr},0)', fmt=CUR)
    put(dt, f"AB{rr}", f"=IFERROR(INDEX('State Rules'!$M${SR_FIRST}:$M${SR_LAST},MATCH($B{rr},'State Rules'!$A${SR_FIRST}:$A${SR_LAST},0)),\"\")", f_link)
    put(dt, f"AC{rr}", f'=IF(AB{rr}="SHORT",P{rr},IF(AB{rr}="LONG",Q{rr},""))', fmt=DATE)
    put(dt, f"AD{rr}", f'=IF(AB{rr}="SHORT",R{rr},IF(AB{rr}="LONG",S{rr},""))')
    for c in range(1, 31): dt.cell(rr, c).border = box
DT_LAST = D0 + len(rows) - 1
put(dt, f"A{DT_LAST+1}", "TOTAL", f_bold)
put(dt, f"H{DT_LAST+1}", f"=SUM(H{D0}:H{DT_LAST})", f_bold, CUR)
for col in ["V","W","X","Z","AA"]: put(dt, f"{col}{DT_LAST+1}", f"=SUM({col}{D0}:{col}{DT_LAST})", f_bold, CUR)
dt.freeze_panes = "E7"
dt.auto_filter.ref = f"A6:AD{DT_LAST}"

# ================================================================ Current Month
cm = wb.create_sheet("Current Month", 2)
put(cm, "A1", "Current-month unbooked deposits ('Deposited checks' block, no 'x' in Sage column) - all September/October 2026", f_title)
put(cm, "A2", "Listed only so the Summary ties to the cash-requirements workbook. These are in the normal matching cycle and are not an unclaimed-property concern today.", wrap=True)
cm.merge_cells("A2:G2")
hdr(cm, 4, ["Entity", "Bank / block", "Source row", "Deposit date", "Check #", "Amount"], [14, 14, 9, 12, 16, 14])
for i, r in enumerate(cur):
    rr = 5 + i
    put(cm, f"A{rr}", r["entity"]); put(cm, f"B{rr}", r["bank"]); put(cm, f"C{rr}", int(r["row"]))
    put(cm, f"D{rr}", datetime.date.fromisoformat(r["date"]) if r["date"] else None, fmt=DATE)
    put(cm, f"E{rr}", r["check"]); put(cm, f"F{rr}", float(r["amount"]), fmt=CUR)
    for c in range(1, 7): cm.cell(rr, c).border = box
put(cm, f"A{5+len(cur)}", "TOTAL", f_bold); put(cm, f"F{5+len(cur)}", f"=SUM(F5:F{4+len(cur)})", f_bold, CUR)

# ================================================================ Gaps & Next Steps
gp = wb.create_sheet("Gaps & Next Steps", 3)
put(gp, "A1", "What is missing, and what to do next", f_title)
hdr(gp, 3, ["#", "Item", "Why it matters", "Who / where"], [4, 55, 80, 30])
gaps = [
    ("DATA GAP: the fee vs. agent split of every check is unknown", "The source workbook has E&O / FEE / Commission columns but they are empty for all 181 items. Per the user, fee checks are typically $644, $1,044, $69 or other small amounts, so items at those amounts are presumed company revenue and items at or under the small-amount threshold are treated as likely fees. Larger checks likely include an agent's commission split, but how much of each is the agent's is still unknown; only that share (plus anything owed back to a payer) can ever be unclaimed property. Until the split is known, every dollar figure on the Summary is an upper bound, not an amount to remit.", "Match each check to its Paperless Pipeline transaction; fill Detail cols U-W."),
    ("DATA GAP: owner identity and last-known address", "The first priority rule sends property to the state of the owner's last-known address. User confirmed that for many older checks the agent is not known. Where the agent can be identified, record the name and state in Detail cols T-U. Where it cannot, the item falls to the second priority rule (state of formation of the LLC).", "Agent master file / Paperless Pipeline; fill Detail cols T-U."),
    ("RESOLVED, with a twist: state of formation", "Entity Management Database (10/6/2026): URE Dallas, URE Houston, Quick Close Properties/TUR, URE Chicago and URE Philadelphia are all Texas LLCs; URE Chicago and URE Philadelphia are merely registered as foreign LLCs in Illinois and Pennsylvania. Gallery (RaySon Partners) is Florida; Leading Edge Realty is Alabama. Consequence: every unknown-agent item at Chicago and Philly is reportable to TEXAS, under Texas's 3-year general period for contractor commissions, not to Illinois (1 year) or Pennsylvania (2 years). Only items where the company's records show an agent address in IL or PA go to those states. The workbook now defaults each entity to its formation state and switches an item to the owner's state once Detail col U is filled in.", "Done; counsel to confirm the priority-rule application."),
    ("PRACTICAL EFFECT: one Texas filing could cover most of it", "With Texas as the default state for five of the seven entities, a single Texas voluntary disclosure (one per holder entity, or coordinated) would address the bulk of the aged population. Pennsylvania and Illinois filings would be limited to items with known PA/IL agent addresses.", "Counsel / Comptroller."),
    ("RESOLVED: office-to-state mapping", "User confirmed 10/6/2026: Dallas, Houston, TUR = Texas; Chicago = Illinois; Philly = Pennsylvania; Leading Edge = Alabama; Gallery = Florida. Office state is shown on Summary col AB; it is NOT the governing state for unknown-owner items (see the formation-state row).", "Done."),
    ("CONFIRM: payable date vs. deposit date", "Dormancy runs from when the amount became payable to the owner, not from when the check was deposited. Deposit date is used as a proxy throughout. If agents are paid on closing, the closing date is the better trigger and could be earlier.", "Paperless Pipeline closing dates."),
    ("RESOLVED: agents are independent contractors", "User confirmed 10/6/2026. Texas' 1-year rule applies to 'wages' as defined in Labor Code 61.001, which excludes independent contractors, so the 3-year general rule is the likely Texas test. Illinois and Alabama reach 'other compensation for personal services' and Pennsylvania lists commissions, so their shorter periods likely apply. Florida's wage rule is not shown to reach contractors, so 5 years is likely. Counsel should confirm each.", "Confirm with counsel."),
    ("RESOLVED: no prior write-offs", "User confirmed nothing has been done with the unapplied deposits so far, so the population in this workbook is the full population and there is no prior 'private escheat' to unwind.", "Done."),
    ("ICA CLAUSES THAT REDUCE WHAT IS OWED TO AGENTS - legal question", "Every ICA lets the Broker hold or apply commissions against amounts the agent owes (dues, fees, E&O deductibles, indemnity) and conditions post-termination payouts on dues being paid in full. The Philly ICA goes further: an incomplete file 30 days past closing is paid 50/50 and 60 days past closing is 'forfeited'; a license not reactivated within 30 days means 'brokerage will keep 100% of commission'. TUR's agreement forfeits accrued commissions if the agent collects in their own name or breaches Section 5. A documented offset reduces the agent's share dollar for dollar. A contractual forfeiture is less certain: Alabama's act (and, per secondary sources, most states' acts) says a contractual limitation on the owner's right does not prevent property from being presumed abandoned. Whether Texas, Pennsylvania, Illinois or Florida would honor these clauses against an escheat claim is for counsel. Do not treat a forfeiture as the company's money until counsel says so AND the triggering facts (closing date, file status at day 60, termination date) are documented per item. See the ICA Terms tab.", "Counsel; Accounting to document per item."),
    ("WHEN THE AGENT'S SHARE BECOMES PAYABLE (dormancy trigger)", "Houston's TXR-2301 form is explicit: the associate's fee is payable when Broker receives Broker's fees (para. 16.C), so the deposit date is the trigger for Houston. Dallas, Chicago and Philly ICAs pay 'promptly after receipt and processing' but 'provided that Broker has received a complete File'. A condition that depends on the agent or the broker completing paperwork probably does not defer dormancy under anti-limitation provisions, but it is an argument counsel may want to assess. TUR agents are normally paid by the title company at closing; TUR's deposited checks are usually the broker's portion, lease or new-home commissions, so TUR's larger items need the Disbursement Authorization form to find the agent share.", "Counsel; Accounting."),
    ("DATA: the split exists on the file", "TUR's Disbursement Authorization form records TUR / Agent / Mentor amounts per transaction. Dallas, Chicago and Philly fees are on the Closing Disclosure. Pull these for each larger item to fill Detail cols V-X; that converts the upper-bound figures into the real reportable amount.", "Accounting / Paperless Pipeline."),
    ("RESOLVED: fee schedules for every office", "The ICA Fee Structure Cheat Sheet (current) and JV ICA Transaction Fee Summary (older $495/$895 schedule) supplied 10/6/2026 cover Chicago, Dallas/Frisco, Houston/TUR, Philly, DC, TUR legacy and Leading Edge. All amounts are in Summary rows 4-7.", "Done."),
    ("NOTE: Houston items at $495", "Two Houston items at exactly $495 (10/9/2024 Fidelity National Title 'Commissions'; 7/21/2026 Patten Title) are classified as fees because $495 was the sale transaction fee under the older schedule. If Houston had moved to $595 by those dates, treat them as commission instead.", "Accounting."),
    ("ACTION: work the oldest items first", "Houston has four items from 2022 and Dallas/TUR/Philly have 2023 items. Under the Texas 3-year test, every Texas item deposited before 3/1/2023 would already have been due on the July 1, 2026 report if it was owed to someone else; the Detail tab's 'Dormant date - LONG' column identifies them.", "Accounting."),
    ("ACTION: run due-diligence letters before any filing", "Every state in scope requires a written notice to the owner before remittance (TX: >$250, 60+ days before delivery; PA/IL/FL/AL/MO: $50 threshold per sources). A response from the owner stops the clock and lets you pay them directly instead of the state.", "Accounting; template letters."),
    ("ACTION: consider a Voluntary Disclosure Agreement for Texas", "Three Texas entities hold about 85% of the aged balance. Texas waives penalty and interest under a VDA but it must be requested before any Comptroller inquiry. Florida has a similar program. Confirm eligibility with the Comptroller or counsel.", "Counsel / Comptroller."),
    ("ACTION: items that are not the company's money", "Memo-flagged items (e.g. 'not for this address', 'Not in PP', 'duplicate?', 'Dallas transaction' deposited in Houston) are owed back to the payer, not to an agent. Return them to the payer now; a refund resolves the item and removes it from the escheat population.", "Accounting."),
    ("ACTION: legal hold item", "Houston row 27 ($1,500, 7/18/2023, memo: on hold per law enforcement). Keep segregated; document the hold; do not escheat or release without counsel.", "Counsel."),
    ("NOT VERIFIED in this session", "Statute and state-treasury web pages were blocked by this session's network policy, so every rule on the State Rules tab rests on search-engine summaries of those pages and on secondary sources (Sovos, UPPO, law-firm alerts). Specific items I could not confirm: Illinois business-holder report deadline; PA December 31 cutoff and VDA terms; FL and AL due-diligence thresholds; MO VDA. Treat the State Rules tab as a starting point for counsel, not a filing position.", "Counsel / unclaimed-property advisor."),
]
for i, (a, b, c) in enumerate(gaps, 4):
    put(gp, f"A{i}", i - 3); put(gp, f"B{i}", a, f_bold, wrap=True); put(gp, f"C{i}", b, wrap=True); put(gp, f"D{i}", c, wrap=True)
    gp.row_dimensions[i].height = 75
    for cc in range(1, 5): gp.cell(i, cc).border = box

# ================================================================ ICA Terms
ic = wb.create_sheet("ICA Terms", 4)
put(ic, "A1", "Independent Contractor Agreement terms relevant to accrued commissions and unclaimed property", f_title)
put(ic, "A2", "Five agreements provided 10/6/2026. Quotations are from the documents; 'relevance' is analysis, not legal advice. Dallas and Chicago PDFs are scans and were read visually; the other three were text-extracted.", wrap=True)
ic.merge_cells("A2:F2"); ic.row_dimensions[2].height = 30
hdr(ic, 4, ["Entity", "Legal entity named in the document", "Document", "Clause / location", "What it says", "Relevance to accrued commissions / unclaimed property"], [11, 30, 34, 26, 80, 70])
ICA = [
 ("Dallas", "URE Dallas LLC dba United Real Estate (Nick Bristow, General Manager, Broker)", "Independent Contractor Agreement for Dallas/Frisco, DocuSign CBD8CDF3..., revised 4/18/2023", "Commissions and Fees (p.2)", "Commissions, compensation, bonuses and fees are paid to Broker. Contractor Commissions 'will be promptly paid to by Broker after receipt & processing of such Commissions, less any amounts owed to Broker, & provided that Broker has received a complete File.'", "Agent's share arises on receipt, net of amounts the agent owes, conditioned on a complete file. Supports using deposit date as the dormancy trigger; the complete-file condition is a counsel question (see Gaps)."),
 ("Dallas", "", "same", "Broker Fees (p.3)", "$595 transaction fee per side for residential sales up to $649,999; $995 at $650,000 and above; commercial 10%; other events 12% (min $75, max $995 per side); Broker-generated lead fee 40%; E&O Fee $49 per transaction; 'All sales related fees due Broker must be included on the transaction's Closing Disclosure or HUD-1 Statement & paid through the close of escrow.' Monthly $65 + $4 convenience fee by credit card. Annual cap: after 24 transactions, $50 fee + E&O per transaction.", "Confirms $644 = $595 + $49 and $1,044 = $995 + $49 arrive from title as company fee revenue, and $69 = monthly dues. These checks are the company's money, not unclaimed property."),
 ("Dallas", "", "same", "Errors & Omissions Insurance (p.4)", "Broker may withhold from Contractor's commissions payable an amount adequate to satisfy uncovered claims, placed in a Claims and Disputes Retention Account pending settlement.", "A documented withholding for a claim is a legitimate reduction of the agent's share; keep the claim file."),
 ("Dallas", "", "same", "Protection of Contractor's Listings and Contracts; Transaction Defined; Termination (p.10)", "Commissions earned and paid on sales contracts completed before termination 'will be disbursed to Contractor in compliance with the commission plan in effect... conditioned on Contractor having any and all dues, fees and expenses owing to Broker paid in full.' 'Broker has the right to hold and/or apply any commissions owing to Contractor, as may be necessary to pay for or secure any obligations of Contractor.' Transactions not closed before termination are paid at a 70/30 split.", "Offset rights: amounts an agent owes the company reduce the agent's reportable share dollar for dollar if documented. No clause forfeits earned commissions outright."),
 ("Dallas", "", "same", "Contractor Information (p.10); Agent Information sheet (p.1); Controlling Law (p.11)", "Contractor must immediately update mailing address, phone and email. The Agent Information sheet collects home address, state and zip. Controlling law: Texas.", "The company should hold a last-known address for every identified agent, which puts those items under the first priority rule (agent's address state, usually Texas)."),
 ("Chicago", "URE Chicago LLC dba United Real Estate - Chicago (Richard Williams, Designated Managing Broker)", "Independent Contractor Agreement, DocuSign A769C223...", "Commissions and Fees; Broker Fees (pp.2-3)", "Same commission language as Dallas (promptly paid after receipt and processing, less amounts owed, provided Broker has a complete File). Fees: $595 per side up to $650,000; $995 at $650,001+; commercial greater of 12% or $595; other events 12% (min $75, max $995); lead fee 40%; E&O $49 per side; monthly $69; annual cap after 24 transactions.", "Same fee structure as Dallas: $644 / $1,044 / $69 are company revenue."),
 ("Chicago", "", "same", "E&O; Protection of Listings; Transaction Defined; Termination; Controlling Law (pp.4, 9-11)", "Same E&O withholding, dues-paid-in-full condition and hold/apply rights as Dallas. No 70/30 post-termination line. Contractor must update mailing address. Controlling law: Texas.", "Offset rights as Dallas. The Texas choice-of-law clause does not change which state takes custody of unclaimed property; the priority rules do."),
 ("Houston", "URE Houston LLC d/b/a United Real Estate (Carol Drake, Managing Broker)", "Texas REALTORS form TXR-2301 (07-08-22), DocuSign 3DE311AB...", "Para. 12.B Receipt of Brokerage Fees", "Associate must deliver any compensation for brokerage services received from any client, escrow agent, title company or other person to Broker for disbursement in accordance with the agreement.", "All commission checks flow through the Broker, which is why they land in the deposited-checks block."),
 ("Houston", "", "same", "Para. 16.A and 16.C Associate's Fees", "'All fees and compensation that Broker or Associate earn for providing brokerage services... are payable to and belong to Broker.' 'Associate's fees under this agreement are earned at the time Broker's fees are earned... Associate's fees under this agreement are payable when Broker receives Broker's fees under the applicable agreements for brokerage services, unless the fees are subject to arbitration, litigation, or a court order.'", "Clearest trigger in the five documents: the agent's share is PAYABLE WHEN THE BROKER RECEIVES THE CHECK. For Houston the deposit date is the dormancy start date with little room for argument."),
 ("Houston", "", "same", "Para. 16.D Disputes; Para. 17.B Special Expenses", "Disputed fees between associates are held in trust pending resolution. Special expenses: $65 monthly fee by auto-debit; each transaction carries a $49 E&O charge plus a transaction fee per the attached fee schedule, deducted from gross fees before calculating Associate's fees.", "Fee schedule itself was not in the copy provided. The $49 E&O and the $644 items imply the same $595 + $49 pattern as Dallas; confirm."),
 ("Houston", "", "same", "Para. 21 Termination; Para. 24.E Controlling Law; signature block", "Either party may terminate on written notice. Any fee unpaid at termination is paid per the fee schedule. No forfeiture clause. Controlling law: Texas. Signature block collects the associate's home address.", "Termination does not extinguish the agent's right to unpaid fees. Last-known address should be on file."),
 ("TUR", "Quick-Close Properties, LLC dba Texas United Realty (TREC license 599460)", "TUR Legacy Sign Up Docs v1 (TREC sponsorship form, Agent Agreement, Commission Policy, Policies, DA Checklist, Disbursement Authorization form)", "Policies item 11 Commission Payment", "'Licensee acknowledges he/she will be paid at closing and funding from the title company for all commissions earned except... 1) Lease commissions are paid through the broker directly to the licensee three business days after broker has received and deposited the commission check. 2) New Home commissions... three business days after received and deposited, broker will cut agent's check.'", "TUR agents are normally paid by title at closing, so checks TUR deposits are usually the broker's portion, a lease commission or a new-home commission. For lease and new-home items the agent's share is payable three business days after deposit, which fixes the dormancy trigger."),
 ("TUR", "", "same", "Commission Policy of Texas United Realty", "Plan 1 New Agent Mentorship: 70% split. Plan 2 Agent Coaching: 85% split less transaction fee. Plan 3 Transaction Fee Plan: $150 (sale up to $10K commission), $250 (up to $20K), $350 (up to $30K), $450 (up to $40K), $60 residential lease. All plans $179 annual fee. Commercial 70/30 or 95/5.", "Explains TUR's recurring $150 / $250 / $60 deposits as fee checks and gives the split to apply to larger items once the plan on the DA form is known."),
 ("TUR", "", "same", "Disbursement Authorization form", "Per-transaction form showing total commission and the amounts to Texas United Realty, Agent and Mentor, approved by Broker/Owner, with a request that title mail the HUD and broker's portion to TUR.", "The agent/company split for every TUR transaction is on this form. Pull it for each larger item to replace the upper-bound figures."),
 ("TUR", "", "same", "Agent Agreement paras. 1.K, 3.B, 5.F, 8.C", "1.K: an agent who collects a fee in their own name 'will forfeit any accrued but uncollected commissions'. 3.B: agent assigns commissions to TUR to secure indemnification. 5.F: breach of Section 5 (confidentiality / non-solicitation) means the agent 'shall lose all rights to any further commissions'. 8.C: amounts owed at or after termination may be deducted from commissions due.", "Offset (8.C, 3.B) is a documented reduction. Forfeiture (1.K, 5.F) is a legal question: anti-limitation provisions in unclaimed property acts may override a contractual forfeiture. Do not rely on it without counsel and a documented breach."),
 ("TUR", "", "same", "Policies item 6 Record Keeping", "All transaction documentation kept for four years from closing or termination of contract.", "Four years is shorter than unclaimed property record-retention expectations (commonly 10 years after a report) and audit look-backs. Preserve DA forms and HUDs for the aged items regardless."),
 ("Philly", "'United Real Estate' (no LLC name in the document); Pennsylvania Real Estate Commission licensee", "Independent Contractor Agreement, Updated 08.2023, 11 pages", "Para. 3 Commissions and Fees", "'100% of any and all such commissions will be promptly paid to Contractor by Broker after receipt and processing, less any amounts owing to Broker. Payment of any and all commissions is subject to Broker receiving a complete sales file no less than five (5) business days prior to closing.'", "Agent's share arises on receipt net of amounts owed; complete-file condition as Dallas."),
 ("Philly", "", "same", "Para. 3 fee schedule", "$595 transaction fee per side for sales up to $799,999.99; $995 at $800,000+; 'charged to the Buyer or Seller and paid to Broker'; agent responsible if the buyer/seller does not pay. E&O $49 per transaction; enhanced E&O $110 on personal transactions. Leases 12% (min $125, or $595/$995 whichever is less) + $49. Referrals 10% (min $100) + $49. Commercial 10% + $49. Monthly $65 + $4 card fee, $25 late fee. Lead fee 25-50%; mentor fee 30% on first four transactions. 'All fees due Broker must be included on the transaction's HUD-1/ALTA/Closing Disclosure statement and paid through the close of escrow.'", "Confirms fee checks arrive from settlement. User adds $495 and $295 as Philly fee amounts (not in this 08.2023 schedule; presumably an earlier schedule)."),
 ("Philly", "", "same", "Para. 21 Protection of Listings; Para. 22 Transaction Defined; Para. 23 Termination", "Post-termination commissions disbursed per the plan in effect, 'conditioned on Contractor having any and all dues, fees and expenses owing to Broker paid in full.' Broker may 'hold and/or apply any commissions owing to Contractor' to secure obligations. On termination, executed contracts split 70/30. 'If the Contractor's license expires or goes inactive, anything that is currently under contract brokerage split will be fifty/fifty. Contractor will have 30 days to bring their license back to active status. If Contractor fails to bring license back to active status within the 30 days, brokerage will keep 100% of commission.'", "Offset rights as elsewhere, plus an express 100% retention clause for lapsed licenses. Whether Pennsylvania unclaimed property law honors a contractual retention against an escheat claim is for counsel."),
 ("Philly", "", "same", "Para. 31 Fines and Penalties", "'3. Any incomplete file(s) 30 days past closing will be paid at a 50/50 split. 4. Any incomplete file(s) 60 days past closing will be forfeited and grounds for termination.'", "The most consequential clause for Philly's aged items: if a file was never completed, the ICA says the agent's share was forfeited at day 60 after closing. If counsel concludes Pennsylvania (or the agent's address state) honors this, those amounts are company revenue, not unclaimed property. Document the closing date and file status for each item before relying on it; anti-limitation provisions may override."),
 ("Philly", "", "same", "Para. 26 Contractor Information; Para. 28 Controlling Law; Para. 30 Paperless Pipeline", "Contractor must update mailing address. Controlling law: Pennsylvania. All documents must be uploaded to Paperless Pipeline within 72 hours.", "Last-known address should be on file. Paperless Pipeline should show whether each aged file was ever completed."),
]
ICA += [
 ("All", "n/a", "ICA Fee Structure Cheat Sheet (xlsx; the 'Fee Structure / ICA Breakdown' PDF is a scan of the same sheet)", "Current fee grid by office", "E&O $49 all offices (Chicago/Dallas: except referral; Philly: except personal sale). Residential sale fee $595 / $995 per side with price breaks at $650,000 (Chicago, Dallas, Houston/TUR) or $800,000 (Philly, DC). Commercial 12% min $595 (Chicago), 10% (Dallas, Houston/TUR, Philly), 5% (DC). Residential lease and referral 12% with minimums of $75 (Chicago, Dallas, max $995), $100 (Houston/TUR lease), $125 (Philly, DC); Houston/TUR and Philly referrals 10%. Caps 24 residential or $14,280 ($12,000 Houston/TUR); fee after cap $50 Dallas, $0 elsewhere. Philly enhanced E&O $110 on personal sales. Mentor 30%; Houston/TUR coaching 15%; broker leads 40% (Chicago, Dallas) or 25-50% (Philly). TUR legacy: 70% for 5 closings, 85% coaching, fee plan $150/$250/$350/$450 by commission, lease $100 or 10%, commercial 70/30 or 95/5. Leading Edge: E&O $0 ($35 in holding company); sale fee by price band $100 / $395 / $695 / $995 / $1,295 / $1,595 / $1,895 / $2,195 / $2,495 / $2,795; lease 10%; outside referral 10%; inside referral $0/$35; commercial 10%; cap $12,000 ($13,200 commercial-only); fee after cap $150; late fee $50 per billing period.", "Source for the all-entities and office-specific fee lists in Summary rows 4-7. Leading Edge's three aged items ($395, $150, $100) are exact schedule amounts."),
 ("All", "n/a", "JV ICA Transaction Fee Summary (docx)", "Older fee schedule", "Residential sale fee $495 / $895 (Philly break $800,000; Chicago $650,001; Houston $650,000; DC $800,001; Dallas $650,000). E&O $49 (Houston originally $45, 'has been changed to $49'; DC $45). Philly leases/referrals lesser of 10% (min $100) or $495/$895. Chicago/Dallas other events 10% min $75 max $895; Houston max $495/$895; DC referral/lease 10% min $100 max $295 plus E&O. Dallas after-cap fee $50. Philly termination: 70/30 on executed contracts; license lapse not cured in 30 days, Broker keeps 100%. Dallas: transactions not closed before termination paid 70/30.", "Explains $495 / $895 / $544 / $944 / $45 / $295 as historical fee amounts; the oldest aged items (2022-2024) likely fall under this schedule. Restates the Philly and Dallas termination terms."),
]
for i, row in enumerate(ICA, 5):
    for j, v in enumerate(row, 1):
        c = ic.cell(i, j, v); c.font = f_bold if j == 1 else f_norm
        c.alignment = Alignment(wrap_text=True, vertical="top"); c.border = box
    ic.row_dimensions[i].height = 120
ic.freeze_panes = "A5"

# ================================================================ Method
mt = wb.create_sheet("Method", 5)
put(mt, "A1", "Method", f_title)
lines = [
    "1. Source file: United_Cash_Requirements.xlsx as provided on 10/6/2026 (file name f519c0da-United_Cash_Requirements.xlsx).",
    "2. For each entity tab, the 'Previous Month's Accrued Commissions' block was read (Dallas M20:R143, Chicago R20:V190, Houston M20:R286, Philly S21:W133, DC R22:V194, TUR M21:R299, Leading Edge AA22:AE195 and AG22:AK195, Gallery M21:R299). Rows without an 'x' in the 'accounted for in Sage' column are the unbooked items that make up the accrued balance.",
    "3. Each entity's extracted unbooked total was tied to the tab's 'outstanding amount' cell (Dallas M18 = 59,991.62; Houston M18 = 76,889.50; Philly S19 = 29,708.78; TUR M19 = 55,528.45; Leading Edge AA20 = 250.00 and AG20 = 395.00; Gallery M19 = 4,001.11; Chicago R18 = 0; DC R20 = 0). All tie exactly.",
    "4. The current-month 'Deposited checks' blocks were read the same way and tie to each tab's A18/A19 (and the second-bank cells) and to the source Summary row 5.",
    "5. URE Corp carries no accrued commissions (A4 = -O58 where O58 is empty).",
    "6. Ages are computed from the deposit date to the as-of date. Dormant dates = deposit date + the state period in months (EDATE).",
    "7. Classification hints are heuristics on amount and memo text only; they are not conclusions.",
    "8. State rules were researched by web search on 10/6/2026. Primary statute pages could not be opened from this session; see the State Rules tab 'Verification status' column for what is and is not confirmed.",
    "9. Three fee documents were also provided 10/6/2026 (ICA Fee Structure Cheat Sheet xlsx, a scanned PDF of the same sheet, and the JV ICA Transaction Fee Summary docx with the older $495/$895 schedule); their amounts are in Summary rows 4-7.",
    "10. Entity Management Database (xlsx) provided 10/6/2026 supplies each holder's legal name and domicile; the domicile is used as the default governing state under the second priority rule.",
    "11. Five independent contractor agreements were provided 10/6/2026 (Dallas/Frisco, Chicago, Houston TXR-2301, TUR legacy sign-up docs, Pennsylvania ICA). Fee amounts in Summary rows 4-6 come from the user and from those fee schedules; the ICA Terms tab cites each clause used.",
]
for i, t in enumerate(lines, 3):
    put(mt, f"A{i}", t, wrap=True); mt.merge_cells(f"A{i}:J{i}"); mt.row_dimensions[i].height = 34
mt.column_dimensions["A"].width = 120

wb.save(OUT)
print("saved", OUT, "detail rows", len(rows), "current rows", len(cur))
