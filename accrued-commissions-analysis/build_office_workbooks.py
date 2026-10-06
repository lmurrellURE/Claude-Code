"""Build one fill-in workbook per office from source-extract/accrued_rows.csv.

Usage: python3 build_office_workbooks.py <out_dir>   (run from the folder containing source-extract/)
Each workbook holds only that office's unapplied prior-month deposits, the adopted dormancy test for
its governing state, a priority flag per item, and yellow fill-in columns for the office.
"""
import csv, datetime, sys, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L

OUT_DIR = sys.argv[1] if len(sys.argv) > 1 else "office-workbooks"
AS_OF = datetime.date(2026, 10, 6)
F = "Arial"
f_norm = Font(name=F, size=10); f_bold = Font(name=F, size=10, bold=True); f_title = Font(name=F, size=14, bold=True)
f_input = Font(name=F, size=10, color="0000FF"); f_hdr = Font(name=F, size=10, bold=True, color="FFFFFF")
fill_hdr = PatternFill("solid", fgColor="1F3864"); fill_in = PatternFill("solid", fgColor="FFFF00")
fill_sub = PatternFill("solid", fgColor="D9E1F2"); fill_p1 = PatternFill("solid", fgColor="F8CBAD"); fill_p2 = PatternFill("solid", fgColor="FFE699")
thin = Side(style="thin", color="BFBFBF"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
CUR = '$#,##0.00;($#,##0.00);-'; DATE = 'mm/dd/yyyy'

# entity, office label, governing state (LLC formation state), adopted period (yrs), next report cutoff, address note, fee note
OFFICES = [
    ("Dallas", "Dallas", "TX", 3, datetime.date(2027, 3, 1), "", "$595 or $995 sale fee per side plus $49 E&O (older files $495/$895 plus $45/$49); $50 after cap"),
    ("Houston", "Houston", "TX", 3, datetime.date(2027, 3, 1), "", "$595 or $995 sale fee per side plus $49 E&O (older files $495/$895 plus $45/$49)"),
    ("TUR", "Texas United Realty", "TX", 3, datetime.date(2027, 3, 1), "", "United plan $595/$995 plus $49 E&O; legacy plan $150/$250/$350/$450 by commission size, $60 lease, $179 annual"),
    ("Chicago", "Chicago", "TX", 3, datetime.date(2027, 3, 1), "If the agent's last-known address is in Illinois, Illinois law applies instead (1 year); send the address and the Director will reassess.", "$595 or $995 sale fee per side plus $49 E&O"),
    ("Philly", "Philadelphia", "TX", 3, datetime.date(2027, 3, 1), "If the agent's last-known address is in Pennsylvania, Pennsylvania law applies instead (2 years); send the address and the Director will reassess.", "$595 or $995 sale fee per side plus $49 E&O ($110 enhanced E&O on personal sales); also $495, $295"),
    ("Gallery", "URE Gallery", "FL", 3, datetime.date(2026, 12, 31), "If the agent's last-known address is in Florida, the period is 5 years; send the address and the Director will reassess.", "Per the Gallery fee schedule"),
    ("Leading Edge", "Leading Edge Realty", "AL", 1, datetime.date(2027, 6, 30), "", "Sale fee by price band $100 to $2,795; E&O $0 or $35; $150 after cap"),
]
STATE_NAME = {"TX": "Texas", "FL": "Florida", "AL": "Alabama"}

rows = [r for r in csv.DictReader(open("source-extract/accrued_rows.csv")) if r["booked"] == "False" and r["amount"]]
for r in rows:
    r["amt"] = float(r["amount"]); r["d"] = datetime.date.fromisoformat(r["date"]) if r["date"] else None

def put(ws, ref, val, font=f_norm, fmt=None, fill=None, wrap=False):
    c = ws[ref]; c.value = val; c.font = font
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if wrap: c.alignment = Alignment(wrap_text=True, vertical="top")
    return c

os.makedirs(OUT_DIR, exist_ok=True)
for ent, label, st, period, cutoff, addr_note, fee_note in OFFICES:
    items = sorted([r for r in rows if r["entity"] == ent], key=lambda r: r["d"] or datetime.date(1900, 1, 1))
    wb = Workbook()

    # ---------------- Settings
    sg = wb.active; sg.title = "Settings"
    put(sg, "A1", f"Settings - {label}", f_title)
    put(sg, "A3", "Office"); put(sg, "B3", label, f_input)
    put(sg, "A4", "Holder entity governing state (LLC formation state)"); put(sg, "B4", st, f_input)
    put(sg, "A5", "Dormancy period applied (years, adopted 10/6/2026)"); put(sg, "B5", period, f_input)
    put(sg, "A6", "Next state report cutoff"); put(sg, "B6", cutoff, f_input, DATE)
    put(sg, "A7", "As-of date for ages"); put(sg, "B7", AS_OF, f_input, DATE, fill_in)
    put(sg, "A8", "Return completed workbook by"); put(sg, "B8", None, f_input, DATE, fill_in)
    put(sg, "A9", "Office fee schedule (reference)"); put(sg, "B9", fee_note, wrap=True)
    put(sg, "A11", "Do not change rows 3-6; they come from the company's unclaimed property policy. Row 7 re-ages every item; row 8 is for the Director to set before sending.", wrap=True)
    sg.merge_cells("A11:D11"); sg.row_dimensions[11].height = 32; sg.merge_cells("B9:D9"); sg.row_dimensions[9].height = 32
    sg.column_dimensions["A"].width = 52; sg.column_dimensions["B"].width = 22; sg.column_dimensions["C"].width = 22; sg.column_dimensions["D"].width = 22

    # ---------------- Instructions
    ins = wb.create_sheet("Instructions", 0)
    put(ins, "A1", f"Unapplied deposits review - {label}", f_title)
    put(ins, "A2", "Prepared 10/6/2026 by Financial Reporting and Analysis. Please complete the yellow columns on the Items tab and return the workbook by the date on the Settings tab.", wrap=True)
    text = [
        ("What this is", f_bold),
        ("We are holding money from closings that was never matched to a transaction. Some of it belongs to agents or to the people who sent it, and state law says that after a set number of years we must either pay it to them or hand it to the state. Once a year, in December, Financial Reporting and Analysis finds the items that have reached that point and clears them; everything younger keeps being matched as usual.", f_norm),
        ("Why it matters", f_bold),
        ("This money belongs to our agents and the people who sent it. Every state we operate in can audit ten years back and charge interest and penalties on anything held past its deadline, and the only thing that gets those waived is showing we had a written process, ran it on schedule and filed on time.", f_norm),
        ("What we need from you", f_bold),
        ("1. Work the Items tab oldest first. Priority 1 items are already past the threshold; Priority 2 items will reach it by the next state cutoff. Do those first, then the rest as time allows.", f_norm),
        ("2. For each item, find the transaction in Paperless Pipeline and enter the file number, property address and closing date.", f_norm),
        ("3. Enter the agent's name and last-known mailing address (city and state). The state matters: it can change which state's rules apply.", f_norm),
        ("4. Enter the split from the Disbursement Authorization or Closing Disclosure: company fee, E&O, agent share. If the document cannot be found, leave the split blank and say so in Notes; the Director will estimate it as the check less the standard fee.", f_norm),
        ("5. If the money is not ours (wrong office, duplicate, unknown), enter the amount to refund in the Refund column and name who it goes back to in Notes.", f_norm),
        ("6. Pick a Status for every item and attach or link the supporting document. The Check column must read 0.00 once the split is entered.", f_norm),
        ("What not to do", f_bold),
        ("Do not pay an agent, refund a payer, or write anything off from this list. Financial Reporting and Analysis makes every entry after the review. Do not delete rows; mark the Status instead. Items already booked in Sage since the extract date: set Status to 'Already booked in Sage' and give the date.", f_norm),
        ("The rule applied to this office", f_bold),
        (f"The holder entity for this office was formed in {STATE_NAME[st]}, so {STATE_NAME[st]} rules apply where the agent's address is unknown: {period} year{'s' if period != 1 else ''} from the deposit date. The next {STATE_NAME[st]} report cutoff is {cutoff.strftime('%m/%d/%Y')}. {addr_note}".strip(), f_norm),
        ("Fee amounts for this office, for reference: " + fee_note + ". A check at exactly a fee amount is the company's fee and needs no agent split; confirm it and set Status to 'Company fee - confirmed'.", f_norm),
        ("Questions: Director of Financial Reporting and Analysis.", f_norm),
    ]
    for i, (t, f) in enumerate(text, 4):
        put(ins, f"A{i}", t, f, wrap=True); ins.merge_cells(f"A{i}:H{i}")
        ins.row_dimensions[i].height = 16 if f is f_bold else (30 if len(t) < 180 else 46)
    ins.column_dimensions["A"].width = 16
    for c in "BCDEFGH": ins.column_dimensions[c].width = 16
    ins.merge_cells("A2:H2"); ins.row_dimensions[2].height = 30

    # ---------------- Items
    it = wb.create_sheet("Items", 1)
    put(it, "A1", f"{label} - unapplied deposits as of " + AS_OF.strftime("%m/%d/%Y"), f_title)
    put(it, "A2", "Columns A-F come from the cash requirements workbook. G-I are formulas. Yellow columns J-W are for the office to complete. Priority: 1 = past the dormancy threshold today; 2 = reaches it by the next state cutoff; 3 = not yet.", wrap=True)
    it.merge_cells("A2:W2"); it.row_dimensions[2].height = 30
    put(it, "A3", "As-of date"); put(it, "B3", "=Settings!$B$7", fmt=DATE); put(it, "C3", "Period (yrs)"); put(it, "D3", "=Settings!$B$5"); put(it, "E3", "Next cutoff"); put(it, "F3", "=Settings!$B$6", fmt=DATE)
    labels = ["#", "Deposit date", "Check #", "Payer / check writer", "Amount", "Memo from cash workbook",
              "Age (yrs)", "Dormant on", "Priority",
              "Paperless Pipeline file / transaction #", "Property address", "Closing date", "Agent name", "Agent last-known mailing address (city, state)", "Agent address state (2 letters)",
              "Company fee $", "E&O $", "Agent share $", "Refund to payer $ (not ours)", "Check (should be 0.00)", "Supporting document", "Status", "Notes"]
    widths = [5, 11, 10, 28, 12, 45, 8, 11, 30, 20, 28, 11, 22, 30, 9, 12, 10, 12, 14, 12, 18, 26, 40]
    H = 5
    for j, lab in enumerate(labels, 1):
        c = it.cell(H, j, lab); c.font = f_hdr; c.fill = fill_hdr; c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = box
        it.column_dimensions[L(j)].width = widths[j - 1]
    it.row_dimensions[H].height = 44
    R0 = H + 1
    for i, r in enumerate(items):
        rr = R0 + i
        put(it, f"A{rr}", i + 1); put(it, f"B{rr}", r["d"], fmt=DATE); put(it, f"C{rr}", r["check"])
        put(it, f"D{rr}", r["payer"]); put(it, f"E{rr}", r["amt"], fmt=CUR); put(it, f"F{rr}", r["memo"])
        put(it, f"G{rr}", f'=IF(B{rr}="","",($B$3-B{rr})/365.25)', fmt="0.00")
        put(it, f"H{rr}", f'=IF(B{rr}="","",EDATE(B{rr},12*$D$3))', fmt=DATE)
        put(it, f"I{rr}", f'=IF(H{rr}="","",IF(H{rr}<=$B$3,"1 - Past threshold: resolve now",IF(H{rr}<=$F$3,"2 - Reaches threshold by next cutoff","3 - Not yet")))')
        for col in "JKLMNOPQRSUVW": put(it, f"{col}{rr}", None, f_input, fill=fill_in)
        it[f"L{rr}"].number_format = DATE
        for col in "PQRS": it[f"{col}{rr}"].number_format = CUR
        put(it, f"T{rr}", f'=IF(COUNT(P{rr}:S{rr})=0,"",ROUND(E{rr}-N(P{rr})-N(Q{rr})-N(R{rr})-N(S{rr}),2))', fmt=CUR)
        for c in range(1, 24): it.cell(rr, c).border = box
    n = len(items)
    RT = R0 + n
    put(it, f"A{RT}", "TOTAL", f_bold)
    if n:
        put(it, f"E{RT}", f"=SUM(E{R0}:E{RT-1})", f_bold, CUR)
        for col in "PQRS": put(it, f"{col}{RT}", f"=SUM({col}{R0}:{col}{RT-1})", f_bold, CUR)
        put(it, f"A{RT+2}", "Priority 1 items", f_bold); put(it, f"E{RT+2}", f'=SUMIF(I{R0}:I{RT-1},"1 -*",E{R0}:E{RT-1})', fmt=CUR)
        put(it, f"A{RT+3}", "Priority 2 items", f_bold); put(it, f"E{RT+3}", f'=SUMIF(I{R0}:I{RT-1},"2 -*",E{R0}:E{RT-1})', fmt=CUR)
        put(it, f"A{RT+4}", "Items with a Status entered", f_bold); put(it, f"E{RT+4}", f'=COUNTA(V{R0}:V{RT-1})')
        it.auto_filter.ref = f"A{H}:W{RT-1}"
        from openpyxl.formatting.rule import FormulaRule
        it.conditional_formatting.add(f"A{R0}:I{RT-1}", FormulaRule(formula=[f'LEFT($I{R0},1)="1"'], fill=fill_p1))
        it.conditional_formatting.add(f"A{R0}:I{RT-1}", FormulaRule(formula=[f'LEFT($I{R0},1)="2"'], fill=fill_p2))
    else:
        put(it, f"A{R0}", "No unapplied prior-month deposits for this office as of the extract date. Use this template for items that age in later.", wrap=False)
    for c in range(1, 24): it.cell(RT, c).border = box; it.cell(RT, c).fill = fill_sub
    last = max(RT - 1, R0 + 200)
    dv_doc = DataValidation(type="list", formula1='"Disbursement Authorization,Closing Disclosure,Both,None found"', allow_blank=True)
    dv_status = DataValidation(type="list", formula1='"Matched - pay agent,Company fee - confirmed,Refund to payer - not ours,Agent unknown - could not identify,Already booked in Sage,Legal hold,Other - see Notes"', allow_blank=True)
    it.add_data_validation(dv_doc); it.add_data_validation(dv_status)
    dv_doc.add(f"U{R0}:U{last}"); dv_status.add(f"V{R0}:V{last}")
    it.freeze_panes = f"G{R0}"

    fn = os.path.join(OUT_DIR, f"{ent.replace(' ', '_')}_Unapplied_Deposits_Review.xlsx")
    wb.save(fn); print("saved", fn, n, "items", round(sum(r["amt"] for r in items), 2))
