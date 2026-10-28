"""Build 'The Constrictor Funding Tracker.xlsx' from research/funding.json, and write the constrictor-funding skill references.
Google Sheets compatible (COUNTIFS, SUMIFS; inline dropdowns without commas).
usage: python3 build_funding.py OUT research/funding.json"""
import sys, os, json, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter as L

OUT = sys.argv[1]; R = json.load(open(sys.argv[2]))
NAVY, GOLD, GOLD_L = "081C2B", "EAA959", "FBEBD3"
F = "Arial"
f_title = Font(name=F, size=16, bold=True, color=GOLD); f_sub = Font(name=F, size=10, color="FFFFFF")
f_head = Font(name=F, size=10, bold=True, color="FFFFFF"); f_lab = Font(name=F, size=10); f_bold = Font(name=F, size=10, bold=True)
f_in = Font(name=F, size=10, color="0000FF"); f_note = Font(name=F, size=9, italic=True, color="555555")
f_sec = Font(name=F, size=11, bold=True, color=NAVY); f_calc = Font(name=F, size=10, bold=True, color="000000"); f_link = Font(name=F, size=9, color="1F5FA8")
fill_navy = PatternFill("solid", fgColor=NAVY); fill_in = PatternFill("solid", fgColor="FFFF00"); fill_goldl = PatternFill("solid", fgColor=GOLD_L)
thin = Side(style="thin", color="BFBFBF"); box = Border(top=thin, bottom=thin, left=thin, right=thin)
wrap = Alignment(wrap_text=True, vertical="top"); MONEY0 = '$#,##0;[Red]-$#,##0;"-"'; DATE = "mmm d, yyyy"
STATUS = ["Not started", "Researching", "Planning to apply", "Drafting", "Submitted", "Approved", "Declined", "Not eligible", "Not now", "Tell families"]
FITS = ["BBG", "CCCL", "Second Chance", "Families"]
ELIG = {"yes": "Yes, now", "after nonprofit": "After nonprofit", "after charity": "After charity", "no": "No"}
OWNERS = ["Cody", "Kyle"]
import datetime as dt
# known next dates from the research (program name prefix -> (date, note))
KNOWN = {"KidSport Ontario": (dt.datetime(2026, 12, 4), "Families apply by Dec 4 for Jan-Feb programs. BBG isn't registered with Boxing Ontario (Cody, Oct 2026), so KidSport can't cover BBG classes until it is."),
         "Community Sport and Recreation Infrastructure Fund": (dt.datetime(2026, 12, 22), "Only for new or expanded facilities; the province pays at most 25%."),
         "Rural Ontario Development (ROD) Program: Business": (dt.datetime(2027, 2, 25), "Fourth window Jan 18 to Feb 25, 2027."),
         "Canada Summer Jobs 2027 (private": (None, "Opens Oct 27, 2026 at 9 am ET. Open BBG's CRA payroll (RP) account first: the form asks for it. Draft application ready for Cody.")}

def short(s, n):
    s = re.sub(r"\s+", " ", str(s or "")).strip()
    return s if len(s) <= n else s[: n - 1].rstrip() + "…"

progs = R.get("programs", [])
order = {"yes": 0, "after nonprofit": 1, "after charity": 2, "no": 3}
progs.sort(key=lambda p: (order.get(str(p.get("eligible_now", "")).lower(), 4), {"high": 0, "medium": 1, "low": 2}.get(p.get("confidence"), 3), p.get("name", "")))

wb = Workbook()
def banner(ws, title, sub, width):
    for r in (1, 2):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=width)
        for c in range(1, width + 1): ws.cell(r, c).fill = fill_navy
    ws["A1"] = title; ws["A1"].font = f_title; ws["A1"].alignment = Alignment(vertical="center")
    ws["A2"] = sub; ws["A2"].font = f_sub; ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 30; ws.row_dimensions[2].height = 32
def head(ws, r, cols):
    for i, (t, w) in enumerate(cols, 1):
        c = ws.cell(r, i, t); c.font = f_head; c.fill = fill_navy; c.border = box; c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[r].height = 32
def inp(c, v=None, fmt=None):
    c.value = v; c.font = f_in; c.fill = fill_in; c.border = box; c.alignment = wrap
    if fmt: c.number_format = fmt
def txt(c, v, font=None):
    c.value = v; c.font = font or f_lab; c.border = box; c.alignment = wrap
def dv_list(ws, items, rng):
    d = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=True); ws.add_data_validation(d); d.add(rng)

# ---------------- Start Here
ws0 = wb.active; ws0.title = "Start Here"
banner(ws0, "THE CONSTRICTOR FUNDING TRACKER", "Grants and programs that fit BBG, CCCL, Second Chance and gym families: who can apply, when, and where each application stands.", 3)
ws0.column_dimensions["A"].width = 4; ws0.column_dimensions["B"].width = 26; ws0.column_dimensions["C"].width = 96
n_now = sum(1 for p in progs if str(p.get("eligible_now", "")).lower() == "yes")
n_later = sum(1 for p in progs if str(p.get("eligible_now", "")).lower() in ("after nonprofit", "after charity"))
steps = [
 ("1", "What's here", f"{len(progs)} programs researched on {R.get('generated', 'October 2026')}: {n_now} BBG, CCCL or families can use now, {n_later} that open up once Second Chance is a nonprofit or charity, and the rest for reference. The Not Eligible tab lists ones that don't fit and why."),
 ("2", "Check before relying", "Programs change every year. Before planning around a date or amount, open the source link and check the current intake. The Confidence column says how sure the research was."),
 ("3", "Track it", "For each program you go after, set Our status, the Next deadline and the Owner. 'Days left' counts down and the Summary tab shows what's due in the next 60 days."),
 ("4", "Apply", "Ask Claude: \"draft the Canada Summer Jobs application\". It drafts from real facts only and marks the gaps for you. You review and submit. The Application Kit tab lists what funders usually ask for."),
 ("5", "Money", "A grant counts on the Money Summary only when it's paid into the bank. Approved grants usually come with reports and receipts: put the report date in 'Report due'."),
 ("6", "Second Chance", "Most nonprofit funding needs an incorporated nonprofit (some need a registered charity). Keep Second Chance's applications and money separate from BBG and CCCL."),
]
r = 4
for n, a, b in steps:
    ws0.cell(r, 1, n).font = Font(name=F, size=12, bold=True, color=GOLD); ws0.cell(r, 2, a).font = f_bold
    ws0.cell(r, 1).alignment = Alignment(vertical="top"); ws0.cell(r, 2).alignment = Alignment(vertical="top", wrap_text=True)
    c = ws0.cell(r, 3, b); c.font = f_lab; c.alignment = wrap; ws0.row_dimensions[r].height = 34; r += 1
r += 1
ws0.cell(r, 2, "Colours").font = f_sec; r += 1
for lab, fill, t in (("Yellow, blue text", fill_in, "You type or pick here"), ("Black, bold", None, "Calculated: don't type over it")):
    c = ws0.cell(r, 2, lab); c.font = f_in if fill else f_lab
    if fill: c.fill = fill
    ws0.cell(r, 3, t).font = f_lab; r += 1

# ---------------- Programs
wp = wb.create_sheet("Programs")
PC = [("Program", 30), ("Funder", 20), ("Level", 10), ("For", 16), ("Can we apply?", 12), ("Who can apply", 26), ("What it funds", 40), ("Amount", 18),
      ("Intake / deadline", 24), ("Program status", 11), ("Matching money", 18), ("How to apply", 32), ("Source", 24), ("Confidence", 10),
      ("Our status", 14), ("Next deadline", 12), ("Days left", 8), ("Owner", 9), ("Amount asked", 11), ("Amount approved", 11), ("Report due", 12), ("Notes", 40)]
banner(wp, "PROGRAMS", "Researched programs (left) and where we stand with each (right, in yellow). Check the source before relying on a date or amount.", len(PC))
head(wp, 4, PC)
PMAX = 4 + len(progs) + 40
for k in range(PMAX - 4):
    rr = 5 + k
    p = progs[k] if k < len(progs) else None
    if p:
        el = ELIG.get(str(p.get("eligible_now", "")).lower(), str(p.get("eligible_now", "")))
        vals = [p.get("name"), p.get("funder"), str(p.get("level", "")).capitalize(), ", ".join(p.get("fits") or []), el, short(p.get("who_can_apply"), 160),
                short(p.get("what_it_funds"), 220), short(p.get("amount"), 120), short(p.get("intake"), 140), (lambda t: t[:1].upper() + t[1:])(str(p.get("status", ""))),
                short(p.get("matching"), 120), short(p.get("how_to_apply"), 170), p.get("source_url"), str(p.get("confidence", "")).capitalize()]
        for i, v in enumerate(vals, 1):
            txt(wp.cell(rr, i), v, f_bold if i == 1 else (f_link if i == 13 else None))
        default = "Not eligible" if el == "No" else ("Not now" if el.startswith("After") else ("Tell families" if (p.get("fits") or []) == ["Families"] else "Not started"))
        if str(p.get("name", "")).startswith("Canada Summer Jobs 2027 (private"): default = "Drafting"
        inp(wp.cell(rr, 15), default)
        wp.row_dimensions[rr].height = 78
    else:
        for i in range(1, 15): inp(wp.cell(rr, i))
        inp(wp.cell(rr, 15))
    kd = next((v for k_, v in KNOWN.items() if p and str(p.get("name", "")).startswith(k_)), None)
    inp(wp.cell(rr, 16), kd[0] if kd else None, DATE)
    c = wp.cell(rr, 17, f'=IF(P{rr}="","",P{rr}-TODAY())'); c.font = f_calc; c.border = box; c.number_format = "0"
    inp(wp.cell(rr, 18)); inp(wp.cell(rr, 19), None, MONEY0); inp(wp.cell(rr, 20), None, MONEY0); inp(wp.cell(rr, 21), None, DATE); inp(wp.cell(rr, 22), ((kd[1] + " ") if kd else "") + short(p.get("notes"), 200) if p else None)
dv_list(wp, STATUS, f"O5:O{PMAX}"); dv_list(wp, OWNERS, f"R5:R{PMAX}")
dv_list(wp, ["Yes, now", "After nonprofit", "After charity", "No"], f"E5:E{PMAX}")
d = DataValidation(type="date", allow_blank=True); wp.add_data_validation(d); d.add(f"P5:P{PMAX}"); d.add(f"U5:U{PMAX}")
wp.conditional_formatting.add(f"E5:E{PMAX}", CellIsRule(operator="equal", formula=['"Yes, now"'], fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(name=F, color="006100", bold=True)))
wp.conditional_formatting.add(f"E5:E{PMAX}", CellIsRule(operator="equal", formula=['"No"'], fill=PatternFill("solid", fgColor="EDEDED"), font=Font(name=F, color="808080")))
wp.conditional_formatting.add(f"Q5:Q{PMAX}", FormulaRule(formula=[f'AND(Q5<>"",Q5<=14,O5<>"Submitted",O5<>"Approved",O5<>"Declined")'], fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=F, color="9C0006", bold=True)))
wp.conditional_formatting.add(f"O5:O{PMAX}", CellIsRule(operator="equal", formula=['"Approved"'], fill=PatternFill("solid", fgColor="C6EFCE")))
wp.freeze_panes = "B5"; wp.auto_filter.ref = f"A4:V{PMAX}"

# ---------------- Summary
wsS = wb.create_sheet("Summary", 1)
banner(wsS, "SUMMARY", "What's open to us, where each application stands, and what's due.", 7)
wsS.column_dimensions["A"].width = 26
for c in "BCDEFG": wsS.column_dimensions[c].width = 15
hdr = ["For", "Programs", "Yes, now", "After nonprofit or charity", "Applied (submitted)", "Approved", "Approved $"]
for i, h in enumerate(hdr, 1):
    c = wsS.cell(4, i, h); c.font = f_head; c.fill = fill_navy; c.border = box; c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
wsS.row_dimensions[4].height = 32
D_ = f"Programs!$D$5:$D${PMAX}"; E_ = f"Programs!$E$5:$E${PMAX}"; O_ = f"Programs!$O$5:$O${PMAX}"; T_ = f"Programs!$T$5:$T${PMAX}"; Q_ = f"Programs!$Q$5:$Q${PMAX}"
for k, fit in enumerate(FITS):
    r = 5 + k
    wsS.cell(r, 1, fit).font = f_bold
    fs = [f'=COUNTIFS({D_},"*{fit}*")', f'=COUNTIFS({D_},"*{fit}*",{E_},"Yes, now")', f'=COUNTIFS({D_},"*{fit}*",{E_},"After*")',
          f'=COUNTIFS({D_},"*{fit}*",{O_},"Submitted")', f'=COUNTIFS({D_},"*{fit}*",{O_},"Approved")', f'=SUMIFS({T_},{D_},"*{fit}*",{O_},"Approved")']
    for j, f in enumerate(fs):
        c = wsS.cell(r, 2 + j, f); c.font = f_calc; c.border = box; c.alignment = Alignment(horizontal="center")
        if j == 5: c.number_format = MONEY0
r = 5 + len(FITS) + 1
wsS.cell(r, 1, "Deadlines").font = f_sec; r += 1
for lab, f in (("Due in the next 14 days", f'=COUNTIFS({Q_},">=0",{Q_},"<=14",{O_},"<>Submitted",{O_},"<>Approved",{O_},"<>Declined")'),
               ("Due in the next 60 days", f'=COUNTIFS({Q_},">=0",{Q_},"<=60",{O_},"<>Submitted",{O_},"<>Approved",{O_},"<>Declined")'),
               ("Deadline passed, not applied", f'=COUNTIFS({Q_},"<0",{O_},"<>Submitted",{O_},"<>Approved",{O_},"<>Declined",{O_},"<>Not eligible",{O_},"<>Not now")'),
               ("Reports due in the next 30 days", f'=COUNTIFS(Programs!$U$5:$U${PMAX},">="&TODAY(),Programs!$U$5:$U${PMAX},"<="&(TODAY()+30))')):
    wsS.cell(r, 1, lab).font = f_lab; c = wsS.cell(r, 2, f); c.font = f_calc; c.border = box; r += 1
wsS.cell(r + 1, 1, "Programs can count under more than one 'For'.").font = f_note

# ---------------- Not Eligible
wn = wb.create_sheet("Not Eligible")
banner(wn, "NOT ELIGIBLE (FOR NOW)", "Programs that came up but don't fit, and why. Worth a second look if something changes (e.g. Second Chance becomes a charity).", 3)
head(wn, 4, [("Program", 36), ("Why not", 80), ("Source", 40)])
for k, x in enumerate(R.get("not_eligible", [])):
    rr = 5 + k
    txt(wn.cell(rr, 1), x.get("name"), f_bold); txt(wn.cell(rr, 2), x.get("why")); txt(wn.cell(rr, 3), x.get("source_url"), f_link)
    wn.row_dimensions[rr].height = max(30, 14 * (len(str(x.get("why", ""))) // 90 + 1) + 6)

# ---------------- Application Kit
wa = wb.create_sheet("Application Kit")
banner(wa, "APPLICATION KIT", "What funders usually ask for, and the facts we can use. Get these ready once and most applications go faster.", 3)
head(wa, 4, [("Item", 34), ("Where it comes from / notes", 90), ("Ready?", 10)])
kit = [("Business or organization legal name and address", "BBG: 17 Snow Road, Bancroft ON. BBG and CCCL are two separate businesses (Cody, Oct 20, 2026): use the right one's legal name."),
       ("Business Number (CRA) and structure", "Sole proprietorship (Cody, Oct 20, 2026). BBG has a business number; open the payroll (RP) account before hiring. Your accountant can confirm anything else a form asks."),
       ("Number of employees", "Canada Summer Jobs and others depend on it (e.g. 50 or fewer for private-sector employers)."),
       ("Last year's financial statements or a budget", "Your accountant; the Money Summary helps for a simple budget."),
       ("Project description and timeline", "What the money will do, by when, and who it helps. Claude drafts it from real facts."),
       ("Project budget with quotes", "Each cost marked as a quote, estimate or known cost."),
       ("Letters of support", "From partners, the town or a sponsor. Ask early; never draft one for someone to sign as if it were theirs."),
       ("Proof of insurance", "Gym and event policies."),
       ("Results so far (real numbers only)", "400 tickets and 15 sponsors at the first show; about 90 gym members by March 2026 (The Highlander). No invented numbers."),
       ("Photos", "Only adults who agreed, or kids with a parent's written OK. Never kids in funder marketing without it."),
       ("Banking details for payment", "Entered on the funder's portal by Cody. Never put them in a sheet or a draft.")]
tips = R.get("tips", [])
for k, (a, b) in enumerate(kit):
    rr = 5 + k
    txt(wa.cell(rr, 1), a, f_bold); txt(wa.cell(rr, 2), b); inp(wa.cell(rr, 3))
    wa.row_dimensions[rr].height = 30
dv_list(wa, ["Yes", "No", "Not needed"], f"C5:C{4+len(kit)}")
r = 6 + len(kit)
if tips:
    wa.cell(r, 1, "Tips from the funders").font = f_sec; r += 1
    for t in tips:
        c = wa.cell(r, 1, "•"); c.font = f_lab; c2 = wa.cell(r, 2, t); c2.font = f_lab; c2.alignment = wrap
        wa.row_dimensions[r].height = max(18, 14 * (len(t) // 100 + 1) + 4); r += 1

# ---------------- Open Questions
wq = wb.create_sheet("Open Questions")
banner(wq, "QUESTIONS FOR CODY", "Eligibility hangs on these. Cody answered most of them on October 20, 2026; the rest are at the bottom.", 2)
head(wq, 4, [("Question", 70), ("Answer", 70)])
OQ = [q if isinstance(q, dict) else {"q": q, "a": ""} for q in R.get("open_questions", [])]
rr = 5
for group, title in ((True, "CODY'S ANSWERS (OCTOBER 20, 2026)"), (False, "STILL TO CONFIRM")):
    c = wq.cell(rr, 1, title); c.font = Font(name=F, size=10, bold=True, color=GOLD); rr += 1
    for q in OQ:
        if bool(q.get("a")) != group: continue
        txt(wq.cell(rr, 1), q["q"]); inp(wq.cell(rr, 2), q.get("a") or None)
        wq.row_dimensions[rr].height = max(18, 14 * (max(len(q["q"]), len(q.get("a") or "")) // 75 + 1) + 4); rr += 1
    rr += 1
if R.get("business_profile"):
    c = wq.cell(rr, 1, "BUSINESS PROFILE FOR APPLICATIONS"); c.font = Font(name=F, size=10, bold=True, color=GOLD); rr += 1
    for line_ in R["business_profile"]:
        txt(wq.cell(rr, 1), line_); wq.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=2)
        wq.row_dimensions[rr].height = max(18, 14 * (len(line_) // 140 + 1) + 4); rr += 1

for w in wb.worksheets:
    w.page_setup.orientation = "landscape"; w.sheet_properties.pageSetUpPr.fitToPage = True; w.page_setup.fitToWidth = 1; w.page_setup.fitToHeight = 0
wb.active = 0
wb.save(OUT); print("saved", OUT, len(progs), "programs", n_now, "now", n_later, "later")

# ---------------- skill references
D = os.path.join(os.path.dirname(os.path.abspath(__file__)), "skills/constrictor-funding/references"); os.makedirs(D, exist_ok=True)
Lm = ["# Funding programs (researched " + str(R.get("generated", "October 2026")) + ")", "",
      "Programs change every year. Check the source page for the current intake before Cody relies on a date or amount.", ""]
for grp, title in (("yes", "Can apply now"), ("after nonprofit", "Once Second Chance is an incorporated nonprofit"), ("after charity", "Once Second Chance is a registered charity"), ("no", "Not eligible (reference)")):
    ps = [p for p in progs if str(p.get("eligible_now", "")).lower() == grp]
    if not ps: continue
    Lm += [f"## {title}", ""]
    for p in ps:
        Lm += [f"### {p.get('name')} ({p.get('funder')})",
               f"- For: {', '.join(p.get('fits') or [])} · Who can apply: {p.get('who_can_apply')}",
               f"- Funds: {p.get('what_it_funds')} · Amount: {p.get('amount')}",
               f"- Intake: {p.get('intake')} ({p.get('status')}) · Matching: {p.get('matching')}",
               f"- How: {p.get('how_to_apply')}", f"- Source ({p.get('confidence')} confidence): {p.get('source_url')}"]
        if p.get("notes"): Lm.append(f"- Notes: {p.get('notes')}")
        Lm.append("")
if R.get("not_eligible"):
    Lm += ["## Came up but doesn't fit", ""] + [f"- **{x.get('name')}**: {x.get('why')} ({x.get('source_url')})" for x in R["not_eligible"]] + [""]
if R.get("open_questions"):
    OQ = [q if isinstance(q, dict) else {"q": q, "a": ""} for q in R["open_questions"]]
    Lm += ["## Cody's answers (October 20, 2026)", ""] + [f"- {q['q']} **{q['a']}**" for q in OQ if q.get("a")] + [""]
    Lm += ["## Still to confirm with Cody", ""] + [f"- {q['q']}" for q in OQ if not q.get("a")] + [""]
open(os.path.join(D, "programs.md"), "w").write("\n".join(Lm))
La = ["# Application kit", "", "What funders usually ask for. Draft only from confirmed facts; mark gaps as [Cody: ...].", ""]
if R.get("business_profile"):
    La += ["## Business profile (Cody's answers, October 2026)", ""] + [f"- {x}" for x in R["business_profile"]] + ["", "## What funders ask for", ""]
La += [f"- **{a}**: {b}" for a, b in kit]
if tips: La += ["", "## Tips from the funders", ""] + [f"- {t}" for t in tips]
open(os.path.join(D, "application-kit.md"), "w").write("\n".join(La) + "\n")
print("refs", os.listdir(D))
