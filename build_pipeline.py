"""Build 'CCCL Sponsorship Pipeline.xlsx' from research/sponsor_prospects.json (Google Sheets compatible)."""
import json, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.utils import get_column_letter as L

SRC, OUT = sys.argv[1], sys.argv[2]
prospects = json.load(open(SRC))
# Real 2026 sponsors Cody has named so far (the other 14 still to come)
PAST = [dict(name="Dillabough Millwrighting & Welding", town="Bancroft", area="North Hastings", category="Millwright & welding shop",
             why_fit="2026 Title Sponsor ($7,500). Thank them first, then offer Title again for 2027 before anyone else.",
             suggested_tier="Title", website="", phone="", email="",
             source_url="https://www.canpages.ca/page/ON/bancroft/dillabough-millwrighting-welding/7238519",
             group="A-existing", supporter="Yes (2026 Title Sponsor)", first=True)]
prospects = PAST + [p for p in prospects if p["name"].lower() not in {x["name"].lower() for x in PAST}]
# Gym-wall businesses Cody confirmed as May 2, 2026 sponsors (Oct 20, 2026); the other wall businesses stay "Yes (logo on gym wall)"
CONFIRMED_2026 = ("dixie lee", "pinestone", "it stix", "ico - international")
for p in prospects:
    if any(k in p["name"].lower() for k in CONFIRMED_2026):
        p["supporter"] = "Yes (2026 sponsor)"
        if p["name"].lower().startswith("ico"):
            p["note"] = "Cody confirmed 'ICO' sponsored May 2, 2026. Which ICO is a guess from the logo: confirm the full name with Cody before contacting."
BRANDS = json.load(open(sys.argv[3]))["brands"] if len(sys.argv) > 3 else []
import re as _re
GROUP_LABEL = {"protein": "Brand: protein", "gear": "Brand: fight gear", "alcohol": "Brand: alcohol", "other": "Brand: other"}
TIER_MAP = {"Silver $2,000": "Silver Sponsor", "Bronze $1,000": "Bronze Sponsor", "Gold $4,000": "Gold Sponsor", "In-kind": "In-kind Partner"}
CHARITY_ONLY = ("Jumpstart", "Kawartha Credit Union", "Ontario Power Generation", "Maple Leaf")
def brand_row(b):
    c = b.get("contact", "")
    email = (_re.findall(r"[\w.+-]+@[\w-]+\.[\w.]+", c) or [""])[0]
    phone = (_re.findall(r"1?-?\(?\d{3}\)?[ -]\d{3}-\d{4}", c) or [""])[0]
    url = (b.get("sponsorship_url", "") or "").split(" ")[0]
    if not url.startswith("http"): url = ""
    notes = "How to reach: " + c
    if b["group"] == "alcohol":
        notes = ("ALCOHOL: Cody decides. Only for a 19+ show or 19+ area, never near youth bouts or kids' content. Check with the AGCO, "
                 "the Athletics Commission and the venue first. " + notes)
    elif "Monster" in b["name"] or "GURU" in b["name"]:
        notes = "ENERGY DRINK: Cody decides. Keep it away from youth bouts and kids' content. " + notes
    elif any(k in b["name"] for k in CHARITY_ONLY):
        notes = "Funds registered charities or non-profits only: needs a non-profit partner to apply. " + notes
    return dict(name=b["name"], town=b.get("hq", "").split(" (")[0].split(";")[0], area=GROUP_LABEL.get(b["group"], "Brand"),
                category=("Energy drink (FLAG)" if ("Monster" in b["name"] or "GURU" in b["name"]) else b["group"].title()), why_fit=b.get("why_fit", "") + " ASK: " + b.get("realistic_ask", ""),
                suggested_tier=TIER_MAP.get(b.get("suggested_tier", ""), b.get("suggested_tier", "")), website=url, phone=phone, email=email,
                source_url=b.get("source_url", ""), group="C-brand", note=notes, confidence=b.get("confidence"))
TOP = ["Hayabusa", "BioSteel", "Canadian Protein", "MuscleTech", "Allmax", "STING"]
def brand_rank(name):
    for i, t in enumerate(TOP):
        if name.startswith(t): return i
    return len(TOP)
_br = [brand_row(b) for b in BRANDS]
for i, b in enumerate(_br): b["rank"] = (brand_rank(b["name"]), i)
prospects = prospects + _br
NAVY, GOLD, GOLD_L = "081C2B", "EAA959", "FBEBD3"
F = "Arial"
f_title = Font(name=F, size=16, bold=True, color=GOLD)
f_sub = Font(name=F, size=10, color="FFFFFF")
f_head = Font(name=F, size=10, bold=True, color="FFFFFF")
f_lab = Font(name=F, size=10)
f_bold = Font(name=F, size=10, bold=True)
f_in = Font(name=F, size=10, color="0000FF")
f_calc = Font(name=F, size=10, color="000000")
f_link = Font(name=F, size=10, color="008000")
f_note = Font(name=F, size=9, italic=True, color="555555")
f_sec = Font(name=F, size=11, bold=True, color=NAVY)
fill_navy = PatternFill("solid", fgColor=NAVY)
fill_goldl = PatternFill("solid", fgColor=GOLD_L)
fill_in = PatternFill("solid", fgColor="FFFF00")
thin = Side(style="thin", color="BFBFBF")
box = Border(top=thin, bottom=thin, left=thin, right=thin)
USD = '$#,##0;($#,##0);"-"'
DATE = 'mmm d, yyyy'
wrap = Alignment(wrap_text=True, vertical="top")
top = Alignment(vertical="top")

STAGES = ["Prospect", "Contacted", "Interested", "Proposal sent", "Signed", "Paid", "Not now"]
REAL = "2026 package (Cody). Same for 2027 unless Cody changes it."
PACKAGES = [  # name, price, what's in it, slots per show (suggested), status note
    ("Title Sponsor", 7500, "Naming rights (\"presented by\"), top logo on the cage and banners, featured on posters, ads and social media, "
     "named at fighter introductions, own posts before and after, VIP seating, a booth or display, recognised as the leading partner. Biggest logo on the event shirt and \"presented by\" on the livestream.", 1,
     REAL + " 2026: Dillabough Millwrighting & Welding. Cody calls this level Platinum."),
    ("Gold Sponsor", 4000, "Priority logo around the cage, featured on posters and ads, social media posts, several announcer mentions, "
     "VIP tickets, room for their banner and display. Large logo on the event shirt and the livestream.", 2, REAL + " Shirt and livestream sizes by level: confirmed by Cody (Oct 20, 2026)."),
    ("Silver Sponsor", 2000, "Logo on selected posters and promos, CCCL social media mention, shared banner space at the venue, "
     "announcer mention, free tickets, listed as an official sponsor. Logo on the event shirt and the livestream.", 4,
     REAL + " Quote $2,000 (the 2026 price) until Cody confirms. He listed Silver at $2,500 in Oct 2026."),
    ("Bronze Sponsor", 1000, "Logo on selected ads, social media sponsor announcement, named in sponsor thank-yous, announcer mention, "
     "free admission tickets. Small logo on the event shirt and the livestream sponsor slide.", 10, REAL),
    ("Community Sponsor", 300, "CCCL social media mention, name or logo in selected sponsor promos, listed as an official community supporter. "
     "Name on the event shirt and the livestream sponsor slide. $300-$500.", 20, REAL + " Price is the minimum of the $300-$500 range."),
    ("In-kind Partner", 0, "Product, prizes or services instead of cash: fighter bags, a sampling table, gear, prize packs. Perks agreed one-on-one, "
     "roughly matching the cash level of the same value.", 6,
     "For brands that give product. Agreed by Cody case by case. Only cash counts toward the sponsorship goal. "
     "Anything a fighter receives must be worth under $300 in total, or they lose amateur status; belts, medals and trophies are fine. No raffle without a lottery licence."),
    ("VIP Corporate Table", 1500, "Cage-side table for 8-10 guests (depends on the venue's tables) for a business to host staff or clients.", 6,
     "$1,500 a table (Cody); 6 sold in 2026. Seats 8-10 depending on the venue's tables (Kyle, Oct 2026)."),
]
SHOWS = ["Not booked yet", "Show 1 · 2027", "Show 2 · 2027", "Show 3 · 2027", "Show 4 · 2027", "All 2027 shows"]   # no 2027 show is booked yet (Cody, Oct 20)
OWNERS = ["Cody", "Kyle", "Staff"]

wb = Workbook()

def banner(ws, title, sub, width):
    for r in (1, 2):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=width)
        for c in range(1, width + 1):
            ws.cell(r, c).fill = fill_navy
    ws["A1"] = title; ws["A1"].font = f_title; ws["A1"].alignment = Alignment(vertical="center")
    ws["A2"] = sub; ws["A2"].font = f_sub; ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 30; ws.row_dimensions[2].height = 30

def head(ws, r, labels, widths=None):
    for i, h in enumerate(labels, 1):
        c = ws.cell(r, i, h); c.font = f_head; c.fill = fill_navy; c.border = box
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = 32
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[L(i)].width = w

# ------------------------------------------------------------------ Pipeline
wsP = wb.active; wsP.title = "Pipeline"
COLS = [  # header, width, key
    ("#", 5, "id"), ("Business", 30, "name"), ("Town", 14, "town"), ("Area", 16, "area"), ("Category", 18, "category"),
    ("Existing supporter?", 11, "group"), ("Suggested package", 17, "tier"), ("Stage", 14, None), ("Package agreed", 17, None),
    ("Amount agreed ($)", 12, None), ("Amount paid ($)", 12, None), ("Show", 20, None), ("Owner", 9, None),
    ("Contact person", 16, None), ("Phone", 15, "phone"), ("Email", 26, "email"), ("Website", 28, "website"),
    ("Last contact", 12, None), ("Next follow-up", 12, None), ("Follow-up due?", 10, None),
    ("Why they fit", 46, "why_fit"), ("Notes", 30, None), ("Source", 30, "source_url"),
]
NCOL = len(COLS)
banner(wsP, "CCCL SPONSORSHIP PIPELINE",
       "One row per business. Move the Stage as things happen. Only 'Paid' counts as money in the bank. "
       "Yellow columns are yours to fill in; grey ones come from research.", NCOL)
HR = 4
head(wsP, HR, [c[0] for c in COLS], [c[1] for c in COLS])
col = {c[0]: i + 1 for i, c in enumerate(COLS)}
CL = {k: L(v) for k, v in col.items()}
input_cols = {"Stage", "Package agreed", "Amount agreed ($)", "Amount paid ($)", "Show", "Owner", "Contact person",
              "Last contact", "Next follow-up", "Notes"}
area_order = {"Haliburton County": 0, "North Hastings": 1, "North Kawartha": 2}
grp = lambda p: 0 if str(p.get("group", "")).startswith("A") else (2 if str(p.get("group", "")).startswith("C") else 1)
rows = sorted(prospects, key=lambda p: (0 if p.get("first") else 1, grp(p), p.get("rank", (0, 0)),
                                       area_order.get(p.get("area", ""), 5), p.get("town", ""), p.get("name", "")))
MAXR = HR + 300
r = HR + 1
for i, p in enumerate(rows, 1):
    vals = {
        "#": i, "Business": p.get("name", ""), "Town": p.get("town", ""), "Area": p.get("area", ""),
        "Category": p.get("category", ""), "Existing supporter?": p.get("supporter") or ("Yes (logo on gym wall)" if grp(p) == 0 else ""),
        "Suggested package": {"Small $250-500": "Community Sponsor", "Standard $1,000": "Bronze Sponsor",
                              "Presenting $2,500+": "Silver Sponsor", "Title": "Title Sponsor"}.get(p.get("suggested_tier", ""), p.get("suggested_tier", "")),
        "Stage": "Prospect", "Phone": p.get("phone", ""), "Email": p.get("email", ""),
        "Website": p.get("website", "") or p.get("contact_page", ""), "Why they fit": p.get("why_fit", ""),
        "Source": p.get("source_url", ""), "Owner": "Cody", "Show": SHOWS[0],
    }
    if p.get("confidence") == "low":
        vals["Notes"] = "Unconfirmed: check the business name and details before contacting."
    if p.get("note"):
        vals["Notes"] = p["note"] + (" (Low confidence: check details first.)" if p.get("confidence") == "low" else "")
    elif "(ALCOHOL)" in p.get("category", "") or "(GAMBLING)" in p.get("category", ""):
        kind = "ALCOHOL" if "(ALCOHOL)" in p["category"] else "GAMBLING"
        vals["Notes"] = (f"{kind}: Cody decides. Keep their logo off youth bouts and kids' content." +
                         (" Unconfirmed details: check first." if p.get("confidence") == "low" else ""))
    for h, _, _ in COLS:
        c = wsP.cell(r, col[h], vals.get(h))
        c.font = f_in if h in input_cols else f_lab
        c.border = box; c.alignment = wrap if h in ("Why they fit", "Notes", "Business") else top
    r += 1
last_data = r - 1
# formulas + formatting for every row up to MAXR (so new rows work too)
for rr in range(HR + 1, MAXR + 1):
    if rr > last_data:
        for h, _, _ in COLS:
            c = wsP.cell(rr, col[h]); c.border = box
            c.font = f_in if h in input_cols else f_lab
        wsP.cell(rr, col["#"], f'=IF({CL["Business"]}{rr}="","",ROW()-{HR})').font = f_calc
    for h in input_cols:
        wsP.cell(rr, col[h]).fill = fill_in
    wsP.cell(rr, col["Amount agreed ($)"]).number_format = USD
    wsP.cell(rr, col["Amount paid ($)"]).number_format = USD
    wsP.cell(rr, col["Last contact"]).number_format = DATE
    wsP.cell(rr, col["Next follow-up"]).number_format = DATE
    st, nx = f'{CL["Stage"]}{rr}', f'{CL["Next follow-up"]}{rr}'
    c = wsP.cell(rr, col["Follow-up due?"],
                 f'=IF(OR({CL["Business"]}{rr}="",{st}="Paid",{st}="Not now",{nx}=""),"",IF({nx}<=TODAY(),"DUE",""))')
    c.font = Font(name=F, size=10, bold=True, color="C00000"); c.border = box; c.alignment = top
for rr in range(HR + 1, last_data + 1):
    wsP.row_dimensions[rr].height = 42

def dv_list(ws, items, rng, prompt=None):
    dv = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=True)
    if prompt: dv.prompt = prompt; dv.showInputMessage = True
    ws.add_data_validation(dv); dv.add(rng)

dv_list(wsP, STAGES, f'{CL["Stage"]}{HR+1}:{CL["Stage"]}{MAXR}')
dv_list(wsP, [p[0] for p in PACKAGES], f'{CL["Package agreed"]}{HR+1}:{CL["Package agreed"]}{MAXR}')
dv_list(wsP, SHOWS, f'{CL["Show"]}{HR+1}:{CL["Show"]}{MAXR}')
dv_list(wsP, OWNERS, f'{CL["Owner"]}{HR+1}:{CL["Owner"]}{MAXR}')
dvd = DataValidation(type="date", allow_blank=True); wsP.add_data_validation(dvd)
dvd.add(f'{CL["Last contact"]}{HR+1}:{CL["Next follow-up"]}{MAXR}')
# stage colours (whole row light tint by stage on the Stage cell)
stage_fill = {"Contacted": "DDEBF7", "Interested": "FFF2CC", "Proposal sent": "FCE4D6", "Signed": "E2EFDA",
              "Paid": "A9D08E", "Not now": "E7E6E6"}
for s, colr in stage_fill.items():
    wsP.conditional_formatting.add(f'{CL["Stage"]}{HR+1}:{CL["Stage"]}{MAXR}',
        CellIsRule(operator="equal", formula=[f'"{s}"'], fill=PatternFill("solid", fgColor=colr)))
wsP.freeze_panes = wsP.cell(HR + 1, col["Town"])
wsP.auto_filter.ref = f"A{HR}:{L(NCOL)}{MAXR}"
P = lambda h: f"Pipeline!${CL[h]}${HR+1}:${CL[h]}${MAXR}"

# ------------------------------------------------------------------ Summary
wsS = wb.create_sheet("Summary", 0)
banner(wsS, "SPONSORSHIP SUMMARY", "Updates by itself from the Pipeline tab. Nothing to type here except the goals in yellow.", 6)
for c, w in zip("ABCDEF", (34, 16, 16, 16, 16, 40)):
    wsS.column_dimensions[c].width = w
r = 4
wsS.cell(r, 1, "GOALS").font = f_sec; r += 1
goals = {}
for key, lab, val, note in (
    ("goal", "Sponsorship goal for the next show ($)", 18000,
     "The first show raised $18,000 from 15 sponsors (Cody). With 4+ shows in 2027, set a goal for each show. "
     "The Event Profit Predictor shows how much each venue needs to break even on a low night."),
    ("goal_n", "Number of sponsors wanted", 15, "Same as 2026."),
):
    wsS.cell(r, 1, lab).font = f_lab
    c = wsS.cell(r, 2, val); c.font = f_in; c.fill = fill_in; c.border = box; c.number_format = USD if key == "goal" else "0"
    n = wsS.cell(r, 6, note); n.font = f_note; n.alignment = wrap
    wsS.row_dimensions[r].height = 30
    goals[key] = f"$B${r}"; r += 1
r += 1
wsS.cell(r, 1, "WHERE THE MONEY IS").font = f_sec; r += 1
for i, h in enumerate(("Stage", "Businesses", "Dollars agreed", "Dollars paid"), 1):
    c = wsS.cell(r, i, h); c.font = f_head; c.fill = fill_navy; c.border = box
r += 1
first = r
for s in STAGES:
    wsS.cell(r, 1, s).font = f_lab
    c = wsS.cell(r, 2, f'=COUNTIFS({P("Stage")},"{s}",{P("Business")},"<>")'); c.number_format = "0"
    c3 = wsS.cell(r, 3, f'=SUMIFS({P("Amount agreed ($)")},{P("Stage")},"{s}")'); c3.number_format = USD
    c4 = wsS.cell(r, 4, f'=SUMIFS({P("Amount paid ($)")},{P("Stage")},"{s}")'); c4.number_format = USD
    for cc in (c, c3, c4): cc.font = f_calc; cc.border = box
    wsS.cell(r, 1).border = box
    r += 1
wsS.cell(r, 1, "Total").font = f_bold
for ci, fmt in ((2, "0"), (3, USD), (4, USD)):
    c = wsS.cell(r, ci, f"=SUM({L(ci)}{first}:{L(ci)}{r-1})"); c.font = f_bold; c.number_format = fmt; c.border = box; c.fill = fill_goldl
r += 2
wsS.cell(r, 1, "SCOREBOARD").font = f_sec; r += 1
score = [
    ("Signed or paid ($ agreed)", f'=SUMIFS({P("Amount agreed ($)")},{P("Stage")},"Signed")+SUMIFS({P("Amount agreed ($)")},{P("Stage")},"Paid")', USD),
    ("Money actually received", f'=SUM({P("Amount paid ($)")})', USD),
    ("Still to collect from signed sponsors", None, USD),
    ("% of goal received", None, "0%"),
    ("Sponsors signed or paid", f'=COUNTIFS({P("Stage")},"Signed",{P("Business")},"<>")+COUNTIFS({P("Stage")},"Paid",{P("Business")},"<>")', "0"),
    ("Follow-ups due today or overdue", f'=COUNTIF({P("Follow-up due?")},"DUE")', "0"),
    ("Businesses not contacted yet", f'=COUNTIFS({P("Stage")},"Prospect",{P("Business")},"<>")', "0"),
]
srow = {}
for lab, fml, fmt in score:
    wsS.cell(r, 1, lab).font = f_lab; wsS.cell(r, 1).border = box
    srow[lab] = r; r += 1
for lab, fml, fmt in score:
    rr = srow[lab]
    if lab == "Still to collect from signed sponsors":
        fml = f'=MAX(0,B{srow["Signed or paid ($ agreed)"]}-B{srow["Money actually received"]})'
    if lab == "% of goal received":
        fml = f'=IF({goals["goal"]}=0,0,B{srow["Money actually received"]}/{goals["goal"]})'
    c = wsS.cell(rr, 2, fml); c.font = Font(name=F, size=11, bold=True); c.number_format = fmt; c.border = box
wsS.conditional_formatting.add(f'B{srow["Follow-ups due today or overdue"]}',
    CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor="F8CBAD")))
r += 1
wsS.cell(r, 1, "PACKAGES SOLD").font = f_sec; r += 1
for i, h in enumerate(("Package", "Price ($)", "Sold (signed or paid)", "Slots", "Left"), 1):
    c = wsS.cell(r, i, h); c.font = f_head; c.fill = fill_navy; c.border = box
r += 1
for j, (name, price, *_rest) in enumerate(PACKAGES):
    pr = 5 + j  # row on Packages tab (set below)
    wsS.cell(r, 1, name).font = f_lab
    c2 = wsS.cell(r, 2, f"=Packages!B{pr}"); c2.font = f_link; c2.number_format = USD
    c3 = wsS.cell(r, 3, f'=COUNTIFS({P("Package agreed")},"{name}",{P("Stage")},"Signed")+COUNTIFS({P("Package agreed")},"{name}",{P("Stage")},"Paid")')
    c4 = wsS.cell(r, 4, f"=Packages!D{pr}"); c4.font = f_link
    c5 = wsS.cell(r, 5, f"=MAX(0,D{r}-C{r})")
    for cc in (wsS.cell(r, 1), c2, c3, c4, c5): cc.border = box
    r += 1
wsS.sheet_view.showGridLines = False

# ------------------------------------------------------------------ Packages
wsK = wb.create_sheet("Packages")
banner(wsK, "SPONSOR PACKAGES", "Cody's 2026 packages, used again for 2027 until he changes them. Prices, perks and slots in yellow can be changed. "
       "Slots per show are a suggestion.", 5)
for c, w in zip("ABCDE", (22, 12, 70, 8, 44)):
    wsK.column_dimensions[c].width = w
head(wsK, 4, ["Package", "Price ($)", "What the sponsor gets", "Slots", "Status"])
for j, (name, price, perks, slots, note) in enumerate(PACKAGES):
    rr = 5 + j
    wsK.cell(rr, 1, name).font = f_bold
    for ci, v, fmt in ((2, price, USD), (3, perks, None), (4, slots, "0")):
        c = wsK.cell(rr, ci, v); c.font = f_in; c.fill = fill_in; c.border = box; c.alignment = wrap
        if fmt: c.number_format = fmt
    n = wsK.cell(rr, 5, note); n.font = f_note; n.alignment = wrap
    wsK.row_dimensions[rr].height = max(58, 14 * (len(perks) // 78 + 1))
dv_pk = DataValidation(type="list", formula1=f"=Packages!$A$5:$A${4 + len(PACKAGES)}", allow_blank=True)

# ------------------------------------------------------------------ Fulfilment
wsF = wb.create_sheet("Fulfilment")
FCOLS = [("Business", 30), ("Package", 17), ("Show", 20), ("Logo file received", 11), ("Banner ordered", 10),
         ("Banner hung at show", 10), ("Own thank-you post", 11), ("Tagged in fight-week posts", 11),
         ("On poster", 9), ("On the livestream", 10), ("On the shirts", 9), ("Tickets / VIP seats sent", 11),
         ("Announcer shout-out", 11), ("Photo + thank-you sent after", 12), ("All done?", 9), ("Notes", 30)]
NCHK = len(FCOLS) - 5                      # tick columns D..
LCHK = L(3 + NCHK)                         # last tick column
banner(wsF, "SPONSOR DELIVERY CHECKLIST",
       "When a sponsor signs, add them here (pick from the list) and tick off each promise as it's done. Every sponsor should get everything they paid for.",
       len(FCOLS))
head(wsF, 4, [c[0] for c in FCOLS], [c[1] for c in FCOLS])
YN = ["Yes", "No", "N/A"]
for rr in range(5, 65):
    for ci in range(1, len(FCOLS) + 1):
        c = wsF.cell(rr, ci); c.border = box; c.font = f_in
        if ci not in (len(FCOLS) - 1,): c.fill = fill_in
    chk = f"D{rr}:{LCHK}{rr}"
    c = wsF.cell(rr, len(FCOLS) - 1,
                 f'=IF(A{rr}="","",IF(COUNTIF({chk},"Yes")+COUNTIF({chk},"N/A")={NCHK},"DONE",{NCHK}-COUNTIF({chk},"Yes")-COUNTIF({chk},"N/A")&" left"))')
    c.font = Font(name=F, size=10, bold=True); c.fill = PatternFill(fill_type=None)
dvb = DataValidation(type="list", formula1=f"=Pipeline!$B${HR+1}:$B${MAXR}", allow_blank=True)
wsF.add_data_validation(dvb); dvb.add("A5:A64")
wsF.add_data_validation(dv_pk); dv_pk.add("B5:B64")
dv_list(wsF, SHOWS, "C5:C64")
dv_list(wsF, YN, f"D5:{LCHK}64")
wsF.conditional_formatting.add(f"D5:{LCHK}64", CellIsRule(operator="equal", formula=['"Yes"'], fill=PatternFill("solid", fgColor="C6EFCE")))
wsF.conditional_formatting.add(f"{L(len(FCOLS)-1)}5:{L(len(FCOLS)-1)}64", CellIsRule(operator="equal", formula=['"DONE"'], fill=PatternFill("solid", fgColor="A9D08E")))
# example row
ex = ["(Example) Lockhart Bros Contracting", "Bronze Sponsor", "Show 1 · 2027", "Yes", "N/A", "N/A", "Yes", "No", "Yes", "No", "No", "No", "No", "No", None,
      "Example row showing the format. Bronze gets a logo on ads and the shirt, not its own banner, so the two banner columns are N/A. Replace it with a real sponsor."]
for ci, v in enumerate(ex, 1):
    if v is not None: wsF.cell(5, ci, v)
wsF.freeze_panes = "B5"
wsP.add_data_validation  # noqa

# ------------------------------------------------------------------ Start Here
ws0 = wb.create_sheet("Start Here", 0)
banner(ws0, "CCCL SPONSORSHIP PIPELINE",
       "Find sponsors, pitch them, follow up, collect, and deliver what was promised. Built for Cody Purtell by Kyle. Opens in Google Sheets or Excel.", 3)
ws0.column_dimensions["A"].width = 4; ws0.column_dimensions["B"].width = 26; ws0.column_dimensions["C"].width = 95
r = 4
def sec(t):
    global r
    ws0.cell(r, 2, t).font = f_sec; r += 1
def line(a, b, n=None):
    global r
    if n: ws0.cell(r, 1, n).font = Font(name=F, size=12, bold=True, color=GOLD)
    ws0.cell(r, 2, a).font = f_bold; ws0.cell(r, 2).alignment = top
    c = ws0.cell(r, 3, b); c.font = f_lab; c.alignment = wrap
    r += 1
sec("THE STAGES")
for a, b in (("Prospect", "Researched, not contacted yet."),
             ("Contacted", "Pitch sent or a conversation started. Set a next follow-up date."),
             ("Interested", "They replied and want to hear more."),
             ("Proposal sent", "They have the package and price."),
             ("Signed", "They said yes and agreed an amount. Not money yet."),
             ("Paid", "The money is in the account. Only this counts as revenue."),
             ("Not now", "Said no or not this year. Leave a note and try again next season.")):
    line(a, b)
r += 1
sec("EVERY WEEK (15 MINUTES)")
for n, (a, b) in enumerate((
        ("Monday", "Open Summary. Anyone marked DUE gets a follow-up today. Claude can draft them."),
        ("New sponsor signed", "Set Stage to Signed, fill in the package and amount, and add them to the Fulfilment tab."),
        ("Money arrives", "Set Stage to Paid and fill in Amount paid."),
        ("Before the show", "Check Fulfilment: banners hung, posts done, tickets sent."),
        ("After the show", "Send every sponsor a photo and a thank-you. Mark it on Fulfilment. That's how they come back next year.")), 1):
    line(a, b, str(n))
r += 1
sec("TABS")
for a, b in (("Summary", "Money agreed vs paid, follow-ups due, packages sold."),
             ("Pipeline", "Every business, one row each. Yellow columns are yours. Grey columns came from research."),
             ("Packages", "Cody's sponsor packages (Title $7,500 down to Community $300-500, plus VIP tables)."),
             ("Fulfilment", "What each sponsor was promised and whether it's done.")):
    line(a, b)
r += 1
sec("ABOUT THE LIST")
line("Where it came from", f"The 2026 Title Sponsor, {len(rows) - len(PAST) - len(BRANDS)} local businesses and {len(BRANDS)} national and Ontario brands (rows marked 'Brand:' in the Area column), researched by Claude from public websites in October 2026. "
     "Phone numbers and emails are each business's public contact details. Check them before reaching out.")
line("Existing supporters", "2026 sponsors (Cody, Oct 20, 2026): Dillabough Millwrighting & Welding (Title, $7,500), Dixie Lee, Pinestone Resort, IT STIX Signs & Graphics and ICO; most were Bronze. "
     "Businesses marked 'Yes (logo on gym wall)' support the gym but weren't May 2 sponsors. Start with both groups: thank them first, then pitch 2027. "
     "Cody has promised some sponsors spots at future Apsley, Bancroft and Haliburton shows: check with him before pitching for those towns. Add the other 2026 sponsors when Cody lists them.")
line("Brands", "Start with Hayabusa (Kingston fight gear), BioSteel / Canadian Protein and MuscleTech / Six Star. Ask for product first (fighter bags, a sampling table, prize packs), "
     "then cash once there are 2027 crowd numbers. Anything a fighter receives must be worth under $300 in total (amateur rule); no raffle without a lottery licence. Alcohol and energy drinks stay away from youth bouts and kids' content; Proper No. Twelve only for a 19+ show or area, and only after checking with the AGCO and the venue.")
line("Show column", "No 2027 show is booked yet, so every row says 'Not booked yet'. Change it once Cody confirms a town and date. Never name a town or date in a pitch until then.")
line("Colour key", "Yellow cells with blue text are yours to fill in. Black text is calculated. Don't type over it.")
ws0.sheet_view.showGridLines = False

for ws in wb.worksheets:
    ws.sheet_properties.tabColor = GOLD if ws.title in ("Start Here", "Summary") else NAVY
wb.active = 0
wb.save(OUT)
print("rows", len(rows), "last_data", last_data, "saved", OUT)
