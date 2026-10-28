"""Build 'Constrictor HQ.xlsx': Cody's command centre (Google Sheets compatible)."""
import sys, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter as L

OUT = sys.argv[1]
NAVY, GOLD, GOLD_L = "081C2B", "EAA959", "FBEBD3"
F = "Arial"
f_title = Font(name=F, size=16, bold=True, color=GOLD); f_sub = Font(name=F, size=10, color="FFFFFF")
f_head = Font(name=F, size=10, bold=True, color="FFFFFF"); f_lab = Font(name=F, size=10); f_bold = Font(name=F, size=10, bold=True)
f_in = Font(name=F, size=10, color="0000FF"); f_note = Font(name=F, size=9, italic=True, color="555555")
f_sec = Font(name=F, size=11, bold=True, color=NAVY); f_big = Font(name=F, size=18, bold=True, color=NAVY)
fill_navy = PatternFill("solid", fgColor=NAVY); fill_in = PatternFill("solid", fgColor="FFFF00"); fill_goldl = PatternFill("solid", fgColor=GOLD_L)
thin = Side(style="thin", color="BFBFBF"); box = Border(top=thin, bottom=thin, left=thin, right=thin)
DATE = "ddd mmm d"; wrap = Alignment(wrap_text=True, vertical="top"); top = Alignment(vertical="top")
BUSINESSES = ["BBG", "CCCL", "Second Chance", "Athlete brand", "Office"]
OWNERS = ["Cody", "Kyle", "Sean", "Jordan", "Front desk", "Volunteer"]
wb = Workbook()

def banner(ws, title, sub, width):
    for r in (1, 2):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=width)
        for c in range(1, width + 1): ws.cell(r, c).fill = fill_navy
    ws["A1"] = title; ws["A1"].font = f_title; ws["A1"].alignment = Alignment(vertical="center")
    ws["A2"] = sub; ws["A2"].font = f_sub; ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 30; ws.row_dimensions[2].height = 30

def table(ws, cols, maxr, hr=4, formula_cols=None, fills=True):
    formula_cols = formula_cols or {}
    for i, (h, w) in enumerate(cols, 1):
        c = ws.cell(hr, i, h); c.font = f_head; c.fill = fill_navy; c.border = box
        c.alignment = Alignment(wrap_text=True, vertical="center"); ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[hr].height = 32
    for r in range(hr + 1, maxr + 1):
        for i in range(1, len(cols) + 1):
            c = ws.cell(r, i); c.border = box; c.alignment = wrap
            if i in formula_cols:
                c.value = formula_cols[i](r); c.font = Font(name=F, size=10, bold=True, color="C00000")
            elif fills:
                c.font = f_in; c.fill = fill_in

def dv(ws, items, rng):
    d = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=True); ws.add_data_validation(d); d.add(rng)

def dvdate(ws, rng):
    d = DataValidation(type="date", allow_blank=True); ws.add_data_validation(d); d.add(rng)

# ------------------------------------------------------------------ Tasks
wsT = wb.active; wsT.title = "Tasks"
TC = [("Task", 44), ("Business", 13), ("Owner", 11), ("Due", 12), ("Priority", 10), ("Status", 12), ("Overdue?", 10),
      ("Added", 12), ("Done on", 12), ("Notes", 40)]
banner(wsT, "TASKS", "Everything the team has to do, one row each. Claude's morning report reads this. Mark Done and the date when it's finished.", len(TC))
MT = 4 + 400
table(wsT, TC, MT, formula_cols={7: lambda r: f'=IF(OR(A{r}="",F{r}="Done",D{r}=""),"",IF(D{r}<TODAY(),"OVERDUE",IF(D{r}=TODAY(),"TODAY","")))'})
for r in range(5, MT + 1):
    for col in ("D", "H", "I"): wsT[f"{col}{r}"].number_format = DATE
dv(wsT, BUSINESSES, f"B5:B{MT}"); dv(wsT, OWNERS, f"C5:C{MT}"); dv(wsT, ["High", "Normal", "Low"], f"E5:E{MT}")
dv(wsT, ["To do", "In progress", "Waiting on someone", "Done"], f"F5:F{MT}"); dvdate(wsT, f"D5:D{MT}"); dvdate(wsT, f"H5:I{MT}")
wsT.conditional_formatting.add(f"F5:F{MT}", CellIsRule(operator="equal", formula=['"Done"'], fill=PatternFill("solid", fgColor="C6EFCE")))
HD = "'Start Here'!$C$12"   # handoff date: starter tasks and the first posts count from it
starter = [   # (task, business, owner, days after the handoff date OR a fixed date, priority)
    ("Open BBG's CRA payroll account (RP) if it isn't done yet: the Canada Summer Jobs form asks for the payroll number", "BBG", "Cody", 2, "High"),
    ("Set up GCOS and submit Canada Summer Jobs 2027 from the draft (Command Centre > CSJ 2027 Application Draft). It opened Oct 27; the close date isn't published (the 2026 call closed Dec 11), so submit early", "BBG", "Cody", 14, "High"),
    ("Change the CCCL Google password now that Kyle's setup is done, and check 2-Step Verification is on", "Office", "Cody", 1, "High"),
    ("Check the morning report arrived in Gmail drafts with the laptop shut overnight", "Office", "Cody", 1, "High"),
    ("Second Chance: make its own Gmail, upload the four Second Chance files to its Drive, share the folder with the board and with ontario.cccl@gmail.com (Handoff Guide, section 4)", "Second Chance", "Cody", 7, "High"),
    ("Second Chance: check where the incorporation filing stands and confirm the board (Second Chance Tracker, Setup and Board tabs)", "Second Chance", "Cody", 7, "High"),
    ("Second Chance: ask a partner charity about leading a United Way application (due Dec 11)", "Second Chance", "Cody", dt.date(2026, 12, 1), "Normal"),
    ("Add a free first class option to the Gymdesk signup page: one class of any kind (boxing, kickboxing, jiu-jitsu, wrestling, MMA)", "BBG", "Cody", 7, "High"),
    ("Thank the 2026 sponsors and gym-wall supporters, and give them first pick for 2027", "CCCL", "Cody", 10, "Normal"),
    ("Call the top venues on the Compare tab for capacity, full event price, 2027 dates and bar terms", "CCCL", "Cody", 14, "High"),
    ("Decide if a 2027 show is adults-only (19+): the 19+ Show Check tab shows what CCCL needs from the bar", "CCCL", "Cody", 21, "High"),
    ("Ask the Ontario Athletics Commission whether CCCL needs a single-event permit ($600) or a tournament permit ($1,600) for an 8-12 bout card with youth grappling", "CCCL", "Cody", 14, "Normal"),
    ("Make the fighter sign-up Google Form (template 3 in the Fight Night Templates)", "CCCL", "Cody", 21, "Normal"),
    ("Import the 16 CCCL starter designs into Canva and save them as CCCL templates", "CCCL", "Kyle", 14, "Normal"),
    ("Sign and date the health and safety policy and the violence and harassment policy (BBG First Employee Basics)", "BBG", "Cody", 30, "Normal"),
    ("Before the summer student starts: WSIB or insurance, posters, awareness training, first aid kit (BBG First Employee Basics checklist)", "BBG", "Cody", 180, "Normal"),
    ("Check whether BBG is registered on Transfer Payment Ontario (most Ontario grants need it)", "Office", "Cody", 14, "Normal"),
    ("Paste last month's bank export into the Money Summary (delete the account-number column first) and sort the Check lines", "Office", "Cody", 14, "Normal"),
    ("Monthly backup: download the Bancroft Boxing folder from Drive to the USB drive (Workstation Plan, section 7)", "Office", "Cody", 30, "Normal"),
    ("Decide whether to register BBG with Boxing Ontario (needed for KidSport to cover kids' fees, and sets coach and insurance rules)", "BBG", "Cody", 30, "Normal"),
    ("NCCP training for the coaches who don't have it yet (the Community Coaching Grant can help)", "BBG", "Cody", 45, "Normal"),
    ("Add the kids and youth to the Skills Passport roster (first names only) and start logging skills", "BBG", "Cody", 21, "Normal"),
    ("Athlete brand: make the Facebook athlete page, switch Instagram to a Creator account, and film the first four Constrictor Tips", "Athlete brand", "Cody", 7, "Normal"),
    ("30-day check-in with Kyle: what's used, what isn't, what's annoying", "Office", "Kyle", 30, "Normal"),
]
today = dt.date(2026, 10, 20)
monday = dt.date(2026, 10, 19)
cal_start = dt.date(2026, 10, 26)   # posts start the Monday after the handoff week
for k, (t, b, o, d, p) in enumerate(starter):
    r = 5 + k
    due = d if isinstance(d, dt.date) else f'=IF({HD}="","",{HD}+{d})'
    for c, v in zip("ABCDEF", (t, b, o, due, p, "To do")): wsT[f"{c}{r}"] = v
    wsT[f"H{r}"] = f'=IF({HD}="","",{HD})'
wsT.freeze_panes = "B5"; wsT.auto_filter.ref = f"A4:{L(len(TC))}{MT}"

# ------------------------------------------------------------------ Approvals
wsA = wb.create_sheet("Approvals")
AC = [("Added", 12), ("What needs approving", 44), ("Business", 13), ("Type", 16), ("Prepared by", 12), ("Where it is", 30),
      ("Decision", 12), ("Decided on", 12), ("Days waiting", 10), ("Notes", 36)]
banner(wsA, "APPROVAL QUEUE", "Anything going out to the public, sponsors, venues or money waits here for a yes. Claude drafts; Cody (or staff, for routine gym replies) approves.", len(AC))
MA = 4 + 300
table(wsA, AC, MA, formula_cols={9: lambda r: f'=IF(OR(B{r}="",A{r}="",G{r}<>"Waiting"),"",MAX(0,TODAY()-A{r}))'})
for r in range(5, MA + 1):
    wsA[f"A{r}"].number_format = DATE; wsA[f"H{r}"].number_format = DATE; wsA[f"I{r}"].number_format = "0"
dv(wsA, BUSINESSES, f"C5:C{MA}")
dv(wsA, ["Sponsor pitch", "Sponsor deal", "Social post", "Member message", "Parent announcement", "Venue", "Money", "Contract", "Press", "Other"], f"D5:D{MA}")
dv(wsA, ["Claude", "Kyle", "Staff", "Cody"], f"E5:E{MA}"); dv(wsA, ["Waiting", "Approved", "Changes needed", "No"], f"G5:G{MA}")
dvdate(wsA, f"A5:A{MA}"); dvdate(wsA, f"H5:H{MA}")
wsA.conditional_formatting.add(f"G5:G{MA}", CellIsRule(operator="equal", formula=['"Approved"'], fill=PatternFill("solid", fgColor="C6EFCE")))
wsA.conditional_formatting.add(f"G5:G{MA}", CellIsRule(operator="equal", formula=['"Waiting"'], fill=PatternFill("solid", fgColor="FFF2CC")))
ex = [today, "(Example) First 5 sponsor pitch drafts: gym-wall sponsors", "CCCL", "Sponsor pitch", "Claude", "Gmail drafts folder", "Waiting", None, None,
      "Example row showing the format. Replace or delete it."]
for i, v in enumerate(ex, 1):
    if v is not None: wsA.cell(5, i, v)
wsA.freeze_panes = "C5"

# ------------------------------------------------------------------ Content Calendar
wsC = wb.create_sheet("Content Calendar")
CC = [("Date", 12), ("Business", 10), ("Platform", 16), ("Post type", 18), ("Draft", 60), ("Photo / video needed", 26),
      ("Status", 11), ("Approved by", 11), ("Notes", 30)]
banner(wsC, "CONTENT CALENDAR", "Planned posts for BBG and CCCL. Claude's social skill drafts them; someone approves; then post and mark Posted. "
       "No photos of kids without a parent's OK. First names only.", len(CC))
MC = 4 + 300
table(wsC, CC, MC)
for r in range(5, MC + 1): wsC[f"A{r}"].number_format = DATE
dv(wsC, ["BBG", "CCCL", "Cody"], f"B5:B{MC}"); dv(wsC, ["Facebook + Instagram", "Facebook", "Instagram", "TikTok", "YouTube"], f"C5:C{MC}")
dv(wsC, ["Class of the week", "Free trial", "Coach spotlight", "Member win", "Training clip", "Kids & parents", "Announcement",
         "Sponsor shout-out", "Fight countdown", "Fighter spotlight", "Results", "Behind the scenes", "Constrictor Tip", "Get back up"], f"D5:D{MC}")
dv(wsC, ["Idea", "Drafted", "Approved", "Posted"], f"G5:G{MC}"); dv(wsC, ["Cody", "Staff"], f"H5:H{MC}"); dvdate(wsC, f"A5:A{MC}")
wsC.conditional_formatting.add(f"G5:G{MC}", CellIsRule(operator="equal", formula=['"Posted"'], fill=PatternFill("solid", fgColor="C6EFCE")))
plan = [
    (0, "BBG", "Class of the week", "Never boxed before? Our Monday boxing class (4-5pm) starts with the basics: stance, jab, footwork. Your first class is free. Sign up at bancroft-boxing-grappling.gymdesk.com/signup or message us. 🥊",
     "Clip of the 4pm class warming up (adults only, or kids with parent OK)"),
    (2, "BBG", "Free trial", "Your first class at Bancroft Boxing & Grappling is FREE. Boxing, wrestling, jiu-jitsu, kickboxing and MMA, with boxing for kids from about 4. After that it's $20 a drop-in or $25 for the week. Sign up: bancroft-boxing-grappling.gymdesk.com/signup",
     "Wide shot of the gym"),
    (4, "BBG", "Coach spotlight", "Meet [coach first name, e.g. Sean or Jordan, who help coach wrestling]. [One line on what they coach and their background]. Catch them on [day] at [time].",
     "Photo of the coach (ask them first)"),
    (7, "BBG", "Class of the week", "Kids Boxing, Wednesdays 5:15-6:15pm. Fitness, focus and confidence, coached by people who care. First class free for new kids. Parents: sign up and do the waiver at bancroft-boxing-grappling.gymdesk.com/signup",
     "Kids class photo ONLY with parents' OK, or the empty mats and bags"),
    (9, "CCCL", "Announcement", "400 tickets and 15 local sponsors at our first show. The Cottage Country Combat League is coming back in 2027, and sponsor spots are open now. Message Cody to put your name cage-side.",
     "Best photo from CCCL 1 in Haliburton"),
    (11, "BBG", "Training clip", "Friday pad work. [One line about the drill]. Come try it: first class free.", "15-30 sec pad-work clip"),
    (14, "BBG", "Class of the week", "Women's Boxing: Wednesdays 7pm and Thursdays 6pm. A women-only class for women and girls 16+, all levels welcome. First class is on us.",
     "Photo from the women's class (ask everyone in it)"),
    (16, "BBG", "Member win", "[First name] [what they achieved: first month, first competition, a goal]. Proud of you! 👊", "Photo with permission"),
    (18, "CCCL", "Behind the scenes", "Planning CCCL 2027. Which town should we hit next: Haliburton, Bancroft, Peterborough, Lindsay or somewhere new? Drop it in the comments.", "None needed"),
    (21, "BBG", "Class of the week", "No-gi Jiu-Jitsu, Tuesdays 4-6pm. Learn to control, escape and submit. Great for self-defence. First class free.",
     "Grappling clip (adults)"),
    (23, "BBG", "Free trial", "Thinking about it? Just come try one class, free. Message us or drop by 17 Snow Road.", "Gym photo"),
    (25, "BBG", "Training clip", "Wrestling, Mondays 3:20pm. Where it all starts. [one line].", "Wrestling clip (kids only with parents' OK)"),
]
# Cody's own page: the first four weeks from the athlete brand plan (bbg-cccl-social/references/athlete-brand.md)
import os as _os, re as _re
_ab = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "skills/bbg-cccl-social/references/athlete-brand.md")
if _os.path.exists(_ab):
    _type = {"Constrictor Tip": "Constrictor Tip", "Behind the scenes": "Behind the scenes", "Get back up": "Get back up", "Shows": "Announcement"}
    for _ln in open(_ab).read().splitlines():
        _m = _re.match(r"\| (\d) \| (Mon|Wed|Fri) \| ([^|]+) \| (.+) \| ([^|]+) \|$", _ln)
        if _m:
            _wk, _day, _pil, _post, _film = _m.groups()
            plan.append((7 * (int(_wk) - 1) + {"Mon": 0, "Wed": 2, "Fri": 4}[_day], "Cody", _type.get(_pil.strip(), "Behind the scenes"), _post.strip(), _film.strip()))
    plan.sort(key=lambda x: (x[0], {"BBG": 0, "CCCL": 1, "Cody": 2}.get(x[1], 3)))
for k, (d, b, t, draft, photo) in enumerate(plan):
    r = 5 + k
    # posts start the Monday after the handoff date
    when = f'=IF({HD}="","",{HD}-WEEKDAY({HD},2)+8+{d})'
    for c, v in zip("ABCDEFG", (when, b, "Facebook + Instagram", t, draft, photo, "Drafted")): wsC[f"{c}{r}"] = v
    wsC[f"I{r}"] = "Starter draft from Kyle. Check times against Gymdesk before posting."
    wsC.row_dimensions[r].height = 58
wsC.freeze_panes = "B5"

# ------------------------------------------------------------------ Links
wsL = wb.create_sheet("Links")
banner(wsL, "THE OTHER SHEETS", "Once each file is in Google Drive, paste its link here so everyone can find it from one place.", 3)
for c, w in zip("ABC", (34, 60, 50)): wsL.column_dimensions[c].width = w
for i, h in enumerate(("File", "Google Drive link (paste it here)", "What it's for"), 1):
    c = wsL.cell(4, i, h); c.font = f_head; c.fill = fill_navy; c.border = box
files = [("CCCL Event Profit Predictor", "Compare towns and venues, break-even, stress test, 25-town scout, ticket tracker"),
         ("CCCL Sponsorship Pipeline", "2026 sponsors, 70 local prospects, national brands, stages, follow-ups, delivery checklist"),
         ("BBG Leads & Trials", "Gym inquiries to free trial to member"),
         ("CCCL Fight Night Kit", "Fighters, bout card, countdown, volunteers, production, run of show, press list, results"),
         ("CCCL Fight Night Templates", "Gym outreach, fighter sign-up form, info sheet, announcer script, press releases, promo copy"),
         ("CCCL Fighter Package", "The 4-page packet each confirmed fighter gets through their coach (fill the yellow fields)"),
         ("CCCL Marketing Kit", "Show campaign, ticket sources and tracking links, ad tests, poster drop, shot list, Canva kit"),
         ("CCCL 2027 Sponsor Packages (PDF)", "Attach to sponsor pitches"),
         ("CCCL Sponsor Pitch Templates", "Email and message templates"),
         ("Staff Playbook", "How the team uses all of this"),
         ("The Constrictor AI Blueprint", "The big plan: all 60 jobs and the build order"),
         ("CCCL Brand Kit + Canva starter designs", "Colours, fonts, logo rules and the 16 designs to import into Canva"),
         ("Handoff folder", "Cody's phone cheat sheet, the Handoff Guide and the skill tests"),
         ("CSJ 2027 Application Draft", "A funded summer student: steps before applying, answers to paste in, the job posting"),
         ("BBG First Employee Basics", "Safety plan, harassment policy, posters, payroll and WSIB before the first hire"),
         ("The Constrictor Athlete Brand Plan", "Cody's own Facebook and Instagram: post types, weekly routine, first four weeks, bios"),
         ("BBG & CCCL Contacts", "Every venue, gym, media outlet, supplier and official, with the contact log"),
         ("BBG Coaching Library", "Class profiles, drills, 4-week lesson plans and the safety rules"),
         ("BBG Skills Passport", "Kids' and youth levels, the roster, progress log and parent cards"),
         ("BBG & CCCL Money Summary", "Bank, Eventbrite and Gymdesk exports sorted by business; monthly totals"),
         ("The Constrictor Funding Tracker", "Grants and programs, eligibility, deadlines and applications")]
for k, (f, w) in enumerate(files):
    r = 5 + k
    wsL.cell(r, 1, f).font = f_bold; c = wsL.cell(r, 2); c.font = f_in; c.fill = fill_in; wsL.cell(r, 3, w).font = f_lab
    for i in (1, 2, 3): wsL.cell(r, i).border = box

# ------------------------------------------------------------------ Dashboard
wsD = wb.create_sheet("Dashboard", 0)
banner(wsD, "THE CONSTRICTOR HQ", "Cody's one screen. Open it every morning (it works on a phone). Tasks, approvals and posts update by themselves; "
       "the weekly numbers are filled in from Monday's report.", 4)
for c, w in zip("ABCD", (40, 22, 4, 48)): wsD.column_dimensions[c].width = w
r = 4
wsD.cell(r, 1, "Today").font = f_lab; c = wsD.cell(r, 2, "=TODAY()"); c.number_format = "ddd mmm d, yyyy"; c.font = f_bold
r += 2
TT = lambda col: f"Tasks!${col}$5:${col}${MT}"
AA = lambda col: f"Approvals!${col}$5:${col}${MA}"
CCr = lambda col: f"'Content Calendar'!${col}$5:${col}${MC}"
def stat(label, fml, fmt="0", warn=False, note=None):
    global r
    wsD.cell(r, 1, label).font = f_lab; wsD.cell(r, 1).border = box
    c = wsD.cell(r, 2, fml); c.font = Font(name=F, size=12, bold=True); c.number_format = fmt; c.border = box
    c.alignment = Alignment(horizontal="center")
    if warn: wsD.conditional_formatting.add(f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor="F8CBAD")))
    if note: n = wsD.cell(r, 4, note); n.font = f_note; n.alignment = wrap
    r += 1
def sec(t):
    global r
    wsD.cell(r, 1, t).font = f_sec; r += 1
sec("NEEDS CODY NOW")
stat("Approvals waiting", f'=COUNTIFS({AA("G")},"Waiting",{AA("B")},"<>")', warn=True)
stat("Oldest approval waiting (days)", f'=IFERROR(_xlfn.MAXIFS({AA("I")},{AA("G")},"Waiting"),0)', warn=True)
stat("Tasks overdue", f'=COUNTIF({TT("G")},"OVERDUE")', warn=True)
stat("Tasks due today", f'=COUNTIF({TT("G")},"TODAY")')
r += 1
sec("TEAM TASKS")
stat("Open tasks", f'=COUNTIFS({TT("A")},"<>",{TT("F")},"<>Done")')
stat("Due in the next 7 days", f'=COUNTIFS({TT("F")},"<>Done",{TT("D")},">="&TODAY(),{TT("D")},"<="&(TODAY()+7))')
stat("Done in the last 7 days", f'=COUNTIFS({TT("F")},"Done",{TT("I")},">="&(TODAY()-7))')
for o in OWNERS:
    stat(f"   Open for {o}", f'=COUNTIFS({TT("C")},"{o}",{TT("F")},"<>Done",{TT("A")},"<>")')
r += 1
sec("POSTS THIS WEEK (MON-SUN)")
wk0 = "(TODAY()-WEEKDAY(TODAY(),3))"
for s in ("Idea", "Drafted", "Approved", "Posted"):
    stat(f"   {s}", f'=COUNTIFS({CCr("A")},">="&{wk0},{CCr("A")},"<="&({wk0}+6),{CCr("G")},"{s}")')
r += 1
sec("WEEKLY NUMBERS (from Monday's report)")
wsD.cell(r, 4, "Claude's Monday report lists these. Type them in, or have Claude update this sheet if the Google Sheets connector is on.").font = f_note
wsD.cell(r, 4).alignment = wrap; wsD.row_dimensions[r].height = 26
for lab, fmt in (("Week of", "mmm d"), ("Sponsorship paid so far ($)", '$#,##0'), ("Sponsorship signed, not paid ($)", '$#,##0'),
                 ("Sponsor follow-ups due", "0"), ("Gym inquiries last week", "0"), ("Free trials last week", "0"),
                 ("New members last week", "0"), ("Tickets sold for next show", "0"), ("Tickets needed to break even", "0")):
    wsD.cell(r, 1, lab).font = f_lab; wsD.cell(r, 1).border = box
    c = wsD.cell(r, 2); c.font = Font(name=F, size=12, bold=True, color="0000FF"); c.fill = fill_in; c.border = box; c.number_format = fmt
    c.alignment = Alignment(horizontal="center"); r += 1
wsD.sheet_view.showGridLines = False

# ------------------------------------------------------------------ Start Here
ws0 = wb.create_sheet("Start Here", 0)
banner(ws0, "CONSTRICTOR HQ: START HERE", "Cody's command centre for BBG and CCCL. Built by Kyle. Opens in Google Sheets or Excel.", 3)
ws0.column_dimensions["A"].width = 4; ws0.column_dimensions["B"].width = 24; ws0.column_dimensions["C"].width = 92
rr = 4
for a, b in (("Dashboard", "The one screen: approvals waiting, overdue tasks, this week's posts, the weekly numbers."),
             ("Tasks", "Every job for the team. Owner, due date, status. Overdue ones turn red on the Dashboard."),
             ("Approvals", "Everything waiting for a yes before it goes out: sponsor pitches, posts, member messages, money."),
             ("Content Calendar", "Planned posts. The first four weeks are drafted. Check class times before posting."),
             ("Links", "Paste the Google Drive link of each other sheet here."),
             ("Every morning", "Claude's morning report reads this sheet plus Gmail, Calendar, the Sponsorship Pipeline and the Leads sheet, and tells Cody what matters today."),
             ("Colour key", "Yellow cells with blue text are for typing. Red text and black text are calculated. Rows marked (Example) show the format; delete them.")):
    ws0.cell(rr, 2, a).font = f_bold; ws0.cell(rr, 2).alignment = top
    c = ws0.cell(rr, 3, b); c.font = f_lab; c.alignment = wrap; rr += 1
assert rr == 11
ws0.cell(12, 2, "Handoff date").font = f_bold
hdc = ws0.cell(12, 3); hdc.font = Font(name=F, size=12, bold=True, color="0000FF"); hdc.fill = fill_in; hdc.border = box; hdc.number_format = "ddd mmm d, yyyy"
dvdate(ws0, "C12")
n = ws0.cell(13, 3, "Kyle types the sit-down date here before upload. The starter tasks' due dates count from it, and the first four weeks of posts start the Monday after.")
n.font = f_lab; n.alignment = wrap
ws0.sheet_view.showGridLines = False
for w in wb.worksheets:
    w.sheet_properties.tabColor = GOLD if w.title in ("Start Here", "Dashboard") else NAVY
    w.page_setup.orientation = "landscape"; w.page_setup.fitToWidth = 1; w.page_setup.fitToHeight = 0
    w.sheet_properties.pageSetUpPr.fitToPage = True
wb.active = 1
wb.save(OUT); print("saved", OUT)
