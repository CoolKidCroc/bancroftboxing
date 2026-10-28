// The Constrictor Handoff Guide (Word / Google Docs).
// usage: NODE_PATH=$(npm root -g) node guide.js OUT.docx BBG_LOGO CCCL_LOGO
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, Footer, PageNumber, ShadingType, LevelFormat, ImageRun,
  Table, TableRow, TableCell, WidthType, VerticalAlign, TableLayoutType } = require("docx");
const OUT = process.argv[2], BBG = process.argv[3], CCCL = process.argv[4];
const NAVY = "081C2B", GOLD = "C8862E", GOLDL = "FBEBD3", PALE = "F2F4F6", FONT = "Arial", W = 9720;

const R = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 20, ...o });
const runs = (x) => (Array.isArray(x) ? x : [x]).flatMap(y => typeof y === "string" ? [R(y)] : (Array.isArray(y) ? y : [y]));
const P = (x, o = {}) => new Paragraph({ spacing: { after: 90 }, ...o, children: runs(x) });
const Bd = (t, o = {}) => R(t, { bold: true, ...o });
let SEC = 0;
const H1 = (t, sub, brk) => { SEC++; return [new Paragraph({ keepNext: true, pageBreakBefore: !!brk, spacing: { before: 280, after: sub ? 20 : 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: GOLD, space: 3 } },
    children: [new TextRun({ text: `${SEC}. ${t}`.toUpperCase(), font: FONT, size: 27, bold: true, color: NAVY, characterSpacing: 10 })] }),
  ...(sub ? [new Paragraph({ keepNext: true, spacing: { after: 110 }, children: [R(sub, { italics: true, color: "5A6672", size: 19 })] })] : [])]; };
const H2 = (t) => new Paragraph({ keepNext: true, spacing: { before: 160, after: 60 }, children: [R(t, { bold: true, size: 22, color: GOLD })] });
const bullet = (x) => new Paragraph({ numbering: { reference: "b", level: 0 }, spacing: { after: 45 }, children: runs(x) });
const chk = (x) => new Paragraph({ keepNext: true, spacing: { after: 60 }, indent: { left: 360, hanging: 360 }, children: [new TextRun({ text: "☐  ", font: "Segoe UI Symbol", size: 22 }), ...runs(x)] });

const line = { style: BorderStyle.SINGLE, size: 4, color: "D7DDE3" };
const borders = { top: line, bottom: line, left: line, right: line };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const cell = (x, w, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: o.borders || borders, verticalAlign: o.v || VerticalAlign.CENTER,
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, margins: { top: 60, bottom: 60, left: 110, right: 110 },
  children: [new Paragraph({ spacing: { after: 20 }, children: runs(x).map(r => r) })] });
function table(cols, rows, o = {}) {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: cols, layout: TableLayoutType.FIXED,
    rows: rows.map((r, i) => new TableRow({ cantSplit: true, tableHeader: i === 0,
      children: r.map((c, j) => i === 0 ? cell([Bd(c, { color: "FFFFFF", size: 18 })], cols[j], { fill: NAVY })
        : cell(c === "☐" ? [new TextRun({ text: "☐", font: "Segoe UI Symbol", size: 22 })] : (typeof c === "string" && j === 0 && o.label ? [Bd(c, { size: 19 })] : c), cols[j], { fill: o.label && j === 0 ? GOLDL : (i % 2 === 0 ? PALE : undefined) })) })) });
}
const callout = (title, lines) => new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], layout: TableLayoutType.FIXED,
  rows: [new TableRow({ cantSplit: true, children: [new TableCell({ width: { size: W, type: WidthType.DXA }, shading: { fill: GOLDL, type: ShadingType.CLEAR, color: "auto" },
    borders: { top: none, bottom: none, right: none, left: { style: BorderStyle.SINGLE, size: 24, color: GOLD } }, margins: { top: 110, bottom: 110, left: 200, right: 160 },
    children: [new Paragraph({ spacing: { after: 50 }, children: [Bd(title, { size: 21, color: NAVY })] }), ...lines.map(l => P(l, { spacing: { after: 50 } }))] })] })] });
const sp = (a = 100) => new Paragraph({ spacing: { after: a }, children: [] });

const cover = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [6720, 3000], layout: TableLayoutType.FIXED,
  rows: [new TableRow({ children: [
    new TableCell({ width: { size: 6720, type: WidthType.DXA }, shading: { fill: NAVY, type: ShadingType.CLEAR, color: "auto" }, borders: { top: none, bottom: none, left: none, right: none },
      margins: { top: 300, bottom: 280, left: 340, right: 120 }, verticalAlign: VerticalAlign.CENTER, children: [
        new Paragraph({ spacing: { after: 40 }, children: [R("THE CONSTRICTOR AI", { size: 20, bold: true, color: "EAA959", characterSpacing: 40 })] }),
        new Paragraph({ spacing: { after: 60 }, children: [R("HANDOFF GUIDE", { font: "Arial Black", size: 50, bold: true, color: "F4F1EA" })] }),
        new Paragraph({ children: [R("How Kyle hands everything to Cody, and how it keeps running after.", { size: 20, color: "C9D3DC" })] })] }),
    new TableCell({ width: { size: 3000, type: WidthType.DXA }, shading: { fill: NAVY, type: ShadingType.CLEAR, color: "auto" }, borders: { top: none, bottom: none, left: none, right: none },
      margins: { top: 160, bottom: 160, left: 100, right: 260 }, verticalAlign: VerticalAlign.CENTER, children: [
        new Paragraph({ alignment: AlignmentType.RIGHT, children: [
          new ImageRun({ type: "png", data: fs.readFileSync(BBG), transformation: { width: 84, height: 84 } }), R("  "),
          new ImageRun({ type: "png", data: fs.readFileSync(CCCL), transformation: { width: 84, height: 84 } })] })] }),
  ] })] });

const children = [
  cover, sp(140),
  P([Bd("Who this is for: "), "Kyle, to set things up and walk Cody through them; Cody, to keep as the manual. Read it once start to finish (about 15 minutes), then use the checklists."]),

  ...H1("What Cody is getting", "Everything is in the Bancroft Boxing folder, with the same layout in Google Drive. Second Chance is the exception: it goes in its own Drive (section 4)."),
  table([3000, 4320, 2400], [
    ["Piece", "What it's for", "Folder"],
    ["START HERE plan", "The big picture: what's built, the order, what's still needed", "Bancroft Boxing (top)"],
    ["The Constrictor AI Blueprint", "A one-page map of every job the AI does", "Bancroft Boxing (top)"],
    ["9 Claude skills + install guide", "The jobs Claude does for Cody (front desk, social and his own page, sponsors, shows, fight night, HQ, coaching, funding, Second Chance)", "The Constrictor (Claude skills)"],
    ["Constrictor HQ", "Dashboard, tasks, approvals, content calendar. The morning report reads it", "Command Centre"],
    ["BBG & CCCL Contacts", "Every venue, gym, media outlet, supplier and official, with a log of every call and email", "Command Centre"],
    ["BBG Coaching Library", "Class profiles, 63 drills, 4-week lesson plans that check time, age, contact and all-ages groups, the gym's equipment and the safety rules", "BBG Coaching"],
    ["BBG Skills Passport", "Four levels of skills for each kids' and youth class, a roster (first names only) and a progress card for parents", "BBG Coaching"],
    ["BBG & CCCL Money Summary", "Paste the bank, Eventbrite and Gymdesk exports; it sorts every line and totals each month by business", "Command Centre"],
    ["The Constrictor Funding Tracker", "Grants and programs that fit, who can apply, deadlines and where each application stands (re-checked with Cody's answers)", "Command Centre"],
    ["CSJ 2027 Application Draft + BBG First Employee Basics", "A funded summer student: what to do before applying, the answers to paste in, the job posting, and the safety and harassment policies a first employer needs", "Command Centre"],
    ["Second Chance Starter Plan, Tracker, Policy Drafts, First Letters", "Setting up the nonprofit: the board, four policies for a lawyer to review, five letters ready to send, partners, placements, participants (first names only) and funding", "Second Chance (not uploaded with the rest)"],
    ["Athlete Brand Plan", "Cody's own Facebook and Instagram in 1 to 3 hours a week", "Athlete Brand"],
    ["BBG Leads & Trials", "Every gym inquiry, the free first class, follow-ups", "BBG Members & Leads"],
    ["Sponsorship Pipeline, packages, pitch templates", "Finding, pitching and tracking sponsors. Only Paid counts", "CCCL / Sponsorships"],
    ["Event Profit Predictor", "Where the next show goes and whether it makes money", "CCCL"],
    ["Fight Night Kit, Templates, Fighter Package", "Fighters, the bout card, the night, press, the report", "CCCL / Fight Night"],
    ["Marketing Kit + 16 Canva designs", "The show campaign, tracking links, ads, livestream, graphics", "CCCL / Marketing"],
    ["Logos + CCCL Brand Kit", "Colours, fonts and how to use the designs in Canva", "Brand"],
    ["Staff Playbook", "Copy-paste prompts for staff who don't have the skills", "Bancroft Boxing (top)"],
    ["Workstation Plan", "Setting up Cody's laptop: what runs where, what each skill can touch, the cloud morning report, security, backups and costs", "Handoff"],
    ["Cheat sheet, skill tests, this guide", "The handoff", "Handoff"],
  ], { label: true }),

  ...H1("Before the sit-down", "Kyle, about 1 to 2 hours. Do these first so the sit-down is all show-and-tell."),
  chk([Bd("Sign in to the CCCL Google account "), "(ontario.cccl@gmail.com). Cody made it and keeps the password; it never goes in a file. Sponsor pitches, venue emails and press releases come from it."]),
  chk([Bd("Upload the folder to Google Drive "), "in the CCCL account, leaving out the Second Chance folder (Cody sets that up himself; section 4) and Kyle's own _Kyle (don't upload) folder. Turn on Drive's Settings > Convert uploads first, so every .xlsx becomes a Google Sheet. If one didn't convert: open it, File > Save as Google Sheets, then delete the .xlsx so there's only one copy."]),
  chk([Bd("Open every sheet's Start Here tab "), "and clear any example rows so Cody starts clean. BBG Leads & Trials has no Start Here tab: its example row is on the Leads tab."]),
  chk([Bd("Constrictor HQ > Start Here: type the sit-down date "), "as the Handoff date. The starter tasks and the first four weeks of posts count from it. Delete any starter task Cody has already done."]),
  chk([Bd("Run the skill tests "), "on your own Claude (The Constrictor Skill Tests sheet, about an hour). The Summary tab must say READY FOR CODY."]),
  chk([Bd("Send Cody the cheat sheet "), "(the 9 phone images) so he can save them to his camera roll."]),
  chk([Bd("Book 90 minutes "), "with Cody at a computer, with his phone."]),

  ...H1("The sit-down with Cody", "About 90 minutes. Cody does the typing; Kyle points."),
  table([1250, 3000, 4570, 900], [
    ["Time", "What", "How", "Done"],
    ["0:00", "The big picture", "Open the Blueprint. Go over the 3 golden rules on cheat sheet card 1: drafts only, Cody and the officials decide bouts, only Paid is money.", "☐"],
    ["0:10", "Claude on Cody's account", "Pro plan. Settings > Capabilities > Code execution on. Upload the 9 skills (install guide). Connect Google Drive, Gmail and Calendar with ontario.cccl@gmail.com.", "☐"],
    ["0:25", "Morning report", "Say \"Morning report\". Then set it up as a Cloud scheduled task for weekdays at 7:05 (Workstation Plan, section 4), so it runs with the laptop shut. Click Run now once to test.", "☐"],
    ["0:35", "The gym", "Paste a real Facebook message: \"Reply to this: …\". Then \"Plan Wednesday's kids boxing\" and open the Coaching Library.", "☐"],
    ["0:50", "Sponsors", "\"How are sponsors going?\" Then draft one real pitch, find it in Gmail drafts, and Cody sends it himself.", "☐"],
    ["1:05", "Shows", "Open the Predictor and the Fight Night Kit. Ask \"What's left for the next show?\" Import one Canva design (Create a design > Import file).", "☐"],
    ["1:15", "What's new", "Open the CSJ draft, the Second Chance Starter Plan and the Athlete Brand Plan. Say \"Posts for my page this week\". Book a separate hour for Second Chance.", "☐"],
    ["1:25", "Sharing", "Set sharing (section 4: only Cody for now). Put the 30-day check-in on the calendar.", "☐"],
  ]),
  sp(60),
  callout("If something doesn't work during the sit-down", ["Don't fix it live. Write it on the HQ Tasks tab (Owner = Kyle) and move on. Section 7 has the usual fixes."]),

  ...H1("Google Drive: who sees what", "Share the least that lets each person do their job. Sheets can't hide single tabs, so share whole files carefully."),
  P([Bd("Cody's choice (Oct 20, 2026): only Cody sees the folder for now. "), "Kyle keeps editor access (on his own Google account) to maintain it until Cody removes him. The right-hand column is what to share when staff start using the sheets in week 2 (section 5); everything else reaches staff by text or printout."]),
  table([3300, 1200, 1500, 3720], [
    ["File or folder", "Cody", "Kyle", "Staff, from week 2"],
    ["Whole Bancroft Boxing folder", "Owner", "Editor", "Nobody gets the whole folder"],
    ["Constrictor HQ, Contacts, Money Summary, Funding Tracker", "Owner", "Editor", "Nobody"],
    ["Sponsorship Pipeline, Predictor, Fight Night Kit, Marketing Kit", "Owner", "Editor", "Nobody"],
    ["BBG Leads & Trials", "Owner", "Editor", "Front desk: Editor"],
    ["BBG Coaching Library, Skills Passport", "Owner", "Editor", "Coaches: Editor"],
    ["Brand folder, Canva designs, Staff Playbook", "Owner", "Editor", "Staff: Viewer"],
    ["Second Chance folder (its own Drive)", "Owner", "No access", "The Second Chance board only. Never gym staff"],
  ]),
  sp(60),
  bullet([Bd("Staff get tasks by text, "), "not by opening HQ: Claude drafts the message, Cody sends it."]),
  bullet([Bd("Volunteers and the fight-night crew "), "get a printout or PDF of the Run of Show and the volunteer briefing, not the Fight Night Kit."]),
  bullet([Bd("Second Chance "), "has its own folder, email and bank account, and Cody sets it up himself after the sit-down (about 20 minutes): make a Gmail for Second Chance, upload the four Second Chance files to its Drive (Kyle emails them to that address), and share the folder with the board. So Claude can read the Tracker, also share it with ontario.cccl@gmail.com, and change the CCCL password first if anyone else (Kyle during setup) has signed into that account. Participants' details are never shared with gym staff and never linked to these sheets."]),
  bullet([Bd("Cody's personal email "), "stays out of every shared file. Sponsors, venues and press get ontario.cccl@gmail.com."]),
  bullet([Bd("Never in any sheet: "), "medical details (Yes / No / Pending only), card numbers, personal phone numbers, kids' last names."]),

  ...H1("The first 30 days", "Small steps. Each week adds one habit."),
  table([1300, 4420, 4000], [
    ["Week", "Cody", "Kyle"],
    ["1", "Read the morning report every day. Say \"Go through my approvals\" once a day.", "Watch for anything wrong (a price, a name, a missing sheet) and fix the skill's facts."],
    ["2", "Share Leads & Trials with the front desk and the Coaching Library and Skills Passport with the coaches (section 4). Front desk starts drafting gym replies from the Staff Playbook; Cody approves. Log every inquiry in Leads & Trials. Coaches start using the lesson plans.", "Check the Leads sheet is being filled in and replies go out the same day. Ask the coaches which drills ran long."],
    ["3", "Mondays: \"Do my follow-ups\" and \"How are sponsors going?\"", "Tidy the Sponsorship Pipeline with Cody: stages, amounts, Paid."],
    ["4", "30-minute check-in with Kyle: what's used, what isn't, what's annoying.", "Update the skills, re-run the tests for any skill you changed, send new zips."],
  ]),

  ...H1("Keeping it running", "Most upkeep is Kyle's. Cody's job is to say when something changes."),
  table([2400, 5020, 2300], [
    ["When", "What", "Who"],
    ["Prices, classes or the schedule change", "Update gym-facts.md in bbg-front-desk AND its copy in bbg-cccl-social, the Classes tab of the Coaching Library (and classes.md in bbg-coach), and the Staff Playbook's facts block; re-zip, and Cody re-uploads the skills (delete the old ones first).", "Cody tells Kyle; Kyle updates"],
    ["Sponsor packages change", "Prices are written in several places: the Pipeline's Packages tab, packages.md, cccl-facts.md and pitch-templates.md (all in cccl-sponsor-desk), the package line in that skill's SKILL.md, the Sponsor Pitch Templates doc and the Sponsor Packages PDF. Update them all, then re-zip.", "Kyle"],
    ["A new show is booked", "Make a copy of the Fight Night Kit and Marketing Kit for the show (File > Make a copy, \"CCCL [Town] [date]\"), fill in Show Setup, then say \"Plan the campaign for [Town] on [date]\".", "Cody and Kyle"],
    ["After every show", "\"Do the report\". Post-event report to the Commission within 14 days. Copy the real numbers into the Predictor.", "Cody"],
    ["Every Monday", "Monday report, sponsor follow-ups, posts for next week.", "Cody"],
    ["Start of every month", "Paste last month's bank export (account-number column deleted) into the Money Summary, sort the Check lines, then ask \"how did we do in [month]?\"", "Cody"],
    ["First week of every month", "Ask \"anything new out there?\" (opportunity scan) and \"is everything being kept up?\" (system check). Fix anything stale.", "Cody with Kyle"],
    ["Every 3 months", "Re-run the skill tests (or after any skill change). Check the Staff Playbook still matches.", "Kyle"],
    ["Every year (before the season)", "Check the Ontario rules, Eventbrite fees, Canva price and sponsor packages, and refresh the funding research. Update ontario-rules.md, the Predictor and the Funding Tracker.", "Kyle"],
  ]),

  ...H1("When something goes wrong"),
  table([3400, 6320], [
    ["Problem", "Fix"],
    ["Claude doesn't use the skill", "Name it: \"Using the sponsor desk, …\". Check the skill is switched on (Customize > Skills) and Code execution is on (Settings > Capabilities)."],
    ["\"I can't find the sheet\"", "Google Drive must be connected with the CCCL account, the file must be a Google Sheet (not .xlsx), and the name must match. Renamed it? Tell Claude the new name."],
    ["A number looks wrong", "Ask \"Where did that number come from?\" Claude only reads the sheets, so fix the sheet. Only Paid counts as money."],
    ["Claude makes something up (a price, a date)", "Correct it in the chat, then tell Kyle so the skill's facts get fixed."],
    ["No Gmail drafts", "Gmail must be connected with ontario.cccl@gmail.com. Drafts sit in the Drafts folder until Cody sends them."],
    ["The morning report didn't come", "Open Scheduled in Claude. Check the task is on and is a Cloud task (a Local one stops while the laptop sleeps). Its History says why a run failed. Saying \"Morning report\" in a chat always works."],
    ["A formula shows #REF! or #N/A", "Someone deleted a row or column. File > Version history, restore yesterday's version."],
    ["An old answer after an update", "Delete the old skill before uploading the new zip, then start a new chat."],
  ]),

  ...H1("Sign-off checklist"),
  chk("Signed in to ontario.cccl@gmail.com; Cody has the password"),
  chk("Folder in Google Drive (without Second Chance), every sheet converted to Google Sheets, sharing set (section 4)"),
  chk("Claude Pro on Cody's account, Code execution on"),
  chk("All 9 skills uploaded and switched on"),
  chk("Google Drive, Gmail and Calendar connected with ontario.cccl@gmail.com"),
  chk("Skill tests: READY FOR CODY"),
  chk("Morning report set up as a Cloud scheduled task (weekdays 7:05) and tested with the laptop shut"),
  chk("Cheat sheet saved on Cody's phone"),
  chk("Staff Playbook with the front desk, Sean and Jordan"),
  chk("30-day check-in on the calendar"),
  sp(160),
  table([3240, 3240, 3240], [["Cody", "Kyle", "Date"], ["", "", ""]]),
];

const doc = new Document({
  creator: "Kyle", title: "The Constrictor Handoff Guide",
  numbering: { config: [{ reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 240 } } } }] }] },
  styles: { default: { document: { run: { font: FONT, size: 20 } } } },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 900, bottom: 900, left: 1260, right: 1260 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      R("The Constrictor handoff guide  ·  page ", { size: 16, color: "888888" }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: "888888" })] })] }) },
    children }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("saved", OUT); });
