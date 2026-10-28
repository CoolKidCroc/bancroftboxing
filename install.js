const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, LevelFormat, Table, TableRow, TableCell, WidthType, ShadingType } = require("docx");
const OUT = process.argv[2];
const NAVY = "081C2B", GOLD = "C8862E", FONT = "Arial";
const T = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 21, ...o });
const P = (runs, o = {}) => new Paragraph({ spacing: { after: 90 }, ...o, children: (Array.isArray(runs) ? runs : [runs]).map(x => typeof x === "string" ? T(x) : x) });
const H = (t) => { SEC++; return new Paragraph({ spacing: { before: 200, after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: GOLD, space: 2 } },
  children: [new TextRun({ text: t, font: FONT, size: 26, bold: true, color: NAVY })] }); };
let SEC = 0;
const step = (runs) => new Paragraph({ numbering: { reference: "num" + SEC, level: 0 }, spacing: { after: 80 }, children: (Array.isArray(runs) ? runs : [runs]).map(x => typeof x === "string" ? T(x) : x) });
const W = 9720;
const cell = (t, w, head) => new TableCell({ width: { size: w, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 100, right: 100 },
  shading: head ? { type: ShadingType.CLEAR, color: "auto", fill: NAVY } : undefined,
  children: [new Paragraph({ children: [new TextRun({ text: t, font: FONT, size: 19, bold: head, color: head ? "FFFFFF" : "000000" })] })] });
const tbl = (widths, rows) => new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths,
  rows: rows.map((r, i) => new TableRow({ cantSplit: true, children: r.map((t, j) => cell(t, widths[j], i === 0)) })) });

const children = [
  new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "ADD THE CONSTRICTOR TO YOUR CLAUDE", font: FONT, size: 34, bold: true, color: NAVY })] }),
  new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "About 10 minutes, once. Do it on a computer.", font: FONT, size: 22, bold: true, color: GOLD })] }),
  P("These are skills: saved jobs Claude knows how to do for Cody's businesses. After this, Cody just asks in plain words, like \"draft pitches for the first five businesses on the sponsor list\" or \"reply to this message\" (paste it in), and Claude uses the right skill."),

  H("What's in the folder"),
  tbl([2600, 4520, 2600], [
    ["Skill file", "What it does", "Try saying"],
    ["cccl-sponsor-desk.zip", "Finds sponsors, drafts pitches and follow-ups, tracks who's paid, preps sponsor deliverables for each show.", "\"Do my sponsor follow-ups\""],
    ["bbg-front-desk.zip", "Drafts replies to gym inquiries with the free class and the signup link, birthday parties, check-ins for members away a month, announcements, payment reminders (\"who owes us?\") and answers about help paying for kids' classes.", "\"Reply to this message: …\""],
    ["cccl-event-scout.zip", "Compares towns and venues for the next show, checks date clashes, runs the break-even numbers and tracks ticket sales.", "\"Compare Peterborough and Lindsay for April\""],
    ["constrictor-hq.zip", "Morning report, tasks and delegation, the approval queue, the inbox sweep (sorts email and leaves reply drafts), contact follow-ups, the monthly money summary, a monthly opportunity scan and system check, and turning a voice note into tasks and drafts.", "\"Morning report\" or \"Sort my inbox\""],
    ["constrictor-funding.zip", "Finds grants and programs that fit BBG, CCCL, Second Chance and gym families, checks eligibility, tracks deadlines and drafts applications for Cody to review.", "\"What funding fits our kids programs?\""],
    ["bbg-coach.zip", "Lesson plans and 4-week blocks for every class, the drill library, sub-coach sheets, whether a drill is OK for an age, parent updates and Skills Passport progress cards.", "\"Plan Wednesday's kids boxing\""],
    ["bbg-cccl-social.zip", "Posts for BBG, CCCL and Cody's own page: weekly gym posts, Cody's \"Get back up\" posts and bio, the 16-week show campaign, spotlights, sponsor shout-outs, text for the 16 Canva starter designs, and whether an ad is worth it.", "\"Plan the campaign for the Peterborough show\""],
    ["cccl-fight-night.zip", "Fighters and the bout card, gym outreach, the Fighter Package for each fighter, the run of show, announcer script, press releases and the report after a show.", "\"Message the gyms for the Peterborough show\""],
    ["second-chance.zip", "Second Chance, Cody's nonprofit: setup steps, the board, partner letters, policy drafts, safe placements and participant check-ins. Kept apart from BBG and CCCL.", "\"Where are we with Second Chance?\""],
  ]),

  H("1. Turn on the setting skills need"),
  step([T("In Claude, open "), T("Settings → Capabilities", { bold: true }), T(".")]),
  step([T("Turn on "), T("Code execution and file creation", { bold: true }), T(". Skills don't run without it.")]),

  H("2. Upload each skill"),
  step([T("Open "), T("Customize → Skills", { bold: true }), T(".")]),
  step([T("Click "), T("+", { bold: true }), T(", then "), T("Create skill", { bold: true }), T(", then "), T("Upload a skill", { bold: true }), T(".")]),
  step([T("Pick one of the .zip files from the "), T("The Constrictor (Claude skills)", { bold: true }), T(" folder. Don't unzip it first.")]),
  step("Make sure the skill's switch is on. Repeat for the other zip files."),

  H("3. Connect Google"),
  step([T("The CCCL Google account is ", { bold: true }), T("ontario.cccl@gmail.com", { bold: true }), T(". Cody made it and keeps the password (never in a file). Sponsor pitches, venue emails and press releases come from it, and Claude's drafts sit in its Gmail.")]),
  step([T("In Claude's settings, find "), T("Connectors", { bold: true }), T(" and connect "), T("Google Drive, Gmail and Google Calendar", { bold: true }), T(" signed in as ontario.cccl@gmail.com.")]),
  step("Put the sheets from this folder in Google Drive (not the Second Chance or _Kyle folders): upload them, then open each with Google Sheets (File → Save as Google Sheets)."),
  step([T("Second Chance is the exception: ", { bold: true }), T("its files go in their own Drive, owned by a Second Chance Gmail, never inside the shared Bancroft Boxing folder. Cody sets it up himself after the sit-down and shares the folder with the board, and with ontario.cccl@gmail.com so Claude can read the Tracker. Change the CCCL password first if anyone else has signed into that account (see the Handoff Guide, section 4).")]),
  step("Claude reads the sheets and writes Gmail drafts. Gmail asks before anything is sent, and Cody sends every message himself."),

  H("4. Test it"),
  P([T("Open a new chat and say: "), T("\"Using the sponsor desk, how are sponsors going?\"", { bold: true }), T(" Claude should open the CCCL Sponsorship Pipeline sheet and answer from it. If it can't find the sheet, check step 3.")]),
  P([T("Before Cody gets the skills, ", { bold: true }), T("Kyle runs the full test list in "), T("Handoff / The Constrictor Skill Tests", { bold: true }), T(" (58 tests, about an hour). The Summary tab says READY FOR CODY when every test passes.")]),

  H("5. Get the morning report every day"),
  P([T("Cody's Pro plan includes scheduled tasks. In Claude, open "), T("Scheduled", { bold: true }), T(" (in some accounts, Routines) > "), T("New task", { bold: true }), T(", and pick "), T("Cloud", { bold: true }), T(" if it asks. Paste the instructions from the Workstation Plan (Handoff folder, section 4), set weekdays at 7:05, keep only Google Drive, Gmail and Calendar, then click Run now to test. A Cloud task runs with the laptop shut; a Local one only runs while the laptop is awake with the Claude app open.")]),
  H("Good to know"),
  P([T("Skills belong to one Claude account. ", { bold: true }), T("On Pro they can't be shared, so the front desk, Sean and Jordan use the Staff Playbook prompts instead (they work in free ChatGPT or Claude). Anyone who wants the skills can upload the same zip files to their own Claude.")]),
  P([T("Updating a skill: ", { bold: true }), T("when Kyle sends a new zip, turn off or delete the old skill in Customize → Skills, then upload the new one the same way.")]),
  P([T("The schedule and prices ", { bold: true }), T("the Front Desk uses are inside bbg-front-desk (gym-facts.md), with an identical copy in bbg-cccl-social. When they change, tell Kyle (he updates both), or tell Claude in the chat. It will use what you say for that conversation.")]),
];
const doc = new Document({
  numbering: { config: [1,2,3,4,5,6,7].map(i => ({ reference: "num" + i, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 300 } } } }] })) },
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 760, bottom: 700, left: 1260, right: 1260 } } }, children }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("saved", OUT); });
