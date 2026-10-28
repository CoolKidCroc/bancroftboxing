const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType,
  ShadingType, AlignmentType, HeadingLevel, LevelFormat, BorderStyle, Footer, PageNumber,
} = require("docx");

const OUT = process.argv[2];
const LOGO = process.argv[3];
const NAVY = "081C2B", GOLD = "C8862E", GOLDL = "FBEBD3";
const FONT = "Arial";
const W = 9360; // content width, US Letter with 1" margins

const p = (text, opts = {}) => new Paragraph({
  spacing: { after: 120 }, ...opts.para,
  children: (Array.isArray(text) ? text : [text]).map(t => typeof t === "string" ? new TextRun({ text: t, font: FONT, size: 21, ...opts.run }) : t),
});
const b = (t) => new TextRun({ text: t, font: FONT, size: 21, bold: true });
const r = (t) => new TextRun({ text: t, font: FONT, size: 21 });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 120 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: GOLD, space: 2 } },
  children: [new TextRun({ text: t, font: FONT, size: 28, bold: true, color: NAVY })] });
const bullet = (parts) => new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 60 },
  children: (Array.isArray(parts) ? parts : [r(parts)]) });
const check = (parts) => new Paragraph({ numbering: { reference: "box", level: 0 }, spacing: { after: 60 },
  children: (Array.isArray(parts) ? parts : [r(parts)]) });

function table(widths, header, rows) {
  const cell = (txt, i, head, shade) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA },
    shading: head ? { type: ShadingType.CLEAR, color: "auto", fill: NAVY } : (shade ? { type: ShadingType.CLEAR, color: "auto", fill: shade } : undefined),
    margins: { top: 80, bottom: 80, left: 110, right: 110 },
    children: [new Paragraph({ children: (Array.isArray(txt) ? txt : [txt]).map(t => typeof t === "string"
      ? new TextRun({ text: t, font: FONT, size: 19, bold: head, color: head ? "FFFFFF" : "000000" }) : t) })],
  });
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, i, true)) }),
      ...rows.map((row, k) => new TableRow({ cantSplit: true, children: row.map((c, i) => cell(c, i, false, k % 2 ? "F5F5F5" : undefined)) }))],
  });
}
const s19b = (t) => new TextRun({ text: t, font: FONT, size: 19, bold: true });
const s19 = (t) => new TextRun({ text: t, font: FONT, size: 19 });
const done = (t) => new TextRun({ text: t, font: FONT, size: 19, bold: true, color: "1D7A3A" });

const children = [
  new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 120 },
    children: [new ImageRun({ type: "png", data: fs.readFileSync(LOGO), transformation: { width: 84, height: 84 } })] }),
  new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "THE CONSTRICTOR AI", font: FONT, size: 40, bold: true, color: NAVY })] }),
  new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "The plan for Bancroft Boxing & Grappling and the Cottage Country Combat League", font: FONT, size: 24, color: GOLD, bold: true })] }),
  new Paragraph({ spacing: { after: 240 }, children: [new TextRun({ text: "Prepared by Kyle for Cody Purtell  ·  October 2026  ·  Start here", font: FONT, size: 19, italics: true, color: "555555" })] }),

  h1("The idea in one minute"),
  p([r("Your 60-bot list is the right destination. The way there is not 60 separate bots. It is "), b("a handful of workflows that share one set of records"),
     r(". Each one prepares the work and a person approves it before anything goes out. Most of your 60 ideas become features of these few workflows.")]),
  p([r("It runs on what you already use: "), b("your Claude Pro account, the CCCL Google account (ontario.cccl@gmail.com: Gmail, Drive, Sheets), Gymdesk and your Facebook/Instagram pages"),
     r(". Extra monthly cost to start: "), b("about $0"), r(". We only add paid tools once a workflow has proven it saves time or makes money.")]),
  p([b("New here? "), r("Open the Handoff folder first: your phone cheat sheet (9 images to save to your camera roll) and the Handoff Guide that walks through setup.")]),

  h1("How the pieces fit"),
  table([2100, 4560, 2700], ["Piece", "What it is", "Where it lives"], [
    [[s19b("The records")], "A few Google Sheets: event numbers, sponsors, gym leads, members to check on, tasks. This is the 'shared database' every workflow reads and updates, so nothing gets done twice.", "A shared Google Drive folder owned by the CCCL Google account"],
    [[s19b("The Constrictor")], "A set of Claude skills you run by asking: 'morning report', 'draft sponsor pitches for these five businesses', 'who hasn't been in for a month?', 'compare Peterborough and Kingston', 'message the gyms for the next show'.", "Your Claude account"],
    [[s19b("Command centre")], "A Dashboard tab in Google Sheets: approval queue, new leads, sponsor pipeline, this week's numbers. Works on your phone.", "Google Sheets"],
    [[s19b("Your team")], "Staff use the same Sheets plus ready-made message templates. The templates also work in ChatGPT, so nobody needs a new account.", "Google Drive"],
  ]),

  h1("Build order"),
  p("One workflow at a time. Each has to prove itself before the next one starts. The nine Claude skills that run these are in The Constrictor (Claude skills) folder, with an install guide."),
  table([760, 2600, 3900, 2100], ["#", "Workflow", "What it does for you", "Status"], [
    ["1", [s19b("CCCL Event Profit Predictor + Town Scout")], "Uses your real 2026 numbers and compares six 2027 options (Haliburton, Bancroft, Marmora, Peterborough, Kingston, Oshawa) with a flag for crowds under 1,000. Plus a check on adults-only shows, 24 venues researched for 2027 and 25 smaller towns ranked.", [done("Built, with your real numbers")]],
    ["2", [s19b("CCCL Sponsorship Pipeline")], "Your real packages (Title $7,500 to Community $300), the 2026 Title Sponsor first, 70 local businesses, 24 national and Ontario brands (Hayabusa, BioSteel, MuscleTech, Proper No. Twelve and more), pitch templates, a one-page sponsor sheet and follow-up reminders.", [done("Built, with your real packages")]],
    ["3", [s19b("BBG Members & Leads")], "Claude drafts replies to inquiries with the free class and the Gymdesk signup link (parents sign the waiver first), tracks leads to trial to member, and lists members who haven't been in for a month.", [done("Built")]],
    ["4", [s19b("Morning Report & Tasks")], "Constrictor HQ sheet (Dashboard, Tasks, Approvals, Content Calendar) and a morning report skill that can run every weekday at 7am.", [done("Built")]],
    ["5", [s19b("Social posts, ticket tracker, staff playbook")], "Post drafts and a content calendar, weekly ticket sales vs break-even, and a playbook with ChatGPT-ready prompts for staff.", [done("Built")]],
    ["6", [s19b("CCCL Fight Night Kit")], "Fighter sign-ups and gym outreach, the bout card with checks, a countdown from booking to the report, volunteers, production, a run of show built from the card, the press list, and the report after each show. Plus a 4-page Fighter Package for every confirmed fighter (bout, fight-day schedule, gear, rules, their ticket link).", [done("Built")]],
    ["7", [s19b("CCCL Marketing Kit + Canva designs")], "A 16-week show campaign, a tracking link for every ticket source (fighters, gyms, sponsors, posters, ads), an ad test log that says keep or stop, a poster drop list, a fight-night shot list, 16 CCCL starter designs to import into Canva (posters, fight cards, countdowns, \"I'm fighting\" posts for each fighter, weigh-ins, results, sponsor posts), and a brand kit (colours, fonts, logo rules).", [done("Built")]],
    ["8", [s19b("Handoff pack")], "A 9-card cheat sheet for your phone (what to say to Claude for every job), a Handoff Guide (the 90-minute sit-down, Google Drive sharing, the first 30 days, upkeep and fixes) and 58 skill tests Kyle runs before you get the skills.", [done("Built")]],
    ["9", [s19b("Contacts, inbox sweep, coaching library")], "One Contacts sheet for every venue, gym, media outlet, supplier and official (110 to start, from the research) with a log of every call; an inbox sweep that sorts email and leaves reply drafts; and a BBG Coaching Library: 63 drills, a 4-week plan for every class that checks time, age and contact, and the safety rules, with a seventh skill, bbg-coach.", [done("Built")]],
    ["10", [s19b("Money Summary + Funding Tracker")], "A sheet that sorts the bank, Eventbrite and Gymdesk exports by business and totals every month (\"how did we do in September?\"), and 35 researched funding programs: 18 that BBG, CCCL or gym families can use now, 13 for once Second Chance is a nonprofit or charity, and 4 that don't fit after your answers, with an eighth skill that checks eligibility and drafts applications. First up: Canada Summer Jobs 2027 opens Oct 27.", [done("Built")]],
    ["11", [s19b("Skills Passport and the last helpers")], "Four levels of skills for Kids Boxing, Wrestling and Boxing 11-14 with a progress card for parents; payment reminders from the Gymdesk export; help-with-fees answers for parents; and a monthly opportunity scan and system check.", [done("Built")]],
    ["12", [s19b("Canada Summer Jobs, Second Chance, your athlete brand")], "A ready-to-paste CSJ 2027 application for a summer student (opens Oct 27) and the first-employee basics (safety plan, harassment policy, posters, payroll); a Second Chance Starter Plan, its own private tracker, four policy drafts and five letters ready to send, with a ninth skill; and a plan for your own Facebook and Instagram in 1 to 3 hours a week (\"Get back up\").", [done("Built, from your Oct 20 answers")]],
  ]),

  h1("Who approves what"),
  p("Nothing goes out on its own. The AI drafts; a person sends."),
  bullet([b("You approve: "), r("sponsor pitches and deals, anything involving money or contracts, fight matchups, press releases, and anything public about CCCL.")]),
  bullet([b("Staff can approve: "), r("routine replies to gym inquiries using approved templates, class reminders, holiday closures, and posts from an approved content calendar.")]),

  h1("Ground rules for privacy"),
  bullet("Kids' and members' personal details stay in Gymdesk. The AI works from first names and attendance only, unless you decide otherwise."),
  bullet("No health information, fighter medicals or payment card details go into the AI or the shared Sheets."),
  bullet("Second Chance has its own folder, email, bank account and rules. Participants are a first name and a code; their records never mix with the gym or CCCL, and nobody writes down what someone was convicted of."),
  bullet("Alcohol, energy-drink and gambling sponsors stay away from youth bouts and anything showing kids. No cannabis or vape sponsors."),
  bullet("Claude and ChatGPT don't share memory. Both can read the same Drive folder, and that folder is the only thing they share."),

  h1("What you told us (October 2026)"),
  bullet([b("First show: "), r("$60 GA, $100 VIP, $1,500 VIP tables (6 sold), 400 tickets, 400-600 people, 15 sponsors for $18,000, costs listed at $29,000. That adds up to about $18,000 profit, which went into the gym and other projects. Eventbrite fees paid by CCCL; the bar was run by others; the livestream didn't make money.")]),
  bullet([b("2027: "), r("4+ shows, up to 3 hours from Bancroft, crowds of 1,000 or more. Months: Jan, Feb, Apr, May, Jun, Aug, Nov. Adults-only (19+) shows are on the table if the money makes sense.")]),
  bullet([b("Sponsors: "), r("same five levels as 2026 plus VIP tables ($1,500, 8-10 seats depending on the venue's tables). Every level gets its logo on the livestream and the event shirt, sized by level. Dream sponsors: Proper No. Twelve and a big protein brand. Pitches come from ontario.cccl@gmail.com.")]),
  bullet([b("Gym: "), r("the free class is one class of any kind (boxing, kickboxing, jiu-jitsu, wrestling, MMA), booked through the Gymdesk signup link. Parents sign the waiver first, check in after a month away, the front desk can reply too. Sean and Jordan help coach wrestling.")]),
  bullet([b("Your answers (October 20): "), r("Boxing takes kids from about 4, Jiu-Jitsu and MMA are 12+, Women's Boxing is 16+, and you coach every class and are the designated person for head injuries. Rowan's Law paperwork is done and every coach has a vulnerable sector check. BBG isn't registered with Boxing Ontario but is signed up for Jumpstart. BBG is a sole proprietorship, BBG and CCCL are two businesses, nobody is on payroll and coaches volunteer. Closed Christmas to New Year's. 2026 sponsors included Dillabough, Dixie Lee, Pinestone Resort, IT STIX and ICO, most at Bronze. Mat McTrash offered CCCL 50% of bar sales. You want a summer student, and you want to start Second Chance and your athlete brand now.")]),

  h1("Is an adults-only show worth it?"),
  p("The 19+ Show Check tab in the Profit Predictor compares three ways to run each venue. It now uses Mat McTrash's offer: CCCL gets 50% of bar sales. Even at 50%, a family show with a fenced 19+ bar area makes the most at every venue (about $18,300 vs $14,900 for adults-only at Haliburton), because an adults-only crowd is smaller and each lost ticket is worth about $73."),
  p([b("When adults-only would win: "), r("only if CCCL gets about $21 of bar money per adult (at 50%, that's adults spending about $41-42 each at the bar), if the crowd drops by about 10% or less, or with an alcohol sponsor on top. Below that, the family show with a 19+ area makes more. Mat McTrash's terms only count where he runs the bar; ask every other venue.")]),

  h1("Still needed from you"),
  check([b("Now: open BBG's CRA payroll account (RP). "), r("Canada Summer Jobs 2027 opened October 27 and asks for the payroll number in its first question; BBG only has a business number. The close date isn't published (the 2026 call closed December 11), so set up GCOS and apply from the CSJ 2027 Application Draft early.")]),
  check([b("The rest of the 2026 sponsors "), r("(5 of 15 are named so far), and which sponsors were promised spots at the Apsley, Bancroft and Haliburton shows, so nobody gets pitched twice.")]),
  check([b("Silver: $2,000 or $2,500? "), r("Pitches quote $2,000 (the 2026 price) until you confirm. Your answer said $2,500.")]),
  check([b("Send Kyle the Eventbrite payout report "), r("you found, to confirm the exact May 2 profit.")]),
  check([b("Add the free first class to Gymdesk: "), r("until it's there, replies tell new people to choose Pay Later.")]),
  check([b("Bar terms from the other venues "), r("you're calling, and which venues Mat McTrash would serve.")]),
  check([b("Second Chance: "), r("is it incorporating federally or in Ontario, are the five board members unrelated to each other and to you, and its own Gmail and bank account (Second Chance Tracker, Setup tab).")]),
  check([b("Transfer Payment Ontario: "), r("is BBG registered? Most Ontario grants need it.")]),
  check([b("Class ages: "), r("Jiu-Jitsu and MMA 12+ and Women's Boxing 16+ come from your Oct 20 answers; Wrestling from 6 and Kickboxing from 15 are what the sheets use, so say if either is wrong. One tick on the follow-up form: we read \"boxing is all ages from about 4\" as Kids Boxing 4-10 plus the Mon/Wed Boxing class open to every age, with kids in their own group.")]),

  h1("How we'll know it's working"),
  table([2600, 6760], ["Workflow", "What we measure every week"], [
    ["Profit Predictor", "Used before every venue decision. After each show: how close the prediction was to reality."],
    ["Sponsorships", "Number of sponsors at each stage (contacted, interested, proposal, signed, paid) and dollars signed and paid. An email sent is not revenue."],
    ["Members & leads", "New inquiries, how fast they got a reply, trials booked, trials that became members, quiet members contacted."],
    ["Morning report", "Hours of admin you don't have to do anymore (your own estimate, once a week)."],
  ]),

  h1("About ChatGPT"),
  p("Keep using it for strategy and thinking out loud. If you want it to see the same information, connect ChatGPT to the same Google Drive folder. Anything beyond that (the two AIs passing work to each other automatically) is a custom build, and it isn't worth it until the basics are running."),
];

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [{ id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: FONT, size: 28, bold: true, color: NAVY }, paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 0 } }] },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    { reference: "box", levels: [{ level: 0, format: LevelFormat.BULLET, text: "☐", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, bottom: 1080, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      new TextRun({ text: "The Constrictor AI plan  ·  page ", font: FONT, size: 16, color: "888888" }),
      new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: "888888" })] })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT, buf); console.log("saved", OUT); });
