const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, BorderStyle, Footer, PageNumber, ShadingType, LevelFormat } = require("docx");
const OUT = process.argv[2];
const NAVY = "081C2B", GOLD = "C8862E", FONT = "Arial";
const T = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 21, ...o });
const P = (runs, o = {}) => new Paragraph({ spacing: { after: 100 }, ...o, children: (Array.isArray(runs) ? runs : [runs]).map(x => typeof x === "string" ? T(x) : x) });
let SEC = 0;
const H1 = (t, br) => { SEC++; return new Paragraph({ heading: HeadingLevel.HEADING_1, keepNext: true, pageBreakBefore: !!br, spacing: { before: 300, after: 90 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: GOLD, space: 2 } },
  children: [new TextRun({ text: t, font: FONT, size: 28, bold: true, color: NAVY })] }); };
const H2 = (t) => new Paragraph({ keepNext: true, spacing: { before: 200, after: 60 }, children: [new TextRun({ text: t, font: FONT, size: 23, bold: true, color: NAVY })] });
const step = (runs) => new Paragraph({ numbering: { reference: "n" + SEC, level: 0 }, spacing: { after: 60 }, children: (Array.isArray(runs) ? runs : [runs]).map(x => typeof x === "string" ? T(x) : x) });
const bullet = (runs) => new Paragraph({ numbering: { reference: "b", level: 0 }, spacing: { after: 60 }, children: (Array.isArray(runs) ? runs : [runs]).map(x => typeof x === "string" ? T(x) : x) });
function box(lines) {
  const sh = { type: ShadingType.CLEAR, color: "auto", fill: "F4F6F8" }, ind = { left: 240, right: 240 };
  return lines.map((l, i) => new Paragraph({ shading: sh, indent: ind, keepNext: i < lines.length - 1, keepLines: true, spacing: { after: 70 },
    children: [new TextRun({ text: l, font: FONT, size: 20 })] })).concat([new Paragraph({ spacing: { after: 120 }, children: [] })]);
}

const FACTS = [
  "ABOUT BANCROFT BOXING & GRAPPLING (use only these facts):",
  "- Gym at 17 Snow Road, Bancroft, Ontario. Owner and head coach: Cody Purtell. Phone 705-760-0938.",
  "- Boxing, kickboxing, wrestling, no-gi jiu-jitsu, MMA and yoga. Boxing takes kids from about 4, wrestling from 6, Jiu-Jitsu and MMA from 12, Kickboxing from 15 (ask Cody about younger teens), Women's Boxing from 16. Cody coaches every class.",
  "- Closed from Christmas to New Year's (check the exact dates with Cody).",
  "- New people get their FIRST CLASS FREE. After that: $20 drop-in, $25 for a week, $50 per 4 weeks for the Kids Program or Women's Boxing, $75/month boxing only (or grappling only), $125/month all access (boxing, jiu-jitsu, wrestling), $1,000 for a year of all access. Personal training $50. Birthday party $225. Fitness camp $250. No refunds once purchased.",
  "- Schedule: Monday Wrestling 3:20-4:00, Boxing 4-5. Tuesday Jiu-Jitsu 4-6, Yoga 6:15-7. Wednesday Boxing 4-5, Kids Boxing 5:15-6:15, Women's Boxing 7-7:45. Thursday Boxing ages 11-14 4-5, Women's Boxing 6-6:45, Kickboxing 7-8. Friday MMA 4-5, Competition 5-6.",
  "- The free first class can be any one class: boxing, kickboxing, jiu-jitsu, wrestling or MMA. To book it, they sign up, pick the free first class and sign the waiver here: https://bancroft-boxing-grappling.gymdesk.com/signup . For anyone under 18, a parent or guardian fills it in and signs the waiver before the first class. They don't pay anything for the first class. Until the free-class option is on the signup page, tell them: fill in the form and sign the waiver, choose Pay Later, and tell us which class you're coming to.",
  "- Cody coaches every class. Sean and Jordan help coach wrestling.",
  "RULES FOR EVERY MESSAGE: warm, short and local, like a friendly coach. Suggest one real class and include the signup link. No medical or injury advice. For kids, talk to the parent. Never pressure. If someone is upset, injured or asking for a refund, don't answer it: say Cody will reach out today.",
];

const children = [
  new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "BBG & CCCL STAFF PLAYBOOK", font: FONT, size: 38, bold: true, color: NAVY })] }),
  new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "How the team uses the sheets and AI. One section per job.", font: FONT, size: 22, bold: true, color: GOLD })] }),

  H1("The three rules"),
  step([T("AI drafts, a person sends. ", { bold: true }), T("Read every draft before it goes out. Fix anything that's wrong or doesn't sound like us.")]),
  step([T("Cody approves the big stuff. ", { bold: true }), T("Anything to sponsors, venues, the press, the Ontario Athletics Commission, or about money or contracts goes to Cody first (he logs it on his Approvals tab). Staff can send routine gym replies that follow the templates.")]),
  step([T("Protect people's privacy. ", { bold: true }), T("First names only in the sheets. No phone numbers, health details or kids' last names. Talk to parents, not kids. No photos of kids without a parent's OK.")]),

  H1("Where things live"),
  P("Cody shares only the sheets your job needs. If you can't open one, ask him."),
  bullet([T("Your tasks: ", { bold: true }), T("Cody sends them by text. Constrictor HQ, the Sponsorship Pipeline, the Fight Night Kit and the Contacts are Cody's.")]),
  bullet([T("BBG Leads & Trials (front desk): ", { bold: true }), T("one row for every person who asks about classes.")]),
  bullet([T("BBG Coaching Library and Skills Passport (coaches): ", { bold: true }), T("class profiles, the drill library, 4-week lesson plans, the safety rules and kids' progress. Coaches check it before every class.")]),
  bullet([T("Fight night (volunteers): ", { bold: true }), T("a printed volunteer list, run of show and banner list from Cody.")]),
  bullet([T("Gymdesk: ", { bold: true }), T("memberships, payments, attendance, waivers. Personal details stay here.")]),

  H1("1. Someone asks about classes"),
  step("Front desk or Cody: reply within a few hours if you can. Use prompt A below (ChatGPT or Claude), or Cody's Front Desk skill if you're on his account."),
  step("Always offer the free first class, suggest one real class time that fits them, and send the signup link (bancroft-boxing-grappling.gymdesk.com/signup). For kids, the parent fills it in and signs the waiver before the class."),
  step([T("Add a row to "), T("Leads", { bold: true }), T(": date, first name, who it's for, what they want, how they reached out, Replied = Yes. If they book, add the free trial date.")]),
  step("If they don't answer in 2 days, send one friendly nudge. Then let it go."),

  H1("2. After a free trial"),
  step("Same day or next morning: a short thank-you text (prompt B). Ask how it felt, and suggest the next class."),
  step([T("On Leads: "), T("Came to trial? ", { bold: true }), T("Yes/No, and "), T("Joined? ", { bold: true }), T("Yes / Not yet / No - not now, with the plan if they joined.")]),
  step("If they joined, make sure Cody has set them up in Gymdesk."),

  H1("3. Fridays: check in on quiet members"),
  step("Cody (or Claude) pulls the list of members who haven't been in for a month from Gymdesk."),
  step("Send each one a short, no-pressure check-in (prompt C) after Cody OKs the list."),
  step("Never mention money or their membership. Just: we miss you, here's a class that fits."),

  H1("4. Posting on social media"),
  step([T("Cody sends the posts he's approved. Only post those. Gym posts go on BBG's pages; fight-night posts go on CCCL's own Facebook and Instagram, then get shared to BBG's.")]),
  step("Check the class time on Gymdesk before posting. Schedules change."),
  step("Everyone in a photo or video has said yes. Kids only with a parent's OK."),
  step("After posting, tell Cody so he can mark it Posted. Reply to comments with a thank-you and the free first class."),

  H1("5. Fight night (volunteers)"),
  step("Check your role and shift on the printed volunteer list, and read the volunteer briefing Cody sends the week of the show."),
  step("Use the printed banner list from Cody. Every sponsor's banner goes up where it was promised."),
  step("Give the announcer the sponsor shout-out list Cody approved."),
  step("Take a clear photo of every sponsor banner with the crowd. Those photos go in the thank-you emails."),
  step("Kids' grappling: no photos or video unless the parent has said yes."),

  H1("6. Coaching a class"),
  step([T("Open the "), T("Lesson Plans", { bold: true }), T(" tab of the BBG Coaching Library and find this week's plan for your class. The Check column should say OK. Set out the equipment the drills need before class.")]),
  step("Start with the warm-up, teach the tap and breakfalls before any live grappling, and give water breaks."),
  step([T("Kids under 12: ", { bold: true }), T("no strikes to the head, ever. Partner work for kids is glove taps or cooperative drilling. Sparring with head contact is Competition Class only, with Cody's OK.")]),
  step([T("Any sign of a head injury ", { bold: true }), T("(headache, dizzy, confused, sick, unsteady): the class's designated person removes them for the day and they don't go back in. Tell a parent for anyone under 18 and tell Cody. Only a doctor or nurse practitioner can clear them to come back (Ontario's Rowan's Law).")]),
  step([T("Rule of Two: ", { bold: true }), T("never be alone with a child. Two screened adults present, one-on-one time only in the open with another adult in sight, and no private messages to kids: group messages that include parents.")]),
  step("After class, write anything to change in the plan's Notes column (a drill that ran long, one the kids loved)."),
  step([T("Kids' and youth classes: when a kid nails a skill, log it on the "), T("Progress", { bold: true }), T(" tab of the BBG Skills Passport (first name, skill, date). Log each skill once.")]),

  H1("Prompt pack: copy, paste, edit, send", true),
  P("These work in ChatGPT or free Claude. First paste the BBG facts block, then the prompt you need. Always read the answer before sending, and check class times on Gymdesk."),
  H2("Paste this first: the BBG facts block"),
  ...box(FACTS),
  H2("A. Reply to a class inquiry"),
  ...box(["Using the BBG facts above, write a reply to the message below. Under 90 words. Answer their question first, offer the free first class, suggest one real class time that fits them, and give the signup link so they can sign up and do the waiver before they come.", "Message: [paste their message]"]),
  H2("B. Thank-you after a free trial"),
  ...box(["Using the BBG facts above, write a short text to [first name], who tried [class] today. Thank them, ask how it felt, and suggest the next class that fits. Mention the $25 week or a membership only if it fits naturally. Under 50 words."]),
  H2("C. Quiet-member check-in"),
  ...box(["Using the BBG facts above, write a short, friendly text from Cody to [first name], who usually comes to [class] but hasn't been in for about a month. No guilt, no mention of money or membership. Invite them to [class] on [day]. Under 40 words."]),
  H2("D. Class-of-the-week post"),
  ...box(["Using the BBG facts above, write a Facebook and Instagram post about [class] for beginners: who it's for, what they'll do, the day and time, and that the first class is free. Under 60 words, up to 2 emoji, 3-5 hashtags starting with #Bancroft."]),
  H2("E. Birthday party inquiry"),
  ...box(["Using the BBG facts above, reply to this birthday party question. The package is $225. Ask for the date, the kids' ages and roughly how many, and say Cody will confirm. Under 60 words.", "Message: [paste]"]),
  H2("F. Closure or schedule change"),
  ...box(["Using the BBG facts above, write a short announcement for Facebook and a text version: [what's changing, which classes, which dates, why if we want to say]. Clear and friendly. Under 60 words."]),
  H2("G. Plan a class (coaches)"),
  ...box(["You're helping a coach at a boxing and grappling gym plan a [minutes]-minute [class] for [ages]. Allowed contact: [e.g. pads and bags only / cooperative partner drilling / live grappling]. Theme: [e.g. the jab].",
          "Plan six parts: warm-up, two skills, a pad or partner drill, a game or live round, and a cool-down, leaving 5-10 minutes for water. Give a timeline with minutes, how to run each drill in two lines, coaching points and safety notes.",
          "Rules: no strikes to the head for anyone under 12; teach the tap and breakfalls before any live grappling; no slams or neck cranks; nothing above the allowed contact."]),
];
const nums = Array.from({ length: 12 }, (_, i) => ({ reference: "n" + i, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 300 } } } }] }));
const doc = new Document({
  numbering: { config: [...nums, { reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 260 } } } }] }] },
  styles: { default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [{ id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 28, bold: true, color: NAVY }, paragraph: { outlineLevel: 0 } }] },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 960, bottom: 960, left: 1260, right: 1260 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      new TextRun({ text: "BBG & CCCL staff playbook  ·  page ", font: FONT, size: 16, color: "888888" }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: "888888" })] })] }) },
    children }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("saved", OUT); });
