// CCCL Fighter Package (editable Word / Google Docs). One copy per fighter, sent through their coach.
// usage: NODE_PATH=$(npm root -g) node fighter_pkg.js OUT.docx LOGO.png
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, Footer, PageNumber, ShadingType, LevelFormat, ImageRun,
  Table, TableRow, TableCell, WidthType, VerticalAlign, PageBreak, TableLayoutType } = require("docx");
const OUT = process.argv[2], LOGO = process.argv[3];
const INK = "151C19", PINE = "354940", RED = "AE2221", CREAM = "F2ECDC", SAGE = "DCE5DF", FONT = "Arial";
const W = 9720; // text width in DXA (Letter, 1260 margins)

// text with [placeholders] highlighted yellow so staff can see what to fill in
const R = (t, o = {}) => String(t).split(/(\[[^\]]+\])/).filter(Boolean).map(p =>
  new TextRun({ text: p, font: FONT, size: 20, ...o, ...(p.startsWith("[") ? { highlight: "yellow" } : {}) }));
const P = (t, o = {}, ro = {}) => new Paragraph({ spacing: { after: 90 }, ...o, children: Array.isArray(t) ? t.flatMap(x => typeof x === "string" ? R(x, ro) : x) : R(t, ro) });
const B = (t, o = {}) => R(t, { bold: true, ...o });
const H1 = (t, sub, brk) => [new Paragraph({ keepNext: true, pageBreakBefore: !!brk, spacing: { before: 260, after: sub ? 20 : 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: RED, space: 3 } },
    children: [new TextRun({ text: t.toUpperCase(), font: FONT, size: 28, bold: true, color: INK, characterSpacing: 20 })] }),
  ...(sub ? [new Paragraph({ keepNext: true, spacing: { after: 110 }, children: [new TextRun({ text: sub, font: FONT, size: 19, italics: true, color: "5C6661" })] })] : [])];
const H2 = (t) => new Paragraph({ keepNext: true, spacing: { before: 170, after: 60 }, children: [new TextRun({ text: t, font: FONT, size: 22, bold: true, color: PINE })] });
const bullet = (t, lvl = 0) => new Paragraph({ numbering: { reference: "b", level: lvl }, spacing: { after: 45 }, children: Array.isArray(t) ? t.flatMap(x => typeof x === "string" ? R(x) : x) : R(t) });
const box = (t) => new TextRun({ text: "☐", font: "Segoe UI Symbol", size: 22 });

const border = { style: BorderStyle.SINGLE, size: 4, color: "C9D3CD" };
const borders = { top: border, bottom: border, left: border, right: border };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: none, bottom: none, left: none, right: none };
function cell(content, w, o = {}) {
  let kids;
  if (typeof content === "string") kids = [new Paragraph({ spacing: { after: 30 }, children: R(content, o.run || {}) })];
  else if (content.every && content.every(c => c instanceof Paragraph)) kids = content;
  else kids = [new Paragraph({ spacing: { after: 30 }, children: content.flat(3).flatMap(x => typeof x === "string" ? R(x, o.run || {}) : x) })];
  return new TableCell({ width: { size: w, type: WidthType.DXA }, borders: o.borders || borders, verticalAlign: o.v || VerticalAlign.CENTER,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: o.mt ?? 50, bottom: o.mb ?? 50, left: 120, right: 120 }, columnSpan: o.span, children: kids });
}
function table(cols, rows, o = {}) {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: cols, layout: TableLayoutType.FIXED,
    rows: rows.map((r, i) => new TableRow({ cantSplit: true, tableHeader: o.header && i === 0,
      children: r.map((c, j) => {
        if (o.header && i === 0) return cell([B(c, { color: "FFFFFF", size: 18 })], cols[j], { fill: PINE });
        const first = j === 0 && o.label;
        return cell(typeof c === "string" ? (first ? [B(c, { size: 19, color: INK })] : c) : c, cols[j], { fill: first ? SAGE : (o.zebra && i % 2 === 0 ? "F6F8F6" : undefined) });
      }) })) });
}
const spacer = (a = 120) => new Paragraph({ spacing: { after: a }, children: [] });
const callout = (title, lines, fill = "FBF3E9", edge = RED) => new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], layout: TableLayoutType.FIXED,
  rows: [new TableRow({ cantSplit: true, children: [new TableCell({ width: { size: W, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR, color: "auto" },
    borders: { top: none, bottom: none, right: none, left: { style: BorderStyle.SINGLE, size: 24, color: edge } }, margins: { top: 110, bottom: 110, left: 200, right: 160 },
    children: [new Paragraph({ spacing: { after: 50 }, children: B(title, { color: INK, size: 21 }) }), ...lines.map(l => P(l, { spacing: { after: 50 } }))] })] })] });

// ---------- cover band ----------
const cover = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [7020, 2700], layout: TableLayoutType.FIXED,
  rows: [new TableRow({ children: [
    new TableCell({ width: { size: 7020, type: WidthType.DXA }, shading: { fill: INK, type: ShadingType.CLEAR, color: "auto" }, borders: noBorders,
      margins: { top: 240, bottom: 220, left: 340, right: 120 }, verticalAlign: VerticalAlign.CENTER, children: [
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "COTTAGE COUNTRY COMBAT LEAGUE", font: FONT, size: 20, bold: true, color: "D9483B", characterSpacing: 40 })] }),
        new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "FIGHTER PACKAGE", font: "Arial Black", size: 52, bold: true, color: CREAM })] }),
        new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: "CCCL ", font: FONT, size: 24, bold: true, color: CREAM }), ...R("[Town] · [day, date]", { size: 24, bold: true, color: INK })] }),
      ] }),
    new TableCell({ width: { size: 2700, type: WidthType.DXA }, shading: { fill: INK, type: ShadingType.CLEAR, color: "auto" }, borders: noBorders,
      margins: { top: 160, bottom: 160, left: 120, right: 260 }, verticalAlign: VerticalAlign.CENTER, children: [
        new Paragraph({ alignment: AlignmentType.RIGHT, children: [new ImageRun({ type: "png", data: fs.readFileSync(LOGO), transformation: { width: 130, height: 130 } })] })] }),
  ] })] });

const children = [
  cover,
  new Paragraph({ spacing: { before: 120, after: 60 }, children: [
    new TextRun({ text: "For staff: ", font: FONT, size: 17, bold: true, color: RED }),
    new TextRun({ text: "fill in everything highlighted yellow, check each time and rule against the Fight Night Kit (Show Setup and Rules tabs), then send one copy per fighter through their coach. Adults only for the promo page. Never ask for or write down medical details.", font: FONT, size: 17, color: "5C6661" })] }),

  ...H1("You're on the card", "Everything you need from now to fight night. Keep it handy and go through it with your coach."),
  P(["Hi ", "[first name]", ", welcome to the Cottage Country Combat League. Your bout is below. If anything is wrong, tell your coach right away."]),
  table([2700, 7020], [
    ["Show", "CCCL [number] · [Town]"],
    ["Date", "[Day, Month date, year]"],
    ["Venue", "[Venue name, street address, town]"],
    ["Your bout", "[Your name] vs [opponent's name] ([opponent's gym])"],
    ["Weight", "[weight] lb. You must make this at the weigh-in on fight day."],
    ["Rules", "[Senior amateur MMA · 3 x 3 min]  or  [Youth A · 3 x 2 min]"],
    ["Experience class", "[Pre-Novice (0-5 contests) / Novice (6-10) / Open (11+)]"],
    ["Weigh-in", "[time] at [place]. Bring photo ID."],
    ["Check-in", "[time] at [entrance], for paperwork and hand wraps"],
    ["Doors · first bout", "[Doors time] · [first bout time]. You're bout [number] of [number]."],
    ["Your corner", "Up to 3 people. Send their first names to your coach by [date]."],
    ["Questions", "Your coach first, or Cody Purtell (CCCL), 705-760-0938 or ontario.cccl@gmail.com"],
  ], { label: true }),
  ...H1("Ready to fight", "Your coach and CCCL track these. You're on only when every box is ticked."),
  table([7020, 2700], [
    ["Item", "Status"],
    ["Medical exam done (the result goes to the physician and the Commission)", "[Yes / No / Pending]"],
    ["Commission paperwork in", "[Yes / No / Pending]"],
    ["Code of conduct and rules declaration signed (fight day)", "[Yes / No]"],
    ["No-suspension form signed (fight day)", "[Yes / No]"],
    ["Waiver signed", "[Yes / No]"],
    ["Parent or guardian consent (under 18 only)", "[Yes / No / Not needed]"],
  ], { header: true, zebra: true }),

  ...H1("Your next six weeks", "Tick them off as you go.", true),
  table([1800, 7040, 880], [
    ["When", "What", "Done"],
    ["Now", "Read this with your coach. Send your walkout song and two lines about you for the announcer (hometown, job, why you fight) by [date].", [box()]],
    ["Now", "Send one sharp photo for your fight graphics by [date]: chest up, bright light, plain background, no sunglasses. (Adults only.)", [box()]],
    ["Now", "Post that you're fighting, with your own ticket link (page 4).", [box()]],
    ["By [date]", "Medical exam done and your Commission paperwork in. Your coach knows the steps. Medical details go to the physician and the Ontario Athletics Commission, never to CCCL.", [box()]],
    ["By [date]", "Your corners' first names to your coach (up to 3).", [box()]],
    ["2 weeks out", "Second post. Check your gear against the list on page 3.", [box()]],
    ["Fight week", "Last post. Plan your weight so you make it on the day: weigh-ins are on fight day, within 3 hours of the first bout. There's no day-before weigh-in.", [box()]],
    ["Fight day", "Follow the schedule on page 3. Bring everything on the gear list.", [box()]],
    ["Day after", "Share your result post (official results only) and thank the people who came out.", [box()]],
  ], { header: true, zebra: true }),

  ...H1("Rules to know", "Amateur rules. The officials have the final say, so ask at your referee's briefing if you're unsure."),
  table([2400, 7320], [
    ["Class", "What's not allowed (on top of the usual fouls)"],
    ["Everyone", "No elbow or forearm strikes, no knees to the head, no heel hooks, toe holds or neck cranks, no stomps."],
    ["Pre-Novice (0-5)", "Also no strikes to the head of a downed opponent, no slams, suplexes or high throws, no up-kicks from the ground."],
    ["Novice (6-10)", "Also no strikes to the head of a downed opponent."],
    ["Open (11+)", "Nothing extra."],
    ["Youth (under 19)", "No strikes to the head or face at all, no knee strikes, no slams or up-kicks, plus extra limits by age group. Your coach has the full list."],
  ], { header: true, zebra: true }),
  spacer(80),
  callout("Fighting under 18?", [
    "A parent or guardian signs your forms and the fight-day no-suspension form, and comes to the weigh-in with you.",
    "CCCL never names anyone under 18 in a press release, never puts you in an ad, and only shares a photo of you if your parent says yes in writing. The promo page is for adult fighters, but your family can still buy tickets at [Eventbrite link].",
  ], "EEF3F0", PINE),

  ...H1("Fight day", "Times are for this show. Officials can change the order on the night.", true),
  table([2400, 7320], [
    ["Time", "What happens"],
    ["[time]", [...B("Weigh-in "), ...R("at [place]. You must make [weight] lb. Over? You get 1 hour to make it, and you can't lose more than 1% of your body weight in that hour or you can't fight.")]],
    ["[time]", [...B("Check-in. "), ...R("Sign the code of conduct, the rules declaration and the no-suspension form (a parent or guardian signs for anyone under 18).")]],
    ["[time]", [...B("Hand wraps. "), ...R("CCCL's wrappers do them. An official checks them before your gloves go on.")]],
    ["1 hr+ before your bout", [...B("Pre-bout medical "), ...R("with the physician. They have the final say on whether you fight.")]],
    ["Before your bout", [...B("Referee's briefing "), ...R("in the dressing room, for you and your lead corner. Ask any rules questions here.")]],
    ["[time] · [time]", [...B("Doors open · first bout. "), ...R("You're bout [number]. Stay in the fighters' area once you're warmed up.")]],
    ["Right after", [...B("Post-bout medical. "), ...R("Go straight to the exam room with at least one corner, win or lose.")]],
  ], { header: true, zebra: true }),

  H2("What to bring"),
  bullet([box(), " Photo ID (we check ages)"]),
  bullet([box(), " Mouthguard (bring one; treat it as required)"]),
  bullet([box(), " Soft neoprene shin guards, tight-fitting, with no buckles, straps or Velcro showing"]),
  bullet([box(), " Groin cup (required for men; optional groin or chest protector for women)"]),
  bullet([box(), " MMA or compression shorts with no pockets, zips or fasteners"]),
  bullet([box(), " Short-sleeve rash guard if you want one (optional)"]),
  bullet([box(), " Water and food for after the weigh-in, a hoodie and sandals for the walk to the cage"]),
  H2("CCCL supplies"),
  bullet("Gloves (everyone wears the same type, so don't bring your own) and hand wraps, plus the people who wrap"),
  bullet("A stool, bucket, water and ice for your corner, and a third corner person for cuts if you need one"),
  H2("Not allowed in the cage"),
  bullet("Headgear, shoes, a gi, jewellery, glasses or hard contacts, makeup"),
  bullet("Tie your hair back, trim your nails, contain facial hair. Petroleum jelly only on the brows, under the eyes and on the nose."),
  H2("Your corner"),
  bullet("Up to 3 people. Two can come in between rounds. They stay seated during rounds and leave at the 10-second whistle."),
  bullet("Corner bag only: petroleum jelly, enswell, cotton swabs, gauze, water in a clear bottle, a clear spray bottle, bagged ice, towels, sponges, tape and scissors."),
  bullet("If your lead corner asks the referee to stop the fight, it's a TKO loss. Talk about that plan with your coach before the night."),

  ...H1("Bring your crowd", "For adult fighters. Your people make the night: here's how to get them there.", true),
  table([2700, 7020], [
    ["Your ticket link", [...B("[your personal tracking link]", { size: 22 })]],
    ["Why it matters", "Every ticket bought through your link shows up under your name, so we can see who brought their crowd and shout you out. Share this link, not the general one."],
  ], { label: true }),
  H2("Three posts you can copy"),
  P([...B("Now: "), ...R("\"I'm fighting at the Cottage Country Combat League in [Town] on [date]! Come yell for me: [your link]\"")]),
  P([...B("Two weeks out: "), ...R("\"Two weeks. Camp's almost done. Tickets for [Town] here: [your link]\"")]),
  P([...B("Fight week: "), ...R("\"[Day] night at [venue]. Doors at [time]. Last chance for tickets: [your link]\"")]),
  H2("Graphics you'll get from us"),
  bullet([...B("\"I'm fighting\" post and story "), ...R("with your photo, once your bout is confirmed. On the story, add an Instagram link sticker with your ticket link (Stickers > Link).")]),
  bullet([...B("Fighter vs fighter "), ...R("when the card is announced.")]),
  bullet([...B("Your result "), ...R("the day after, once the result is official.")]),
  P("Tag @[CCCL Instagram] and CCCL on Facebook so we can reshare you."),
  H2("Keep it clean"),
  bullet("Hype yourself and the night, not trash talk. Respect your opponent, their gym and the officials, online and in person."),
  bullet("No betting or odds talk, and don't post anyone else's photo without their OK."),
  bullet("Results only once they're official."),

  spacer(100),
  P([...B("On the night: "), ...R("no alcohol or drugs on fight day, follow every call the referee and physician make, and keep the dressing room to fighters, corners, officials and medical staff.")]),
  P([...B("See you on fight night. ", { color: RED }), ...R("Cody \"The Constrictor\" Purtell, Cottage Country Combat League")]),
];

const doc = new Document({
  creator: "Cottage Country Combat League", title: "CCCL Fighter Package",
  numbering: { config: [{ reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 240 } } } }] }] },
  styles: { default: { document: { run: { font: FONT, size: 20 } } } },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 900, bottom: 900, left: 1260, right: 1260 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      new TextRun({ text: "CCCL fighter package  ·  page ", font: FONT, size: 16, color: "888888" }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: "888888" })] })] }) },
    children }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("saved", OUT); });
