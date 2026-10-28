const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, BorderStyle, Footer, PageNumber, ShadingType, LevelFormat, ImageRun } = require("docx");
const OUT = process.argv[2];
const LOGO = process.argv[3];
const INK = "151C19", PINE = "354940", RED = "AE2221", FONT = "Arial";

const T = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 21, ...o });
const P = (runs, o = {}) => new Paragraph({ spacing: { after: 100 }, ...o, children: (Array.isArray(runs) ? runs : [runs]).map(x => typeof x === "string" ? T(x) : x) });
let SEC = 0;
const H1 = (t) => { SEC++; return new Paragraph({ heading: HeadingLevel.HEADING_1, keepNext: true, spacing: { before: 320, after: 80 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: RED, space: 2 } },
  children: [new TextRun({ text: t, font: FONT, size: 28, bold: true, color: INK })] }); };
const H2 = (t) => new Paragraph({ keepNext: true, spacing: { before: 200, after: 60 }, children: [new TextRun({ text: t, font: FONT, size: 23, bold: true, color: PINE })] });
const WHEN = (t) => P([T(t, { italics: true, color: "666666", size: 19 })], { spacing: { after: 120 }, keepNext: true });
const bullet = (runs) => new Paragraph({ numbering: { reference: "b", level: 0 }, spacing: { after: 50 }, children: (Array.isArray(runs) ? runs : [runs]).map(x => typeof x === "string" ? T(x) : x) });
const SIG = ["Thanks,", "Cody Purtell", "Cottage Country Combat League  ·  Bancroft Boxing & Grappling", "705-760-0938"];
function box(subject, lines) {
  const sh = { type: ShadingType.CLEAR, color: "auto", fill: "EEF2EF" };
  const ind = { left: 240, right: 240 };
  const out = [];
  if (subject) out.push(new Paragraph({ shading: sh, indent: ind, keepNext: true, spacing: { before: 60, after: 140 }, children: [T("Subject: ", { bold: true }), T(subject)] }));
  const sig = new Set(SIG);
  const real = lines.filter(l => l !== "");
  real.forEach((l, i) => out.push(new Paragraph({ shading: sh, indent: ind, keepNext: i < real.length - 1, keepLines: true,
    spacing: { after: sig.has(l) && l !== SIG[0] ? 0 : (l === SIG[0] ? 40 : 120) },
    children: l.startsWith("## ") ? [T(l.slice(3), { bold: true })] : [T(l)] })));
  out.push(new Paragraph({ spacing: { after: 160 }, children: [] }));
  return out;
}

const children = [
  new Paragraph({ spacing: { after: 80 }, children: [new ImageRun({ type: "png", data: fs.readFileSync(LOGO), transformation: { width: 78, height: 78 } })] }),
  new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "CCCL FIGHT NIGHT & PROMO TEMPLATES", font: FONT, size: 36, bold: true, color: INK })] }),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "Messages, forms, scripts and promo copy for running and selling a show. Fill in the [brackets], read it once, send.", font: FONT, size: 22, color: RED, bold: true })] }),
  P([T("House rules: ", { bold: true }), T("Cody sends everything himself. Contact fighters through their gym or coach. Never ask for or write down medical details: the physician and the Athletics Commission handle those. Youth athletes: talk to the parent, first name and last initial only, no photos without a parent's OK. No prize worth $300 or more (it would make a fighter a pro). Results only once they're official. Check every rule against the Rules tab of the Fight Night Kit and the Commission.")]),

  H1("1. Message to a gym's head coach"),
  WHEN("First contact with a gym on the Recruitment tab, about 12 weeks before the show. Email or the gym's Facebook/Instagram page."),
  ...box("Amateur MMA bouts for your fighters: CCCL [Town], [date]", [
    "Hi Coach [first name],",
    "",
    "I'm Cody Purtell from Bancroft Boxing & Grappling. I run the Cottage Country Combat League, amateur MMA sanctioned by the Ontario Athletics Commission. Our first show in Haliburton sold 400 tickets, and our next one is [Town] on [date] at [venue].",
    "",
    "I'm building the card now and looking for [weight classes, e.g. 'lightweights and welterweights with 0-5 amateur fights']. Fighters get a proper show: Commission officials, a physician cage-side, a livestream and a real crowd. [Youth grappling: we're also running youth grappling matches if your kids' program wants in.]",
    "",
    "If you've got athletes who are ready, have them fill in the sign-up form here: [form link], or send me their names, weights and records and I'll take it from there.",
    "",
    ...SIG]),
  H2("Short version for Facebook or Instagram"),
  ...box(null, ["Hey Coach! Cody from Bancroft Boxing & Grappling here 👊 We're running CCCL [Town] on [date], amateur MMA sanctioned by the Ontario Athletics Commission. Looking for [weight classes]. Got anyone ready? Sign-up form: [link]"]),

  H1("2. Follow-up to a gym"),
  WHEN("No reply after a week. Once only."),
  ...box("Re: Amateur MMA bouts for your fighters", ["Hi Coach [first name],", "", "Just bumping this up. I'm locking in the [Town] card by [date] and still need [weight classes]. Happy to talk on the phone if that's easier.", "", ...SIG]),

  H1("3. Fighter sign-up form (make it in Google Forms)"),
  WHEN("Takes about 10 minutes. In Google Forms (signed in as ontario.cccl@gmail.com), make a new form called 'CCCL [Town] fighter sign-up', add these questions, and link the responses to a Google Sheet. Copy accepted fighters onto the Fighters tab."),
  H2("Adult amateur MMA"),
  bullet("Full name, as it should appear on the card (short answer, required)"),
  bullet("Nickname, if any (short answer)"),
  bullet("Gym or team (short answer, required)"),
  bullet("Head coach's name and the gym's email (short answer, required)"),
  bullet("Fight weight you'll make (lb) (short answer, required)"),
  bullet("Amateur MMA record: wins and losses, and a link to it if it's online (short answer, required)"),
  bullet("Other amateur contests in any combat sport (boxing, kickboxing, wrestling...): how many and where (short answer, required). This sets the experience class: Pre-Novice 0-5, Novice 6-10, Open 11+."),
  bullet("Age division: Senior (19+), Youth A (16-18), Youth B (14-15), Youth C (12-13) (multiple choice, required). Under 18: a parent or guardian must also sign."),
  bullet("Any professional fights, or ever paid a purse or prize worth $300 or more? (Yes / No, required). If yes, they may not count as an amateur: check with the Commission."),
  bullet("Which show dates can you do? (checkboxes)"),
  bullet("Walkout song (short answer)"),
  bullet("Two lines about you for the announcer: hometown, job, why you fight (paragraph)"),
  bullet("OK to use your name, photo and video to promote the show? (Yes / No, required)"),
  bullet("I understand I'll need a medical exam and Ontario Athletics Commission paperwork, and that medical information goes to the Commission and the physician, not to CCCL (checkbox, required)"),
  P([T("Don't add questions about injuries, medical history or health. ", { bold: true }), T("That information goes through the Commission's process, not the promoter's form.")]),
  H2("Youth grappling (filled in by a parent or guardian)"),
  bullet("Athlete's first name and last initial (required)"),
  bullet("Age group (required): use the grappling ruleset's age divisions (see the Rules tab of the Fight Night Kit)"),
  bullet("Weight (lb) and club (required)"),
  bullet("Years of training: under 1, 1-3, 3 or more (required)"),
  bullet("I've reviewed the concussion awareness resources with my child (Rowan's Law) (checkbox, required)"),
  bullet("Parent or guardian's name and email (required)"),
  bullet("I give permission for my child to compete in youth grappling at CCCL [Town] on [date] (checkbox, required)"),
  bullet("OK to show my child in photos or video? (Yes / No, required)"),

  H1("4. Bout offer to a coach"),
  WHEN("Once Cody and the officials are happy with a matchup. Send to both coaches."),
  ...box("Bout offer: [fighter] vs [opponent], CCCL [Town] [date]", [
    "Hi Coach [first name],",
    "",
    "We'd like to match [fighter] with [opponent] from [opponent's gym] on CCCL [Town], [date].",
    "## [Senior / Youth A] amateur MMA · [weight] lb · [3 x 3 min / 3 x 2 min] · [experience class]",
    "[Opponent] is [record], [one line on style if known]. Both fighters weigh in on the day, within 3 hours of the first bout.",
    "",
    "Can you let me know by [date] if [fighter] is in? Once confirmed I'll send [fighter]'s Fighter Package: weigh-in and fight-day times, the gear list, the rules and their ticket link.",
    "",
    ...SIG]),

  H1("5. Fighter info sheet (cover email for the Fighter Package)"),
  WHEN("To every confirmed fighter (through their coach) about 5 weeks out, with the Fighter Package attached (CCCL / Fight Night / CCCL Fighter Package: fill in everything highlighted yellow, one copy per fighter). This email is the short version. Check every time and rule against the Rules tab and the Commission before sending."),
  ...box("You're on the card: CCCL [Town], [date]", [
    "Hi [first name],",
    "",
    "You're confirmed for CCCL [Town] on [date] at [venue, address]. Your Fighter Package is attached: your bout, the fight-day schedule, the gear list, the rules and your ticket link. Go through it with your coach. Here's the short version.",
    "## Before fight week",
    "Finish your medical exam and Commission paperwork by [date]. Your coach knows the steps. Send your walkout song by [date].",
    "## Weigh-ins (show day)",
    "[Time] at [place], within 3 hours of the first bout. Bring photo ID. You must make [weight] lb. If you're over, you get 1 hour to make it, and you can't fight if you lose more than 1% of your body weight in that hour.",
    "## Fight day",
    "Check in at [time] to sign the codes of conduct, the rules declaration and the no-suspension form (a parent signs for anyone under 18). Referee's instructions at [time]. Your pre-bout medical is at least 1 hour before your bout. Doors open at [time], first bout at [time].",
    "## What to bring",
    "Photo ID, mouthguard, soft neoprene shin guards (no straps or Velcro), a cup (men), fight shorts and a rash guard if you want one, water. CCCL supplies the gloves and hand wraps, so don't bring your own gloves. No headgear or shoes. Up to 3 corners.",
    "## Rules to know",
    "Amateur rules: no elbows, no knees to the head, no heel hooks, toe holds or neck cranks, no stomps. [Pre-Novice: no ground strikes to the head, slams or up-kicks. Novice: no ground strikes to the head.] [Youth: no head or face strikes at all.]",
    "## Tickets for family",
    "Your people can buy tickets here: [Eventbrite link].",
    "",
    "Any questions, go through your coach or reply here. See you on fight night.",
    "",
    ...SIG]),

  H1("6. Volunteer briefing"),
  WHEN("Send the week of the show and read it out at the volunteer meeting on the night."),
  ...box(null, [
    "## CCCL [Town] · Volunteer briefing",
    "Arrive at [time] at [entrance]. Ask for [crew lead first name]. Wear [black / the volunteer shirt].",
    "Your role and shift are on the list from Cody. If you're unsure where to be, find the crew lead.",
    "Radios: channel [number]. Keep talk short.",
    "Tickets: every person gets scanned or is on the comp list. [19+ area: wristbands only after a photo ID check. No wristband, no bar.]",
    "Fighters' areas: only fighters, corners, officials and medical staff in the dressing rooms and behind the cage.",
    "Youth athletes: never film or photograph a child unless the parent has said yes. Kids stay with their coach or parent.",
    "Emergencies: tell [security lead] or the physician right away. Don't move anyone who's hurt. Keep the ambulance path at [location] clear.",
    "Thank you. The night doesn't happen without you.",
  ]),

  H1("7. Announcer script"),
  WHEN("Print it for the announcer with the bout card. Sponsor reads follow each sponsor's level: check the Packages and Fulfilment tabs."),
  ...box(null, [
    "## Opening (about 5 minutes before bout 1)",
    "\"Good evening [Town]! Welcome to the Cottage Country Combat League, presented by [Title Sponsor]! Tonight: [number] bouts of amateur mixed martial arts, sanctioned by the Ontario Athletics Commission.\"",
    "\"A few things: stay in your seats during the bouts, keep the aisles clear, and cheer loud. Washrooms are [where]. [The licensed area is 19+ only.]\"",
    "## Sponsor reads by level",
    "Title: at the opening, before the main event and at the close. \"[Business], your Title Sponsor: [one line about them].\"",
    "Gold: two reads during the night. \"This bout is brought to you by [Business].\"",
    "Silver: one read. \"Thanks to [Business] for supporting tonight's show.\"",
    "Bronze: grouped read at intermission. \"Big thanks to our Bronze sponsors: [names].\"",
    "Community: grouped thank-you at the close. \"And thank you to our community supporters: [names].\"",
    "## Fighter introductions",
    "\"Fighting out of the blue corner, from [gym] in [town], with a record of [W-L], weighing in at [weight] pounds: [name]!\" Then the red corner the same way.",
    "## Results (only what the officials hand you)",
    "\"Ladies and gentlemen, we go to the judges' scorecards...\" or \"Your winner, by [method] at [time] of round [number]: [name]!\"",
    "## Intermission",
    "\"We'll take a [number]-minute break. Grab a drink, visit the merch table, and thank our sponsors: [Bronze names].\"",
    "## Close",
    "\"Thank you [Town]! Thanks to our fighters, coaches, officials, volunteers and [Title Sponsor]. Our next show is [Town] on [date]. Get home safe!\"",
  ]),

  H1("8. Press release: show announced"),
  WHEN("About 14 weeks out, to every outlet on the Press tab, through the way each one asks for news."),
  ...box("Amateur MMA returns: Cottage Country Combat League comes to [Town] on [date]", [
    "FOR IMMEDIATE RELEASE",
    "[Town], Ont., [date of release]: The Cottage Country Combat League (CCCL) brings sanctioned amateur mixed martial arts to [venue] in [Town] on [day, date]. Doors open at [time], first bout at [time].",
    "CCCL's first show, at the A.J. LaRue Arena in Haliburton in May 2026, sold 400 tickets with 15 local businesses as sponsors. The [Town] card will feature [number] amateur bouts [and youth grappling], with fighters from gyms across [region].",
    "\"[One or two sentences from Cody: why this town, what the night means for local athletes.]\" said Cody \"The Constrictor\" Purtell, CCCL promoter and owner of Bancroft Boxing & Grappling.",
    "Tickets: [prices] at [Eventbrite link]. Businesses interested in sponsoring can contact Cody at 705-760-0938.",
    "Media contact: Cody Purtell, 705-760-0938, ontario.cccl@gmail.com.",
  ]),

  H1("9. Press release: the card"),
  WHEN("About 4 weeks out, once fighters have agreed to be named. Never name anyone under 18 in a press release."),
  ...box("Fight card set for CCCL [Town] on [date]", [
    "FOR IMMEDIATE RELEASE",
    "[Town], Ont.: The Cottage Country Combat League has announced the card for [date] at [venue]. The main event features [fighter] ([gym], [record]) against [fighter] ([gym], [record]) at [weight class].",
    "Also on the card: [two or three more bouts, same format].",
    "\"[Quote from Cody about the card.]\"",
    "Tickets: [link]. The event will also be livestreamed at [link].",
    "Media contact: Cody Purtell, 705-760-0938, ontario.cccl@gmail.com.",
  ]),

  H1("10. Press release: results"),
  WHEN("The day after the show. Official results only, copied from the Bout Card tab once the officials confirm them."),
  ...box("Results: CCCL [Town]", [
    "FOR IMMEDIATE RELEASE",
    "[Town], Ont.: [Number] people filled [venue] on [date] for the Cottage Country Combat League's [number] show. In the main event, [winner] ([gym]) defeated [opponent] by [method] [at time, round].",
    "Full results:",
    "[Winner] def. [opponent], [method], round [number]",
    "[...one line per bout...]",
    "\"[Quote from Cody: thanks to the town, sponsors and volunteers; what's next.]\"",
    "CCCL returns [Town / date if set]. Media contact: Cody Purtell, 705-760-0938, ontario.cccl@gmail.com.",
  ]),
  H1("11. Fighter promo pack"),
  WHEN("To every confirmed adult fighter (through their coach) about 6 weeks out, with their own Eventbrite tracking link from the Marketing Kit's Ticket Sources tab and their \"I'm fighting\" post and story (starter designs 11 and 12). After the show, send their result post (design 14; for a loss, change WIN to THANK YOU)."),
  ...box("Your promo pack: CCCL [Town], [date]", [
    "Hi [first name],",
    "",
    "Your friends and family are the best crowd you'll ever have, so here's everything to bring them out.",
    "## Your ticket link",
    "[personal tracking link]. Every ticket bought through it counts as yours, and we'll shout out the fighters who bring the biggest crowds.",
    "## Three posts you can copy",
    "Post 1 (now): \"I'm fighting at the Cottage Country Combat League in [Town] on [date]! Come yell for me: [link]\"",
    "Post 2 (2 weeks out): \"Two weeks. Camp's almost done. Tickets for [Town] here: [link]\"",
    "Post 3 (fight week): \"[Day] night at [venue]. Doors at [time]. Last chance for tickets: [link]\"",
    "## Your graphics",
    "Attached: your \"I'm fighting\" post and story. On the story, add an Instagram link sticker with your link (Stickers > Link). Tag @[CCCL Instagram] so we can reshare. You'll get a fighter vs fighter graphic when the card is announced and your result post the day after.",
    "",
    "Keep it respectful toward your opponent: that's how we do it at CCCL. See you on fight night.",
    "",
    ...SIG]),

  H1("12. Gym members: group tickets"),
  WHEN("For the BBG members' chat or email, and for partner gyms whose fighters are on the card. About 12 weeks out, again at 4 weeks."),
  ...box(null, [
    "Fight night is coming to [Town] on [date], and [our fighters / fighters from [gym]] are on the card. Let's pack the place.",
    "Grab your tickets through our gym link so we can see how many of us are coming: [gym tracking link]. Bring your family: there's [youth grappling and] a great night for all ages.",
    "[Optional: groups of 10 or more, talk to Cody about a VIP table for 8-10.]",
  ]),

  H1("13. Sponsor share kit"),
  WHEN("When a sponsor's post goes up (about 4 weeks out), so they share it with their own customers. Their own tracking link shows how many tickets their followers bought."),
  ...box("Your CCCL [Town] post is live", [
    "Hi [first name],",
    "",
    "Your sponsor post is up on our Facebook and Instagram: [post link]. If you share it, here's a caption you can use:",
    "\"We're proud to back the Cottage Country Combat League in [Town] on [date]. Local fighters, a great night out. Tickets: [sponsor tracking link]\"",
    "Your \"Proud sponsor\" graphic with your logo is attached if you'd rather post it yourself. Tag @[CCCL Instagram] and we'll reshare.",
    "",
    ...SIG]),

  H1("14. Radio pitch"),
  WHEN("About 5 weeks out, to the local stations on the Press tab (Moose FM, Quinte News and others)."),
  ...box("Interview idea: amateur MMA comes to [Town] on [date]", [
    "Hi [name],",
    "",
    "I'm Cody Purtell, a Bancroft fighter who runs Bancroft Boxing & Grappling and the Cottage Country Combat League. On [date] we're bringing sanctioned amateur MMA to [venue] in [Town], with [number] bouts [and youth grappling].",
    "Happy to come in or call in for five minutes: what a local fight night is like, the athletes from [region] on the card, and how it's run safely with the Ontario Athletics Commission. A fighter from the card could join too.",
    "",
    ...SIG]),

  H1("15. Community calendar listing"),
  WHEN("For kawarthaNOW, Moose FM, town calendars and Facebook event pages. Week 14."),
  ...box(null, [
    "## Short (one line)",
    "Cottage Country Combat League: sanctioned amateur MMA at [venue], [Town]. [Day, date], doors [time]. Tickets: [link]",
    "## Long",
    "Live amateur mixed martial arts comes to [Town]! The Cottage Country Combat League brings [number] bouts of amateur MMA, sanctioned by the Ontario Athletics Commission, to [venue] on [day, date]. Doors open at [time], first bout at [time]. Local fighters, a full crowd and a great night out. General admission $[price], VIP $[price], VIP tables for 8-10. Tickets: [link]",
  ]),

  H1("16. Eventbrite event description"),
  WHEN("Paste into the Eventbrite page when tickets go on sale. Check the prices, times and age policy first."),
  ...box(null, [
    "## Title",
    "Cottage Country Combat League: Fight Night [Town]",
    "## Summary",
    "Sanctioned amateur MMA live at [venue], [Town]. [Day, date]. Doors [time].",
    "## About",
    "The Cottage Country Combat League brings professionally run amateur mixed martial arts to Ontario's cottage country. Our first show at the A.J. LaRue Arena in Haliburton sold 400 tickets. Now we're coming to [Town].",
    "[Number] amateur MMA bouts [plus youth grappling], sanctioned by the Ontario Athletics Commission, with fighters from gyms across [region].",
    "## Tickets",
    "General admission $[price]. VIP seat $[price]. VIP tables seat 8-10 cage-side: $1,500 (message us for business tables).",
    "## Good to know",
    "Doors [time], first bout [time], ends about [time]. [All ages / 19+ in the licensed area.] Parking: [details]. Can't make it? Watch the livestream: [link].",
    "Questions: ontario.cccl@gmail.com or 705-760-0938.",
  ]),

];

const doc = new Document({
  numbering: { config: [{ reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 260 } } } }] }] },
  styles: { default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [{ id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 28, bold: true, color: INK }, paragraph: { outlineLevel: 0 } }] },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 960, bottom: 960, left: 1260, right: 1260 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      new TextRun({ text: "CCCL fight night templates  ·  page ", font: FONT, size: 16, color: "888888" }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: "888888" })] })] }) },
    children }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("saved", OUT); });
