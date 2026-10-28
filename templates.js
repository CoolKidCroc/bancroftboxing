const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, BorderStyle, Footer, PageNumber, ShadingType, LevelFormat } = require("docx");
const OUT = process.argv[2];
const NAVY = "081C2B", GOLD = "C8862E", FONT = "Arial";

const T = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 21, ...o });
const P = (runs, o = {}) => new Paragraph({ spacing: { after: 100 }, ...o, children: (Array.isArray(runs) ? runs : [runs]).map(x => typeof x === "string" ? T(x) : x) });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, keepNext: true, spacing: { before: 320, after: 80 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: GOLD, space: 2 } },
  children: [new TextRun({ text: t, font: FONT, size: 28, bold: true, color: NAVY })] });
const WHEN = (t) => P([T(t, { italics: true, color: "666666", size: 19 })], { spacing: { after: 120 }, keepNext: true });
const SIG = ["Thanks,", "Cody Purtell", "Cottage Country Combat League  ·  Bancroft Boxing & Grappling", "705-760-0938"];
// an email: subject + body lines inside a shaded box (paragraph shading so it copies cleanly)
function email(subject, lines) {
  const sh = { type: ShadingType.CLEAR, color: "auto", fill: "F4F6F8" };
  const ind = { left: 240, right: 240 };
  const out = [];
  if (subject) out.push(new Paragraph({ shading: sh, indent: ind, keepNext: true, spacing: { before: 60, after: 140 },
    children: [T("Subject: ", { bold: true }), T(subject)] }));
  const sig = new Set(SIG);
  const real = lines.filter(l => l !== "");
  real.forEach((l, i) => out.push(new Paragraph({ shading: sh, indent: ind, keepNext: i < real.length - 1, keepLines: true,
    spacing: { after: sig.has(l) && l !== SIG[0] ? 0 : (l === SIG[0] ? 40 : 140) }, children: [T(l)] })));
  out.push(new Paragraph({ spacing: { after: 160 }, children: [] }));
  return out;
}

const children = [
  new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "CCCL SPONSOR PITCH TEMPLATES", font: FONT, size: 36, bold: true, color: NAVY })] }),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "Copy, fill in the [brackets], send from ontario.cccl@gmail.com.", font: FONT, size: 22, color: GOLD, bold: true })] }),
  P("Each one is short on purpose. Local business owners read emails on their phone between jobs. Before sending, swap every [bracket] for the real detail, and read it out loud once so it sounds like you."),
  P([T("House rules: ", { bold: true }), T("Cody sends every pitch himself, from ontario.cccl@gmail.com. Never promise a perk that isn't on the Packages tab: a $1,000 Bronze Sponsor gets their logo on ads, not their own banner by the cage. Log every send on the Pipeline tab and set the next follow-up date.")]),

  H1("1. First pitch: a business we haven't worked with"),
  WHEN("Send to Prospects. Set the next follow-up for 6 days later."),
  ...email("Put [Business] cage-side at CCCL [Town] this [Month]", [
    "Hi [first name or 'there'],",
    "",
    "I'm Cody Purtell. I run Bancroft Boxing & Grappling and the Cottage Country Combat League, the amateur MMA show that ran its first night at the A.J. LaRue Arena in Haliburton in May 2026.",
    "",
    "We sold 400 tickets to that first show and had 15 local businesses behind it. We're back in 2027 with [number] shows, starting in [Town] on [date]. It's sanctioned by the Ontario Athletics Commission, livestreamed, and the crowd is local: families, trades, cottagers and a lot of people who [one line about why their customers fit, e.g. 'ride the trails you service'].",
    "",
    "I'd love to have [Business] as a Bronze Sponsor. For $1,000 you get your logo on our ads, the livestream and the event shirt, a sponsor announcement on our social media, a live mention from the announcer on fight night and free tickets. Packages run from $300 up to Title Sponsor, and they're all on the attached sheet.",
    "",
    "Could I drop by for ten minutes next week, or call you? Whatever's easier.",
    "",
    ...SIG]),

  H1("2. First pitch: an existing gym supporter"),
  WHEN("For businesses with a banner on the gym wall. Thank them first."),
  ...email("Thank you, and a first look at CCCL 2027", [
    "Hi [first name],",
    "",
    "Every day I train under the [Business] banner at the gym, so first: thank you. Your support is a big part of why Bancroft Boxing & Grappling exists.",
    "",
    "Our first Cottage Country Combat League show in Haliburton in May 2026 went really well, and we're planning [number] shows for 2027, starting in [Town] on [date]. I wanted you to have first pick of the sponsor spots before I go to anyone else.",
    "",
    "Bronze is $1,000: your logo on our ads, the livestream and the event shirt, a social media announcement, a mention from the announcer and free tickets. If you'd like your banner at the show, Silver ($2,000) adds shared banner space and Gold ($4,000) puts your logo right around the cage. The full sheet is attached.",
    "",
    "Can I swing by this week to walk you through it?",
    "",
    ...SIG]),

  H1("3. First pitch: a 2026 sponsor"),
  WHEN("For the 15 businesses that backed the first show. Send these before any new pitches go out."),
  ...email("Thank you for 2026, and first pick for 2027", [
    "Hi [first name],",
    "",
    "[Business] was one of the 15 businesses that made our first fight night in Haliburton happen[, and as our Title Sponsor, the biggest one]. Thank you. We sold 400 tickets, and your name was in front of every one of those people.",
    "",
    "We're planning [number] shows for 2027, starting in [Town] on [date]. Before I offer anything to anyone else, I wanted to give you first pick of the same package, [Package] at $[amount]. If you'd like to be part of more than one show, let's talk about that too.",
    "",
    "Can I call you this week?",
    "",
    ...SIG]),

  H1("4. Follow-up 1: about a week later"),
  WHEN("No reply after 6-7 days. Keep it to three lines."),
  ...email("Re: Put [Business] cage-side at CCCL [Town]", [
    "Hi [first name],",
    "",
    "Just bumping this up in case it got buried. We're locking in sponsors for [Town] on [date], and I'd like [Business] to be part of it. Happy to call or stop in, whatever works for you.",
    "",
    ...SIG]),

  H1("5. Follow-up 2: last nudge"),
  WHEN("Two weeks after follow-up 1. After this, set the stage to 'Not now' and try again next season."),
  ...email("Last call for [Town] fight night sponsors", [
    "Hi [first name],",
    "",
    "Last note from me on this one. Banners for [Town] go to print on [date], so if you'd like [Business] up there, just reply 'in' and I'll take care of the rest. If the timing's not right, no worries at all. I'll check in before the next show.",
    "",
    ...SIG]),

  H1("6. Phone or walk-in script"),
  WHEN("Thirty seconds. Then stop talking and let them ask questions."),
  ...email(null, [
    "\"Hi, I'm Cody from Bancroft Boxing & Grappling. I run the Cottage Country Combat League, the amateur MMA night we held at the arena in Haliburton in May.",
    "Our first show sold 400 tickets, and we're doing [number] shows in 2027. I'm looking for a few local businesses to back it. For a thousand dollars you get your logo on our ads, the livestream and the event shirt, a shout-out on our social media and from the announcer, and free tickets. Community spots start at three hundred, and the bigger packages put your banner by the cage.",
    "Would you be open to it? I can leave you the sheet.\"",
  ]),

  H1("7. Facebook or Instagram message"),
  WHEN("For businesses that answer messages faster than email."),
  ...email(null, [
    "Hey! Cody here from Bancroft Boxing & Grappling 🥊 We're bringing the Cottage Country Combat League back in 2027, starting in [Town] on [date], and we're looking for a few local sponsors. $1,000 gets [Business] its logo on our ads, the livestream and the event shirt, a social media shout-out, a mention from the announcer and free tickets (community spots from $300, bigger packages with banners by the cage). Want me to send the details?",
  ]),

  H1("8. They said yes: confirmation"),
  WHEN("Same day they say yes. Then set Stage to 'Signed' and add them to the Fulfilment tab."),
  ...email("You're in: [Business] x CCCL [Town]", [
    "Hi [first name],",
    "",
    "Awesome, thank you! [Business] is confirmed as a [Package] for CCCL [Town] on [date].",
    "",
    "Two things I need from you:",
    "1. Your logo, as a high-quality file (PNG or PDF), by [date] so the banner can go to print.",
    "2. Payment of $[amount] by [date], by [e-transfer to (email) / cheque to (name) / invoice: pick one].",
    "",
    "Your [number] tickets will come by email the week of the show. I'll tag you in our posts as we build up to fight night. [Title, Gold and Silver: we'll also need a banner file, or your banner itself, by (date).]",
    "",
    ...SIG]),

  H1("9. After the show: thank you"),
  WHEN("Within three days of the show, with a photo of their banner. This is what gets them back next year."),
  ...email("Thank you from CCCL [Town]", [
    "Hi [first name],",
    "",
    "Thank you for backing CCCL [Town]. Here's a photo from fight night [with your banner / your logo on the screen]: [attach photo]. We had [crowd number] people in the building and [number] watching the livestream, and your name was part of all of it.",
    "",
    "We're planning the next show for [Town] on [date]. You'll hear from me first.",
    "",
    ...SIG]),

  H1("10. National brand (their sponsorship form or email)"),
  WHEN("For brands like Hayabusa, BioSteel or MuscleTech. Most want their own form: paste this into it, trimmed to fit. Ask for product first. Alcohol brands: 19+ shows or areas only, and only after Cody checks with the AGCO and the venue."),
  ...email("Ontario amateur MMA: [Brand] x Cottage Country Combat League 2027", [
    "Hi [Brand] team,",
    "",
    "I'm Cody Purtell, a former pro fighter and the promoter of the Cottage Country Combat League (CCCL), amateur MMA sanctioned by the Ontario Athletics Commission. Our first show in Haliburton in May 2026 sold 400 tickets with 15 local sponsors behind it. In 2027 we're running [number] shows across Ontario, including [Town] on [date].",
    "",
    "Our crowd is [one line on why it fits them, e.g. 'young athletes and fight fans who train', 'trades and rural families'], and our fighters come from gyms across [region].",
    "",
    "We'd love [Brand] on board as [what you're asking for: e.g. 'the official fight-gear partner', 'product for every fighter bag and a sampling table at each show']. In return: [the perks that match, e.g. logo on the cage banner and livestream, announcer mentions, social posts, VIP seats].",
    "",
    "Happy to send our sponsor sheet and talk through what would work for you.",
    "",
    ...SIG]),
];

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [{ id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: FONT, size: 28, bold: true, color: NAVY }, paragraph: { spacing: { before: 320, after: 80 }, outlineLevel: 0 } }] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, bottom: 1080, left: 1260, right: 1260 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      new TextRun({ text: "CCCL sponsor pitch templates  ·  page ", font: FONT, size: 16, color: "888888" }),
      new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: "888888" })] })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("saved", OUT); });
