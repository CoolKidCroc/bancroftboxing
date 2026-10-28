// Second Chance Starter Plan (Word), from research/second_chance_plan.json.
// usage: NODE_PATH=$(npm root -g) node second_chance.js OUT.docx research/second_chance_plan.json
const fs = require("fs");
// Second Chance gets its own look: deep teal and warm sand, so it's never mistaken for a gym or CCCL document
const S = require("./docstyle.js")({ navy: "12343B", gold: "B8732E", goldLight: "F6E9D8", goldOnNavy: "F0C38E" });
const { R, P, Bd, H1, H2, bullet, num, chk, table, callout, sp, cover, save } = S;
const OUT = process.argv[2], D = JSON.parse(fs.readFileSync(process.argv[3], "utf8"));

const children = [
  cover("SECOND CHANCE", "STARTER PLAN", "How Second Chance gets set up, who it helps, how it keeps people safe and private, and where the money can come from.", []),
  sp(140),
  callout("What Cody told us (October 20, 2026)", [
    [Bd("Mission: "), D.about.mission],
    [Bd("For: "), D.about.who],
    [Bd("What it does: "), D.about.does],
    [Bd("Where it stands: "), D.about.status],
  ]),
  sp(80),
  callout("Three things to get right from day one", [
    [Bd("Keep it separate. "), "Its own bank account, its own books, its own Drive folder. Participants' information never goes in a BBG or CCCL file."],
    [Bd("Don't tie it to one employer. "), "CRA wants training, not jobs, and not for one employer. Fight-night roles at CCCL are fine as one placement among several, never the only one."],
    [Bd("Cody declares his businesses. "), "He owns BBG and CCCL, so he steps out of any Second Chance decision that involves them."],
  ], "EEF3F3", "12343B"),

  ...H1("The first 90 days", "In order. Every step is on the Setup tab of the Second Chance Tracker."),
  table([1100, 5720, 1300, 1600], [["Area", "Step", "Who", "When"], ...D.checklist], { top: true }),

  ...H1("The board", "Cody named five people (first names only in our files)."),
  P([Bd(D.board.join(", ")), ". Second Chance needs at least 3 directors under Ontario's law; with 5 it can keep meeting when someone's away."]),
  bullet("More than half must be at arm's length: not related to each other or to Cody by blood, marriage, common-law partnership or adoption, and not acting together on Cody's behalf. Otherwise CRA treats it as a private foundation."),
  bullet("No more than one-third of the directors can be employees of Second Chance."),
  bullet("Each director agrees in writing, gets a vulnerable sector check (the OPP lists volunteer checks as free; confirm with the Bancroft detachment), and signs the conflict-of-interest policy."),
  bullet("Suggested roles: a chair, a treasurer (keeps the books and the bank account) and a secretary (minutes). [Cody: who does what]"),

  ...H1("How the program works", "A pathway a person can follow from first meeting to a job."),
  table([2200, 7520], [["Step", "What happens"], ...D.pathway], { label: true }),
  sp(60),
  H2("Fight-night roles"),
  table([2000, 7720], [["", "Roles"], ...D.fight_night_roles], { label: true, top: true }),
  sp(60),
  P(["Placement hosts to ask beyond CCCL ", R("[Cody: ask first; none are confirmed]", { highlight: "yellow" }), ": local businesses that already know Cody, such as the 2026 CCCL sponsors (a millwright and welding shop, a resort, a restaurant, a sign shop), plus the partners below."]),

  ...H1("Safety, fairness and privacy"),
  H2("Screening and the Rule of Two"),
  bullet("Police checks come in three types: criminal record; criminal record and judicial matters; and vulnerable sector. A vulnerable sector check is only for roles with trust or authority over people under 18 or other vulnerable people."),
  bullet("The person consents, gets the result first, and decides whether to share it."),
  bullet("Two adults with any participant under 18, in the open (Rule of Two). Participants are never alone with a youth athlete."),
  bullet("Decide each person one at a time, never with a blanket ban. In jobs, Ontario's Human Rights Code protects people with a pardoned record of offences."),
  H2("Privacy"),
  ...D.privacy.map(p => bullet(p)),
  H2("Charity status cautions"),
  ...D.charity_cautions.map(p => bullet(p)),

  ...H1("Partners near Bancroft", "Possible referral partners, co-applicants or a charity to lead a grant while Second Chance gets established."),
  table([3300, 4820, 1600], [["Organization", "What they do", "Where"], ...D.partners.map(([a, b, c, d]) => [S.lines([Bd(a, { size: 19 })], [R(d.replace(/^https?:\/\//, "").replace(/\/$/, ""), { size: 16, color: "5A6672" })]), b, c])], { top: true }),

  ...H1("Where the money can come from", "Most grants need incorporation first, and some need charity status. A registered charity partner can also grant money to Second Chance."),
  table([2900, 5020, 1800], [["Program", "What and when", "Needs"], ...D.funding], { top: true }),
  sp(60),
  H2("First-year costs we know"),
  table([5720, 4000], [["Item", "Cost"], ...D.budget]),

  ...H1("What Claude can do", "With the second-chance skill installed and the Second Chance Tracker in its own Drive folder."),
  bullet([Bd("\"Where are we with Second Chance?\" "), "Reads the Setup tab and says what's done, what's next and what's overdue."]),
  bullet([Bd("\"Draft a letter to [partner]\" "), "A short email introducing Second Chance and asking for a meeting. Draft only; Cody sends it."]),
  bullet([Bd("\"Draft the screening policy\" "), "(or privacy, conflict of interest, Rule of Two) for the board and a lawyer to review."]),
  bullet([Bd("\"Plan a fight-night placement for [first name]\" "), "Picks roles from the safe list, with no youth contact unless the checks are done."]),
  bullet([Bd("\"Draft the United Way application with [partner]\" "), "Uses the funding skill and the facts here. Never invented numbers."]),
  sp(40),
  P([R("Sources (checked Oct 20, 2026): " + D.sources.map(s => s[0]).join("; ") + ". Research, not legal or accounting advice: a lawyer should review the purposes, by-law and policies, and an accountant the books.", { size: 16, color: "5A6672", italics: true })]),
];
save(OUT, children, "Second Chance starter plan");
