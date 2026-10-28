// Second Chance: policy drafts + first letters (two Word docs) and the matching second-chance skill references.
// usage: NODE_PATH=$(npm root -g) node sc_docs.js POLICIES.docx LETTERS.docx SKILL_REF_DIR
const fs = require("fs"), path = require("path");
const mk = () => require("./docstyle.js")({ navy: "12343B", gold: "B8732E", goldLight: "F6E9D8", goldOnNavy: "F0C38E" });
const S = mk(), S2 = mk();  // separate section numbering for each doc
const { R, P, Bd, H1, H2, bullet, num, chk, table, callout, sp, cover, save, answer } = S;
const [POL, LET, REF] = process.argv.slice(2);

// ---------------- policies: [title, purpose, [ [heading, [points]] ... ]]
const POLICIES = [
  ["Screening", "How Second Chance chooses and checks the people who help run it, and how it decides what a participant can do. Fair to everyone, safe for young people.", [
    ["Who it covers", ["Directors, staff, volunteers, mentors, and participants doing a placement or a fight-night role."]],
    ["The steps", [
      "Write down what the role involves and whether it includes anyone under 18 or other vulnerable people.",
      "Directors, mentors and staff: a short application, an interview, and two references.",
      "A police check only when the role needs one. A vulnerable sector check is only for roles with trust or authority over people under 18 or other vulnerable people. The OPP lists volunteer checks as free (confirm with the Bancroft detachment).",
      "The person consents to any check, sees the result first, and decides whether to share it.",
      "Orientation before starting: these four policies, the Rule of Two, and who to call.",
      "Check-ins: a mentor or Cody checks in at least once a month. Every role is reviewed once a year.",
    ]],
    ["Participants", [
      "Each person is considered one at a time. No blanket bans because someone has a record.",
      "In jobs, Ontario's Human Rights Code protects people with a pardoned record of offences. Ask a lawyer before treating a record as a reason to say no.",
      "We record only the date of a check and the outcome as Cleared or Role-limited, plus that written consent is on file. Never what someone was charged with or convicted of.",
      "Role-limited means no role with young people present, and nothing else is assumed.",
    ]],
    ["Who decides", ["Two directors together. If a decision involves BBG or CCCL, Cody isn't one of them."]],
  ]],
  ["Rule of Two and young people", "Keeping young people safe whenever Second Chance is in the same place as them, especially at fight nights and the gym.", [
    ["Always", [
      "Two screened adults with anyone under 18, in the open where others can see. Never one adult alone with a young person, out of sight.",
      "No private messages to anyone under 18. Group messages that include a parent only.",
      "Participants never have authority over young athletes and are never alone with one.",
      "Youth warm-up areas and youth athletes' rooms at fight nights are for staff and coaches only.",
    ]],
    ["Fight-night roles that fit", ["Setup and teardown, chairs and barriers, the merch table with staff, cleanup, runner for the production crew, ticket scanning alongside staff."]],
    ["Not without more", ["Any role with youth contact needs a vulnerable sector check that came back Cleared, plus the Rule of Two. Security work needs Ontario's security licensing rules checked first."]],
    ["If something worries you", ["Stop it, make everyone safe, and tell Cody and the board chair the same day. Write down what happened (facts only) in the Second Chance folder."]],
  ]],
  ["Privacy", "What Second Chance writes down about people, where it's kept, and who sees it. Based on the 10 fair information principles that Canada's privacy law sets out, which nonprofits are encouraged to follow.", [
    ["What we keep", [
      "In the tracker: first name, a code (SC-001), who referred them, stage, consent on file, check date and outcome, roles allowed, placement, mentor's first name, check-in dates.",
      "Contact details stay in Second Chance's own email contacts, not in the tracker.",
    ]],
    ["What we never write down", ["What anyone was charged with or convicted of, health details, addresses, or anything about their family."]],
    ["Consent", ["Written consent before anything is recorded or shared, saying what and with whom. People can see what we hold about them and ask us to fix it."]],
    ["Who sees it", ["Cody and the board. A mentor sees only their own participant's first name and schedule. Never BBG or CCCL staff, and never in a BBG or CCCL file."]],
    ["Where it lives", ["Second Chance's own Google Drive folder and email only. No copies on personal phones beyond contacts, and no printouts left out."]],
    ["Sharing", ["Only with the person's consent, or where the law requires it. [Board: ask a lawyer what that covers.]"]],
    ["Photos and stories", ["Only with written consent for that exact use, first name at most, and never a last name."]],
    ["How long we keep it", ["[Board: set how long records are kept after someone leaves, then delete them.]"]],
    ["If something goes wrong", ["If information gets to the wrong person, tell the person affected and the board chair right away, and fix what caused it."]],
  ]],
  ["Conflict of interest", "So decisions are made for Second Chance's purposes, not for anyone's business or family.", [
    ["Declaring", ["Every director declares their interests in writing when they join and once a year. Cody declares that he owns BBG and CCCL."]],
    ["At meetings", ["A director with a conflict says so, leaves the discussion and doesn't vote. The minutes record it."]],
    ["Money to a director's business", ["Any payment to BBG, CCCL or another director's business (gym time, event fees, supplies) needs approval from the other directors, and Ontario has extra rules for charities paying a director's business. Ask a lawyer before any such payment."]],
    ["Placements", ["CCCL fight nights can be one placement among several, never the only one. Training must build skills people can use with any employer."]],
    ["No private benefit", ["Nobody gets a personal benefit from Second Chance beyond something small and incidental."]],
  ]],
];

// ---------------- letters: [to, why, subject, body paragraphs]
const SIGN = ["Thanks,", "Cody Purtell", "Founder, Second Chance (Bancroft)", "705-760-0938 · [Second Chance email]"];
const LETTERS = [
  ["John Howard Society of the Kawarthas", "They serve Bancroft and Haliburton County and already run job training and record suspensions. A referral partner from day one.",
   "Second Chance: a new reintegration program in Bancroft",
   ["Hi [name],",
    "I'm Cody Purtell. I own Bancroft Boxing & Grappling and run the Cottage Country Combat League, and I'm starting Second Chance, a nonprofit in Bancroft that helps people with criminal records get back to work through job skills, work experience, mentoring and community support.",
    "Your team already does this work across our area, and I'd rather work alongside you than start from scratch. Could we meet for 30 minutes in the next few weeks? I'd like to learn how your Get SET training and referrals work, and whether Second Chance could be a useful next step for people you support.",
    "Second Chance is being incorporated now. We're setting up its board, screening and privacy policies before we take anyone on."]],
  ["Loyalist College Community Employment Services (Bancroft)", "The Employment Ontario office in town: job matching and Better Jobs Ontario. The job-connection step.",
   "Working together on jobs for people with records",
   ["Hi [name],",
    "I'm Cody Purtell from Bancroft Boxing & Grappling. I'm starting Second Chance, a local nonprofit that helps people with criminal records rebuild through job training, work experience and mentoring.",
    "Your office is where people in Bancroft go to find work, so I'd like Second Chance to send people to you for job matching and Better Jobs Ontario, and to hear what local employers need. Could we meet for half an hour at your office on Hastings Street?",
    "We're early: incorporation is under way and we're lining up partners before we take anyone on."]],
  ["North Hastings Community Trust", "A registered charity in Bancroft with grant experience. Could lead a United Way application (due Dec 11) while Second Chance waits for charity status.",
   "Partnering on a United Way application",
   ["Hi [name],",
    "I'm Cody Purtell. I'm starting Second Chance, a Bancroft nonprofit that helps people with criminal records get back to work and back into the community.",
    "United Way Hastings & Prince Edward's community fund is due December 11, and it needs a registered charity to apply. Second Chance won't have charity status in time. Would the Trust consider leading an application with Second Chance delivering the program? We'd do the writing and the reporting, and the program would be [Cody: one line, e.g. a 10-week job-skills and work-experience pilot for 6 people].",
    "If it's not a fit, I'd still be grateful for 20 minutes of advice on getting started."]],
  ["A board member (one letter each)", "Invite each of the five people Cody named. The questions at the end check CRA's arm's-length rule.",
   "Would you join the Second Chance board?",
   ["Hi [first name],",
    "I'm starting Second Chance, a nonprofit that helps people with criminal records rebuild their lives through job training, work experience and mentoring. I'd like you on the first board of directors.",
    "What it involves: [Cody: e.g. one evening meeting a month for the first six months, then every two months], reading short updates, and helping make decisions about policies, money and partners. Directors aren't paid. You'd get a vulnerable sector check through the OPP (it lists volunteer checks as free) and sign our conflict-of-interest policy, like everyone else.",
    "Two quick questions, because the charity rules need most of the board to be independent: are you related to me or to anyone else on the board by blood, marriage, common-law partnership or adoption? And do you work for me? Either answer is fine; it just decides how we set the board up.",
    "Could you let me know by [date]?"]],
  ["A local business (placement host)", "For businesses that might host a short work placement. Several hosts keep the training fair and keep charity status safe.",
   "Would you host a short work placement?",
   ["Hi [name],",
    "I'm Cody Purtell from Bancroft Boxing & Grappling. I'm starting Second Chance, a local nonprofit that helps people with criminal records get back to work.",
    "Would [business] consider hosting a short work placement: [Cody: e.g. 2 to 4 weeks, a few days a week], with a Second Chance mentor checking in and nothing for you to pay? It gives someone real experience and a reference, and gives you a look at a possible hire.",
    "Happy to drop by and explain. Nothing is promised either way."]],
];

// ---------- Policies doc
const pol = [
  cover("SECOND CHANCE", "POLICY DRAFTS", "Screening, the Rule of Two, privacy and conflict of interest. For the board to adopt after a lawyer reviews them.", []),
  sp(120),
  callout("How to use these", [
    "They're short on purpose: one page each, plain words. The board reads them, a lawyer reviews them, and the board adopts them at its first meeting (the minutes say so).",
    "Yellow brackets are decisions only the board can make.",
    "Research, not legal advice. Sources are in the Second Chance Starter Plan.",
  ], "EEF3F3", "12343B"),
];
POLICIES.forEach(([title, purpose, secs], i) => {
  pol.push(...H1(title + " policy", purpose, i > 0));
  secs.forEach(([h, pts]) => { pol.push(H2(h)); pts.forEach(t => pol.push(bullet(t))); });
  pol.push(sp(60), P([R("Draft for the Second Chance board and a lawyer to review. Adopted on: [date]   Review by: [date]", { size: 17, color: "5A6672", italics: true })]));
});

// ---------- Letters doc
const let_ = [
  S2.cover("SECOND CHANCE", "FIRST LETTERS", "Five drafts for Cody to send: two referral partners, a charity to lead a grant, the board, and placement hosts.", []),
  sp(120),
  callout("Before sending", [
    "Fill the yellow brackets, read it out loud once, then send it yourself from Second Chance's own email (or ontario.cccl@gmail.com until it has one).",
    "Find the right person's name on each organization's website or by phone. Never promise money, jobs or charity status.",
    "Log each one on the Partners tab of the Second Chance Tracker with the date and the next step.",
  ], "EEF3F3", "12343B"),
];
LETTERS.forEach(([to, why, subj, body], i) => {
  let_.push(...S2.H1(to, why, i > 0));
  let_.push(answer("SUBJECT", subj), sp(40), answer("EMAIL", [...body, "", ...SIGN].filter((x, k, a) => !(x === "" && a[k - 1] === "")).map(x => x || " ")));
});

Promise.all([save(POL, pol, "Second Chance policy drafts"), S2.save(LET, let_, "Second Chance first letters")]).then(() => {
  const P1 = ["# Second Chance policy drafts", "", "The same text as 'Second Chance Policy Drafts' (Word). Drafts for the board and a lawyer to review; never legal advice.", ""];
  POLICIES.forEach(([t, p, secs]) => { P1.push(`## ${t} policy`, "", p, ""); secs.forEach(([h, pts]) => { P1.push(`### ${h}`, ...pts.map(x => `- ${x}`), ""); }); });
  fs.writeFileSync(path.join(REF, "policies.md"), P1.join("\n"));
  const L1 = ["# Second Chance first letters", "", "Drafts for Cody to send himself. Fill the [brackets]; never promise money, jobs or charity status.", ""];
  LETTERS.forEach(([to, why, subj, body]) => { L1.push(`## ${to}`, "", `*${why}*`, "", `**Subject:** ${subj}`, "", ...body.map(b => `> ${b}\n>`), ...SIGN.map(s => `> ${s}  `), ""); });
  fs.writeFileSync(path.join(REF, "letters.md"), L1.join("\n"));
  console.log("refs written");
});
