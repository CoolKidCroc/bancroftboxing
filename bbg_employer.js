// BBG First Employee Basics (Word): what a first-time Ontario employer needs before the summer student starts,
// with draft health & safety and workplace violence/harassment policies. Facts from research/first_employee.md.
// usage: NODE_PATH=$(npm root -g) node bbg_employer.js OUT.docx BBG_LOGO
const S = require("./docstyle.js")();
const { R, P, Bd, H1, H2, bullet, num, chk, table, callout, sp, cover, save, answer } = S;
const OUT = process.argv[2], BBG = process.argv[3];
const F = require("./research/first_employee_facts.json");   // facts checked against the sources (see that file)

const children = [
  cover("BANCROFT BOXING & GRAPPLING", "FIRST EMPLOYEE BASICS", "What BBG needs in place before its first paid worker (the Canada Summer Jobs student): safety, harassment, posters, pay and WSIB. Drafts for Cody to adjust.", [BBG]),
  sp(140),
  callout("The short version", F.short.map(t => t)),

  ...H1("Checklist", "Tick these off between now and the student's first day."),
  table([3400, 4520, 1800], [["What", "How", "When"], ...F.checklist], { top: true }),

  ...H1("Health and safety policy", "Draft. Cody signs and dates it, posts it where staff can see it, and reviews it once a year.", true),
  answer("BANCROFT BOXING & GRAPPLING: HEALTH AND SAFETY POLICY", [
    "Bancroft Boxing & Grappling is committed to a safe and healthy workplace for everyone who works, coaches, volunteers and trains here. We follow Ontario's Occupational Health and Safety Act and its regulations.",
    "Cody Purtell, the owner, is responsible for health and safety: providing training and equipment, fixing hazards, and making sure this policy is followed. Workers and volunteers work safely, follow these rules, wear the equipment required, and report hazards, injuries and near-misses right away.",
    "Our main hazards and how we control them:",
    "• Mats and equipment: mats are cleaned and sanitized every day; bags, anchors, pads and ropes are checked weekly and taken out of use if damaged.",
    "• Slips, trips and lifting: keep floors clear and dry; two people move heavy items like mats and bags.",
    "• Heat and hydration: water breaks in every session; extra breaks on hot days.",
    "• Cleaning products: read the label and the safety data sheet; anyone using them has WHMIS training; store them away from children.",
    "• Head injuries: anyone with signs of a concussion stops for the day and the head-injury steps are followed (Rowan's Law). Cody is the designated person.",
    "• Working with children: two screened adults with anyone under 18 (Rule of Two); no one alone with a child out of sight.",
    "• Emergencies: call 911 and give the address, 17 Snow Road, Bancroft. The first aid kit is at [Cody: where]. [Cody: name] holds current first aid training.",
    "Every worker completes Ontario's free health and safety awareness training on their first day. Report every injury to Cody the same day; it's recorded and, where required, reported to WSIB.",
    "Signed: ____________________  Cody Purtell, Owner      Date: ____________   Review by: ____________",
  ]),

  ...H1("Workplace violence and harassment policy", "Draft. Ontario requires every employer to have these policies and a written program for harassment. Cody signs, posts and reviews them yearly.", true),
  answer("BANCROFT BOXING & GRAPPLING: WORKPLACE VIOLENCE AND HARASSMENT POLICY", [
    "Bancroft Boxing & Grappling does not tolerate workplace violence or workplace harassment, including sexual harassment, from anyone: owner, staff, volunteers, members, parents or visitors. Training and sparring under a coach's rules is sport, not violence; anything outside those rules is not.",
    "Workplace harassment means a course of comments or conduct against a worker in the workplace that is known, or ought reasonably to be known, to be unwelcome. Workplace sexual harassment includes unwelcome comments or conduct because of sex, sexual orientation, gender identity or gender expression, and unwelcome sexual advances by someone in a position to give or deny a benefit. Reasonable direction from a supervisor is not harassment.",
    "Workplace violence means physical force, or an attempt or threat of physical force, against a worker in the workplace that causes or could cause injury.",
  ]),
  sp(40),
  answer("THE PROGRAM: HOW TO REPORT, HOW WE INVESTIGATE", [
    "How to report: tell Cody Purtell in person, by phone (705-760-0938) or in writing. If the concern is about Cody, report to [Cody: a named person outside the business, e.g. a trusted adult the worker can reach, with their phone or email].",
    "Violence or a threat: if anyone is in danger, call 911 first, then tell Cody. Members or visitors who become violent are asked to leave and may lose their membership.",
    "Investigation: every report is looked into promptly and fairly. The people involved are interviewed and notes are kept. If Cody is involved, the person named above arranges an outside investigator.",
    "Privacy: information is kept confidential and shared only as needed to investigate, take action, or as the law requires.",
    "Results: the worker who reported and the person accused are both told the results in writing, and what's being done about it.",
    "No reprisal: nobody is punished for reporting in good faith or for helping with an investigation.",
    "This policy and program are reviewed at least once a year, and every worker reads them on their first day.",
    "Signed: ____________________  Cody Purtell, Owner      Date: ____________   Review by: ____________",
  ]),

  ...H1("Pay, hours and WSIB", "What applies to a summer student in Ontario.", true),
  table([2600, 7120], [["Topic", "What applies"], ...F.pay], { label: true, top: true }),
  sp(80),
  P([R("Sources (checked " + F.checked + "): " + F.sources.join("; ") + ". Research, not legal or tax advice: ask WSIB, CRA or a lawyer about anything that matters.", { size: 16, color: "5A6672", italics: true })]),
];
save(OUT, children, "BBG first employee basics");
