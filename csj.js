// Canada Summer Jobs 2027: BBG application draft (Word).
// usage: NODE_PATH=$(npm root -g) node csj.js OUT.docx BBG_LOGO
const S = require("./docstyle.js")();
const { R, P, Bd, H1, H2, bullet, num, chk, table, callout, sp, answer, cover, save, lines } = S;
const OUT = process.argv[2], BBG = process.argv[3];

// cost example (2026 job rules, Ontario minimum wage $17.95 from Oct 1, 2026)
const MIN = 17.95, WEEKS = 8, HRS = 35, H = WEEKS * HRS, SUB = +(MIN * 0.5 * H).toFixed(0);
const money = (x) => "$" + Math.round(x).toLocaleString("en-CA");
const row = (wage) => { const wages = wage * H, vac = wages * 0.04; return [`$${wage.toFixed(2)} an hour`, money(wages), money(vac), money(SUB), money(wages + vac - SUB)]; };

const children = [
  cover("BANCROFT BOXING & GRAPPLING", "CANADA SUMMER JOBS 2027", "A draft application for a funded summer student. Cody reviews it, fills the yellow gaps and submits it himself.", [BBG]),
  sp(140),
  callout("The short version", [
    [Bd("Opens: "), "Tuesday, October 27, 2026 at 9:00 am ET. The closing date isn't posted yet (last year it was December 11). Applying early doesn't score higher, but don't leave it late."],
    [Bd("What BBG gets: "), `up to half of Ontario's minimum wage for each hour the student works. At $17.95 an hour that's about $8.98 an hour, or about ${money(SUB)} for ${WEEKS} weeks at ${HRS} hours.`],
    [Bd("Do first: "), "open a CRA payroll account. The very first question on the form asks for BBG's payroll (RP) number, and BBG only has a business number today."],
    [Bd("The job: "), "Summer Programs and Community Outreach Assistant. Gym only, no CCCL work (see the rules below)."],
  ]),

  ...H1("Before you apply", "About an hour, and only Cody can do these. The call opened October 27, so do them now."),
  chk([Bd("Open BBG's CRA payroll account (RP). "), "Fastest through CRA's Business Registration Online, or in My Business Account. It gives a 15-digit number like 123456789 RP0001. The form asks for it, and CRA needs it before the first payroll remittance anyway."]),
  chk([Bd("Set up GCOS "), "(Grants and Contributions Online Services): sign in with your bank or a GCKey, register BBG's business number, then confirm your identity through CRA My Account or at a Service Canada office. GCOS is the recommended way to apply. The online form and paper form EMP5396A also work; email and fax don't."]),
  chk([Bd("Decide the wage, the weeks and the start date "), "(section 4). Paying a little above minimum wage scores points."]),
  chk([Bd("Sort out WSIB or insurance. "), "Gyms are probably not required to have WSIB (call WSIB to confirm: 1-800-387-0750), but CSJ requires the student to be covered by WSIB or equivalent insurance from the first day. Optional WSIB coverage would cost roughly $45 for an 8-week job."]),
  chk([Bd("Have a short health-and-safety plan and an anti-harassment policy. "), "The form asks about both. Drafts are in BBG First Employee Basics (Command Centre) for you to adjust."]),

  ...H1("The rules that matter for BBG", "From the 2026 applicant guide; ESDC hasn't published anything different for 2027 yet."),
  bullet([Bd("The student can't replace volunteers. "), "BBG's coaches are volunteers, so the job adds new summer work (camps, front desk, outreach, social media). It doesn't take over a volunteer coach's class."]),
  bullet([Bd("BBG must manage the job itself. "), "CCCL is a separate business, so the draft leaves CCCL out. If you want the student at a CCCL show, ask CSJ first (1-800-935-5555)."]),
  bullet([Bd("No personal errands for you, "), "no fundraising to pay the student's wage, and nothing political."]),
  bullet([Bd("30 to 40 hours a week for 6 to 16 weeks in a row. "), "Part-time jobs don't qualify. ESDC usually approves about 8 weeks at 35 hours, and it may approve fewer weeks than you ask for."]),
  bullet([Bd("The student is 15 to 30 "), "when the job starts, a citizen, permanent resident or protected person, with a SIN."]),
  bullet([Bd("Don't start the job before written approval. "), "Hours before approval aren't paid back. Decisions came in April last year, by email, and there's no appeal."]),
  bullet([Bd("Hiring a family member "), "needs CSJ's written OK first."]),
  bullet([Bd("Answer CSJ within 5 business days "), "if they email with questions, or the application can be dropped."]),
  bullet([Bd("With anyone under 18: "), "the student gets a vulnerable sector check before starting and follows the Rule of Two like every BBG coach."]),

  ...H1("How it's scored", "Out of 100. The draft answers below aim at every line."),
  table([3600, 1300, 4820], [
    ["Part", "Points", "What earns them"],
    ["Wage above minimum", "10", "Pay more than $17.95 an hour"],
    ["Plans to keep the student", "5", "Say honestly whether there could be fall hours"],
    ["Supervision", "15", "Cody on site every shift, a clear check-in routine"],
    ["Mentoring", "13", "A real plan: orientation, weekly one-on-ones, goals, a reference"],
    ["Skills the student builds", "27", "Skills they can take to any job, named specifically"],
    ["Youth facing barriers", "10", "Actively recruiting youth the form lists"],
    ["National priorities", "10", "2026: construction; green jobs; technology and digital skills, including AI"],
    ["Local priorities", "10", "2026, our riding: not-for-profits, cultural events, community events, cultural community support, local and regional tourism"],
  ]),
  P(["Our riding is ", Bd("Hastings—Lennox and Addington—Tyendinaga"), " (Elections Canada lists the Town of Bancroft in it). The 2027 priorities aren't out yet; check the CSJ local priorities page after October 27 and adjust the two priority answers."], { spacing: { before: 100, after: 90 } }),

  ...H1("The numbers", "An example for 8 weeks at 35 hours (280 hours). Change it to suit."),
  table([2400, 1700, 1700, 1960, 1960], [
    ["Wage", "Wages", "Vacation pay (4%)", "CSJ pays back", "BBG pays"],
    row(17.95), row(18.50), row(19.00),
  ]),
  P("CSJ pays back half the minimum wage per hour worked, whatever wage BBG pays. BBG also pays the employer's share of CPP and EI, which CSJ doesn't cover (a few hundred dollars for a job this size; CRA's payroll calculator gives the exact figure). These are estimates, not tax advice.", { spacing: { before: 90, after: 90 } }),
  P(["Suggested: ", Bd("$18.50 an hour, 8 weeks at 35 hours, "), "starting ", "[Cody: start date, e.g. Monday June 28, 2027, after school ends]", ". CSJ's 2026 jobs ran between April 20 and August 29; check the 2027 dates."]),

  ...H1("Draft answers", "Copy each box into the form. Fill the yellow gaps; change anything that isn't how you'd say it."),
  H2("Part A: about BBG"),
  table([3200, 6520], [
    ["Form question", "Answer"],
    ["Payroll (RP) number", "[Cody: 15-digit number from CRA, e.g. 123456789 RP0001]"],
    ["Legal and operating name", "[Cody: your legal name as a sole proprietor], operating as Bancroft Boxing & Grappling"],
    ["Year established", "2025 [Cody: confirm]"],
    ["Main activities", "A community boxing and grappling gym in Bancroft. Classes in boxing, kickboxing, wrestling, no-gi jiu-jitsu and MMA for kids from about 4, youth and adults, plus a kids' program, women's boxing, summer fitness camps and birthday parties."],
    ["Full-time employees", "0"],
    ["Where the job is", "17 Snow Road, Bancroft, ON K0L 1C0"],
    ["Other wage funding for this job", "No [Cody: confirm]"],
    ["Debts to the Government of Canada", "[Cody]"],
    ["Health and safety, WSIB, anti-harassment", "Yes: a written health and safety policy and a written harassment program (BBG First Employee Basics). Workers' compensation: [Cody: optional WSIB coverage, or the insurer that covers the student]"],
  ], { label: true }),
  sp(80),
  H2("Part B: the job"),
  answer("JOB TITLE", "Summer Programs and Community Outreach Assistant"),
  sp(30),
  answer("WAGE AND HOURS", "[Cody: $18.50] an hour, 35 hours a week for 8 weeks, [Cody: start and end dates]."),
  sp(30),
  answer("MAIN TASKS", [
    "1. Help Cody plan and run BBG's summer fitness camp days and drop-in youth sessions: check-in, warm-up games, setting out equipment, and keeping groups organized, always alongside Cody.",
    "2. Welcome new members and families at the front desk: answer questions, book free first classes in the Gymdesk system, and help parents with the online sign-up and waiver.",
    "3. Plan, photograph and schedule the gym's social media posts using BBG's content calendar, following the gym's photo-consent rules for children.",
    "4. Organize community outreach: a free family try-it day, open-gym sessions for summer visitors and cottagers, and a BBG table at local summer events [Cody: name the events].",
    "5. Keep the gym safe and clean: set up and sanitize mats and equipment, check first-aid supplies and keep the equipment list up to date.",
    "6. Use the gym's digital tools: Gymdesk for scheduling and attendance, Google Sheets for tracking, and AI writing tools to draft posts and member messages that Cody reviews before anything is sent.",
    "7. Track simple numbers each week (camp sign-ups, first-class visitors, social media results) and present a short weekly summary to Cody.",
  ]),
  sp(30),
  answer("SKILLS THE STUDENT WILL BUILD", [
    "Customer service and communication with the public, parents and children; organizing and running events; digital marketing and social media; using scheduling software, spreadsheets and AI tools responsibly; basic data tracking and reporting; teamwork and leadership with groups of kids; workplace safety, first aid and concussion awareness (Ontario's Rowan's Law resources).",
    "By the end of the summer the student will have planned and run a community event, built a month of social media content, and kept a weekly report, all of which they can show future employers.",
  ]),
  sp(30),
  answer("SUPERVISION PLAN", [
    "Cody Purtell, the owner and head coach, supervises directly and is on site for every shift. Each shift starts with a 5-minute plan for the day and ends with a quick check of what got done. The student never supervises children alone: BBG follows the Rule of Two (two screened adults with anyone under 18), and the student gets a vulnerable sector check before starting.",
  ]),
  sp(30),
  answer("MENTORING PLAN", [
    "Week 1: orientation (gym tour, safety rules, emergency and concussion procedures, the Rule of Two, customer service standards) and first aid/CPR training [Cody: course and provider]. Together we set three goals for the summer.",
    "Every week: a 30-minute one-on-one with Cody to review the weekly summary, give feedback and plan the next week.",
    "Mid-summer and final reviews against the three goals. At the end, a written reference letter and an introduction to local business owners who support the gym.",
  ]),
  sp(30),
  answer("PLANS TO KEEP THE STUDENT", "[Cody: honest answer, e.g. \"If the summer programs grow, we'd like to offer part-time hours in the fall.\"]"),
  sp(30),
  answer("YOUTH FACING BARRIERS", "[Cody: pick the groups from the form's list.] We'll share the job with the local Employment Ontario office and North Hastings High School, and invite applications from youth facing barriers."),
  sp(30),
  answer("NATIONAL PRIORITIES (2026 list; update if 2027 changes)", "Technology and digital skills, including foundational AI skills: the student learns scheduling software, spreadsheets, social media tools and AI writing tools, with clear rules about checking AI drafts and protecting members' privacy."),
  sp(30),
  answer("LOCAL PRIORITIES (2026 list; update if 2027 changes)", "Community events and local and regional tourism development: the student runs a free family try-it day, open-gym sessions that welcome summer visitors and cottagers, and a BBG presence at local summer events in Bancroft."),
  sp(80),
  H2("Part C"),
  bullet("Costs: the wage and hours above. CSJ works out its share."),
  bullet("Advance payment: probably not needed; leave it unticked unless cash flow is tight."),
  bullet("Attestation: Cody reads and signs it himself."),

  ...H1("After you submit"),
  num("CSJ may email with questions: answer within 5 business days.", "n2"),
  num("Decisions arrive by email, starting about April. Nothing starts before written approval.", "n2"),
  num("Post the job (Job Bank is free), interview, and do the vulnerable sector check.", "n2"),
  num("Have WSIB or equivalent insurance in place before day one, and work through the checklist in BBG First Employee Basics (posters, training, first aid, payroll).", "n2"),
  num("Fill in the Employer and Employee Declaration, then claim. CSJ pays after the claim is processed, not up front.", "n2"),
  sp(60),
  H2("The job posting (once CSJ approves)"),
  S.answer("POST ON JOB BANK, FACEBOOK AND AT NORTH HASTINGS HIGH SCHOOL", [
    "Summer job: Programs and Community Outreach Assistant, Bancroft Boxing & Grappling",
    "Help run summer camp days and drop-in sessions, welcome new members at the front desk, make our social media posts and organize a free family try-it day. You'll learn scheduling software, digital marketing and event planning, and get first aid training.",
    "[Cody: 35] hours a week for 8 weeks, starting [Cody: date]. $[Cody: wage] an hour. Ages 15 to 30 (Canada Summer Jobs). No boxing experience needed: we'll teach you.",
    "To apply, send a short note about yourself to bancroftboxingandgrappling@outlook.com by [date]. Funded in part by the Government of Canada's Canada Summer Jobs program.",
  ]),
  P([R("Sources, checked Oct 20, 2026: the canada.ca CSJ applicant guide, CRA payroll pages, ontario.ca minimum wage, Elections Canada boundaries, WSIB FAQs. A draft, not legal or tax advice.", { size: 16, color: "5A6672", italics: true })]),
];
save(OUT, children, "CSJ 2027 draft · Bancroft Boxing & Grappling");
