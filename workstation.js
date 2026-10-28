// The Constrictor Workstation Plan (Word): setting up Cody's HP ProBook 640 G5 as his command centre.
// What runs where, what to install, what each skill can touch, the cloud morning report, the dashboard, security, backups and costs.
// usage: NODE_PATH=$(npm root -g) node workstation.js OUT.docx BBG_LOGO CCCL_LOGO
const S = require("./docstyle.js")();
const { R, P, Bd, H1, H2, bullet, num, chk, table, callout, sp, answer, cover, save, lines } = S;
const OUT = process.argv[2], BBG = process.argv[3], CCCL = process.argv[4];

const PROMPT = "Run my Constrictor morning report with the constrictor-hq skill. Use Google Drive, Gmail and Google Calendar for ontario.cccl@gmail.com. "
  + "On Mondays, include the weekly numbers. Read only: don't send, share, move or delete anything, and don't change any sheet. "
  + "When the report is ready, save it as a Gmail draft to ontario.cccl@gmail.com with the subject \"Morning report\" and today's date.";

const children = [
  cover("THE CONSTRICTOR AI", "WORKSTATION PLAN", "Setting up the HP ProBook 640 G5 as Cody's command centre: what to install, what runs where, what each part of the AI can touch, what it costs, and how it keeps running when the laptop is off.", [BBG, CCCL]),
  sp(140),
  callout("The short version", [
    [Bd("The laptop is a window, not the engine. "), "Everything the businesses depend on already runs in the cloud: Claude (the 9 skills and the morning report), Google (the sheets, Gmail and Calendar), Gymdesk and Eventbrite. Close the laptop and nothing stops."],
    [Bd("Nothing to code or host. "), "No servers, GitHub repositories, API keys or automation tools. They'd add monthly costs and more things to break, and they wouldn't add anything the system needs."],
    [Bd("One new monthly cost: "), "Claude Pro, US$20 a month (about CA$28), or US$17 a month if paid yearly. Everything else is free or something BBG already pays for."],
    [Bd("Your dashboard "), "is the Dashboard tab of Constrictor HQ, plus a morning report waiting in Gmail at 7 every weekday."],
  ]),

  ...H1("What runs where", "If it's in the cloud, it works with the laptop shut. The laptop and phone are just how Cody gets to it."),
  table([2300, 5320, 2100], [
    ["Where", "What", "Laptop off?"],
    [lines([Bd("Claude (cloud)")], "Cody's Pro account"), "The 9 skills, every chat, and the 7am morning report (a Cloud scheduled task)", "Keeps running"],
    [lines([Bd("Google (cloud)")], "ontario.cccl@gmail.com"), "Every sheet and doc in Drive, Gmail and its drafts, Google Calendar", "Keeps running"],
    [lines([Bd("Google (cloud)")], "Second Chance Gmail"), "The Second Chance folder only, kept apart from BBG and CCCL", "Keeps running"],
    [lines([Bd("Gymdesk, Eventbrite")]), "Members, payments, attendance, waivers and signups; tickets", "Keeps running"],
    [lines([Bd("Facebook, Instagram, Canva")]), "The pages and the designs. People post; Claude never does", "Keeps running"],
    [lines([Bd("The ProBook (local)")]), "Chrome (where Cody works), the Claude desktop app (optional), the monthly backup to a USB drive", "Not needed"],
    [lines([Bd("Cody's phone")]), "Claude app, Gmail app, Google Authenticator: the morning report and approvals from anywhere", "Not needed"],
  ], { top: true }),
  sp(60),
  P([Bd("Installation files: "), "the 9 skill zip files and the install guide are in the folder The Constrictor (Claude skills). They're uploaded to Claude once (Handoff Guide, section 3). Kyle keeps the build files that make the sheets and skills, so updates arrive as new zips. Cody never needs GitHub."]),

  ...H1("Set up the laptop", "About an hour, in this order. Everything here is free or included with Windows 11 Pro.", true),
  table([600, 6720, 2400], [
    ["#", "What to do", "Why"],
    ["1", lines([Bd("Windows. "), "Run Windows Update until nothing's left. Sign in with a PIN or fingerprint (Windows Hello). Turn on BitLocker (search \"BitLocker\" or \"Device encryption\" in Settings) and save the recovery key to Cody's Microsoft account, or print it and keep it with the business papers. Set the screen to lock after 5 minutes."]), "A lost or stolen laptop gives away nothing"],
    ["2", lines([Bd("Google Chrome, one profile per account: "), "CCCL (ontario.cccl@gmail.com), Second Chance (its own Gmail, once it exists) and Personal. Each profile keeps its own logins and saved passwords."]), "Keeps the businesses, Second Chance and personal email apart"],
    ["3", lines([Bd("Bookmarks in the CCCL profile: "), "claude.ai, Constrictor HQ (pin it as the first tab), Google Drive, Gymdesk, Eventbrite, Canva and the Facebook pages."]), "One click to everything"],
    ["4", lines([Bd("Claude desktop app "), "(claude.ai/download), signed in to Cody's Claude account. Optional: claude.ai in Chrome does the same jobs."]), "Claude in its own window"],
    ["5", lines([Bd("On the phone: "), "the Claude app, the Gmail app (CCCL account) and Google Authenticator."]), "Morning report and approvals on the go"],
  ], { top: true }),
  sp(60),
  P([Bd("Don't install: "), "Python, Node, Docker, GitHub Desktop, n8n, Zapier, a home server, a second antivirus (Microsoft Defender is built in), or Microsoft Office (Google Sheets and Docs open every .xlsx and .docx in the folder)."]),
  P([Bd("The ProBook itself: "), "16 GB of memory is plenty, because the heavy work happens on Claude's and Google's servers and the laptop mostly runs Chrome. Check the battery holds a charge, and keep it plugged in at the desk."]),

  ...H1("The AI team and what each part can touch", "Each \"agent\" is a Claude skill. It only works when Cody asks (or when the morning report runs), and it drafts; a person sends."),
  table([1900, 2950, 2650, 2220], [
    ["Skill", "Reads", "Can make", "Never"],
    [lines([Bd("constrictor-hq")], "Chief of staff"), "Constrictor HQ, Sponsorship Pipeline, Leads & Trials, Contacts, Money Summary, Predictor, the show kits; Gmail inbox; Calendar", "The morning report, tasks, staff messages, Gmail drafts", "Send, delete, or change a sheet unless Cody says so"],
    [lines([Bd("bbg-front-desk")], "Gym replies"), "The gym's facts, prices and schedule (inside the skill)", "Reply drafts and a lead note", "Send; health details or kids' last names"],
    [lines([Bd("bbg-cccl-social")], "Posts, Cody's page"), "HQ Content Calendar, Marketing Kit, Fight Night Kit", "Post drafts, Canva text", "Post; anyone under 18 in an ad"],
    [lines([Bd("cccl-sponsor-desk")], "Sponsors"), "Sponsorship Pipeline; the web", "Gmail drafts; new sponsor rows to paste", "Send; count anything but Paid"],
    [lines([Bd("cccl-event-scout")], "Venues and profit"), "Profit Predictor, Contacts; the web", "Venue quote requests, comparisons", "Book or sign anything"],
    [lines([Bd("cccl-fight-night")], "Shows"), "Fight Night Kit, Marketing Kit, Contacts", "Gmail drafts, press releases, fighter packets", "Decide bouts (Cody and the officials do)"],
    [lines([Bd("bbg-coach")], "Classes"), "Coaching Library, Skills Passport", "Lesson plans, sub-coach sheets", "Medical advice"],
    [lines([Bd("constrictor-funding")], "Grants"), "Funding Tracker, CSJ draft; the web", "Application drafts", "Submit anything"],
    [lines([Bd("second-chance")], "The nonprofit"), "The Second Chance Tracker only", "Letters, board notes, plans", "Copy anything into a BBG or CCCL file; record offences, health or addresses"],
  ], { top: true }),
  H2("Claude settings (the permissions)"),
  table([3100, 6620], [
    ["Setting", "How to set it"],
    ["Google Drive, Gmail, Google Calendar", "Connected with ontario.cccl@gmail.com. Gmail sending stays on \"ask every time\" (the default), so nothing goes out without Cody clicking."],
    ["Code execution", "On (Settings > Capabilities). The skills need it."],
    ["Web search", "On. Sponsors, venues and grants are researched on the web."],
    ["Skills", "All 9 uploaded and switched on (Customize > Skills)."],
    ["Everything else", "Off for now: Claude in Chrome, computer use, other connectors. None of the jobs need them."],
  ], { label: true }),

  ...H1("Keeping it running with the laptop off", "Claude has two kinds of scheduled tasks. Only one keeps going when the laptop is shut.", true),
  table([2200, 7520], [
    ["Kind", "What happens"],
    ["Cloud task", "Runs on Anthropic's servers on schedule, whether the laptop is on, asleep or shut. This is the one to use."],
    ["Local (desktop) task", "Only runs while the laptop is awake with the Claude app open. A missed run waits until the laptop wakes up. Don't use it for the morning report."],
  ], { label: true }),
  H2("Set up the morning report (at the sit-down, 10 minutes)"),
  num("In Claude, open Scheduled (in some accounts it's called Routines) and click New task. If it asks Local or Cloud, pick Cloud.", "n2"),
  num("Name: Constrictor morning report.", "n2"),
  num("Paste the instructions in the box below.", "n2"),
  num("Schedule: weekdays at 7:05 am. A few minutes past the hour starts closer to on time than 7:00.", "n2"),
  num("Connectors: keep Google Drive, Gmail and Google Calendar. Remove any others.", "n2"),
  num("Click Create, then Run now once to test it. The report should land in Gmail drafts.", "n2"),
  num("The real test: shut the laptop that night. At 7:10 the next morning, the report should be in Gmail on Cody's phone.", "n2"),
  sp(40),
  answer("THE INSTRUCTIONS TO PASTE", PROMPT),
  sp(60),
  P([R("Scheduled tasks are still a preview feature, so their screens and limits can change. If a run ever fails, the task's History shows why, and saying \"Morning report\" in any chat still works.", { italics: true, color: "5A6672", size: 18 })]),

  ...H1("Your dashboard", "One place to see every business, from any computer or phone."),
  bullet([Bd("Constrictor HQ > Dashboard: "), "approvals waiting, overdue tasks, this week's posts, last week's numbers and sponsorship paid against the goal. The Links tab opens every other sheet."]),
  bullet([Bd("The morning report: "), "today's calendar, what needs Cody's OK, money, new gym inquiries, the next show, the inbox and the top 3 moves, in one phone screen."]),
  bullet([Bd("Watching the AI: "), "the morning report's History page shows every run and whether it worked. The skills don't run on their own, so there's nothing else to watch. Once a month, ask \"Is everything being kept up?\" and HQ checks every sheet is being used."]),
  bullet([Bd("Second Chance stays off the dashboard "), "on purpose, to protect participants' privacy."]),
  P([Bd("Why not a separate dashboard app: "), "it would copy the same numbers into a second place that goes out of date, and most of them charge monthly."]),

  ...H1("Secure access", "Ten minutes, once."),
  chk([Bd("2-Step Verification "), "on ontario.cccl@gmail.com, the Second Chance Gmail and the email Cody signs in to Claude with. Use Google Authenticator or phone prompts, and keep the backup codes on paper."]),
  chk([Bd("Change the CCCL password "), "after Kyle's setup. Nobody else signs in to Cody's Claude or Google accounts; staff get single sheets shared with them (Handoff Guide, section 4)."]),
  chk([Bd("Gmail sending stays on \"ask every time\" "), "in Claude."]),
  chk([Bd("Laptop: "), "PIN or fingerprint, BitLocker on, screen locks after 5 minutes, Windows Update on."]),
  chk([Bd("If the laptop is lost: "), "from the phone, change the Google and Claude passwords and sign out the laptop (Google Account > Security > Your devices). Nothing is lost, because it all lives in the cloud."]),
  chk([Bd("Cody's personal email "), "stays in the Personal Chrome profile and never goes in a business file."]),

  ...H1("Backups", "Free, about 15 minutes on the first of each month."),
  bullet([Bd("Automatic: "), "every Google Sheet and Doc keeps its own version history (File > Version history), so any day's version can be restored. Gymdesk and Eventbrite keep their own records."]),
  bullet([Bd("Monthly: "), "in Google Drive, right-click the Bancroft Boxing folder > Download. Drive zips it, with every sheet as Excel and every doc as Word. Copy the zip to a USB drive named with the month, and keep the last 12. Add the month's Gymdesk and Eventbrite exports (the same ones the Money Summary uses)."]),
  bullet([Bd("Second Chance: "), "its own monthly download, kept separately. The board decides where."]),
  bullet([Bd("The skills: "), "the 9 zip files are in the folder, so they're in every backup."]),
  P([Bd("Why the Drive app on the laptop isn't a backup: "), "it shows Google Sheets as links to the online file, not as copies."]),

  ...H1("What it costs", "Only Claude Pro is new. Prices checked October 2026."),
  table([3300, 3720, 2700], [
    ["Item", "Cost", "Status"],
    ["Claude Pro", "US$20 a month (about CA$28), or US$17 a month paid yearly", "Needed. Cody chose Pro"],
    ["Google accounts (CCCL, Second Chance)", "Free, 15 GB each", "Have"],
    ["Gymdesk", "What BBG pays now", "Have, no change"],
    ["Eventbrite", "Fees per ticket", "Have, no change"],
    ["Canva", "The free plan covers the starter designs", "Have"],
    ["Chrome, Claude app, Authenticator, Defender, BitLocker", "Free", "Install"],
    ["USB drive for backups", "One-time, small", "Buy once"],
    ["API or usage fees", "None. The skills run inside the Claude plan, with no API keys", "Nothing to buy"],
  ], { label: true }),
  H2("Don't buy (each one would duplicate something you have)"),
  table([3300, 6420], [
    ["Skip", "Because"],
    ["Claude Team", "Only needed to give staff their own copy of the skills. Staff use the Staff Playbook prompts instead."],
    ["Claude Max", "Only if Cody keeps hitting Pro's usage limits. Upgrade then, not before."],
    ["ChatGPT Plus or other AI plans", "Claude already does these jobs."],
    ["Zapier, Make, n8n, a server", "The Cloud morning report covers what needs to run on its own."],
    ["Microsoft 365", "Google already opens and edits every file."],
    ["A dashboard app, paid antivirus", "HQ is the dashboard; Defender is built in."],
    ["More Google storage", "Only if a 15 GB account fills up."],
  ], { label: true }),

  ...H1("The rollout, in order"),
  table([2300, 1600, 5820], [
    ["When", "Who", "What"],
    ["The day the laptop arrives", "Cody, Kyle on the phone", "Section 2 (set up the laptop) and 2-Step Verification (section 6). About an hour."],
    ["Before the sit-down", "Kyle", "Handoff Guide, section 2: upload to Drive, run the skill tests."],
    ["The sit-down (90 min)", "Cody and Kyle", "Handoff Guide, section 3, on this laptop: Pro plan, skills, connectors, then the Cloud morning report (section 4 here)."],
    ["The next morning", "Cody", "Laptop shut overnight. Check the report is in Gmail at 7:10."],
    ["Once Kyle's done", "Cody", "Change the CCCL password. Set up the Second Chance Gmail and folder in its own Chrome profile."],
    ["Week 2", "Cody", "Share Leads & Trials with the front desk and the coaching sheets with the coaches (Handoff Guide, section 4)."],
    ["First of every month", "Cody", "The backup (section 7) and \"Is everything being kept up?\""],
    ["Every year", "Kyle", "The yearly check in the Handoff Guide, section 6."],
  ], { label: true, top: true }),
  sp(80),
  P([R("Sources: Claude Pro pricing (chatgpt.ca, July 2026); Claude Help Center: Schedule recurring tasks in Cowork, and Use Google Workspace connectors; Claude Code docs: Automate work with scheduled tasks (cloud tasks run when the laptop is closed). Checked October 27, 2026.", { size: 16, color: "5A6672", italics: true })]),
];
save(OUT, children, "The Constrictor workstation plan");
