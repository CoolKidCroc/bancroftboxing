// Cody's phone cheat sheet: 7 phone-size cards (1080x1920) as PNGs + one PDF.
// usage: NODE_PATH=$(npm root -g) node cheat.js OUTDIR CCCL_LOGO BBG_LOGO
const fs = require("fs"), path = require("path");
const { chromium } = require("playwright");
const OUT = process.argv[2]; fs.mkdirSync(OUT, { recursive: true });
const CCCL = "file://" + path.resolve(process.argv[3]), BBG = "file://" + path.resolve(process.argv[4]);
const FS = path.resolve(__dirname, "../canva/node_modules/@fontsource");
const f = (p) => "file://" + FS + "/" + p;
const CSS = `
@font-face{font-family:Saira;font-weight:700;src:url(${f("saira-condensed/files/saira-condensed-latin-700-normal.woff2")})}
@font-face{font-family:Saira;font-weight:800;src:url(${f("saira-condensed/files/saira-condensed-latin-800-normal.woff2")})}
@font-face{font-family:Saira;font-weight:600;src:url(${f("saira-condensed/files/saira-condensed-latin-600-normal.woff2")})}
@font-face{font-family:Plex;font-weight:400;src:url(${f("ibm-plex-sans/files/ibm-plex-sans-latin-400-normal.woff2")})}
@font-face{font-family:Plex;font-weight:500;src:url(${f("ibm-plex-sans/files/ibm-plex-sans-latin-500-normal.woff2")})}
@font-face{font-family:Plex;font-weight:600;src:url(${f("ibm-plex-sans/files/ibm-plex-sans-latin-600-normal.woff2")})}
:root{--navy:#081C2B;--navy2:#11283B;--navy3:#183650;--gold:#EAA959;--gold2:#C8862E;--on:#F4F1EA;--muted:#9AABBA;--line:#24405A}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--navy)}
.card{width:1080px;height:1920px;position:relative;overflow:hidden;color:var(--on);font-family:Plex,sans-serif;page-break-after:always;
  background:radial-gradient(90% 50% at 85% 0%,rgba(234,169,89,.13),rgba(234,169,89,0) 70%),linear-gradient(180deg,#0B2236,#081C2B 40%,#06141F)}
.top{padding:150px 76px 0}
.eyebrow{font-family:Saira;font-weight:700;letter-spacing:.16em;text-transform:uppercase;color:var(--gold);font-size:30px;display:flex;justify-content:space-between}
.title{font-family:Saira;font-weight:800;text-transform:uppercase;font-size:132px;line-height:.92;margin-top:18px}
.lede{font-size:34px;line-height:1.4;color:#C9D3DC;margin-top:22px}
.list{padding:34px 76px 0;display:flex;flex-direction:column;gap:18px}
.it{background:rgba(17,40,59,.85);border:2px solid var(--line);border-left:8px solid var(--gold2);border-radius:14px;padding:22px 28px}
.say{font-family:Plex;font-weight:600;font-size:41px;line-height:1.25}
.say::before{content:"“";color:var(--gold)} .say::after{content:"”";color:var(--gold)}
.get{font-size:30px;line-height:1.38;color:var(--muted);margin-top:8px}
.b{color:var(--gold)}
.sec{font-family:Saira;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:var(--gold);font-size:30px;margin:14px 0 -4px}
.foot{position:absolute;left:76px;right:76px;bottom:70px;display:flex;justify-content:space-between;align-items:center;font-family:Saira;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);font-size:26px}
.foot img{width:74px;height:74px;border-radius:50%;margin-left:12px}
.step{display:flex;gap:26px;align-items:flex-start;margin-top:26px}
.num{flex:0 0 74px;height:74px;border-radius:50%;background:var(--gold);color:var(--navy);font-family:Saira;font-weight:800;font-size:46px;display:flex;align-items:center;justify-content:center}
.rule{display:flex;gap:20px;align-items:flex-start;font-size:33px;line-height:1.35;margin-top:18px}
.rule i{flex:0 0 14px;height:14px;border-radius:50%;background:var(--gold);margin-top:16px}
`;
const br = (s) => s.replace(/\[([^\]]+)\]/g, '<span class="b">[$1]</span>');
const items = (arr) => arr.map(x => x.sec ? `<div class="sec">${x.sec}</div>` :
  `<div class="it"><div class="say">${br(x[0])}</div><div class="get">${x[1]}</div></div>`).join("");
const foot = (n) => `<div class="foot"><span>The Constrictor AI · ${n} of 9</span><span style="display:flex;align-items:center">Swipe for more<img src="${BBG}"><img src="${CCCL}"></span></div>`;
const card = (n, eyebrow, title, lede, list) => `<div class="card"><div class="top"><div class="eyebrow"><span>${eyebrow}</span><span>${n} / 9</span></div>
  <div class="title">${title}</div>${lede ? `<div class="lede">${lede}</div>` : ""}</div><div class="list">${items(list)}</div>${foot(n)}</div>`;

const C = [];
// 1. How it works
C.push(`<div class="card"><div class="top">
  <div class="eyebrow"><span>Cody's cheat sheet</span><span>1 / 9</span></div>
  <div style="display:flex;gap:22px;margin-top:40px"><img src="${BBG}" style="width:170px;border-radius:50%"><img src="${CCCL}" style="width:170px;border-radius:50%"></div>
  <div class="title" style="font-size:122px;margin-top:40px">Talk to Claude like you'd talk to Kyle</div>
  <div class="lede" style="font-size:36px">Say what you want in plain words. Claude checks your sheets, writes the draft and hands it back. <b style="color:var(--on)">You (or staff) hit send.</b></div>
  <div class="step"><div class="num">1</div><div style="font-size:35px;line-height:1.4">Open Claude on your phone or computer.</div></div>
  <div class="step"><div class="num">2</div><div style="font-size:35px;line-height:1.4">Say a line from these cards. Fill in the <span class="b">[gold brackets]</span>.</div></div>
  <div class="step"><div class="num">3</div><div style="font-size:35px;line-height:1.4">Read what comes back. Change it by just saying so: "shorter", "friendlier", "add the price". Then send it yourself.</div></div>
  <div style="margin-top:50px;background:rgba(17,40,59,.85);border:2px solid var(--gold2);border-radius:16px;padding:30px 34px">
    <div class="sec" style="margin:0 0 6px">The 3 golden rules</div>
    <div class="rule"><i></i><div><b>Drafts only.</b> Claude never sends, posts or pays anything.</div></div>
    <div class="rule"><i></i><div><b>You and the officials decide every bout.</b> Claude only flags problems.</div></div>
    <div class="rule"><i></i><div><b>Only "Paid" counts as money.</b> Promised isn't paid.</div></div></div>
</div>${foot(1)}</div>`);
// 2. Every day
C.push(card(2, "Every day", "Your day", "Constrictor HQ: tasks, approvals, staff and the inbox.", [
  ["Morning report", "Today's schedule, what needs your OK, overdue tasks, money, new gym inquiries, the next show, inbox and your top 3 moves."],
  ["Monday report", "The morning report plus last week's numbers and the week ahead."],
  ["What's on my plate?", "Your open tasks, most urgent first."],
  ["Brain-dump: [just talk for a minute or two]", "Turns a voice note or rant into tasks, drafts and reminders."],
  ["Go through my approvals", "One at a time: yes, change it, or no."],
  ["Ask [Sean] to [do this] by [Friday]", "Makes the task for HQ and drafts the message for you to send."],
  ["Sort my inbox", "Sorts the last 3 days of email, leaves reply drafts in Gmail and flags what needs you."],
]));
// 3. The gym
C.push(card(3, "Bancroft Boxing & Grappling", "The gym", "Messages, members, posts and classes.", [
  ["Reply to this: [paste the message]", "A friendly reply with the right class, price and the free first class."],
  ["Who should I check in with? [attach the Gymdesk attendance export]", "Members not seen in a month, with a short message for each."],
  ["Posts for next week", "Three gym posts with captions, ready to copy."],
  ["Plan [Wednesday's kids boxing]", "A timed class plan from the Coaching Library, safe for the ages."],
  ["Progress card for [first name]", "Their Skills Passport level and what's next, for the parent."],
  ["Write an announcement: [holiday hours]", "A post and a members' message."],
  ["Who owes us? [attach the Gymdesk payments export]", "Overdue payments with a friendly reminder for each, for your OK."],
]));
// 4. Plan a show
C.push(card(4, "CCCL · planning", "Plan a show", "Where, when and will it make money.", [
  ["Where should the next show be?", "Ranks towns and venues, with likely profit."],
  ["Compare [Town] and [Town]", "Side by side: venue, crowd, costs, likely profit."],
  ["Does [date] clash with anything?", "Checks local events, holidays and other fight shows."],
  ["Should [Town] be adults-only?", "All ages vs a 19+ area vs adults-only, in dollars."],
  ["Draft a venue quote request for [venue]", "An email asking for price, dates and what's included."],
  ["How are ticket sales?", "Sold vs goal, and which links and posts are selling."],
  ["What's left for the [Town] show?", "The checklist, with what's overdue first."],
  ["Who do I owe a call?", "Venues, gyms and media due a follow-up."],
]));
// 5. Fight night
C.push(card(5, "CCCL · fight night", "Fight night", "Fighters, the card and the night itself.", [
  ["Find fighters for [Town]", "Messages to gym coaches, and the sign-up form."],
  ["Who could fight who?", "Matchups that fit Ontario's rules, with any problems flagged. You decide."],
  ["Send the bout offers", "A draft offer to each coach."],
  ["Send the fighter packages", "Each fighter's package and promo link, through their coach."],
  ["Who's ready to fight?", "Paperwork and medical status: Yes, No or Pending. Never details."],
  ["Run of show and announcer script", "The night minute by minute, and every sponsor read."],
  ["Do the report", "After the show: results, money, the Commission report due in 14 days."],
]));
// 6. Sponsors & tickets
C.push(card(6, "CCCL · money", "Sponsors & tickets", "Sponsors, ads and graphics.", [
  ["Find sponsors in [Town]", "Local businesses that fit, with a reason for each."],
  ["Draft pitches for [businesses]", "A personal pitch for each, with the right package."],
  ["Do my follow-ups", "Everyone due a nudge today. Every Monday."],
  ["How are sponsors going?", "Paid vs goal, signed-but-unpaid, and who to call."],
  ["[Business] said yes", "Drafts the thank-you and lists what to update and what they're owed."],
  ["Is this ad worth it?", "Keep it if a ticket costs about $20 or less to sell."],
  ["Canva text for the [fight card]", "Which starter design to use and the exact words for it."],
]));
// 7. Money & funding
C.push(card(7, "BBG · CCCL · money", "Money & funding", "The monthly picture and grants worth chasing.", [
  ["How did we do in [September]?", "In, out and what's left for BBG and CCCL, from the Money Summary."],
  ["Anything new out there?", "Monthly scan: new funding, events, rule changes, sponsors."],
  ["What funding fits [kids programs]?", "Programs we can apply to now, with amounts and dates."],
  ["Are we eligible for [Canada Summer Jobs]?", "Goes through the rules one by one, with the source."],
  ["Draft the [Canada Summer Jobs] application", "Real facts only, gaps marked for you. You submit."],
  ["Funding deadlines", "Everything due in the next 60 days, with the next step."],
  ["A parent asked about help with fees", "A kind reply about Jumpstart, which BBG is signed up for."],
]));
// 8. Second Chance & your page
C.push(card(8, "Second Chance · Cody", "Second Chance & your page", "Your nonprofit, kept separate, and your own posts.", [
  ["Where are we with Second Chance?", "Setup steps done, what's next, and whether the board is ready."],
  ["Draft a letter to [partner]", "A short intro and one clear ask. You send it."],
  ["Plan a placement for [first name]", "Safe roles only, and never CCCL alone."],
  ["Who's due for a check-in?", "First names and a friendly message for each."],
  ["Posts for my page this week", "Monday tip, Wednesday behind the scenes, Friday Get back up."],
  ["Write my speaker intro", "For schools, banquets and events."],
]));
// 9. Rules & where things live
C.push(`<div class="card"><div class="top"><div class="eyebrow"><span>Good to know</span><span>9 / 9</span></div>
  <div class="title">Rules & where things live</div></div>
  <div style="padding:30px 76px 0">
    <div class="sec">Claude follows these. So do we.</div>
    ${["First names only in the sheets. No card numbers or personal phone numbers.",
       "Medical: Yes, No or Pending only. Details go to the physician and the Commission.",
       "Kids: never in ads or press releases. Photos only with a parent's written OK.",
       "No prize worth $300 or more. It would make a fighter a pro.",
       "No cannabis or vape sponsors. Alcohol and betting get flagged first.",
       "Real numbers only: 400 tickets and 15 sponsors are real."].map(r => `<div class="rule" style="font-size:31px;margin-top:10px"><i></i><div>${r}</div></div>`).join("")}
    <div class="sec" style="margin-top:36px">Google Drive › Bancroft Boxing</div>
    ${[["Command Centre", "HQ, Contacts, Money Summary, Funding"], ["BBG Members & Leads", "Leads & Trials"], ["BBG Coaching", "Class plans, drills, safety, Skills Passport"],
       ["CCCL", "Profit Predictor · Sponsorships · Fight Night · Marketing"], ["Brand", "Logos and the CCCL brand kit"], ["Second Chance", "Its own folder: you and the board only"], ["Handoff", "This cheat sheet and the setup guide"]].map(r =>
      `<div style="display:flex;gap:20px;font-size:30px;line-height:1.35;margin-top:14px"><b style="flex:0 0 330px;color:var(--gold)">${r[0]}</b><span>${r[1]}</span></div>`).join("")}
    <div style="margin-top:36px;background:rgba(17,40,59,.85);border:2px solid var(--line);border-radius:16px;padding:26px 30px;font-size:30px;line-height:1.45">
      <b style="color:var(--gold)">Stuck?</b> Claude can't see a sheet or email? In Settings › Connectors, turn on Google Drive and Gmail as <b>ontario.cccl@gmail.com</b>. Anything else: message Kyle.</div>
  </div>
  <div class="foot"><span>The Constrictor AI · 9 of 9</span><span style="display:flex;align-items:center"><img src="${BBG}"><img src="${CCCL}"></span></div></div>`);

(async () => {
  const html = `<!doctype html><html><head><meta charset="utf-8"><style>${CSS}@page{size:1080px 1920px;margin:0}</style></head><body>${C.join("")}</body></html>`;
  const hf = path.join(OUT, "_cheat.html"); fs.writeFileSync(hf, html);
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1080, height: 1920 } });
  await p.goto("file://" + path.resolve(hf), { waitUntil: "load" }); await p.evaluate(() => document.fonts.ready); await p.waitForTimeout(300);
  const over = await p.evaluate(() => [...document.querySelectorAll(".card")].map((c, i) => {
    const cr = c.getBoundingClientRect(); let low = 0;
    c.querySelectorAll(".it,.rule,.step,.lede,div").forEach(e => { if (e.closest(".foot")) return; const r = e.getBoundingClientRect(); if (r.bottom > low) low = r.bottom; });
    return (i + 1) + ": lowest " + Math.round(low - cr.top);
  }));
  console.log(over.join(" | "));
  const n = await p.evaluate(() => document.querySelectorAll(".card").length);
  for (let i = 0; i < n; i++) {
    const el = (await p.$$(".card"))[i];
    await el.screenshot({ path: path.join(OUT, `Cheat sheet ${i + 1} of ${n}.png`) });
  }
  await p.pdf({ path: path.join(OUT, "Cody's Cheat Sheet (phone).pdf"), width: "1080px", height: "1920px", printBackground: true, margin: { top: 0, right: 0, bottom: 0, left: 0 } });
  await b.close(); fs.unlinkSync(hf);
})();
