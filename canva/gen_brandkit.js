// CCCL Brand Kit one-pager (Letter, 2 pages) -> PDF + PNG preview
// usage: node gen_brandkit.js OUT.pdf LOGO_PNG [research/marketing_tools.json]
const fs = require("fs"), path = require("path");
const { chromium } = require("playwright");
const OUT = process.argv[2], LOGO = process.argv[3];
const MT = process.argv[4] && fs.existsSync(process.argv[4]) ? JSON.parse(fs.readFileSync(process.argv[4])) : null;
const font = (f) => "file://" + path.resolve(__dirname, "node_modules/@fontsource/" + f);
const logo = "file://" + path.resolve(LOGO);
let sizes = [["Instagram + Facebook feed post (use this)", "1080 × 1350", "4:5"],
             ["Instagram 3:4 post (keep text off the top and bottom edges)", "1080 × 1440", "3:4"],
             ["Stories and Reels (text inside the middle 1080 × 1610)", "1080 × 1920", "9:16"],
             ["Facebook event cover (title and date in the middle)", "1920 × 1005", "≈1.91:1"],
             ["Facebook link / landscape", "1200 × 628", "1.91:1"], ["Poster", "11 × 17 in (also letter)", ""]];
const canvaSteps = [
  "In Canva, click Create a design, then Import file, and pick a starter design PDF from the Canva starter designs folder. Every word comes in as its own text box you can edit.",
  "The background (dark gradient, pines, red glow) is one image. Click it and Lock it so it doesn't move while you work.",
  "Check the fonts: if Canva swapped one, select the text and choose Anton (headlines), Oswald (names and dates) or Inter (small text). All three are free in Canva.",
  "Replace the sample text. For photos: delete the grey photo box, then drag a Frame (Elements > Frames) to the same spot and drop the photo into it, so it crops neatly.",
  "Add the QR code (Apps > QR code) with that source's tracking link. On stories, add the Instagram link sticker when posting.",
  "Keep a copy of each finished design and change the town, date and names for the next show.",
  "Free Canva is enough. Canva Pro (about $142 CAD a year) adds a full Brand Kit with every CCCL colour, saved brand templates and one-click resizing. Canva's connector for Claude works on free Canva."];
const TH = (n) => "file://" + path.resolve(__dirname, "thumbs/" + n + ".jpg");
const DES = [["01", "Tickets on sale", "Week 14: tickets go on sale. Pin it."], ["02", "Countdown story", "21, 14 and 7 days out"], ["03", "Poster", "Week 10, before the card is set. Print with the QR code."],
  ["04", "Fight card", "When the card is announced"], ["05", "Fighter vs fighter", "One per bout, weeks 5 to 1"], ["06", "Fighter spotlight", "Weeks 6 to 2, adults who agree"],
  ["07", "Sponsor thank-you", "Week 8 and after the show"], ["08", "Results", "Day after, official results only"], ["09", "Facebook event cover", "Week 16, the Facebook event"],
  ["10", "Poster with main event", "Once the main event is signed"], ["11", "I'm fighting (post)", "Each fighter, with their own ticket link"], ["12", "I'm fighting (story)", "Each fighter; add the link sticker"],
  ["13", "Weigh-in results", "Fight day, after weigh-ins"], ["14", "Win", "Day after. For a loss, change WIN to THANK YOU"], ["15", "Proud sponsor", "Sponsors share it, with their link"],
  ["16", "Fight card (story)", "Card announced and fight week"]];
const html = `<!doctype html><html><head><meta charset="utf-8"><style>
@font-face{font-family:Anton;src:url(${font("anton/files/anton-latin-400-normal.woff2")})}
@font-face{font-family:Oswald;font-weight:600;src:url(${font("oswald/files/oswald-latin-600-normal.woff2")})}
@font-face{font-family:Oswald;font-weight:500;src:url(${font("oswald/files/oswald-latin-500-normal.woff2")})}
@font-face{font-family:Inter;font-weight:400;src:url(${font("inter/files/inter-latin-400-normal.woff2")})}
@font-face{font-family:Inter;font-weight:600;src:url(${font("inter/files/inter-latin-600-normal.woff2")})}
@page{size:Letter;margin:0}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Inter;color:#151C19;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.pg{width:8.5in;height:11in;position:relative;overflow:hidden;page-break-after:always;background:#F4F1E8}
.top{background:#151C19;color:#F2ECDC;padding:.45in .6in .4in;display:flex;justify-content:space-between;align-items:center}
.h{font-family:Anton;text-transform:uppercase;line-height:.95}
.o{font-family:Oswald;text-transform:uppercase;letter-spacing:.06em}
.body{padding:.35in .6in}
h2{font-family:Oswald;font-weight:600;text-transform:uppercase;letter-spacing:.06em;font-size:13pt;color:#354940;margin:.22in 0 .1in;border-bottom:2px solid #AE2221;padding-bottom:3px}
p,li{font-size:10pt;line-height:1.45}
ul{padding-left:.2in}
.sw{display:grid;grid-template-columns:repeat(3,1fr);gap:.12in}
.sw div{border-radius:6px;padding:.12in;height:.95in;display:flex;flex-direction:column;justify-content:flex-end;font-size:9pt}
.sw b{font-family:Oswald;font-weight:600;font-size:11pt;text-transform:uppercase}
table{width:100%;border-collapse:collapse;font-size:9.5pt}
td,th{border-bottom:1px solid #D6CFBD;padding:5px 4px;text-align:left}
th{font-family:Oswald;font-weight:600;text-transform:uppercase;font-size:9pt;color:#354940}
.two{display:grid;grid-template-columns:1fr 1fr;gap:.3in}
</style></head><body>
<div class="pg">
  <div class="top"><div><div class="o" style="color:#D9483B;font-weight:600;font-size:12pt">Cottage Country Combat League</div>
    <div class="h" style="font-size:52pt;margin-top:6px">Brand kit</div>
    <div style="font-size:10pt;color:#C9C3B3;margin-top:6px">Colours, fonts, logo and sizes for every CCCL poster and post. For Canva and anyone making CCCL graphics.</div></div>
    <img src="${logo}" style="width:1.55in;border-radius:50%"></div>
  <div class="body">
    <h2>Logo</h2>
    <div class="two"><div>
      <p>Files: <b>Brand/CCCL logo.png</b> (with white corners) and <b>Brand/CCCL logo (transparent).png</b> (round, for dark backgrounds and Canva).</p>
      <ul><li>Keep it round and whole: no cropping, stretching or recolouring.</li><li>Put it on near-black or forest green, or on cream with space around it.</li>
      <li>Leave clear space around it of at least a quarter of its width.</li><li>Sponsor logos never sit on top of it.</li></ul></div>
      <div style="display:flex;gap:.2in;align-items:center;justify-content:center">
        <div style="background:#151C19;padding:.15in;border-radius:8px"><img src="${logo}" style="width:1.2in;border-radius:50%"></div>
        <div style="background:#F2ECDC;padding:.15in;border-radius:8px;border:1px solid #D6CFBD"><img src="${logo}" style="width:1.2in;border-radius:50%"></div></div></div>
    <h2>Colours (from the logo)</h2>
    <div class="sw">
      <div style="background:#151C19;color:#F2ECDC"><b>Near-black</b>#151C19 · backgrounds</div>
      <div style="background:#354940;color:#F2ECDC"><b>Forest green</b>#354940 · panels</div>
      <div style="background:#26362F;color:#F2ECDC"><b>Dark green</b>#26362F · cards, trees</div>
      <div style="background:#F2ECDC;color:#151C19;border:1px solid #D6CFBD"><b>Cream</b>#F2ECDC · text on dark</div>
      <div style="background:#D9483B;color:#fff"><b>Bright red</b>#D9483B · big words on dark</div>
      <div style="background:#AE2221;color:#fff"><b>Maple red</b>#AE2221 · small accents</div></div>
    <p style="margin-top:8px">Red is for the one big thing to notice (the town, "VS", "fighting") and small labels. Everything else is cream on dark.</p>
    <h2>Fonts (all free in Canva)</h2>
    <table><tr><th>Font</th><th>Use</th><th>Looks like</th></tr>
      <tr><td><b>Anton</b></td><td>Headlines: town, FIGHT NIGHT, names on the card</td><td class="h" style="font-size:20pt">Haliburton</td></tr>
      <tr><td><b>Oswald</b> (SemiBold)</td><td>Dates, venues, labels, fighter details</td><td class="o" style="font-weight:600;font-size:13pt">Sat May 1 · Doors 6 PM</td></tr>
      <tr><td><b>Inter</b></td><td>Small print: prices, sponsors, quotes</td><td style="font-size:10.5pt">General admission $60 · VIP $100</td></tr></table>
  </div>
</div>
<div class="pg"><div class="body" style="padding-top:.5in">
  <h2 style="margin-top:0">Sizes</h2>
  <table><tr><th>Format</th><th>Size</th><th>Ratio</th></tr>${sizes.map(s => `<tr><td>${s[0]}</td><td>${s[1]}</td><td>${s[2]}</td></tr>`).join("")}</table>
  <h2>The look</h2>
  <ul><li>Dark background, tall all-caps headline, the town in red, a thin red line, then date and venue in Oswald.</li>
    <li>Layered mountains and pines along the bottom with a soft red glow, like the logo at night.</li><li>The round logo top right (or centred on stories). A thin footer band for the date or "Tickets: link in bio".</li>
    <li>Real photos of real fighters beat any graphic: bright, close, sharp.</li></ul>
  <h2>The starter designs</h2>
  <p>16 designs in <b>CCCL/Marketing/Canva starter designs</b>, each as a PNG (to look at) and a PDF (to import into Canva). Page 3 shows them all and when to use each.</p>
  <h2>Using them in Canva</h2>
  <ol style="padding-left:.22in">${canvaSteps.map(s => `<li style="margin-bottom:4px">${s}</li>`).join("")}</ol>
  <h2>Rules</h2>
  <ul><li>Real numbers only: 400 tickets and 15 sponsors at the first show are real. No invented crowd or viewer numbers.</li>
    <li>Fighters: adults who agreed, official results only, respect for every fighter.</li>
    <li>Kids: never in ads, never named in press, photos only with a parent's written OK.</li>
    <li>Sponsors: logos sized by their level, only what their package includes. No alcohol logos on anything youth-facing.</li>
    <li>Every poster and post links to its own Eventbrite tracking link (Marketing Kit, Ticket Sources tab).</li></ul>
</div></div>
<div class="pg"><div class="body" style="padding-top:.45in">
  <h2 style="margin-top:0">The 16 starter designs</h2>
  <p style="margin-bottom:.12in">Sample text (Haliburton, May 1, fighter names) is there to replace. Fighters get 11, 12 and 14 with their own tracking link; sponsors get 15.</p>
  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:.14in .16in">${DES.map(d => `<div>
    <div style="background:#151C19;border-radius:6px;height:1.72in;display:flex;align-items:center;justify-content:center;padding:4px"><img src="${TH(d[0])}" style="max-width:100%;max-height:100%;border-radius:3px"></div>
    <div class="o" style="font-weight:600;font-size:9pt;margin-top:5px;color:#151C19">${d[0]} · ${d[1]}</div>
    <div style="font-size:7.8pt;line-height:1.3;color:#4a544f">${d[2]}</div></div>`).join("")}</div>
</div></div>
</body></html>`;
(async () => {
  const hf = path.join(path.dirname(OUT), "_brandkit.html"); fs.writeFileSync(hf, html);
  const b = await chromium.launch(); const p = await b.newPage();
  await p.goto("file://" + path.resolve(hf), { waitUntil: "load" }); await p.evaluate(() => document.fonts.ready); await p.waitForTimeout(200);
  await p.pdf({ path: OUT, format: "Letter", printBackground: true, margin: { top: 0, right: 0, bottom: 0, left: 0 } });
  await b.close(); fs.unlinkSync(hf); console.log("saved", OUT);
})();
