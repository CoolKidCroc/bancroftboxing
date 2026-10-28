// Three starter designs for Cody's own page (navy and gold "Constrictor" look): Constrictor Tip post, Constrictor Tip story, Get Back Up quote.
// PNG for posting, PDF to import into Canva and edit. usage: NODE_PATH=$(npm root -g) node athlete_designs.js OUTDIR
const fs = require("fs"), path = require("path");
const { chromium } = require("playwright");
const OUT = process.argv[2]; fs.mkdirSync(OUT, { recursive: true });
const FS = path.resolve(__dirname, "canva/node_modules/@fontsource");
const f = (p) => "file://" + FS + "/" + p;
const CSS = `
@font-face{font-family:Saira;font-weight:800;src:url(${f("saira-condensed/files/saira-condensed-latin-800-normal.woff2")})}
@font-face{font-family:Saira;font-weight:700;src:url(${f("saira-condensed/files/saira-condensed-latin-700-normal.woff2")})}
@font-face{font-family:Plex;font-weight:400;src:url(${f("ibm-plex-sans/files/ibm-plex-sans-latin-400-normal.woff2")})}
@font-face{font-family:Plex;font-weight:600;src:url(${f("ibm-plex-sans/files/ibm-plex-sans-latin-600-normal.woff2")})}
@font-face{font-family:Mono;font-weight:500;src:url(${f("ibm-plex-mono/files/ibm-plex-mono-latin-500-normal.woff2")})}
:root{--navy:#081C2B;--deep:#04111B;--gold:#EAA959;--gold2:#C8862E;--cream:#F4F1EA;--muted:#9FB0BF}
*{box-sizing:border-box;margin:0;padding:0}
body{background:#000}
.art{position:relative;overflow:hidden;color:var(--cream);font-family:Plex,sans-serif;
  background:radial-gradient(120% 80% at 85% 10%,#123049 0%,var(--navy) 45%,var(--deep) 100%)}
.coil{position:absolute;border-radius:50%;border:3px solid rgba(234,169,89,.13)}
.eyebrow{font-family:Mono;font-weight:500;letter-spacing:.18em;text-transform:uppercase;color:var(--gold);font-size:30px}
.title{font-family:Saira;font-weight:800;text-transform:uppercase;line-height:.9;letter-spacing:.005em}
.photo{position:absolute;left:0;right:0;top:0;background:linear-gradient(180deg,#1b3a52 0%,#0d2436 100%);display:grid;place-items:center;
  color:rgba(244,241,234,.45);font-family:Mono;letter-spacing:.2em;font-size:26px;text-transform:uppercase}
.photo::after{content:"";position:absolute;inset:auto 0 0 0;height:45%;background:linear-gradient(180deg,transparent,var(--navy))}
.bar{position:absolute;left:0;right:0;bottom:0;display:flex;justify-content:space-between;align-items:center;padding:30px 64px;
  border-top:3px solid var(--gold);background:rgba(4,17,27,.92);font-family:Saira;font-weight:700;letter-spacing:.1em;text-transform:uppercase;font-size:28px}
.bar b{color:var(--gold)}
.chip{display:inline-block;border:2px solid var(--gold);color:var(--gold);font-family:Saira;font-weight:700;letter-spacing:.14em;text-transform:uppercase;padding:10px 20px;font-size:28px}
`;
const coils = (cx, cy, n, r0, step) => Array.from({ length: n }, (_, i) => { const r = r0 + i * step; return `<div class="coil" style="width:${2*r}px;height:${2*r}px;left:${cx-r}px;top:${cy-r}px"></div>`; }).join("");

const D = [
  ["01 Constrictor Tip (post 1080x1350)", 1080, 1350, `
    <div class="photo" style="height:760px">Your photo or video still (full width)</div>
    ${coils(980, 1180, 6, 90, 70)}
    <div style="position:absolute;left:64px;right:64px;top:640px">
      <div class="eyebrow">Constrictor Tip #01</div>
      <div class="title" style="font-size:150px;margin-top:18px">The rear naked<br><span style="color:var(--gold)">choke grip</span></div>
      <div style="font-size:36px;line-height:1.35;color:#C9D3DC;margin-top:26px;max-width:860px">The detail that won my first pro fight. Watch the clip, then come learn it properly.</div>
    </div>
    <div class="bar"><span>Cody <b>"The Constrictor"</b> Purtell</span><span>First class free · <b>link in bio</b></span></div>`],
  ["02 Constrictor Tip (story 1080x1920)", 1080, 1920, `
    <div class="photo" style="height:1080px">Your video (full width)</div>
    ${coils(1000, 1720, 7, 90, 70)}
    <div style="position:absolute;left:64px;right:64px;top:960px">
      <div class="eyebrow">Constrictor Tip #01</div>
      <div class="title" style="font-size:170px;margin-top:22px">The rear naked<br><span style="color:var(--gold)">choke grip</span></div>
      <div style="font-size:40px;line-height:1.35;color:#C9D3DC;margin-top:30px">Bancroft Boxing &amp; Grappling · 17 Snow Road</div>
      <div style="margin-top:44px"><span class="chip">Put your link sticker here</span></div>
    </div>
    <div class="bar" style="padding:36px 64px 80px"><span>Cody <b>"The Constrictor"</b> Purtell</span><span><b>First class free</b></span></div>`],
  ["03 Get Back Up quote (post 1080x1350)", 1080, 1350, `
    ${coils(160, 200, 7, 60, 64)}
    ${coils(1010, 1150, 6, 80, 70)}
    <div style="position:absolute;left:84px;right:84px;top:330px">
      <div class="title" style="font-size:240px;color:var(--gold);line-height:.6">&ldquo;</div>
      <div style="font-family:Saira;font-weight:700;text-transform:uppercase;font-size:84px;line-height:1.02;margin-top:10px">If you have a dream, fight for it! No matter how many times you get knocked down,
        <span style="color:var(--gold)">get back up</span> and keep going.</div>
      <div style="font-size:34px;color:#C9D3DC;margin-top:44px">Cody "The Constrictor" Purtell</div>
    </div>
    <div class="bar"><span><b>Get back up</b></span><span>Bancroft Boxing &amp; Grappling</span></div>`],
];

(async () => {
  const b = await chromium.launch();
  for (const [name, w, h, body] of D) {
    const p = await b.newPage({ viewport: { width: w, height: h } });
    const html = `<!doctype html><html><head><meta charset="utf-8"><style>${CSS}@page{size:${w}px ${h}px;margin:0}</style></head><body><div class="art" style="width:${w}px;height:${h}px">${body}</div></body></html>`;
    const hf = path.join(OUT, "_a.html"); fs.writeFileSync(hf, html);
    await p.goto("file://" + path.resolve(hf), { waitUntil: "load" }); await p.evaluate(() => document.fonts.ready); await p.waitForTimeout(250);
    const over = await p.evaluate(() => { const a = document.querySelector(".art").getBoundingClientRect(); let low = 0;
      document.querySelectorAll(".art > div:not(.coil):not(.photo):not(.bar)").forEach(e => { const r = e.getBoundingClientRect(); low = Math.max(low, r.bottom); });
      const bar = document.querySelector(".bar").getBoundingClientRect().top; return [Math.round(low - a.top), Math.round(bar - a.top)]; });
    console.log(name, "text bottom", over[0], "bar top", over[1]);
    await p.screenshot({ path: path.join(OUT, name + ".png") });
    await p.pdf({ path: path.join(OUT, name + ".pdf"), width: w + "px", height: h + "px", printBackground: true, margin: { top: 0, right: 0, bottom: 0, left: 0 } });
    fs.unlinkSync(hf); await p.close();
  }
  await b.close();
})();
