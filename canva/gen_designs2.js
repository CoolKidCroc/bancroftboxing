// CCCL starter designs v2: PNG (to look at) + PDF (to import into Canva).
// Each design = one background image (gradient, glow, grain, layered pines) + live text and boxes on top,
// so in Canva the background is one piece and every word is editable.
// usage: node gen_designs2.js OUTDIR LOGO_PNG [only-name-prefix]
const fs = require("fs"), path = require("path");
const { chromium } = require("playwright");
const OUT = process.argv[2], LOGO = process.argv[3], ONLY = process.argv[4] || "";
fs.mkdirSync(OUT, { recursive: true });
const BGDIR = path.join(OUT, "_bg"); fs.mkdirSync(BGDIR, { recursive: true });
const font = (f) => "file://" + path.resolve(__dirname, "node_modules/@fontsource/" + f);
const logo = "file://" + path.resolve(LOGO);

// ---- sample show (all of it is sample text to replace in Canva) ----
const S = { num: "CCCL 2", town: "Haliburton", date: "Saturday, May 1, 2027", short: "Sat May 1", venue: "A.J. LaRue Arena",
  doors: "Doors 6 PM", first: "First bout 7 PM" };
const PRICES = [["General admission", "$60"], ["VIP seat", "$100"], ["VIP table · 8-10 guests", "$1,500"]];

const CSS = `
@font-face{font-family:Anton;src:url(${font("anton/files/anton-latin-400-normal.woff2")})}
@font-face{font-family:Oswald;font-weight:500;src:url(${font("oswald/files/oswald-latin-500-normal.woff2")})}
@font-face{font-family:Oswald;font-weight:600;src:url(${font("oswald/files/oswald-latin-600-normal.woff2")})}
@font-face{font-family:Oswald;font-weight:700;src:url(${font("oswald/files/oswald-latin-700-normal.woff2")})}
@font-face{font-family:Inter;font-weight:400;src:url(${font("inter/files/inter-latin-400-normal.woff2")})}
@font-face{font-family:Inter;font-weight:600;src:url(${font("inter/files/inter-latin-600-normal.woff2")})}
:root{--ink:#151C19;--ink2:#0C110F;--pine:#354940;--pine2:#26362F;--cream:#F2ECDC;--red:#D9483B;--red2:#AE2221;--sage:#A9BDB0;--line:rgba(169,189,176,.22)}
*{box-sizing:border-box;margin:0;padding:0}
html,body{background:var(--ink)}
.c{position:relative;overflow:hidden;color:var(--cream);font-family:Inter,sans-serif;background-size:100% 100%;background-repeat:no-repeat}
.h{font-family:Anton,sans-serif;text-transform:uppercase;line-height:.9;letter-spacing:.004em}
.o{font-family:Oswald,sans-serif;text-transform:uppercase;letter-spacing:.06em}
.red{color:var(--red)} .sage{color:var(--sage)} .cream{color:var(--cream)}
.rule{height:4px;background:var(--red);border-radius:2px}
.logo{position:absolute;border-radius:50%;box-shadow:0 8px 28px rgba(0,0,0,.55)}
.pill{display:inline-flex;align-items:center;background:var(--cream);color:var(--ink);border-radius:6px;font-family:Oswald;font-weight:700;text-transform:uppercase;letter-spacing:.06em}
.kick{font-family:Oswald;font-weight:700;text-transform:uppercase;letter-spacing:.14em;color:var(--red)}
.ph{position:relative;overflow:hidden;background:linear-gradient(160deg,#2A3A33,#18221E);border:2px solid #3F544A;display:flex;align-items:center;justify-content:center}
.ph .lbl{position:absolute;left:0;right:0;bottom:7%;text-align:center;font-family:Oswald;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:#8EA398}
.ph svg{position:absolute;bottom:0;left:50%;transform:translateX(-50%);height:78%;opacity:.5}
.ph.full svg{height:64%;opacity:.42}
.ph.full .lbl{bottom:auto;top:30%}
.ph.story .lbl{top:44%}
.foot{position:absolute;left:0;right:0;bottom:0;display:flex;justify-content:space-between;align-items:center;font-family:Oswald;font-weight:600;text-transform:uppercase;letter-spacing:.08em;color:var(--sage)}
.chip{border:2px solid var(--line);border-radius:6px;padding:.35em .7em;font-family:Oswald;font-weight:600;text-transform:uppercase;letter-spacing:.06em}
.row{display:flex;justify-content:space-between;align-items:center;border-bottom:2px solid var(--line)}
.slogo{background:#EFE8D6;border:2px dashed #C9C0AA;color:#8A8574;display:flex;align-items:center;justify-content:center;font-family:Oswald;font-weight:600;letter-spacing:.12em;text-transform:uppercase;border-radius:6px}
`;
// fighter-in-guard silhouette for photo boxes (shows "a photo goes here" without looking empty)
const SIL = `<svg viewBox="0 0 200 220" fill="#0E1512"><circle cx="100" cy="72" r="40"/>
<path d="M14 220 C16 160 52 128 100 128 C148 128 184 160 186 220Z"/></svg>`;
const ph = (w, h, label, fs = 24, extra = "", cls = "") => `<div class="ph ${cls}" style="width:${w}px;height:${h}px;${extra}">${SIL}<div class="lbl" style="font-size:${fs}px">${label}</div></div>`;
const logoBox = (w, h, label, fs = 18, extra = "") => `<div class="slogo" style="width:${w};height:${h}px;font-size:${fs}px;${extra}">${label}</div>`;

// ---------- background: gradient + arena glow + grain + layered mountains and pines ----------
function rng(seed) { return () => { seed |= 0; seed = seed + 0x6D2B79F5 | 0; let t = Math.imul(seed ^ seed >>> 15, 1 | seed); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }
function pine(x, base, H, W) {
  const p = [[0, 0], [.45, .33], [.24, .33], [.72, .64], [.45, .64], [1, .93], [.1, .93], [.1, 1], [-.1, 1], [-.1, .93], [-1, .93], [-.45, .64], [-.72, .64], [-.24, .33], [-.45, .33]];
  return p.map(([dx, dy]) => `${(x + dx * W).toFixed(1)},${(base - H + dy * H).toFixed(1)}`).join(" ");
}
function ridge(w, h, seed = 7, foot = 0) {
  const r = rng(seed), g = h - foot;
  // far mountains
  let far = `0,${h} 0,${g * .55}`; const nf = 7;
  for (let i = 0; i <= nf; i++) { const x = w * i / nf, pk = x + w / nf * .5; far += ` ${x.toFixed(0)},${(g * (.42 + r() * .2)).toFixed(0)} ${pk.toFixed(0)},${(g * (.05 + r() * .25)).toFixed(0)}`; }
  far += ` ${w},${g * .5} ${w},${h}`;
  // mid hills
  let mid = `0,${h} 0,${g * .7}`; const nm = 11;
  for (let i = 0; i <= nm; i++) { const x = w * i / nm; mid += ` ${x.toFixed(0)},${(g * (.48 + r() * .18)).toFixed(0)}`; }
  mid += ` ${w},${h}`;
  // back pines (mid layer) and front pines
  const back = [], front = [];
  for (let x = -20; x < w + 40; x += w / 26 * (0.7 + r() * .6)) { const H = g * (.32 + r() * .2); back.push(`<polygon points="${pine(x, g * .82, H, H * .3)}"/>`); }
  for (let x = -30; x < w + 60; x += w / 15 * (0.6 + r() * .8)) { const H = g * (.42 + r() * .45); front.push(`<polygon points="${pine(x, g + 2, H, H * .3)}"/>`); }
  return `<svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" style="position:absolute;left:0;bottom:0;display:block">
  <defs><linearGradient id="fa" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#2B3D35"/><stop offset="1" stop-color="#1A2620"/></linearGradient>
  <linearGradient id="mi" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#1E2C26"/><stop offset="1" stop-color="#131C18"/></linearGradient>
  <linearGradient id="mist" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#151C19" stop-opacity="0"/><stop offset="1" stop-color="#1D2A24" stop-opacity=".9"/></linearGradient></defs>
  <polygon points="${far}" fill="url(#fa)" opacity=".85"/>
  <rect x="0" y="${g * .35}" width="${w}" height="${g * .5}" fill="url(#mist)"/>
  <polygon points="${mid}" fill="url(#mi)"/>
  <g fill="#17221D">${back.join("")}</g>
  <g fill="#0B100E">${front.join("")}</g>
  <rect x="0" y="${g}" width="${w}" height="${foot + 2}" fill="#0B100E"/>
  <rect x="0" y="${g}" width="${w}" height="3" fill="#AE2221" opacity=".0"/></svg>`;
}
function bgHTML(d) {
  const b = d.bg || {}, w = d.w, h = d.h;
  const glow = b.glow || "50% -8%";
  return `<div style="position:relative;width:${w}px;height:${h}px;overflow:hidden;
    background:radial-gradient(85% 60% at ${glow},rgba(242,236,220,.13),rgba(242,236,220,0) 70%),
      radial-gradient(90% 45% at 50% 104%,rgba(174,34,33,${b.redGlow ?? .22}),rgba(174,34,33,0) 70%),
      linear-gradient(180deg,#1C2924 0%,#151C19 45%,#0F1513 100%)">
    ${b.ridge === 0 ? "" : ridge(w, b.ridge || Math.round(h * .17), b.seed || 7, b.foot || 0)}
    <svg width="${w}" height="${h}" style="position:absolute;inset:0;mix-blend-mode:overlay;opacity:.32"><filter id="n"><feTurbulence type="fractalNoise" baseFrequency=".9" numOctaves="2" stitchTiles="stitch"/><feColorMatrix type="saturate" values="0"/></filter><rect width="100%" height="100%" filter="url(#n)"/></svg>
    <div style="position:absolute;inset:0;box-shadow:inset 0 0 ${Math.round(w * .18)}px rgba(0,0,0,.55)"></div>
  </div>`;
}

const footBar = (h, l, r, fs = 24, pad = 70) => `<div class="foot" style="height:${h}px;padding:0 ${pad}px;font-size:${fs}px"><span>${l}</span><span>${r}</span></div>`;

const D = [];
// 01 Tickets on sale
D.push({ name: "01 Tickets on sale (post 1080x1350)", w: 1080, h: 1350, bg: { ridge: 215, foot: 74, seed: 3 }, html: `
<div style="padding:64px 76px 0">
  <img class="logo" src="${logo}" style="width:200px;right:66px;top:62px">
  <div class="kick" style="font-size:30px">Cottage Country Combat League</div>
  <div class="o sage" style="font-size:24px;font-weight:600;margin-top:6px">${S.num} · Amateur MMA</div>
  <div class="h" style="font-size:226px;margin-top:18px">Fight<br>Night</div>
  <div class="h red" style="font-size:150px;margin-top:4px">${S.town}</div>
  <div style="display:flex;align-items:center;gap:22px;margin-top:30px"><div class="rule" style="width:90px"></div>
    <div class="o" style="font-size:46px;font-weight:600">${S.date}</div></div>
  <div class="o sage" style="font-size:30px;font-weight:500;margin-top:10px;padding-left:112px">${S.venue} · ${S.doors} · ${S.first}</div>
  <div style="display:flex;gap:14px;margin-top:40px">${PRICES.map((p, i) => `
    <div style="flex:${i == 2 ? 1.25 : 1};border:2px solid ${i == 1 ? "var(--red)" : "var(--line)"};border-radius:8px;padding:16px 18px;background:rgba(21,28,25,.55)">
      <div class="o sage" style="font-size:19px;font-weight:600">${p[0]}</div><div class="h" style="font-size:58px;margin-top:6px">${p[1]}</div></div>`).join("")}</div>
  <div style="display:flex;gap:20px;align-items:center;margin-top:26px">
    <span class="pill" style="font-size:34px;padding:12px 24px">Tickets on sale now</span>
    <span class="o" style="font-size:28px;font-weight:600">Link in bio · Eventbrite</span></div>
</div>
${footBar(74, "Sanctioned by the Ontario Athletics Commission", "Amateur MMA", 20)}` });

// 02 Countdown story
D.push({ name: "02 Countdown story (1080x1920)", w: 1080, h: 1920, bg: { ridge: 420, foot: 150, seed: 11 }, html: `
<div style="padding:175px 90px 0;text-align:center">
  <img src="${logo}" style="width:220px;border-radius:50%;box-shadow:0 8px 28px rgba(0,0,0,.55)">
  <div class="kick" style="font-size:32px;margin-top:34px">${S.num} · ${S.town}</div>
  <div class="h red" style="font-size:520px;line-height:.82;margin-top:44px">7</div>
  <div class="h" style="font-size:150px;margin-top:6px">Days</div>
  <div class="o sage" style="font-size:42px;font-weight:600;margin-top:16px">Until fight night</div>
  <div style="margin:56px auto 0;max-width:820px;border:2px solid var(--line);border-radius:10px;padding:26px 30px;background:rgba(21,28,25,.6)">
    <div class="kick" style="font-size:24px">Main event</div>
    <div class="h" style="font-size:72px;margin-top:10px">Fighter name<br><span class="red">vs</span> Fighter name</div></div>
  <div class="o" style="font-size:42px;font-weight:600;margin-top:40px">${S.short} · ${S.venue}</div>
  <span class="pill" style="font-size:40px;padding:16px 34px;margin-top:36px">Get tickets</span>
</div>` });

// 03 Poster (early, no card yet)
D.push({ name: "03 Poster (11x17 in)", w: 1056, h: 1632, pdfIn: [11, 17], scale: 3, bg: { ridge: 330, foot: 176, seed: 5 }, html: `
<div style="padding:76px 80px 0">
  <div style="display:flex;justify-content:space-between;align-items:flex-start">
    <div style="padding-top:18px"><div class="kick" style="font-size:30px">Cottage Country Combat League</div>
      <div class="o sage" style="font-size:22px;font-weight:600;margin-top:10px;line-height:1.35">Live amateur MMA<br>Sanctioned by the Ontario Athletics Commission</div></div>
    <img src="${logo}" style="width:230px;border-radius:50%;box-shadow:0 8px 28px rgba(0,0,0,.55)"></div>
  <div class="h" style="font-size:240px;margin-top:18px">Fight Night</div>
  <div class="h red" style="font-size:168px;margin-top:4px">${S.town}</div>
  <div style="display:flex;align-items:center;gap:24px;margin-top:30px"><div class="rule" style="width:110px"></div>
    <div class="o" style="font-size:58px;font-weight:600">${S.date}</div></div>
  <div class="o sage" style="font-size:31px;font-weight:500;margin-top:10px">${S.venue} · ${S.doors} · ${S.first}</div>
  <div style="display:flex;gap:34px;margin-top:38px;align-items:stretch">
    <div class="slogo" style="width:220px;height:220px;flex-shrink:0;font-size:22px;text-align:center;line-height:1.3;background:#F2ECDC">QR code<br>poster's<br>tracking link</div>
    <div style="flex:1;border:2px solid var(--line);border-radius:10px;padding:22px 28px;background:rgba(21,28,25,.55)">
      <div class="o" style="font-size:40px;font-weight:700">Tickets on Eventbrite</div>
      ${PRICES.map(p => `<div style="display:flex;justify-content:space-between;font-family:Oswald;font-weight:600;font-size:28px;text-transform:uppercase;letter-spacing:.05em;margin-top:12px"><span class="sage">${p[0]}</span><span>${p[1]}</span></div>`).join("")}</div></div>
</div>
<div style="position:absolute;left:80px;right:80px;bottom:34px">
  <div class="o sage" style="font-size:21px;font-weight:600;margin-bottom:12px">Presented by</div>
  <div style="display:flex;gap:14px">${logoBox("100%", 92, "Title sponsor", 18, "flex:1.6")}${[1, 2, 3].map(() => logoBox("100%", 92, "Sponsor", 16, "flex:1")).join("")}</div></div>` });

// 04 Fight card
const B = [["155 lb", "Main event"], ["170 lb", "Co-main"], ["145 lb"], ["185 lb"], ["135 lb"], ["205 lb"], ["160 lb"]];
D.push({ name: "04 Fight card (post 1080x1350)", w: 1080, h: 1350, bg: { ridge: 230, foot: 74, seed: 9 }, html: `
<div style="padding:66px 70px 0">
  <img class="logo" src="${logo}" style="width:160px;right:60px;top:52px">
  <div class="kick" style="font-size:28px">${S.num} · ${S.town} · ${S.short}</div>
  <div class="h" style="font-size:150px;margin-top:10px">Fight card</div>
  <div style="margin-top:36px;border:3px solid var(--red);border-radius:10px;padding:24px 30px;background:rgba(53,73,64,.38)">
    <div class="kick" style="font-size:22px">Main event · ${B[0][0]}</div>
    <div class="h" style="font-size:70px;margin-top:8px">Fighter name <span class="red">vs</span> Fighter name</div></div>
  <div style="margin-top:14px;border:2px solid var(--line);border-radius:10px;padding:18px 30px;background:rgba(21,28,25,.5)">
    <div class="o sage" style="font-size:20px;font-weight:700;letter-spacing:.14em">Co-main · ${B[1][0]}</div>
    <div class="h" style="font-size:52px;margin-top:6px">Fighter name <span class="red">vs</span> Fighter name</div></div>
  <div style="margin-top:8px">${B.slice(2).map(b => `
    <div class="row" style="padding:15px 8px">
      <div class="o" style="font-size:38px;font-weight:600">Fighter name <span class="red">vs</span> Fighter name</div>
      <div class="o sage" style="font-size:26px;font-weight:600">${b[0]}</div></div>`).join("")}</div>
  <div class="o sage" style="font-size:24px;font-weight:600;margin-top:18px">+ Youth grappling</div>
</div>
${footBar(74, `${S.doors} · ${S.first}`, "Tickets: link in bio", 22)}` });

// 05 Fighter vs fighter
const corner = (side, col) => `
  <div style="flex:1;display:flex;flex-direction:column;align-items:center">
    <div style="width:100%;height:10px;background:${col};border-radius:4px 4px 0 0"></div>
    ${ph(450, 590, "Fighter photo", 24, "border-top:none")}
    <div class="h" style="font-size:78px;margin-top:24px;text-align:center">Fighter<br>name</div>
    <div class="o sage" style="font-size:26px;font-weight:600;margin-top:12px">Gym name · Town</div>
    <div class="o" style="font-size:28px;font-weight:600;margin-top:6px">2-1 · ${side} corner</div></div>`;
D.push({ name: "05 Fighter vs fighter (post 1080x1350)", w: 1080, h: 1350, bg: { ridge: 170, foot: 74, seed: 13, glow: "50% 30%" }, html: `
<div style="padding:54px 60px 0">
  <div style="text-align:center"><div class="kick" style="font-size:30px">Main event · 155 lb</div>
  <div class="o" style="font-size:26px;font-weight:600;margin-top:6px">${S.num} · ${S.town} · ${S.short}</div></div>
  <div style="display:flex;gap:40px;margin-top:34px;position:relative">${corner("Red", "#AE2221")}${corner("Blue", "#3E5B7A")}
    <div class="h red" style="position:absolute;left:50%;top:250px;transform:translateX(-50%);font-size:160px;text-shadow:0 8px 26px rgba(0,0,0,.7)">VS</div></div>
</div>
<img class="logo" src="${logo}" style="width:120px;left:50%;transform:translateX(-50%);bottom:104px">
${footBar(74, S.venue, "Tickets: link in bio", 22, 60)}` });

// 06 Fighter spotlight
D.push({ name: "06 Fighter spotlight (post 1080x1350)", w: 1080, h: 1350, bg: { ridge: 0 }, html: `
${ph(1080, 830, "Fighter photo (full width)", 30, "position:absolute;left:0;top:0;border:none", "full")}
<div style="position:absolute;left:0;right:0;top:560px;height:280px;background:linear-gradient(rgba(21,28,25,0),#151C19)"></div>
<img class="logo" src="${logo}" style="width:140px;right:48px;top:48px">
<div style="position:absolute;left:70px;right:70px;top:700px">
  <div class="kick" style="font-size:30px">Fighter spotlight</div>
  <div class="h" style="font-size:150px;margin-top:8px">Fighter name</div>
  <div style="display:flex;gap:12px;margin-top:22px;flex-wrap:wrap">${["Gym name", "Hometown", "155 lb", "2-1"].map(c => `<span class="chip" style="font-size:24px">${c}</span>`).join("")}</div>
  <div style="font-family:Inter;font-size:34px;line-height:1.45;margin-top:26px;max-width:900px">"One line from the fighter: why they fight, who they're fighting for."</div>
</div>
${footBar(80, `${S.num} · ${S.town} · ${S.short}`, '<span class="red">Tickets: link in bio</span>', 24)}` });

// 07 Sponsor thank-you
const tier = (label, n, h, fs = 16) => `<div class="o sage" style="font-size:22px;font-weight:700;letter-spacing:.14em;margin:22px 0 10px">${label}</div>
  <div style="display:flex;gap:14px">${Array.from({ length: n }, () => logoBox("100%", h, "Sponsor logo", fs, "flex:1")).join("")}</div>`;
D.push({ name: "07 Sponsor thank-you (post 1080x1350)", w: 1080, h: 1350, bg: { ridge: 170, foot: 74, seed: 17 }, html: `
<div style="padding:66px 70px 0">
  <img class="logo" src="${logo}" style="width:150px;right:60px;top:52px">
  <div class="kick" style="font-size:28px">${S.num} · ${S.town}</div>
  <div class="h" style="font-size:132px;margin-top:10px">Thank you<br>to our sponsors</div>
  <div class="o sage" style="font-size:22px;font-weight:700;letter-spacing:.14em;margin:28px 0 10px">Title sponsor</div>
  ${logoBox("100%", 150, "Title sponsor logo", 22)}
  ${tier("Gold", 2, 112, 18)}${tier("Silver", 4, 82)}
  <div class="o" style="font-size:24px;font-weight:600;margin-top:22px;line-height:1.5"><span class="sage">Bronze:</span> Name · Name · Name · Name<br><span class="sage">Community:</span> Name · Name · Name</div>
</div>
${footBar(74, "Want in? Sponsor the next show", "Message us", 22)}` });

// 08 Results
D.push({ name: "08 Results (post 1080x1350)", w: 1080, h: 1350, bg: { ridge: 200, foot: 74, seed: 19 }, html: `
<div style="padding:66px 70px 0">
  <img class="logo" src="${logo}" style="width:160px;right:60px;top:52px">
  <div class="kick" style="font-size:28px">${S.num} · ${S.town} · Official results</div>
  <div class="h" style="font-size:160px;margin-top:10px">Results</div>
  <div style="margin-top:34px">${["Sub · R2", "TKO · R1", "Decision", "Sub · R1", "Decision", "TKO · R2", "Decision", "Decision"].map((m, i) => `
    <div class="row" style="padding:15px 8px">
      <div class="o" style="font-size:36px;font-weight:600">Winner name <span class="sage" style="font-weight:500;text-transform:lowercase">def.</span> Opponent</div>
      <span class="chip" style="font-size:20px;color:${i == 0 ? "var(--cream)" : "var(--sage)"};border-color:${i == 0 ? "var(--red)" : "var(--line)"}">${m}</span></div>`).join("")}</div>
</div>
${footBar(74, "Thank you to every fighter, coach and fan", "Next show: date", 21)}` });

// 09 Facebook event cover
D.push({ name: "09 Facebook event cover (1920x1005)", w: 1920, h: 1005, bg: { ridge: 250, seed: 23, glow: "35% -10%" }, html: `
<div style="padding:120px 150px 0">
  <img class="logo" src="${logo}" style="width:360px;right:170px;top:150px">
  <div class="kick" style="font-size:40px">Cottage Country Combat League · ${S.num}</div>
  <div class="h" style="font-size:236px;margin-top:20px">Fight night <span class="red">${S.town}</span></div>
  <div style="display:flex;align-items:center;gap:26px;margin-top:34px"><div class="rule" style="width:110px"></div>
    <div class="o" style="font-size:60px;font-weight:600">${S.date} · ${S.venue}</div></div>
  <div style="display:flex;align-items:center;gap:22px;margin-top:30px;padding-left:136px">
    <span class="pill" style="font-size:34px;padding:10px 24px">Tickets on Eventbrite</span>
    <span class="o sage" style="font-size:32px;font-weight:600">Amateur MMA · Sanctioned by the Ontario Athletics Commission</span></div>
</div>` });

// 10 Poster with the main event (once the card is set)
D.push({ name: "10 Poster with main event (11x17 in)", w: 1056, h: 1632, pdfIn: [11, 17], scale: 3, bg: { ridge: 290, foot: 120, seed: 29, glow: "50% 34%" }, html: `
<div style="padding:66px 80px 0">
  <img class="logo" src="${logo}" style="width:178px;right:72px;top:56px">
  <div class="kick" style="font-size:28px">Cottage Country Combat League · ${S.num}</div>
  <div class="h" style="font-size:150px;margin-top:14px">Fight Night</div>
  <div class="h red" style="font-size:122px;margin-top:2px">${S.town}</div>
  <div style="display:flex;gap:30px;margin-top:38px;position:relative">
    ${["Red", "Blue"].map((c, i) => `<div style="flex:1;text-align:center">
      <div style="height:9px;background:${i ? "#3E5B7A" : "#AE2221"};border-radius:4px 4px 0 0"></div>
      ${ph(433, 440, "Fighter photo", 22, "border-top:none")}
      <div class="h" style="font-size:60px;margin-top:16px">Fighter name</div>
      <div class="o sage" style="font-size:22px;font-weight:600;margin-top:6px">Gym · Town · 2-1</div></div>`).join("")}
    <div style="position:absolute;left:50%;top:140px;transform:translateX(-50%);text-align:center">
      <div class="kick" style="font-size:22px;color:var(--cream);background:var(--red2);padding:6px 14px;border-radius:4px;white-space:nowrap">Main event</div>
      <div class="h red" style="font-size:150px;margin-top:6px;text-shadow:0 8px 26px rgba(0,0,0,.7)">VS</div></div></div>
  <div class="o sage" style="font-size:22px;font-weight:600;margin-top:14px;text-align:center">+ a full undercard of amateur MMA</div>
  <div style="display:flex;align-items:center;gap:24px;margin-top:28px"><div class="rule" style="width:100px"></div>
    <div class="o" style="font-size:52px;font-weight:600">${S.date}</div></div>
  <div class="o sage" style="font-size:29px;font-weight:500;margin-top:8px">${S.venue} · ${S.doors} · ${S.first}</div>
  <div style="display:flex;gap:24px;margin-top:30px;align-items:stretch">
    <div class="slogo" style="width:184px;height:184px;flex-shrink:0;font-size:19px;text-align:center;line-height:1.3;background:#F2ECDC">QR code<br>poster's<br>tracking link</div>
    <div style="flex:1;border:2px solid var(--line);border-radius:10px;padding:14px 24px;background:rgba(21,28,25,.6)">
      ${PRICES.map(p => `<div style="display:flex;justify-content:space-between;font-family:Oswald;font-weight:600;font-size:25px;text-transform:uppercase;letter-spacing:.05em;margin-top:8px"><span class="sage">${p[0]}</span><span>${p[1]}</span></div>`).join("")}
      <div class="o" style="font-size:23px;font-weight:700;margin-top:12px">Tickets on Eventbrite</div></div>
    <div style="width:230px;display:flex;flex-direction:column;gap:12px">${logoBox("100%", 86, "Title sponsor", 16)}${logoBox("100%", 86, "Sponsor", 16)}</div></div>
</div>
${footBar(120, "Scan for tickets", "Amateur MMA · Sanctioned by the OAC", 24, 80)}` });

// 11 I'm fighting (post) — fighters share this with their own tracking link
D.push({ name: "11 Fighter - I'm fighting (post 1080x1350)", w: 1080, h: 1350, bg: { ridge: 0 }, html: `
${ph(1080, 860, "Your photo (full width)", 30, "position:absolute;left:0;top:0;border:none", "full")}
<div style="position:absolute;left:0;right:0;top:520px;height:340px;background:linear-gradient(rgba(21,28,25,0),#151C19 92%)"></div>
<img class="logo" src="${logo}" style="width:140px;right:48px;top:48px">
<div class="kick" style="position:absolute;left:66px;top:62px;font-size:26px;color:var(--cream);background:var(--red2);padding:8px 16px;border-radius:4px">${S.num} · ${S.town}</div>
<div style="position:absolute;left:66px;right:66px;top:640px">
  <div class="h" style="font-size:198px">I'm <span class="red">fighting</span></div>
  <div class="o" style="font-size:52px;font-weight:600;margin-top:18px">Fighter name <span class="sage" style="font-weight:500">vs</span> Opponent name</div>
  <div style="display:flex;gap:12px;margin-top:20px;flex-wrap:wrap">${["155 lb", "Amateur MMA", "Gym name"].map(c => `<span class="chip" style="font-size:24px">${c}</span>`).join("")}</div>
  <div class="o sage" style="font-size:34px;font-weight:600;margin-top:28px">${S.date} · ${S.venue}</div>
  <div style="display:flex;gap:18px;align-items:center;margin-top:30px">
    <span class="pill" style="font-size:32px;padding:12px 24px">Get tickets with my link</span>
    <span class="o" style="font-size:28px;font-weight:600">Link in bio</span></div>
</div>
${footBar(74, "Cottage Country Combat League", "Amateur MMA", 21, 66)}` });

// 12 I'm fighting (story) — leaves a spot for the Instagram link sticker
D.push({ name: "12 Fighter - I'm fighting (story 1080x1920)", w: 1080, h: 1920, bg: { ridge: 0 }, html: `
${ph(1080, 1080, "Your photo", 34, "position:absolute;left:0;top:0;border:none", "full story")}
<div style="position:absolute;left:0;right:0;top:700px;height:400px;background:linear-gradient(rgba(21,28,25,0),#151C19 92%)"></div>
<img src="${logo}" style="position:absolute;left:50%;transform:translateX(-50%);top:170px;width:190px;border-radius:50%;box-shadow:0 8px 28px rgba(0,0,0,.55)">
<div style="position:absolute;left:70px;right:70px;top:800px;text-align:center">
  <div class="h" style="font-size:230px">I'm <span class="red">fighting</span></div>
  <div class="o" style="font-size:54px;font-weight:600;margin-top:22px">Fighter name <span class="sage" style="font-weight:500">vs</span> Opponent</div>
  <div class="o sage" style="font-size:38px;font-weight:600;margin-top:16px">${S.num} · ${S.town} · ${S.short}</div>
  <div class="o" style="font-size:32px;font-weight:600;margin-top:8px">${S.venue} · ${S.first}</div>
  <div style="margin:44px auto 0;width:640px;height:150px;border:3px dashed var(--sage);border-radius:16px;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:6px">
    <div class="o" style="font-size:30px;font-weight:700">Put your link sticker here</div>
    <div style="font-size:22px;color:var(--sage)">Instagram: Stickers > Link > your ticket link</div></div>
  <div class="o" style="font-size:34px;font-weight:600;margin-top:26px">Tickets with my link <span class="red">↑</span></div>
</div>` });

// 13 Weigh-in results (fight day)
D.push({ name: "13 Weigh-in results (post 1080x1350)", w: 1080, h: 1350, bg: { ridge: 200, foot: 74, seed: 31 }, html: `
<div style="padding:66px 70px 0">
  <img class="logo" src="${logo}" style="width:150px;right:60px;top:52px">
  <div class="kick" style="font-size:28px">${S.num} · ${S.town} · Fight day</div>
  <div class="h" style="font-size:136px;margin-top:10px">Weigh-ins<br><span class="red">done</span></div>
  <div class="o sage" style="font-size:28px;font-weight:600;margin-top:16px">Every bout is on · ${S.doors} · ${S.first}</div>
  <div style="margin-top:26px">${[["155.0", "154.4"], ["169.8", "170.0"], ["144.6", "145.0"], ["184.2", "185.0"], ["135.0", "134.8"], ["204.0", "203.6"]].map(w => `
    <div class="row" style="padding:15px 8px">
      <div class="o" style="font-size:34px;font-weight:600;flex:1">Fighter name <span class="sage" style="font-weight:500">${w[0]}</span></div>
      <div class="h red" style="font-size:30px;width:70px;text-align:center">vs</div>
      <div class="o" style="font-size:34px;font-weight:600;flex:1;text-align:right"><span class="sage" style="font-weight:500">${w[1]}</span> Fighter name</div></div>`).join("")}</div>
</div>
${footBar(74, "Official weights", "Last tickets: link in bio", 22)}` });

// 14 Win post (after the show; for a fighter who lost, change WIN to THANK YOU)
D.push({ name: "14 Fighter - Win (post 1080x1350)", w: 1080, h: 1350, bg: { ridge: 0 }, html: `
${ph(1080, 900, "Fight photo (full width)", 30, "position:absolute;left:0;top:0;border:none", "full")}
<div style="position:absolute;left:0;right:0;top:560px;height:340px;background:linear-gradient(rgba(21,28,25,0),#151C19 92%)"></div>
<img class="logo" src="${logo}" style="width:140px;right:48px;top:48px">
<div style="position:absolute;left:66px;right:66px;top:660px">
  <div class="kick" style="font-size:28px">${S.num} · ${S.town} · Official result</div>
  <div class="h" style="font-size:250px;margin-top:6px">Win</div>
  <div class="o" style="font-size:50px;font-weight:600;margin-top:14px">Fighter name <span class="sage" style="font-weight:500;text-transform:lowercase">def.</span> Opponent</div>
  <div style="display:flex;gap:12px;margin-top:18px">${["Submission · Round 2", "155 lb"].map((c, i) => `<span class="chip" style="font-size:24px;${i ? "" : "border-color:var(--red)"}">${c}</span>`).join("")}</div>
  <div style="font-family:Inter;font-size:30px;line-height:1.45;margin-top:24px;max-width:900px">"Thank you to my coaches, my team and everyone who came out."</div>
</div>
${footBar(74, "Cottage Country Combat League", "Amateur MMA", 21, 66)}` });

// 15 Proud sponsor (sponsors share this)
D.push({ name: "15 Sponsor - Proud sponsor (post 1080x1350)", w: 1080, h: 1350, bg: { ridge: 230, foot: 74, seed: 37, glow: "50% 40%" }, html: `
<div style="padding:110px 80px 0;text-align:center">
  <div class="kick" style="font-size:30px">Proud sponsor of</div>
  <div class="h" style="font-size:120px;margin-top:16px">Fight Night <span class="red">${S.town}</span></div>
  <div class="o sage" style="font-size:30px;font-weight:600;margin-top:16px">${S.num} · ${S.date}</div>
  ${logoBox("820px", 360, "Your logo", 34, "margin:56px auto 0")}
  <div style="display:flex;align-items:center;justify-content:center;gap:26px;margin-top:50px">
    <img src="${logo}" style="width:150px;border-radius:50%;box-shadow:0 8px 28px rgba(0,0,0,.55)">
    <div style="text-align:left"><div class="o" style="font-size:30px;font-weight:700">Cottage Country Combat League</div>
      <div class="o sage" style="font-size:24px;font-weight:600;margin-top:6px">${S.venue} · Amateur MMA</div></div></div>
</div>
${footBar(74, "See you there", "Tickets: link in bio", 22)}` });

// 16 Fight card (story)
D.push({ name: "16 Fight card (story 1080x1920)", w: 1080, h: 1920, bg: { ridge: 380, foot: 150, seed: 41 }, html: `
<div style="padding:170px 80px 0">
  <div style="display:flex;justify-content:space-between;align-items:flex-start">
    <div><div class="kick" style="font-size:30px">${S.num} · ${S.town}</div>
      <div class="h" style="font-size:150px;margin-top:10px">Fight<br>card</div></div>
    <img src="${logo}" style="width:200px;border-radius:50%;box-shadow:0 8px 28px rgba(0,0,0,.55)"></div>
  <div style="margin-top:40px;border:3px solid var(--red);border-radius:10px;padding:26px 30px;background:rgba(53,73,64,.38)">
    <div class="kick" style="font-size:24px">Main event · 155 lb</div>
    <div class="h" style="font-size:76px;margin-top:8px">Fighter name<br><span class="red">vs</span> Fighter name</div></div>
  <div style="margin-top:14px">${["170 lb", "145 lb", "185 lb", "135 lb", "205 lb", "160 lb"].map(w => `
    <div class="row" style="padding:15px 8px">
      <div class="o" style="font-size:40px;font-weight:600">Fighter name <span class="red">vs</span> Fighter name</div>
      <div class="o sage" style="font-size:26px;font-weight:600">${w}</div></div>`).join("")}</div>
  <div class="o" style="font-size:34px;font-weight:600;margin-top:30px;text-align:center">${S.short} · ${S.venue} · ${S.first}</div>
  <div style="text-align:center;margin-top:26px"><span class="pill" style="font-size:36px;padding:14px 30px">Get tickets</span></div>
</div>` });

(async () => {
  const b = await chromium.launch();
  const report = [];
  for (const d of D) {
    if (ONLY && !d.name.startsWith(ONLY)) continue;
    const sc = d.scale || 2;
    // 1) background image
    const bp = await b.newPage({ viewport: { width: d.w, height: d.h }, deviceScaleFactor: sc });
    await bp.setContent(`<!doctype html><html><head><style>*{margin:0;padding:0}body{background:#151C19}</style></head><body>${bgHTML(d)}</body></html>`);
    const bgFile = path.join(BGDIR, d.name + ".jpg");
    await bp.screenshot({ path: bgFile, type: "jpeg", quality: 90, clip: { x: 0, y: 0, width: d.w, height: d.h } });
    await bp.close();
    // 2) design with live text on top
    const p = await b.newPage({ viewport: { width: d.w, height: d.h } });
    const hf = path.join(OUT, "_" + d.name + ".html");
    fs.writeFileSync(hf, `<!doctype html><html><head><meta charset="utf-8"><style>${CSS}@page{margin:0;size:${d.pdfIn ? d.pdfIn[0] + "in " + d.pdfIn[1] + "in" : d.w + "px " + d.h + "px"}}</style></head><body>
      <div class="c" style="width:${d.w}px;height:${d.h}px;background-image:url('${"file://" + path.resolve(bgFile)}')">${d.html}</div></body></html>`);
    await p.goto("file://" + path.resolve(hf), { waitUntil: "load" });
    await p.evaluate(() => document.fonts.ready); await p.waitForTimeout(250);
    // overflow check: any element poking outside the canvas or text running past its box
    const issues = await p.evaluate(([W, H]) => {
      const out = [];
      document.querySelectorAll(".c *").forEach(e => {
        const r = e.getBoundingClientRect(); if (!r.width) return;
        if (r.right > W + 1 || r.bottom > H + 1 || r.left < -1) out.push((e.className || e.tagName) + " " + Math.round(r.left) + "," + Math.round(r.top) + "-" + Math.round(r.right) + "," + Math.round(r.bottom));
        if (e.scrollWidth > e.clientWidth + 2 && getComputedStyle(e).overflow !== "visible") out.push("clip " + (e.className || e.tagName));
      });
      let low = 0; document.querySelectorAll(".c > div:not(.foot):not(.ph) *").forEach(e => { if (e.closest(".ph") || e.closest(".foot")) return; const r = e.getBoundingClientRect(); if (r.width && r.bottom > low) low = r.bottom; });
      out.push("lowest content y=" + Math.round(low));
      return out.slice(-7);
    }, [d.w, d.h]);
    await p.screenshot({ path: path.join(OUT, d.name + ".png"), clip: { x: 0, y: 0, width: d.w, height: d.h } });
    await p.pdf({ path: path.join(OUT, d.name + ".pdf"), width: d.pdfIn ? d.pdfIn[0] + "in" : d.w + "px", height: d.pdfIn ? d.pdfIn[1] + "in" : d.h + "px", printBackground: true, margin: { top: 0, right: 0, bottom: 0, left: 0 } });
    const fonts = await p.evaluate(() => [...document.fonts].filter(f => f.status === "loaded").length);
    report.push(`${d.name} | fonts loaded ${fonts} | ${issues.join(" ; ")}`);
    fs.unlinkSync(hf); await p.close();
  }
  await b.close();
  console.log(report.join("\n"));
})();
