const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({viewport:{width:816,height:1056}});
  await p.goto('file://' + process.argv[2], { waitUntil: 'load' });
  const r = await p.evaluate(() => { const o = {}; for (const s of ['.top','.body','.foot']) { const e=document.querySelector(s); o[s]=[e.scrollHeight,e.clientHeight]; } return o; });
  console.log(JSON.stringify(r)); await b.close();
})();
