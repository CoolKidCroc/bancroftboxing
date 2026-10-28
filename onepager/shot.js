const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({viewport:{width:816,height:1200}});
  await p.goto('file://' + process.argv[2], { waitUntil: 'load' });
  await p.screenshot({path: process.argv[3], fullPage: true}); await b.close();
})();
