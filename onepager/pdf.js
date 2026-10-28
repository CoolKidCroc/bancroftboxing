const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage();
  await p.goto('file://' + process.argv[2], { waitUntil: 'load' }); await p.waitForTimeout(300);
  const h = await p.evaluate(() => document.querySelector('.page').scrollHeight); console.log('page height px', h, '(Letter = 1056)');
  await p.pdf({ path: process.argv[3], format: 'Letter', printBackground: true, margin: {top:0,bottom:0,left:0,right:0} });
  await b.close();
})();
