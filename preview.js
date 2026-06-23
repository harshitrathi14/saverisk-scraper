const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const dir = process.argv[2];
  const b = await chromium.launch({ headless: true });
  const p = await b.newPage({ viewport: { width: 1240, height: 1600 } });
  await p.goto('file://' + path.join(dir, 'dashboard.html'));
  await p.waitForTimeout(600);
  await p.screenshot({ path: 'dashboard_preview.png', fullPage: true });
  await b.close();
  console.log('ok');
})().catch(e => { console.error(e.message); process.exit(1); });
