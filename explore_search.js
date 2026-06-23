// explore_search.js — type a query into the global search box and capture the
// autocomplete/search responses (to learn the name -> company-hash mapping).
const { chromium } = require('playwright');
const path = require('path');

const USER_DATA_DIR = path.join(__dirname, '.session');
const queries = process.argv.slice(2);
if (!queries.length) queries.push('NEOGROWTH', 'AADISHAKTI');

(async () => {
  const ctx = await chromium.launchPersistentContext(USER_DATA_DIR, { headless: true });
  const page = ctx.pages()[0] || (await ctx.newPage());

  const captured = [];
  page.on('response', async (resp) => {
    const u = resp.url();
    if (/search|autocomplete|suggest|getcomp|lookup|ac\.aspx|\.ashx/i.test(u) && resp.request().method() !== 'OPTIONS') {
      let body = '';
      try { body = (await resp.text()).slice(0, 1200); } catch {}
      captured.push({ status: resp.status(), method: resp.request().method(), url: u, postData: resp.request().postData()?.slice(0, 300), body });
    }
  });

  await page.goto('https://www.saverisk.com/myorders.aspx', { waitUntil: 'networkidle', timeout: 60000 });

  for (const q of queries) {
    console.log('\n================ QUERY:', q, '================');
    captured.length = 0;
    const box = page.locator('#txtsearch');
    await box.click();
    await box.fill('');
    await box.type(q, { delay: 90 });
    await page.waitForTimeout(2500); // let autocomplete fire
    // dump any visible suggestion dropdown
    const sugg = await page.evaluate(() => {
      const items = [...document.querySelectorAll('.ui-autocomplete li, .autocomplete li, [id*=suggest] li, ul li a')]
        .map((e) => ({ t: (e.innerText || '').trim().slice(0, 70), href: e.getAttribute && e.getAttribute('href') }))
        .filter((x) => x.t).slice(0, 15);
      return items;
    });
    captured.forEach((c) => {
      console.log(`  [XHR] ${c.status} ${c.method} ${c.url}`);
      if (c.postData) console.log('        POST:', c.postData);
      if (c.body) console.log('        BODY:', c.body.replace(/\n/g, ' ').slice(0, 600));
    });
    console.log('  [DROPDOWN]'); sugg.forEach((s) => console.log('    -', JSON.stringify(s)));
  }

  await ctx.close();
  process.exit(0);
})().catch((e) => { console.error('ERR', e); process.exit(1); });
