// explore_charges.js — open a company page, click the Charges menu, capture the AJAX
// calls and the rendered charge grid(s). Saves screenshots + html + a log of XHRs.
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const USER_DATA_DIR = path.join(__dirname, '.session');
const COMPANY_URL = process.argv[2] ||
  'https://www.saverisk.com/company/0x922933F85C2BE38B55711BEDF5487882/neogrowth credit private limited';

(async () => {
  const ctx = await chromium.launchPersistentContext(USER_DATA_DIR, { headless: true });
  const page = ctx.pages()[0] || (await ctx.newPage());

  const xhrLog = [];
  page.on('response', async (resp) => {
    const u = resp.url();
    if (/charge|grid|dashboard|getcompany|ashx|handler|json/i.test(u) && resp.request().method() !== 'OPTIONS') {
      let ct = resp.headers()['content-type'] || '';
      xhrLog.push(`${resp.status()} ${resp.request().method()} ${ct.split(';')[0]}  ${u}`);
    }
  });

  await page.goto(COMPANY_URL, { waitUntil: 'networkidle', timeout: 60000 });
  await page.waitForTimeout(1500);

  // Click the Charges left-menu item: element whose data attr contains "Charges~Charges"
  const chargesMenu = page.locator('[data*="Charges~Charges"]').first();
  const n = await chargesMenu.count();
  console.log('Charges menu elements found:', n);
  if (n > 0) {
    await chargesMenu.scrollIntoViewIfNeeded().catch(() => {});
    await chargesMenu.click({ timeout: 8000 }).catch((e) => console.log('click err', e.message));
    await page.waitForLoadState('networkidle').catch(() => {});
    await page.waitForTimeout(3500);
  }

  await page.screenshot({ path: 'explore_charges.png', fullPage: true }).catch(() => {});

  // Capture the main content panel HTML (right side), and list any sub-tabs that appeared
  const info = await page.evaluate(() => {
    const subtabs = [...document.querySelectorAll('[data*="charge"],[data*="Charge"]')]
      .map((e) => ({ id: e.id, data: e.getAttribute('data'), txt: (e.innerText || '').trim().slice(0, 40) }))
      .filter((e) => e.data).slice(0, 40);
    // find tables with charge-ish headers
    const tables = [...document.querySelectorAll('table')].map((t) => {
      const head = [...t.querySelectorAll('th')].map((th) => th.innerText.trim()).filter(Boolean);
      const firstRow = [...(t.querySelector('tbody tr')?.querySelectorAll('td') || [])].map((td) => td.innerText.trim().slice(0, 30));
      return { id: t.id, rows: t.querySelectorAll('tr').length, head: head.slice(0, 15), firstRow: firstRow.slice(0, 15) };
    }).filter((t) => t.rows > 1 && (t.head.join(' ') + t.firstRow.join(' ')).length > 0);
    return { subtabs, tables };
  });

  fs.writeFileSync('explore_charges_panel.html', await page.content());
  console.log('\n--- XHR / network (charge-ish) ---'); xhrLog.forEach((l) => console.log(l));
  console.log('\n--- charge sub-tabs ---'); info.subtabs.forEach((s) => console.log(JSON.stringify(s)));
  console.log('\n--- tables on page ---'); info.tables.forEach((t) => console.log(JSON.stringify(t)));

  await ctx.close();
  process.exit(0);
})().catch((e) => { console.error('ERR', e); process.exit(1); });
