// discover_funded.js — probe the "Companies Funded" endpoint on NACL's page.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const L = require('./lib');

const HASH = '0x025C786EDFCF09A2AE99CE990D79E14A'; // Northern Arc Capital Limited
const SESSION_DIR = path.join(__dirname, '.session');

(async () => {
  const ctx = await chromium.launchPersistentContext(SESSION_DIR, { headless: true });
  try {
    const ss = JSON.parse(fs.readFileSync(path.join(__dirname, 'storageState.json'), 'utf8'));
    await ctx.addCookies((ss.cookies || []).filter((c) => /saverisk/.test(c.domain || '')));
  } catch {}
  const page = ctx.pages()[0] || (await ctx.newPage());
  await page.goto(L.BASE + '/myorders.aspx', { waitUntil: 'domcontentloaded', timeout: 60000 });
  const userid = await page.evaluate(() => { const e = document.querySelector('[id*="hdn_userid"],[id*="hdn_username"]'); return e ? e.value : ''; });
  if (!userid) { console.error('*** Not logged in. Run: node login.js ***'); await ctx.close(); process.exit(2); }
  console.log('Session OK. userid:', userid);

  // try several dashboardurl + parameter variants
  const candidates = [
    { dashboardurl: 'Companies Funded', parameterNvals: '' },
    { dashboardurl: 'companies funded', parameterNvals: '' },
    { dashboardurl: 'Companies Funded', parameterNvals: '{ddl}|;{period}|;{unit}|;{ddltemplate}|;{ddlsort}|;{ddlto}|;{IsSort}|' },
    { dashboardurl: 'Company Funded', parameterNvals: '' },
    { dashboardurl: 'Funded Companies', parameterNvals: '' },
  ];

  for (const cand of candidates) {
    const res = await page.evaluate(async ({ hash, cand }) => {
      try {
        const r = await fetch('https://www.saverisk.com/CmpAsyncDataService.aspx/ExecuteMethodStaticAsync', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json; charset=UTF-8', 'X-Requested-With': 'XMLHttpRequest' },
          body: JSON.stringify({ cinno: hash, parameterNvals: cand.parameterNvals, pgno: '1', dashboardurl: cand.dashboardurl, category: 'Company' }),
        });
        const t = await r.text();
        if (!r.ok) return { error: 'http ' + r.status, raw: t.slice(0, 300) };
        let inner;
        try { inner = JSON.parse(JSON.parse(t).d.Result || '{}'); } catch (e) { return { error: 'parse', raw: t.slice(0, 300) }; }
        const tbl = inner.Table || inner.table || [];
        return { ok: true, keys: Object.keys(inner), rowCount: tbl.length, firstRow: tbl[0] || null };
      } catch (e) { return { error: String(e) }; }
    }, { hash: HASH, cand });
    console.log('\n=== dashboardurl=' + JSON.stringify(cand.dashboardurl) + ' params=' + JSON.stringify(cand.parameterNvals) + ' ===');
    console.log(JSON.stringify(res, null, 2));
    if (res.ok && res.rowCount) break;
  }

  await ctx.close();
  process.exit(0);
})().catch((e) => { console.error('FATAL', e); process.exit(1); });
