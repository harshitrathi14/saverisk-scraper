// explore.js — reuses the saved ./.session to inspect authenticated pages.
// Usage: node explore.js <url> <tag>
// Saves: explore_<tag>.png and explore_<tag>.html, prints links/forms/inputs.

const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const USER_DATA_DIR = path.join(__dirname, '.session');
const url = process.argv[2] || 'https://www.saverisk.com/myorders.aspx';
const tag = process.argv[3] || 'home';

(async () => {
  const ctx = await chromium.launchPersistentContext(USER_DATA_DIR, { headless: true });
  const page = ctx.pages()[0] || (await ctx.newPage());
  const resp = await page.goto(url, { waitUntil: 'networkidle', timeout: 60000 }).catch((e) => { console.log('goto err', e.message); return null; });
  await page.waitForTimeout(1500);

  console.log('FINAL URL:', page.url());
  console.log('STATUS:', resp && resp.status());
  console.log('TITLE:', await page.title());

  await page.screenshot({ path: `explore_${tag}.png`, fullPage: true }).catch(() => {});
  fs.writeFileSync(`explore_${tag}.html`, await page.content());

  const data = await page.evaluate(() => {
    const txt = (e) => (e.innerText || e.value || e.placeholder || '').trim().replace(/\s+/g, ' ').slice(0, 60);
    const links = [...document.querySelectorAll('a[href]')]
      .map((a) => ({ t: txt(a), href: a.getAttribute('href') }))
      .filter((l) => l.href && !l.href.startsWith('javascript:void') === false ? true : (l.t || l.href))
      .slice(0, 120);
    const inputs = [...document.querySelectorAll('input,select,textarea')]
      .map((e) => ({ tag: e.tagName, type: e.type, id: e.id, name: e.name, ph: e.placeholder, val: txt(e) }))
      .filter((e) => e.id || e.name || e.ph)
      .slice(0, 60);
    const buttons = [...document.querySelectorAll('button,[role=button],input[type=button],input[type=submit]')]
      .map((b) => ({ id: b.id, t: txt(b) })).filter((b) => b.id || b.t).slice(0, 50);
    return { links, inputs, buttons };
  });

  console.log('\n--- INPUTS ---'); data.inputs.forEach((i) => console.log(JSON.stringify(i)));
  console.log('\n--- BUTTONS ---'); data.buttons.forEach((b) => console.log(JSON.stringify(b)));
  console.log('\n--- LINKS (first 120) ---'); data.links.forEach((l) => console.log(`${l.t}  ->  ${l.href}`));

  await ctx.close();
  process.exit(0);
})().catch((e) => { console.error('ERR', e); process.exit(1); });
