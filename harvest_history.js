// harvest_history.js — pull the FULL "Charge History" (Creation/Modification/Satisfaction events,
// incl. closed/satisfied charges) for every onboarded entity, last 5 years by default.
// Reuses the saved login session and the hashes already resolved in approved_matches.json
// (so NO matching/search step). Resumable + incremental-save; leaves the weekly pipeline untouched.
//
//   node harvest_history.js                 # harvest all entities older than --max-age (default: only un-harvested)
//   node harvest_history.js --limit 5       # smoke test: first 5 entities only
//   node harvest_history.js --years 5       # lookback window (default 5)
//   node harvest_history.js --max-age 6     # re-harvest cache older than 6 days (weekly refresh)
//   node harvest_history.js --refresh       # ignore cache, re-harvest everyone
//   node harvest_history.js --fast          # shorter delays

const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');
const L = require('./lib.js');

const ROOT = __dirname;
const SESSION_DIR = path.join(ROOT, '.session');
const CACHE_FILE = path.join(ROOT, 'charge_history_cache.json');
const APPROVED_FILE = path.join(ROOT, 'approved_matches.json');
const INPUT = path.join(ROOT, 'input.csv');

const arg = (k, d) => { const i = process.argv.indexOf('--' + k); return i >= 0 ? (process.argv[i + 1] && !process.argv[i + 1].startsWith('--') ? process.argv[i + 1] : true) : d; };
const LIMIT = Number(arg('limit', 0)) || 0;
const YEARS = Number(arg('years', 5)) || 5;
const MAX_AGE_DAYS = arg('max-age', undefined) !== undefined ? Number(arg('max-age', 6)) : (arg('refresh', false) ? 0 : Infinity);
const REFRESH = !!arg('refresh', false);
const FAST = !!arg('fast', false);
const GENTLE = FAST ? { min: 400, max: 900 } : { min: 900, max: 1900 };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const rnd = (a, b) => Math.floor(a + Math.random() * (b - a));

const NOW = new Date();
const SINCE = new Date(NOW.getFullYear() - YEARS, NOW.getMonth(), NOW.getDate()).getTime();

// minimal CSV parser (handles quoted fields)
function parseCsv(txt) {
  const rows = []; let i = 0, field = '', row = [], q = false;
  const pushF = () => { row.push(field); field = ''; };
  const pushR = () => { if (row.length) rows.push(row); row = []; };
  while (i < txt.length) {
    const c = txt[i];
    if (q) { if (c === '"') { if (txt[i + 1] === '"') { field += '"'; i++; } else q = false; } else field += c; }
    else if (c === '"') q = true;
    else if (c === ',') pushF();
    else if (c === '\n') { pushF(); pushR(); }
    else if (c === '\r') {} else field += c;
    i++;
  }
  if (field.length || row.length) { pushF(); pushR(); }
  const head = rows.shift();
  return rows.map((r) => Object.fromEntries(head.map((h, j) => [h, r[j] ?? ''])));
}

(async () => {
  const approved = L.loadJson(APPROVED_FILE, {});
  const cache = L.loadJson(CACHE_FILE, {});
  const inputRows = parseCsv(fs.readFileSync(INPUT, 'utf8'));

  // build work list: input entity -> resolved hash from approved_matches
  let work = [];
  for (const r of inputRows) {
    const key = (r.name || '').toUpperCase();
    const m = approved[key];
    if (!m || !m.hash) continue;
    work.push({ key, name: r.name, short_name: r.short_name, cin: r.cin, sector: r.sector, exposure: r.exposure, onboarded_date: r.onboarded_date, hash: m.hash });
  }
  const totalMatched = work.length;
  // skip already-harvested (fresh) entities unless refreshing
  work = work.filter((w) => {
    const rec = cache[w.key];
    if (!rec) return true;
    if (REFRESH) return true;
    if (MAX_AGE_DAYS === Infinity) return false; // default: only un-harvested
    return !rec.harvestedAt || (NOW - new Date(rec.harvestedAt)) > MAX_AGE_DAYS * 864e5;
  });
  if (LIMIT) work = work.slice(0, LIMIT);

  console.log(`Charge-History harvest | lookback ${YEARS}y (since ${new Date(SINCE).toISOString().slice(0, 10)})`);
  console.log(`Matched entities: ${totalMatched} | already cached: ${totalMatched - (work.length)} (approx) | to harvest now: ${work.length}\n`);
  if (!work.length) { console.log('Nothing to harvest. Use --refresh or --max-age N to re-pull.'); process.exit(0); }

  const ctx = await chromium.launchPersistentContext(SESSION_DIR, { headless: true });
  try { const ss = JSON.parse(fs.readFileSync(path.join(ROOT, 'storageState.json'), 'utf8')); await ctx.addCookies((ss.cookies || []).filter((c) => /saverisk/.test(c.domain || ''))); } catch {}
  const page = ctx.pages()[0] || (await ctx.newPage());
  await page.goto(L.BASE + '/myorders.aspx', { waitUntil: 'domcontentloaded', timeout: 60000 });
  let userid = await page.evaluate(() => { const e = document.querySelector('[id*="hdn_userid"],[id*="hdn_username"]'); return e ? e.value : ''; });
  if (!userid) { console.error('\n*** Not logged in. Run `node login.js` first. ***'); await ctx.close(); process.exit(2); }
  console.log('Session OK. user', userid, '\n');

  async function refresh() {
    await page.goto(L.BASE + '/myorders.aspx', { waitUntil: 'domcontentloaded', timeout: 60000 }).catch(() => {});
    await sleep(1200);
    return page.evaluate(() => { const e = document.querySelector('[id*="hdn_userid"],[id*="hdn_username"]'); return e ? e.value : ''; });
  }

  let done = 0, consecErr = 0, totalEvents = 0, totalSatis = 0;
  for (const w of work) {
    let res = await L.fetchChargeHistory(page, w.hash, SINCE);
    if (res.error) {
      consecErr++;
      if (consecErr >= 4) { const uid = await refresh(); if (!uid) { console.error(`\n*** SESSION EXPIRED after ${done}. Re-login + re-run to resume (cache kept). ***`); break; } userid = uid; consecErr = 0; res = await L.fetchChargeHistory(page, w.hash, SINCE); }
      if (res.error) { console.log(`  ! ${w.short_name || w.name}: ${res.error}`); continue; }
    }
    consecErr = 0;
    const ev = res.events || [];
    const satis = ev.filter((e) => /satisf/i.test(e.eventType)).length;
    const creat = ev.filter((e) => /creat/i.test(e.eventType)).length;
    totalEvents += ev.length; totalSatis += satis;
    cache[w.key] = {
      matched_company: (approved[w.key] || {}).name || w.name, cin: w.cin, hash: w.hash,
      sector: w.sector, exposure: w.exposure, onboarded_date: w.onboarded_date,
      events: ev, eventCount: ev.length, fetched: res.fetched, harvestedAt: NOW.toISOString(),
    };
    L.saveJson(CACHE_FILE, cache); // incremental, crash-safe
    done++;
    console.log(`  [${done}/${work.length}] ${(w.short_name || w.name).slice(0, 34).padEnd(34)} events:${String(ev.length).padStart(4)}  creat:${String(creat).padStart(4)}  satis:${String(satis).padStart(4)}`);
    await sleep(rnd(GENTLE.min, GENTLE.max));
  }

  console.log(`\n================= HARVEST SUMMARY =================`);
  console.log(`Harvested this run: ${done}/${work.length} | total events: ${totalEvents} | satisfaction events: ${totalSatis}`);
  console.log(`Cache: ${CACHE_FILE} (${Object.keys(cache).length} entities)`);
  await ctx.close();
  process.exit(0);
})();
