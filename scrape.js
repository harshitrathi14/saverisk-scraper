// scrape.js — Saverisk lending-intelligence scraper (Northern Arc vs Others).
// Gentle, resumable, cache-backed. Running this NEVER uses AI tokens.
//
// Run:  node scrape.js                 process all (uses cache; only fetches new/uncached)
//       node scrape.js --max-new 150    fetch at most 150 NEW entities this run, then stop (batching)
//       node scrape.js --limit 15       only consider first 15 input rows (testing)
//       node scrape.js --refresh        ignore cache and re-fetch everything
//       node scrape.js --fast           shorter delays (less "gentle")
//
// Outputs -> ./output/<timestamp>/ : Excel, charge_creation_dashboard.html, deck.pptx, CSVs.

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const L = require('./lib');
const R = require('./reports');
const OUT = require('./output');
const PPT = require('./ppt');

const SESSION_DIR = path.join(__dirname, '.session');
const STATE_FILE = path.join(__dirname, 'state.json');
const APPROVED_FILE = path.join(__dirname, 'approved_matches.json');
const CACHE_FILE = path.join(__dirname, 'charges_cache.json');
const INPUT_CSV = path.join(__dirname, 'input.csv');
const OUT_ROOT = path.join(__dirname, 'output');
const NA_RE = /northern arc/i;

const args = process.argv.slice(2);
const argVal = (f) => { const i = args.indexOf(f); return i >= 0 ? args[i + 1] : null; };
const LIMIT = argVal('--limit') ? parseInt(argVal('--limit'), 10) : null;
const MAX_NEW = argVal('--max-new') ? parseInt(argVal('--max-new'), 10) : Infinity;
const REFRESH = args.includes('--refresh');
const MAX_AGE_DAYS = argVal('--max-age') ? parseFloat(argVal('--max-age')) : Infinity; // re-fetch cache older than N days
const FAST = args.includes('--fast');
const RATING = args.includes('--rating');
const NO_OVERVIEW = args.includes('--no-overview');
const REPORT_ONLY = args.includes('--report-only'); // rebuild outputs from cache, no browser/session
const NOW = process.env.RUN_NOW ? new Date(process.env.RUN_NOW) : new Date();

// gentle pacing: random delay between fetched entities, longer pause every batch
const GENTLE = FAST
  ? { min: 300, max: 700, batchEvery: 80, batchPause: [3000, 6000] }
  : { min: 1300, max: 3000, batchEvery: 40, batchPause: [20000, 40000] };
const rnd = (a, b) => Math.floor(a + Math.random() * (b - a));
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ---------- CSV ----------
function parseCsv(text) {
  const rows = []; let row = [], f = '', q = false;
  for (let i = 0; i < text.length; i++) { const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else if (c === '"') q = true; else if (c === ',') { row.push(f); f = ''; }
    else if (c === '\n') { row.push(f); rows.push(row); row = []; f = ''; }
    else if (c !== '\r') f += c; }
  if (f.length || row.length) { row.push(f); rows.push(row); }
  return rows.filter((r) => r.some((c) => c.trim() !== ''));
}
function loadEntities() {
  const rows = parseCsv(fs.readFileSync(INPUT_CSV, 'utf8'));
  const hdr = rows[0].map((h) => h.trim().toLowerCase());
  const col = (...n) => { for (const x of n) { const i = hdr.indexOf(x); if (i >= 0) return i; } return -1; };
  const iN = col('name', 'entity'), iS = col('short_name', 'short name'), iC = col('cin'),
        iE = col('exposure'), iSec = col('sector'), iO = col('onboarded_date', 'onboardeddate');
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]; const name = (r[iN] || '').trim(); if (!name) continue;
    out.push({ name, short_name: iS >= 0 ? (r[iS] || '').trim() : '', cin: iC >= 0 ? (r[iC] || '').trim() : '',
      exposure: iE >= 0 ? r[iE] : '', sector: (iSec >= 0 ? (r[iSec] || '').trim() : '') || '', onboarded: iO >= 0 ? r[iO] : '' });
  }
  return LIMIT ? out.slice(0, LIMIT) : out;
}

(async () => {
  const entities = loadEntities();
  for (const e of entities) if (!e.sector || /^na$/i.test(e.sector)) e.sector = 'Zero Exposure';

  const approved = L.loadJson(APPROVED_FILE, {});
  const state = L.loadJson(STATE_FILE, { lastRun: null, charges: {} });
  const prevState = JSON.parse(JSON.stringify(state));
  const cache = L.loadJson(CACHE_FILE, {});      // KEY -> { matched_company, cin, hash, saverisk_sector, saverisk_industry, charges:[...] }

  const cached0 = Object.keys(cache).length;
  console.log(`Loaded ${entities.length} entities. Cached already: ${cached0}. Mode: ${FAST ? 'fast' : 'gentle'}, max-new: ${MAX_NEW}`);

  let ctx = null, page = null, userid = '';
  const unmatched = [];
  let newFetches = 0, consecErr = 0, aborted = false, reachedCap = false;

  if (!REPORT_ONLY) {
  ctx = await chromium.launchPersistentContext(SESSION_DIR, { headless: true });
  // The auth cookie (api_session_xx) is a session cookie that the persistent profile drops on close,
  // so re-inject the snapshot saved by login.js at the moment of login.
  try {
    const ss = JSON.parse(fs.readFileSync(path.join(__dirname, 'storageState.json'), 'utf8'));
    await ctx.addCookies((ss.cookies || []).filter((c) => /saverisk/.test(c.domain || '')));
  } catch {}
  page = ctx.pages()[0] || (await ctx.newPage());
  await page.goto(L.BASE + '/myorders.aspx', { waitUntil: 'domcontentloaded', timeout: 60000 });
  userid = await page.evaluate(() => { const e = document.querySelector('[id*="hdn_userid"],[id*="hdn_username"]'); return e ? e.value : ''; });
  if (!userid) { console.error('\n*** Not logged in. Run the login step first. ***'); await ctx.close(); process.exit(2); }
  console.log('Session OK. user id:', userid);

  async function refreshSession() {
    await page.goto(L.BASE + '/myorders.aspx', { waitUntil: 'domcontentloaded', timeout: 60000 }).catch(() => {});
    await sleep(1200);
    return page.evaluate(() => { const el = document.querySelector('[id*="hdn_userid"],[id*="hdn_username"]'); return el ? el.value : ''; });
  }

  for (const e of entities) {
    const key = e.name.toUpperCase();

    // use cache unless refreshing or the cached snapshot is older than --max-age days
    const rec0 = cache[key];
    const fresh = rec0 && (MAX_AGE_DAYS === Infinity || (rec0.scrapedAt && (NOW - new Date(rec0.scrapedAt)) <= MAX_AGE_DAYS * 864e5));
    if (!REFRESH && fresh) { e._cached = true; continue; }
    if (aborted || reachedCap) { e._skipped = true; continue; }
    if (newFetches >= MAX_NEW) { reachedCap = true; e._skipped = true; continue; }

    // ---- match (CIN-first) ----
    let match = approved[key] && approved[key].hash ? approved[key] : null;
    if (!match) {
      const cin = (e.cin || '').trim().toUpperCase();
      const q = cin || e.short_name || e.name;
      let sr = await L.apiSearch(page, q, userid);
      if (sr.error || !sr.results || !sr.results.length) {
        await sleep(1500); const uid = await refreshSession(); if (uid) userid = uid;
        sr = await L.apiSearch(page, q, userid);
      }
      if (sr.error) {
        consecErr++; unmatched.push({ entity: e.name, cin, reason: sr.error });
        if (consecErr >= 6) { const uid = await refreshSession(); if (!uid) { aborted = true; console.error(`\n*** SESSION EXPIRED after ${newFetches} new entities. Re-login and re-run to resume (cache keeps progress). ***`); } else { userid = uid; consecErr = 0; } }
        continue;
      }
      consecErr = 0;
      const results = sr.results || [];
      let hit = cin ? results.find((r) => (r.CIN || '').toUpperCase() === cin) : null;
      if (!hit && results.length) hit = results[0];
      if (!hit) { unmatched.push({ entity: e.name, cin, reason: cin ? 'CIN not found' : 'no results' }); continue; }
      match = { name: L.stripHtml(hit.Company_Name).replace(/-\s*company$/i, '').replace(/^invalid cin of\s*/i, '').trim(), cin: hit.CIN || cin, hash: (hit.HIDE_HC || '').replace(/^company\//, ''), source: 'cin' };
      approved[key] = match;
    }

    // ---- charges ----
    const cr = await L.fetchAllCharges(page, match.hash);
    if (cr.error) { unmatched.push({ entity: e.name, cin: match.cin, reason: 'charge ' + cr.error }); continue; }
    consecErr = 0;
    const rec = { matched_company: match.name, cin: match.cin, hash: match.hash, charges: cr.charges, scrapedAt: NOW.toISOString() };
    if (!NO_OVERVIEW) { try { const ov = await L.fetchOverview(page, match.hash); if (ov) { rec.saverisk_sector = ov.sector; rec.saverisk_industry = ov.industry; } } catch {} }
    if (RATING) { try { rec.rating = await L.fetchRatingUI(page, match.hash); } catch {} }
    cache[key] = rec;
    L.saveJson(CACHE_FILE, cache);     // persist incrementally (safe on crash)
    L.saveJson(APPROVED_FILE, approved);

    newFetches++;
    if (newFetches % 25 === 0) console.log(`  …fetched ${newFetches} new (cache now ${Object.keys(cache).length})`);
    // gentle pacing
    await sleep(rnd(GENTLE.min, GENTLE.max));
    if (newFetches % GENTLE.batchEvery === 0) { const p = rnd(GENTLE.batchPause[0], GENTLE.batchPause[1]); console.log(`  …batch pause ${Math.round(p / 1000)}s (gentle)`); await sleep(p); }
  }
  } // end if (!REPORT_ONLY)
  if (REPORT_ONLY) console.log(`Report-only: building from cache (${Object.keys(cache).length} entities), no scraping.`);

  // ---------- assemble dataset from cache for all input entities present ----------
  const okEntities = [];
  const allCharges = [];
  for (const e of entities) {
    const key = e.name.toUpperCase();
    const rec = cache[key];
    if (!rec) continue;
    e.status = 'ok'; e.hash = rec.hash; e.matched_company = rec.matched_company; e.matched_cin = rec.cin;
    e.saverisk_sector = rec.saverisk_sector || ''; e.saverisk_industry = rec.saverisk_industry || ''; e.rating = rec.rating || null;
    okEntities.push(e);
    for (const c of rec.charges) {
      allCharges.push({ entity: e.name, short_name: e.short_name, sector: e.sector, exposure: e.exposure,
        lender: c.lender, lenderCin: c.lenderCin, amountCr: c.amountCr, creationDate: c.creationDate,
        date: L.parseDate(c.creationDate), chargeId: c.chargeId, isNA: NA_RE.test(c.lender || ''), isTrustee: R.TRUSTEE_RE.test(c.lender || '') });
    }
  }

  const fullyDone = okEntities.length >= entities.length - unmatched.length && !aborted && !reachedCap;

  // ---------- analyses + outputs ----------
  const A = R.buildAnalyses(okEntities, allCharges, NOW, prevState);
  const ts = NOW.toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const outDir = path.join(OUT_ROOT, ts);
  fs.mkdirSync(outDir, { recursive: true });

  const toCsv = (rows, cols) => (cols.join(',') + '\n') + (rows || []).map((r) => cols.map((c) => {
    const v = r[c] == null ? '' : String(r[c]); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join(',')).join('\n') + '\n';
  const writeCsv = (n, rows) => { if (rows && rows.length) fs.writeFileSync(path.join(outDir, n), toCsv(rows, Object.keys(rows[0]))); };

  OUT.writeExcel(A, path.join(outDir, `Saverisk_Lending_Intelligence_${ts}.xlsx`));
  OUT.writeHtml(A, path.join(outDir, 'charge_creation_dashboard.html'), { entityCount: okEntities.length, generated: NOW.toUTCString() });
  try { await PPT.writePpt(A, path.join(outDir, `Saverisk_Deck_${ts}.pptx`), { entityCount: okEntities.length, generated: NOW.toUTCString() }); }
  catch (err) { console.error('PPT failed:', err.message); }
  writeCsv('entity_summary.csv', A.summaryRows);
  writeCsv('new_since_last_run.csv', A.newSinceRun);
  for (const w of R.WINDOWS) { writeCsv(`funded_by_others_${w.key}.csv`, A.externalCharges[w.key]); writeCsv(`funded_by_northern_arc_${w.key}.csv`, A.naCharges[w.key]); writeCsv(`first_time_funded_${w.key}.csv`, A.firstTimeFunded[w.key]); writeCsv(`active_lenders_${w.key}.csv`, A.activeLenders[w.key]); writeCsv(`first_time_lenders_${w.key}.csv`, A.firstTime[w.key]); writeCsv(`new_lenders_to_book_${w.key}.csv`, A.newLenders[w.key]); }
  writeCsv('na_vs_others.csv', A.naVsOthers);
  writeCsv('nacl_share_per_entity.csv', A.naclShare);        // NACL % share of each entity's charges
  OUT.writeSheetExcel(A.naclShare, 'NACL Share per Entity', path.join(outDir, `NACL_Share_per_Entity_${ts}.xlsx`));
  writeCsv('all_charges.csv', A.allCharges);                 // the full charge list (big CSV)
  if (unmatched.length) writeCsv('unmatched.csv', unmatched);
  // snapshot the raw cache into the run folder too (full per-entity scraped data)
  try { fs.copyFileSync(CACHE_FILE, path.join(outDir, 'charges_cache.json')); } catch {}

  // update baseline state only when the WHOLE input set is done in one complete pass
  for (const e of okEntities) state.charges[e.name.toUpperCase()] = (cache[e.name.toUpperCase()].charges || []).map((c) => c.chargeId);
  if (fullyDone) state.lastRun = NOW.toISOString();
  L.saveJson(STATE_FILE, state);

  console.log('\n================= RUN SUMMARY =================');
  if (aborted) console.log('*** ABORTED (session expired) — re-login + re-run to resume. ***');
  if (reachedCap) console.log(`*** Reached --max-new ${MAX_NEW} cap — re-run to fetch the next batch. ***`);
  console.log(`Input: ${entities.length} | scraped/cached: ${okEntities.length} | new this run: ${newFetches} | unmatched: ${unmatched.length}`);
  console.log(`Total charges: ${allCharges.length} | full baseline complete: ${fullyDone}`);
  for (const row of A.naVsOthers) console.log(`  ${row.window.padEnd(13)} | Others: ${row.other_charges} / ₹${row.other_amount_cr}Cr | NA: ${row.na_charges} / ₹${row.na_amount_cr}Cr`);
  console.log(`Funded by others (1m): ${A.externalCharges['1m'].length} charges | Funded by Northern Arc (1m): ${A.naCharges['1m'].length} charges`);
  console.log(`\nOutputs: ${outDir}`);

  if (ctx) await ctx.close();
  process.exit(0);
})().catch((e) => { console.error('FATAL', e); process.exit(1); });
