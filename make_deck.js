// make_deck.js — build ONE combined PowerPoint for the whole project:
//   Part 1: Charge-Creation Lending Intelligence (NACL vs Others)   [from ppt.js]
//   Part 2: Renewal Radar — predicted next-tranche leads            [from renewal_radar/ppt_radar.js]
// Saved into the project ROOT as Saverisk_Lending_Intelligence_Deck.pptx (no scraping; reads caches).
//
// Run:  node make_deck.js
const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');
const L = require('./lib');
const R = require('./reports');
const { addMainSlides } = require('./ppt');
const { buildLifecycle } = require('./renewal_radar/lifecycle');
const { buildForecast } = require('./renewal_radar/predict');
const { buildLeads } = require('./renewal_radar/leads');
const { addRadarSlides } = require('./renewal_radar/ppt_radar');

const NA_RE = /northern arc/i;
const NOW = process.env.RUN_NOW ? new Date(process.env.RUN_NOW) : new Date();
const CACHE_FILE = path.join(__dirname, 'charges_cache.json');
const HIST_FILE = path.join(__dirname, 'charge_history_cache.json');
const INPUT_CSV = path.join(__dirname, 'input.csv');
const STATE_FILE = path.join(__dirname, 'state.json');

// ---- minimal CSV + entity loader (mirrors scrape.js) ----
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
  return out;
}

(async () => {
  // ----- Part 1 data: assemble okEntities + allCharges from the cache (same as scrape.js report step) -----
  const entities = loadEntities();
  for (const e of entities) if (!e.sector || /^na$/i.test(e.sector)) e.sector = 'Zero Exposure';
  const cache = L.loadJson(CACHE_FILE, {});
  const state = L.loadJson(STATE_FILE, { lastRun: null, charges: {} });
  const prevState = JSON.parse(JSON.stringify(state));

  const okEntities = [], allCharges = [];
  for (const e of entities) {
    const rec = cache[e.name.toUpperCase()];
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
  if (!okEntities.length) { console.error('No cached entities found — run the scraper first.'); process.exit(2); }
  const A = R.buildAnalyses(okEntities, allCharges, NOW, prevState);
  const mainMeta = { entityCount: okEntities.length, generated: NOW.toUTCString() };
  console.log(`Part 1 (NACL vs Others): ${okEntities.length} entities, ${allCharges.length} charges`);

  // ----- Part 2 data: Renewal Radar -----
  const lifecycle = buildLifecycle();
  const forecast = buildForecast(lifecycle, 12);
  const leads = buildLeads(forecast);
  let events = 0, satis = 0;
  try { const h = require(HIST_FILE); for (const k in h) for (const ev of (h[k].events || [])) { events++; if (/satisf/i.test(ev.eventType || '')) satis++; } } catch {}
  const radarMeta = { entityCount: lifecycle.entities.length, eventCount: events, satisfactionCount: satis, generated: NOW.toUTCString() };
  console.log(`Part 2 (Renewal Radar): ${lifecycle.entities.length} entities, ${forecast.length} forecasts, ${leads.length} competible leads`);

  // ----- one combined deck -----
  const p = new PptxGenJS();
  p.defineLayout({ name: 'WIDE', width: 13.33, height: 7.5 });
  p.layout = 'WIDE';
  p.theme = { headFontFace: 'Segoe UI', bodyFontFace: 'Segoe UI' };
  const state2 = { pageNo: 0 };       // shared continuous page numbering
  addMainSlides(p, A, mainMeta, state2);
  addRadarSlides(p, lifecycle, forecast, leads, radarMeta, state2);

  const file = path.join(__dirname, 'Saverisk_Lending_Intelligence_Deck.pptx');
  await p.writeFile({ fileName: file });
  console.log(`\nCombined deck written to project root:\n  ${file}`);
})();
