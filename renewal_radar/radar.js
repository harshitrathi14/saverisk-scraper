// renewal_radar/radar.js — orchestrator for the Renewal Radar (separate from the weekly pipeline).
//
// Pipeline:  lifecycle (open + history caches) -> forecast (next 12m) -> leads (NACL-competible)
//            -> Excel + self-contained HTML dashboard + CSVs, written to output_radar/<timestamp>/.
//
// Inputs (repo root, produced by the existing scraper + harvest_history.js):
//   charges_cache.json, charge_history_cache.json, input.csv
// Run:  node renewal_radar/radar.js [--horizon 12]
const fs = require('fs');
const path = require('path');
const { buildLifecycle } = require('./lifecycle.js');
const { buildForecast } = require('./predict.js');
const { buildLeads } = require('./leads.js');
const { writeRadar } = require('./report_radar.js');
const { buildEmail } = require('./email.js');
const RT = require('./ratings.js');

const arg = (k, d) => { const i = process.argv.indexOf('--' + k); return i >= 0 && process.argv[i + 1] ? process.argv[i + 1] : d; };
const HORIZON = Number(arg('horizon', 12)) || 12;

(async () => {
  const t0 = Date.now();
  console.log('Renewal Radar — building from caches…');

  const lifecycle = buildLifecycle();
  console.log(`  lifecycle: ${lifecycle.entities.length} entities, ` +
    `${lifecycle.entities.reduce((s, e) => s + e.lenderCadence.length, 0)} cadence relationships`);

  const forecast = buildForecast(lifecycle, HORIZON);
  const byBasis = forecast.reduce((m, r) => ((m[r.basis] = (m[r.basis] || 0) + 1), m), {});
  console.log(`  forecast: ${forecast.length} rows (next ${HORIZON}m) — ${JSON.stringify(byBasis)}`);

  const leads = buildLeads(forecast);
  const leadCr = leads.reduce((s, l) => s + (Number(l.expectedAmtCr) || 0), 0);
  const tierMix = leads.reduce((m, l) => ((m[l.lenderTier] = (m[l.lenderTier] || 0) + 1), m), {});
  console.log(`  leads (NACL-competible): ${leads.length} | expected ₹${leadCr.toFixed(0)}Cr | tiers ${JSON.stringify(tierMix)}`);
  if (RT.OVERRIDES_COUNT === 0) {
    console.log('  NOTE: lender_ratings.json not populated — UNKNOWN-tier lenders are included as candidates.');
    console.log('        Drop a {"<lender name lowercased>":"<tier>"} map at renewal_radar/lender_ratings.json to refine.');
  }

  const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const outDir = path.join(__dirname, '..', 'output_radar', ts);
  fs.mkdirSync(outDir, { recursive: true });
  writeRadar(lifecycle, forecast, leads, outDir);

  // history stats (for the email header) + email draft -> outDir AND a copy in the project root
  let events = 0, satis = 0;
  try { const h = require(path.join(__dirname, '..', 'charge_history_cache.json')); for (const k in h) for (const ev of (h[k].events || [])) { events++; if (/satisf/i.test(ev.eventType || '')) satis++; } } catch {}
  const email = buildEmail(lifecycle, forecast, leads, { entityCount: lifecycle.entities.length, eventCount: events, satisfactionCount: satis });
  fs.writeFileSync(path.join(outDir, 'email_draft.md'), email);
  fs.writeFileSync(path.join(__dirname, '..', 'Renewal_Radar_email.md'), email);

  console.log(`\nDone in ${((Date.now() - t0) / 1000).toFixed(1)}s`);
  console.log(`Outputs: ${outDir}`);
  console.log('  - renewal_radar.html   (open this)');
  console.log('  - Renewal_Radar.xlsx');
  console.log('  - forecast_calendar.csv / leads.csv / entity_cadence.csv / monthly_flow.csv');
  console.log('  - email_draft.md   (also copied to project root: Renewal_Radar_email.md)');
})();
