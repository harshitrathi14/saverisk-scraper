// renewal_radar/predict.js — forecast each entity's NEXT loan/renewal/top-up so NACL can
// approach BEFORE the incumbent renews. Consumes lifecycle.js's per-entity funding cadence.
//
// Two signals, both per upcoming month inside the horizon:
//   - "recurring-cadence" (primary): a lender<->entity relationship with >=2 charges and a
//     sane median gap → project lastDate + medianGap forward. Overdue cycles roll to next >= NOW.
//   - "entity-cadence" (secondary): when an entity has <2 usable relationships, fall back to its
//     coarse inter-creation rhythm from monthly[] → emit "(unknown — entity rhythm)" at low conf.
const fs = require('fs');
const path = require('path');

const clamp = (x, lo, hi) => Math.max(lo, Math.min(hi, x));
const num = (x) => (typeof x === 'number' && isFinite(x) ? x : null);
// NACL exposure: raw value is in RUPEES → ₹ Crore (1e7). 0/blank/invalid → null.
const exposureCr = (raw) => { const x = Number(raw); return (isFinite(x) && x > 0) ? +(x / 1e7).toFixed(2) : null; };

// ---------- date math (month-granular, UTC-safe) ----------
function addMonths(d, months) {
  const dt = new Date(d.getTime());
  const day = dt.getUTCDate();
  dt.setUTCDate(1);
  dt.setUTCMonth(dt.getUTCMonth() + Math.round(months));
  // clamp day to end-of-month
  const last = new Date(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth() + 1, 0)).getUTCDate();
  dt.setUTCDate(Math.min(day, last));
  return dt;
}
const ymOf = (d) => d.toISOString().slice(0, 7);
const parseD = (s) => { const d = new Date(s); return isNaN(d) ? null : d; };

// ---------- load lifecycle (file preferred, function fallback) ----------
function loadLifecycle() {
  const fp = path.join(__dirname, '_lifecycle.json');
  try { return JSON.parse(fs.readFileSync(fp, 'utf8')); } catch {}
  try { return require('./lifecycle.js').buildLifecycle(); } catch (e) {
    throw new Error('no _lifecycle.json and lifecycle.js unavailable: ' + e.message);
  }
}

// ---------- core ----------
function buildForecast(lifecycle, horizonMonths = 12) {
  const NOW = new Date();
  const horizonEnd = addMonths(NOW, horizonMonths);
  const out = [];
  const seen = new Set(); // entity|lender|ym dedupe

  const entities = (lifecycle && Array.isArray(lifecycle.entities)) ? lifecycle.entities : [];

  for (const ent of entities) {
    if (!ent) continue;
    const entity = ent.entity || ent.name || '';
    const sector = ent.sector || '';
    const naclExposureCr = exposureCr(ent.exposure);
    const cadences = Array.isArray(ent.lenderCadence) ? ent.lenderCadence : [];
    let emitted = 0;

    // ----- PRIMARY: recurring-cadence -----
    for (const lc of cadences) {
      if (!lc) continue;
      const n = num(lc.n);
      const gap = num(lc.medianGapMonths);
      const lastD = parseD(lc.lastDate);
      if (n == null || n < 2 || gap == null || gap < 0.5 || gap > 36 || !lastD) continue;

      let predicted = addMonths(lastD, gap);
      let rolled = false;
      if (predicted < NOW) {
        rolled = true;
        let guard = 0;
        while (predicted < NOW && guard++ < 600) predicted = addMonths(predicted, gap);
      }
      if (predicted > horizonEnd) continue;

      const reg = clamp(num(lc.regularity) ?? 0, 0, 1);
      const overduePenalty = rolled ? 0.15 : 0;
      const confidence = clamp(0.25 + 0.5 * reg + 0.1 * Math.min(n, 5) / 5 - overduePenalty, 0.05, 0.95);

      const ym = ymOf(predicted);
      const lender = lc.lender || '(unknown lender)';
      const key = entity + '|' + lender + '|' + ym;
      if (seen.has(key)) continue;
      seen.add(key);

      out.push({
        ym, date: predicted.toISOString(), entity, sector, lender,
        isNACL: !!lc.isNACL, basis: 'recurring-cadence',
        expectedAmtCr: num(lc.lastAmtCr), naclExposureCr,
        confidence, nObs: n, medianGapMonths: gap, lastDate: lastD.toISOString(),
      });
      emitted++;
    }

    // ----- SECONDARY: entity-cadence (few/no recurring relationships) -----
    if (emitted < 2) {
      const monthly = Array.isArray(ent.monthly) ? ent.monthly : [];
      // active months = those with a creation, within last 24 months
      const cutoff = addMonths(NOW, -24);
      const cutoffYm = ymOf(cutoff);
      const active = monthly
        .filter((m) => m && num(m.createdCount) > 0 && m.ym && m.ym >= cutoffYm)
        .map((m) => m.ym)
        .sort();

      if (active.length >= 2) {
        // typical inter-creation gap (months) between consecutive active months
        const ymToIdx = (s) => { const [y, mo] = s.split('-').map(Number); return y * 12 + (mo - 1); };
        const gaps = [];
        for (let i = 1; i < active.length; i++) gaps.push(ymToIdx(active[i]) - ymToIdx(active[i - 1]));
        gaps.sort((a, b) => a - b);
        const medGap = gaps[Math.floor(gaps.length / 2)];

        if (medGap >= 1 && medGap <= 36) {
          const lastActive = active[active.length - 1];
          const [ly, lm] = lastActive.split('-').map(Number);
          let cursor = new Date(Date.UTC(ly, lm - 1, 15));
          // walk forward into the horizon
          let guard = 0;
          while (guard++ < 600) {
            cursor = addMonths(cursor, medGap);
            if (cursor < NOW) continue;
            if (cursor > horizonEnd) break;
            const ym = ymOf(cursor);
            const lender = '(unknown — entity rhythm)';
            const key = entity + '|' + lender + '|' + ym;
            if (seen.has(key)) continue;
            seen.add(key);
            // coarse signal: cap confidence at 0.3, scale gently by #observations
            const confidence = clamp(0.12 + 0.03 * Math.min(active.length, 6), 0.05, 0.3);
            out.push({
              ym, date: cursor.toISOString(), entity, sector, lender,
              isNACL: false, basis: 'entity-cadence',
              expectedAmtCr: null, naclExposureCr, confidence,
              nObs: active.length, medianGapMonths: medGap,
              lastDate: new Date(Date.UTC(ly, lm - 1, 15)).toISOString(),
            });
          }
        }
      }
    }
  }

  out.sort((a, b) => (a.ym < b.ym ? -1 : a.ym > b.ym ? 1 : b.confidence - a.confidence));
  return out;
}

// ---------- mock lifecycle for self-test ----------
function mockLifecycle() {
  const now = new Date();
  const monthsAgo = (m) => addMonths(now, -m).toISOString();
  return {
    asOf: now.toISOString(),
    entities: [
      { // clean HDFC ~5-month cadence, 6 creations, last ~4 months ago → near-future high conf
        entity: 'Acme Finance Pvt Ltd', sector: 'SBL', exposure: 250000000, onboardedDate: monthsAgo(40), // ₹25 Cr (raw rupees)
        openCharges: [{ lender: 'HDFC Bank Limited', amountCr: 25, creationDate: monthsAgo(4), isNACL: false, isTrustee: false }],
        monthly: [
          { ym: ymOf(addMonths(now, -24)), createdCount: 1, createdCr: 20, satisfiedCount: 0, satisfiedCr: 0, netCr: 20 },
          { ym: ymOf(addMonths(now, -19)), createdCount: 1, createdCr: 21, satisfiedCount: 0, satisfiedCr: 0, netCr: 21 },
          { ym: ymOf(addMonths(now, -14)), createdCount: 1, createdCr: 22, satisfiedCount: 0, satisfiedCr: 0, netCr: 22 },
          { ym: ymOf(addMonths(now, -9)), createdCount: 1, createdCr: 24, satisfiedCount: 0, satisfiedCr: 0, netCr: 24 },
          { ym: ymOf(addMonths(now, -4)), createdCount: 1, createdCr: 25, satisfiedCount: 0, satisfiedCr: 0, netCr: 25 },
        ],
        lenderCadence: [
          { lender: 'HDFC Bank Limited', isNACL: false, n: 6, firstDate: monthsAgo(27), lastDate: monthsAgo(4), medianGapMonths: 5, regularity: 0.85, lastAmtCr: 25, source: 'charges' },
          { lender: 'Northern Arc Capital Limited', isNACL: true, n: 2, firstDate: monthsAgo(14), lastDate: monthsAgo(8), medianGapMonths: 6, regularity: 0.4, lastAmtCr: 10, source: 'charges' },
        ],
        summary: { totalOpenCr: 35, openCount: 2, last12CreatedCr: 49, last12SatisfiedCr: 0, distinctActiveLenders: 2 },
      },
      { // overdue single lender — should roll forward, get overdue penalty
        entity: 'Beta Microfin Ltd', sector: 'MFI', exposure: 30, onboardedDate: monthsAgo(50),
        openCharges: [],
        monthly: [],
        lenderCadence: [
          { lender: 'Kotak Mahindra Bank Limited', isNACL: false, n: 4, firstDate: monthsAgo(40), lastDate: monthsAgo(14), medianGapMonths: 6, regularity: 0.6, lastAmtCr: 15, source: 'charges' },
        ],
        summary: { totalOpenCr: 15, openCount: 1, last12CreatedCr: 0, last12SatisfiedCr: 0, distinctActiveLenders: 1 },
      },
      { // no usable relationships → entity-cadence fallback from monthly rhythm (~4 months)
        entity: 'Gamma Housing Fin Ltd', sector: 'AHF', exposure: 0, onboardedDate: monthsAgo(30), // zero exposure → null
        openCharges: [],
        monthly: [
          { ym: ymOf(addMonths(now, -12)), createdCount: 1, createdCr: 8, satisfiedCount: 0, satisfiedCr: 0, netCr: 8 },
          { ym: ymOf(addMonths(now, -8)), createdCount: 1, createdCr: 9, satisfiedCount: 0, satisfiedCr: 0, netCr: 9 },
          { ym: ymOf(addMonths(now, -4)), createdCount: 1, createdCr: 10, satisfiedCount: 0, satisfiedCr: 0, netCr: 10 },
        ],
        lenderCadence: [
          { lender: 'Some One-Off NBFC', isNACL: false, n: 1, firstDate: monthsAgo(4), lastDate: monthsAgo(4), medianGapMonths: null, regularity: 0, lastAmtCr: 10, source: 'charges' },
        ],
        summary: { totalOpenCr: 0, openCount: 0, last12CreatedCr: 27, last12SatisfiedCr: 0, distinctActiveLenders: 1 },
      },
    ],
  };
}

// ---------- CLI ----------
function printSummary(rows, horizon = 12) {
  const NOW = new Date();
  console.log('total forecast rows:', rows.length);
  const byBasis = {};
  for (const r of rows) byBasis[r.basis] = (byBasis[r.basis] || 0) + 1;
  console.log('by basis:', JSON.stringify(byBasis));

  console.log('per upcoming month (next ' + horizon + '):');
  const counts = {};
  for (let i = 0; i < horizon; i++) counts[ymOf(addMonths(NOW, i))] = 0;
  for (const r of rows) if (r.ym in counts) counts[r.ym]++;
  for (const ym of Object.keys(counts)) console.log('  ' + ym + ': ' + counts[ym]);

  console.log('top 10 by confidence:');
  const top = [...rows].sort((a, b) => b.confidence - a.confidence).slice(0, 10);
  for (const r of top) {
    console.log('  ' + r.ym + '  ' + r.confidence.toFixed(2) + '  ' + (r.isNACL ? 'NACL ' : '     ') +
      r.entity + ' <- ' + r.lender + '  [' + r.basis + ', n=' + r.nObs +
      ', gap=' + r.medianGapMonths + ', amt=' + r.expectedAmtCr + ']');
  }
}

if (require.main === module) {
  const mock = process.argv.includes('--mock');
  if (mock) {
    const rows = buildForecast(mockLifecycle(), 12);
    console.log('=== MOCK forecast ===');
    for (const r of rows) console.log(JSON.stringify(r));
    console.log('---');
    printSummary(rows);
  } else {
    const lifecycle = loadLifecycle();
    const rows = buildForecast(lifecycle, 12);
    fs.writeFileSync(path.join(__dirname, '_forecast.json'), JSON.stringify(rows, null, 2));
    printSummary(rows);
    console.log('wrote', path.join(__dirname, '_forecast.json'));
  }
}

module.exports = { buildForecast, loadLifecycle, addMonths, ymOf };
