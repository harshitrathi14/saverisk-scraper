// renewal_radar/lifecycle.js — per-entity funding LIFECYCLE signals for the Renewal Radar.
//
// Reads the two read-only caches (open charges + 5-year history) and input.csv metadata, and emits
// one record per entity (keyed off the history cache, the forecasting superset) carrying:
//   - openCharges    : current open positions (real lender names) with NACL/trustee flags
//   - monthly        : created/satisfied flow per active month over the last 60 months
//   - lenderCadence  : per-NAMED-lender creation cadence (count, gaps, regularity) — the renewal signal
//   - summary        : open totals + last-12m flow + distinct active lenders
//
// Pure read; never mutates the inputs. Date parsing goes through L.parseDate; bad years are dropped.
const fs = require('fs');
const path = require('path');
const L = require('../lib.js');
const RT = require('./ratings.js');

const ROOT = path.join(__dirname, '..');
const NOW = new Date();
const NOW_MS = NOW.getTime();
const MIN_YEAR = 1990, MAX_YEAR = NOW.getUTCFullYear() + 1;
const MS_60M = 60, MS_12M = 12; // month windows
const DAYS_PER_MONTH = 30.44;

// ---------- helpers ----------
// Parse a "DD Mon YYYY" date and reject typo years (e.g. "08 Jun 3198").
function validDate(s) {
  const d = L.parseDate(s);
  if (!d) return null;
  const y = d.getUTCFullYear();
  return (y < MIN_YEAR || y > MAX_YEAR) ? null : d;
}
const iso = (d) => d.toISOString();
const num = (v) => (v == null || v === '' ? null : Number(v));      // amountCr -> number|null
const amt0 = (v) => { const n = num(v); return n == null || Number.isNaN(n) ? 0 : n; }; // for sums
const ym = (d) => d.toISOString().slice(0, 7);                       // "YYYY-MM"
// whole-month difference between two ISO instants (UTC), used for the 60/12-month windows
function monthsBetween(a, b) {
  return (b.getUTCFullYear() - a.getUTCFullYear()) * 12 + (b.getUTCMonth() - a.getUTCMonth());
}
function median(arr) {
  if (!arr.length) return null;
  const s = [...arr].sort((x, y) => x - y), m = s.length >> 1;
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}
const clamp = (x, lo, hi) => Math.max(lo, Math.min(hi, x));

// A history event is a "creation" if its eventType contains "Creat" (covers Creation,
// Creation(Debenture), Creation modification(Other than Debenture)).
const isCreate = (t) => /creat/i.test(t || '');
// Repayment events: "Satisfaction" / "Satisfied".
const isSatisfy = (t) => /satisf/i.test(t || '');
// A lender name is usable for cadence only if NAMED: not blank, not the bucketed literal "Others",
// and not a debenture trustee (those hide the true lender).
function isNamedLender(name) {
  const n = String(name || '').trim();
  if (!n || /^others$/i.test(n)) return false;
  if (RT.isTrustee(n)) return false;
  return true;
}

// ---------- input.csv (entity metadata) ----------
// minimal CSV reader (handles quoted fields with embedded commas)
function readCsv(file) {
  const text = fs.readFileSync(file, 'utf8');
  const lines = text.split(/\r?\n/).filter((l) => l.length);
  const split = (line) => {
    const out = []; let cur = '', q = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (q) {
        if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') q = false;
        else cur += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',') { out.push(cur); cur = ''; }
      else cur += ch;
    }
    out.push(cur);
    return out;
  };
  const head = split(lines[0]);
  return lines.slice(1).map((l) => {
    const cells = split(l), row = {};
    head.forEach((h, i) => { row[h] = cells[i] == null ? '' : cells[i]; });
    return row;
  });
}

function loadInputMeta() {
  const meta = {};
  try {
    for (const r of readCsv(path.join(ROOT, 'input.csv'))) {
      const key = String(r.name || '').trim().toUpperCase();
      if (key) meta[key] = r;
    }
  } catch {}
  return meta;
}

// ---------- per-lender cadence ----------
// Accumulate creation events into per-lender buckets keyed by normalized name. Tracks distinct-dated
// creations (date+amount), display-name casing votes, and which source(s) it appeared in.
function addCreation(buckets, rawLender, date, amountCr, source) {
  const key = L.normName(rawLender) || rawLender.trim().toLowerCase();
  let b = buckets.get(key);
  if (!b) { b = { display: {}, dates: new Map(), sources: new Set() }; buckets.set(key, b); }
  b.display[rawLender] = (b.display[rawLender] || 0) + 1;   // casing vote
  b.sources.add(source);
  const dkey = iso(date);
  // dedup by lender+ISO-date (folding open into history); keep the amount we first see for that date
  if (!b.dates.has(dkey)) b.dates.set(dkey, { d: date, amt: num(amountCr) });
}

// Build lenderCadence rows from accumulated buckets (only lenders with >=2 distinct-dated creations).
function buildCadence(buckets) {
  const rows = [];
  for (const b of buckets.values()) {
    const entries = [...b.dates.values()].sort((a, c) => a.d - c.d);
    if (entries.length < 2) continue;
    const dates = entries.map((e) => e.d);
    // gaps between consecutive creations, in months (days / 30.44)
    const gaps = [];
    for (let i = 1; i < dates.length; i++) gaps.push((dates[i] - dates[i - 1]) / 86400000 / DAYS_PER_MONTH);
    let regularity;
    if (gaps.length === 1) {
      regularity = 0.5; // single gap -> not enough to judge regularity
    } else {
      const mean = gaps.reduce((s, x) => s + x, 0) / gaps.length;
      const variance = gaps.reduce((s, x) => s + (x - mean) ** 2, 0) / gaps.length;
      const cv = mean > 0 ? Math.sqrt(variance) / mean : 0; // coefficient of variation
      regularity = clamp(1 - cv, 0, 1);
    }
    // display name = most-common original casing
    const display = Object.entries(b.display).sort((a, c) => c[1] - a[1])[0][0];
    const sources = b.sources;
    const source = sources.has('open') && sources.has('history') ? 'both' : (sources.has('open') ? 'open' : 'history');
    rows.push({
      lender: display,
      isNACL: RT.isNACL(display),
      n: entries.length,
      firstDate: iso(dates[0]),
      lastDate: iso(dates[dates.length - 1]),
      medianGapMonths: median(gaps),
      regularity,
      lastAmtCr: entries[entries.length - 1].amt,
      source,
    });
  }
  rows.sort((a, b) => b.n - a.n);
  return rows;
}

// ---------- main ----------
function buildLifecycle() {
  const history = L.loadJson(path.join(ROOT, 'charge_history_cache.json'), {});
  const open = L.loadJson(path.join(ROOT, 'charges_cache.json'), {});
  const meta = loadInputMeta();

  const entities = [];
  // Iterate the history cache (superset of forecastable entities); enrich from open + input.csv.
  for (const key of Object.keys(history)) {
    const hist = history[key] || {};
    const op = open[key] || {};
    const m = meta[key] || {};

    // --- openCharges (current positions, real lender names) ---
    const openCharges = [];
    for (const c of (op.charges || [])) {
      const d = validDate(c.creationDate);
      if (!d) continue;
      openCharges.push({
        lender: c.lender,
        amountCr: num(c.amountCr),
        creationDate: iso(d),
        isNACL: RT.isNACL(c.lender),
        isTrustee: RT.isTrustee(c.lender),
      });
    }

    // --- monthly flow (history) over the last 60 months ---
    const monthsMap = new Map(); // ym -> bucket
    for (const e of (hist.events || [])) {
      const d = validDate(e.date);
      if (!d) continue;
      if (monthsBetween(d, NOW) > MS_60M || d > NOW) continue; // last 60 months only
      const created = isCreate(e.eventType), satisfied = isSatisfy(e.eventType);
      if (!created && !satisfied) continue;                    // skip plain Modification etc.
      const k = ym(d);
      let bk = monthsMap.get(k);
      if (!bk) { bk = { ym: k, createdCount: 0, createdCr: 0, satisfiedCount: 0, satisfiedCr: 0, netCr: 0 }; monthsMap.set(k, bk); }
      if (created) { bk.createdCount++; bk.createdCr += amt0(e.amountCr); }
      if (satisfied) { bk.satisfiedCount++; bk.satisfiedCr += amt0(e.amountCr); }
    }
    const monthly = [...monthsMap.values()].sort((a, b) => a.ym.localeCompare(b.ym));
    for (const bk of monthly) bk.netCr = bk.createdCr - bk.satisfiedCr;

    // --- lenderCadence (named-lender creation cadence; folds open creations into history) ---
    const buckets = new Map();
    for (const e of (hist.events || [])) {
      if (!isCreate(e.eventType) || !isNamedLender(e.lender)) continue;
      const d = validDate(e.date);
      if (!d) continue;
      addCreation(buckets, e.lender, d, e.amountCr, 'history');
    }
    for (const c of (op.charges || [])) {
      if (!isNamedLender(c.lender)) continue;
      const d = validDate(c.creationDate);
      if (!d) continue;
      addCreation(buckets, c.lender, d, c.amountCr, 'open');
    }
    const lenderCadence = buildCadence(buckets);

    // --- summary ---
    let totalOpenCr = 0;
    const activeLenders = new Set();
    for (const oc of openCharges) {
      totalOpenCr += amt0(oc.amountCr);
      if (isNamedLender(oc.lender)) activeLenders.add(L.normName(oc.lender) || oc.lender.toLowerCase());
    }
    let last12CreatedCr = 0, last12SatisfiedCr = 0;
    for (const e of (hist.events || [])) {
      const d = validDate(e.date);
      if (!d || d > NOW || monthsBetween(d, NOW) >= MS_12M) continue; // last 12 months
      if (isCreate(e.eventType)) last12CreatedCr += amt0(e.amountCr);
      else if (isSatisfy(e.eventType)) last12SatisfiedCr += amt0(e.amountCr);
    }

    entities.push({
      entity: key,
      sector: m.sector || op.saverisk_sector || hist.sector || '',
      exposure: m.exposure != null && m.exposure !== '' ? m.exposure : (hist.exposure || ''),
      onboardedDate: m.onboarded_date || hist.onboarded_date || '',
      openCharges,
      monthly,
      lenderCadence,
      summary: {
        totalOpenCr,
        openCount: openCharges.length,
        last12CreatedCr,
        last12SatisfiedCr,
        distinctActiveLenders: activeLenders.size,
      },
    });
  }

  return { asOf: NOW.toISOString(), entities };
}

module.exports = { buildLifecycle };

// ---------- CLI ----------
if (require.main === module) {
  const res = buildLifecycle();
  const outPath = path.join(__dirname, '_lifecycle.json');
  fs.writeFileSync(outPath, JSON.stringify(res, null, 2));

  const totalOpen = res.entities.reduce((s, e) => s + e.openCharges.length, 0);
  const totalCadence = res.entities.reduce((s, e) => s + e.lenderCadence.length, 0);
  console.log('entities:', res.entities.length);
  console.log('total open charges:', totalOpen);
  console.log('total lenderCadence relationships:', totalCadence);
  console.log('wrote', outPath);

  console.log('\nsample top lenderCadence rows (3 entities):');
  res.entities.filter((e) => e.lenderCadence.length).slice(0, 3).forEach((e) => {
    console.log('\n  ' + e.entity);
    e.lenderCadence.slice(0, 3).forEach((r) => {
      console.log(`    ${r.lender} | n=${r.n} | medianGap=${r.medianGapMonths == null ? 'NA' : r.medianGapMonths.toFixed(1)}m | reg=${r.regularity.toFixed(2)} | src=${r.source}`);
    });
  });

  // Validation: MAS FINANCIAL SERVICES LIMITED should show HDFC Bank with n>=5, medianGap ~4-7m.
  const mas = res.entities.find((e) => /MAS FINANCIAL SERVICES/i.test(e.entity));
  if (mas) {
    const hdfc = mas.lenderCadence.find((r) => /hdfc bank/i.test(r.lender));
    console.log('\nMAS / HDFC cadence row:', hdfc ? JSON.stringify(hdfc) : 'NOT FOUND');
  }
}
