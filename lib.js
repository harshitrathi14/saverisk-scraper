// lib.js — shared helpers for the Saverisk scraper (matching, parsing, API calls).
const fs = require('fs');
const path = require('path');

const BASE = 'https://www.saverisk.com';

// ---------- name normalization & similarity ----------
const SUFFIXES = [
  'private limited', 'pvt ltd', 'pvt. ltd.', 'pvt limited', 'p ltd', 'pvt',
  'public limited', 'limited', 'ltd', 'llp', 'incorporated', 'inc',
  'corporation', 'corp', 'company', 'co', 'india', '(india)', 'nbfc', '(nbfc)',
];

function normName(s) {
  if (!s) return '';
  let t = String(s).toLowerCase();
  t = t.replace(/-\s*company\s*$/i, ' ');         // strip trailing "- Company"
  t = t.replace(/[.,&/()\-]/g, ' ');               // punctuation -> space
  t = t.replace(/\s+/g, ' ').trim();
  // drop trailing legal suffix words (repeatedly)
  let changed = true;
  while (changed) {
    changed = false;
    for (const suf of SUFFIXES) {
      const re = new RegExp('(^|\\s)' + suf.replace(/[.()]/g, '\\$&') + '\\s*$');
      if (re.test(t)) { t = t.replace(re, '').trim(); changed = true; }
    }
  }
  return t.replace(/\s+/g, ' ').trim();
}

function tokens(s) { return new Set(normName(s).split(' ').filter(Boolean)); }

// Jaccard token overlap + prefix bonus; returns 0..1
function similarity(a, b) {
  const na = normName(a), nb = normName(b);
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  const ta = tokens(a), tb = tokens(b);
  if (!ta.size || !tb.size) return 0;
  let inter = 0; for (const x of ta) if (tb.has(x)) inter++;
  const jac = inter / (ta.size + tb.size - inter);
  // containment: all of the shorter name's tokens present in the longer
  const [small, big] = ta.size <= tb.size ? [ta, tb] : [tb, ta];
  let cont = 0; for (const x of small) if (big.has(x)) cont++;
  const containment = cont / small.size;
  // startsWith bonus (handles "alpha marine" vs "alpha marine foods")
  const pref = (nb.startsWith(na) || na.startsWith(nb)) ? 0.15 : 0;
  return Math.min(1, Math.max(jac, 0.7 * containment) + pref);
}

// ---------- date parsing ----------
const MON = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
function parseDate(s) {
  if (!s) return null;
  const m = String(s).trim().match(/(\d{1,2})\s+([A-Za-z]{3,})\s+(\d{4})/);
  if (!m) return null;
  const mon = MON[m[2].slice(0, 3).toLowerCase()];
  if (mon == null) return null;
  return new Date(Date.UTC(+m[3], mon, +m[1]));
}
function daysAgo(n, now) { const d = new Date(now); d.setUTCDate(d.getUTCDate() - n); return d; }
function fmtDate(d) {
  if (!d) return '';
  return d.toISOString().slice(0, 10);
}

// ---------- html strip ----------
function stripHtml(s) {
  if (s == null) return '';
  return String(s).replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
}

// ---------- Saverisk JSON APIs (run from a page already on saverisk origin) ----------
// page: a Playwright Page loaded on www.saverisk.com. Uses in-page fetch (cookies + CSRF).
async function apiSearch(page, searchstr, userid) {
  return page.evaluate(async ({ searchstr, userid }) => {
    const r = await fetch('https://www.saverisk.com/AsyncService.aspx/SearchAddCompany', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json; charset=UTF-8', 'X-Requested-With': 'XMLHttpRequest' },
      body: JSON.stringify({ searchstr, type: 'Company', id: userid }),
    });
    const t = await r.text();
    if (!r.ok) return { error: 'http ' + r.status, raw: t.slice(0, 200) };
    try {
      const outer = JSON.parse(t);
      const arr = JSON.parse(outer.d.Result || '[]');
      return { ok: true, results: arr };
    } catch (e) { return { error: 'parse', raw: t.slice(0, 200) }; }
  }, { searchstr, userid });
}

async function apiCharges(page, hash, pgno) {
  return page.evaluate(async ({ hash, pgno }) => {
    const r = await fetch('https://www.saverisk.com/CmpAsyncDataService.aspx/ExecuteMethodStaticAsync', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json; charset=UTF-8', 'X-Requested-With': 'XMLHttpRequest' },
      body: JSON.stringify({
        cinno: hash,
        parameterNvals: '{ddl}|;{period}|;{unit}|;{ddltemplate}|;{ddlsort}|;{ddlto}|;{IsSort}|',
        pgno: String(pgno),
        dashboardurl: 'Open Charge',
        category: 'Company',
      }),
    });
    const t = await r.text();
    if (!r.ok) return { error: 'http ' + r.status, raw: t.slice(0, 200) };
    try {
      const outer = JSON.parse(t);
      const inner = JSON.parse(outer.d.Result || '{}');
      return { ok: true, table: inner.Table || [] };
    } catch (e) { return { error: 'parse', raw: t.slice(0, 200) }; }
  }, { hash, pgno });
}

// latest credit rating via the UI (reliable but slower: navigates + clicks the Credit Rating tab).
// Returns { ratingDate, agency, grade, outlook } or null.
async function fetchRatingUI(page, hash) {
  try {
    await page.goto(BASE + '/company/' + hash + '/x', { waitUntil: 'domcontentloaded', timeout: 60000 });
    const menu = page.locator('[data*="credit rating~Credit Rating"]').first();
    if (!(await menu.count())) return null;
    await menu.click({ timeout: 8000 }).catch(() => {});
    await page.waitForTimeout(3000);
    const rows = await page.evaluate(() => {
      const heads = [...document.querySelectorAll('th')].map((th) => th.innerText.trim());
      const t = document.querySelector('#table_data');
      if (!t) return [];
      const idx = (re) => heads.findIndex((h) => re.test(h));
      const iD = idx(/rating date/i), iA = idx(/agency/i), iR = idx(/^rating$/i), iO = idx(/outlook/i);
      return [...t.querySelectorAll('tr')].map((tr) => {
        const td = [...tr.querySelectorAll('td')].map((x) => x.innerText.trim());
        if (!td.length) return null;
        return { ratingDate: iD >= 0 ? td[iD] : '', agency: iA >= 0 ? td[iA] : '', grade: iR >= 0 ? td[iR] : '', outlook: iO >= 0 ? td[iO] : '' };
      }).filter((r) => r && (r.grade || r.agency));
    });
    rows.sort((a, b) => (parseDate(b.ratingDate) || 0) - (parseDate(a.ratingDate) || 0));
    return rows[0] || null;
  } catch { return null; }
}

// company overview profile (Saverisk's own Sector + Industry, etc.)
async function fetchOverview(page, hash) {
  const res = await page.evaluate(async ({ hash }) => {
    const r = await fetch('https://www.saverisk.com/CmpAsyncDataService.aspx/ExecuteMethodStaticAsync', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json; charset=UTF-8', 'X-Requested-With': 'XMLHttpRequest' },
      body: JSON.stringify({ cinno: hash, parameterNvals: '', pgno: '', dashboardurl: 'overview', category: 'Company' }),
    });
    const t = await r.text();
    if (!r.ok) return { error: 'http ' + r.status };
    try { const inner = JSON.parse(JSON.parse(t).d.Result || '{}'); return { ok: true, row: (inner.Table && inner.Table[0]) || {} }; }
    catch (e) { return { error: 'parse' }; }
  }, { hash });
  if (res.error || !res.row) return null;
  return {
    sector: stripHtml(res.row.Sector || ''),
    industry: stripHtml(res.row.Industry || ''),
    state: stripHtml(res.row.State || res.row.Registered_State || ''),
    status: stripHtml(res.row.Company_Status || res.row.Status || ''),
  };
}

// latest credit rating for a company (dashboardurl "credit rating")
async function apiRating(page, hash) {
  return page.evaluate(async ({ hash }) => {
    const r = await fetch('https://www.saverisk.com/CmpAsyncDataService.aspx/ExecuteMethodStaticAsync', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json; charset=UTF-8', 'X-Requested-With': 'XMLHttpRequest' },
      body: JSON.stringify({ cinno: hash, parameterNvals: '{filter1}|;{filter2}|', pgno: '', dashboardurl: 'credit rating', category: 'Company' }),
    });
    const t = await r.text();
    if (!r.ok) return { error: 'http ' + r.status };
    try { const inner = JSON.parse(JSON.parse(t).d.Result || '{}'); return { ok: true, table: inner.Table || [] }; }
    catch (e) { return { error: 'parse' }; }
  }, { hash });
}

// returns { ratingDate, agency, grade, outlook } for the most recent rating row
async function fetchLatestRating(page, hash) {
  const res = await apiRating(page, hash);
  if (res.error || !res.table || !res.table.length) return null;
  const rows = res.table.map((r) => ({
    ratingDate: stripHtml(r['Rating Date'] || r.Rating_Date),
    agency: stripHtml(r.Agency),
    grade: stripHtml(r.Rating),
    instrument: stripHtml(r['Instrument Category'] || r.Instrument_Category),
    development: stripHtml(r.Development),
    outlook: stripHtml(r.Outlook),
  })).filter((r) => r.grade || r.agency);
  rows.sort((a, b) => (parseDate(b.ratingDate) || 0) - (parseDate(a.ratingDate) || 0));
  return rows[0] || null;
}

// fetch ALL charge pages for a company (handles pagination, dedup by Charge_Id)
async function fetchAllCharges(page, hash) {
  const seen = new Set();
  const out = [];
  for (let pg = 1; pg <= 20; pg++) {
    const res = await apiCharges(page, hash, pg);
    if (res.error) return { error: res.error, raw: res.raw, charges: out };
    const rows = res.table;
    if (!rows.length) break;
    let added = 0;
    for (const row of rows) {
      const id = String(row.Charge_Id || row.HIDE_RowNo || '');
      if (!id || seen.has(id)) continue;
      seen.add(id); added++;
      out.push({
        chargeId: id,
        lender: stripHtml(row.Charge_Holder),
        lenderCin: row.HIDE_CIN || '',
        amountCr: row['Charge_Amount(Rs Cr)'] === '' ? null : Number(row['Charge_Amount(Rs Cr)']),
        creationDate: stripHtml(row.Date_of_Creation),
        modificationDate: stripHtml(row.Date_of_Modification),
        assets: stripHtml(row.Assets_under_charge),
        address: stripHtml(row.Address),
      });
    }
    if (added === 0) break;        // no new rows -> stop
    if (rows.length < 50) break;   // last page
  }
  return { charges: out };
}

// low-level call to the "Charge History" view (full lifecycle: Creation/Modification/Satisfaction
// events for BOTH open and closed/satisfied charges — the superset of "Open Charge").
async function apiChargeHistory(page, hash, pgno) {
  return page.evaluate(async ({ hash, pgno }) => {
    const r = await fetch('https://www.saverisk.com/CmpAsyncDataService.aspx/ExecuteMethodStaticAsync', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json; charset=UTF-8', 'X-Requested-With': 'XMLHttpRequest' },
      body: JSON.stringify({
        cinno: hash,
        parameterNvals: '{ddl}|;{period}|;{unit}|;{ddltemplate}|;{ddlsort}|;{ddlto}|;{IsSort}|',
        pgno: String(pgno),
        dashboardurl: 'Charge History',
        category: 'Company',
      }),
    });
    const t = await r.text();
    if (!r.ok) return { error: 'http ' + r.status, raw: t.slice(0, 200) };
    try {
      const inner = JSON.parse(JSON.parse(t).d.Result || '{}');
      return { ok: true, table: inner.Table || [] };
    } catch (e) { return { error: 'parse', raw: t.slice(0, 200) }; }
  }, { hash, pgno });
}

// fetch ALL charge-history events (paginates to the end; no artificial page cap).
// Returns { events:[{chargeId, lender, amountCr, date, eventType}] }, optionally filtered to
// events on/after `sinceMs` (epoch ms). Undated rows are kept (flagged via date:'').
async function fetchChargeHistory(page, hash, sinceMs = 0, maxPages = 120) {
  const seen = new Set();
  const out = [];
  for (let pg = 1; pg <= maxPages; pg++) {
    const res = await apiChargeHistory(page, hash, pg);
    if (res.error) return { error: res.error, raw: res.raw, events: out };
    const rows = res.table;
    if (!rows.length) break;
    let added = 0;
    for (const row of rows) {
      const id = String(row.Charge_ID || row.Charge_Id || '');
      const eventType = stripHtml(row['Creation_or_Modification_or_Satisfaction'] || '');
      const rawDate = stripHtml(row['Date_Charge_Creation/Modification/Satisfaction'] || '');
      const dkey = id + '|' + eventType + '|' + rawDate;
      if (seen.has(dkey)) continue;
      seen.add(dkey); added++;
      const amt = row['Charge_Amount_Secured_in_(Rs._Crore)'];
      out.push({
        chargeId: id,
        lender: stripHtml(row.Charge_Holder),
        amountCr: amt === '' || amt == null ? null : Number(amt),
        date: rawDate,
        eventType,
      });
    }
    if (added === 0) break;        // a fully-duplicate page -> stop
    if (rows.length < 50) break;   // last page
  }
  // optional 5-year (or any) cutoff filter, keeping undated rows
  const events = sinceMs ? out.filter((e) => { const d = parseDate(e.date); return !d || d.getTime() >= sinceMs; }) : out;
  return { events, fetched: out.length };
}

// ---------- json persistence ----------
function loadJson(file, fallback) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return fallback; }
}
function saveJson(file, obj) { fs.writeFileSync(file, JSON.stringify(obj, null, 2)); }

module.exports = {
  BASE, normName, similarity, parseDate, daysAgo, fmtDate, stripHtml,
  apiSearch, apiCharges, fetchAllCharges, apiChargeHistory, fetchChargeHistory, apiRating, fetchLatestRating, fetchRatingUI, fetchOverview, loadJson, saveJson,
};
