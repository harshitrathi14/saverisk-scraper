// reports.js — turn the scraped charge dataset into the NA-vs-Others analyses,
// then write an Excel workbook + a self-contained HTML dashboard.
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const L = require('./lib');

const WINDOWS = [
  { key: '1w', days: 7, label: 'Last 1 Week' },
  { key: '1m', days: 30, label: 'Last 1 Month' },
  { key: '3m', days: 90, label: 'Last 3 Months' },
  { key: '6m', days: 180, label: 'Last 6 Months' },
];
const TRUSTEE_RE = /trusteeship|debenture trustee|catalyst trustee|beacon trustee|vistra|axis trustee|idbi trustee/i;

function inWindow(date, now, days) { return date && date >= L.daysAgo(days, now) && date <= now; }
const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const sum = (arr, f) => arr.reduce((a, x) => a + (Number(f(x)) || 0), 0);

// ---------- build all analyses from the flat charge list ----------
// charges: [{ entity, short_name, sector, exposure, lender, lenderCin, amountCr, date(Date|null),
//             creationDate(str), chargeId, isNA, isTrustee }]
// entities: [{ name, short_name, sector, exposure, rating, status, ... }]
function buildAnalyses(entities, charges, now, prevState) {
  const dated = charges.filter((c) => c.date);

  // first-time lender->borrower: earliest charge per (entity,lender)
  const pairMin = {};
  for (const c of dated) {
    const k = c.entity + '||' + c.lender;
    if (!pairMin[k] || c.date < pairMin[k]) pairMin[k] = c.date;
  }
  for (const c of dated) c.firstTime = (c.date && c.date.getTime() === pairMin[c.entity + '||' + c.lender].getTime());

  // ---- NA vs Others, overall per window ----
  const naVsOthers = WINDOWS.map((w) => {
    const inW = dated.filter((c) => inWindow(c.date, now, w.days));
    const na = inW.filter((c) => c.isNA), oth = inW.filter((c) => !c.isNA);
    return {
      window: w.label,
      na_charges: na.length, na_amount_cr: r2(sum(na, (x) => x.amountCr)), na_borrowers: new Set(na.map((x) => x.entity)).size,
      other_charges: oth.length, other_amount_cr: r2(sum(oth, (x) => x.amountCr)), other_borrowers: new Set(oth.map((x) => x.entity)).size,
    };
  });

  // ---- NA vs Others by sector, per window ----
  const sectorsAll = [...new Set(entities.map((e) => e.sector || 'Unclassified'))];
  const naVsOthersBySector = {};
  for (const w of WINDOWS) {
    const inW = dated.filter((c) => inWindow(c.date, now, w.days));
    const bySec = {};
    for (const c of inW) {
      const s = c.sector || 'Unclassified';
      if (!bySec[s]) bySec[s] = { sector: s, na_charges: 0, na_amount_cr: 0, other_charges: 0, other_amount_cr: 0 };
      if (c.isNA) { bySec[s].na_charges++; bySec[s].na_amount_cr += Number(c.amountCr) || 0; }
      else { bySec[s].other_charges++; bySec[s].other_amount_cr += Number(c.amountCr) || 0; }
    }
    naVsOthersBySector[w.key] = Object.values(bySec)
      .map((x) => ({ ...x, na_amount_cr: r2(x.na_amount_cr), other_amount_cr: r2(x.other_amount_cr) }))
      .sort((a, b) => (b.other_charges + b.na_charges) - (a.other_charges + a.na_charges));
  }

  // ---- charge creations by OTHER (non-NA) lenders, per window ----
  // One row per charge (entity × lender × charge), never clubbed, each with its own date.
  const externalCharges = {};
  for (const w of WINDOWS) {
    externalCharges[w.key] = dated
      .filter((c) => !c.isNA && inWindow(c.date, now, w.days))
      .map((c) => ({
        entity: c.entity, sector: c.sector || 'Unclassified', exposure_cr: r2((Number(c.exposure) || 0) / 1e7),
        lender: c.lender, amount_cr: r2(c.amountCr), charge_date: c.creationDate,
        type: c.isTrustee ? 'Debenture/Trustee' : 'Charge/Loan', charge_id: c.chargeId,
      }))
      .sort((a, b) => (L.parseDate(b.charge_date) || 0) - (L.parseDate(a.charge_date) || 0) || (b.amount_cr - a.amount_cr));
  }

  // ---- charge creations BY Northern Arc, per window (parallel to externalCharges) ----
  const naCharges = {};
  for (const w of WINDOWS) {
    naCharges[w.key] = dated
      .filter((c) => c.isNA && inWindow(c.date, now, w.days))
      .map((c) => ({
        entity: c.entity, sector: c.sector || 'Unclassified', exposure_cr: r2((Number(c.exposure) || 0) / 1e7),
        lender: c.lender, amount_cr: r2(c.amountCr), charge_date: c.creationDate,
        type: c.isTrustee ? 'Debenture/Trustee' : 'Charge/Loan', charge_id: c.chargeId,
      }))
      .sort((a, b) => (L.parseDate(b.charge_date) || 0) - (L.parseDate(a.charge_date) || 0) || (b.amount_cr - a.amount_cr));
  }

  // ---- most active lenders per window (all lenders, NA flagged) ----
  const activeLenders = {};
  for (const w of WINDOWS) {
    const inW = dated.filter((c) => inWindow(c.date, now, w.days));
    const m = {};
    for (const c of inW) {
      const k = c.lender || '(blank)';
      if (!m[k]) m[k] = { lender: k, is_na: c.isNA ? 'Yes' : '', charges: 0, amount_cr: 0, borrowers: new Set() };
      m[k].charges++; m[k].amount_cr += Number(c.amountCr) || 0; m[k].borrowers.add(c.entity);
    }
    activeLenders[w.key] = Object.values(m)
      .map((x) => ({ lender: x.lender, is_na: x.is_na, charges_created: x.charges, total_amount_cr: r2(x.amount_cr), distinct_borrowers: x.borrowers.size }))
      .sort((a, b) => b.charges_created - a.charges_created || b.total_amount_cr - a.total_amount_cr);
  }

  // ---- first-time lender->borrower per window ----
  const firstTime = {};
  for (const w of WINDOWS) {
    firstTime[w.key] = dated.filter((c) => c.firstTime && inWindow(c.date, now, w.days))
      .map((c) => ({ entity: c.entity, sector: c.sector || 'Unclassified', lender: c.lender, is_na: c.isNA ? 'Yes' : '', amount_cr: r2(c.amountCr), charge_date: c.creationDate, type: c.isTrustee ? 'Debenture/Trustee' : 'Charge/Loan' }))
      .sort((a, b) => (L.parseDate(b.charge_date) || 0) - (L.parseDate(a.charge_date) || 0));
  }

  // ---- per-entity summary + latest charge ----
  const summaryRows = [];
  const latestRows = [];
  for (const e of entities) {
    const ec = charges.filter((c) => c.entity === e.name);
    const ecD = [...ec].sort((a, b) => (b.date || 0) - (a.date || 0));
    const latest = ecD[0];
    const latestExt = ecD.find((c) => !c.isNA);
    const naC = ec.filter((c) => c.isNA), othC = ec.filter((c) => !c.isNA);
    summaryRows.push({
      entity: e.name, short_name: e.short_name, sector: e.sector || 'Unclassified',
      saverisk_sector: e.saverisk_sector || '', saverisk_industry: e.saverisk_industry || '',
      exposure_cr: r2((Number(e.exposure) || 0) / 1e7), rating: e.rating ? `${e.rating.grade} (${e.rating.agency}) ${e.rating.ratingDate}` : '',
      status: e.status, total_open_charges: ec.length, na_charges: naC.length, other_charges: othC.length,
      latest_lender: latest ? latest.lender : '', latest_amount_cr: latest ? r2(latest.amountCr) : '', latest_date: latest ? latest.creationDate : '',
      latest_external_lender: latestExt ? latestExt.lender : '', latest_external_date: latestExt ? latestExt.creationDate : '',
    });
    if (latest) latestRows.push({ entity: e.name, sector: e.sector || 'Unclassified', exposure_cr: r2((Number(e.exposure) || 0) / 1e7), lender: latest.lender, is_na: latest.isNA ? 'Yes' : '', amount_cr: r2(latest.amountCr), charge_date: latest.creationDate, type: latest.isTrustee ? 'Debenture/Trustee' : 'Charge/Loan' });
  }

  // ---- new since last run (by chargeId) ----
  const prevCharges = (prevState && prevState.charges) || {};
  const firstRun = !(prevState && prevState.lastRun);
  const newSinceRun = [];
  for (const c of charges) {
    const prevIds = new Set(prevCharges[c.entity.toUpperCase()] || []);
    if (!firstRun && !prevIds.has(c.chargeId)) {
      newSinceRun.push({ entity: c.entity, sector: c.sector || 'Unclassified', lender: c.lender, is_na: c.isNA ? 'Yes' : '', amount_cr: r2(c.amountCr), charge_date: c.creationDate, charge_id: c.chargeId });
    }
  }
  newSinceRun.sort((a, b) => (L.parseDate(b.charge_date) || 0) - (L.parseDate(a.charge_date) || 0));

  const allCharges = charges.map((c) => ({ entity: c.entity, sector: c.sector || 'Unclassified', lender: c.lender, is_na: c.isNA ? 'Yes' : '', amount_cr: r2(c.amountCr), charge_date: c.creationDate, charge_id: c.chargeId, type: c.isTrustee ? 'Debenture/Trustee' : 'Charge/Loan' }));

  return { WINDOWS, now, firstRun, naVsOthers, naVsOthersBySector, externalCharges, naCharges, activeLenders, firstTime, summaryRows, latestRows, newSinceRun, allCharges, entities };
}

module.exports = { WINDOWS, TRUSTEE_RE, buildAnalyses };
