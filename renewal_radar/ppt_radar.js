// renewal_radar/ppt_radar.js — polished PowerPoint deck for the Renewal Radar module.
// Matches the house style of the main ppt.js (navy/gold, Segoe UI). Consumes the in-memory
// lifecycle / forecast / leads objects produced by the pipeline.
const PptxGenJS = require('pptxgenjs');

const C = {
  navy: '0B1F3A', navy2: '13294B', bar: '15315B',
  oth: 'E8590C', na: '1D6FE0', naSoft: 'DCE9FB', green: '1E8E5A', greenSoft: 'DDF1E7',
  ink: '1A2233', mut: '6B7785', line: 'E6E9EF', soft: 'F7F9FC', white: 'FFFFFF', gold: 'F2B705',
};
const fmtCr = (n) => '₹' + (Number(n) || 0).toLocaleString('en-IN', { maximumFractionDigits: 1 }) + ' Cr';
const fmtN = (n) => (Number(n) || 0).toLocaleString('en-IN');
const fmtPct = (n) => Math.round((Number(n) || 0) * 100) + '%';
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const fmtDate = (iso) => { if (!iso) return ''; const d = new Date(iso); return isNaN(d) ? '' : `${String(d.getUTCDate()).padStart(2, '0')} ${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}`; };
const ymLabel = (ym) => { if (!ym) return ''; const [y, m] = ym.split('-'); return `${MONTHS[(+m) - 1]} ${y}`; };

// Canonical disclaimer — kept identical across the HTML dashboard, Excel and this deck.
const DISCLAIMER = "Basis of estimation & disclaimer — These predictions extrapolate the historical lender↔borrower charge-creation cadence; they are indicative prospecting signals, not assured events. Predicted dates may vary by ~1–2 months (charge-filing / registration lag and deal timing). Whether a drawdown materialises — and from which lender — depends on the entity's evolving credit rating and capital position: a stronger balance sheet or an upgraded rating can move the borrower to lower-cost lenders (e.g. large banks), away from the NACL-competible set; a weaker position can do the reverse. Amounts are a secured-charge proxy, not sanctioned loan values. Validate each lead against Northern Arc's internal rating and capital-position view before acting.";
const DISCLAIMER_SHORT = "Indicative extrapolation of historical cadence — not assured. Dates ±1–2 months; may not materialise if the entity's rating/capital position changes (a stronger entity may shift to lower-cost banks).";

// Append the Renewal Radar section to an existing deck `p`. `state` = { pageNo } shares the
// continuous page counter with the main slides. (Standalone writeRadarPpt kept below for reuse.)
function addRadarSlides(p, lifecycle, forecast, leads, meta, state) {
  meta = meta || {};
  state = state || { pageNo: 0 };

  const footer = (s) => {
    state.pageNo++;
    s.addShape('line', { x: 0.5, y: 7.06, w: 12.33, h: 0, line: { color: C.line, width: 1 } });
    s.addText('Saverisk · Renewal Radar — predicted next-tranche lead intelligence', { x: 0.5, y: 7.1, w: 8, h: 0.3, color: C.mut, fontSize: 8 });
    s.addText('CONFIDENTIAL — Northern Arc', { x: 8.5, y: 7.1, w: 3.0, h: 0.3, color: C.mut, fontSize: 8, align: 'right' });
    s.addText(String(state.pageNo), { x: 12.6, y: 7.1, w: 0.4, h: 0.3, color: C.mut, fontSize: 8, align: 'right' });
  };
  const header = (title, sub, accent = C.bar) => {
    const s = p.addSlide(); s.background = { color: C.white };
    s.addShape('rect', { x: 0, y: 0, w: 13.33, h: 0.86, fill: { color: C.navy } });
    s.addShape('rect', { x: 0, y: 0, w: 0.16, h: 0.86, fill: { color: accent } });
    s.addText(title, { x: 0.45, y: 0.06, w: 12.4, h: 0.52, color: C.white, fontSize: 21, bold: true, valign: 'middle' });
    if (sub) s.addText(sub, { x: 0.45, y: 0.56, w: 12.4, h: 0.26, color: 'A9C3E6', fontSize: 11.5 });
    footer(s);
    return s;
  };
  const divider = (title, sub, accent = C.gold) => {
    const d = p.addSlide(); d.background = { color: C.navy };
    d.addShape('rect', { x: 0.8, y: 3.18, w: 1.7, h: 0.09, fill: { color: accent } });
    d.addText(title, { x: 0.8, y: 2.2, w: 11.7, h: 1.0, color: C.white, fontSize: 33, bold: true });
    if (sub) d.addText(sub, { x: 0.8, y: 3.42, w: 11.7, h: 0.5, color: 'C7D4E6', fontSize: 14.5 });
  };
  const tbl = (s, cols, rows, max, colW, y = 1.12) => {
    if (!rows || !rows.length) { s.addText('No records.', { x: 0.5, y: 3.2, w: 12, h: 0.5, color: C.mut, fontSize: 14, italic: true }); return; }
    const head = cols.map((c) => ({ text: c.label, options: { bold: true, color: C.white, fill: { color: C.bar }, fontSize: 10.5, align: c.right ? 'right' : 'left', valign: 'middle' } }));
    const body = rows.slice(0, max).map((r, i) => cols.map((c) => {
      let v = r[c.key]; if (c.cr) v = fmtCr(v); else if (c.num) v = fmtN(v); else if (c.pct) v = fmtPct(v); else if (c.date) v = fmtDate(v); else if (c.ym) v = ymLabel(v);
      return { text: String(v == null ? '' : v), options: { fontSize: 9.5, color: C.ink, fill: { color: i % 2 ? C.soft : C.white }, align: c.right ? 'right' : 'left', valign: 'middle' } };
    }));
    s.addTable([head, ...body], { x: 0.4, y, w: 12.5, colW, rowH: 0.3, border: { type: 'solid', color: C.line, pt: 0.5 }, autoPage: false });
    if (rows.length > max) s.addText(`… and ${fmtN(rows.length - max)} more — see leads.csv / the Excel workbook`, { x: 0.4, y: 6.78, w: 12, h: 0.25, color: C.mut, fontSize: 9, italic: true });
  };
  const kpi = (s, items, y = 1.3) => {
    const tileW = 12.5 / items.length;
    items.forEach((it, i) => {
      const x = 0.4 + i * tileW;
      s.addShape('roundRect', { x: x + 0.06, y, w: tileW - 0.12, h: 1.7, fill: { color: C.soft }, line: { color: C.line, width: 1 }, rectRadius: 0.06 });
      s.addText(String(it.v), { x: x + 0.06, y: y + 0.22, w: tileW - 0.12, h: 0.6, color: it.c || C.na, fontSize: 26, bold: true, align: 'center' });
      s.addText(it.l, { x: x + 0.06, y: y + 0.95, w: tileW - 0.12, h: 0.6, color: C.mut, fontSize: 10, align: 'center', valign: 'top' });
    });
  };

  // ---- derive aggregates ----
  const leadCr = (leads || []).reduce((a, l) => a + (Number(l.expectedAmtCr) || 0), 0);
  const leadEntities = new Set((leads || []).map((l) => l.entity)).size;
  // leads per upcoming month
  const byMonth = {};
  (leads || []).forEach((l) => { byMonth[l.ym] = (byMonth[l.ym] || 0) + 1; });
  const monthsSorted = Object.keys(byMonth).sort();
  // top competible lenders by lead count
  const byLender = {};
  (leads || []).forEach((l) => { const k = l.lender; byLender[k] = (byLender[k] || 0) + 1; });
  const topLenders = Object.entries(byLender).sort((a, b) => b[1] - a[1]).slice(0, 12);
  // an illustrative recurring relationship (highest-n named, non-NACL) from lifecycle
  let ex = null;
  for (const e of (lifecycle.entities || [])) {
    for (const lc of (e.lenderCadence || [])) {
      if (lc.isNACL) continue;
      if (!ex || lc.n > ex.lc.n) ex = { entity: e.entity, sector: e.sector, lc };
    }
  }

  // ---------- SECTION DIVIDER (Part 2 of the combined deck) ----------
  let s = p.addSlide(); s.background = { color: C.navy };
  s.addShape('rect', { x: 0.8, y: 3.18, w: 1.7, h: 0.09, fill: { color: C.gold } });
  s.addText('Part 2 · Renewal Radar', { x: 0.8, y: 1.7, w: 11.7, h: 0.6, color: C.gold, fontSize: 16, bold: true });
  s.addText('Predicting the next loan tranche — before the competition lends', { x: 0.8, y: 2.2, w: 11.7, h: 1.0, color: C.white, fontSize: 30, bold: true });
  s.addText('For each onboarded entity: when the next renewal / top-up is due, which lender is likely to give it, and where Northern Arc can win on rate.', { x: 0.8, y: 3.42, w: 11.2, h: 0.7, color: 'C7D4E6', fontSize: 14 });
  s.addText(`${fmtN(meta.entityCount || (lifecycle.entities || []).length)} onboarded entities    ·    5-year charge history (open + satisfied)    ·    Source: Saverisk (MCA)`, { x: 0.8, y: 5.6, w: 11.7, h: 0.4, color: C.mut, fontSize: 12 });

  // ---------- THE OPPORTUNITY ----------
  s = header('The opportunity', 'Turn charge-creation history into a forward lead list for the business team', C.gold);
  s.addText([
    { text: 'The problem.  ', options: { bold: true, color: C.ink, fontSize: 13 } },
    { text: 'Other lenders keep funding Northern Arc’s onboarded entities — renewals and top-ups land on a rhythm. By the time we see a new charge, the loan is already given.', options: { color: C.ink, fontSize: 13 } },
  ], { x: 0.5, y: 1.25, w: 12.3, h: 0.9, valign: 'top' });
  s.addText([
    { text: 'The idea.  ', options: { bold: true, color: C.ink, fontSize: 13 } },
    { text: 'Most wholesale facilities recur. Learn each entity→lender cadence from 5 years of charge data, project the next drawdown date, and surface it as a dated lead — so the team can approach the entity first.', options: { color: C.ink, fontSize: 13 } },
  ], { x: 0.5, y: 2.15, w: 12.3, h: 0.9, valign: 'top' });
  s.addShape('roundRect', { x: 0.4, y: 3.25, w: 12.5, h: 1.15, fill: { color: C.navy }, rectRadius: 0.06 });
  s.addText([
    { text: 'The edge.   ', options: { color: C.gold, fontSize: 12, bold: true } },
    { text: 'Northern Arc can only win where the incumbent lender is similarly- or lower-rated (similar cost of funds). So the radar filters to NACL-competible lenders — screening out the AAA banks NACL can’t out-price — and hands the team only the winnable renewals.', options: { color: C.white, fontSize: 12.5 } },
  ], { x: 0.7, y: 3.37, w: 11.9, h: 0.9, valign: 'middle' });
  s.addText('What makes it possible now', { x: 0.5, y: 4.6, w: 12, h: 0.3, color: C.ink, fontSize: 12, bold: true });
  s.addText([
    'Saverisk’s “Charge History” view exposes loan SATISFACTION (repayment/closure) events — not just open charges — so we see the full raise-and-repay rhythm over 5 years.',
    'Open charges give the real current lender names (incl. NBFCs); history gives the timing. Combined, they yield a per-entity funding cadence.',
    'A NACL-competible filter (lender credit tier) converts the raw forecast into a ranked, dated lead list.',
  ].map((t) => ({ text: t, options: { bullet: { code: '2022' }, color: C.ink, fontSize: 11 } })), { x: 0.7, y: 4.95, w: 12, h: 1.8, lineSpacingMultiple: 1.12 });

  // ---------- DATA & METHOD ----------
  s = header('How it works', 'Three inputs → cadence → forecast → competible leads', C.na);
  const steps = [
    { t: '1 · Open charges', d: 'Who lends now + live exposure. Real lender names & CINs.', src: 'charges_cache.json', c: C.na },
    { t: '2 · Charge history (5y)', d: 'Creation + Satisfaction events incl. closed loans → raise-vs-repay rhythm.', src: 'charge_history_cache.json', c: C.green },
    { t: '3 · Lender tier', d: 'NACL-competible filter (≤ AA-). Populated from NACL internal counterparty ratings.', src: 'lender_ratings.json', c: C.gold },
  ];
  steps.forEach((st, i) => {
    const x = 0.4 + i * (12.5 / 3);
    s.addShape('roundRect', { x: x + 0.08, y: 1.25, w: 12.5 / 3 - 0.16, h: 2.15, fill: { color: C.soft }, line: { color: st.c, width: 1.5 }, rectRadius: 0.06 });
    s.addText(st.t, { x: x + 0.25, y: 1.4, w: 12.5 / 3 - 0.5, h: 0.4, color: st.c, fontSize: 14, bold: true });
    s.addText(st.d, { x: x + 0.25, y: 1.9, w: 12.5 / 3 - 0.5, h: 1.0, color: C.ink, fontSize: 11 });
    s.addText(st.src, { x: x + 0.25, y: 3.0, w: 12.5 / 3 - 0.5, h: 0.3, color: C.mut, fontSize: 9, italic: true });
  });
  s.addText('Pipeline', { x: 0.5, y: 3.7, w: 12, h: 0.3, color: C.ink, fontSize: 12, bold: true });
  s.addText('lifecycle.js  (per-entity cadence + monthly created/satisfied flow)   →   predict.js  (forward 12-month calendar, each row scored by regularity & history depth)   →   leads.js + ratings.js  (keep only NACL-competible incumbents)   →   Excel · self-contained HTML dashboard · CSVs.', { x: 0.5, y: 4.0, w: 12.3, h: 0.7, color: C.ink, fontSize: 11.5 });
  s.addText('Confidence', { x: 0.5, y: 4.85, w: 12, h: 0.3, color: C.ink, fontSize: 12, bold: true });
  s.addText('Each forecast carries a confidence from how regular and how deep the relationship is. A steady lender with 5+ prior drawdowns on a clean cadence scores high (~0.8); a sparse or noisy one scores low (~0.3). The team works the high-confidence, near-term rows first.', { x: 0.5, y: 5.15, w: 12.3, h: 0.9, color: C.ink, fontSize: 11.5 });

  // ---------- KEY NUMBERS ----------
  s = header('By the numbers', 'What the radar currently sees across the onboarded book', C.bar);
  kpi(s, [
    { v: fmtN(meta.entityCount || (lifecycle.entities || []).length), l: 'onboarded entities tracked', c: C.ink },
    { v: fmtN(meta.eventCount), l: 'charge events analysed (5y)', c: C.ink },
    { v: fmtN(meta.satisfactionCount), l: 'loan repayments (satisfactions) seen', c: C.green },
    { v: fmtN((forecast || []).length), l: 'forecast events, next 12 months', c: C.na },
  ]);
  kpi(s, [
    { v: fmtN((leads || []).length), l: 'NACL-competible leads', c: C.oth },
    { v: fmtN(leadEntities), l: 'distinct entities with a lead', c: C.oth },
    { v: fmtCr(leadCr), l: 'expected ₹ in competible leads (proxy)', c: C.oth },
    { v: fmtN(topLenders.length ? Object.keys(byLender).length : 0), l: 'distinct competible incumbent lenders', c: C.na },
  ], 3.25);
  s.addShape('roundRect', { x: 0.4, y: 5.25, w: 12.5, h: 1.1, fill: { color: C.navy }, rectRadius: 0.06 });
  s.addText([
    { text: 'Read me   ', options: { color: C.gold, fontSize: 12, bold: true } },
    { text: '₹ figures are a secured-charge proxy (charges are over-secured), so treat them as relative magnitude, not a quote. Leads with UNKNOWN-tier lenders are candidates pending the internal ratings map.', options: { color: C.white, fontSize: 12 } },
  ], { x: 0.7, y: 5.37, w: 11.9, h: 0.85, valign: 'middle' });

  // ---------- LEADS BY MONTH (chart) ----------
  s = header('Upcoming competible leads — by month', 'When the winnable renewals/top-ups are due across the next 12 months', C.oth);
  if (monthsSorted.length) {
    s.addChart(p.ChartType.bar, [{ name: 'NACL-competible leads', labels: monthsSorted.map(ymLabel), values: monthsSorted.map((m) => byMonth[m]) }],
      { x: 0.5, y: 1.2, w: 12.3, h: 5.5, barDir: 'col', chartColors: [C.oth], showLegend: false, showValue: true, dataLabelFontSize: 11, dataLabelColor: C.ink, catAxisLabelFontSize: 12, valGridLine: { style: 'none' } });
  } else s.addText('No leads in horizon.', { x: 0.5, y: 3.2, w: 12, h: 0.5, color: C.mut, fontSize: 14, italic: true });

  // ---------- TOP COMPETIBLE LENDERS (chart) ----------
  s = header('Who we’d be competing with', 'Most frequent NACL-competible incumbent lenders in the lead list', C.na);
  if (topLenders.length) {
    s.addChart(p.ChartType.bar, [{ name: 'Leads', labels: topLenders.map((l) => l[0]), values: topLenders.map((l) => l[1]) }],
      { x: 0.5, y: 1.15, w: 12.3, h: 5.5, barDir: 'bar', chartColors: [C.na], showLegend: false, showValue: true, dataLabelFontSize: 10, catAxisLabelFontSize: 10, valGridLine: { style: 'none' } });
  } else s.addText('No competible lenders found.', { x: 0.5, y: 3.2, w: 12, h: 0.5, color: C.mut, fontSize: 14, italic: true });

  // ---------- TOP LEADS — BY QUARTER ----------
  const leadCols = [
    { key: 'ym', label: 'Month', ym: true }, { key: 'entity', label: 'Onboarded Entity' },
    { key: 'naclExposureCr', label: 'NACL Exp.', cr: true, right: true },
    { key: 'sector', label: 'Sector' }, { key: 'lender', label: 'Likely Lender (incumbent)' },
    { key: 'lenderTier', label: 'Tier', right: true }, { key: 'expectedAmtCr', label: 'Exp. ₹Cr', cr: true, right: true },
    { key: 'confidence', label: 'Conf.', pct: true, right: true }, { key: 'lastDate', label: 'Last Drawn', date: true },
  ];
  const leadW = [1.2, 2.8, 1.0, 0.9, 2.7, 0.8, 1.0, 0.7, 1.4]; // Σ = 12.5 (fits x:0.4 + w:12.5)
  const QLAB = ['Jan–Mar', 'Apr–Jun', 'Jul–Sep', 'Oct–Dec'];
  const qKey = (ym) => { const [y, m] = ym.split('-').map(Number); return y * 4 + Math.floor((m - 1) / 3); };
  const qName = (k) => `${QLAB[k % 4]} ${Math.floor(k / 4)}`;
  // bucket leads into calendar quarters
  const buckets = {};
  for (const l of (leads || [])) { const k = qKey(l.ym); (buckets[k] = buckets[k] || []).push(l); }
  const qKeys = Object.keys(buckets).map(Number).sort((a, b) => a - b);
  if (qKeys.length) {
    divider('Top leads by quarter', 'Approach these entities before the listed lender renews — highest confidence first', C.oth);
    for (const k of qKeys) {
      const rows = buckets[k].sort((a, b) => (b.confidence || 0) - (a.confidence || 0));
      const cr = rows.reduce((a, r) => a + (Number(r.expectedAmtCr) || 0), 0);
      const ents = new Set(rows.map((r) => r.entity)).size;
      s = header(`Top leads · ${qName(k)}`, `${fmtN(rows.length)} competible leads · ${fmtN(ents)} entities · ~${fmtCr(cr)} estimated (proxy)`, C.oth);
      tbl(s, leadCols, rows, 14, leadW);                    // header+14 rows -> table bottom ≈ 5.62"
      s.addText('⚠  ' + DISCLAIMER_SHORT, { x: 0.4, y: 5.8, w: 12.5, h: 0.5, color: '8A6D1A', fontSize: 9, italic: true });
    }
  }

  // ---------- WORKED EXAMPLE ----------
  if (ex) {
    s = header('Worked example — a recurring relationship', 'How one entity→lender cadence becomes a dated lead', C.green);
    s.addText([
      { text: `${ex.entity}`, options: { bold: true, color: C.ink, fontSize: 15 } },
      { text: `   (${ex.sector || '—'})`, options: { color: C.mut, fontSize: 13 } },
    ], { x: 0.5, y: 1.25, w: 12.3, h: 0.4 });
    s.addText([
      { text: 'Incumbent lender:  ', options: { color: C.mut, fontSize: 13 } },
      { text: `${ex.lc.lender}`, options: { color: C.ink, fontSize: 13, bold: true } },
    ], { x: 0.5, y: 1.75, w: 12.3, h: 0.35 });
    kpi(s, [
      { v: fmtN(ex.lc.n), l: 'prior drawdowns seen', c: C.ink },
      { v: (ex.lc.medianGapMonths != null ? ex.lc.medianGapMonths.toFixed(1) : '—') + ' mo', l: 'typical gap between draws', c: C.na },
      { v: fmtPct(ex.lc.regularity), l: 'cadence regularity', c: C.green },
      { v: fmtDate(ex.lc.lastDate), l: 'last drawdown', c: C.ink },
    ], 2.35);
    s.addShape('roundRect', { x: 0.4, y: 4.35, w: 12.5, h: 1.0, fill: { color: C.greenSoft }, line: { color: C.green, width: 1 }, rectRadius: 0.06 });
    const nextGuess = ex.lc.lastDate && ex.lc.medianGapMonths ? (() => { const d = new Date(ex.lc.lastDate); d.setMonth(d.getMonth() + Math.round(ex.lc.medianGapMonths)); return fmtDate(d.toISOString()); })() : '—';
    s.addText([
      { text: 'Prediction:   ', options: { color: C.green, fontSize: 13, bold: true } },
      { text: `${ex.lc.lender} is likely to extend the next tranche to ${ex.entity} around `, options: { color: C.ink, fontSize: 13 } },
      { text: `${nextGuess}`, options: { color: C.ink, fontSize: 13, bold: true } },
      { text: `. If ${ex.lc.lender} is NACL-competible, the team approaches now — ahead of that renewal.`, options: { color: C.ink, fontSize: 13 } },
    ], { x: 0.7, y: 4.5, w: 11.9, h: 0.75, valign: 'middle' });
    s.addText('The radar runs this logic for every entity→lender relationship with enough history, then keeps only the competible, near-term ones.', { x: 0.5, y: 5.6, w: 12.3, h: 0.5, color: C.mut, fontSize: 11, italic: true });
  }

  // ---------- LIMITATIONS (honesty) ----------
  s = header('What this is — and is not', 'Read before acting on a number', C.bar);
  s.addText([
    'Probabilistic, not certain. Each lead is a timing + likely-lender estimate with a confidence — strong for steady recurring relationships, weak for sparse ones.',
    '₹ amount ≠ loan size. Charges are over-secured and can cover multi-tranche facilities; expected ₹ is order-of-magnitude.',
    'No exact per-loan tenure. Charge IDs in the history view are unreliable, so we use aggregate raise-vs-repay flow, not per-loan maturity.',
    'Partial lender attribution in history (~⅔ bucketed as “Others”), so cadence leans on named lenders + open-charge names.',
    'Lender ratings are NOT from Saverisk (it exposes none) — populate lender_ratings.json from NACL’s internal counterparty ratings to activate true AA-/A/BBB tiering.',
    'No debt/equity feasibility gate yet — Saverisk has no net worth/borrowings; join NACL’s internal financials by CIN to gate “can the entity absorb more?”.',
  ].map((t) => ({ text: t, options: { bullet: { code: '2022' }, color: C.ink, fontSize: 11 } })), { x: 0.6, y: 1.25, w: 12.2, h: 3.85, lineSpacingMultiple: 1.12 });
  // amber disclaimer callout on the same slide (clears the bullets, ends above the footer line)
  s.addShape('roundRect', { x: 0.4, y: 5.25, w: 12.5, h: 1.5, fill: { color: 'FFF6E0' }, line: { color: C.gold, width: 1.5 }, rectRadius: 0.06 });
  s.addShape('rect', { x: 0.4, y: 5.25, w: 0.1, h: 1.5, fill: { color: C.gold } });
  s.addText([
    { text: 'Disclaimer.  ', options: { bold: true, color: '8A6D1A' } },
    { text: DISCLAIMER.replace(/^Basis of estimation & disclaimer — /, ''), options: { color: C.ink } },
  ], { x: 0.65, y: 5.32, w: 12.0, h: 1.36, fontSize: 9.5, valign: 'middle', lineSpacingMultiple: 1.04 });

  // ---------- NEXT STEPS ----------
  s = header('To make leads sharper — two NACL-internal inputs', 'Both keyed by CIN / lender name; neither is scraped', C.gold);
  s.addShape('roundRect', { x: 0.4, y: 1.3, w: 6.1, h: 4.6, fill: { color: C.soft }, line: { color: C.line, width: 1 }, rectRadius: 0.06 });
  s.addText('1 · Lender ratings', { x: 0.7, y: 1.5, w: 5.5, h: 0.4, color: C.na, fontSize: 15, bold: true });
  s.addText([
    'Drop renewal_radar/lender_ratings.json:',
    '{ "vivriti finance limited": "A+",',
    '  "incred financial services limited": "A+",',
    '  "poonawalla fincorp limited": "AAA" }',
    '',
    'Activates true tier filtering (≤ AA- = competible). Until then, non-AAA-major lenders show as UNKNOWN candidates.',
  ].map((t, i) => ({ text: t, options: { color: i >= 1 && i <= 3 ? C.ink : C.ink, fontSize: i >= 1 && i <= 3 ? 10 : 11.5, fontFace: i >= 1 && i <= 3 ? 'Consolas' : 'Segoe UI' } })), { x: 0.7, y: 2.0, w: 5.5, h: 3.7, lineSpacingMultiple: 1.15 });
  s.addShape('roundRect', { x: 6.8, y: 1.3, w: 6.1, h: 4.6, fill: { color: C.soft }, line: { color: C.line, width: 1 }, rectRadius: 0.06 });
  s.addText('2 · D/E feasibility', { x: 7.1, y: 1.5, w: 5.5, h: 0.4, color: C.green, fontSize: 15, bold: true });
  s.addText([
    'Export CIN → { net worth, total borrowings, D/E } from onboarding / monitoring.',
    '',
    'Joins on the same CIN the dataset already uses. Gates and ranks leads by headroom — “is the entity even able to take more debt, and at what leverage?”',
    '',
    'More accurate and more current than any external/scraped source — and most onboarded entities are unlisted, so external feeds won’t have them.',
  ].map((t) => ({ text: t, options: { color: C.ink, fontSize: 11.5 } })), { x: 7.1, y: 2.0, w: 5.5, h: 3.7, lineSpacingMultiple: 1.15 });
  s.addText('Run anytime:  node make_deck.js  → combined deck (this) in the project root;  node renewal_radar/radar.js  → HTML dashboard · Excel · CSVs.', { x: 0.5, y: 6.1, w: 12.3, h: 0.4, color: C.mut, fontSize: 11, italic: true });
}

module.exports = { addRadarSlides };
