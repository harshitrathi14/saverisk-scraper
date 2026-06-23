// ppt.js — build a shareable PowerPoint deck from the analyses.
const PptxGenJS = require('pptxgenjs');

const NA = '0D6EFD', OTH = 'E8590C', INK = '1A2233', MUT = '6B7785', LINE = 'E6E9EF', BG = 'F4F6FA';
const fmtCr = (n) => '₹' + (Number(n) || 0).toLocaleString('en-IN', { maximumFractionDigits: 1 }) + ' Cr';
const fmtN = (n) => (Number(n) || 0).toLocaleString('en-IN');

function writePpt(A, file, meta) {
  const p = new PptxGenJS();
  p.defineLayout({ name: 'WIDE', width: 13.33, height: 7.5 });
  p.layout = 'WIDE';
  p.theme = { headFontFace: 'Segoe UI', bodyFontFace: 'Segoe UI' };

  const header = (s, title, sub) => {
    s.background = { color: 'FFFFFF' };
    s.addShape('rect', { x: 0, y: 0, w: 13.33, h: 0.9, fill: { color: '15315B' } });
    s.addText(title, { x: 0.5, y: 0.12, w: 12.3, h: 0.66, color: 'FFFFFF', fontSize: 22, bold: true, valign: 'middle' });
    if (sub) s.addText(sub, { x: 0.5, y: 0.95, w: 12.3, h: 0.3, color: MUT, fontSize: 11 });
  };
  const tbl = (s, cols, rows, x, y, w, max, opts = {}) => {
    const head = cols.map((c) => ({ text: c.label, options: { bold: true, color: 'FFFFFF', fill: { color: '15315B' }, fontSize: 9 } }));
    const body = (rows || []).slice(0, max).map((r, i) => cols.map((c) => {
      let v = r[c.key]; if (c.cr) v = fmtCr(v); else if (c.num) v = fmtN(v);
      const isNA = c.na && r[c.key] === 'Yes';
      return { text: String(v == null ? '' : v), options: { fontSize: 8.5, color: isNA ? NA : INK, bold: isNA, fill: { color: i % 2 ? 'F7F9FC' : 'FFFFFF' }, align: c.right ? 'right' : 'left' } };
    }));
    s.addTable([head, ...body], { x, y, w, colW: opts.colW, border: { type: 'solid', color: LINE, pt: 0.5 }, autoPage: false });
  };

  // 1) Title
  let s = p.addSlide(); s.background = { color: '0B1220' };
  s.addText('Lending Intelligence', { x: 0.8, y: 2.5, w: 11.7, h: 0.9, color: 'FFFFFF', fontSize: 40, bold: true });
  s.addText('Northern Arc vs Other Lenders — Charge-Creation Comparison', { x: 0.8, y: 3.4, w: 11.7, h: 0.6, color: '8AB4F8', fontSize: 20 });
  s.addText(`Onboarded portfolio: ${fmtN(meta.entityCount)} entities   ·   Source: Saverisk (MCA charges)   ·   ${meta.generated}`, { x: 0.8, y: 4.2, w: 11.7, h: 0.4, color: MUT, fontSize: 12 });

  // 2) Headline — NA vs Others by window (clustered bar of charge counts)
  s = p.addSlide(); header(s, 'Charge creation: Northern Arc vs Others', 'New charges created across the portfolio by window');
  const labels = A.naVsOthers.map((w) => w.window);
  s.addChart(p.ChartType.bar, [
    { name: 'Other lenders', labels, values: A.naVsOthers.map((w) => w.other_charges) },
    { name: 'Northern Arc', labels, values: A.naVsOthers.map((w) => w.na_charges) },
  ], { x: 0.5, y: 1.3, w: 7.2, h: 5.6, barDir: 'col', chartColors: [OTH, NA], showLegend: true, legendPos: 'b', showValue: true, dataLabelFontSize: 10, catAxisLabelFontSize: 11, valAxisHidden: false });
  // side KPIs
  let ky = 1.5;
  for (const w of A.naVsOthers) {
    s.addShape('roundRect', { x: 8.1, y: ky, w: 4.7, h: 1.25, fill: { color: 'FFFFFF' }, line: { color: LINE, width: 1 }, rectRadius: 0.06 });
    s.addText(w.window.toUpperCase(), { x: 8.3, y: ky + 0.08, w: 4.3, h: 0.3, color: MUT, fontSize: 10, bold: true });
    s.addText([{ text: `${fmtN(w.other_charges)} `, options: { color: OTH, bold: true, fontSize: 18 } }, { text: 'by others   ', options: { color: MUT, fontSize: 11 } }, { text: `${fmtN(w.na_charges)} `, options: { color: NA, bold: true, fontSize: 18 } }, { text: 'NA', options: { color: MUT, fontSize: 11 } }], { x: 8.3, y: ky + 0.38, w: 4.3, h: 0.4 });
    s.addText(`Others ${fmtCr(w.other_amount_cr)} · NA ${fmtCr(w.na_amount_cr)}`, { x: 8.3, y: ky + 0.82, w: 4.3, h: 0.3, color: INK, fontSize: 10 });
    ky += 1.4;
  }

  // 3) Sector-wise (use 1 month window) — clustered bar
  const secW = A.naVsOthersBySector['1m'] || [];
  s = p.addSlide(); header(s, 'Charge creation by sector (last 1 month)', 'Where competitors are most active vs Northern Arc');
  const secTop = secW.slice(0, 10);
  if (secTop.length) {
    s.addChart(p.ChartType.bar, [
      { name: 'Other lenders', labels: secTop.map((r) => r.sector), values: secTop.map((r) => r.other_charges) },
      { name: 'Northern Arc', labels: secTop.map((r) => r.sector), values: secTop.map((r) => r.na_charges) },
    ], { x: 0.5, y: 1.3, w: 12.3, h: 5.7, barDir: 'bar', chartColors: [OTH, NA], showLegend: true, legendPos: 'b', showValue: true, dataLabelFontSize: 9, catAxisLabelFontSize: 10 });
  } else s.addText('No charge activity in the last month.', { x: 0.5, y: 3, w: 12, h: 0.5, color: MUT, fontSize: 14 });

  // 4) Entities funded by other lenders (1 month) — one row per charge
  const chargeColsP = [
    { key: 'entity', label: 'Entity' }, { key: 'sector', label: 'Sector' },
    { key: 'lender', label: 'Lender' }, { key: 'amount_cr', label: 'Amount', cr: true, right: true },
    { key: 'charge_date', label: 'Date' }, { key: 'type', label: 'Type' },
  ];
  s = p.addSlide(); header(s, 'Entities funded by other lenders (last 1 month)', 'Each charge by a non-Northern-Arc lender, with its date');
  tbl(s, chargeColsP, A.externalCharges['1m'], 0.4, 1.25, 12.5, 15, { colW: [3.2, 2.4, 3.4, 1.3, 1.3, 1.3] });

  // 4b) Entities funded by Northern Arc (1 month)
  s = p.addSlide(); header(s, 'Entities funded by Northern Arc (last 1 month)', 'Each charge created by Northern Arc, with its date');
  tbl(s, chargeColsP, A.naCharges['1m'], 0.4, 1.25, 12.5, 15, { colW: [3.2, 2.4, 3.4, 1.3, 1.3, 1.3] });

  // 5) Most active competitor lenders (1 month)
  s = p.addSlide(); header(s, 'Most active lenders (last 1 month)', 'Ranked by charges created across the portfolio');
  tbl(s, [
    { key: 'lender', label: 'Lender' }, { key: 'is_na', label: 'NA?', na: true },
    { key: 'charges_created', label: 'Charges', num: true, right: true },
    { key: 'total_amount_cr', label: 'Amount', cr: true, right: true },
    { key: 'distinct_borrowers', label: 'Borrowers', num: true, right: true },
  ], A.activeLenders['1m'], 0.6, 1.25, 11.5, 16, { colW: [6.0, 1.0, 1.5, 1.5, 1.5] });

  // 6) First-time lenders (1 month)
  s = p.addSlide(); header(s, 'First-time lender → borrower (last 1 month)', 'New lending relationships, not top-ups');
  tbl(s, [
    { key: 'entity', label: 'Borrower' }, { key: 'sector', label: 'Sector' },
    { key: 'lender', label: 'First-time Lender' }, { key: 'is_na', label: 'NA?', na: true },
    { key: 'amount_cr', label: 'Amount', cr: true, right: true }, { key: 'charge_date', label: 'Date' },
  ], A.firstTime['1m'], 0.4, 1.25, 12.5, 15, { colW: [3.0, 2.2, 3.6, 0.8, 1.4, 1.5] });

  return p.writeFile({ fileName: file });
}

module.exports = { writePpt };
