// ppt.js — polished PowerPoint deck. Organised BY ANALYSIS, with 1m/2m/3m/6m inside each.
const PptxGenJS = require('pptxgenjs');

const C = {
  navy: '0B1F3A', navy2: '13294B', bar: '15315B',
  oth: 'E8590C', na: '1D6FE0', naSoft: 'DCE9FB',
  ink: '1A2233', mut: '6B7785', line: 'E6E9EF', soft: 'F7F9FC', white: 'FFFFFF', gold: 'F2B705',
};
const fmtCr = (n) => '₹' + (Number(n) || 0).toLocaleString('en-IN', { maximumFractionDigits: 1 }) + ' Cr';
const fmtN = (n) => (Number(n) || 0).toLocaleString('en-IN');
const PPTW = ['1m', '2m', '3m', '6m']; // windows shown inside each analysis section

function writePpt(A, file, meta) {
  const p = new PptxGenJS();
  p.defineLayout({ name: 'WIDE', width: 13.33, height: 7.5 });
  p.layout = 'WIDE';
  p.theme = { headFontFace: 'Segoe UI', bodyFontFace: 'Segoe UI' };
  const wlabel = (k) => (A.WINDOWS.find((w) => w.key === k) || {}).label || k;
  let pageNo = 0;

  const footer = (s) => {
    pageNo++;
    s.addShape('line', { x: 0.5, y: 7.06, w: 12.33, h: 0, line: { color: C.line, width: 1 } });
    s.addText('Saverisk · Charge-Creation Lending Intelligence', { x: 0.5, y: 7.1, w: 8, h: 0.3, color: C.mut, fontSize: 8 });
    s.addText('CONFIDENTIAL — Northern Arc', { x: 8.5, y: 7.1, w: 3.0, h: 0.3, color: C.mut, fontSize: 8, align: 'right' });
    s.addText(String(pageNo), { x: 12.6, y: 7.1, w: 0.4, h: 0.3, color: C.mut, fontSize: 8, align: 'right' });
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
    d.addText('1 month   ·   2 months   ·   3 months   ·   6 months', { x: 0.8, y: 5.6, w: 11.7, h: 0.3, color: C.mut, fontSize: 12 });
  };
  const tbl = (s, cols, rows, max, colW) => {
    if (!rows || !rows.length) { s.addText('No records in this window.', { x: 0.5, y: 3.2, w: 12, h: 0.5, color: C.mut, fontSize: 14, italic: true }); return; }
    const head = cols.map((c) => ({ text: c.label, options: { bold: true, color: C.white, fill: { color: C.bar }, fontSize: 10.5, align: c.right ? 'right' : 'left', valign: 'middle' } }));
    const body = rows.slice(0, max).map((r, i) => cols.map((c) => {
      let v = r[c.key]; if (c.cr) v = fmtCr(v); else if (c.num) v = fmtN(v);
      const isNA = c.na && r[c.key] === 'Yes';
      return { text: String(v == null ? '' : (isNA ? 'NACL' : v)), options: { fontSize: 9.5, color: isNA ? C.na : C.ink, bold: isNA, fill: { color: isNA ? C.naSoft : (i % 2 ? C.soft : C.white) }, align: c.right ? 'right' : 'left', valign: 'middle' } };
    }));
    s.addTable([head, ...body], { x: 0.4, y: 1.12, w: 12.5, colW, rowH: 0.3, border: { type: 'solid', color: C.line, pt: 0.5 }, autoPage: false });
    if (rows.length > max) s.addText(`… and ${rows.length - max} more — see the Excel workbook`, { x: 0.4, y: 6.78, w: 12, h: 0.25, color: C.mut, fontSize: 9, italic: true });
  };

  // column definitions
  const chargeCols = [
    { key: 'entity', label: 'Onboarded Entity' }, { key: 'sector', label: 'Nimbus Sector' },
    { key: 'lender', label: 'Lender' }, { key: 'amount_cr', label: 'Amount', cr: true, right: true },
    { key: 'charge_date', label: 'Date' }, { key: 'type', label: 'Type' },
  ]; const chargeW = [3.4, 1.5, 3.7, 1.3, 1.3, 1.3];
  const firstFundedCols = [
    { key: 'entity', label: 'Onboarded Entity' }, { key: 'sector', label: 'Nimbus Sector' },
    { key: 'first_lender', label: 'First Lender' }, { key: 'is_na', label: 'NACL', na: true },
    { key: 'amount_cr', label: 'Amount', cr: true, right: true }, { key: 'first_charge_date', label: 'First Charge Date' },
  ]; const firstFundedW = [3.4, 1.5, 3.7, 0.9, 1.3, 1.7];
  const lenderCols = [
    { key: 'lender', label: 'Lender' }, { key: 'is_na', label: 'NACL', na: true },
    { key: 'charges_created', label: 'Charges', num: true, right: true },
    { key: 'total_amount_cr', label: 'Amount', cr: true, right: true },
    { key: 'distinct_borrowers', label: 'Onboarded Borrowers', num: true, right: true },
  ]; const lenderW = [6.1, 1.0, 1.5, 1.7, 2.2];
  const firstTimeCols = [
    { key: 'entity', label: 'Onboarded Entity' }, { key: 'sector', label: 'Nimbus Sector' },
    { key: 'lender', label: 'First-time Lender' }, { key: 'is_na', label: 'NACL', na: true },
    { key: 'amount_cr', label: 'Amount', cr: true, right: true }, { key: 'charge_date', label: 'Date' },
  ]; const firstTimeW = [3.4, 1.5, 3.7, 0.9, 1.3, 1.7];
  const newLenderCols = [
    { key: 'lender', label: 'Lender (new to book)' }, { key: 'is_na', label: 'NACL', na: true },
    { key: 'first_borrower', label: 'First Onboarded Borrower' }, { key: 'sector', label: 'Nimbus Sector' },
    { key: 'amount_cr', label: 'Amount', cr: true, right: true }, { key: 'first_charge_date', label: 'First Charge Date' },
  ]; const newLenderW = [3.6, 0.9, 3.6, 1.5, 1.3, 1.5];

  // per-analysis renderers (draw onto a content slide for a given window)
  const renderers = {
    sector: (s, wk) => {
      const top = (A.naVsOthersBySector[wk] || []).slice(0, 10);
      if (top.length) s.addChart(p.ChartType.bar, [
        { name: 'Other lenders', labels: top.map((r) => r.sector), values: top.map((r) => r.other_charges) },
        { name: 'Northern Arc (NACL)', labels: top.map((r) => r.sector), values: top.map((r) => r.na_charges) },
      ], { x: 0.5, y: 1.15, w: 12.3, h: 5.4, barDir: 'bar', barGapWidthPct: 30, chartColors: [C.oth, C.na], showLegend: true, legendPos: 'b', showValue: true, dataLabelFontSize: 11, catAxisLabelFontSize: 11, valGridLine: { style: 'none' } });
      else s.addText('No charge activity in this window.', { x: 0.5, y: 3.2, w: 12, h: 0.5, color: C.mut, fontSize: 14, italic: true });
    },
    externalCharges: (s, wk) => tbl(s, chargeCols, A.externalCharges[wk], 16, chargeW),
    naCharges: (s, wk) => tbl(s, chargeCols, A.naCharges[wk], 16, chargeW),
    firstFunded: (s, wk) => tbl(s, firstFundedCols, A.firstTimeFunded[wk], 16, firstFundedW),
    activeLenders: (s, wk) => tbl(s, lenderCols, A.activeLenders[wk], 18, lenderW),
    firstTime: (s, wk) => tbl(s, firstTimeCols, A.firstTime[wk], 16, firstTimeW),
    newLenders: (s, wk) => tbl(s, newLenderCols, A.newLenders[wk], 16, newLenderW),
  };
  const SECTIONS = [
    { key: 'sector', title: 'Charge creation by sector — NACL vs Others', sub: 'Across Nimbus sectors (top 10 by activity)', accent: C.oth },
    { key: 'externalCharges', title: 'Onboarded entities funded by other lenders', sub: 'Each charge by a non-NACL lender, with its date', accent: C.oth },
    { key: 'naCharges', title: 'Onboarded entities funded by Northern Arc (NACL)', sub: 'Each charge created by Northern Arc, with its date', accent: C.na },
    { key: 'firstFunded', title: 'Onboarded entities funded for the FIRST time', sub: 'Earliest-ever charge in the window — and the first lender', accent: C.bar },
    { key: 'activeLenders', title: 'Most active lenders', sub: 'Ranked by charges created across the onboarded portfolio', accent: C.bar },
    { key: 'firstTime', title: 'First-time lender → onboarded borrower', sub: 'A lender’s first ever charge on that borrower — new money, not a top-up', accent: C.oth },
    { key: 'newLenders', title: 'Lenders new to the onboarded book', sub: 'First-ever charge across the whole portfolio falls in the window', accent: C.oth },
  ];

  // ---------- COVER ----------
  let s = p.addSlide(); s.background = { color: C.navy };
  s.addShape('rect', { x: 0, y: 0, w: 13.33, h: 2.0, fill: { color: C.navy2 } });
  s.addShape('rect', { x: 0.8, y: 2.55, w: 1.7, h: 0.09, fill: { color: C.gold } });
  s.addText('Charge-Creation Lending Intelligence', { x: 0.8, y: 2.75, w: 11.7, h: 1.0, color: C.white, fontSize: 38, bold: true });
  s.addText([
    { text: 'Northern Arc ', options: { color: C.white, fontSize: 20, bold: true } },
    { text: '(NACL) ', options: { color: '8AB4F8', fontSize: 20, bold: true } },
    { text: 'vs Other Lenders', options: { color: C.oth, fontSize: 20, bold: true } },
  ], { x: 0.8, y: 3.75, w: 11.7, h: 0.5 });
  s.addText('Who is creating new charges across the onboarded portfolio — by sector and over time.', { x: 0.8, y: 4.35, w: 11.0, h: 0.4, color: 'C7D4E6', fontSize: 13 });
  s.addText(`${fmtN(meta.entityCount)} Northern Arc onboarded entities    ·    Source: Saverisk (MCA charge data)    ·    ${meta.generated}`, { x: 0.8, y: 6.4, w: 11.7, h: 0.4, color: C.mut, fontSize: 11 });
  s.addText('Windows count back from the generation date. Amounts in ₹ Crore.', { x: 0.8, y: 6.75, w: 11.7, h: 0.3, color: C.mut, fontSize: 10, italic: true });

  // ---------- EXECUTIVE SUMMARY ----------
  s = header('Executive summary', 'New charge creation on the onboarded portfolio — Northern Arc vs all other lenders');
  const wins = A.naVsOthers, tileW = 12.5 / wins.length;
  wins.forEach((w, i) => {
    const x = 0.4 + i * tileW;
    s.addShape('roundRect', { x: x + 0.06, y: 1.3, w: tileW - 0.12, h: 2.0, fill: { color: C.soft }, line: { color: C.line, width: 1 }, rectRadius: 0.06 });
    s.addText(w.window.toUpperCase(), { x: x + 0.06, y: 1.42, w: tileW - 0.12, h: 0.3, color: C.mut, fontSize: 9, bold: true, align: 'center' });
    s.addText(fmtN(w.other_charges), { x: x + 0.06, y: 1.74, w: tileW - 0.12, h: 0.5, color: C.oth, fontSize: 25, bold: true, align: 'center' });
    s.addText('charges by others', { x: x + 0.06, y: 2.26, w: tileW - 0.12, h: 0.25, color: C.mut, fontSize: 8.5, align: 'center' });
    s.addText(fmtCr(w.other_amount_cr), { x: x + 0.06, y: 2.5, w: tileW - 0.12, h: 0.28, color: C.ink, fontSize: 10, bold: true, align: 'center' });
    s.addText(`NACL: ${fmtN(w.na_charges)} · ${fmtCr(w.na_amount_cr)}`, { x: x + 0.06, y: 2.86, w: tileW - 0.12, h: 0.3, color: C.na, fontSize: 8.5, align: 'center' });
  });
  const six = wins.find((w) => /6 month/i.test(w.window)) || wins[wins.length - 1];
  const ratio = six && six.na_charges ? Math.round(six.other_charges / six.na_charges) : null;
  s.addShape('roundRect', { x: 0.4, y: 3.6, w: 12.5, h: 1.15, fill: { color: C.navy }, rectRadius: 0.06 });
  s.addText([
    { text: 'Key takeaway   ', options: { color: C.gold, fontSize: 12, bold: true } },
    { text: ratio ? `Over ${six.window.toLowerCase()}, other lenders created ${fmtN(six.other_charges)} new charges (${fmtCr(six.other_amount_cr)}) on Northern Arc’s onboarded borrowers — about ${ratio}× Northern Arc’s own ${fmtN(six.na_charges)} (${fmtCr(six.na_amount_cr)}).` : `Over ${six ? six.window.toLowerCase() : 'the period'}, other lenders created ${fmtN(six ? six.other_charges : 0)} new charges on the onboarded portfolio.`, options: { color: C.white, fontSize: 12.5 } },
  ], { x: 0.7, y: 3.72, w: 11.9, h: 0.9, valign: 'middle' });
  s.addText('How this deck is organised', { x: 0.4, y: 4.95, w: 12, h: 0.3, color: C.ink, fontSize: 12, bold: true });
  s.addText([
    'One section per analysis; inside each, the same view for 1 month / 2 months / 3 months / 6 months.',
    'Sections: sector split · funded by others · funded by NACL · first-time funded · most active lenders · first-time lender→borrower · lenders new to the book.',
    'Final section: Northern Arc’s share of total charge creation per entity (highest & lowest).',
  ].map((t) => ({ text: t, options: { bullet: { code: '2022' }, color: C.ink, fontSize: 11 } })), { x: 0.7, y: 5.3, w: 12, h: 1.5, lineSpacingMultiple: 1.15 });

  // ---------- HEADLINE CHART ----------
  s = header('Charge creation: Northern Arc (NACL) vs Others', 'New charges created across the onboarded portfolio, by window', C.oth);
  const labels = A.naVsOthers.map((w) => w.window);
  s.addChart(p.ChartType.bar, [
    { name: 'Other lenders', labels, values: A.naVsOthers.map((w) => w.other_charges) },
    { name: 'Northern Arc (NACL)', labels, values: A.naVsOthers.map((w) => w.na_charges) },
  ], { x: 0.5, y: 1.2, w: 12.3, h: 5.6, barDir: 'col', barGapWidthPct: 40, chartColors: [C.oth, C.na], showLegend: true, legendPos: 'b', legendFontSize: 12, showValue: true, dataLabelFontSize: 11, dataLabelColor: C.ink, catAxisLabelFontSize: 13, valGridLine: { style: 'none' } });

  // ---------- SECTIONS (by analysis; each expands into 1m/2m/3m/6m) ----------
  for (const sec of SECTIONS) {
    divider(sec.title, sec.sub, sec.accent);
    for (const wk of PPTW) {
      const cs = header(`${sec.title} · ${wlabel(wk)}`, sec.sub, sec.accent);
      renderers[sec.key](cs, wk);
    }
  }

  // ---------- FINAL: NACL share of total charge creation, per entity ----------
  divider('Northern Arc share of total charge creation', 'NACL amount ÷ total charge amount per onboarded entity (all current open charges)', C.na);
  const shareCols = [
    { key: 'entity', label: 'Onboarded Entity' }, { key: 'sector', label: 'Nimbus Sector' },
    { key: 'nacl_amount_cr', label: 'NACL Amt', cr: true, right: true },
    { key: 'total_amount_cr', label: 'Total Amt', cr: true, right: true },
    { key: 'nacl_share_amount_pct', label: 'NACL Share %', right: true },
  ];
  const shareW = [4.6, 2.0, 1.9, 1.9, 2.1];
  const withCharges = (A.naclShare || []).filter((r) => r.total_charges > 0);
  s = header('Northern Arc share — HIGHEST', 'Where NACL is the dominant charge-holder', C.na);
  tbl(s, shareCols, withCharges, 18, shareW);
  s = header('Northern Arc share — LOWEST', 'Where other lenders dominate the onboarded borrower’s charges', C.oth);
  tbl(s, shareCols, [...withCharges].reverse(), 18, shareW);

  return p.writeFile({ fileName: file });
}

module.exports = { writePpt };
