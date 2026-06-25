// report_radar.js — Renewal Radar reporting module.
// Renders the per-month next-lender forecast into a self-contained HTML dashboard,
// an Excel workbook, and CSVs. Style mirrors the project's root output.js.
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// ---------- disclaimer (canonical, defined once) ----------
const DISCLAIMER = "Basis of estimation & disclaimer — These predictions extrapolate the historical lender↔borrower charge-creation cadence; they are indicative prospecting signals, not assured events. Predicted dates may vary by ~1–2 months (charge-filing / registration lag and deal timing). Whether a drawdown materialises — and from which lender — depends on the entity's evolving credit rating and capital position: a stronger balance sheet or an upgraded rating can move the borrower to lower-cost lenders (e.g. large banks), away from the NACL-competible set; a weaker position can do the reverse. Amounts are a secured-charge proxy, not sanctioned loan values. Validate each lead against Northern Arc's internal rating and capital-position view before acting.";
const DISCLAIMER_SHORT = "Indicative extrapolation — see Disclaimer sheet. Dates ±1–2 months; may not materialise if rating/capital position changes.";

// ---------- format helpers ----------
const esc = (s) => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const num = (n) => (Number.isFinite(Number(n)) ? Number(n) : 0);
const fmtCr = (n) => '₹' + num(n).toLocaleString('en-IN', { maximumFractionDigits: 2 }) + ' Cr';
const fmtCrRaw = (n) => Math.round(num(n) * 100) / 100; // numeric, 2dp — for Excel/CSV
const fmtN = (n) => num(n).toLocaleString('en-IN');
const fmtPct = (c) => {
  const v = Number(c);
  if (!Number.isFinite(v)) return '';
  return Math.round((v <= 1 ? v * 100 : v)) + '%';
};
const pctVal = (c) => {
  const v = Number(c);
  if (!Number.isFinite(v)) return 0;
  return Math.max(0, Math.min(100, Math.round(v <= 1 ? v * 100 : v)));
};
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function fmtDate(d) {
  if (!d) return '';
  const dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt.getTime())) return String(d);
  return `${String(dt.getUTCDate()).padStart(2, '0')} ${MONTHS[dt.getUTCMonth()]} ${dt.getUTCFullYear()}`;
}
function ymLabel(ym) {
  if (!ym) return '';
  const m = String(ym).match(/^(\d{4})-(\d{2})/);
  if (!m) return String(ym);
  return `${MONTHS[Number(m[2]) - 1] || m[2]} ${m[1]}`;
}
// NACL exposure cell for Excel/CSV: numeric ₹Cr, blank when null/0/invalid.
const expCell = (v) => (v == null || !Number.isFinite(Number(v)) || Number(v) === 0) ? '' : fmtCrRaw(v);
// exposure ₹Cr from a lifecycle entity's `exposure` (RUPEES); null when 0/invalid.
function expCrFromRupees(rupees) {
  const v = Number(rupees);
  if (!Number.isFinite(v) || v === 0) return null;
  return v / 1e7;
}
// build entity-name -> exposureCr (₹Cr) map from lifecycle.entities
function exposureMap(lifecycle) {
  const m = {};
  for (const e of ((lifecycle && lifecycle.entities) || [])) {
    m[e.entity || ''] = expCrFromRupees(e.exposure);
  }
  return m;
}

// ---------- row builders (shared by Excel + CSV) ----------
function forecastRows(forecast) {
  return (forecast || []).map((r) => ({
    'Month': ymLabel(r.ym),
    'Est. Date': fmtDate(r.date),
    'Entity': r.entity || '',
    'NACL Exposure (₹Cr)': expCell(r.naclExposureCr),
    'Sector': r.sector || '',
    'Likely Lender': r.lender || '',
    'Basis': r.basis || '',
    'Expected ₹Cr': fmtCrRaw(r.expectedAmtCr),
    'Confidence': fmtPct(r.confidence),
    'Observations': num(r.nObs),
    'Median Gap (mo)': r.medianGapMonths == null ? '' : num(r.medianGapMonths),
    'Last Drawn': fmtDate(r.lastDate),
  }));
}
function leadRows(leads) {
  return (leads || []).map((r) => ({
    'Month': ymLabel(r.ym),
    'Est. Date': fmtDate(r.date),
    'Entity': r.entity || '',
    'NACL Exposure (₹Cr)': expCell(r.naclExposureCr),
    'Sector': r.sector || '',
    'Likely Lender': r.lender || '',
    'Lender Tier': r.lenderTier || 'UNKNOWN',
    'Expected ₹Cr': fmtCrRaw(r.expectedAmtCr),
    'Confidence': fmtPct(r.confidence),
    'Observations': num(r.nObs),
    'Median Gap (mo)': r.medianGapMonths == null ? '' : num(r.medianGapMonths),
    'Last Drawn': fmtDate(r.lastDate),
  }));
}
function cadenceRows(lifecycle) {
  const out = [];
  for (const e of ((lifecycle && lifecycle.entities) || [])) {
    const expCr = expCrFromRupees(e.exposure);
    for (const c of (e.lenderCadence || [])) {
      out.push({
        'Entity': e.entity || '',
        'NACL Exposure (₹Cr)': expCell(expCr),
        'Sector': e.sector || '',
        'Lender': c.lender || '',
        'NACL?': c.isNACL ? 'Yes' : 'No',
        '#Raises': num(c.n),
        'First': fmtDate(c.firstDate),
        'Last': fmtDate(c.lastDate),
        'Median Gap (mo)': c.medianGapMonths == null ? '' : num(c.medianGapMonths),
        'Regularity': c.regularity == null ? '' : c.regularity,
        'Last ₹Cr': fmtCrRaw(c.lastAmtCr),
      });
    }
  }
  return out;
}
function monthlyRows(lifecycle, maxMonths = 24) {
  const out = [];
  for (const e of ((lifecycle && lifecycle.entities) || [])) {
    const expCr = expCrFromRupees(e.exposure);
    const months = (e.monthly || []).slice().sort((a, b) => String(a.ym).localeCompare(String(b.ym)));
    const recent = months.slice(-maxMonths);
    for (const m of recent) {
      out.push({
        'Entity': e.entity || '',
        'NACL Exposure (₹Cr)': expCell(expCr),
        'Month': m.ym || '',
        'Created #': num(m.createdCount),
        'Created ₹Cr': fmtCrRaw(m.createdCr),
        'Satisfied #': num(m.satisfiedCount),
        'Satisfied ₹Cr': fmtCrRaw(m.satisfiedCr),
        'Net ₹Cr': fmtCrRaw(m.netCr),
      });
    }
  }
  return out;
}

// ---------- Excel ----------
function writeExcel(lifecycle, forecast, leads, file) {
  const wb = XLSX.utils.book_new();
  const add = (name, rows) => {
    const data = (rows && rows.length) ? rows : [{ info: 'no rows' }];
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  };
  // sheet with a one-line disclaimer note row above the data headers
  const addWithNote = (name, rows) => {
    const data = (rows && rows.length) ? rows : [{ info: 'no rows' }];
    const ws = XLSX.utils.aoa_to_sheet([[DISCLAIMER_SHORT]]); // row 1 = note
    XLSX.utils.sheet_add_json(ws, data, { origin: 'A2' });      // headers on row 2, data below
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  };

  // dedicated Disclaimer sheet (full text wrapped across a few rows for readability)
  const discSheet = XLSX.utils.aoa_to_sheet([
    ['Renewal Radar — Basis of estimation & disclaimer'],
    [''],
    [DISCLAIMER],
    [''],
    ['How to read:'],
    ['• Predicted dates may vary by ~1–2 months (charge-filing / registration lag and deal timing).'],
    ['• "NACL-competible" leads = forecast events a non-bank/NACL-type lender is likely to win; validate against internal rating.'],
    ['• Amounts are a secured-charge proxy (₹Cr), not sanctioned loan values.'],
  ]);
  discSheet['!cols'] = [{ wch: 120 }];
  XLSX.utils.book_append_sheet(wb, discSheet, 'Disclaimer');

  addWithNote('Forecast Calendar', forecastRows(forecast));
  addWithNote('Leads (NACL-competible)', leadRows(leads));
  add('Entity Cadence', cadenceRows(lifecycle));
  add('Monthly Flow', monthlyRows(lifecycle, 24));
  XLSX.writeFile(wb, file);
}

// ---------- CSV ----------
function toCsv(rows) {
  if (!rows || !rows.length) return '';
  const cols = Object.keys(rows[0]);
  const escCell = (v) => {
    v = v == null ? '' : String(v);
    return /[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v;
  };
  const lines = [cols.map(escCell).join(',')];
  for (const r of rows) lines.push(cols.map((c) => escCell(r[c])).join(','));
  return lines.join('\n') + '\n';
}
function writeCsv(file, rows, headerCols) {
  if (!rows || !rows.length) {
    fs.writeFileSync(file, (headerCols || []).join(',') + '\n');
    return;
  }
  fs.writeFileSync(file, toCsv(rows));
}

// ---------- HTML helpers ----------
function htmlTable(cols, rows, opts = {}) {
  const max = opts.max || 1000;
  if (!rows || !rows.length) return '<p class="empty">No records.</p>';
  const head = '<tr>' + cols.map((c) => {
    const numeric = (c.cr || c.numCol) ? ' num' : '';
    return `<th class="sortable${numeric}" title="Click to sort">${esc(c.label)}</th>`;
  }).join('') + '</tr>';
  const body = rows.slice(0, max).map((r) => '<tr>' + cols.map((c) => {
    if (c.render) return c.render(r);
    let v = r[c.key];
    let raw = v;
    if (c.cr) { raw = num(v); v = fmtCr(v); }
    else if (c.numCol) { raw = num(v); v = fmtN(v); }
    else { raw = String(v == null ? '' : v); }
    return `<td${c.right ? ' style="text-align:right"' : ''} data-v="${esc(String(raw))}">${esc(v)}</td>`;
  }).join('') + '</tr>').join('');
  const more = rows.length > max ? `<tr class="morerow"><td colspan="${cols.length}" class="muted">…and ${rows.length - max} more (see Excel)</td></tr>` : '';
  return `<table>${head}${body}${more}</table>`;
}

function confCell(c) {
  const p = pctVal(c);
  return `<td data-v="${p}"><div class="confwrap"><div class="confbar" style="width:${p}%"></div><span>${p}%</span></div></td>`;
}

// NACL exposure HTML cell (₹Cr): "—" when null/0/invalid; numeric data-v for sorting.
function expCellHtml(v) {
  const raw = (v == null || !Number.isFinite(Number(v)) || Number(v) === 0) ? '' : num(v);
  const txt = raw === '' ? '—' : fmtCr(v);
  return `<td style="text-align:right" data-v="${esc(String(raw === '' ? 0 : raw))}">${esc(txt)}</td>`;
}

function tierBadge(tier) {
  const t = tier || 'UNKNOWN';
  const cls = t === 'UNKNOWN' ? 'tier unk' : 'tier';
  return `<span class="${cls}">${esc(t)}</span>`;
}

function writeHtml(lifecycle, forecast, leads, file, meta) {
  // build next-12-month window from asOf
  const asOf = (meta && meta.asOf) ? new Date(meta.asOf) : new Date();
  const base = isNaN(asOf.getTime()) ? new Date() : asOf;
  const monthsList = [];
  for (let i = 0; i < 12; i++) {
    const d = new Date(Date.UTC(base.getUTCFullYear(), base.getUTCMonth() + i, 1));
    const ym = `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
    monthsList.push(ym);
  }
  // months that actually have leads, in chronological order, restricted to next-12 first, then any others present
  const leadsByYm = {};
  for (const l of (leads || [])) {
    const ym = String(l.ym || '');
    (leadsByYm[ym] = leadsByYm[ym] || []).push(l);
  }
  const presentYms = Object.keys(leadsByYm).sort();
  const orderedYms = [...monthsList.filter((m) => leadsByYm[m]), ...presentYms.filter((m) => !monthsList.includes(m))];
  const tabYms = orderedYms.length ? orderedYms : monthsList.slice(0, 1);

  // KPIs
  const entityCount = ((lifecycle && lifecycle.entities) || []).length;
  const next12Set = new Set(monthsList);
  const forecastNext12 = (forecast || []).filter((r) => next12Set.has(String(r.ym))).length;
  const leadsCount = (leads || []).length;
  const leadsExpectedCr = (leads || []).reduce((s, r) => s + num(r.expectedAmtCr), 0);

  const leadCols = [
    { key: 'entity', label: 'Entity' },
    { key: 'naclExposureCr', label: 'NACL Exp. ₹Cr', numCol: true, right: true, render: (r) => expCellHtml(r.naclExposureCr) },
    { key: 'sector', label: 'Sector' },
    { key: 'lender', label: 'Likely Lender' },
    { key: 'tier', label: 'Tier', render: (r) => `<td data-v="${esc(r.lenderTier || 'UNKNOWN')}">${tierBadge(r.lenderTier)}</td>` },
    { key: 'expectedAmtCr', label: 'Expected ₹Cr', cr: true, right: true },
    { key: 'confidence', label: 'Confidence', render: (r) => confCell(r.confidence) },
    { key: 'lastDate', label: 'Last Drawn', render: (r) => `<td data-v="${esc(r.lastDate || '')}">${esc(fmtDate(r.lastDate))}</td>` },
  ];

  // per-month lead panes
  const def = 0;
  const tabBtns = tabYms.map((ym, i) =>
    `<button class="tab${i === def ? ' active' : ''}" data-t="lead-${esc(ym)}">${esc(ymLabel(ym))} <small>(${(leadsByYm[ym] || []).length})</small></button>`
  ).join('');
  const tabPanes = tabYms.map((ym, i) => {
    const rows = (leadsByYm[ym] || []).slice().sort((a, b) => pctVal(b.confidence) - pctVal(a.confidence));
    return `<div class="pane${i === def ? ' active' : ''}" id="lead-${esc(ym)}">${htmlTable(leadCols, rows, { max: 500 })}</div>`;
  }).join('');
  const leadTabs = (tabBtns && tabPanes)
    ? `<div class="tabs">${tabBtns}</div>${tabPanes}`
    : '<p class="empty">No NACL-competible leads forecast.</p>';

  // full forecast calendar (with text filter)
  const fcCols = [
    { key: 'ym', label: 'Month', render: (r) => `<td data-v="${esc(r.ym || '')}">${esc(ymLabel(r.ym))}</td>` },
    { key: 'date', label: 'Est. Date', render: (r) => `<td data-v="${esc(r.date || '')}">${esc(fmtDate(r.date))}</td>` },
    { key: 'entity', label: 'Entity' },
    { key: 'naclExposureCr', label: 'NACL Exp. ₹Cr', numCol: true, right: true, render: (r) => expCellHtml(r.naclExposureCr) },
    { key: 'sector', label: 'Sector' },
    { key: 'lender', label: 'Likely Lender', render: (r) => `<td data-v="${esc(r.lender || '')}"${r.isNACL ? ' class="na"' : ''}>${esc(r.lender || '')}</td>` },
    { key: 'basis', label: 'Basis' },
    { key: 'expectedAmtCr', label: 'Expected ₹Cr', cr: true, right: true },
    { key: 'confidence', label: 'Confidence', render: (r) => confCell(r.confidence) },
    { key: 'nObs', label: 'Obs', numCol: true, right: true },
    { key: 'medianGapMonths', label: 'Gap (mo)', render: (r) => `<td style="text-align:right" data-v="${num(r.medianGapMonths)}">${r.medianGapMonths == null ? '' : esc(num(r.medianGapMonths))}</td>` },
    { key: 'lastDate', label: 'Last Drawn', render: (r) => `<td data-v="${esc(r.lastDate || '')}">${esc(fmtDate(r.lastDate))}</td>` },
  ];
  const forecastTable = htmlTable(fcCols, forecast || [], { max: 3000 });

  const generated = fmtDate(base);

  // disclaimer HTML: bold the "… —" lead-in phrase, escape the rest
  const discIdx = DISCLAIMER.indexOf('—');
  const disclaimerHtml = (discIdx > -1)
    ? `<strong>${esc(DISCLAIMER.slice(0, discIdx + 1))}</strong>${esc(DISCLAIMER.slice(discIdx + 1))}`
    : esc(DISCLAIMER);

  const html = `<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Renewal Radar — Next-Lender Forecast (NACL)</title>
<style>
:root{--na:#0d6efd;--oth:#e8590c;--ok:#1f9d55;--bg:#0b1220;--card:#fff;--ink:#1a2233;--muted:#6b7785;--line:#e6e9ef}
*{box-sizing:border-box}body{margin:0;font:14px/1.5 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:var(--ink);background:#f4f6fa}
header{background:linear-gradient(120deg,#0b1220,#15315b);color:#fff;padding:26px 32px}
header h1{margin:0;font-size:22px;letter-spacing:.2px}
header .sub{opacity:.85;margin-top:6px;font-size:13px}
main{padding:24px 32px;max-width:1180px;margin:0 auto}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:8px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;box-shadow:0 1px 2px rgba(16,24,40,.04)}
.kpi-win{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
.kpi-main{font-size:26px;font-weight:700;margin:6px 0 2px;color:var(--ink)}
.kpi-main.lead{color:var(--na)}.kpi-main.cr{color:var(--ok)}
.kpi-sub{font-size:12px;color:var(--muted)}
section{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px 20px;margin:18px 0;box-shadow:0 1px 2px rgba(16,24,40,.04)}
section h2{margin:0 0 4px;font-size:16px}section .desc{color:var(--muted);font-size:12.5px;margin-bottom:12px}
.tabs{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:12px}
.tab{border:1px solid var(--line);background:#f7f9fc;border-radius:999px;padding:5px 13px;font-size:12.5px;cursor:pointer;color:var(--muted)}
.tab small{opacity:.7}
.tab.active{background:var(--ink);color:#fff;border-color:var(--ink)}
.pane{display:none}.pane.active{display:block}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th{text-align:left;color:var(--muted);font-weight:600;border-bottom:2px solid var(--line);padding:7px 8px;white-space:nowrap}
th.sortable{cursor:pointer;user-select:none}th.sortable:hover{color:var(--ink)}
th.sortable::after{content:'\\2195';font-size:9px;opacity:.35;margin-left:4px}
th[data-dir=asc]::after{content:'\\25B2';opacity:.9}th[data-dir=desc]::after{content:'\\25BC';opacity:.9}
td{border-bottom:1px solid var(--line);padding:7px 8px;vertical-align:middle}
tr:hover td{background:#fafbfe}
td.na{color:var(--na);font-weight:600}.muted,.empty{color:var(--muted)}.empty{padding:14px 4px}
.tier{display:inline-block;background:#eef4ff;color:#1d4ed8;border:1px solid #d6e4ff;border-radius:999px;padding:1px 9px;font-size:11px;font-weight:600}
.tier.unk{background:#f1f3f6;color:var(--muted);border-color:var(--line)}
.confwrap{position:relative;background:#eef1f6;border-radius:4px;height:18px;min-width:120px}
.confbar{position:absolute;left:0;top:0;height:18px;background:linear-gradient(90deg,#1d6fe0,#0d6efd);border-radius:4px}
.confwrap span{position:absolute;right:6px;top:0;line-height:18px;font-size:11px;color:var(--ink);font-weight:600}
.filterbox{margin-bottom:10px}
.filterbox input{width:100%;max-width:360px;padding:7px 10px;border:1px solid var(--line);border-radius:8px;font-size:13px}
.disclaimer{background:#fff8ec;border:1px solid #f5d99a;border-left:4px solid #e8950c;border-radius:10px;padding:12px 16px;margin:14px 0 4px;font-size:12.5px;line-height:1.55;color:#6b4e16}
.disclaimer strong{color:#9a5b00}
.disclaimer .tag{display:inline-block;background:#e8950c;color:#fff;font-size:10.5px;font-weight:700;letter-spacing:.4px;text-transform:uppercase;border-radius:4px;padding:1px 7px;margin-right:8px;vertical-align:middle}
footer{color:var(--muted);font-size:11.5px;text-align:center;padding:24px}
footer .footdisc{max-width:920px;margin:0 auto 12px;text-align:left;background:#fff8ec;border:1px solid #f5d99a;border-left:4px solid #e8950c;border-radius:10px;padding:11px 15px;color:#6b4e16;font-size:11.5px;line-height:1.55}
</style></head><body>
<header>
  <h1>Renewal Radar — predicted next lender per entity, by month</h1>
  <div class="sub">Forecasts which lender is likely to give each Northern Arc onboarded entity its next loan / renewal, so the NACL business team can approach first · as of ${esc(generated)} · ${esc(entityCount)} entities</div>
</header>
<main>
  <div class="kpis">
    <div class="kpi"><div class="kpi-win">Entities Tracked</div><div class="kpi-main">${fmtN(entityCount)}</div><div class="kpi-sub">with lifecycle history</div></div>
    <div class="kpi"><div class="kpi-win">Forecast Events · next 12m</div><div class="kpi-main">${fmtN(forecastNext12)}</div><div class="kpi-sub">predicted charge creations</div></div>
    <div class="kpi"><div class="kpi-win">NACL-competible Leads</div><div class="kpi-main lead">${fmtN(leadsCount)}</div><div class="kpi-sub">entities NACL can approach</div></div>
    <div class="kpi"><div class="kpi-win">Expected ₹Cr in Leads</div><div class="kpi-main cr">${fmtCr(leadsExpectedCr)}</div><div class="kpi-sub">addressable pipeline</div></div>
  </div>

  <div class="disclaimer"><span class="tag">Read this first</span>${disclaimerHtml}</div>

  <section>
    <h2>Leads by month — who to approach, and when</h2>
    <div class="desc">NACL-competible forecast events for the next 12 months. Pick a month tab; rows are sorted by confidence. <b>Click any column header to sort.</b></div>
    ${leadTabs}
  </section>

  <section>
    <h2>Full forecast calendar</h2>
    <div class="desc">Every predicted next-lender event across all entities and months (NACL likely-lender rows highlighted blue). Type to filter by entity, sector, lender, or month.</div>
    <div class="filterbox"><input id="fc-filter" type="text" placeholder="Filter forecast calendar — e.g. an entity, sector, or lender name…"></div>
    <div id="fc-wrap">${forecastTable}</div>
  </section>
</main>
<footer>
  <div class="footdisc">${disclaimerHtml}</div>
  Source: Saverisk (MCA charge data). Amounts in ₹ Crore. Forecast is heuristic (lender cadence × recency); confidence reflects observation count and gap regularity.
</footer>
<script>
// month tabs
document.querySelectorAll('.tab').forEach(function(b){b.addEventListener('click',function(){
  var id=b.dataset.t, group=b.parentElement;
  group.querySelectorAll('.tab').forEach(function(x){x.classList.remove('active');});
  b.classList.add('active');
  var sec=group.parentElement;
  sec.querySelectorAll('.pane').forEach(function(p){p.classList.toggle('active',p.id===id);});
});});
// text filter on the forecast calendar
(function(){
  var inp=document.getElementById('fc-filter');
  if(!inp)return;
  inp.addEventListener('input',function(){
    var q=inp.value.toLowerCase();
    var tbl=document.querySelector('#fc-wrap table');
    if(!tbl)return;
    [].slice.call(tbl.rows,1).forEach(function(r){
      if(r.classList.contains('morerow'))return;
      r.style.display = (!q || r.textContent.toLowerCase().indexOf(q)>-1) ? '' : 'none';
    });
  });
})();
// click-to-sort on any table header
document.querySelectorAll('table').forEach(function(tbl){
  var ths=[].slice.call(tbl.rows[0] ? tbl.rows[0].cells : []);
  ths.forEach(function(th,idx){
    if(!th.classList.contains('sortable'))return;
    th.addEventListener('click',function(){
      var numeric=th.classList.contains('num');
      var dir=th.getAttribute('data-dir')==='asc'?'desc':'asc';
      ths.forEach(function(t){t.removeAttribute('data-dir');});
      th.setAttribute('data-dir',dir);
      var rows=[].slice.call(tbl.rows,1).filter(function(r){return !r.classList.contains('morerow') && r.cells.length===ths.length;});
      var rest=[].slice.call(tbl.rows,1).filter(function(r){return r.classList.contains('morerow')||r.cells.length!==ths.length;});
      rows.sort(function(a,b){
        var av=a.cells[idx]?a.cells[idx].getAttribute('data-v'):'';
        var bv=b.cells[idx]?b.cells[idx].getAttribute('data-v'):'';
        var cmp=numeric?((parseFloat(av)||0)-(parseFloat(bv)||0)):String(av).localeCompare(String(bv));
        return dir==='asc'?cmp:-cmp;
      });
      var tb=tbl.tBodies[0]||tbl.rows[0].parentNode;
      rows.concat(rest).forEach(function(r){tb.appendChild(r);});
    });
  });
});
</script>
</body></html>`;
  fs.writeFileSync(file, html);
  return html.length;
}

// ---------- main entry ----------
function writeRadar(lifecycle, forecast, leads, outDir) {
  lifecycle = lifecycle || { entities: [] };
  forecast = forecast || [];
  leads = leads || [];
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const xlsxFile = path.join(outDir, 'Renewal_Radar.xlsx');
  const htmlFile = path.join(outDir, 'renewal_radar.html');

  writeExcel(lifecycle, forecast, leads, xlsxFile);
  const htmlBytes = writeHtml(lifecycle, forecast, leads, htmlFile, { asOf: lifecycle.asOf });

  writeCsv(path.join(outDir, 'forecast_calendar.csv'), forecastRows(forecast),
    ['Month', 'Est. Date', 'Entity', 'NACL Exposure (₹Cr)', 'Sector', 'Likely Lender', 'Basis', 'Expected ₹Cr', 'Confidence', 'Observations', 'Median Gap (mo)', 'Last Drawn']);
  writeCsv(path.join(outDir, 'leads.csv'), leadRows(leads),
    ['Month', 'Est. Date', 'Entity', 'NACL Exposure (₹Cr)', 'Sector', 'Likely Lender', 'Lender Tier', 'Expected ₹Cr', 'Confidence', 'Observations', 'Median Gap (mo)', 'Last Drawn']);
  writeCsv(path.join(outDir, 'entity_cadence.csv'), cadenceRows(lifecycle),
    ['Entity', 'NACL Exposure (₹Cr)', 'Sector', 'Lender', 'NACL?', '#Raises', 'First', 'Last', 'Median Gap (mo)', 'Regularity', 'Last ₹Cr']);
  writeCsv(path.join(outDir, 'monthly_flow.csv'), monthlyRows(lifecycle, 24),
    ['Entity', 'NACL Exposure (₹Cr)', 'Month', 'Created #', 'Created ₹Cr', 'Satisfied #', 'Satisfied ₹Cr', 'Net ₹Cr']);

  return {
    files: ['Renewal_Radar.xlsx', 'renewal_radar.html', 'forecast_calendar.csv', 'leads.csv', 'entity_cadence.csv', 'monthly_flow.csv']
      .map((f) => path.join(outDir, f)),
    htmlBytes,
  };
}

module.exports = { writeRadar };

// ---------- self-test ----------
if (require.main === module) {
  const asOf = '2026-06-25';
  const d = (s) => new Date(s + 'T00:00:00Z');
  const lifecycle = {
    asOf,
    entities: [
      {
        entity: 'Acme Finance Pvt Ltd', sector: 'SBL', exposure: 219000000, onboardedDate: '2023-04-01',
        openCharges: [{ lender: 'HDFC Bank', amountCr: 50, creationDate: d('2025-11-10'), isNACL: false, isTrustee: false }],
        monthly: [
          { ym: '2026-03', createdCount: 2, createdCr: 30, satisfiedCount: 1, satisfiedCr: 10, netCr: 20 },
          { ym: '2026-04', createdCount: 1, createdCr: 15, satisfiedCount: 0, satisfiedCr: 0, netCr: 15 },
        ],
        lenderCadence: [
          { lender: 'HDFC Bank', isNACL: false, n: 5, firstDate: d('2023-05-01'), lastDate: d('2025-11-10'), medianGapMonths: 6, regularity: 'high', lastAmtCr: 50, source: 'charge' },
          { lender: 'Northern Arc', isNACL: true, n: 2, firstDate: d('2024-02-01'), lastDate: d('2025-08-01'), medianGapMonths: 9, regularity: 'medium', lastAmtCr: 20, source: 'charge' },
        ],
        summary: { totalOpenCr: 70, openCount: 2, last12CreatedCr: 45, last12SatisfiedCr: 10, distinctActiveLenders: 2 },
      },
      {
        entity: 'Bharat Microfin Ltd', sector: 'MFI', exposure: 805000000, onboardedDate: '2022-09-15',
        openCharges: [{ lender: 'Catalyst Trusteeship', amountCr: 40, creationDate: d('2026-01-05'), isNACL: false, isTrustee: true }],
        monthly: [
          { ym: '2026-01', createdCount: 1, createdCr: 40, satisfiedCount: 0, satisfiedCr: 0, netCr: 40 },
          { ym: '2026-05', createdCount: 1, createdCr: 25, satisfiedCount: 1, satisfiedCr: 20, netCr: 5 },
        ],
        lenderCadence: [
          { lender: 'IDFC First Bank', isNACL: false, n: 4, firstDate: d('2023-01-01'), lastDate: d('2026-01-05'), medianGapMonths: 4, regularity: 'high', lastAmtCr: 40, source: 'charge' },
        ],
        summary: { totalOpenCr: 40, openCount: 1, last12CreatedCr: 65, last12SatisfiedCr: 20, distinctActiveLenders: 1 },
      },
      {
        entity: 'Coastal Vehicle Finance', sector: 'VF', exposure: 1500000000, onboardedDate: '2021-06-01',
        openCharges: [],
        monthly: [{ ym: '2026-02', createdCount: 1, createdCr: 60, satisfiedCount: 0, satisfiedCr: 0, netCr: 60 }],
        lenderCadence: [
          { lender: 'Northern Arc', isNACL: true, n: 3, firstDate: d('2023-03-01'), lastDate: d('2026-02-01'), medianGapMonths: 7, regularity: 'medium', lastAmtCr: 60, source: 'charge' },
          { lender: 'Kotak Mahindra Bank', isNACL: false, n: 6, firstDate: d('2022-07-01'), lastDate: d('2025-12-01'), medianGapMonths: 5, regularity: 'high', lastAmtCr: 35, source: 'charge' },
        ],
        summary: { totalOpenCr: 0, openCount: 0, last12CreatedCr: 60, last12SatisfiedCr: 0, distinctActiveLenders: 2 },
      },
      {
        entity: 'Deccan Housing Fin', sector: 'AHF', exposure: 0, onboardedDate: '2024-01-10',
        openCharges: [],
        monthly: [],
        lenderCadence: [
          { lender: 'SBI', isNACL: false, n: 1, firstDate: d('2025-09-01'), lastDate: d('2025-09-01'), medianGapMonths: null, regularity: 'low', lastAmtCr: 18, source: 'charge' },
        ],
        summary: { totalOpenCr: 0, openCount: 0, last12CreatedCr: 18, last12SatisfiedCr: 0, distinctActiveLenders: 1 },
      },
    ],
  };

  const mk = (ym, date, entity, sector, lender, isNACL, basis, amt, conf, nObs, gap, last, expCr) =>
    ({ ym, date: d(date), entity, sector, lender, isNACL, basis, expectedAmtCr: amt, confidence: conf, nObs, medianGapMonths: gap, lastDate: d(last), naclExposureCr: expCr });

  // mix of populated (e.g. 21.9) and null naclExposureCr
  const forecast = [
    mk('2026-07', '2026-07-12', 'Acme Finance Pvt Ltd', 'SBL', 'HDFC Bank', false, 'cadence', 52, 0.84, 5, 6, '2025-11-10', 21.9),
    mk('2026-07', '2026-07-20', 'Bharat Microfin Ltd', 'MFI', 'IDFC First Bank', false, 'cadence', 42, 0.78, 4, 4, '2026-01-05', 80.5),
    mk('2026-08', '2026-08-05', 'Coastal Vehicle Finance', 'VF', 'Kotak Mahindra Bank', false, 'cadence', 36, 0.71, 6, 5, '2025-12-01', 150),
    mk('2026-08', '2026-08-15', 'Acme Finance Pvt Ltd', 'SBL', 'Northern Arc', true, 'cadence', 22, 0.55, 2, 9, '2025-08-01', 21.9),
    mk('2026-09', '2026-09-02', 'Coastal Vehicle Finance', 'VF', 'Northern Arc', true, 'cadence', 60, 0.62, 3, 7, '2026-02-01', 150),
    mk('2026-09', '2026-09-18', 'Bharat Microfin Ltd', 'MFI', 'IDFC First Bank', false, 'cadence', 28, 0.66, 4, 4, '2026-01-05', 80.5),
    mk('2026-10', '2026-10-01', 'Acme Finance Pvt Ltd', 'SBL', 'HDFC Bank', false, 'cadence', 50, 0.80, 5, 6, '2025-11-10', 21.9),
    mk('2026-10', '2026-10-22', 'Deccan Housing Fin', 'AHF', 'SBI', false, 'recency', 18, 0.40, 1, null, '2025-09-01', null),
    mk('2026-11', '2026-11-09', 'Coastal Vehicle Finance', 'VF', 'Kotak Mahindra Bank', false, 'cadence', 35, 0.69, 6, 5, '2025-12-01', 150),
    mk('2026-12', '2026-12-03', 'Bharat Microfin Ltd', 'MFI', 'IDFC First Bank', false, 'cadence', 44, 0.74, 4, 4, '2026-01-05', 80.5),
  ];

  const tierOf = { 'HDFC Bank': 'AA-', 'IDFC First Bank': 'A', 'Kotak Mahindra Bank': 'AA-', 'SBI': 'UNKNOWN' };
  const leads = forecast.filter((r) => !r.isNACL).slice(0, 6).map((r) => ({
    ym: r.ym, date: r.date, entity: r.entity, sector: r.sector, lender: r.lender,
    lenderTier: tierOf[r.lender] || 'UNKNOWN', expectedAmtCr: r.expectedAmtCr, confidence: r.confidence,
    nObs: r.nObs, medianGapMonths: r.medianGapMonths, lastDate: r.lastDate, naclExposureCr: r.naclExposureCr, naclCompetible: true,
  }));

  const outDir = path.join(__dirname, '_sample_out');
  const res = writeRadar(lifecycle, forecast, leads, outDir);
  console.log('Files written:');
  for (const f of res.files) {
    const ok = fs.existsSync(f);
    console.log('  ', ok ? 'OK' : 'MISSING', f, ok ? `(${fs.statSync(f).size} bytes)` : '');
  }
  console.log('HTML bytes:', res.htmlBytes);

  // also try empty inputs (must not crash)
  const emptyDir = path.join(__dirname, '_sample_out_empty');
  const r2 = writeRadar({ asOf, entities: [] }, [], [], emptyDir);
  console.log('Empty-input run OK, files:', r2.files.length);
}
