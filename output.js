// output.js — write the analyses to an Excel workbook + a self-contained HTML dashboard.
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// ---------- Excel ----------
function writeExcel(A, file) {
  const wb = XLSX.utils.book_new();
  const add = (name, rows) => {
    const data = (rows && rows.length) ? rows : [{ info: 'no rows' }];
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  };
  // headline NA vs Others
  add('NA vs Others', A.naVsOthers);
  for (const w of A.WINDOWS) add(`NAvsOther by Sector ${w.key}`, A.naVsOthersBySector[w.key]);
  for (const w of A.WINDOWS) add(`Funded by Others ${w.key}`, A.externalCharges[w.key]);
  for (const w of A.WINDOWS) add(`Funded by Northern Arc ${w.key}`, A.naCharges[w.key]);
  for (const w of A.WINDOWS) add(`Active Lenders ${w.key}`, A.activeLenders[w.key]);
  for (const w of A.WINDOWS) add(`First-Time Lenders ${w.key}`, A.firstTime[w.key]);
  add('Entity Summary', A.summaryRows);
  add('Latest Charge', A.latestRows);
  add('New Since Last Run', A.newSinceRun);
  add('All Charges', A.allCharges);
  XLSX.writeFile(wb, file);
}

// ---------- HTML helpers ----------
const esc = (s) => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const fmtN = (n) => (Number(n) || 0).toLocaleString('en-IN');
const fmtCr = (n) => '₹' + (Number(n) || 0).toLocaleString('en-IN', { maximumFractionDigits: 1 }) + ' Cr';

function table(cols, rows, opts = {}) {
  if (!rows || !rows.length) return '<p class="empty">No records.</p>';
  const max = opts.max || 100;
  const head = '<tr>' + cols.map((c) => `<th>${esc(c.label)}</th>`).join('') + '</tr>';
  const body = rows.slice(0, max).map((r) => '<tr>' + cols.map((c) => {
    let v = r[c.key];
    if (c.cr) v = fmtCr(v); else if (c.num) v = fmtN(v);
    const cls = c.na && (r[c.key] === 'Yes') ? ' class="na"' : '';
    return `<td${cls}${c.right ? ' style="text-align:right"' : ''}>${esc(v)}</td>`;
  }).join('') + '</tr>').join('');
  const more = rows.length > max ? `<tr><td colspan="${cols.length}" class="muted">…and ${rows.length - max} more (see Excel)</td></tr>` : '';
  return `<table>${head}${body}${more}</table>`;
}

// grouped horizontal bars for NA vs Others by sector (CSS only, offline-safe)
function sectorBars(rows) {
  if (!rows || !rows.length) return '<p class="empty">No charge activity in this window.</p>';
  const top = rows.slice(0, 12);
  const maxV = Math.max(1, ...top.map((r) => Math.max(r.na_charges, r.other_charges)));
  return '<div class="bars">' + top.map((r) => `
    <div class="barrow">
      <div class="barlabel" title="${esc(r.sector)}">${esc(r.sector)}</div>
      <div class="bargroup">
        <div class="bar na" style="width:${(r.na_charges / maxV * 100).toFixed(1)}%">${r.na_charges ? r.na_charges : ''}</div>
        <div class="bar oth" style="width:${(r.other_charges / maxV * 100).toFixed(1)}%">${r.other_charges ? r.other_charges : ''}</div>
      </div>
    </div>`).join('') + '</div>';
}

function kpiRow(A) {
  return A.naVsOthers.map((w) => `
    <div class="kpi">
      <div class="kpi-win">${esc(w.window)}</div>
      <div class="kpi-main"><span class="oth-t">${fmtN(w.other_charges)}</span> <small>charges by others</small></div>
      <div class="kpi-sub">${fmtCr(w.other_amount_cr)} · ${fmtN(w.other_borrowers)} borrowers</div>
      <div class="kpi-na">Northern Arc: ${fmtN(w.na_charges)} charges · ${fmtCr(w.na_amount_cr)}</div>
    </div>`).join('');
}

function windowTabs(A, builder, idPrefix) {
  // default to "Last 3 Months" (index 2) — 1-week/1-month are sparse due to MCA filing lag
  const def = Math.min(2, A.WINDOWS.length - 1);
  const btns = A.WINDOWS.map((w, i) => `<button class="tab${i === def ? ' active' : ''}" data-t="${idPrefix}-${w.key}">${esc(w.label)}</button>`).join('');
  const panes = A.WINDOWS.map((w, i) => `<div class="pane${i === def ? ' active' : ''}" id="${idPrefix}-${w.key}">${builder(w)}</div>`).join('');
  return `<div class="tabs">${btns}</div>${panes}`;
}

function writeHtml(A, file, meta) {
  const chargeCols = [
    { key: 'entity', label: 'Entity' }, { key: 'sector', label: 'Sector' },
    { key: 'exposure_cr', label: 'NA Exposure', cr: true, right: true },
    { key: 'lender', label: 'Lender' },
    { key: 'amount_cr', label: 'Amount', cr: true, right: true },
    { key: 'charge_date', label: 'Charge Date' }, { key: 'type', label: 'Type' },
  ];
  const fundedTabs = windowTabs(A, (w) => table(chargeCols, A.externalCharges[w.key], { max: 100 }), 'funded');
  const naFundedTabs = windowTabs(A, (w) => table(chargeCols, A.naCharges[w.key], { max: 100 }), 'nafunded');

  const lenderTabs = windowTabs(A, (w) => table([
    { key: 'lender', label: 'Lender' }, { key: 'is_na', label: 'NA?', na: true },
    { key: 'charges_created', label: 'Charges', num: true, right: true },
    { key: 'total_amount_cr', label: 'Amount', cr: true, right: true },
    { key: 'distinct_borrowers', label: 'Borrowers', num: true, right: true },
  ], A.activeLenders[w.key], { max: 25 }), 'lenders');

  const firstTabs = windowTabs(A, (w) => table([
    { key: 'entity', label: 'Borrower' }, { key: 'sector', label: 'Sector' },
    { key: 'lender', label: 'First-time Lender' }, { key: 'is_na', label: 'NA?', na: true },
    { key: 'amount_cr', label: 'Amount', cr: true, right: true }, { key: 'charge_date', label: 'Date' },
  ], A.firstTime[w.key], { max: 50 }), 'first');

  const sectorTabs = windowTabs(A, (w) => sectorBars(A.naVsOthersBySector[w.key]), 'sector');

  const html = `<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Saverisk Lending Intelligence — Northern Arc vs Others</title>
<style>
:root{--na:#0d6efd;--oth:#e8590c;--bg:#0b1220;--card:#fff;--ink:#1a2233;--muted:#6b7785;--line:#e6e9ef}
*{box-sizing:border-box}body{margin:0;font:14px/1.5 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:var(--ink);background:#f4f6fa}
header{background:linear-gradient(120deg,#0b1220,#15315b);color:#fff;padding:26px 32px}
header h1{margin:0;font-size:22px;letter-spacing:.2px}
header .sub{opacity:.8;margin-top:6px;font-size:13px}
.legend{margin-top:10px;font-size:12px}.legend span{display:inline-block;margin-right:14px}
.dot{display:inline-block;width:10px;height:10px;border-radius:2px;vertical-align:middle;margin-right:5px}
.dot.na{background:var(--na)}.dot.oth{background:var(--oth)}
main{padding:24px 32px;max-width:1180px;margin:0 auto}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:8px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;box-shadow:0 1px 2px rgba(16,24,40,.04)}
.kpi-win{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
.kpi-main{font-size:26px;font-weight:700;margin:6px 0 2px}.kpi-main .oth-t{color:var(--oth)}.kpi-main small{font-size:12px;font-weight:500;color:var(--muted)}
.kpi-sub{font-size:12px;color:var(--ink)}.kpi-na{font-size:12px;color:var(--na);margin-top:6px;border-top:1px dashed var(--line);padding-top:6px}
section{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px 20px;margin:18px 0;box-shadow:0 1px 2px rgba(16,24,40,.04)}
section h2{margin:0 0 4px;font-size:16px}section .desc{color:var(--muted);font-size:12.5px;margin-bottom:12px}
.tabs{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:12px}
.tab{border:1px solid var(--line);background:#f7f9fc;border-radius:999px;padding:5px 13px;font-size:12.5px;cursor:pointer;color:var(--muted)}
.tab.active{background:var(--ink);color:#fff;border-color:var(--ink)}
.pane{display:none}.pane.active{display:block}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th{text-align:left;color:var(--muted);font-weight:600;border-bottom:2px solid var(--line);padding:7px 8px;white-space:nowrap}
td{border-bottom:1px solid var(--line);padding:7px 8px;vertical-align:top}
tr:hover td{background:#fafbfe}
td.na{color:var(--na);font-weight:600}.muted,.empty{color:var(--muted)}.empty{padding:14px 4px}
.bars{display:flex;flex-direction:column;gap:9px}
.barrow{display:flex;align-items:center;gap:10px}
.barlabel{width:210px;font-size:12px;color:var(--ink);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.bargroup{flex:1;display:flex;flex-direction:column;gap:3px}
.bar{height:15px;border-radius:3px;color:#fff;font-size:10px;line-height:15px;padding:0 5px;min-width:14px;white-space:nowrap}
.bar.na{background:var(--na)}.bar.oth{background:var(--oth)}
footer{color:var(--muted);font-size:11.5px;text-align:center;padding:24px}
</style></head><body>
<header>
  <h1>Saverisk Lending Intelligence — Northern Arc vs Other Lenders</h1>
  <div class="sub">Charge-creation (MCA) comparison across the onboarded portfolio · ${esc(meta.entityCount)} entities · generated ${esc(meta.generated)}</div>
  <div class="legend"><span><i class="dot na"></i>Northern Arc</span><span><i class="dot oth"></i>Other lenders</span>${A.firstRun ? '<span style="opacity:.7">· first run = baseline (no "new since last run" yet)</span>' : ''}</div>
</header>
<main>
  <div class="kpis">${kpiRow(A)}</div>

  <section>
    <h2>Charge creation by sector — Northern Arc vs Others</h2>
    <div class="desc">Number of new charges created per sector in the selected window. Orange = competitor lending into your portfolio; blue = Northern Arc.</div>
    ${sectorTabs}
  </section>

  <section>
    <h2>Entities funded by other lenders</h2>
    <div class="desc">Every charge created by a non-Northern-Arc lender in the window — one row per lender per charge, with its own date.</div>
    ${fundedTabs}
  </section>

  <section>
    <h2>Entities funded by Northern Arc</h2>
    <div class="desc">Every charge created by Northern Arc in the window — one row per charge, with its own date.</div>
    ${naFundedTabs}
  </section>

  <section>
    <h2>Most active lenders</h2>
    <div class="desc">Lenders ranked by charges created in the window across your portfolio (Northern Arc flagged).</div>
    ${lenderTabs}
  </section>

  <section>
    <h2>First-time lender → borrower relationships</h2>
    <div class="desc">A lender creating its first ever charge on that borrower in the window — new money, not a top-up.</div>
    ${firstTabs}
  </section>
</main>
<footer>Source: Saverisk (MCA charge data). Amounts in ₹ Crore. Debenture-trustee-held charges show the trustee, not the underlying lender.</footer>
<script>
document.querySelectorAll('.tab').forEach(b=>b.addEventListener('click',()=>{
  const id=b.dataset.t, group=b.parentElement;
  group.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));
  b.classList.add('active');
  const sec=group.parentElement;
  sec.querySelectorAll('.pane').forEach(p=>p.classList.toggle('active',p.id===id));
}));
</script>
</body></html>`;
  fs.writeFileSync(file, html);
}

module.exports = { writeExcel, writeHtml };
