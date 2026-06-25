// renewal_radar/email.js — generate a ready-to-send email (Markdown) presenting the Renewal Radar
// calendar: estimated future fundings to the onboarded entities, per month, with the competible leads.
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const fmtCr = (n) => '₹' + (Number(n) || 0).toLocaleString('en-IN', { maximumFractionDigits: 1 }) + ' Cr';
const fmtPct = (n) => Math.round((Number(n) || 0) * 100) + '%';
const ymLabel = (ym) => { if (!ym) return ''; const [y, m] = ym.split('-'); return `${MONTHS[(+m) - 1]} ${y}`; };
const fmtDate = (iso) => { if (!iso) return ''; const d = new Date(iso); return isNaN(d) ? '' : `${String(d.getUTCDate()).padStart(2, '0')} ${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}`; };
const exp = (n) => (n == null ? '—' : fmtCr(n));

const DISCLAIMER = "These predictions extrapolate the historical lender↔borrower charge-creation cadence; they are indicative prospecting signals, not assured events. Predicted dates may vary by ~1–2 months (charge-filing / registration lag and deal timing). Whether a drawdown materialises — and from which lender — depends on the entity's evolving credit rating and capital position: a stronger balance sheet or an upgraded rating can move the borrower to lower-cost lenders (e.g. large banks), away from the NACL-competible set; a weaker position can do the reverse. Amounts are a secured-charge proxy, not sanctioned loan values. Validate each lead against Northern Arc's internal rating and capital-position view before acting.";

function buildEmail(lifecycle, forecast, leads, meta) {
  meta = meta || {};
  const L = leads || [];
  const totalCr = L.reduce((a, l) => a + (Number(l.expectedAmtCr) || 0), 0);
  const entities = new Set(L.map((l) => l.entity)).size;

  // per-month rollup (chronological)
  const months = {};
  for (const l of L) {
    const m = (months[l.ym] = months[l.ym] || { n: 0, cr: 0, ent: new Set(), top: [] });
    m.n++; m.cr += Number(l.expectedAmtCr) || 0; m.ent.add(l.entity); m.top.push(l);
  }
  const ymSorted = Object.keys(months).sort();
  const calRows = ymSorted.slice(0, 6).map((ym) => {
    const m = months[ym];
    const top3 = m.top.sort((a, b) => (b.confidence || 0) - (a.confidence || 0)).slice(0, 3)
      .map((l) => `${l.entity.replace(/\s+(private|limited|ltd\.?|pvt\.?)\b/gi, '').trim()} (${l.lender.replace(/\s+(private|limited|ltd\.?)\b/gi, '').trim()})`).join('; ');
    return `| ${ymLabel(ym)} | ${m.n} | ${m.ent.size} | ${fmtCr(m.cr)} | ${top3} |`;
  }).join('\n');

  // top near-term leads (highest confidence)
  const topLeads = [...L].sort((a, b) => (b.confidence || 0) - (a.confidence || 0)).slice(0, 12)
    .map((l) => `| ${ymLabel(l.ym)} | ${l.entity} | ${exp(l.naclExposureCr)} | ${l.lender} | ${exp(l.expectedAmtCr)} | ${fmtPct(l.confidence)} | ${fmtDate(l.lastDate)} |`).join('\n');

  return `Subject: Renewal Radar — estimated upcoming fundings to our onboarded entities (next 12 months)

Hi team,

Sharing the first cut of the **Renewal Radar** — a forward calendar of when our onboarded entities are likely to take their next loan / renewal / top-up over the coming 12 months, and from which lender. The intent is simple: **reach the entity before the incumbent lender renews**, on relationships where Northern Arc can realistically compete on rate.

How it's built: we learn each entity↔lender charge-creation cadence from ~5 years of Saverisk (MCA) charge history — including loans that have since been satisfied (repaid) — and project the next likely drawdown. We then keep only the **NACL-competible** incumbents (similar / weaker credit tier; the AAA banks we can't out-price are screened out).

**Headline (NACL-competible leads, next 12 months)**
- ${L.length} predicted renewal/top-up leads across ${entities} onboarded entities
- ~${fmtCr(totalCr)} of estimated funding in play (secured-charge proxy)
- ${meta.satisfactionCount ? meta.satisfactionCount.toLocaleString('en-IN') + ' historical repayments analysed across ' : 'Across '}${(meta.entityCount || (lifecycle.entities || []).length)} entities

**Estimated funding calendar — next 6 months**

| Month | Leads | Entities | Est. ₹ (proxy) | Sample entities (likely lender) |
|---|---:|---:|---:|---|
${calRows || '| — | 0 | 0 | ₹0 Cr | — |'}

**Top near-term leads to action first** (highest confidence)

| Month | Onboarded Entity | NACL Exposure | Likely Lender (incumbent) | Est. ₹ | Confidence | Last Drawn |
|---|---|---:|---|---:|---:|---|
${topLeads || '| — | — | — | — | — | — | — |'}

The full month-by-month calendar (every entity, sortable, with NACL exposure) is in the dashboard — **renewal_radar.html** — and the **Renewal_Radar.xlsx** workbook (“Leads” and “Forecast Calendar” sheets). The combined deck (Saverisk_Lending_Intelligence_Deck.pptx) carries the summary.

**Please read before acting —** ${DISCLAIMER}

Happy to walk through any specific entity or sector.

Best regards,
[Name]
Northern Arc
`;
}

module.exports = { buildEmail };
