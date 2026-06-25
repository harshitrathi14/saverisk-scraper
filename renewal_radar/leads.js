// renewal_radar/leads.js — turn the raw forecast into a NACL-competible business lead list.
// A lead = an upcoming renewal/top-up where the likely incumbent lender is one NACL can plausibly
// outcompete on rate (similar/weaker credit tier, per ratings.js), so the business team can
// approach the entity BEFORE that lender renews.
const RT = require('./ratings.js');

// forecastRows: array from predict.buildForecast. Returns the competible subset, enriched.
function buildLeads(forecastRows) {
  const rows = (forecastRows || [])
    // only rows tied to an identifiable incumbent lender (entity-rhythm rows have no lender to pre-empt)
    .filter((r) => r && r.basis === 'recurring-cadence' && !r.isNACL && RT.isCompetible(r.lender))
    .map((r) => ({
      ym: r.ym,
      date: r.date,
      entity: r.entity,
      sector: r.sector,
      lender: r.lender,
      lenderTier: RT.tierOf(r.lender).tier,
      naclExposureCr: r.naclExposureCr != null ? r.naclExposureCr : null,
      expectedAmtCr: r.expectedAmtCr,
      confidence: r.confidence,
      nObs: r.nObs,
      medianGapMonths: r.medianGapMonths,
      lastDate: r.lastDate,
      naclCompetible: true,
    }));
  rows.sort((a, b) => (a.ym < b.ym ? -1 : a.ym > b.ym ? 1 : b.confidence - a.confidence));
  return rows;
}

module.exports = { buildLeads };
