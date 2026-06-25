# Renewal Radar

Predicts, for each onboarded entity, **when its next loan / renewal / top-up is likely due and from
which lender** over the next 12 months — and flags the ones where the likely incumbent lender is
**NACL-competible** (similar/weaker credit tier), so the business team can approach the entity
*before* that lender renews. Separate from the weekly NACL-vs-others pipeline; writes to
`output_radar/<timestamp>/` and touches none of the existing outputs.

## Run
```bash
# 1) one-time / weekly: harvest full charge history (open + closed/satisfied, 5y) — needs login
node login.js                       # mobile + OTP
node harvest_history.js             # -> charge_history_cache.json  (~20-30 min, resumable)

# 2) build the radar (instant, no browser — pure analysis on the caches)
node renewal_radar/radar.js         # -> output_radar/<ts>/renewal_radar.html + .xlsx + CSVs
node renewal_radar/radar.js --horizon 18   # longer forecast window
```

## Data sources & how the signal is built
| Signal | Source | File |
|---|---|---|
| *Who* lends now / live exposure | **Open charges** (real lender names + CIN) | `charges_cache.json` (from `scrape.js`) |
| *When* — raise vs repay rhythm | **Charge History** (Creation + **Satisfaction** events, 5y, incl. closed loans) | `charge_history_cache.json` (from `harvest_history.js`) |
| *Winnable?* — lender credit tier | **`lender_ratings.json`** (populate from NACL internal) | `renewal_radar/lender_ratings.json` |

Pipeline: `lifecycle.js` (per-entity cadence + monthly created/satisfied flow) → `predict.js`
(forward 12-month calendar with confidence) → `leads.js` (NACL-competible filter via `ratings.js`)
→ `report_radar.js` (Excel + self-contained HTML + CSVs).

## Honest limitations (read before trusting a number)
- **Charge amount ≠ loan amount.** Charges are over-secured (≈1.1–1.25× cover) and a charge can
  secure a multi-tranche facility, so `Expected ₹Cr` is an *order-of-magnitude* estimate, not a quote.
- **No exact per-loan tenure.** In the Charge History view `Charge_ID` is a placeholder/blank ~half
  the time, so a repayment can't be linked to its originating drawdown. We use **aggregate** create-vs-satisfy
  flow, not per-loan maturity.
- **Lender attribution is partial in history.** ~⅔ of history creation events bucket the lender as
  `"Others"`. So per-lender cadence is anchored on **named** lenders (banks + the larger NBFCs) and on
  the open-charge names; long-tail NBFCs may be under-counted. Confidence scores reflect this.
- **Predictions are probabilistic.** Strong for steady recurring relationships (e.g. MAS←HDFC, ~5-month
  cadence, n≥5); low-confidence for sparse/`Others`-dominated entities. Every row carries a confidence.
- **Ratings are NOT from Saverisk.** Saverisk exposes no lender ratings on this account (every rating
  endpoint returns 0 rows). Until `lender_ratings.json` is populated, the filter only hard-excludes the
  unambiguous AAA majors (see `ratings.js` → `SEED_AAA`); everything else is tier `UNKNOWN` and shown as
  a *candidate* lead — including some genuinely high-grade NBFCs (e.g. Poonawalla/Shriram) that proper
  ratings would screen out. **Populate the map to activate true AA-/A/BBB tiering.**
- **No D/E feasibility gate yet.** Saverisk has no net-worth/borrowings (only nominal paid-up share
  capital, which ≠ equity). The feasibility gate ("can the entity absorb more debt?") should join
  NACL's own internal counterparty financials by CIN — not scraped. Hook left in `leads.js`.

## Activating the two NACL-internal inputs (both keyed by CIN / lender name)
1. **Lender ratings** → create `renewal_radar/lender_ratings.json`:
   ```json
   { "vivriti finance limited": "A+", "incred financial services limited": "A+",
     "poonawalla fincorp limited": "AAA", "shriram finance limited": "AA+" }
   ```
   Keys are lender names **lowercased**; tiers from `ratings.js → TIER_ORDER`. Competible cutoff =
   `NACL_TIER` (`AA-`) and weaker. Edit `ratings.js` to change the benchmark.
2. **D/E feasibility** → export `CIN → {net worth, total borrowings, D/E}` from onboarding/monitoring
   and join in `leads.js` to gate/rank leads by headroom.

## Files
`ratings.js` lender tiers + competible test · `lifecycle.js` cadence + flow · `predict.js` forecast ·
`leads.js` competible filter · `report_radar.js` outputs · `radar.js` orchestrator.
Intermediates `_lifecycle.json` / `_forecast.json` are debug dumps (git-ignored).
