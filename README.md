# Saverisk Lending Intelligence Scraper

Compares **Northern Arc** wholesale lending vs other lenders using **MCA charge-creation**
data from Saverisk, across 1-week / 1-month / 3-month / 6-month windows, sector-wise.

For each onboarded entity it pulls all charges, tags each lender as **Northern Arc** vs
**Other**, and produces an Excel workbook, a self-contained HTML dashboard, and a PPT deck.

> Running the scraper uses **no AI** — it is plain browser automation (Playwright).

## Setup

```bash
npm install
npx playwright install chromium
```

Provide an input file `input.csv` with columns: `name, short_name, cin, exposure, sector`
(CIN drives exact matching).

## Usage

```bash
node login.js              # one-time: log in (mobile + OTP); session is saved locally
node scrape.js             # gentle, cache-backed run over all entities
node scrape.js --max-new 150   # fetch at most 150 new entities (batching)
node scrape.js --report-only   # rebuild outputs from cache, no scraping
node scrape.js --rating        # also scrape credit rating (slower, UI-based)
```

Outputs land in `output/<timestamp>/`: Excel, `dashboard.html`, `*.pptx`, CSVs.

## How it works

- **Match**: `AsyncService.aspx/SearchAddCompany` by CIN (resolves outdated CINs).
- **Charges**: `CmpAsyncDataService.aspx/ExecuteMethodStaticAsync` (`Open Charge`).
- **Sector/Industry**: same endpoint, `overview` dashboard.
- **Rating** (opt-in): UI-driven Credit Rating tab.
- Searches run **sequentially** (the server serializes search per session).
- Incremental: `state.json` diffs charge-IDs to report new lending since the last run.

## Files

| File | Purpose |
|------|---------|
| `login.js` | Launch browser for one-time OTP login; saves session |
| `scrape.js` | Main scraper + report orchestration |
| `lib.js` | Saverisk API helpers, matching, parsing |
| `reports.js` | Builds the NA-vs-Others analyses |
| `output.js` | Excel workbook + HTML dashboard |
| `ppt.js` | PowerPoint deck |

## Notes

Session cookies (`.session/`, `storageState.json`) and all portfolio data
(`input.csv`, `charges_cache.json`, `output/`) are git-ignored — never commit them.
