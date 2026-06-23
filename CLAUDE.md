# CLAUDE.md — Saverisk Lending-Intelligence Scraper

Project memory for future sessions. Read this first.

## What this is
A Node.js tool that compares **Northern Arc (NACL)** wholesale lending vs **other lenders**,
using **MCA charge-creation** data scraped from **Saverisk** (saverisk.com), for a fixed list of
**Northern Arc onboarded entities**. It produces an Excel workbook, a self-contained HTML
dashboard, and a PowerPoint deck — all built around the NACL-vs-others comparison, sector-wise,
over time windows.

> Running the scraper uses **no AI / zero tokens** — it is plain Playwright browser automation.

## Run it
```bash
node login.js                  # one-time: opens a browser; log in with mobile + OTP (auto-saves session)
node scrape.js                 # gentle, cache-backed run; only fetches uncached entities
node scrape.js --report-only   # rebuild ALL outputs from cache, no browser/session (instant)
node scrape.js --max-new 150   # fetch at most 150 new entities, then stop (batching across sessions)
node scrape.js --refresh       # ignore cache, re-fetch everything
node scrape.js --rating        # also scrape credit rating (slow, UI-based; off by default)
node scrape.js --fast          # shorter delays (less "gentle")
```
Outputs land in `output/<timestamp>/`: `charge_creation_dashboard.html`, Excel, `*.pptx`,
per-window CSVs, `all_charges.csv`, and a `charges_cache.json` snapshot.

## Input
`input.csv` (regenerated from `CIN_exposure_onboarded_nimbus entities.xlsx`) with columns:
`name, short_name, cin, exposure, sector, onboarded_date`. CIN drives exact matching.
The `sector` is Northern Arc's internal **Nimbus Sector** (e.g. CORP, SBL, VF, MFI, AHF, BD,
CONS, AGRI, SFB). Blank/"NA" sector → labelled "Zero Exposure".

## Time windows
`reports.js → WINDOWS`: 1w / 1m / 2m / 3m / 6m. Each counts **back from the run date** (NOW).
To add/remove a window, edit that array — everything else (dashboard tabs, Excel sheets, CSVs,
PPT) loops over it automatically.

## Architecture / files
| File | Role |
|------|------|
| `login.js` | Launch headed browser; auto-detect login; save `.session/` + `storageState.json` |
| `lib.js` | Saverisk API helpers, name matching, date/HTML parsing |
| `reports.js` | `buildAnalyses()` — all the NACL-vs-others analyses (pure functions) |
| `output.js` | Excel workbook + self-contained HTML dashboard |
| `ppt.js` | Polished PowerPoint deck (cover, exec summary, per-window sections) |
| `scrape.js` | Orchestrates: match → fetch → cache → analyse → write outputs |
| `preview.js` | Render the HTML to `dashboard_preview.png` for a visual check |

## Saverisk internals (reverse-engineered)
- **Login**: `flogin.aspx`, **mobile + OTP** (also username/password but that has reCAPTCHA —
  avoid). Session cookie `api_session_xx` has `expires:-1` (session cookie) so the persistent
  profile drops it on close → `scrape.js` re-injects cookies from `storageState.json` on launch.
  Session lifetime looks ~time-based (~90 min); a fresh login completes a full run in one pass.
- **Match**: `POST AsyncService.aspx/SearchAddCompany` `{searchstr, type:'Company', id:userid}`.
  Resolves outdated/invalid CINs to the correct company ("Invalid CIN of <name>") — accept the
  top result and use its corrected CIN. **Searches MUST be sequential** — concurrent calls return
  empty (server serializes search per session).
- **Charges**: `POST CmpAsyncDataService.aspx/ExecuteMethodStaticAsync` `dashboardurl:'Open Charge'`
  → `{Table:[{Charge_Holder, 'Charge_Amount(Rs Cr)', Date_of_Creation, Charge_Id, HIDE_CIN, ...}]}`.
  Paginates by `pgno`; dedupe by `Charge_Id` (50/page).
- **Sector/Industry**: same endpoint, `dashboardurl:'overview'` → `Sector`, `Industry`.
- **Rating** (opt-in, slow): NOT the JSON API — must drive the UI (`fetchRatingUI`): open the
  company page, click the Credit Rating tab, read `#table_data`.

## Classification & key analyses
- **NACL vs Other**: a charge is NACL if `Charge_Holder` matches `/northern arc/i`.
- **Debenture-trustee-held** charges (Catalyst/Beacon/Vistra/Axis/IDBI) hide the true lender —
  flagged as `Debenture/Trustee` type.
- Analyses (all per window): NACL-vs-others overall + by sector; entities funded by other lenders
  (one row per charge); entities funded by NACL; entities funded for the FIRST time (earliest-ever
  charge); most active lenders; first-time lender→borrower; lenders new to the whole book; entity
  summary; new-since-last-run (incremental diff via `state.json`).

## State / caching
- `charges_cache.json` — full per-entity scraped data; resume skips cached entities (zero requests).
- `approved_matches.json` — entity→Saverisk hash/CIN, so matching isn't repeated.
- `state.json` — `lastRun` + per-entity charge-IDs for the "new since last run" diff. `lastRun`
  is only stamped on a **complete** run; an aborted/partial run stays resumable.

## Gotchas
- If `scrape.js` prints "Not logged in", re-run `login.js` (session expired).
- Concurrency breaks search — keep it sequential.
- Don't commit secrets/data: `.session/`, `storageState.json`, `input.csv`, `*.xlsx`,
  `charges_cache.json`, `output/` are git-ignored. The GitHub repo is **public** (code only).
- `Math.random()`/`Date.now()` are fine here (plain Node) — only restricted inside Workflow scripts.
