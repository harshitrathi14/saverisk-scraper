# CLAUDE.md — Saverisk Lending-Intelligence Scraper

Project memory for future sessions. Read this first.

## What this is
A Node.js tool that compares **Northern Arc (NACL)** wholesale lending vs **other lenders**,
using **MCA charge-creation** data scraped from **Saverisk** (saverisk.com), for a fixed list of
**Northern Arc onboarded entities**. It produces an Excel workbook, a self-contained HTML
dashboard, and a PowerPoint deck — all built around the NACL-vs-others comparison, sector-wise,
over time windows.

> Running the scraper uses **no AI / zero tokens** — it is plain Playwright browser automation.

## WEEKLY REFRESH — runbook (do this, minimal tokens)
The session expires (~time-based), so each weekly refresh starts with a fresh login.

1. **Update the input** (if the onboarded list/sectors changed): re-export
   `CIN_exposure_onboarded_nimbus entities.xlsx`, then regenerate `input.csv`:
   ```bash
   node -e 'const X=require("xlsx"),fs=require("fs");const wb=X.readFile("CIN_exposure_onboarded_nimbus entities.xlsx",{cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];const rows=X.utils.sheet_to_json(ws,{defval:""});const esc=v=>{v=v==null?"":String(v);return /[",\n]/.test(v)?`"${v.replace(/"/g,'""')}"`:v};const out=["name,short_name,cin,exposure,sector,onboarded_date"];for(const r of rows){const od=r.onboardedDate instanceof Date?r.onboardedDate.toISOString().slice(0,10):(r.onboardedDate||"");out.push([r.name,r.short_name,r.cin,r.Exposure,r.Sector,od].map(esc).join(","))}fs.writeFileSync("input.csv",out.join("\n")+"\n");console.log("input.csv:",rows.length)'
   ```
2. **Log in** (one-time per refresh): `node login.js` → a browser opens on the desktop;
   log in with **mobile number + OTP** (NOT username/password — that has reCAPTCHA). It auto-saves.
   - If the window doesn't appear: `rm -f .session/Singleton*` and re-run. `pgrep` self-matches
     the command string — confirm a real browser via `ps -eo comm | grep '^chrome$'`.
3. **Refresh the data** (re-pull charges that are older than ~6 days; resumes cleanly):
   ```bash
   node scrape.js --max-age 6
   ```
   ~25–35 min at gentle pace. It diffs against last week and fills the "new since last run" report.
   - If it prints **SESSION EXPIRED** mid-run: just re-run `node login.js` then
     `node scrape.js --max-age 6` again — cache + freshness make it resume from where it stopped.
4. **Outputs** land in `output/<timestamp>/`. Open `charge_creation_dashboard.html`; share the
   Excel + `*.pptx`. To only rebuild reports after an input/label tweak (no scraping):
   `node scrape.js --report-only`.

That's it — steps 2–4 are the whole weekly job. The agent should NOT re-derive the analysis or
re-read every file; this runbook + the sections below are the source of truth.

## Run it (all commands)
```bash
node login.js                  # one-time: opens a browser; log in with mobile + OTP (auto-saves session)
node scrape.js                 # gentle, cache-backed run; only fetches uncached entities
node scrape.js --max-age 6     # WEEKLY: re-fetch cache older than 6 days (resumable); rebuilds reports
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
