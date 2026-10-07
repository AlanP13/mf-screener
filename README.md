# Indian Mutual Fund Screener

Weekly screener for every open-ended Indian mutual fund (Regular plan, Growth option, ~1,700 schemes).
A Python pipeline publishes CSVs to `data/`; a Google Sheet pulls them in hourly through Apps Script
and applies the screen, persistence rules, history and dashboard.

```
Saturday ~8:47 AM ET          hourly (Apps Script trigger)
pipeline/mf_pipeline.py  ──►  data/*.csv  ──►  Google Sheet: All Funds → screen formulas
  AMFI NAVAll, mfapi.in,      (this repo)      → status engine → Changes / History / Archive
  AMFI AAUM + TER + scheme                     → Dashboard, Rankings, Watchlist, Portfolio
  master, Kuvera
```

## Why Saturday
AMFI publishes Friday NAVs Friday night IST. Saturday morning US Eastern is Saturday evening IST,
so the full week is in, and mfapi.in has caught up. Sunday adds nothing.
Quarterly average AUM arrives from AMFI in the first weeks after quarter end and is picked up
automatically by the next Saturday run.

## Files
| Path | What |
|---|---|
| `pipeline/mf_pipeline.py` | Universe, NAV history, returns, 52W range, quarterly average NAV/AUM, risk metrics, TER, backfill |
| `pipeline/validate.py` | Blocks publishing if the run looks wrong (fund count, failed histories, stale NAV, missing AUM) |
| `apps_script/Code.gs` | Paste into the Sheet (Extensions → Apps Script). Builds tabs, formulas, status engine |
| `data/funds.csv` | One row per fund |
| `data/quarterly_nav.csv`, `data/quarterly_aum.csv` | Calendar-quarter averages |
| `data/backfill.csv` | 24 month-end snapshots, used once to seed history so streaks are meaningful on day one |
| `run_weekly.sh` | Refresh + validate + commit + push |
| `.github/workflows/refresh.yml` | Manual fallback run on GitHub Actions |

## Method notes
* Returns: 3M, 6M, 1Y point-to-point; 2Y+ CAGR. NAV on or before each anniversary date.
* Since inception: AMFI history starts 03-Apr-2006. For older equity/hybrid funds SI assumes the
  Rs 10 launch NAV (matches AMC factsheets closely); otherwise SI is "since Apr-2006" and labelled.
* AUM: latest month fund-level AUM (Kuvera, via the Direct plan ISIN), falling back to AMFI quarterly
  average AUM summed across all plans. Quarterly average AUM is AMFI's, fund level.
* Face-value changes in NAV history (e.g. Rs 10 → Rs 1,000) are detected and rebased.
* Consistency: share of monthly rolling 1Y windows over 5 years at or above the category median.
* Sharpe/Sortino: 3Y daily, risk-free 6.5%.
