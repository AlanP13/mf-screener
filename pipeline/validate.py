"""Sanity checks on a pipeline run; exits non-zero so a bad run is never published."""
import json, sys
from pathlib import Path
import pandas as pd

d = Path(sys.argv[1] if len(sys.argv) > 1 else "data")
m = json.loads((d / "manifest.json").read_text())
f = pd.read_csv(d / "funds.csv")
errors = []
if len(f) < 1200:
    errors.append(f"only {len(f)} funds (expected 1,500+)")
if m["history_failed"] > 0.05 * m["funds"]:
    errors.append(f"{m['history_failed']} NAV histories failed")
if f.aum_cr.isna().mean() > 0.10:
    errors.append("over 10% of funds missing AUM")
if f.ret_1y.notna().sum() < 1000:
    errors.append("too few 1Y returns")
if f.code.duplicated().any():
    errors.append("duplicate scheme codes")
age = (pd.Timestamp(m["asof_ist"]) - pd.Timestamp(m["latest_nav_date"])).days
if age > 5:
    errors.append(f"latest NAV is {age} days old")
for name in ("quarterly_nav.csv", "quarterly_aum.csv", "backfill.csv"):
    if not (d / name).exists():
        errors.append(f"missing {name}")
print(json.dumps({"funds": len(f), "nav_date": m["latest_nav_date"], "errors": errors}, indent=2))
sys.exit(1 if errors else 0)
