#!/usr/bin/env python3
"""Indian Mutual Fund Screener - weekly data pipeline.

Builds the Regular-plan / Growth-option universe of open-ended Indian mutual
funds and writes the CSVs that the Google Sheet imports:

  data/funds.csv          one row per fund: NAV, 52W range, returns, risk, AUM
  data/quarterly_nav.csv  average NAV per calendar quarter
  data/quarterly_aum.csv  fund-level average AUM per calendar quarter (AMFI)
  data/manifest.json      run metadata, source status and validation counts

Sources
  AMFI NAVAll.txt        universe, category, latest NAV
  mfapi.in               full NAV history per scheme code (AMFI history from Apr-2006)
  AMFI average-AUM API   scheme-wise quarterly average AUM (Rs lakh)
  Kuvera (mf.captnemo.in) latest month AUM, expense ratio, fund manager, true inception date

Usage: python mf_pipeline.py --out ../data [--workers 16] [--limit N]
"""
from __future__ import annotations

import argparse
import concurrent.futures as cf
import csv
import io
import datetime as dt
import json
import math
import re
import sys
import time
from collections import defaultdict
from pathlib import Path

import numpy as np
import pandas as pd
import requests

NAVALL_URL = "https://portal.amfiindia.com/spages/NAVAll.txt"
MFAPI_URL = "https://api.mfapi.in/mf/{code}"
AAUM_URL = "https://www.amfiindia.com/api/average-aum-schemewise"
KUVERA_URL = "https://mf.captnemo.in/kuvera/{isin}"

RISK_FREE = 0.065          # annual, used for Sharpe / Sortino
TRADING_DAYS = 250
QUARTERS_BACK = 8          # completed calendar quarters to report
BACKFILL_MONTHS = 24       # month-end snapshots used to seed screening history
STALE_NAV_DAYS = 30        # schemes whose last NAV is older than this are dropped
AMFI_HISTORY_START = dt.date(2006, 4, 3)

SESSION = requests.Session()
SESSION.headers["User-Agent"] = "mf-screener/1.0 (personal research)"


def get(url, params=None, tries=4, timeout=60):
    last = None
    for i in range(tries):
        try:
            r = SESSION.get(url, params=params, timeout=timeout)
            if r.status_code == 200:
                return r
            last = f"HTTP {r.status_code}"
            if r.status_code == 404:
                break
        except requests.RequestException as e:  # network hiccup
            last = str(e)
        time.sleep(1.5 * (i + 1))
    raise RuntimeError(f"{url}: {last}")


# --------------------------------------------------------------------------- categories
SUB_MAP = [
    (r"^large cap", "Large Cap"),
    (r"^large & mid cap", "Large & Mid Cap"),
    (r"^mid cap", "Mid Cap"),
    (r"^small cap", "Small Cap"),
    (r"^multi cap", "Multi Cap"),
    (r"^flexi cap", "Flexi Cap"),
    (r"^focused", "Focused"),
    (r"^value", "Value"),
    (r"^contra", "Contra"),
    (r"^dividend yield", "Dividend Yield"),
    (r"^elss", "ELSS"),
    (r"sectoral|thematic", "Sectoral/Thematic"),
    (r"aggressive hybrid", "Aggressive Hybrid"),
    (r"conservative hybrid", "Conservative Hybrid"),
    (r"balanced hybrid", "Balanced Hybrid"),
    (r"balanced advantage|dynamic asset allocation", "Balanced Advantage"),
    (r"multi asset", "Multi Asset Allocation"),
    (r"equity savings", "Equity Savings"),
    (r"arbitrage", "Arbitrage"),
    (r"overnight", "Overnight"),
    (r"^liquid", "Liquid"),
    (r"money market", "Money Market"),
    (r"ultra short to short", "Low Duration"),
    (r"low duration", "Low Duration"),
    (r"ultra short", "Ultra Short Duration"),
    (r"short duration|short term", "Short Duration"),
    (r"medium to long", "Medium to Long Duration"),
    (r"medium duration|medium term", "Medium Duration"),
    (r"long duration|long term", "Long Duration"),
    (r"dynamic bond|dynamic term", "Dynamic Bond"),
    (r"corporate bond", "Corporate Bond"),
    (r"credit risk", "Credit Risk"),
    (r"banking and psu", "Banking & PSU"),
    (r"10.year constant|constant duration", "Gilt 10Y Constant Duration"),
    (r"gilt", "Gilt"),
    (r"floater|floating", "Floater"),
    (r"retirement", "Retirement"),
    (r"child", "Children's"),
]


def classify(header: str) -> tuple[str, str, str]:
    """Return (broad, category, raw) from an AMFI NAVAll section header."""
    raw = header[header.find("(") + 1: header.rfind(")")] if "(" in header else header
    grp, _, sub = raw.partition(" - ")
    g, s = grp.lower(), sub.lower()
    if "index" in g or "index fund" in s:
        broad = "Index Funds"
        cat = "Index - Debt" if "debt" in s else "Index - Equity"
        return broad, cat, raw
    if "etf" in g or "etf" in s:
        return "ETF", "ETF", raw
    if "fof" in s or "fund of funds" in g or "fund of funds" in s:
        over = "overseas" in g or "overseas" in s
        return "FoF", "FoF Overseas" if over else "FoF Domestic", raw
    if "equity" in g:
        broad = "Equity"
    elif "hybrid" in g:
        broad = "Hybrid"
    elif "debt" in g or "income" in g:
        broad = "Debt"
    elif "solution" in g or "child" in g:
        broad = "Solution Oriented"
    else:
        broad = "Other"
    for pat, name in SUB_MAP:
        if re.search(pat, s or g):
            return broad, name, raw
    return broad, (sub or grp).strip() or "Other", raw


# --------------------------------------------------------------------------- universe
def load_universe(out_raw: Path) -> pd.DataFrame:
    txt = get(NAVALL_URL).text
    (out_raw / "NAVAll.txt").write_text(txt, encoding="utf-8")
    header, amc, rows = None, None, []
    for line in txt.splitlines():
        line = line.strip()
        if not line:
            continue
        p = line.split(";")
        if len(p) >= 8 and p[0].isdigit():
            rows.append(dict(header=header, amc=amc, code=int(p[0]), isin=p[1].strip(),
                             isin_reinv=p[2].strip(), base_name=p[3].strip(), plan=p[4].strip(),
                             option=p[5].strip(), nav=p[6].strip(), nav_date=p[7].strip()))
        elif len(p) == 1:
            if re.match(r"^(Open|Close|Interval)", line):
                header = line
            else:
                amc = line
    df = pd.DataFrame(rows)
    df["all_rows"] = len(df)
    return df


def select_regular_growth(df: pd.DataFrame, asof: dt.date) -> pd.DataFrame:
    d = df[df.header.str.startswith("Open Ended")].copy()
    d = d[d.plan.eq("Regular Plan")]
    opt = d.option.str.lower()
    d = d[opt.str.contains("growth") & ~opt.str.contains(r"idcw|dividend|bonus")]
    d["nav_date"] = pd.to_datetime(d.nav_date, format="%d-%b-%Y").dt.date
    d["nav"] = pd.to_numeric(d.nav, errors="coerce")
    d = d[(d.nav > 0) & d.nav_date.map(lambda x: (asof - x).days <= STALE_NAV_DAYS)]
    cls = d.header.map(classify)
    d["broad"] = cls.map(lambda t: t[0])
    d["category"] = cls.map(lambda t: t[1])
    d["amfi_category"] = cls.map(lambda t: t[2])
    d = d[d.broad != "ETF"]
    return d.reset_index(drop=True)


# --------------------------------------------------------------------------- history
def fetch_history(code: int):
    r = get(MFAPI_URL.format(code=code))
    js = r.json()
    data = js.get("data") or []
    if not data:
        return code, None
    s = pd.Series({pd.Timestamp(dt.datetime.strptime(x["date"], "%d-%m-%Y")): float(x["nav"])
                   for x in data if x.get("nav") not in (None, "", "0", "0.0")})
    s = s[s > 0].sort_index()
    s = s[~s.index.duplicated(keep="last")]
    return code, clean_nav(s)


def clean_nav(s: pd.Series) -> pd.Series:
    """Remove one-day spikes and rebase face-value changes (e.g. Rs 10 -> Rs 1000)."""
    s = s.copy()
    s.attrs["rebased"] = 0
    for _ in range(20):
        r = (s / s.shift(1)).iloc[1:]
        bad = r[(r > 1.5) | (r < 0.67)]
        if bad.empty:
            break
        t = bad.index[0]
        i = s.index.get_loc(t)
        nxt = s.iloc[i + 1] / s.iloc[i] if i + 1 < len(s) else None
        if nxt is not None and 0.8 < r.loc[t] * nxt < 1.25:
            s = s.drop(t)                               # isolated bad print
        else:
            s.iloc[:i] = s.iloc[:i] * r.loc[t]          # rebase history before the jump
            s.attrs["rebased"] += 1
    return s


def nav_on_or_before(s: pd.Series, when: pd.Timestamp):
    idx = s.index.searchsorted(when, side="right") - 1
    if idx < 0:
        return None, None
    return s.index[idx], s.iloc[idx]


def period_return(s: pd.Series, end: pd.Timestamp, years: float | None = None, months: int | None = None):
    start_target = end - pd.DateOffset(years=int(years)) if years else end - pd.DateOffset(months=months)
    if s.index[0] > start_target + pd.Timedelta(days=7):   # not enough history
        return None
    d0, v0 = nav_on_or_before(s, start_target)
    if d0 is None or (start_target - d0).days > 10:
        return None
    v1 = s.iloc[-1]
    if years and years >= 1:
        return (v1 / v0) ** (1 / years) - 1
    return v1 / v0 - 1


def si_return(s: pd.Series, inception, first, truncated: bool, broad: str):
    """Since-inception CAGR (absolute if under a year) up to the last point of s."""
    end = s.index[-1]
    yrs_hist = (end - s.index[0]).days / 365.25
    if yrs_hist <= 0:
        return None, None
    if not truncated:
        if yrs_hist >= 1:
            return (s.iloc[-1] / s.iloc[0]) ** (1 / yrs_hist) - 1, "NAV history"
        return s.iloc[-1] / s.iloc[0] - 1, "NAV history (absolute, <1Y)"
    # NAV before Apr-2006 is not in AMFI history. Equity-oriented funds launched at Rs 10,
    # so SI CAGR = (NAV/10)^(1/age); otherwise report CAGR since Apr-2006.
    age = (end.date() - inception).days / 365.25 if inception else 0
    pre = None
    if inception and inception < first:
        pre = (s.iloc[0] / 10.0) ** (1 / max((first - inception).days / 365.25, 0.5)) - 1
    if broad in ("Equity", "Hybrid", "Solution Oriented") and age > yrs_hist and pre is not None \
            and -0.15 < pre < 0.80:
        return (s.iloc[-1] / 10.0) ** (1 / age) - 1, "Launch NAV Rs 10 assumed (pre-2006 NAV not in AMFI history)"
    return (s.iloc[-1] / s.iloc[0]) ** (1 / yrs_hist) - 1, "Since Apr-2006 (earlier NAV not in AMFI history)"


def backfill_rows(code, s, inception, first, truncated, broad, aum_by_q, months):
    """Month-end return snapshots used to seed the sheet's screening history."""
    rows = []
    for me in months:
        sub = s[s.index <= me]
        if len(sub) < 2 or (me - sub.index[-1]).days > 7:
            continue
        end = sub.index[-1]
        row = dict(month=me.strftime("%Y-%m"), date=end.date().isoformat(), code=code)
        for yrs in (1, 2, 3, 5, 10):
            v = period_return(sub, end, years=yrs)
            row[f"ret_{yrs}y"] = round(v, 5) if v is not None else None
        si, _ = si_return(sub, inception, first, truncated, broad)
        row["ret_si"] = round(si, 5) if si is not None else None
        q = pd.Period(me, freq="Q")
        aum = None
        for k in range(0, 3):
            aum = aum_by_q.get(q - k)
            if aum is not None:
                break
        row["aum_cr"] = aum
        row["age_years"] = round((end.date() - inception).days / 365.25, 2) if inception else None
        rows.append(row)
    return rows


def risk_metrics(s: pd.Series, end: pd.Timestamp, years: int):
    start = end - pd.DateOffset(years=years)
    if s.index[0] > start + pd.Timedelta(days=7):
        return dict(vol=None, mdd=None, sharpe=None, sortino=None)
    w = s[s.index >= start]
    r = np.log(w).diff().dropna()
    if len(r) < 100:
        return dict(vol=None, mdd=None, sharpe=None, sortino=None)
    vol = r.std(ddof=1) * math.sqrt(TRADING_DAYS)
    cagr = (w.iloc[-1] / w.iloc[0]) ** (1 / years) - 1
    rf_d = math.log(1 + RISK_FREE) / TRADING_DAYS
    downside = np.minimum(r - rf_d, 0)
    dd = math.sqrt((downside ** 2).mean()) * math.sqrt(TRADING_DAYS)
    mdd = (w / w.cummax() - 1).min()
    return dict(vol=vol, mdd=mdd,
                sharpe=(cagr - RISK_FREE) / vol if vol > 0 else None,
                sortino=(cagr - RISK_FREE) / dd if dd > 0 else None)


def quarter_label(p: pd.Period) -> str:
    names = {1: "Jan-Mar", 2: "Apr-Jun", 3: "Jul-Sep", 4: "Oct-Dec"}
    return f"{names[p.quarter]} {p.year}"


# --------------------------------------------------------------------------- AUM
PLAN_TOKENS = re.compile(
    r"\b(direct|regular|retail|institutional|plan|option|growth|idcw|dividend|payout|reinvestment|"
    r"re-investment|daily|weekly|fortnightly|monthly|quarterly|half yearly|half-yearly|annual|yearly|"
    r"bonus|income distribution cum capital withdrawal|income distribution|capital withdrawal|cum|"
    r"unclaimed|segregated|portfolio|existing|number|of|no\.|formerly known as.*)\b", re.I)


def fund_key(name: str) -> str:
    n = name.lower().replace("&", " and ")
    n = re.sub(r"\(.*?\)", " ", n)
    n = PLAN_TOKENS.sub(" ", n)
    n = re.sub(r"[^a-z0-9]+", " ", n)
    n = re.sub(r"\b(fund|scheme|the|mutual)\b", " ", n)
    return re.sub(r"\s+", " ", n).strip()


def fetch_aaum(quarters: list[pd.Period]) -> tuple[dict, dict, list]:
    """Return ({period: {amfi_code: lakh}}, {amfi_code: name}, notes)."""
    years = get(AAUM_URL, params=dict(strType="Categorywise", MF_ID=0)).json()["data"]
    fy_ids = {y["financial_year"]: y["id"] for y in years}
    out, names, notes = {}, {}, []
    month = {"January": 1, "April": 4, "July": 7, "October": 10}
    for fy_label, fy_id in fy_ids.items():
        fy_start = int(fy_label.split()[1])
        if fy_start < quarters[0].year - 1:
            continue
        per = get(AAUM_URL, params=dict(fyId=fy_id, strType="Categorywise", MF_ID=0)).json()
        for p in per["data"]["periods"]:
            first = p["period"].split()[0]
            yr = int(p["period"].split()[-1])
            if first not in month:
                continue
            q = pd.Period(year=yr, quarter=(month[first] - 1) // 3 + 1, freq="Q")
            if q not in quarters:
                continue
            js = get(AAUM_URL, params=dict(strType="Categorywise", fyId=fy_id, periodId=p["id"], MF_ID=0),
                     timeout=120).json()
            vals = {}
            for grp in js.get("data", []):
                for sc in grp.get("schemes", []):
                    a = sc.get("AverageAumForTheMonth") or {}
                    v = (a.get("ExcludingFundOfFundsDomesticButIncludingFundOfFundsOverseas") or 0) + \
                        (a.get("FundOfFundsDomestic") or 0)
                    vals[int(sc["AMFI_Code"])] = v
                    names[int(sc["AMFI_Code"])] = sc.get("SchemeNAVName", "")
            out[q] = vals
            notes.append(f"{quarter_label(q)}: {len(vals)} scheme rows")
    return out, names, notes


def fetch_kuvera(isin: str):
    try:
        js = get(KUVERA_URL.format(isin=isin), tries=2, timeout=30).json()
    except Exception:
        return isin, None
    if not isinstance(js, list) or not js:
        return isin, None
    d = js[0]
    ret = d.get("returns") or {}
    return isin, dict(
        k_aum_cr=(d.get("aum") / 10.0) if isinstance(d.get("aum"), (int, float)) else None,
        k_ter=d.get("expense_ratio"), k_ter_date=d.get("expense_ratio_date"),
        k_manager=d.get("fund_manager"), k_start=d.get("start_date"),
        k_si=(ret.get("inception") / 100.0) if isinstance(ret.get("inception"), (int, float)) else None,
        k_rating=d.get("fund_rating"), k_risk=d.get("crisil_rating"))



SCHEME_MASTER_URL = "https://portal.amfiindia.com/DownloadSchemeData_Po.aspx?mf=0"
TER_URL = "https://www.amfiindia.com/api/populate-te-rdata-revised"


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def load_scheme_master(raw: Path) -> pd.DataFrame:
    txt = get(SCHEME_MASTER_URL, timeout=120).text
    (raw / "scheme_master.csv").write_text(txt, encoding="utf-8")
    m = pd.read_csv(io.StringIO(txt), dtype=str)
    m.columns = [c.strip() for c in m.columns]
    m["Code"] = pd.to_numeric(m["Code"], errors="coerce")
    m = m.dropna(subset=["Code"])
    m["Code"] = m["Code"].astype(int)
    m["launch"] = pd.to_datetime(m["Launch Date"], format="%d-%b-%Y", errors="coerce").dt.date
    m["fund_norm"] = m["Scheme Name"].map(norm)
    return m


def load_regular_ter(raw: Path, asof: dt.date) -> tuple[dict, str]:
    """Latest Regular-plan total TER per fund name from AMFI's TER disclosure.

    The current month is usually partial (AMCs file as they go), so the previous
    two months are loaded too and the most recent dated figure wins."""
    frames, months = [], []
    for back in (2, 1, 0):
        y, mth = asof.year, asof.month - back
        while mth <= 0:
            mth += 12
            y -= 1
        month = f"{mth:02d}-{y}"
        try:
            r = get(TER_URL, params=dict(MF_ID="All", Month=month, strCat="-1", strType="-1", excel="true"),
                    timeout=300)
            df = pd.read_excel(io.BytesIO(r.content))
        except Exception:
            continue
        if not df.empty:
            frames.append(df)
            months.append(month)
    if not frames:
        return {}, ""
    df = pd.concat(frames)
    col = [c for c in df.columns if c.startswith("Regular Plan") and "Total TER" in c][0]
    df = df.dropna(subset=["Scheme Name", col]).sort_values("TER Date")
    last = df.groupby("Scheme Name").tail(1)
    return {norm(k): v for k, v in zip(last["Scheme Name"], last[col])}, ", ".join(months)


# --------------------------------------------------------------------------- main
def pct_rank_within(df: pd.DataFrame, col: str, group: str, higher_better=True, min_n=5):
    def f(x):
        v = x[col]
        n = v.notna().sum()
        if n < min_n:
            return pd.Series([np.nan] * len(x), index=x.index)
        r = v.rank(pct=True, ascending=higher_better)
        return (r * 100).round(1)
    return df.groupby(group, group_keys=False)[[col]].apply(f)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="data")
    ap.add_argument("--workers", type=int, default=16)
    ap.add_argument("--limit", type=int, default=0)
    a = ap.parse_args()
    out = Path(a.out)
    raw = out / "_raw"
    out.mkdir(parents=True, exist_ok=True)
    raw.mkdir(exist_ok=True)
    started = dt.datetime.now(dt.timezone.utc)
    asof = (started + dt.timedelta(hours=5, minutes=30)).date()   # IST date

    allrows = load_universe(raw)
    uni = select_regular_growth(allrows, asof)
    uni = uni.sort_values(["code"]).drop_duplicates("code")
    if a.limit:
        uni = uni.head(a.limit)
    print(f"universe: {len(uni)} regular-growth schemes", flush=True)

    # ---- NAV history
    hist, failed = {}, []
    with cf.ThreadPoolExecutor(a.workers) as ex:
        for i, fut in enumerate(cf.as_completed([ex.submit(fetch_history, c) for c in uni.code])):
            try:
                code, s = fut.result()
                if s is not None and len(s) > 1:
                    hist[code] = s
            except Exception as e:
                failed.append(str(e)[:120])
            if i % 200 == 0:
                print(f"  history {i}/{len(uni)}", flush=True)
    print(f"history ok {len(hist)}, failed {len(failed)}", flush=True)

    # ---- scheme master (fund grouping, launch date) and regular TER
    master = load_scheme_master(raw)
    fund_of_code = dict(zip(master.Code, master.fund_norm))
    launch_of_code = dict(zip(master.Code, master.launch))
    ter_map, ter_month = load_regular_ter(raw, asof)
    print(f"scheme master {len(master)}, TER funds {len(ter_map)} ({ter_month})", flush=True)

    # ---- Kuvera (direct-plan ISIN of the same fund: fund-level AUM, manager)
    oe = allrows[allrows.header.str.startswith("Open Ended") & allrows.plan.eq("Direct Plan")]
    oe = oe[oe.option.str.lower().str.contains("growth") & ~oe.option.str.lower().str.contains("idcw|dividend|bonus")]
    direct_isin = {(a_, b_): i_ for a_, b_, i_ in zip(oe.amc, oe.base_name, oe["isin"]) if str(i_).startswith("INF")}
    uni["direct_isin"] = [direct_isin.get((a_, b_)) for a_, b_ in zip(uni.amc, uni.base_name)]
    kmeta = {}
    with cf.ThreadPoolExecutor(8) as ex:
        for isin, m in ex.map(fetch_kuvera, [i for i in uni["direct_isin"].dropna().unique()]):
            if m:
                kmeta[isin] = m
    print(f"kuvera ok {len(kmeta)}", flush=True)

    # ---- quarters
    cur_q = pd.Period(pd.Timestamp(asof), freq="Q")
    quarters = [cur_q - i for i in range(QUARTERS_BACK, 0, -1)]
    aaum, aaum_names, aaum_notes = {}, {}, []
    try:
        aaum, aaum_names, aaum_notes = fetch_aaum(quarters)
    except Exception as e:
        aaum_notes.append(f"AAUM fetch failed: {e}")

    # fund-level AAUM: sum every plan/option of the same fund within the AMC
    key_of_code = {c: fund_of_code.get(c) or ("~" + fund_key(n)) for c, n in aaum_names.items()}
    fund_aaum = {q: defaultdict(float) for q in aaum}
    for q, vals in aaum.items():
        for c, v in vals.items():
            fund_aaum[q][key_of_code[c]] += v

    nav_name = dict(zip(master.Code, master["Scheme NAV Name"]))
    dup = uni.duplicated(["amc", "base_name"], keep=False)
    uni["display"] = [nav_name.get(c, b) if d else b for c, b, d in zip(uni.code, uni.base_name, dup)]

    # ---- per fund metrics
    end_ts = {}
    recs, qnav_rows, qaum_rows, bf_rows = [], [], [], []
    bf_months = [(pd.Timestamp(asof).to_period("M") - k).to_timestamp(how="end").normalize()
                 for k in range(BACKFILL_MONTHS, 0, -1)]
    monthly = {}
    for r in uni.itertuples():
        s = hist.get(r.code)
        rec = dict(code=r.code, isin=r.isin, name=r.display, amc=r.amc, broad=r.broad,
                   category=r.category, amfi_category=r.amfi_category, plan="Regular", option="Growth",
                   nav=r.nav, nav_date=r.nav_date.isoformat())
        km = kmeta.get(r.direct_isin, {}) if r.direct_isin else {}
        first = s.index[0].date() if s is not None else None
        launch = launch_of_code.get(r.code)
        launch = launch if isinstance(launch, dt.date) else None
        truncated = first is not None and first <= AMFI_HISTORY_START + dt.timedelta(days=7)
        if launch and (first is None or launch <= first):
            inception, incep_src = launch, "AMFI scheme master"
        else:
            inception = first
            incep_src = ("Before Apr-2006 (AMFI history starts 03-Apr-2006)" if truncated
                         else "First NAV in AMFI history")
        rec.update(inception=inception.isoformat() if inception else None, inception_source=incep_src,
                   age_years=round((asof - inception).days / 365.25, 2) if inception else None)
        if s is not None:
            end = s.index[-1]
            end_ts[r.code] = end
            rec["hist_nav_date"] = end.date().isoformat()
            rec["rebased"] = s.attrs.get("rebased", 0)
            y1 = s[s.index > end - pd.DateOffset(years=1)]
            rec["high_52w"] = y1.max()
            rec["low_52w"] = y1.min()
            rec["high_52w_date"] = y1.idxmax().date().isoformat()
            rec["low_52w_date"] = y1.idxmin().date().isoformat()
            rec["pct_from_high"] = s.iloc[-1] / y1.max() - 1
            rec["ret_3m"] = period_return(s, end, months=3)
            rec["ret_6m"] = period_return(s, end, months=6)
            for yrs in (1, 2, 3, 5, 10):
                rec[f"ret_{yrs}y"] = period_return(s, end, years=yrs)
            rec["ret_si"], rec["si_source"] = si_return(s, inception, first, truncated, r.broad)
            m3, m5 = risk_metrics(s, end, 3), risk_metrics(s, end, 5)
            rec.update(vol_3y=m3["vol"], mdd_3y=m3["mdd"], sharpe_3y=m3["sharpe"], sortino_3y=m3["sortino"],
                       mdd_5y=m5["mdd"])
            monthly[r.code] = s.resample("ME").last().dropna()
            # quarterly average NAV
            qrow = dict(code=r.code, name=r.base_name, category=r.category)
            sp = s.groupby(s.index.to_period("Q")).mean()
            for q in quarters + [cur_q]:
                lab = quarter_label(q) + (" (QTD)" if q == cur_q else "")
                qrow[lab] = round(sp.get(q), 4) if q in sp.index else None
            qnav_rows.append(qrow)
        # AUM
        rec["aum_cr"] = km.get("k_aum_cr")
        rec["aum_source"] = "Kuvera (latest month)" if rec["aum_cr"] is not None else None
        key = fund_of_code.get(r.code) or ("~" + fund_key(r.base_name))
        arow = dict(code=r.code, name=r.base_name, category=r.category)
        latest_aaum = None
        for q in quarters:
            v = fund_aaum.get(q, {}).get(key)
            v = round(v / 100.0, 2) if v else None        # lakh -> crore
            arow[quarter_label(q)] = v
            if v is not None:
                latest_aaum = (q, v)
        qaum_rows.append(arow)
        if s is not None:
            aum_by_q = {q: arow[quarter_label(q)] for q in quarters if arow.get(quarter_label(q))}
            bf_rows += backfill_rows(r.code, s, inception, first, truncated, r.broad, aum_by_q, bf_months)
        rec["aaum_latest_q_cr"] = latest_aaum[1] if latest_aaum else None
        rec["aaum_latest_q"] = quarter_label(latest_aaum[0]) if latest_aaum else None
        plan_v = aaum.get(quarters[-1], {}).get(r.code)
        rec["aaum_regular_growth_cr"] = round(plan_v / 100.0, 2) if plan_v else None
        if rec["aum_cr"] is None and latest_aaum:
            rec["aum_cr"], rec["aum_source"] = latest_aaum[1], f"AMFI AAUM {quarter_label(latest_aaum[0])}"
        ter = ter_map.get(fund_of_code.get(r.code, ""))
        rec["expense_ratio"] = ter / 100.0 if ter is not None else None
        rec["fund_manager"] = km.get("k_manager")
        rec["risk_level"] = km.get("k_risk")
        recs.append(rec)

    f = pd.DataFrame(recs)

    # ---- consistency: share of monthly rolling-1Y windows (last 5Y) in top half of category
    mdf = pd.DataFrame(monthly)
    roll = mdf / mdf.shift(12) - 1
    roll = roll[roll.index >= roll.index.max() - pd.DateOffset(years=5)]
    cat_of = dict(zip(f.code, f.category))
    cons = {}
    for cat, codes in f.groupby("category").code:
        cols = [c for c in codes if c in roll.columns]
        if len(cols) < 5:
            continue
        sub = roll[cols]
        med = sub.median(axis=1)
        valid = sub.notna().sum(axis=1) >= 5
        for c in cols:
            x = sub[c][valid]
            m = med[valid]
            ok = x.notna()
            if ok.sum() >= 12:
                cons[c] = float((x[ok] >= m[ok]).mean())
    f["consistency_5y"] = f.code.map(cons)
    f["consistency_windows"] = f.code.map(lambda c: int(roll[c].notna().sum()) if c in roll.columns else None)

    # ---- category-relative figures
    for col in ("ret_1y", "ret_3y", "ret_5y"):
        med = f.groupby("category")[col].transform("median")
        f[f"{col}_vs_cat"] = f[col] - med
    f["pct_sharpe"] = pct_rank_within(f, "sharpe_3y", "category")
    f["pct_mdd"] = pct_rank_within(f, "mdd_3y", "category")           # less negative ranks higher
    f["pct_consistency"] = pct_rank_within(f, "consistency_5y", "category")
    f["peer_count"] = f.groupby("category").code.transform("count")

    # ---- validation flags
    def flag(r):
        out = []
        if pd.notna(r.get("hist_nav_date")) and r["hist_nav_date"] != r["nav_date"]:
            out.append("history lags NAVAll")
        if r.get("ret_1y") is not None and pd.notna(r.get("ret_1y")) and abs(r["ret_1y"]) > 1.5:
            out.append("1Y return > 150%")
        if pd.isna(r.get("aum_cr")):
            out.append("no AUM")
        if r.get("rebased"):
            out.append("NAV history rebased (face value change)")
        if pd.isna(r.get("hist_nav_date")):
            out.append("no NAV history")
        return "; ".join(out)
    f["data_flags"] = f.apply(flag, axis=1)
    f["data_date"] = asof.isoformat()

    cols = ["code", "isin", "name", "amc", "broad", "category", "amfi_category", "plan", "option",
            "inception", "inception_source", "age_years", "nav", "nav_date", "high_52w", "high_52w_date",
            "low_52w", "low_52w_date", "pct_from_high", "aum_cr", "aum_source", "aaum_latest_q_cr",
            "aaum_latest_q", "aaum_regular_growth_cr", "expense_ratio", "fund_manager", "risk_level",
            "ret_3m", "ret_6m", "ret_1y", "ret_2y", "ret_3y", "ret_5y", "ret_10y", "ret_si", "si_source",
            "ret_1y_vs_cat", "ret_3y_vs_cat", "ret_5y_vs_cat", "vol_3y", "mdd_3y", "mdd_5y", "sharpe_3y",
            "sortino_3y", "consistency_5y", "pct_sharpe", "pct_mdd", "pct_consistency", "peer_count",
            "data_flags", "data_date"]
    f = f[cols].sort_values(["broad", "category", "name"])
    f.to_csv(out / "funds.csv", index=False, float_format="%.6g", quoting=csv.QUOTE_MINIMAL)
    pd.DataFrame(qnav_rows).to_csv(out / "quarterly_nav.csv", index=False)
    pd.DataFrame(qaum_rows).to_csv(out / "quarterly_aum.csv", index=False)
    pd.DataFrame(bf_rows).to_csv(out / "backfill.csv", index=False)

    manifest = dict(
        run_id=started.strftime("%Y%m%dT%H%M%SZ"), run_utc=started.isoformat(timespec="seconds"),
        asof_ist=asof.isoformat(),
        latest_nav_date=max(f.nav_date), universe_rows=int(allrows.shape[0]), funds=int(len(f)),
        history_ok=len(hist), history_failed=len(failed), kuvera_ok=len(kmeta),
        aaum_quarters=aaum_notes, risk_free=RISK_FREE,
        sources=[
            dict(name="AMFI NAVAll", url=NAVALL_URL, used_for="Universe, category, latest NAV"),
            dict(name="mfapi.in", url="https://api.mfapi.in/mf/<scheme code>",
                 used_for="NAV history: returns, 52W range, quarterly average NAV, risk metrics"),
            dict(name="AMFI average AUM (scheme-wise)", url=AAUM_URL,
                 used_for="Quarterly average AUM, summed across all plans of a fund"),
            dict(name="AMFI scheme master", url=SCHEME_MASTER_URL,
                 used_for="Launch (inception) date; groups plans/options into one fund"),
            dict(name="AMFI TER disclosure", url=TER_URL, used_for="Regular-plan total expense ratio"),
            dict(name="Kuvera via mf.captnemo.in", url="https://mf.captnemo.in/kuvera/<direct ISIN>",
                 used_for="Latest month fund-level AUM and fund manager (looked up via the Direct plan ISIN)"),
        ],
        ter_month=ter_month,
        flags={k: int(v) for k, v in f.data_flags.str.split("; ").explode().value_counts().items() if k},
    )
    (out / "manifest.json").write_text(json.dumps(manifest, indent=2))
    print(json.dumps(manifest, indent=2))


if __name__ == "__main__":
    sys.exit(main())
