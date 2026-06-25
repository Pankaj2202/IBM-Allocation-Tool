"""
Old-logic enrichment, ported from the .xlsb (Pendency sheet formulas).
Adds: Invoice Receipt Date, Tenure (working-day aging), Aging Bucket, Priority (P-tier
+ number), Status (Parked-but-not-posted vs Fresh Allocation), VG->Team, Final_Tag
eligibility. Status drives the precise Parker/Poster split; VG->Team + scanner fallback
drive routing (incl. blank vendor groups).
"""
from __future__ import annotations
import datetime as dt
import numpy as np
import pandas as pd


def _s(v):
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()


def _to_date(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return pd.Timestamp(v).normalize()
    try:
        return pd.to_datetime(v).normalize()
    except Exception:
        return pd.NaT


def tenure_workdays(receipt, report_date, holidays):
    """Excel NETWORKDAYS.INTL(receipt, report, 1, holidays) — inclusive of both ends."""
    hol = np.array([np.datetime64(h) for h in holidays], dtype="datetime64[D]")
    out = []
    rep = np.datetime64(pd.Timestamp(report_date).date(), "D")
    for r in receipt:
        if pd.isna(r):
            out.append(np.nan); continue
        start = np.datetime64(pd.Timestamp(r).date(), "D")
        end = rep + np.timedelta64(1, "D")          # +1 day => inclusive of report date
        if start > rep:
            out.append(0); continue
        out.append(int(np.busday_count(start, end, holidays=hol)))
    return pd.Series(out)


def bucket_for(tenure, buckets):
    """Approximate-match (VLOOKUP ...,1): largest threshold <= tenure."""
    th = sorted(buckets)  # [[0,'0-2 Days'], ...]
    def f(t):
        if pd.isna(t): return ""
        label = th[0][1]
        for thr, lab in th:
            if t >= thr: label = lab
            else: break
        return label
    return tenure.map(f)


# canonical team names
def _canon_team(team_raw, division):
    t = (team_raw or "").strip().lower()
    div = (division or "").strip().lower()
    if t == "plant": return "Plant"
    if t == "freight": return "Freight"
    if t == "ho grey": return "HO-Grey"
    if t == "ho white": return "HO-White"
    if "grey & white" in t or "gery & white" in t or t == "all teams":
        return "HO-White" if div == "white" else "HO-Grey"
    return None  # unknown -> caller decides


def enrich(out: pd.DataFrame, df: pd.DataFrame, cfg: dict, report_date=None) -> pd.DataFrame:
    old = cfg["old_logic"]
    cols = cfg["dump"]["cols"]
    report_date = report_date or dt.date.today()

    def col(key):
        c = cols.get(key)
        return df[c] if c and c in df.columns else pd.Series([None] * len(df), index=df.index)

    receipt = col("receipt_date").map(_to_date)
    out["Invoice Receipt Date"] = receipt.dt.strftime("%d-%b-%Y").fillna("")
    out["Tenure"] = tenure_workdays(list(receipt), report_date, old["holidays"]).values
    out["Aging Bucket"] = bucket_for(out["Tenure"], old["tenure_buckets"])

    # Priority (P-tier by vendor group, then number)
    vg_pri = {k.strip().lower(): v for k, v in old["vg_priority"].items()}
    out["Priority"] = out["Vendor Group"].str.lower().map(vg_pri).fillna("Other")
    pn = old["priority_number"]
    out["Priority No"] = out["Priority"].map(lambda p: pn.get(p, 99))

    # Status: Parked But Not Posted vs Fresh Allocation
    park = col("parking_ref").map(_s)
    sap = col("sap_ref").map(_s)
    parked = (park != "") & (sap == "")
    out["Status"] = np.where(parked, "Parked But Not Posted", "Fresh Allocation")

    # Eligibility (Logic 1/3/4/5 -> Final_Tag)
    amt = out["Amount"]
    thr = old.get("amount_threshold", 100000)
    ten = out["Tenure"].fillna(0)
    pno = out["Priority No"]
    hold = col("hold_comments").map(_s)
    L1 = ((ten > 3) & (pno < 6) & (out["Team"] == "Plant")) | ((ten > 4) & (pno < 6))
    L4 = ((ten > 11) | (amt > thr)) & (pno == 6)
    L5 = (ten < 5) & (amt > thr)
    L3 = hold.str.len() > 5
    out["Final_Tag"] = np.where(L1 | L3 | L4 | L5, "Eligible", "Not Eligible")
    return out


def assign_supervisor(out: pd.DataFrame, cfg: dict) -> pd.Series:
    """Supervisor (team-leader) layer, reverse-engineered from the old workbook:
    Current Status == 'Hold'  -> hold supervisor (Anas, all teams)
    otherwise                 -> the team's supervisor.
    """
    old = cfg["old_logic"]
    hold_sup = old.get("hold_supervisor", "Anas")
    tsup = old.get("team_supervisor", {})
    on_hold = out["Current Status"].astype(str).str.strip().str.lower().eq("hold")
    team_sup = out["Team"].map(lambda t: tsup.get(t, hold_sup))
    return team_sup.where(~on_hold, hold_sup)


def team_via_old(out: pd.DataFrame, df: pd.DataFrame, cfg: dict) -> pd.Series:
    """Route using the maintained VG->Team table; blank VG -> scanner-location fallback."""
    old = cfg["old_logic"]
    cols = cfg["dump"]["cols"]
    vg_team = {k.strip().lower(): v for k, v in old["vg_team"].items()}
    scan = old["scanner_tag"]
    sender = (df[cols["email_sender"]] if cols.get("email_sender") in df.columns
              else pd.Series([None] * len(df), index=df.index)).map(_s).str.lower()

    teams = []
    for i in out.index:
        vg = out.at[i, "Vendor Group"]
        div = out.at[i, "Division"]
        if out.at[i, "Auto/Manual"] == "Auto":
            teams.append("Auto/API"); continue
        if vg == "":
            tag = scan.get(sender.at[i], "")
            teams.append("Plant" if tag.lower() == "plant"
                         else ("HO-Grey" if tag.lower() == "ho grey" else "Plant"))
            continue
        raw = vg_team.get(vg.lower())
        canon = _canon_team(raw, div) if raw else None
        if canon is None:
            # fall back to division / default Plant
            canon = "HO-White" if div.lower() == "white" else "Plant"
        teams.append(canon)
    return pd.Series(teams, index=out.index)
