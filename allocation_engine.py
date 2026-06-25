"""
JKCL AP Allocation Engine  (new logic, user-independent)
--------------------------------------------------------
Pure logic layer: ingest Exela DCN-level detail -> classify -> route -> assign.
All mappings live in allocation_config.json so the team can edit rules without code.
"""
from __future__ import annotations
import json, os, itertools
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))


def load_config(path: str | None = None) -> dict:
    return json.load(open(path or os.path.join(HERE, "allocation_config.json")))


# ----------------------------- ingestion -----------------------------
def read_dump(path: str, cfg: dict) -> pd.DataFrame:
    """Read the Exela detailed report. Header is on a non-first row; find it."""
    hdr = cfg["dump"]["header_row"]
    # Try the configured header row first; fall back to auto-detect on 'DCN'.
    raw = pd.read_excel(path, header=None, dtype=object)
    target = cfg["dump"]["cols"]["dcn"]
    header_idx = hdr - 1
    if target not in list(raw.iloc[header_idx].astype(str)):
        for i in range(min(15, len(raw))):
            if target in list(raw.iloc[i].astype(str)):
                header_idx = i
                break
    df = pd.read_excel(path, header=header_idx, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def _g(df, cfg, key):
    """Resolve a logical field name to its actual column, return the Series or blanks."""
    col = cfg["dump"]["cols"][key]
    return df[col] if col in df.columns else pd.Series([None] * len(df))


def _s(v):
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()


# ----------------------------- classification -----------------------------
def classify(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    vg_util = {k.strip().lower(): v for k, v in cfg["vendor_group_utility"].items()}
    grey_cat = {k.strip().lower(): v for k, v in cfg["hogrey_vendorgroup_to_category"].items()}
    api_move = {x.strip().lower() for x in cfg["api_move_auto_categories"]}
    fr_prefix = tuple(p.lower() for p in cfg["freight_vendor_groups_prefix"])
    fr_doctypes = {x.lower() for x in cfg["freight_doc_types"]}

    out = pd.DataFrame(index=df.index)
    out["DCN"] = _g(df, cfg, "dcn").map(_s)
    out["Vendor Name"] = _g(df, cfg, "vendor_name").map(_s)
    out["Vendor Group"] = _g(df, cfg, "vendor_group").map(_s)
    out["Division"] = _g(df, cfg, "division").map(_s)
    out["Doc Type"] = _g(df, cfg, "doc_type").map(_s)
    out["Invoice No"] = _g(df, cfg, "invoice_no").map(_s)
    out["Invoice Date"] = _g(df, cfg, "invoice_date").map(_s)
    out["Amount"] = pd.to_numeric(_g(df, cfg, "amount"), errors="coerce").fillna(0.0)
    out["Current Queue"] = _g(df, cfg, "current_queue").map(_s)
    out["Current Status"] = _g(df, cfg, "current_status").map(_s)
    out["Processing Type"] = _g(df, cfg, "processing_type").map(_s)
    out["Claimed By"] = _g(df, cfg, "claimed_by").map(_s)
    msme_no = _g(df, cfg, "msme_number").map(_s)
    msme_cat = _g(df, cfg, "msme_category").map(_s)

    vg_l = out["Vendor Group"].str.lower()
    dt_l = out["Doc Type"].str.lower()

    # Utility / Non-Utility
    out["Utility"] = vg_l.map(vg_util).fillna("Unmapped")
    # MSME
    out["MSME Flag"] = ((msme_no != "") | (msme_cat != "")).map({True: "MSME", False: "—"})
    # Auto processing (already system-tagged)
    pt = out["Processing Type"].str.lower()
    cq = out["Current Queue"].str.lower()
    out["Auto/Manual"] = ((pt == "auto processing") | (cq == "auto processing")).map(
        {True: "Auto", False: "Manual"})
    # API-move auto-eligible category (Fly Ash / Gypsum) -> plant must be verified
    out["Auto-Eligible Cat"] = vg_l.isin(api_move).map(
        {True: "Yes (verify plant)", False: "—"})

    # Flags used by routing
    out["_is_freight"] = vg_l.str.startswith(fr_prefix) | dt_l.isin(fr_doctypes)
    out["_grey_cat"] = vg_l.map(grey_cat)
    # EMP CLAIM doc type maps to an HO-Grey category
    emp = dt_l.eq("emp claim")
    out.loc[emp & out["_grey_cat"].isna(), "_grey_cat"] = cfg["emp_claim_category"]
    return out


# ----------------------------- team routing -----------------------------
def route_team(out: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    white = {x.lower() for x in cfg["division_white_values"]}
    paints = {x.lower() for x in cfg["division_paints_values"]}
    paints_to = cfg["paints_route_to"]
    div_l = out["Division"].str.lower()

    team = pd.Series(["Plant"] * len(out), index=out.index)
    auto = out["Auto/Manual"].eq("Auto")
    is_fr = out["_is_freight"]
    is_white = div_l.isin(white)
    is_paints = div_l.isin(paints)
    has_grey_cat = out["_grey_cat"].notna()

    team = team.mask(has_grey_cat & ~is_fr, "HO-Grey")
    team = team.mask(is_white & ~is_fr, "HO-White")
    team = team.mask(is_paints & ~is_fr, paints_to)
    team = team.mask(is_fr, "Freight")
    team = team.mask(auto, "Auto/API")           # auto overrides -> API users
    # Unclassified: blank vendor group AND not routed by division
    blank_vg = out["Vendor Group"].eq("")
    team = team.mask(blank_vg & ~is_white & ~is_paints & ~is_fr & ~auto & ~has_grey_cat,
                     "Plant")  # keep in Plant but tagged below
    out["Team"] = team
    out["Status / Review"] = ""
    vg_blank = out["Vendor Group"].eq("")
    div_known = out["Division"].astype(str).str.strip().ne("")
    routable_other = out["_grey_cat"].notna() | div_known   # EMP CLAIM / division gives a route

    # Genuinely new vendor groups the master doesn't map (actionable: add a config line)
    out.loc[out["Utility"].eq("Unmapped") & out["Vendor Group"].ne(""),
            "Status / Review"] = "Unmapped vendor group"
    # Blank + nothing else to route on = waiting for Exela to finish coding (informational)
    out.loc[vg_blank & ~routable_other, "Status / Review"] = "Pending coding (Indexing)"
    # Fly Ash / Gypsum plant check (by design per master logic)
    out.loc[out["Auto-Eligible Cat"].str.startswith("Yes"),
            "Status / Review"] = "Verify auto-process plant"
    return out


# ----------------------------- role (Parker/Poster) -----------------------------
def role_for_row(out: pd.DataFrame, cfg: dict) -> pd.Series:
    """Authoritative split from old-logic Status: Parked But Not Posted -> Poster,
    Fresh Allocation -> Parker. Falls back to status/queue heuristic if Status absent."""
    if "Status" in out.columns:
        return out["Status"].map(lambda s: "Poster" if s == "Parked But Not Posted" else "Parker")
    poster_status = {x.lower() for x in cfg["poster_if_status_in"]}
    poster_queue = {x.lower() for x in cfg["poster_if_queue_in"]}
    is_poster = out["Current Status"].str.lower().isin(poster_status) | \
        out["Current Queue"].str.lower().isin(poster_queue)
    return is_poster.map({True: "Poster", False: "Parker"})


# ----------------------------- assignment -----------------------------
class RoundRobin:
    def __init__(self, names):
        self.names = list(names)
        self._it = itertools.cycle(self.names) if self.names else None
    def next(self):
        return next(self._it) if self._it else None


def _users_by_team(cfg):
    by = {}
    for u in cfg["users"]:
        by.setdefault(u["team"], []).append(u)
    return by


def _role_pool(members, role):
    """Members whose role list contains the wanted role (Parker/Poster), prime-first."""
    want = role.lower()
    pool = [m for m in members if any(want in r.lower() for r in m.get("roles", []))]
    return pool or members  # fall back to all members if none carry that role


def assign(out: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    out["Role"] = role_for_row(out, cfg)
    by_team = _users_by_team(cfg)

    assigned_name = [""] * len(out)
    assigned_email = [""] * len(out)

    # Build round-robin allocators lazily per (team, role/category) for even load.
    rr_cache: dict = {}

    def rr(key, names):
        if key not in rr_cache:
            rr_cache[key] = RoundRobin(names)
        return rr_cache[key]

    # API pool (auto) = users with any API role across teams
    api_users = [u for u in cfg["users"] if any("api" in r.lower() for r in u.get("roles", []))]

    # HO-Grey: index category -> owners
    grey_members = by_team.get("HO-Grey", [])
    def grey_owners(catcode):
        cats = set(c.strip().upper() for c in str(catcode).replace("&", ",").split(","))
        owners = [m for m in grey_members
                  if cats & set(c.strip().upper() for c in m.get("categories", []))]
        return owners or grey_members

    # Freight: split by scope (Grey/White) using division; balance within.
    freight_members = by_team.get("Freight", [])
    def freight_pool(division):
        d = division.lower()
        pool = []
        for m in freight_members:
            scope = (m.get("scope") or "").lower()
            if d == "white" and "white" in scope: pool.append(m)
            elif d != "white" and "grey" in scope: pool.append(m)
        return pool or freight_members

    for i in out.index:
        team = out.at[i, "Team"]
        role = out.at[i, "Role"]
        if team == "Auto/API":
            pool = api_users or by_team.get("Plant", [])
            m = rr(("api",), [u["name"] for u in pool]).next()
            email = next((u["ibm_email"] for u in pool if u["name"] == m), "")
        elif team == "HO-Grey":
            owners = grey_owners(out.at[i, "_grey_cat"])
            pool = _role_pool(owners, role)
            m = rr(("grey", str(out.at[i, "_grey_cat"]), role), [u["name"] for u in pool]).next()
            email = next((u["ibm_email"] for u in pool if u["name"] == m), "")
        elif team == "Freight":
            pool = freight_pool(out.at[i, "Division"])
            m = rr(("freight", out.at[i, "Division"]), [u["name"] for u in pool]).next()
            email = next((u["ibm_email"] for u in pool if u["name"] == m), "")
        else:  # Plant / HO-White
            members = by_team.get(team, [])
            pool = _role_pool(members, role)
            m = rr((team, role), [u["name"] for u in pool]).next()
            email = next((u["ibm_email"] for u in pool if u["name"] == m), "")
        assigned_name[out.index.get_loc(i)] = m or ""
        assigned_email[out.index.get_loc(i)] = email or ""

    out["Assigned To"] = assigned_name
    out["Assigned Email"] = assigned_email
    return out


def run(dump_path: str, cfg: dict | None = None, report_date=None) -> pd.DataFrame:
    import old_logic as OL
    cfg = cfg or load_config()
    df = read_dump(dump_path, cfg)
    out = classify(df, cfg)
    out = route_team(out, cfg)                       # sets flags + provisional Team
    if "old_logic" in cfg:                            # override Team with maintained VG->Team table
        out["Team"] = OL.team_via_old(out, df, cfg)
        out = OL.enrich(out, df, cfg, report_date)    # tenure/bucket/priority/status/eligibility
        out["Supervisor"] = OL.assign_supervisor(out, cfg)  # team-leader layer
    out = assign(out, cfg)                            # individual user (role from Status)
    has_old = "Tenure" in out.columns
    cols = ["DCN", "Vendor Name", "Vendor Group", "Division", "Doc Type",
            "Invoice No", "Invoice Date"]
    if has_old:
        cols += ["Invoice Receipt Date", "Tenure", "Aging Bucket", "Priority", "Status"]
    cols += ["Amount", "Utility", "MSME Flag", "Auto/Manual", "Auto-Eligible Cat"]
    if has_old:
        cols += ["Final_Tag"]
    cols += ["Current Queue", "Current Status", "Team"]
    if has_old:
        cols += ["Supervisor"]
    cols += ["Role", "Assigned To", "Assigned Email", "Claimed By", "Status / Review"]
    return out[[c for c in cols if c in out.columns]]


if __name__ == "__main__":
    import sys
    res = run(sys.argv[1])
    print("ROWS:", len(res), "| Unassigned:", (res["Assigned To"] == "").sum())
    print("\nBy Team:\n", res["Team"].value_counts())
    print("\nBy Role:\n", res["Role"].value_counts())
    if "Tenure" in res.columns:
        print("\nBy Aging Bucket:\n", res["Aging Bucket"].value_counts())
        print("\nBy Priority:\n", res["Priority"].value_counts())
        print("\nBy Status:\n", res["Status"].value_counts())
        print("\nFinal_Tag:\n", res["Final_Tag"].value_counts())
        print("\nTenure describe:\n", res["Tenure"].describe())
    print("\nStatus / Review:\n", res["Status / Review"].value_counts())
