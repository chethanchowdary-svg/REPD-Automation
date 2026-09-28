"""
UK REPD -> CoreData (DB) NEW PROJECTS builder.
Logic carried over from the Colab script; changes for the repo:
  - secrets come from config (environment), not code
  - build_new_projects() returns rows instead of writing Excel
  - the "Lead QC Changes" comment sheet is replaced by the QC upload sheet (see updater)
  - investment comment wording aligned with the updater ("$X.XX million/MW")
New = REPD row whose Ref ID is absent from DB "State_data_plant_ID".
"""
import calendar
import datetime as dt
import re
import time

import pandas as pd
import requests

from . import config as CFG

PRESENT_DATE = dt.date.today()          # main.py can override for a given run

OPENCAGE_URL = "https://api.opencagedata.com/geocode/v1/json"
OPENCAGE_MIN_INTERVAL_SECONDS = 1.05
OPENCAGE_MAX_RETRIES = 3
UK_COUNTRY_CODES = {"gb", "im", "je", "gg"}

EARLY_MONTHS_PER_50      = 8
LATE_MONTHS_PER_50       = 6
OP_MONTHS_PER_50         = 6
CONSTRUCTION_OP_PER_50   = 9
EARLY_FC_FLOOR_MONTHS    = 15
LATE_FC_FLOOR_MONTHS     = 13
FC_EST_MAX_FROM_TODAY_MO = 24
OP_MIN_MONTHS_WIND       = 6
OP_MIN_MONTHS_OTHER      = 3
OP_MAX_MONTHS            = 48
FC_BACKDATE_MAX_MONTHS   = 36
CONSTR_MIN_FUTURE_OP_MO  = 4

METRIC_RANGES = {
    ("Wind",    "dev"): (1.55, 1.63),
    ("Wind",    "op"):  (1.60, 1.80),
    ("Solar",   "dev"): (1.07, 1.14),
    ("Solar",   "op"):  (1.10, 1.40),
    ("Battery", "any"): (1.00, 1.00),
}
INVESTMENT_BILLION_THRESHOLD = 1000.0
SQM_PER_ACRE       = 4046.8564224
SITE_AREA_DECIMALS = 1
SKIP_TERMINAL_NEW  = False

ST_REPORTED  = "Reported"
ST_ESTIMATED = "Estimated"
ST_NONE      = "Not reported, need more data to estimate"

TECH_TO_SECTOR = {
    "large hydro":                     ("Hydro",   "Hydro: Large"),
    "small hydro":                     ("Hydro",   "Hydro: Small"),
    "battery":                         ("Battery", "Battery: Lithium-ion"),
    "solar photovoltaics":             ("Solar",   "Solar: Utility-scale"),
    "wind offshore":                   ("Wind",    "Wind: Offshore"),
    "wind onshore":                    ("Wind",    "Wind: Onshore"),
    "pumped storage hydroelectricity": ("Hydro",   "Hydro: Pumped Storage"),
}
TECH_TO_OFFON = {"wind offshore": "Offshore", "wind onshore": "Onshore"}
HYDRO_SECTOR = "Hydro"

BUCKET_CANCELLED   = "Cancelled"
BUCKET_EARLY       = "Early Stage"
BUCKET_LATE        = "Late Stage"
BUCKET_CONSTR      = "Under Construction"
BUCKET_OPERATIONAL = "Operational"
BUCKET_DECOMM      = "Decommissioned"

REPD_STATUS_TO_BUCKET = {
    "planning permission refused":    BUCKET_CANCELLED,
    "abandoned":                      BUCKET_CANCELLED,
    "appeal refused":                 BUCKET_CANCELLED,
    "revised":                        BUCKET_CANCELLED,
    "appeal withdrawn":               BUCKET_CANCELLED,
    "secretary of state - refusal":   BUCKET_CANCELLED,
    "planning application withdrawn": BUCKET_CANCELLED,
    "planning permission expired":    BUCKET_EARLY,
    "planning application submitted": BUCKET_EARLY,
    "appeal lodged":                  BUCKET_EARLY,
    "planning permission granted":    BUCKET_LATE,
    "no application required":        BUCKET_LATE,
    "appeal granted":                 BUCKET_LATE,
    "secretary of state - granted":   BUCKET_LATE,
    "under construction":             BUCKET_CONSTR,
    "operational":                    BUCKET_OPERATIONAL,
    "decommissioned":                 BUCKET_DECOMM,
}
BUCKET_TO_DB_LABEL = {
    BUCKET_CANCELLED:   "Cancelled",
    BUCKET_EARLY:       "Early-stage",
    BUCKET_LATE:        "Late-stage",
    BUCKET_CONSTR:      "In construction",
    BUCKET_OPERATIONAL: "Operational",
    BUCKET_DECOMM:      "Decommissioned",
}
TERMINAL_BUCKETS = {BUCKET_CANCELLED, BUCKET_DECOMM}

LIFECYCLE_FLAGS = {
    BUCKET_CANCELLED:   ("No",  "No",  "No"),
    BUCKET_EARLY:       ("No",  "No",  "No"),
    BUCKET_LATE:        ("Yes", "No",  "No"),
    BUCKET_CONSTR:      ("Yes", "Yes", "Yes - Complete"),
    BUCKET_OPERATIONAL: ("Yes", "Yes", "Yes - Complete"),
    BUCKET_DECOMM:      ("Yes", "Yes", "Yes - Complete"),
}

# REPD source columns
R_REFID    = "Ref ID"
R_OPER     = "Operator (or Applicant)"
R_SITE     = "Site Name"
R_TECH     = "Technology Type"
R_COLOC    = "Storage Co-location REPD Ref ID"
R_CAP      = "Installed Capacity (MWelec)"
R_CFDAR    = "CfD Allocation Round"
R_CFDCAP   = "CfD Capacity (MW)"
R_STATUS   = "Development Status"
R_POSTCODE = "Post Code"
R_COUNTY   = "County"
R_REGION   = "Region"
R_COUNTRY  = "Country"
R_XCOORD   = "X-coordinate"
R_YCOORD   = "Y-coordinate"
R_SUBMIT   = "Planning Application Submitted"
R_PPGRANT  = "Planning Permission  Granted"
R_UCONST   = "Under Construction"
R_OPDATE   = "Operational"
R_SITEAREA = "Solar Site Area (sqm)"

# DB output columns
C_PID     = "Project_ID"
C_SDPID   = "State_data_plant_ID"
C_NAME    = "Project_name"
C_CAP     = "Power_Capacity_(MW)"
C_SECTOR  = "Sector"
C_SUBSECT = "Sub_sector"
C_LIFE    = "Project_life_cycle"
C_REGION  = "State/Region"
C_COUNTRY = "Country"
C_COMMENT = "Analyst_comments"
C_COLOC   = "Co_located_storage"
C_COLOCAP = "Co_located_storage_capacity_(MW)"
C_CONT    = "Continent"
C_EIA     = "EIA_approval_received"
C_ENVDATE = "Environmental_approval_date"
C_FC_EST  = "Financial_close_(estimated)"
C_FC_REP  = "Financial_close_date_(reported)"
C_FC_ST   = "Financial_close_date_status"
C_FIN     = "Financing_secured"
C_FINTYPE = "Financing_type"
C_GRID    = "Grid_connection_approved"
C_OP_EST  = "Operational_date_(estimated)"
C_OP_REP  = "Operational_date_(reported)"
C_OP_ST   = "Operational_date_status"
C_PSA     = "Power_sales_arrangement"
C_PPA     = "PPA_signed"
C_INV_EST = "Project_investment_($M)___Estimated"
C_INV_REP = "Project_investment_($M)___Reported"
C_INV_ST  = "Project_investment_status"
C_SHOWCOM = "Show_analyst_comment"
C_SHOWSPV = "Show_SPV"
C_VIEWER  = "Show_to_viewer"
C_AREA    = "Site_area_(acres)"
C_SOLAR   = "Solar_technology"
C_SPV     = "SPV"
C_SUBLOC  = "Sub_location"
C_LAT     = "Latitude"
C_LON     = "Longitude"
C_TERR    = "Territory_(UK)"
C_CFDCAP  = "UK_CFD_Allocated_Capacity_(MW)"
C_CFDAR   = "UK_CFD_Allocation_Round_(AR)"
C_OFFON   = "Offshore/Onshore"

OUT_COLS = [
    C_PID, C_SDPID, C_NAME, C_CAP, C_SECTOR, C_LIFE, C_REGION, C_COUNTRY,
    C_COMMENT, C_COLOC, C_COLOCAP, C_CONT, C_EIA, C_ENVDATE,
    C_FC_EST, C_FC_REP, C_FC_ST, C_FIN, C_FINTYPE, C_GRID,
    C_OP_EST, C_OP_REP, C_OP_ST, C_PSA, C_PPA,
    C_INV_EST, C_INV_REP, C_INV_ST, C_SHOWCOM, C_SHOWSPV, C_VIEWER,
    C_AREA, C_SOLAR, C_SPV, C_SUBLOC, C_LAT, C_LON, C_SUBSECT, C_TERR,
    C_CFDCAP, C_CFDAR,
]


# =============================================================================
# Helpers
# =============================================================================
def _blank(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""


def normalize_multiselect(v):
    from .updater import normalize_multiselect as _n
    return _n(v)


def parse_capacity(v):
    if _blank(v):
        return None
    s = re.sub(r"[^0-9.]", "", str(v))
    try:
        return float(s) if s else None
    except ValueError:
        return None


def parse_repd_date(v):
    if _blank(v):
        return None
    s = str(v).strip()
    m = re.match(r"/Date\((-?\d+)\)/", s)
    if m:
        try:
            return (dt.datetime(1970, 1, 1) + dt.timedelta(milliseconds=int(m.group(1)))).date()
        except Exception:
            return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def fmt_db_date(d):
    return d.strftime("%m-%d-%Y") if d else ""


def fmt_month_year(d):
    return d.strftime("%B %Y") if d else ""


def add_months(d, n):
    n = int(round(n))
    total = (d.year * 12 + (d.month - 1)) + n
    y, m = divmod(total, 12)
    m += 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def end_of_month(d):
    return dt.date(d.year, d.month, calendar.monthrange(d.year, d.month)[1])


def eom_add(a, m):
    return end_of_month(add_months(a, m))


def months_for_capacity(cap, p50):
    return (cap / 50.0) * p50


def estimate_fc_forward(anchor, cap, p50, floor_m):
    cand = eom_add(anchor, months_for_capacity(cap, p50))
    floor = eom_add(PRESENT_DATE, floor_m)
    cap_hi = eom_add(PRESENT_DATE, FC_EST_MAX_FROM_TODAY_MO)
    return min(max(cand, floor), cap_hi)


def estimate_op_from_fc(fc, cap, sector, p50=OP_MONTHS_PER_50):
    cand = eom_add(fc, months_for_capacity(cap, p50))
    lo = eom_add(fc, OP_MIN_MONTHS_WIND if sector == "Wind" else OP_MIN_MONTHS_OTHER)
    hi = eom_add(fc, OP_MAX_MONTHS)
    return min(max(cand, lo), hi)


def backdate_fc_from_op(op, cap):
    return eom_add(op, -min(months_for_capacity(cap, OP_MONTHS_PER_50), FC_BACKDATE_MAX_MONTHS))


def metric_for(sector, stage):
    key = (sector, "any") if sector == "Battery" else (sector, stage)
    rng = METRIC_RANGES.get(key)
    if rng is None:
        return None
    lo, hi = rng
    return round((lo + hi) / 2.0, 2)


def compute_investment(sector, bucket, cap):
    if sector == HYDRO_SECTOR or cap is None or cap <= 0:
        return None, None
    stage = ("op" if bucket == BUCKET_OPERATIONAL
             else "dev" if bucket in (BUCKET_EARLY, BUCKET_LATE, BUCKET_CONSTR) else None)
    if stage is None:
        return None, None
    metric = metric_for(sector, stage)
    if metric is None:
        return None, None
    inv = round(metric * cap, 1)
    return (None, None) if inv <= 0 else (inv, metric)


def sector_phrase(sector, offon):
    if sector == "Wind":
        oo = (offon or "").strip().lower()
        return "offshore wind" if oo == "offshore" else "onshore wind" if oo == "onshore" else "wind"
    return sector.lower()


def money_phrase(m):
    if m is None:
        return ""
    return f"${m / 1000.0:.1f} billion" if m > INVESTMENT_BILLION_THRESHOLD else f"${m:,.0f} million"


def build_comment(sect, inv_est, inv_m, metric, fc, fc_est, op, op_est, is_op):
    blocks = []
    if inv_est and inv_m is not None:
        blocks.append(
            "Estimated Project Investment:\n"
            f"-- Enerdatics calculates the project's investment to be {money_phrase(inv_m)}, "
            f"representing an investment metric of ${metric:.2f} million/MW.\n"
            f"-- This estimate is in line with reported figures for {sect} projects "
            "of this scale and size in the region.")
    timeline = None
    if op_est and fc_est:
        timeline = (
            "Estimated Development Timeline:\n"
            f"-- The project is estimated to be commercially operational by {fmt_month_year(op)}.\n"
            "-- Further, Enerdatics believes that the developer will have secured financing to fund "
            f"construction activities on the project by {fmt_month_year(fc)}.\n"
            f"-- This estimate is based on reported development timelines for {sect} projects "
            "of this scale and size in the region.")
    elif op_est and not fc_est and not is_op:
        timeline = (
            "Estimated Development Timeline:\n"
            f"-- The project is estimated to be commercially operational by {fmt_month_year(op)}.\n"
            f"-- This estimate is based on reported development timelines for {sect} projects "
            "of this scale and size in the region.\n"
            "-- Further, The developer will secure financing to fund construction activities "
            f"on the project by {fmt_month_year(fc)}.")
    elif is_op and fc_est:
        timeline = (
            "Estimated Development Timeline:\n"
            f"-- The project reached commercial operations in {fmt_month_year(op)}.\n"
            "-- Further, Enerdatics believes that developer would have secured financing to fund "
            f"construction activities on the project by {fmt_month_year(fc)}.\n"
            f"-- This estimate is based on reported development timelines for {sect} projects "
            "of this scale and size in the region.")
    if timeline:
        blocks.append(timeline)
    return "\n\n".join(blocks)


# =============================================================================
# Co-location helpers
# =============================================================================
def split_ids(s):
    return [t.strip() for t in re.split(r"[;,&/]|\bOR\b|\band\b", str(s), flags=re.I) if t.strip().isdigit()]


def coloc_fields(coloc_raw, this_tech, repd_idx):
    ids = split_ids(coloc_raw)
    if not ids:
        return "No", ""
    yes, total = False, None
    for cid in ids:
        rr = repd_idx.get(cid)
        if rr is None:
            continue
        if (rr.get(R_TECH) or "").strip().lower() in TECH_TO_SECTOR:
            yes = True
        cap = parse_capacity(rr.get(R_CAP))
        if cap is not None:
            total = (total or 0) + cap
    coloc = "Yes" if yes else "No"
    if this_tech.strip().lower() == "battery":
        return coloc, ""
    return coloc, ("" if total is None else f"{total:g}")


# =============================================================================
# Location helpers (X/Y -> lat/lon -> OpenCage, postcodes.io fallback)
# =============================================================================
def clean_name(s):
    s = (s or "").strip()
    if not s:
        return ""
    s = re.sub(r"\boffshore\b", "", s, flags=re.I)
    s = re.split(r"[/\\&,;]", s)[0]
    s = re.sub(r"^\s*Co\.?\s+", "", s, flags=re.I)
    return s.strip(" .-")


def _geo_from_pc_result(res):
    subloc = res.get("admin_district") or res.get("parish") or ""
    region = res.get("region") or res.get("admin_county") or res.get("admin_district") or ""
    territory = res.get("country") or ""
    return clean_name(subloc), clean_name(region), clean_name(territory)


def _bulk_postcode_lookup(postcodes, batch_size=100):
    results = [None] * len(postcodes)
    idx_map = {i: pc.strip().upper() for i, pc in enumerate(postcodes) if pc and pc.strip()}
    if not idx_map:
        return results
    positions = list(idx_map.keys())
    pc_list = [idx_map[i] for i in positions]
    for start in range(0, len(pc_list), batch_size):
        batch_pos, batch_pc = positions[start:start + batch_size], pc_list[start:start + batch_size]
        try:
            resp = requests.post("https://api.postcodes.io/postcodes", json={"postcodes": batch_pc},
                                 headers={"Content-Type": "application/json"}, timeout=30)
            if not resp.ok:
                continue
            for i, item in enumerate(resp.json().get("result", [])):
                res = item.get("result")
                if res:
                    results[batch_pos[i]] = _geo_from_pc_result(res)
        except Exception:
            continue
    return results


_BNG_TRANSFORMER = None
_OPENCAGE_LAST_REQUEST_AT = 0.0


def _bng_to_latlon(easting, northing):
    global _BNG_TRANSFORMER
    try:
        if _blank(easting) or _blank(northing):
            return None, None
        if _BNG_TRANSFORMER is None:
            from pyproj import Transformer
            _BNG_TRANSFORMER = Transformer.from_crs("EPSG:27700", "EPSG:4326", always_xy=True)
        lon, lat = _BNG_TRANSFORMER.transform(float(easting), float(northing))
        if not (44.0 <= lat <= 66.0 and -16.0 <= lon <= 12.0):
            return None, None
        return round(lat, 6), round(lon, 6)
    except Exception:
        return None, None


def _first_component(components, keys):
    for key in keys:
        value = clean_name(components.get(key))
        if value:
            return value
    return ""


UK_TERRITORIES = {"england", "scotland", "wales", "northern ireland", "isle of man", "jersey", "guernsey"}


def _geo_from_opencage_result(result):
    components = result.get("components") or {}
    country_code = str(components.get("country_code") or components.get("ISO_3166-1_alpha-2") or "").strip().lower()
    territory = _first_component(components, ("state", "province"))
    if not territory:
        country_name = clean_name(components.get("country"))
        if country_name.strip().lower() in UK_TERRITORIES:
            territory = country_name
    region = _first_component(components, ("state_district", "county", "region"))
    subloc = _first_component(components, ("city", "town", "village", "municipality", "hamlet", "locality",
                                           "city_district", "district", "suburb", "quarter", "neighbourhood"))
    if not (country_code in UK_COUNTRY_CODES or territory.strip().lower() in UK_TERRITORIES):
        return None
    return subloc, region, territory


def _opencage_reverse(lat, lon):
    global _OPENCAGE_LAST_REQUEST_AT
    if not CFG.OPENCAGE_API_KEY:
        return None
    params = {"q": f"{lat:.6f},{lon:.6f}", "key": CFG.OPENCAGE_API_KEY, "language": "en",
              "no_annotations": 1, "no_record": 1, "address_only": 1}
    for attempt in range(OPENCAGE_MAX_RETRIES):
        try:
            wait_for = OPENCAGE_MIN_INTERVAL_SECONDS - (time.monotonic() - _OPENCAGE_LAST_REQUEST_AT)
            if wait_for > 0:
                time.sleep(wait_for)
            resp = requests.get(OPENCAGE_URL, params=params, timeout=30)
            _OPENCAGE_LAST_REQUEST_AT = time.monotonic()
            if resp.status_code == 429:
                time.sleep(OPENCAGE_MIN_INTERVAL_SECONDS * (attempt + 1))
                continue
            if resp.status_code >= 500:
                time.sleep(1.5 * (attempt + 1))
                continue
            if not resp.ok:
                return None
            results = resp.json().get("results") or []
            return _geo_from_opencage_result(results[0]) if results else None
        except Exception:
            if attempt + 1 < OPENCAGE_MAX_RETRIES:
                time.sleep(1.5 * (attempt + 1))
    return None


def build_location_cache(repd_df):
    """{ref_id: (subloc, region, territory, lat, lon)} via X/Y -> OpenCage, postcodes.io fallback."""
    rows = repd_df.to_dict("records")
    cache, refs_by_latlon = {}, {}
    for rr in rows:
        ref = str(rr.get(R_REFID, "")).strip()
        lat, lon = _bng_to_latlon(rr.get(R_XCOORD), rr.get(R_YCOORD))
        cache[ref] = (clean_name(rr.get(R_COUNTY)), clean_name(rr.get(R_REGION)),
                      clean_name(rr.get(R_COUNTRY)), lat, lon)
        if lat is not None and lon is not None:
            refs_by_latlon.setdefault((lat, lon), []).append(ref)
    print(f"  OpenCage reverse geocoding: {len(refs_by_latlon)} unique coordinates")
    for i, (lat, lon) in enumerate(refs_by_latlon, start=1):
        geo = _opencage_reverse(lat, lon)
        if geo:
            for ref in refs_by_latlon[(lat, lon)]:
                old = cache[ref]
                cache[ref] = (geo[0] or old[0], geo[1] or old[1], geo[2] or old[2], old[3], old[4])
        if i % 100 == 0:
            print(f"    {i}/{len(refs_by_latlon)}")
    fb_pos, fb_pc = [], []
    for pos, rr in enumerate(rows):
        ref = str(rr.get(R_REFID, "")).strip()
        sub, reg = cache[ref][0], cache[ref][1]
        pc = str(rr.get(R_POSTCODE) or "").strip()
        if not (sub and reg) and pc:
            fb_pos.append(pos)
            fb_pc.append(pc)
    if fb_pc:
        for pos, geo in zip(fb_pos, _bulk_postcode_lookup(fb_pc)):
            if geo:
                ref = str(rows[pos].get(R_REFID, "")).strip()
                s, r, t, la, lo = cache[ref]
                cache[ref] = (s or geo[0], r or geo[1], t or geo[2], la, lo)
    return cache


def apply_location_rules(subloc, region, territory):
    if region and territory and region.lower() == territory.lower():
        region, subloc = subloc, ""
    if region and subloc and region.lower() == subloc.lower():
        subloc = ""
    if not region and subloc:
        region, subloc = subloc, ""
    return subloc, region, territory


# =============================================================================
# Build one new DB row (unchanged rules)
# =============================================================================
def build_row(rrow, repd_idx, loc_cache=None):
    tech = (rrow.get(R_TECH) or "").strip()
    sector, subsector = TECH_TO_SECTOR[tech.lower()]
    offon = TECH_TO_OFFON.get(tech.lower(), "")
    cap = parse_capacity(rrow.get(R_CAP))

    row = {c: "" for c in OUT_COLS}
    row[C_OFFON] = offon
    row[C_SDPID] = str(rrow.get(R_REFID) or "").strip()
    row[C_NAME] = (rrow.get(R_SITE) or "").strip()
    row[C_CAP] = "" if cap is None else f"{cap:g}"
    row[C_SECTOR], row[C_SUBSECT] = sector, subsector
    row[C_CONT], row[C_COUNTRY] = "Europe", "United Kingdom"

    ref = row[C_SDPID]
    raw = loc_cache.get(ref, ("", "", "", None, None)) if loc_cache else ("", "", "", None, None)
    subloc, region, territory = apply_location_rules(raw[0], raw[1], raw[2])
    row[C_SUBLOC], row[C_REGION], row[C_TERR] = subloc, region, territory
    row[C_LAT] = "" if raw[3] is None else raw[3]
    row[C_LON] = "" if raw[4] is None else raw[4]

    spv = (rrow.get(R_OPER) or "").strip()
    row[C_SPV], row[C_SHOWSPV] = spv, ("Yes" if spv else "No")
    ar = (rrow.get(R_CFDAR) or "").strip()
    row[C_CFDAR], row[C_CFDCAP] = ar, (rrow.get(R_CFDCAP) or "").strip()
    if sector == "Solar":
        row[C_SOLAR] = "PV"
    row[C_ENVDATE] = fmt_db_date(parse_repd_date(rrow.get(R_PPGRANT)))
    sqm = parse_capacity(rrow.get(R_SITEAREA))
    if sqm is not None and sqm > 0:
        row[C_AREA] = f"{sqm / SQM_PER_ACRE:,.{SITE_AREA_DECIMALS}f}"
    row[C_COLOC], row[C_COLOCAP] = coloc_fields(rrow.get(R_COLOC), tech, repd_idx)

    bucket = REPD_STATUS_TO_BUCKET.get((rrow.get(R_STATUS) or "").strip().lower())
    if bucket is None:
        for s in (C_INV_ST, C_OP_ST, C_FC_ST):
            row[s] = ST_NONE
        row[C_SHOWCOM] = "No"
        return row, bucket

    row[C_LIFE] = BUCKET_TO_DB_LABEL[bucket]
    eia, grid, fin = LIFECYCLE_FLAGS.get(bucket, ("", "", ""))
    row[C_EIA], row[C_GRID], row[C_FIN] = eia, grid, fin
    row[C_FINTYPE] = "Debt; Equity" if fin == "Yes - Complete" else ""
    if ar:
        row[C_PSA] = "Contracts-for-Difference"
    row[C_PPA] = ("Yes - Complete" if (bucket == BUCKET_OPERATIONAL or row[C_PSA] == "Contracts-for-Difference")
                  else "No")

    if bucket in TERMINAL_BUCKETS:
        for s in (C_INV_ST, C_OP_ST, C_FC_ST):
            row[s] = ST_NONE
        row[C_SHOWCOM] = "No"
        return row, bucket

    repd_op = parse_repd_date(rrow.get(R_OPDATE))
    repd_fc = parse_repd_date(rrow.get(R_UCONST))

    if sector == HYDRO_SECTOR:
        if repd_op:
            row[C_OP_REP], row[C_OP_ST] = fmt_db_date(repd_op), ST_REPORTED
        else:
            row[C_OP_ST] = ST_NONE
        if repd_fc:
            row[C_FC_REP], row[C_FC_ST] = fmt_db_date(repd_fc), ST_REPORTED
        else:
            row[C_FC_ST] = ST_NONE
        row[C_INV_ST], row[C_SHOWCOM] = ST_NONE, "No"
        return row, bucket

    if repd_op:
        row[C_OP_REP], row[C_OP_ST] = fmt_db_date(repd_op), ST_REPORTED
    if repd_fc:
        row[C_FC_REP], row[C_FC_ST] = fmt_db_date(repd_fc), ST_REPORTED
    op_status, fc_status = row[C_OP_ST], row[C_FC_ST]
    is_operational = bucket == BUCKET_OPERATIONAL or (repd_op is not None and repd_op <= PRESENT_DATE)

    if cap is None or cap <= 0:
        for s in (C_INV_ST, C_OP_ST, C_FC_ST):
            if row[s] == "":
                row[s] = ST_NONE
        row[C_SHOWCOM] = "No"
        return row, bucket

    fc_date, fc_estimated = repd_fc, (fc_status != ST_REPORTED)
    if fc_status != ST_REPORTED:
        if bucket == BUCKET_OPERATIONAL and repd_op:
            fc_date = backdate_fc_from_op(repd_op, cap)
        elif bucket == BUCKET_CONSTR:
            fc_date = estimate_fc_forward(PRESENT_DATE, cap, LATE_MONTHS_PER_50, LATE_FC_FLOOR_MONTHS)
        elif bucket == BUCKET_EARLY:
            anchor = parse_repd_date(rrow.get(R_SUBMIT)) or PRESENT_DATE
            fc_date = estimate_fc_forward(anchor, cap, EARLY_MONTHS_PER_50, EARLY_FC_FLOOR_MONTHS)
        elif bucket == BUCKET_LATE:
            anchor = parse_repd_date(rrow.get(R_SUBMIT)) or PRESENT_DATE
            fc_date = estimate_fc_forward(anchor, cap, LATE_MONTHS_PER_50, LATE_FC_FLOOR_MONTHS)
        if fc_date:
            row[C_FC_EST], row[C_FC_ST], fc_estimated = fmt_db_date(fc_date), ST_ESTIMATED, True

    op_date, op_estimated = repd_op, (op_status != ST_REPORTED)
    if op_status != ST_REPORTED:
        if bucket == BUCKET_CONSTR:
            base_fc = fc_date or estimate_fc_forward(PRESENT_DATE, cap, LATE_MONTHS_PER_50, LATE_FC_FLOOR_MONTHS)
            cand = estimate_op_from_fc(base_fc, cap, sector, p50=CONSTRUCTION_OP_PER_50)
            min_fut = eom_add(PRESENT_DATE, CONSTR_MIN_FUTURE_OP_MO)
            if cand < min_fut:
                back = min(12, max(3, round(months_for_capacity(cap, 3))))
                base_fc = eom_add(PRESENT_DATE, -back)
                cand = estimate_op_from_fc(base_fc, cap, sector, p50=CONSTRUCTION_OP_PER_50)
                if cand < min_fut:
                    cand = min_fut
                if fc_status != ST_REPORTED:
                    row[C_FC_EST], row[C_FC_ST] = fmt_db_date(base_fc), ST_ESTIMATED
                    fc_date, fc_estimated = base_fc, True
            op_date = cand
        elif fc_date:
            op_date = estimate_op_from_fc(fc_date, cap, sector)
        if op_date and bucket != BUCKET_OPERATIONAL:
            row[C_OP_EST], row[C_OP_ST], op_estimated = fmt_db_date(op_date), ST_ESTIMATED, True
        elif bucket == BUCKET_OPERATIONAL and not repd_op:
            row[C_OP_ST] = row[C_FC_ST] = ST_NONE
            row[C_OP_EST] = row[C_OP_REP] = row[C_FC_EST] = row[C_FC_REP] = ""
            op_estimated = fc_estimated = False
            op_date = fc_date = None

    inv_m, metric = compute_investment(sector, bucket, cap)
    if inv_m is not None:
        row[C_INV_EST], row[C_INV_ST] = f"{inv_m:,.0f}", ST_ESTIMATED
    else:
        row[C_INV_ST] = ST_NONE

    comment = build_comment(sector_phrase(sector, offon), inv_m is not None, inv_m, metric,
                            fc_date, fc_estimated, op_date, op_estimated, is_operational)
    row[C_COMMENT], row[C_SHOWCOM] = (comment, "Yes") if comment.strip() else ("", "No")
    return row, bucket


# =============================================================================
# Driver
# =============================================================================
def first_tok(s):
    for tok in re.split(r"[;,]|\bOR\b|/", str(s)):
        tok = tok.strip()
        if tok.isdigit():
            return tok
    return None


def build_new_projects(db_ids: set[str], repd: pd.DataFrame) -> tuple[list[dict], dict]:
    """New-project rows for REPD Ref IDs not in `db_ids`, plus run stats."""
    repd_idx = {str(r[R_REFID]).strip(): r for _, r in repd.iterrows()}
    cands, stats = [], {"REPD rows": len(repd), "Ignored (sector)": 0,
                        "Skipped (already in DB)": 0, "Skipped (capacity blank/<1 MW)": 0}
    for _, rr in repd.iterrows():
        if (rr.get(R_TECH) or "").strip().lower() not in TECH_TO_SECTOR:
            stats["Ignored (sector)"] += 1
            continue
        if str(rr.get(R_REFID) or "").strip() in db_ids:
            stats["Skipped (already in DB)"] += 1
            continue
        cap = parse_capacity(rr.get(R_CAP))
        if cap is None or cap < 1:
            stats["Skipped (capacity blank/<1 MW)"] += 1
            continue
        cands.append(rr.to_dict())
    loc_cache = build_location_cache(pd.DataFrame(cands, columns=repd.columns))
    rows = []
    for rr in cands:
        row, bucket = build_row(rr, repd_idx, loc_cache)
        if SKIP_TERMINAL_NEW and bucket in TERMINAL_BUCKETS:
            continue
        for col in (C_PSA, C_CFDAR):
            row[col] = normalize_multiselect(row[col])
        rows.append(row)
    return rows, stats


def db_id_set(db: pd.DataFrame) -> set[str]:
    """Every REPD Ref ID already in the DB (all tokens of multi-ID cells, not just the first)."""
    col = "State_data_plant_ID" if "State_data_plant_ID" in db.columns else "State data plant ID"
    ids = set()
    for v in db.get(col, pd.Series(dtype=str)):
        ids.update(t.strip() for t in re.split(r"[;,]|\bOR\b|/", str(v)) if t.strip().isdigit())
    return ids
