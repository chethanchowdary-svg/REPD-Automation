"""
UK REPD -> CoreData (DB) quarterly project UPDATER.
Logic carried over from the Colab script; changes for the repo:
  - secrets come from config (environment), not code
  - run_update() returns data instead of writing Excel
  - Lead QC rows are split into their own upload list
Join key: DB "State_data_plant_ID" == REPD "Ref ID".
"""
import ast
import calendar
import datetime as dt
import re
import time

import pandas as pd
import requests

from . import config as CFG

PRESENT_DATE = dt.date.today()          # main.py can override for a given run

OPENCAGE_MIN_INTERVAL_SEC = 1.05
OPENCAGE_TIMEOUT_SEC = 30

# ---- Date-estimate rates ----
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
MIN_ESTIMATED_INVESTMENT_M   = 1.01

SQM_PER_ACRE        = 4046.8564224
OVERWRITE_SITE_AREA = True
SITE_AREA_DECIMALS  = 1

ST_REPORTED  = "Reported"
ST_ESTIMATED = "Estimated"
ST_NONE      = "Not reported, need more data to estimate"

# =============================================================================
# DB column names
# =============================================================================
C_JOIN    = "State_data_plant_ID"
C_NAME    = "Project_name"
C_CAP     = "Power_Capacity_(MW)"
C_SECTOR  = "Sector"
C_LIFE    = "Project_life_cycle"
C_OFFON   = "Offshore/Onshore"
C_INV_ST  = "Project_investment_status"
C_INV_EST = "Project_investment_($M)___Estimated"
C_INV_REP = "Project_investment_($M)___Reported"
C_OP_ST   = "Operational_date_status"
C_OP_EST  = "Operational_date_(estimated)"
C_OP_REP  = "Operational_date_(reported)"
C_FC_ST   = "Financial_close_date_status"
C_FC_EST  = "Financial_close_(estimated)"
C_FC_REP  = "Financial_close_date_(reported)"
C_SITEAREA = "Site_area_(acres)"
C_ENVDATE  = "Environmental_approval_date"
C_COMMENT = "Analyst_comments"
C_SHOWCOM = "Show_analyst_comment"
C_EIA     = "EIA_approval_received"
C_GRID    = "Grid_connection_approved"
C_FIN     = "Financing_secured"
C_FINTYPE = "Financing_type"
C_PPA     = "PPA_signed"
C_PSA     = "Power_sales_arrangement"
C_REVMECH = "Revenue_mechanism"
C_COLOC   = "Co_located_storage"
C_CFDAR   = "UK_CFD_Allocation_Round_(AR)"
C_SPV     = "SPV"
C_SHOWSPV = "Show_SPV"
C_REGION  = "State/Region"
C_SUBLOC  = "Sub_location"

REVMECH_MANAGED_VALUES = {"not applicable", "pending finalization", "undisclosed"}

OFFSHORE_WIND_METRIC_MIN = 1.85
OFFSHORE_WIND_METRIC_MAX = 2.90
OFFSHORE_WIND_CAP_MIN_MW = 50
OFFSHORE_WIND_CAP_MAX_MW = 1200
OFFSHORE_WIND_OP_UPLIFT  = 0.10

# REPD column names
R_REFID    = "Ref ID"
R_OPERATOR = "Operator (or Applicant)"
R_STATUS   = "Development Status"
R_SUBMIT   = "Planning Application Submitted"
R_UCONST   = "Under Construction"
R_OPER     = "Operational"
R_SITEAREA = "Solar Site Area (sqm)"
R_PPGRANT  = "Planning Permission  Granted"      # two spaces
R_APPEAL_GRANTED = "Appeal Granted"
R_XCOORD   = "X-coordinate"
R_YCOORD   = "Y-coordinate"
R_CAP      = "Installed Capacity (MWelec)"

DATE_COLS = [C_OP_EST, C_OP_REP, C_FC_EST, C_FC_REP, C_ENVDATE]

# =============================================================================
# Lifecycle buckets
# =============================================================================
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
DB_LABEL_TO_BUCKET = {
    "cancelled":             BUCKET_CANCELLED,
    "suspended":             BUCKET_CANCELLED,
    "decommissioned":        BUCKET_DECOMM,
    "early-stage":           BUCKET_EARLY,
    "late-stage":            BUCKET_LATE,
    "ready-to-build":        BUCKET_LATE,
    "in construction":       BUCKET_CONSTR,
    "construction complete": BUCKET_CONSTR,
    "operational":           BUCKET_OPERATIONAL,
}
TERMINAL_BUCKETS     = {BUCKET_CANCELLED, BUCKET_DECOMM}
BLANKCAP_WIPE_LABELS = {"suspended", "cancelled"}

LIFECYCLE_FLAGS = {
    BUCKET_CANCELLED:   ("No",  "No",  "No"),
    BUCKET_EARLY:       ("No",  "No",  "No"),
    BUCKET_LATE:        ("Yes", "No",  "No"),
    BUCKET_CONSTR:      ("Yes", "Yes", "Yes - Complete"),
    BUCKET_OPERATIONAL: ("Yes", "Yes", "Yes - Complete"),
    BUCKET_DECOMM:      ("Yes", "Yes", "Yes - Complete"),
}

HYDRO_SECTOR   = "Hydro"
BATTERY_SECTOR = "Battery"
LIFECYCLE_CAPACITY_CHANGE_COLS = {C_LIFE, C_CAP}
SITE_AREA_CHANGE_COLS = {C_SITEAREA}

OLD_TO_NEW = {
    "State data plant ID":                 "State_data_plant_ID",
    "Project name":                        "Project_name",
    "Power Capacity (MW)":                 "Power_Capacity_(MW)",
    "Project life cycle":                  "Project_life_cycle",
    "Project investment status":           "Project_investment_status",
    "Project investment ($M) - Estimated": "Project_investment_($M)___Estimated",
    "Project investment ($M) - Reported":  "Project_investment_($M)___Reported",
    "Operational date status":             "Operational_date_status",
    "Operational date (estimated)":        "Operational_date_(estimated)",
    "Operational date (reported)":         "Operational_date_(reported)",
    "Financial close date status":         "Financial_close_date_status",
    "Financial close (estimated)":         "Financial_close_(estimated)",
    "Financial close date (reported)":     "Financial_close_date_(reported)",
    "Site area (acres)":                   "Site_area_(acres)",
    "Analyst comments":                    "Analyst_comments",
    "Show analyst comments":               "Show_analyst_comment",
    "Show analyst comment":                "Show_analyst_comment",
    "EIA approval received":               "EIA_approval_received",
    "Grid connection approved":            "Grid_connection_approved",
    "Financing secured":                   "Financing_secured",
    "Financing type":                      "Financing_type",
    "PPA signed":                          "PPA_signed",
    "Power sales arrangement":             "Power_sales_arrangement",
    "Show SPV":                            "Show_SPV",
    "Sub-sector":                          "Sub_sector",
    "Sub-location":                        "Sub_location",
    "State/Region":                        "State/Region",
    "Territory (UK)":                      "Territory_(UK)",
    "UK CFD Allocation Round (AR)":        "UK_CFD_Allocation_Round_(AR)",
    "UK CFD Allocated Capacity (MW)":      "UK_CFD_Allocated_Capacity_(MW)",
    "Co-located storage":                  "Co_located_storage",
    "Co-located storage capacity (MW)":    "Co_located_storage_capacity_(MW)",
    "Solar technology":                    "Solar_technology",
    "Environmental approval date":         "Environmental_approval_date",
    "Show to viewer":                      "Show_to_viewer",
}


# =============================================================================
# Helpers
# =============================================================================
def _blank(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""


def normalize_multiselect(v):
    if v is None:
        return ""
    if isinstance(v, (list, tuple)):
        items = list(v)
    else:
        s = str(v).strip()
        if s in ("", "[]"):
            return ""
        if s.startswith("[") and s.endswith("]"):
            items = None
            try:
                parsed = ast.literal_eval(s)
                if isinstance(parsed, (list, tuple)):
                    items = list(parsed)
            except Exception:
                items = None
            if items is None:
                items = s[1:-1].split(",")
        else:
            return s
    cleaned = [str(it).strip().strip("'\"").strip() for it in items if str(it).strip().strip("'\"").strip()]
    if not cleaned:
        return ""
    return cleaned[0] if len(cleaned) == 1 else ";".join(cleaned)


def parse_capacity(v):
    if _blank(v):
        return None
    s = re.sub(r"[^0-9.]", "", str(v))
    try:
        return float(s) if s else None
    except ValueError:
        return None


parse_money = parse_capacity


def parse_any_date(v):
    if _blank(v):
        return None
    s = str(v).strip()
    m = re.match(r"/Date\((-?\d+)\)/", s)
    if m:
        try:
            return (dt.datetime(1970, 1, 1) + dt.timedelta(milliseconds=int(m.group(1)))).date()
        except Exception:
            return None
    for fmt in ("%m-%d-%Y", "%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


parse_db_date = parse_any_date


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
    return d.strftime("%m/%d/%Y") if d else ""


def normalise_date_value(v):
    if _blank(v):
        return ""
    d = parse_any_date(v)
    return fmt_db_date(d) if d else str(v).strip()


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


def eom_add(anchor, months):
    return end_of_month(add_months(anchor, months))


def find_revmech_column(columns):
    for c in columns:
        norm = " ".join(str(c).lower().replace("_", " ").replace("-", " ").split())
        if norm in ("revenue mechanism", "revenue mechanisms"):
            return c
    return None


def find_column(columns, *names):
    want = {" ".join(n.lower().replace("_", " ").split()) for n in names}
    for c in columns:
        if " ".join(str(c).lower().replace("_", " ").split()) in want:
            return c
    return None


# =============================================================================
# Estimation logic
# =============================================================================
def months_for_capacity(cap_mw, per_50):
    return (cap_mw / 50.0) * per_50


def estimate_fc_forward(anchor, cap_mw, per_50, floor_months):
    cand = eom_add(anchor, months_for_capacity(cap_mw, per_50))
    floor = eom_add(PRESENT_DATE, floor_months)
    cap_hi = eom_add(PRESENT_DATE, FC_EST_MAX_FROM_TODAY_MO)
    return min(max(cand, floor), cap_hi)


def estimate_op_from_fc(fc, cap_mw, sector, per_50=OP_MONTHS_PER_50):
    cand = eom_add(fc, months_for_capacity(cap_mw, per_50))
    min_mo = OP_MIN_MONTHS_WIND if sector == "Wind" else OP_MIN_MONTHS_OTHER
    return min(max(cand, eom_add(fc, min_mo)), eom_add(fc, OP_MAX_MONTHS))


def backdate_fc_from_op(op, cap_mw):
    months = min(months_for_capacity(cap_mw, OP_MONTHS_PER_50), FC_BACKDATE_MAX_MONTHS)
    return eom_add(op, -months)


def metric_for(sector, stage, cap_mw):
    key = (sector, "any") if sector == "Battery" else (sector, stage)
    rng = METRIC_RANGES.get(key)
    if rng is None:
        return None
    lo, hi = rng
    return round((lo + hi) / 2.0, 2)


def offshore_wind_metric(cap_mw, stage):
    if cap_mw <= OFFSHORE_WIND_CAP_MIN_MW:
        metric = OFFSHORE_WIND_METRIC_MAX
    elif cap_mw >= OFFSHORE_WIND_CAP_MAX_MW:
        metric = OFFSHORE_WIND_METRIC_MIN
    else:
        t = (cap_mw - OFFSHORE_WIND_CAP_MIN_MW) / (OFFSHORE_WIND_CAP_MAX_MW - OFFSHORE_WIND_CAP_MIN_MW)
        metric = OFFSHORE_WIND_METRIC_MAX - t * (OFFSHORE_WIND_METRIC_MAX - OFFSHORE_WIND_METRIC_MIN)
    if stage == "op":
        metric += OFFSHORE_WIND_OP_UPLIFT
    floor = OFFSHORE_WIND_METRIC_MIN + (OFFSHORE_WIND_OP_UPLIFT if stage == "op" else 0)
    return round(max(metric, floor), 4)


def compute_investment(sector, bucket, cap_mw, offon=""):
    if sector == HYDRO_SECTOR:
        return None, None
    if cap_mw is None or cap_mw <= 0:
        return None, None
    stage = ("op" if bucket == BUCKET_OPERATIONAL
             else "dev" if bucket in (BUCKET_EARLY, BUCKET_LATE, BUCKET_CONSTR) else None)
    if stage is None:
        return None, None
    if sector == "Wind" and (offon or "").strip().lower() == "offshore":
        metric = offshore_wind_metric(cap_mw, stage)
    else:
        metric = metric_for(sector, stage, cap_mw)
        if metric is None:
            return None, None
    inv = round(metric * cap_mw, 2)
    return max(inv, MIN_ESTIMATED_INVESTMENT_M), metric


# =============================================================================
# Analyst comment builder
# =============================================================================
def sector_phrase(sector, offon):
    if sector == "Wind":
        oo = (offon or "").strip().lower()
        return "offshore wind" if oo == "offshore" else "onshore wind" if oo == "onshore" else "wind"
    return sector.lower()


def money_phrase_for_comment(inv_m):
    if inv_m is None:
        return ""
    return (f"${inv_m / 1000.0:,.2f} billion" if inv_m >= INVESTMENT_BILLION_THRESHOLD
            else f"${inv_m:,.2f} million")


def investment_column_value_millions(inv_m):
    return round(float(inv_m), 2) if inv_m is not None else ""


def build_comment(sector_txt, inv_estimated, inv_m, cap_mw,
                  fc_date, fc_estimated, op_date, op_estimated, is_operational):
    blocks = []
    if inv_estimated and inv_m is not None and cap_mw and cap_mw > 0:
        blocks.append(
            "Estimated Project Investment:\n"
            f"-- Enerdatics calculates the project's investment to be {money_phrase_for_comment(inv_m)}, "
            f"representing an investment metric of ${inv_m / cap_mw:.2f} million/MW.\n"
            f"-- This estimate is in line with reported figures for {sector_txt} "
            "projects of this scale and size in the region."
        )
    timeline = None
    if op_estimated and fc_estimated:
        timeline = (
            "Estimated Development Timeline:\n"
            f"-- The project is estimated to be commercially operational by {fmt_month_year(op_date)}.\n"
            "-- Further, Enerdatics believes that the developer will have secured financing to fund "
            f"construction activities on the project by {fmt_month_year(fc_date)}.\n"
            f"-- This estimate is based on reported development timelines for {sector_txt} "
            "projects of this scale and size in the region."
        )
    elif op_estimated and not fc_estimated and not is_operational:
        timeline = (
            "Estimated Development Timeline:\n"
            f"-- The project is estimated to be commercially operational by {fmt_month_year(op_date)}.\n"
            f"-- This estimate is based on reported development timelines for {sector_txt} "
            "projects of this scale and size in the region.\n"
            "-- Further, The developer will secure financing to fund construction activities "
            f"on the project by {fmt_month_year(fc_date)}."
        )
    elif is_operational and fc_estimated:
        timeline = (
            "Estimated Development Timeline:\n"
            f"-- The project reached commercial operations in {fmt_month_year(op_date)}.\n"
            "-- Further, Enerdatics believes that developer would have secured financing to fund "
            f"construction activities on the project by {fmt_month_year(fc_date)}.\n"
            f"-- This estimate is based on reported development timelines for {sector_txt} "
            "projects of this scale and size in the region."
        )
    if timeline:
        blocks.append(timeline)
    return "\n\n".join(blocks)


# =============================================================================
# Location helpers (OpenCage)
# =============================================================================
_BNG_TRANSFORMER = None
_OPENCAGE_CACHE = {}
_OPENCAGE_LAST_REQUEST = 0.0
_OPENCAGE_WARNING_SHOWN = False


def _coord_float(v):
    if _blank(v):
        return None
    try:
        return float(str(v).strip().replace(",", ""))
    except (TypeError, ValueError):
        return None


def _clean_geo_name(value):
    if _blank(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip(" ,.-")


def bng_to_latlon(easting, northing):
    global _BNG_TRANSFORMER
    x, y = _coord_float(easting), _coord_float(northing)
    if x is None or y is None:
        return None, None
    try:
        if _BNG_TRANSFORMER is None:
            from pyproj import Transformer
            _BNG_TRANSFORMER = Transformer.from_crs("EPSG:27700", "EPSG:4326", always_xy=True)
        lon, lat = _BNG_TRANSFORMER.transform(x, y)
        if not (-90 <= lat <= 90 and -180 <= lon <= 180):
            return None, None
        return round(lat, 6), round(lon, 6)
    except Exception:
        return None, None


def _extract_opencage_location(result):
    components = result.get("components") or {}
    cc = str(components.get("country_code") or "").strip().lower()
    cn = str(components.get("country") or "").strip().lower()
    if cc and cc not in {"gb", "uk"} and "united kingdom" not in cn:
        return "", ""
    sub_location = ""
    for key in ("city", "town", "village", "municipality", "hamlet", "locality",
                "suburb", "neighbourhood", "city_district", "borough"):
        sub_location = _clean_geo_name(components.get(key))
        if sub_location:
            break
    state_region = ""
    for key in ("state_district", "county", "region"):
        state_region = _clean_geo_name(components.get(key))
        if state_region:
            break
    if not state_region:
        state_region = _clean_geo_name(components.get("state"))
    if sub_location and state_region and sub_location.casefold() == state_region.casefold():
        sub_location = ""
    return sub_location, state_region


def reverse_geocode_opencage(latitude, longitude):
    global _OPENCAGE_LAST_REQUEST, _OPENCAGE_WARNING_SHOWN
    cache_key = (round(float(latitude), 6), round(float(longitude), 6))
    if cache_key in _OPENCAGE_CACHE:
        return _OPENCAGE_CACHE[cache_key]
    if not CFG.OPENCAGE_API_KEY:
        if not _OPENCAGE_WARNING_SHOWN:
            print("  WARNING: OPENCAGE_API_KEY is blank; location filling skipped.")
            _OPENCAGE_WARNING_SHOWN = True
        _OPENCAGE_CACHE[cache_key] = ("", "")
        return ("", "")
    elapsed = time.monotonic() - _OPENCAGE_LAST_REQUEST
    if elapsed < OPENCAGE_MIN_INTERVAL_SEC:
        time.sleep(OPENCAGE_MIN_INTERVAL_SEC - elapsed)
    try:
        response = requests.get(
            "https://api.opencagedata.com/geocode/v1/json",
            params={"q": f"{latitude:.6f},{longitude:.6f}", "key": CFG.OPENCAGE_API_KEY,
                    "language": "en", "no_annotations": 1, "no_record": 1},
            timeout=OPENCAGE_TIMEOUT_SEC)
        _OPENCAGE_LAST_REQUEST = time.monotonic()
        if not response.ok:
            if not _OPENCAGE_WARNING_SHOWN:
                print(f"  WARNING: OpenCage HTTP {response.status_code}.")
                _OPENCAGE_WARNING_SHOWN = True
            result = ("", "")
        else:
            results = response.json().get("results") or []
            result = _extract_opencage_location(results[0]) if results else ("", "")
    except Exception as exc:
        _OPENCAGE_LAST_REQUEST = time.monotonic()
        if not _OPENCAGE_WARNING_SHOWN:
            print(f"  WARNING: OpenCage failed: {exc}")
            _OPENCAGE_WARNING_SHOWN = True
        result = ("", "")
    _OPENCAGE_CACHE[cache_key] = result
    return result


def fill_blank_location_from_repd_xy(row, rrow, changes):
    need_sub, need_reg = _blank(row.get(C_SUBLOC)), _blank(row.get(C_REGION))
    if not (need_sub or need_reg):
        return
    lat, lon = bng_to_latlon(rrow.get(R_XCOORD), rrow.get(R_YCOORD))
    if lat is None:
        return
    sub_location, state_region = reverse_geocode_opencage(lat, lon)
    if need_sub and sub_location:
        _set_logged(row, C_SUBLOC, sub_location, changes, label="Sub_location (from REPD X/Y)")
    if need_reg and state_region:
        _set_logged(row, C_REGION, state_region, changes, label="State/Region (from REPD X/Y)")


# =============================================================================
# Lifecycle-driven flags + Show SPV
# =============================================================================
def _set_logged(row, col, val, changes, label=None):
    old = (row.get(col) or "").strip()
    row[col] = val
    if val != old and (val or old):
        changes.append(f"{label or col}: '{old}' -> '{val}'")


def apply_lifecycle_flags(row, work_bucket, changes, sector=""):
    eia, grid, fin = LIFECYCLE_FLAGS.get(work_bucket, ("", "", ""))
    psa = normalize_multiselect(row.get(C_PSA) or "").strip()
    cfd_round = normalize_multiselect(row.get(C_CFDAR) or "").strip()
    is_cfd = bool(cfd_round) or psa == "Contracts-for-Difference"
    ppa = "Yes - Complete" if (work_bucket == BUCKET_OPERATIONAL or is_cfd) else "No"

    _set_logged(row, C_EIA, eia, changes)
    _set_logged(row, C_GRID, grid, changes)
    _set_logged(row, C_FIN, fin, changes)
    _set_logged(row, C_PPA, ppa, changes)
    _set_logged(row, C_FINTYPE, "Debt; Equity" if fin == "Yes - Complete" else "", changes)

    if cfd_round:
        _set_logged(row, C_PSA, "Contracts-for-Difference", changes)
        return

    if sector == BATTERY_SECTOR:
        _set_logged(row, C_PSA, "Not applicable", changes)
        existing = (row.get(C_REVMECH) or "").strip()
        if existing == "" or existing.lower() in REVMECH_MANAGED_VALUES:
            if work_bucket in (BUCKET_CANCELLED, BUCKET_DECOMM):
                target = "Not applicable"
            elif work_bucket == BUCKET_OPERATIONAL:
                target = "Undisclosed"
            elif work_bucket in (BUCKET_EARLY, BUCKET_LATE, BUCKET_CONSTR):
                target = "Pending finalization"
            else:
                target = existing
            if target != existing:
                _set_logged(row, C_REVMECH, target, changes)
        return

    if (row.get(C_REVMECH) or "").strip():
        existing = (row.get(C_REVMECH) or "").strip()
        if existing.lower() in REVMECH_MANAGED_VALUES:
            _set_logged(row, C_REVMECH, "", changes, label="Revenue_mechanism cleared (non-Battery)")

    if sector in ("Solar", "Wind") and _blank(psa):
        if (row.get(C_PPA) or "").strip() == "Yes - Complete":
            _set_logged(row, C_PSA, "Yet to be determined/Unavailable", changes)


def apply_show_spv(row, rrow, changes):
    spv = (row.get(C_SPV) or "").strip()
    if not spv and rrow is not None:
        op = (rrow.get(R_OPERATOR) or "").strip()
        if op:
            _set_logged(row, C_SPV, op, changes, label="SPV (from REPD operator)")
            spv = op
    _set_logged(row, C_SHOWSPV, "Yes" if spv else "No", changes)


def clear_estimates_and_dates(row, changes, reason):
    wiped = []
    for stcol, estcol, repcol in [(C_INV_ST, C_INV_EST, C_INV_REP),
                                  (C_OP_ST, C_OP_EST, C_OP_REP),
                                  (C_FC_ST, C_FC_EST, C_FC_REP)]:
        if (not _blank(row.get(estcol)) or not _blank(row.get(repcol))
                or str(row.get(stcol) or "").strip() != ST_NONE):
            wiped.append(stcol)
        row[estcol] = ""
        row[repcol] = ""
        row[stcol] = ST_NONE
    if not _blank(row.get(C_COMMENT)) or str(row.get(C_SHOWCOM) or "").strip() != "No":
        wiped.append(C_COMMENT)
    row[C_COMMENT] = ""
    row[C_SHOWCOM] = "No"
    if wiped:
        changes.append(reason)
    return bool(wiped)


BLANKCAP_WIPE_REASON = ("suspended/cancelled + blank capacity -> statuses set to 'Not reported', "
                        "estimates and reported values cleared")


# =============================================================================
# Date ordering validation: EIA <= FC <= OP
# =============================================================================
def _get_date_and_status(row, date_col, est_col, rep_col, status_col):
    rep = parse_db_date(row.get(rep_col))
    est = parse_db_date(row.get(est_col))
    stat = (row.get(status_col) or "").strip()
    if stat == ST_REPORTED and rep:
        return rep, True
    if stat == ST_ESTIMATED and est:
        return est, False
    return None, False


def validate_and_fix_dates(row, cap, sector, work_bucket, rrow, changes):
    eia_date = parse_db_date(row.get(C_ENVDATE))

    def _read():
        fc, fc_rep = _get_date_and_status(row, C_FC_EST, C_FC_EST, C_FC_REP, C_FC_ST)
        op, op_rep = _get_date_and_status(row, C_OP_EST, C_OP_EST, C_OP_REP, C_OP_ST)
        return fc, fc_rep, op, op_rep

    def _clear_fc_and_reestimate(reason):
        row[C_FC_REP] = ""
        row[C_FC_EST] = ""
        row[C_FC_ST] = ST_NONE
        anch = parse_repd_date(rrow.get(R_SUBMIT)) or PRESENT_DATE
        if work_bucket == BUCKET_EARLY:
            new_fc = estimate_fc_forward(anch, cap, EARLY_MONTHS_PER_50, EARLY_FC_FLOOR_MONTHS)
        elif work_bucket in (BUCKET_LATE, BUCKET_CONSTR):
            new_fc = estimate_fc_forward(anch, cap, LATE_MONTHS_PER_50, LATE_FC_FLOOR_MONTHS)
        elif work_bucket == BUCKET_OPERATIONAL:
            op = parse_db_date(row.get(C_OP_REP)) or parse_db_date(row.get(C_OP_EST))
            new_fc = backdate_fc_from_op(op, cap) if op else None
        else:
            new_fc = None
        if new_fc:
            row[C_FC_EST] = fmt_db_date(new_fc)
            row[C_FC_ST] = ST_ESTIMATED
            changes.append(f"{reason}: FC re-estimated {fmt_db_date(new_fc)}")
        else:
            changes.append(f"{reason}: FC cleared (cannot re-estimate)")

    def _clear_op_and_reestimate(fc_date, reason):
        row[C_OP_REP] = ""
        row[C_OP_EST] = ""
        row[C_OP_ST] = ST_NONE
        if fc_date:
            if work_bucket == BUCKET_CONSTR:
                new_op = estimate_op_from_fc(fc_date, cap, sector, per_50=CONSTRUCTION_OP_PER_50)
            else:
                new_op = estimate_op_from_fc(fc_date, cap, sector)
            if new_op and new_op > PRESENT_DATE:
                row[C_OP_EST] = fmt_db_date(new_op)
                row[C_OP_ST] = ST_ESTIMATED
                changes.append(f"{reason}: OP re-estimated {fmt_db_date(new_op)}")
                return new_op
        changes.append(f"{reason}: OP cleared (cannot re-estimate)")
        return None

    # Rule A: FC must not be in the future for Operational / In construction
    if work_bucket in (BUCKET_OPERATIONAL, BUCKET_CONSTR):
        fc_date, fc_rep, op_date, op_rep = _read()
        if fc_date and fc_date > PRESENT_DATE:
            op_anchor = parse_db_date(row.get(C_OP_REP)) or parse_db_date(row.get(C_OP_EST))
            if op_anchor:
                new_fc = backdate_fc_from_op(op_anchor, cap)
                if new_fc >= PRESENT_DATE:
                    new_fc = eom_add(PRESENT_DATE, -1)
            else:
                new_fc = eom_add(PRESENT_DATE, -1)
            row[C_FC_REP] = ""
            row[C_FC_EST] = fmt_db_date(new_fc)
            row[C_FC_ST] = ST_ESTIMATED
            changes.append(f"DATE CHECK Rule A: FC ({fmt_db_date(fc_date)}) in the future for "
                           f"{BUCKET_TO_DB_LABEL.get(work_bucket, '')}, backdated to {fmt_db_date(new_fc)}")

    # Rule B: OP must not be in the past for Early/Late/In construction
    if work_bucket in (BUCKET_EARLY, BUCKET_LATE, BUCKET_CONSTR):
        fc_date, fc_rep, op_date, op_rep = _read()
        if op_date and op_date <= PRESENT_DATE:
            new_op = _clear_op_and_reestimate(
                fc_date, f"DATE CHECK Rule B: OP ({fmt_db_date(op_date)}) in the past for "
                         f"{BUCKET_TO_DB_LABEL.get(work_bucket, '')}")
            if new_op is None or new_op <= PRESENT_DATE:
                row[C_OP_EST] = ""
                row[C_OP_REP] = ""
                row[C_OP_ST] = ST_NONE

    # Rule C: FC must not be in the past for Early/Late
    if work_bucket in (BUCKET_EARLY, BUCKET_LATE):
        fc_date, fc_rep, op_date, op_rep = _read()
        if fc_date and fc_date <= PRESENT_DATE:
            _clear_fc_and_reestimate(f"DATE CHECK Rule C: FC ({fmt_db_date(fc_date)}) in the past for "
                                     f"{BUCKET_TO_DB_LABEL.get(work_bucket, '')}")
            fc_date, fc_rep, op_date, op_rep = _read()
            if op_date and op_date <= PRESENT_DATE and fc_date:
                _clear_op_and_reestimate(fc_date, "DATE CHECK Rule C follow-up: OP still in the past")

    # Ordering checks EIA <= FC <= OP (reported OP of an operational project is locked)
    fc_date, fc_rep, op_date, op_rep = _read()
    op_locked = (work_bucket == BUCKET_OPERATIONAL and op_rep and op_date is not None)

    if eia_date and fc_date and eia_date > fc_date:
        new_fc = eom_add(eia_date, 1)
        if op_locked and new_fc >= op_date:
            new_fc = eom_add(op_date, -max(3, int(round(months_for_capacity(cap, OP_MONTHS_PER_50)))))
        row[C_FC_REP] = ""
        row[C_FC_EST] = fmt_db_date(new_fc)
        row[C_FC_ST] = ST_ESTIMATED
        changes.append(f"DATE CHECK: EIA ({fmt_db_date(eia_date)}) > FC ({fmt_db_date(fc_date)}), "
                       f"FC moved to {fmt_db_date(new_fc)}")
        fc_date = new_fc

    fc_date, fc_rep, op_date, op_rep = _read()
    if fc_date and op_date and fc_date > op_date:
        if op_locked:
            new_fc = eom_add(op_date, -max(3, int(round(months_for_capacity(cap, OP_MONTHS_PER_50)))))
            row[C_FC_REP] = ""
            row[C_FC_EST] = fmt_db_date(new_fc)
            row[C_FC_ST] = ST_ESTIMATED
            changes.append(f"DATE CHECK: FC ({fmt_db_date(fc_date)}) > reported OP "
                           f"({fmt_db_date(op_date)}), FC backdated to {fmt_db_date(new_fc)}")
        else:
            if work_bucket == BUCKET_CONSTR:
                new_op = estimate_op_from_fc(fc_date, cap, sector, per_50=CONSTRUCTION_OP_PER_50)
            else:
                new_op = estimate_op_from_fc(fc_date, cap, sector)
            if new_op is None or new_op <= fc_date:
                new_op = eom_add(fc_date, 6)
            row[C_OP_REP] = ""
            row[C_OP_EST] = fmt_db_date(new_op)
            row[C_OP_ST] = ST_ESTIMATED
            changes.append(f"DATE CHECK: FC ({fmt_db_date(fc_date)}) > OP ({fmt_db_date(op_date)}), "
                           f"OP pushed to {fmt_db_date(new_op)}")

    # Minimum FC -> OP gap
    fc_date, fc_rep, op_date, op_rep = _read()
    if fc_date and op_date:
        min_gap = min(max(3, int(round(months_for_capacity(cap, OP_MONTHS_PER_50)))), OP_MAX_MONTHS)
        if (op_date.year * 12 + op_date.month) - (fc_date.year * 12 + fc_date.month) < min_gap:
            if op_locked:
                new_fc = eom_add(op_date, -min_gap)
                row[C_FC_REP] = ""
                row[C_FC_EST] = fmt_db_date(new_fc)
                row[C_FC_ST] = ST_ESTIMATED
                changes.append(f"DATE CHECK: FC->OP gap < {min_gap}mo, FC backdated to {fmt_db_date(new_fc)}")
            else:
                required_op = eom_add(fc_date, min_gap)
                row[C_OP_REP] = ""
                row[C_OP_EST] = fmt_db_date(required_op)
                row[C_OP_ST] = ST_ESTIMATED
                changes.append(f"DATE CHECK: FC->OP gap < {min_gap}mo, OP pushed to {fmt_db_date(required_op)}")

    # Blank-date check
    for st_col, est_col, rep_col in ((C_FC_ST, C_FC_EST, C_FC_REP), (C_OP_ST, C_OP_EST, C_OP_REP)):
        st = (row.get(st_col) or "").strip()
        if st not in (ST_NONE, ""):
            if _blank(row.get(est_col)) and _blank(row.get(rep_col)):
                row[st_col] = ST_NONE
                changes.append(f"{st_col}: blank date, set to 'Not reported'")


def update_env_approval_date(row, rrow, changes):
    if (row.get(C_EIA) or "").strip().lower() == "no":
        return
    if not _blank(row.get(C_ENVDATE)):
        return
    status = (rrow.get(R_STATUS) or "").strip().lower()
    date_raw = rrow.get(R_APPEAL_GRANTED) if status == "appeal granted" else rrow.get(R_PPGRANT)
    ppg = parse_repd_date(date_raw)
    if ppg is None:
        return
    _set_logged(row, C_ENVDATE, fmt_db_date(ppg), changes, label="Environmental_approval_date (from REPD)")


def rebuild_comment_from_row(row, sector, offon, cap):
    if sector == HYDRO_SECTOR:
        row[C_COMMENT] = ""
        row[C_SHOWCOM] = "No"
        return
    fc_d = parse_db_date(row.get(C_FC_EST)) or parse_db_date(row.get(C_FC_REP))
    op_d = parse_db_date(row.get(C_OP_EST)) or parse_db_date(row.get(C_OP_REP))
    fc_e = (row.get(C_FC_ST) or "").strip() == ST_ESTIMATED
    op_e = (row.get(C_OP_ST) or "").strip() == ST_ESTIMATED
    inv_e = (row.get(C_INV_ST) or "").strip() == ST_ESTIMATED
    inv_m = parse_money(row.get(C_INV_EST)) if inv_e else parse_money(row.get(C_INV_REP))
    is_op = ((row.get(C_LIFE) or "").strip().lower() == "operational"
             or (op_d is not None and op_d <= PRESENT_DATE))
    txt = build_comment(sector_phrase(sector, offon), inv_e, inv_m, cap, fc_d, fc_e, op_d, op_e, is_op)
    row[C_COMMENT], row[C_SHOWCOM] = (txt, "Yes") if txt.strip() else ("", "No")


def _clear_psa_if_ppa_no(row, sector, changes):
    if sector != BATTERY_SECTOR:
        cfd_r = normalize_multiselect(row.get(C_CFDAR) or "").strip()
        if not cfd_r and (row.get(C_PPA) or "").strip() == "No" and (row.get(C_PSA) or "").strip():
            _set_logged(row, C_PSA, "", changes, label="PSA cleared (PPA=No)")


# =============================================================================
# Core per-row processor (unchanged rules)
# =============================================================================
def match_repd_id(join_raw, repd_idx):
    for tok in re.split(r"[;,]|\bOR\b|/", join_raw):
        tok = tok.strip()
        if tok.isdigit() and tok in repd_idx:
            return tok
    return join_raw if join_raw in repd_idx else None


def process_row(row, repd_idx, log):
    changes, tracked_cols = [], set()
    name = row.get(C_NAME, "")
    join_raw = "" if _blank(row.get(C_JOIN)) else str(row.get(C_JOIN)).strip()
    rid = match_repd_id(join_raw, repd_idx)
    sector = (row.get(C_SECTOR) or "").strip()

    if rid is None:
        cur_label = str(row.get(C_LIFE) or "").strip().lower()
        wb = DB_LABEL_TO_BUCKET.get(cur_label)
        apply_lifecycle_flags(row, wb, changes, sector=sector)
        apply_show_spv(row, None, changes)
        cap_nm = parse_capacity(row.get(C_CAP))
        if cur_label in BLANKCAP_WIPE_LABELS and (cap_nm is None or cap_nm <= 0):
            clear_estimates_and_dates(row, changes, BLANKCAP_WIPE_REASON)
        _clear_psa_if_ppa_no(row, sector, changes)
        log.append({"Project name": name, "State data plant ID": join_raw,
                    "Action": "SKIPPED: no REPD match" + (" (flags set)" if changes else ""),
                    "Detail": " | ".join(changes)})
        return row, bool(changes), {"__NO_MATCH__"}

    rrow = repd_idx[rid]
    fill_blank_location_from_repd_xy(row, rrow, changes)
    cap = parse_capacity(row.get(C_CAP))
    offon = row.get(C_OFFON)

    if cap is None or cap <= 0:
        repd_status = (rrow.get(R_STATUS) or "").strip().lower()
        bucket = REPD_STATUS_TO_BUCKET.get(repd_status)
        cur_label = str(row.get(C_LIFE) or "").strip().lower()
        work_bucket = DB_LABEL_TO_BUCKET.get(cur_label, bucket)
        apply_lifecycle_flags(row, work_bucket, changes, sector=sector)
        apply_show_spv(row, rrow, changes)
        if cur_label in BLANKCAP_WIPE_LABELS:
            clear_estimates_and_dates(row, changes, BLANKCAP_WIPE_REASON)
        if work_bucket not in (BUCKET_OPERATIONAL, BUCKET_CANCELLED, BUCKET_DECOMM):
            has_fc_now = not _blank(row.get(C_FC_REP)) or not _blank(row.get(C_FC_EST))
            if has_fc_now and _blank(row.get(C_OP_ST)):
                row[C_OP_ST] = ST_NONE
                changes.append("OP status set to Not reported (FC present but no capacity to estimate OP)")
        _clear_psa_if_ppa_no(row, sector, changes)
        rebuild_comment_from_row(row, sector, offon, cap)
        log.append({"Project name": name, "State data plant ID": rid,
                    "Action": "SKIPPED: blank/zero capacity" + (" (flags set)" if changes else ""),
                    "Detail": " | ".join(changes)})
        return row, bool(changes), set()

    # 0. site area
    sqm = parse_capacity(rrow.get(R_SITEAREA))
    if sqm is not None and sqm > 0:
        new_area = f"{sqm / SQM_PER_ACRE:,.{SITE_AREA_DECIMALS}f}"
        old_area = (row.get(C_SITEAREA) or "").strip()
        if (OVERWRITE_SITE_AREA or old_area == "") and new_area != old_area:
            row[C_SITEAREA] = new_area
            changes.append(f"site area: {new_area} acres ({sqm:,.0f} sqm)")
            tracked_cols.add(C_SITEAREA)

    # 1. lifecycle
    repd_status = (rrow.get(R_STATUS) or "").strip().lower()
    bucket = REPD_STATUS_TO_BUCKET.get(repd_status)
    if bucket is not None:
        new_label, old_label = BUCKET_TO_DB_LABEL[bucket], (row.get(C_LIFE) or "").strip()
        if new_label != old_label:
            row[C_LIFE] = new_label
            changes.append(f"life cycle: '{old_label}' -> '{new_label}'")
            tracked_cols.add(C_LIFE)
    cur_label = (row.get(C_LIFE) or "").strip().lower()
    work_bucket = DB_LABEL_TO_BUCKET.get(cur_label, bucket)

    # 1b. lifecycle flags + Show SPV
    psa_before = (row.get(C_PSA) or "").strip()
    apply_lifecycle_flags(row, work_bucket, changes, sector=sector)
    if (row.get(C_PSA) or "").strip() != psa_before:
        tracked_cols.add(C_PSA)
    apply_show_spv(row, rrow, changes)

    # 2. terminal: wipe
    if work_bucket in TERMINAL_BUCKETS:
        clear_estimates_and_dates(row, changes,
                                  "terminal lifecycle -> cleared estimates, status set to 'Not reported'")
        if sector != BATTERY_SECTOR:
            cfd_r = normalize_multiselect(row.get(C_CFDAR) or "").strip()
            if not cfd_r and (row.get(C_PSA) or "").strip():
                _set_logged(row, C_PSA, "", changes, label="PSA cleared (terminal)")
        if changes:
            log.append({"Project name": name, "State data plant ID": rid,
                        "Action": "UPDATED (terminal)", "Detail": " | ".join(changes)})
        return row, bool(changes), tracked_cols

    # 3. reported dates from REPD
    repd_op = parse_repd_date(rrow.get(R_OPER))
    repd_fc = parse_repd_date(rrow.get(R_UCONST))
    op_status = (row.get(C_OP_ST) or "").strip()
    fc_status = (row.get(C_FC_ST) or "").strip()
    inv_status = (row.get(C_INV_ST) or "").strip()

    if repd_op is not None:
        new_op = fmt_db_date(repd_op)
        if new_op != (row.get(C_OP_REP) or "").strip() or op_status != ST_REPORTED:
            changes.append(f"OP reported: {new_op}")
        row[C_OP_REP], row[C_OP_EST], row[C_OP_ST] = new_op, "", ST_REPORTED
        op_status = ST_REPORTED
    if repd_fc is not None:
        new_fc = fmt_db_date(repd_fc)
        if new_fc != (row.get(C_FC_REP) or "").strip() or fc_status != ST_REPORTED:
            changes.append(f"FC reported: {new_fc}")
        row[C_FC_REP], row[C_FC_EST], row[C_FC_ST] = new_fc, "", ST_REPORTED
        fc_status = ST_REPORTED

    fc_reported = parse_db_date(row.get(C_FC_REP))
    op_reported = parse_db_date(row.get(C_OP_REP))
    is_operational = (work_bucket == BUCKET_OPERATIONAL
                      or (op_reported is not None and op_reported <= PRESENT_DATE))

    # 3b. hydro: no estimates
    if sector == HYDRO_SECTOR:
        if inv_status != ST_REPORTED:
            row[C_INV_ST], row[C_INV_EST] = ST_NONE, ""
            changes.append("hydro -> investment not estimated")
        if op_status != ST_REPORTED:
            row[C_OP_ST], row[C_OP_EST] = ST_NONE, ""
        if fc_status != ST_REPORTED:
            row[C_FC_ST], row[C_FC_EST] = ST_NONE, ""
        row[C_COMMENT], row[C_SHOWCOM] = "", "No"
        log.append({"Project name": name, "State data plant ID": rid,
                    "Action": "UPDATED" if changes else "no change", "Detail": " | ".join(changes)})
        return row, bool(changes), tracked_cols

    # 3c. environmental approval date
    env_before = (row.get(C_ENVDATE) or "").strip()
    update_env_approval_date(row, rrow, changes)
    if (row.get(C_ENVDATE) or "").strip() != env_before:
        tracked_cols.add(C_ENVDATE)

    # 3d. operational without reported OP date -> no estimates
    if work_bucket == BUCKET_OPERATIONAL and op_reported is None:
        row[C_OP_ST] = row[C_FC_ST] = row[C_INV_ST] = ST_NONE
        row[C_OP_EST] = row[C_OP_REP] = row[C_FC_EST] = row[C_FC_REP] = ""
        row[C_INV_EST] = row[C_INV_REP] = ""
        row[C_COMMENT], row[C_SHOWCOM] = "", "No"
        changes.append("operational w/o reported OP date -> no estimates (dates + investment)")
        log.append({"Project name": name, "State data plant ID": rid,
                    "Action": "UPDATED" if changes else "no change", "Detail": " | ".join(changes)})
        return row, bool(changes), tracked_cols

    # 4. FC estimate
    fc_date, fc_estimated = fc_reported, (fc_status != ST_REPORTED)
    if fc_status != ST_REPORTED:
        if work_bucket == BUCKET_OPERATIONAL and op_reported:
            fc_date = backdate_fc_from_op(op_reported, cap)
        elif work_bucket == BUCKET_CONSTR:
            fc_date = estimate_fc_forward(PRESENT_DATE, cap, LATE_MONTHS_PER_50, LATE_FC_FLOOR_MONTHS)
        elif work_bucket == BUCKET_EARLY:
            anchor = parse_repd_date(rrow.get(R_SUBMIT)) or PRESENT_DATE
            fc_date = estimate_fc_forward(anchor, cap, EARLY_MONTHS_PER_50, EARLY_FC_FLOOR_MONTHS)
        elif work_bucket == BUCKET_LATE:
            anchor = parse_repd_date(rrow.get(R_SUBMIT)) or PRESENT_DATE
            fc_date = estimate_fc_forward(anchor, cap, LATE_MONTHS_PER_50, LATE_FC_FLOOR_MONTHS)
        if fc_date:
            row[C_FC_EST], row[C_FC_REP], row[C_FC_ST] = fmt_db_date(fc_date), "", ST_ESTIMATED
            fc_estimated = True
            changes.append(f"FC estimated: {fmt_db_date(fc_date)}")

    # 5. OP estimate
    op_date, op_estimated = op_reported, (op_status != ST_REPORTED)
    if op_status != ST_REPORTED and work_bucket != BUCKET_OPERATIONAL:
        if work_bucket == BUCKET_CONSTR:
            base_fc = fc_date or estimate_fc_forward(PRESENT_DATE, cap, LATE_MONTHS_PER_50, LATE_FC_FLOOR_MONTHS)
            cand = estimate_op_from_fc(base_fc, cap, sector, per_50=CONSTRUCTION_OP_PER_50)
            min_fut = eom_add(PRESENT_DATE, CONSTR_MIN_FUTURE_OP_MO)
            if cand < min_fut:
                back = min(12, max(3, round(months_for_capacity(cap, 3))))
                base_fc = eom_add(PRESENT_DATE, -back)
                cand = estimate_op_from_fc(base_fc, cap, sector, per_50=CONSTRUCTION_OP_PER_50)
                if cand < min_fut:
                    cand = min_fut
                if fc_status != ST_REPORTED:
                    row[C_FC_EST], row[C_FC_ST] = fmt_db_date(base_fc), ST_ESTIMATED
                    fc_date, fc_estimated = base_fc, True
            op_date = cand
        elif fc_date and work_bucket in (BUCKET_EARLY, BUCKET_LATE, BUCKET_CONSTR):
            op_date = estimate_op_from_fc(fc_date, cap, sector)
        if op_date and work_bucket != BUCKET_OPERATIONAL:
            row[C_OP_EST], row[C_OP_REP], row[C_OP_ST] = fmt_db_date(op_date), "", ST_ESTIMATED
            op_estimated = True
            changes.append(f"OP estimated from FC: {fmt_db_date(op_date)}")

    # 5b. PSA rules based on PPA signed (except battery)
    if sector != BATTERY_SECTOR:
        current_ppa = (row.get(C_PPA) or "").strip()
        current_psa = (row.get(C_PSA) or "").strip()
        cfd_round_check = normalize_multiselect(row.get(C_CFDAR) or "").strip()
        if current_ppa == "No" and current_psa and not cfd_round_check:
            _set_logged(row, C_PSA, "", changes, label="Power_sales_arrangement cleared (PPA signed = No)")
            tracked_cols.add(C_PSA)
        elif current_ppa == "Yes - Complete" and not current_psa and not cfd_round_check:
            _set_logged(row, C_PSA, "Yet to be determined/Unavailable", changes,
                        label="Power_sales_arrangement set (PPA=Yes-Complete, PSA was blank)")
            tracked_cols.add(C_PSA)

    # 6. date ordering checks
    validate_and_fix_dates(row, cap, sector, work_bucket, rrow, changes)

    # 6b. FC present but OP missing -> estimate OP
    if work_bucket not in (BUCKET_OPERATIONAL, BUCKET_CANCELLED, BUCKET_DECOMM):
        has_op_date = (not _blank(row.get(C_OP_EST))) or (not _blank(row.get(C_OP_REP)))
        fc_now = parse_db_date(row.get(C_FC_REP)) or parse_db_date(row.get(C_FC_EST))
        if not has_op_date and fc_now:
            anchor_fc = fc_now if fc_now > PRESENT_DATE else PRESENT_DATE
            if work_bucket == BUCKET_CONSTR:
                new_op = estimate_op_from_fc(anchor_fc, cap, sector, per_50=CONSTRUCTION_OP_PER_50)
            else:
                new_op = estimate_op_from_fc(anchor_fc, cap, sector)
            min_fut = eom_add(PRESENT_DATE, CONSTR_MIN_FUTURE_OP_MO if work_bucket == BUCKET_CONSTR else 1)
            if new_op and new_op < min_fut:
                new_op = min_fut
            if new_op:
                row[C_OP_EST], row[C_OP_REP], row[C_OP_ST] = fmt_db_date(new_op), "", ST_ESTIMATED
                changes.append(f"OP estimated (post-validation) from FC: {fmt_db_date(new_op)}")

    # 7. investment
    inv_estimated, inv_m = False, None
    if inv_status != ST_REPORTED and work_bucket not in TERMINAL_BUCKETS:
        inv_m, _metric = compute_investment(sector, work_bucket, cap, offon)
        if inv_m is not None:
            row[C_INV_EST] = investment_column_value_millions(inv_m)
            row[C_INV_REP], row[C_INV_ST] = "", ST_ESTIMATED
            inv_estimated = True
            changes.append(f"investment: {inv_m:,.2f}M (@${inv_m / cap:.2f} M/MW)")
    elif inv_status == ST_REPORTED:
        inv_m = parse_money(row.get(C_INV_REP))

    # 8. analyst comment
    fc_final = parse_db_date(row.get(C_FC_EST)) or parse_db_date(row.get(C_FC_REP))
    op_final = parse_db_date(row.get(C_OP_EST)) or parse_db_date(row.get(C_OP_REP))
    comment = build_comment(sector_phrase(sector, offon), inv_estimated, inv_m, cap,
                            fc_final, (row.get(C_FC_ST) or "").strip() == ST_ESTIMATED,
                            op_final, (row.get(C_OP_ST) or "").strip() == ST_ESTIMATED, is_operational)
    row[C_COMMENT], row[C_SHOWCOM] = (comment, "Yes") if comment.strip() else ("", "No")

    log.append({"Project name": name, "State data plant ID": rid,
                "Action": "UPDATED" if changes else "no change", "Detail": " | ".join(changes)})
    return row, bool(changes), tracked_cols


# =============================================================================
# Driver
# =============================================================================
def prepare_db(db: pd.DataFrame) -> pd.DataFrame:
    """Rename to canonical names, add missing columns, normalise dates (as in the Colab script)."""
    global C_REVMECH
    db = db.copy()
    db.rename(columns={k: v for k, v in OLD_TO_NEW.items() if k in db.columns}, inplace=True)
    existing_rev = find_revmech_column(db.columns)
    if existing_rev and existing_rev != C_REVMECH:
        db.rename(columns={existing_rev: C_REVMECH}, inplace=True)
    for c in (C_COMMENT, C_SHOWCOM, C_EIA, C_GRID, C_FIN, C_FINTYPE, C_PPA, C_PSA, C_REVMECH,
              C_SPV, C_SHOWSPV, C_REGION, C_SUBLOC, C_ENVDATE, C_COLOC, C_JOIN):
        if c not in db.columns:
            db[c] = ""
    if C_COLOC in db.columns:
        db.loc[db[C_COLOC].str.strip() == "", C_COLOC] = "No"
    for col in DATE_COLS:
        if col in db.columns:
            db[col] = db[col].map(normalise_date_value)
    return db


LIFECYCLE_RANK = {"cancelled": 0, "suspended": 0, "early-stage": 1, "late-stage": 2, "ready-to-build": 3,
                  "in construction": 4, "construction complete": 5, "operational": 6, "decommissioned": 7}
_TL_RE = re.compile(r"(?im)^[^\w\r\n]*e?stimated[ \t]+development[ \t]+timeline[ \t]*:.*?"
                    r"(?=(\r?\n[ \t]*\r?\n)|\Z)", re.S)


def timeline_paragraph(comment: str) -> str:
    m = _TL_RE.search(comment or "")
    return m.group(0).strip() if m else ""


def replace_timeline_in_place(existing: str, tl_sec: str) -> str:
    """QC rows: swap only the timeline paragraph; every other character stays as it was."""
    text = existing or ""
    nl = "\r\n" if "\r\n" in text else "\n"
    new = (tl_sec or "").replace("\r\n", "\n").replace("\n", nl)
    m = _TL_RE.search(text)
    if m:
        out = text[:m.start()] + new + text[m.end():]
        return re.sub(r"(\r?\n[ \t]*){3,}", nl + nl, out).strip() if not new else out
    if not new:
        return text
    return (text.rstrip() + nl + nl + new) if text.strip() else new


def qc_process(row: dict, repd_idx: dict, log: list) -> dict:
    """Lead QC rows: REPD capacity applied; lifecycle only moves UP; investment untouched;
    only the 'Estimated Development Timeline' paragraph of the analyst comment is replaced."""
    original = dict(row)
    rid = match_repd_id(_cmp(row.get(C_JOIN)), repd_idx)
    if rid:
        repd_cap = parse_capacity(repd_idx[rid].get(R_CAP))
        db_cap = parse_capacity(row.get(C_CAP))
        if repd_cap is not None and (db_cap is None or abs(repd_cap - db_cap) > 0.01):
            row[C_CAP] = f"{repd_cap:g}"
        bucket = REPD_STATUS_TO_BUCKET.get((repd_idx[rid].get(R_STATUS) or "").strip().lower())
        new_label = BUCKET_TO_DB_LABEL.get(bucket, "")
        cur = (original.get(C_LIFE) or "").strip()
        if new_label and new_label.lower() != cur.lower() and \
                LIFECYCLE_RANK.get(new_label.lower(), -1) <= LIFECYCLE_RANK.get(cur.lower(), -1):
            log.append({"Project name": row.get(C_NAME, ""), "State data plant ID": rid,
                        "Action": "QC: lifecycle downgrade not applied",
                        "Detail": f"REPD points to {cur} -> {new_label}; only capacity synced"})
            return row
    new_row, _c, tracked = process_row(dict(row), repd_idx, log)
    if "__NO_MATCH__" in tracked:
        return original
    for f in (C_INV_ST, C_INV_EST, C_INV_REP):
        if f in original:
            new_row[f] = original[f]
    comment = replace_timeline_in_place(original.get(C_COMMENT, ""), timeline_paragraph(new_row.get(C_COMMENT, "")))
    new_row[C_COMMENT], new_row[C_SHOWCOM] = comment, ("Yes" if comment.strip() else "No")
    return new_row


def is_lead_qc(v) -> bool:
    return str(v).strip().lower() in ("true", "yes", "1", "y")


def _cmp(v):
    return "" if _blank(v) else str(v).strip()


def run_update(db: pd.DataFrame, repd: pd.DataFrame, qc_extra: pd.DataFrame | None = None,
               forced_rows: list[dict] | None = None) -> dict:
    """
    db          : updater report (already prepare_db'd)
    qc_extra    : Lead QC rows from the other report that are not in `db` (prepare_db'd)
    forced_rows : rows matched by GPT (State_data_plant_ID just filled), processed like DB rows
    Returns dict with lists of row dicts (each carries "_changed") and the log DataFrame.
    Lead QC rows never go to the non-QC sheets; they go to "qc" with REPD capacity applied.
    """
    repd_idx = {str(r[R_REFID]).strip(): r for _, r in repd.iterrows()}
    columns = list(db.columns)
    qc_col = find_column(columns, "Lead QC")
    out = {"lifecycle_capacity": [], "site_area": [], "other": [], "full": [], "qc": [], "columns": columns}
    log = []

    def handle(row: dict, source: str):
        original = dict(row)
        qc = bool(qc_col) and is_lead_qc(row.get(qc_col, ""))
        if qc:
            if match_repd_id(_cmp(row.get(C_JOIN)), repd_idx) is None:
                return
            new_row = qc_process(dict(row), repd_idx, log)
        else:
            new_row, _changed, tracked = process_row(dict(row), repd_idx, log)
            if "__NO_MATCH__" in tracked and source != "gpt":
                return
        for col in (C_PSA, C_CFDAR, C_REVMECH):
            if col in new_row:
                new_row[col] = normalize_multiselect(new_row[col])
        if C_COMMENT in new_row and _blank(new_row.get(C_COMMENT)):
            new_row[C_SHOWCOM] = "No"
        changed_cols = {c for c in set(original) | set(new_row)
                        if not str(c).startswith("_") and _cmp(original.get(c)) != _cmp(new_row.get(c))}
        if source == "gpt":
            changed_cols.add(C_JOIN)          # Ref ID added by the GPT / manual match
        new_row["_changed"] = changed_cols
        new_row["_source"] = source
        out["full"].append(new_row)
        if not changed_cols:
            return
        if qc:
            out["qc"].append(new_row)
            return
        if changed_cols & LIFECYCLE_CAPACITY_CHANGE_COLS:
            out["lifecycle_capacity"].append(new_row)
        if changed_cols & SITE_AREA_CHANGE_COLS:
            out["site_area"].append(new_row)
        if changed_cols - LIFECYCLE_CAPACITY_CHANGE_COLS - SITE_AREA_CHANGE_COLS:
            out["other"].append(new_row)

    for row in db.to_dict("records"):
        handle(row, "db")
    if qc_extra is not None and len(qc_extra):
        for row in qc_extra.to_dict("records"):
            handle({c: row.get(c, "") for c in columns} | row, "qc_report")
    for row in forced_rows or []:
        handle(row, "gpt")
    out["log"] = pd.DataFrame(log)
    return out
