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
import unicodedata

import pandas as pd
import requests

from . import config as CFG

PRESENT_DATE = dt.date.today()          # main.py can override for a given run

OPENCAGE_MIN_INTERVAL_SEC = 1.05
OPENCAGE_TIMEOUT_SEC = 30

# ---- Date-estimate rates (house rules section 6) ----
# Early-stage FC = present + MW/50 x 8 months (>= 6 months); Late-stage FC = present + MW/50 x 6
# (>= 3 months); COD = FC + MW/50 x 6 (wind >= 6, others >= 3, <= 48 months); operational FC =
# COD - MW/50 x 6 (<= 36 months back); in-construction FC = present - MW/50 x 6 (not after present).
EARLY_MONTHS_PER_50      = 8
LATE_MONTHS_PER_50       = 6
OP_MONTHS_PER_50         = 6
CONSTRUCTION_OP_PER_50   = 6
EARLY_FC_FLOOR_MONTHS    = 6
LATE_FC_FLOOR_MONTHS     = 3
FC_EST_MAX_FROM_TODAY_MO = 24       # house cap for very large projects ("adjustable by size")
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
# Allowed $M/MW per house rules section 7 (an existing estimate outside these is re-estimated).
ALLOWED_METRIC = {
    ("Solar", "dev"): (1.0, 1.2), ("Solar", "op"): (1.1, 1.4),
    ("Wind", "dev"): (1.4, 1.6), ("Wind", "op"): (1.6, 1.8),
    ("Battery", "dev"): (0.9, 1.6), ("Battery", "op"): (0.9, 1.6),
}
METRIC_TOLERANCE = 0.02
INVESTMENT_BILLION_THRESHOLD = 1000.0
MIN_ESTIMATED_INVESTMENT_M   = 1.0

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

REVMECH_MANAGED_VALUES = {"not applicable", "pending finalization", "undisclosed",
                          "yet to be determined/unavailable"}
REVMECH_UNKNOWN = "Yet to be determined/Unavailable"     # SOP section 7: unknown battery mechanism
PSA_UNKNOWN     = "Yet to be determined/Unavailable"
PSA_BATTERY     = "Not Applicable"                       # house rules section 20 (capital A)
PSA_CFD         = "Contracts-for-Difference"

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
R_STATUS_SHORT = "Development Status (short)"


def repd_awaiting_construction(rrow):
    """REPD short status 'Awaiting Construction': permission granted, construction not started.
    Its Under Construction / Operational dates are only plans."""
    return (str(rrow.get(R_STATUS_SHORT) or "").strip().lower() == "awaiting construction")

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
            items = s.split(";")
    cleaned = []
    for it in items:
        t = str(it).strip().strip("'\"").strip()
        if t and t not in cleaned:
            cleaned.append(t)
    return ";".join(cleaned)


def multiselect_items(v):
    s = normalize_multiselect(v)
    return [t for t in s.split(";") if t] if s else []


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
    """Early/Late-stage FC: present + MW/50 x rate, never below the floor (anchor kept for callers)."""
    cand = eom_add(PRESENT_DATE, months_for_capacity(cap_mw, per_50))
    floor = eom_add(PRESENT_DATE, floor_months)
    cap_hi = eom_add(PRESENT_DATE, FC_EST_MAX_FROM_TODAY_MO)
    return min(max(cand, floor), cap_hi)


def last_month_end_on_or_before(d):
    e = end_of_month(d)
    return e if e <= d else end_of_month(add_months(dt.date(d.year, d.month, 1), -1))


def estimate_fc_construction(cap_mw):
    """In construction, FC missing: present - MW/50 x 6 months, month-end, not later than present."""
    cand = eom_add(PRESENT_DATE, -months_for_capacity(cap_mw, OP_MONTHS_PER_50))
    return min(cand, last_month_end_on_or_before(PRESENT_DATE))


def min_build_months(sector):
    return OP_MIN_MONTHS_WIND if sector == "Wind" else OP_MIN_MONTHS_OTHER


def estimate_op_from_fc(fc, cap_mw, sector, per_50=OP_MONTHS_PER_50):
    cand = eom_add(fc, months_for_capacity(cap_mw, per_50))
    return min(max(cand, eom_add(fc, min_build_months(sector))), eom_add(fc, OP_MAX_MONTHS))


def backdate_fc_from_op(op, cap_mw, sector=""):
    """Operational, FC missing: COD - MW/50 x 6 months, at least the minimum build time
    (wind 6, others 3) and at most 36 months back; month-end."""
    months = months_for_capacity(cap_mw, OP_MONTHS_PER_50)
    months = min(max(months, min_build_months(sector)), FC_BACKDATE_MAX_MONTHS)
    return eom_add(op, -months)


def months_between(a, b):
    return (b.year * 12 + b.month) - (a.year * 12 + a.month)


def metric_for(sector, stage, cap_mw):
    key = (sector, "any") if sector == "Battery" else (sector, stage)
    rng = METRIC_RANGES.get(key)
    if rng is None:
        return None
    lo, hi = rng
    return round((lo + hi) / 2.0, 2)


def metric_allowed(sector, offon, bucket, inv_m, cap_mw):
    """True when inv/MW is inside the house-rule range (offshore wind has no house range: always OK)."""
    if inv_m is None or not cap_mw:
        return False
    if sector == "Wind" and (offon or "").strip().lower() == "offshore":
        return True
    stage = "op" if bucket == BUCKET_OPERATIONAL else "dev"
    rng = ALLOWED_METRIC.get((sector, stage))
    if rng is None:
        return True
    m = inv_m / cap_mw
    return rng[0] - METRIC_TOLERANCE <= m <= rng[1] + METRIC_TOLERANCE


def round_investment(sector, inv):
    """House rules section 7: nearest $1M; BESS keeps one decimal."""
    return round(inv, 1) if sector == BATTERY_SECTOR else float(round(inv))


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
    inv = round_investment(sector, metric * cap_mw)
    return max(inv, MIN_ESTIMATED_INVESTMENT_M), metric


# =============================================================================
# Analyst comment builder (house rules section 8 templates)
# =============================================================================
def sector_phrase(sector, offon):
    if sector == "Wind":
        return "offshore wind" if (offon or "").strip().lower() == "offshore" else "onshore wind"
    if sector == BATTERY_SECTOR:
        return "battery"
    return sector.lower()


def _num(v):
    return f"{v:,.1f}".rstrip("0").rstrip(".") if abs(v - round(v)) > 1e-9 else f"{v:,.0f}"


def money_phrase_for_comment(inv_m):
    return "" if inv_m is None else f"${_num(inv_m)} million"


def investment_column_value_millions(inv_m):
    return round(float(inv_m), 2) if inv_m is not None else ""


def investment_section(sector_txt, inv_m, cap_mw):
    metric = round(inv_m / cap_mw, 2)
    return ("Estimated Project Investment:\n"
            f"-- Enerdatics calculates the project's investment to be {money_phrase_for_comment(inv_m)}, "
            f"representing an investment metric of ${metric:g} million/MW.\n"
            f"-- This estimate is in line with reported figures for {sector_txt} "
            "projects of this scale and size in the region.")


def timeline_section(sector_txt, fc_date, fc_estimated, op_date, op_estimated, is_operational):
    basis = (f"-- This estimate is based on reported development timelines for {sector_txt} "
             "projects of this scale and size in the region.")
    if is_operational and fc_estimated and op_date:                         # Case 3
        return ("Estimated Development Timeline:\n"
                f"-- The project reached commercial operations in {fmt_month_year(op_date)}.\n"
                "-- Further, Enerdatics believes that developer would have secured financing to fund "
                f"construction activities on the project by {fmt_month_year(fc_date)}.\n" + basis)
    if is_operational:
        return ""
    if op_estimated and fc_estimated:                                       # Case 4
        return ("Estimated Development Timeline:\n"
                f"-- The project is estimated to be commercially operational by {fmt_month_year(op_date)}.\n"
                "-- Further, Enerdatics believes that the developer will have secured financing to fund "
                f"construction activities on the project by {fmt_month_year(fc_date)}.\n" + basis)
    if op_estimated and not fc_estimated:                                   # Case 2
        return ("Estimated Development Timeline:\n"
                f"-- The project is estimated to be commercially operational by {fmt_month_year(op_date)}.\n"
                + basis + "\n"
                "-- Further, the developer will secure financing to fund construction activities "
                f"on the project by {fmt_month_year(fc_date)}.")
    if fc_estimated and not op_estimated and op_date:                       # Case 5
        return ("Estimated development timeline:\n"
                f"-- The project is expected to be commercially operational by {fmt_month_year(op_date)}.\n"
                "-- Further, Enerdatics believes that the developer will have secured financing to fund "
                f"construction activities on the project by {fmt_month_year(fc_date)}.\n" + basis)
    return ""


def build_comment(sector_txt, inv_estimated, inv_m, cap_mw,
                  fc_date, fc_estimated, op_date, op_estimated, is_operational):
    blocks = []
    if inv_estimated and inv_m is not None and cap_mw and cap_mw > 0:
        blocks.append(investment_section(sector_txt, inv_m, cap_mw))
    tl = timeline_section(sector_txt, fc_date, fc_estimated, op_date, op_estimated, is_operational)
    if tl:
        blocks.append(tl)
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


FIN_DEFAULT_BUCKETS = (BUCKET_CONSTR, BUCKET_OPERATIONAL, BUCKET_DECOMM)
KEEP_AS_IS_COLS = (C_OP_ST, C_OP_EST, C_OP_REP, C_FC_ST, C_FC_EST, C_FC_REP,
                   C_INV_ST, C_INV_EST, C_INV_REP, C_COMMENT, C_SHOWCOM)


def apply_lifecycle_flags(row, work_bucket, changes, sector="", repd_bucket=None):
    """SOP section 6 matrix. Fixed defaults are applied only where the SOP fixes them; researched
    values (Late-stage / Ready-to-build approvals and financing, PPA before operation) are kept."""
    cfd_round = normalize_multiselect(row.get(C_CFDAR) or "")
    cur = {c: (row.get(c) or "").strip() for c in (C_EIA, C_GRID, C_FIN, C_PPA)}
    label = (row.get(C_LIFE) or "").strip().lower()

    if work_bucket in FIN_DEFAULT_BUCKETS:
        _set_logged(row, C_EIA, "Yes", changes)
        _set_logged(row, C_GRID, "Yes", changes)
        _set_logged(row, C_FIN, "Yes - Complete", changes)
    elif work_bucket == BUCKET_EARLY:
        _set_logged(row, C_EIA, "No", changes)
        _set_logged(row, C_GRID, "No", changes)
        _set_logged(row, C_FIN, "No", changes)
    elif work_bucket == BUCKET_LATE:
        for c in (C_EIA, C_GRID):
            if cur[c] not in ("Yes", "No"):
                _set_logged(row, c, "No", changes)
        if cur[C_FIN] not in ("Yes - Complete", "Yes - Partial", "No"):
            _set_logged(row, C_FIN, "No", changes)
        # REPD planning permission granted = environmental approval obtained
        if repd_bucket in (BUCKET_LATE, BUCKET_CONSTR, BUCKET_OPERATIONAL) and row.get(C_EIA) != "Yes":
            _set_logged(row, C_EIA, "Yes", changes, label="EIA_approval_received (REPD permission granted)")
        if label == "late-stage" and row.get(C_EIA) == "No" and row.get(C_GRID) == "No":
            # house rules section 10: no approval evidence -> Early-stage
            _set_logged(row, C_LIFE, "Early-stage", changes,
                        label="life cycle (HR10: Late-stage without EIA or grid approval)")
            _set_logged(row, C_FIN, "No", changes)
            work_bucket = BUCKET_EARLY

    # PPA signed: Yes - Complete for Operational or a CfD award; otherwise researched (blank -> No)
    if work_bucket == BUCKET_OPERATIONAL or cfd_round:
        _set_logged(row, C_PPA, "Yes - Complete", changes)
    elif not cur[C_PPA]:
        _set_logged(row, C_PPA, "No", changes)

    # Financing type (SOP section 7): Debt + Equity from construction on, other instruments kept
    ft = multiselect_items(row.get(C_FINTYPE))
    if work_bucket in FIN_DEFAULT_BUCKETS:
        ft = ["Debt", "Equity"] + [t for t in ft if t not in ("Debt", "Equity", "Undisclosed")]
    elif (row.get(C_FIN) or "") == "No":
        ft = [t for t in ft if t in ("Pending finalization", "Financing planned/under negotiation")]
    _set_logged(row, C_FINTYPE, ";".join(ft), changes)

    apply_commercial_fields(row, sector, changes)
    return work_bucket


def apply_commercial_fields(row, sector, changes):
    """Power sales arrangement / Revenue mechanism (SOP section 7, house rules section 20)."""
    cfd_round = normalize_multiselect(row.get(C_CFDAR) or "")
    psa = multiselect_items(row.get(C_PSA))
    if sector == BATTERY_SECTOR:
        _set_logged(row, C_PSA, PSA_BATTERY, changes)
        rm = multiselect_items(row.get(C_REVMECH))
        rm = [t for t in rm if t.lower() not in REVMECH_MANAGED_VALUES]
        _set_logged(row, C_REVMECH, ";".join(rm) or REVMECH_UNKNOWN, changes)
        return
    rm = multiselect_items(row.get(C_REVMECH))
    if rm:
        _set_logged(row, C_REVMECH, "", changes, label="Revenue_mechanism cleared (battery-only field)")
    psa = [t for t in psa if t.lower() != "not applicable"]
    if cfd_round and PSA_CFD not in psa:
        psa = [PSA_CFD] + [t for t in psa if t != PSA_UNKNOWN]
    if len(psa) > 1 and PSA_UNKNOWN in psa:
        psa = [t for t in psa if t != PSA_UNKNOWN]
    _set_logged(row, C_PSA, ";".join(psa) or PSA_UNKNOWN, changes)


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
        anch = PRESENT_DATE
        if work_bucket == BUCKET_EARLY:
            new_fc = estimate_fc_forward(anch, cap, EARLY_MONTHS_PER_50, EARLY_FC_FLOOR_MONTHS)
        elif work_bucket == BUCKET_LATE:
            new_fc = estimate_fc_forward(anch, cap, LATE_MONTHS_PER_50, LATE_FC_FLOOR_MONTHS)
        elif work_bucket == BUCKET_CONSTR:
            new_fc = estimate_fc_construction(cap)
        elif work_bucket == BUCKET_OPERATIONAL:
            op = parse_db_date(row.get(C_OP_REP)) or parse_db_date(row.get(C_OP_EST))
            new_fc = backdate_fc_from_op(op, cap, sector) if op else None
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
            if new_op and new_op <= PRESENT_DATE:
                # e.g. Ready-to-build with a past reported FC: COD still has to be in the future
                new_op = eom_add(PRESENT_DATE, min_build_months(sector))
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
            op_anchor = parse_db_date(row.get(C_OP_REP))
            if work_bucket == BUCKET_OPERATIONAL and op_anchor:
                new_fc = min(backdate_fc_from_op(op_anchor, cap, sector), last_month_end_on_or_before(PRESENT_DATE))
            else:
                new_fc = estimate_fc_construction(cap)
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

    # Rule C: an ESTIMATED FC must not be in the past for Early/Late (a reported FC, e.g. on a
    # Ready-to-build row, is kept)
    if work_bucket in (BUCKET_EARLY, BUCKET_LATE):
        fc_date, fc_rep, op_date, op_rep = _read()
        if fc_date and fc_date <= PRESENT_DATE and not fc_rep:
            _clear_fc_and_reestimate(f"DATE CHECK Rule C: FC ({fmt_db_date(fc_date)}) in the past for "
                                     f"{BUCKET_TO_DB_LABEL.get(work_bucket, '')}")
            fc_date, fc_rep, op_date, op_rep = _read()
            if op_date and op_date <= PRESENT_DATE and fc_date:
                _clear_op_and_reestimate(fc_date, "DATE CHECK Rule C follow-up: OP still in the past")

    # Ordering checks EIA <= FC <= OP (reported OP of an operational project is locked)
    fc_date, fc_rep, op_date, op_rep = _read()
    op_locked = bool(op_rep and op_date is not None)

    if eia_date and fc_date and eia_date > fc_date and not fc_rep:
        new_fc = eom_add(eia_date, 1)
        if op_locked and new_fc >= op_date:
            new_fc = eom_add(op_date, -min_build_months(sector))
        row[C_FC_REP] = ""
        row[C_FC_EST] = fmt_db_date(new_fc)
        row[C_FC_ST] = ST_ESTIMATED
        changes.append(f"DATE CHECK: EIA ({fmt_db_date(eia_date)}) > FC ({fmt_db_date(fc_date)}), "
                       f"FC moved to {fmt_db_date(new_fc)}")
        fc_date = new_fc

    # FC -> OP gap: only an ESTIMATED date is moved; two reported dates are never touched.
    fc_date, fc_rep, op_date, op_rep = _read()
    if fc_date and op_date and not (fc_rep and op_rep):
        gap = months_between(fc_date, op_date)
        floor = min_build_months(sector)
        pre_constr = work_bucket in (BUCKET_EARLY, BUCKET_LATE)
        if op_rep and pre_constr and fc_date <= PRESENT_DATE:
            # not yet built: FC cannot be in the past; Late-stage floor from present, COD kept
            new_fc = eom_add(PRESENT_DATE, EARLY_FC_FLOOR_MONTHS if work_bucket == BUCKET_EARLY
                             else LATE_FC_FLOOR_MONTHS)
            latest = eom_add(op_date, -floor)          # must still leave the minimum build time
            if new_fc > latest:
                new_fc = latest if latest > PRESENT_DATE else eom_add(PRESENT_DATE, 1)
            row[C_FC_REP], row[C_FC_EST], row[C_FC_ST] = "", fmt_db_date(new_fc), ST_ESTIMATED
            changes.append(f"DATE CHECK: estimated FC {fmt_db_date(fc_date)} in the past for a "
                           f"pre-construction project, moved to {fmt_db_date(new_fc)}")
        elif op_rep and not pre_constr:
            # reported COD, estimated FC: FC = COD - MW/50 x 6 within [floor, 36] months before COD
            if gap < floor or gap > FC_BACKDATE_MAX_MONTHS:
                want = min(max(months_for_capacity(cap, OP_MONTHS_PER_50), floor), FC_BACKDATE_MAX_MONTHS)
                new_fc = eom_add(op_date, -want)
                if work_bucket in (BUCKET_OPERATIONAL, BUCKET_CONSTR):
                    new_fc = min(new_fc, last_month_end_on_or_before(PRESENT_DATE))
                row[C_FC_REP], row[C_FC_EST], row[C_FC_ST] = "", fmt_db_date(new_fc), ST_ESTIMATED
                changes.append(f"DATE CHECK: FC->reported OP gap {gap}mo outside {floor}-"
                               f"{FC_BACKDATE_MAX_MONTHS}mo, FC re-estimated {fmt_db_date(new_fc)}")
        elif gap < floor or gap > OP_MAX_MONTHS:
            # estimated COD: FC + MW/50 x 6 within [floor, 48] months
            new_op = estimate_op_from_fc(fc_date, cap, sector)
            if work_bucket != BUCKET_OPERATIONAL and new_op <= PRESENT_DATE:
                new_op = eom_add(PRESENT_DATE, CONSTR_MIN_FUTURE_OP_MO if work_bucket == BUCKET_CONSTR else 1)
            row[C_OP_REP], row[C_OP_EST], row[C_OP_ST] = "", fmt_db_date(new_op), ST_ESTIMATED
            changes.append(f"DATE CHECK: FC->OP gap {gap}mo outside {floor}-{OP_MAX_MONTHS}mo, "
                           f"OP re-estimated {fmt_db_date(new_op)}")

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
    """Kept for callers. PSA is no longer cleared when PPA signed = No: an unknown arrangement is
    'Yet to be determined/Unavailable' (SOP section 7), set by apply_commercial_fields."""
    apply_commercial_fields(row, sector, changes)


# =============================================================================
# Core per-row processor
# =============================================================================
def match_repd_id(join_raw, repd_idx):
    for tok in re.split(r"[;,]|\bOR\b|/", join_raw):
        tok = tok.strip()
        if tok.isdigit() and tok in repd_idx:
            return tok
    return join_raw if join_raw in repd_idx else None


def process_row(row, repd_idx, log, keep_lifecycle=False):
    """keep_lifecycle=True (Lead QC rows where REPD would move the lifecycle down): the DB
    lifecycle is kept, but every other rule still runs."""
    changes, tracked_cols = [], set()
    orig = dict(row)
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
    if bucket is not None and not keep_lifecycle:
        new_label, old_label = BUCKET_TO_DB_LABEL[bucket], (row.get(C_LIFE) or "").strip()
        if new_label != old_label:
            row[C_LIFE] = new_label
            changes.append(f"life cycle: '{old_label}' -> '{new_label}'")
            tracked_cols.add(C_LIFE)
    cur_label = (row.get(C_LIFE) or "").strip().lower()
    work_bucket = DB_LABEL_TO_BUCKET.get(cur_label, bucket)

    # 1b. lifecycle flags + Show SPV
    # 1a. house rules section 10: a reported financial close that has passed -> In construction.
    # Ready-to-build rows are kept as they are (team decision).
    fc_rep_now = parse_db_date(row.get(C_FC_REP)) or parse_repd_date(rrow.get(R_UCONST))
    awaiting = repd_awaiting_construction(rrow)
    if (work_bucket in (BUCKET_EARLY, BUCKET_LATE) and cur_label != "ready-to-build" and not awaiting
            and fc_rep_now and fc_rep_now <= PRESENT_DATE
            and (row.get(C_FC_ST) or "").strip() == ST_REPORTED):
        _set_logged(row, C_LIFE, "In construction", changes,
                    label=f"life cycle (HR10: reported FC {fmt_db_date(fc_rep_now)} has passed)")
        tracked_cols.add(C_LIFE)
        work_bucket = BUCKET_CONSTR

    # 1b. lifecycle flags + Show SPV
    psa_before = (row.get(C_PSA) or "").strip()
    life_before = (row.get(C_LIFE) or "").strip()
    work_bucket = apply_lifecycle_flags(row, work_bucket, changes, sector=sector, repd_bucket=bucket)
    if (row.get(C_LIFE) or "").strip() != life_before:
        tracked_cols.add(C_LIFE)
    if (row.get(C_PSA) or "").strip() != psa_before:
        tracked_cols.add(C_PSA)
    apply_show_spv(row, rrow, changes)

    # 2. terminal. Cancelled/Suspended: statuses 'Not reported', no dates, no comment (HR5).
    # Decommissioned is a distinct stage (SOP section 9): its data is kept, not wiped or estimated.
    if work_bucket == BUCKET_DECOMM:
        log.append({"Project name": name, "State data plant ID": rid,
                    "Action": "UPDATED (decommissioned)" if changes else "no change",
                    "Detail": " | ".join(changes)})
        return row, bool(changes), tracked_cols
    if work_bucket in TERMINAL_BUCKETS:
        clear_estimates_and_dates(row, changes,
                                  "terminal lifecycle -> cleared estimates, status set to 'Not reported'")
        if changes:
            log.append({"Project name": name, "State data plant ID": rid,
                        "Action": "UPDATED (terminal)", "Detail": " | ".join(changes)})
        return row, bool(changes), tracked_cols

    # 3. reported dates from REPD
    repd_op = parse_repd_date(rrow.get(R_OPER))
    repd_fc = parse_repd_date(rrow.get(R_UCONST))
    op_status = (row.get(C_OP_ST) or "").strip()
    fc_status = (row.get(C_FC_ST) or "").strip()
    if awaiting:
        # planned dates that have passed are not events: ignored, estimated forward instead
        repd_op = repd_op if repd_op and repd_op > PRESENT_DATE else None
        repd_fc = repd_fc if repd_fc and repd_fc > PRESENT_DATE else None
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

    # 3a. a reported FC (REPD "Under Construction") after the reported COD is wrong: remove it,
    #     so step 4 estimates FC (e.g. operational: COD - MW/50 x 6)
    fc_r, op_r = parse_db_date(row.get(C_FC_REP)), parse_db_date(row.get(C_OP_REP))
    if fc_status == ST_REPORTED and op_status == ST_REPORTED and fc_r and op_r and fc_r > op_r:
        row[C_FC_REP], row[C_FC_EST], row[C_FC_ST] = "", "", ""
        fc_status = ""
        changes.append(f"reported FC {fmt_db_date(fc_r)} after reported COD {fmt_db_date(op_r)}: "
                       "removed, FC estimated")

    # 3b. Awaiting Construction (kept Late-stage): a passed reported FC is pushed forward
    if awaiting and work_bucket == BUCKET_LATE and cur_label != "ready-to-build":
        fc_r = parse_db_date(row.get(C_FC_REP))
        if fc_status == ST_REPORTED and fc_r and fc_r <= PRESENT_DATE:
            row[C_FC_REP], row[C_FC_EST], row[C_FC_ST] = "", "", ""
            fc_status = ""
            changes.append(f"REPD Awaiting Construction: passed FC {fmt_db_date(fc_r)} removed, "
                           "FC estimated forward")

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

    # 3d. operational without reported OP date -> nothing estimated, dates / investment / comment
    # kept exactly as they are in the DB (team decision)
    if work_bucket == BUCKET_OPERATIONAL and op_reported is None:
        for c in KEEP_AS_IS_COLS:
            if c in orig:
                row[c] = orig[c]
        tracked_cols.add("__KEEP_AS_IS__")
        changes.append("operational w/o reported OP date -> dates, investment and comment kept as is")
        log.append({"Project name": name, "State data plant ID": rid,
                    "Action": "UPDATED" if changes else "no change", "Detail": " | ".join(changes)})
        return row, bool(changes), tracked_cols

    # 4. FC estimate
    fc_date, fc_estimated = fc_reported, (fc_status != ST_REPORTED)
    if fc_status != ST_REPORTED:
        if work_bucket == BUCKET_OPERATIONAL and op_reported:
            fc_date = backdate_fc_from_op(op_reported, cap, sector)
        elif work_bucket == BUCKET_CONSTR:
            fc_date = estimate_fc_construction(cap)
        elif work_bucket == BUCKET_EARLY:
            fc_date = estimate_fc_forward(PRESENT_DATE, cap, EARLY_MONTHS_PER_50, EARLY_FC_FLOOR_MONTHS)
        elif work_bucket == BUCKET_LATE:
            fc_date = estimate_fc_forward(PRESENT_DATE, cap, LATE_MONTHS_PER_50, LATE_FC_FLOOR_MONTHS)
        if fc_date:
            row[C_FC_EST], row[C_FC_REP], row[C_FC_ST] = fmt_db_date(fc_date), "", ST_ESTIMATED
            fc_estimated = True
            changes.append(f"FC estimated: {fmt_db_date(fc_date)}")

    # 5. OP estimate
    op_date, op_estimated = op_reported, (op_status != ST_REPORTED)
    if op_status != ST_REPORTED and work_bucket != BUCKET_OPERATIONAL:
        if work_bucket == BUCKET_CONSTR:
            base_fc = fc_date or estimate_fc_construction(cap)
            cand = estimate_op_from_fc(base_fc, cap, sector)
            min_fut = eom_add(PRESENT_DATE, CONSTR_MIN_FUTURE_OP_MO)
            op_date = max(cand, min_fut)
        elif fc_date and work_bucket in (BUCKET_EARLY, BUCKET_LATE, BUCKET_CONSTR):
            op_date = estimate_op_from_fc(fc_date, cap, sector)
        if op_date and work_bucket != BUCKET_OPERATIONAL:
            row[C_OP_EST], row[C_OP_REP], row[C_OP_ST] = fmt_db_date(op_date), "", ST_ESTIMATED
            op_estimated = True
            changes.append(f"OP estimated from FC: {fmt_db_date(op_date)}")

    # 5b. PSA / revenue mechanism (already set in 1b; PSA is no longer cleared when PPA = No)

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


_SECTION_HEAD = re.compile(r"^[^\w\r\n]*e?stimated[ \t]+(project[ \t]+investment|development[ \t]+timeline)"
                           r"[ \t]*:", re.I)


def _paragraphs(text):
    return [p for p in re.split(r"\r?\n[ \t]*\r?\n", (text or "").strip()) if p.strip()]


def _section_kind(par):
    m = _SECTION_HEAD.match(par.strip())
    return None if not m else ("inv" if "investment" in m.group(1).lower() else "tl")


def timeline_paragraph(comment: str) -> str:
    for p in _paragraphs(comment):
        if _section_kind(p) == "tl":
            return p.strip()
    return ""


def merge_generated_sections(existing: str, inv_sec: str, tl_sec: str) -> str:
    """QC rows: the 'Estimated Project Investment' and 'Estimated Development Timeline' paragraphs
    are replaced in place (or removed when nothing is estimated); every other analyst paragraph is
    kept as written. A missing section is added: investment before the timeline."""
    pars = _paragraphs(existing)
    new = {"inv": (inv_sec or "").strip(), "tl": (tl_sec or "").strip()}
    out, placed = [], set()
    for p in pars:
        k = _section_kind(p)
        if k is None:
            out.append(p.strip())
        elif k not in placed:
            placed.add(k)
            if new[k]:
                out.append(new[k])
    if new["inv"] and "inv" not in placed:
        idx = next((i for i, p in enumerate(out) if p == new["tl"]), len(out))
        out.insert(idx, new["inv"])
    if new["tl"] and "tl" not in placed:
        out.append(new["tl"])
    return "\n\n".join(out)


def replace_timeline_in_place(existing: str, tl_sec: str) -> str:
    inv = next((p for p in _paragraphs(existing) if _section_kind(p) == "inv"), "")
    return merge_generated_sections(existing, inv, tl_sec)


def qc_process(row: dict, repd_idx: dict, log: list) -> dict:
    """Lead QC rows: the DB capacity is never changed (a REPD difference is only logged); lifecycle
    only moves UP (a REPD downgrade keeps the DB lifecycle, but every other rule still runs); a valid
    estimated or reported investment is kept; the generated comment sections are rebuilt."""
    original = dict(row)
    rid = match_repd_id(_cmp(row.get(C_JOIN)), repd_idx)
    keep = False
    if rid:
        rr = repd_idx[rid]
        repd_cap, db_cap = parse_capacity(rr.get(R_CAP)), parse_capacity(row.get(C_CAP))
        if repd_cap is not None and (db_cap is None or abs(repd_cap - db_cap) > 0.01):
            log.append({"Project name": row.get(C_NAME, ""), "State data plant ID": rid,
                        "Action": "QC: capacity not changed",
                        "Detail": f"REPD capacity {repd_cap:g} MW vs DB {row.get(C_CAP, '')} MW (Lead QC)"})
        bucket = REPD_STATUS_TO_BUCKET.get((rr.get(R_STATUS) or "").strip().lower())
        new_label = BUCKET_TO_DB_LABEL.get(bucket, "")
        cur = (original.get(C_LIFE) or "").strip()
        if new_label and new_label.lower() != cur.lower() and                 LIFECYCLE_RANK.get(new_label.lower(), -1) <= LIFECYCLE_RANK.get(cur.lower(), -1):
            keep = True
            log.append({"Project name": row.get(C_NAME, ""), "State data plant ID": rid,
                        "Action": "QC: lifecycle downgrade not applied",
                        "Detail": f"REPD points to {cur} -> {new_label}; DB lifecycle kept"})
    new_row, _c, tracked = process_row(dict(row), repd_idx, log, keep_lifecycle=keep)
    if "__NO_MATCH__" in tracked:
        return original
    new_row[C_CAP] = original.get(C_CAP, "")
    if "__KEEP_AS_IS__" in tracked:
        return new_row
    cap = parse_capacity(new_row.get(C_CAP))
    sector = (new_row.get(C_SECTOR) or "").strip()
    offon = new_row.get(C_OFFON)

    # investment: keep the QC'd value unless it breaks a rule
    o_st = (original.get(C_INV_ST) or "").strip()
    o_est, o_rep = parse_money(original.get(C_INV_EST)), parse_money(original.get(C_INV_REP))
    bucket_now = DB_LABEL_TO_BUCKET.get((new_row.get(C_LIFE) or "").strip().lower())
    terminal = bucket_now in TERMINAL_BUCKETS
    if not terminal and o_rep is not None and o_st in (ST_REPORTED, ST_ESTIMATED):
        new_row[C_INV_ST], new_row[C_INV_REP], new_row[C_INV_EST] = ST_REPORTED, original.get(C_INV_REP), ""
    elif (not terminal and o_st == ST_ESTIMATED and o_est is not None
          and metric_allowed(sector, offon, bucket_now, o_est, cap)):
        new_row[C_INV_ST], new_row[C_INV_EST], new_row[C_INV_REP] = ST_ESTIMATED, original.get(C_INV_EST), ""
    # otherwise process_row's estimate (or terminal wipe) stands

    if terminal:
        new_row[C_COMMENT] = ""
    elif sector != HYDRO_SECTOR:
        st = lambda c: (new_row.get(c) or "").strip() == ST_ESTIMATED
        inv_m = parse_money(new_row.get(C_INV_EST))
        sect = sector_phrase(sector, offon)
        inv_sec = investment_section(sect, inv_m, cap) if st(C_INV_ST) and inv_m and cap else ""
        fc_d = parse_db_date(new_row.get(C_FC_EST)) or parse_db_date(new_row.get(C_FC_REP))
        op_d = parse_db_date(new_row.get(C_OP_EST)) or parse_db_date(new_row.get(C_OP_REP))
        is_op = bucket_now == BUCKET_OPERATIONAL
        tl_sec = timeline_section(sect, fc_d, st(C_FC_ST), op_d, st(C_OP_ST), is_op)
        new_row[C_COMMENT] = merge_generated_sections(original.get(C_COMMENT, ""), inv_sec, tl_sec)
    return new_row


# =============================================================================
# Final pass on every output row (house rules 2, 6, 8, 14, 20)
# =============================================================================
_PLAIN_MAP = str.maketrans({"‘": "'", "’": "'", "‚": "'", "‛": "'",
                            "“": '"', "”": '"', "–": "-", "—": "-",
                            "−": "-", " ": " ", "…": "...", "ß": "ss",
                            "ø": "o", "Ø": "O", "æ": "ae", "Æ": "AE",
                            "ł": "l", "Ł": "L", "đ": "d", "Đ": "D",
                            "þ": "th", "ð": "d", "œ": "oe", "Œ": "OE"})
_BOOL_COLS = (C_SHOWSPV, C_SHOWCOM)
MULTISELECT_COLS = (C_PSA, C_CFDAR, C_REVMECH, C_FINTYPE)
_TRUE_WORDS = {"true", "yes", "1", "y"}
_BATTERY_NAME_RE = re.compile(r"BESS|Battery|Energy Storage", re.I)


_CP1252_RUN = re.compile("[\u0080-ÿŒœŠšŸŽžƒˆ˜"
                         "–—‘-„†-•…‰‹›€™]{2,}")


def fix_mojibake(v):
    """UTF-8 text that was read as cp1252 (once or twice), e.g. 'â€™' or 'Ã¢â‚¬â„¢' -> '’'."""
    def _one(m):
        s = m.group(0)
        for _ in range(3):
            try:
                t = s.encode("cp1252").decode("utf-8")
            except (UnicodeEncodeError, UnicodeDecodeError):
                break
            if t == s:
                break
            s = t
        return s
    return _CP1252_RUN.sub(_one, v)


def plain_letters(v):
    """House rules 14: accented/special letters -> plain letters (U+FFFD is left for the QC flags)."""
    if not isinstance(v, str) or v.isascii():
        return v
    # double-encoded replacement character already stored in Quickbase -> one marker (QC Flags)
    v = fix_mojibake(v.replace("ï¿½", "�"))
    s = unicodedata.normalize("NFD", v.translate(_PLAIN_MAP))      # letters only; symbols such as £ kept
    return unicodedata.normalize("NFC", "".join(ch for ch in s if not unicodedata.combining(ch)))


def finalize_row(row: dict) -> dict:
    sector = (row.get(C_SECTOR) or "").strip()
    for c, v in list(row.items()):
        if not str(c).startswith("_") and isinstance(v, str):
            row[c] = plain_letters(v)
    for col in (C_PSA, C_CFDAR, C_REVMECH, C_FINTYPE):
        if col in row:
            row[col] = normalize_multiselect(row[col])
    if sector in (BATTERY_SECTOR, "Solar", "Wind"):
        apply_commercial_fields(row, sector, [])
    for col in _BOOL_COLS:
        if col in row and not _blank(row.get(col)):
            row[col] = "Yes" if str(row[col]).strip().lower() in _TRUE_WORDS else "No"
    if sector == BATTERY_SECTOR and row.get(C_NAME) and not _BATTERY_NAME_RE.search(row[C_NAME]):
        row[C_NAME] = row[C_NAME].rstrip() + " BESS"
    # status <-> value consistency; a reported value always wins over an old estimate
    for st, est, rep in ((C_INV_ST, C_INV_EST, C_INV_REP), (C_FC_ST, C_FC_EST, C_FC_REP),
                         (C_OP_ST, C_OP_EST, C_OP_REP)):
        s_ = (row.get(st) or "").strip()
        if s_ == ST_REPORTED and not _blank(row.get(rep)):
            row[est] = ""
        elif s_ == ST_ESTIMATED and not _blank(row.get(est)):
            row[rep] = ""
        elif s_ == ST_ESTIMATED and not _blank(row.get(rep)):
            row[st] = ST_REPORTED
    for col in (C_FC_EST, C_OP_EST):
        d = parse_db_date(row.get(col))
        if d and d != end_of_month(d):
            row[col] = fmt_db_date(end_of_month(d))
    # environmental approval date: removed when it is in the future or EIA approval = No
    env = parse_db_date(row.get(C_ENVDATE))
    if env and (env > PRESENT_DATE or (row.get(C_EIA) or "").strip() == "No"):
        row[C_ENVDATE] = ""
    any_est = any((row.get(c) or "").strip() == ST_ESTIMATED for c in (C_INV_ST, C_FC_ST, C_OP_ST))
    row[C_SHOWCOM] = "Yes" if (any_est and not _blank(row.get(C_COMMENT))) else "No"
    return row


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
        new_row = finalize_row(new_row)

        def same(c):
            a, b = _cmp(original.get(c)), _cmp(new_row.get(c))
            if c in MULTISELECT_COLS:          # "Debt; Equity" == "Debt;Equity" == "['Debt', 'Equity']"
                return set(multiselect_items(a)) == set(multiselect_items(b))
            return a == b
        changed_cols = {c for c in set(original) | set(new_row) if not str(c).startswith("_") and not same(c)}
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
