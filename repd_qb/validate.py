"""QC Flags tab: problems the rules cannot fix on their own and an analyst has to decide.
Runs on the final rows (after every rule), so whatever is listed here is still in the upload sheets."""
from __future__ import annotations

import re
from collections import defaultdict

from . import updater as U

# replacement char, or typical mojibake of UTF-8 read as cp1252 (e.g. "â€™", "Ã©", "Â£")
_BROKEN = re.compile("�|[ÂÃâ][\u0080-¿€‚-›Œœ"
                     "ŠšŸŽžˆ˜™]")

COLS = ["Sheet", "Project ID", "Project name", "State data plant ID", "Sector", "Life cycle", "Flag", "Detail"]


def _pid(r):
    return r.get("Project ID") or r.get("Project_ID") or ""


def _flag(out, sheet, r, flag, detail):
    out.append({"Sheet": sheet, "Project ID": _pid(r), "Project name": r.get(U.C_NAME, ""),
                "State data plant ID": r.get(U.C_JOIN, ""), "Sector": r.get(U.C_SECTOR, ""),
                "Life cycle": r.get(U.C_LIFE, ""), "Flag": flag, "Detail": detail})


def row_flags(r, sheet, out, today):
    cap = U.parse_capacity(r.get(U.C_CAP))
    life = (r.get(U.C_LIFE) or "").strip()
    sector = (r.get(U.C_SECTOR) or "").strip()
    for c, v in r.items():
        if not isinstance(v, str):
            continue
        m = _BROKEN.search(v)
        if m:
            i = m.start()
            _flag(out, sheet, r, "Broken character", f"{c}: ...{v[max(0, i - 30):i + 15]}... (retype the original letter)")
    psa = U.multiselect_items(r.get(U.C_PSA))
    if U.PSA_CFD in psa and not U.normalize_multiselect(r.get(U.C_CFDAR)):
        _flag(out, sheet, r, "CfD without allocation round", "Power sales arrangement has CfD, AR blank")
    env = U.parse_db_date(r.get(U.C_ENVDATE))
    if env and env > today:
        _flag(out, sheet, r, "Environmental approval date in future", U.fmt_db_date(env))
    if env and (r.get(U.C_EIA) or "") == "No" and sheet != "New Projects":
        _flag(out, sheet, r, "Approval date but EIA = No", U.fmt_db_date(env))
    fc = U.parse_db_date(r.get(U.C_FC_REP)) or U.parse_db_date(r.get(U.C_FC_EST))
    op = U.parse_db_date(r.get(U.C_OP_REP)) or U.parse_db_date(r.get(U.C_OP_EST))
    if fc and op:
        gap = U.months_between(fc, op)
        if op < fc:
            _flag(out, sheet, r, "COD before FC", f"FC {U.fmt_db_date(fc)}, COD {U.fmt_db_date(op)}")
    if sector == U.BATTERY_SECTOR and (r.get(U.C_COLOC) or "") == "Potential Hybridization":
        _flag(out, sheet, r, "Battery co-located storage", "Potential Hybridization: use Yes or No (HR20)")


def duplicate_flags(rows, sheet, out):
    by_id, by_name = defaultdict(list), defaultdict(list)
    for r in rows:
        sid = str(r.get(U.C_JOIN) or "").strip()
        if sid:
            by_id[(sid, r.get(U.C_SECTOR, ""))].append(r)
        n = str(r.get(U.C_NAME) or "").strip().lower()
        if n:
            by_name[n].append(r)
    for (sid, _s), g in by_id.items():
        if len(g) > 1:
            for r in g:
                _flag(out, sheet, r, "Duplicate REPD ID (same sector)",
                      "also on " + ", ".join(_pid(x) for x in g if x is not r))
    for _n, g in by_name.items():
        if len(g) > 1:
            for r in g:
                _flag(out, sheet, r, "Duplicate project name", "also on " + ", ".join(_pid(x) for x in g if x is not r))


def build(upd: dict, new_rows: list, today) -> list[dict]:
    out = []
    full = upd.get("full", [])
    qc_ids = {id(r) for r in upd.get("qc", [])}
    for r in full:
        row_flags(r, "Lead QC Update" if id(r) in qc_ids else "Full DB", out, today)
    for r in new_rows:
        row_flags(r, "New Projects", out, today)
    log = upd.get("log")
    if log is not None and len(log):
        for e in log.to_dict("records"):
            if e.get("Action") == "QC: capacity not changed":
                out.append({"Sheet": "Lead QC Update", "Project ID": "", "Project name": e.get("Project name", ""),
                            "State data plant ID": e.get("State data plant ID", ""), "Sector": "",
                            "Life cycle": "", "Flag": "REPD capacity differs (Lead QC, not applied)",
                            "Detail": e.get("Detail", "")})
    return out
