"""REPD -> Quickbase quarterly run, driven by ONE Quickbase report (Country = United Kingdom, all rows).

The same report feeds every step:
  - rows WITH a State_data_plant_ID  -> updated from REPD (Lead QC rows go to their own sheet)
  - all Ref IDs in the report        -> decide which REPD projects are new
  - rows WITHOUT a State_data_plant_ID -> candidates for GPT matching of new REPD projects

  python -m repd_qb.main                                   # REPD from GOV.UK + Quickbase API
  python -m repd_qb.main --repd-csv repd.csv --db-csv db.csv  # offline
"""
from __future__ import annotations

import argparse
import datetime as dt
import sys
from pathlib import Path

from . import config as C
from . import id_matcher, new_projects, sources, updater, writer

ROOT = Path(__file__).resolve().parent.parent


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="UK REPD -> Quickbase upload sheets")
    p.add_argument("--repd-csv", help="Local REPD quarterly CSV (skips GOV.UK download)")
    p.add_argument("--db-csv", help="CSV export of the Quickbase report (skips the API)")
    p.add_argument("--no-id-match", action="store_true", help="Skip the GPT matching step")
    p.add_argument("--run-date", help="YYYY-MM-DD used as present date (default: today)")
    p.add_argument("--out-dir", default="outputs")
    return p.parse_args(argv)


def main(argv=None) -> int:
    a = parse_args(argv)
    out = Path(a.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    run_date = dt.date.fromisoformat(a.run_date) if a.run_date else dt.date.today()
    updater.PRESENT_DATE = new_projects.PRESENT_DATE = run_date

    # 1. sources
    if a.repd_csv:
        repd_name, repd = Path(a.repd_csv).name, sources.read_csv_path(a.repd_csv)
    else:
        repd_name, repd = sources.fetch_repd_csv()
        repd.to_csv(out / repd_name, index=False)
    print(f"REPD: {repd_name} ({len(repd):,} rows)")
    raw = sources.read_csv_path(a.db_csv) if a.db_csv else sources.fetch_quickbase(C.QB_REPORT_ID)
    if not a.db_csv:
        raw.to_csv(out / "quickbase_snapshot.csv", index=False)

    db = updater.prepare_db(raw)
    has_id = db[updater.C_JOIN].astype(str).str.strip() != ""
    repd_idx = {str(r[updater.R_REFID]).strip(): r for _, r in repd.iterrows()}

    # 2. new projects (REPD Ref IDs not anywhere in the report)
    new_rows, stats = new_projects.build_new_projects(new_projects.db_id_set(db), repd)

    # 3. GPT matching: new REPD projects vs report rows without a Ref ID
    forced, id_rows = [], []
    if not a.no_id_match and (~has_id).any():
        applied, decisions, mstats, _m = id_matcher.match(
            new_rows, db[~has_id], repd_idx, ROOT / "data" / "id_match_overrides.csv")
        id_rows = id_matcher.decisions_table(decisions)
        stats.update(mstats)
        for rid, i in applied.items():
            row = db.loc[i].to_dict()
            row[updater.C_JOIN] = rid
            forced.append(row)
        new_rows = [r for r in new_rows if r["State_data_plant_ID"] not in applied]

    # 4. update rows that have a Ref ID (+ GPT-matched rows); Lead QC rows get their own sheet
    upd = updater.run_update(db[has_id], repd, None, forced)

    stats = {
        "REPD file": repd_name,
        "Run date": run_date.isoformat(),
        "Quickbase report rows": len(db),
        "Rows with State_data_plant_ID": int(has_id.sum()),
        "Rows without State_data_plant_ID": int((~has_id).sum()),
        "Lifecycle/capacity changes (non-QC)": len(upd["lifecycle_capacity"]),
        "Site-area changes (non-QC)": len(upd["site_area"]),
        "Other changes (non-QC)": len(upd["other"]),
        "Lead QC rows with changes": len(upd["qc"]),
        "New projects": len(new_rows),
        **stats,
    }
    path = out / f"REPD_QB_Update_{run_date:%Y%m%d}.xlsx"
    writer.write(path, upd, new_rows, new_projects.OUT_COLS, id_rows, stats)
    for k, v in stats.items():
        print(f"  {k}: {v}")
    print(f"Output: {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
