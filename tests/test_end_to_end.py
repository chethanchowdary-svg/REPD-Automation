from pathlib import Path

import pandas as pd

from repd_qb import id_matcher, main

FX = Path(__file__).parent / "fixtures"


def fake_gpt(payload, key, model):
    e = payload["repd_entry"]
    for c in payload["candidates"]:
        if id_matcher.norm_name(c["name"]) == id_matcher.norm_name(e["name"]):
            return {"match_project_id": c["project_id"], "confidence": "high", "reason": "same name and county"}
    return {"match_project_id": None, "confidence": "low", "reason": "no match"}


def test_full_run(tmp_path, monkeypatch):
    monkeypatch.setattr(id_matcher, "ask_gpt", fake_gpt)
    monkeypatch.setattr(id_matcher.CFG, "OPENAI_API_KEY", "test")
    monkeypatch.setattr(id_matcher.CFG, "OPENCAGE_API_KEY", "")
    main.main(["--repd-csv", str(FX / "repd.csv"), "--db-csv", str(FX / "db.csv"),
               "--run-date", "2026-09-25", "--out-dir", str(tmp_path)])
    out = next(tmp_path.glob("REPD_QB_Update_*.xlsx"))
    x = pd.ExcelFile(out)
    assert {"Lifecycle Capacity", "Lead QC Update", "New Projects", "ID Matches (GPT)"} <= set(x.sheet_names)

    new = pd.read_excel(x, "New Projects", dtype=str).fillna("")
    assert set(new["State_data_plant_ID"]) == {"1005", "1007", "1008"}      # 1004 moved to existing by GPT

    qc = pd.read_excel(x, "Lead QC Update", dtype=str).fillna("")
    z = qc[qc["Project ID"] == "P-6"].iloc[0]
    assert z["Project_life_cycle"] == "Late-stage" and z["Power_Capacity_(MW)"] == "50"
    assert "P-6" not in set(pd.read_excel(x, "Lifecycle Capacity", dtype=str)["Project ID"])

    other = pd.concat([pd.read_excel(x, s, dtype=str) for s in ("Lifecycle Capacity", "Other Changes")]).fillna("")
    d = other[other["Project ID"] == "P-500"].iloc[0]
    assert d["State_data_plant_ID"] == "1004"
    m = pd.read_excel(x, "ID Matches (GPT)", dtype=str)
    assert (m["Applied"] == "Yes").sum() == 1


def test_qc_keeps_analyst_text_and_blocks_downgrade():
    from repd_qb import updater
    import datetime as dt
    updater.PRESENT_DATE = dt.date(2026, 9, 25)
    repd = pd.read_csv(FX / "repd.csv", dtype=str, keep_default_na=False)
    idx = {r["Ref ID"]: r for _, r in repd.iterrows()}
    log = []
    row = {"Project_name": "Zeta Storage", "State_data_plant_ID": "1006", "Power_Capacity_(MW)": "45",
           "Sector": "Battery", "Project_life_cycle": "Early-stage", "Project_investment_status": "Reported",
           "Project_investment_($M)___Reported": "70",
           "Analyst_comments": "Other details:\n-- keep me\n\nEstimated Development Timeline:\n-- old"}
    out = updater.qc_process(dict(row), idx, log)
    assert out["Project_life_cycle"] == "Late-stage" and out["Power_Capacity_(MW)"] == "50"
    assert out["Analyst_comments"].startswith("Other details:\n-- keep me\n\nEstimated Development Timeline:")
    assert "-- old" not in out["Analyst_comments"]
    assert out["Project_investment_($M)___Reported"] == "70"
    down = dict(row, **{"Project_life_cycle": "Operational"})
    out2 = updater.qc_process(down, idx, log)
    assert out2["Project_life_cycle"] == "Operational" and out2["Power_Capacity_(MW)"] == "50"
