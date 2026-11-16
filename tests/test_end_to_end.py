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
    assert z["Project_life_cycle"] == "Late-stage" and z["Power_Capacity_(MW)"] == "45"   # Lead QC: capacity never changed
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
    assert out["Project_life_cycle"] == "Late-stage" and out["Power_Capacity_(MW)"] == "45"
    assert out["Analyst_comments"].startswith("Other details:\n-- keep me\n\nEstimated Development Timeline:")
    assert "-- old" not in out["Analyst_comments"]
    assert out["Project_investment_($M)___Reported"] == "70"
    down = dict(row, **{"Project_life_cycle": "Operational"})
    out2 = updater.qc_process(down, idx, log)
    assert out2["Project_life_cycle"] == "Operational" and out2["Power_Capacity_(MW)"] == "45"


# ---- house rules (Projects-Entry-Bot house_rules.md) ----
def _qc_row(**kw):
    base = {"Project ID": "Q-1", "Project_name": "Test", "State_data_plant_ID": "1006", "Lead QC": "true",
            "Power_Capacity_(MW)": "45", "Sector": "Solar", "Offshore/Onshore": "Onshore",
            "Project_life_cycle": "Late-stage"}
    base.update(kw)
    return base


def _run_qc(row):
    from repd_qb import updater
    import datetime as dt
    updater.PRESENT_DATE = dt.date(2026, 11, 16)
    repd = pd.read_csv(FX / "repd.csv", dtype=str, keep_default_na=False)
    idx = {r["Ref ID"]: r for _, r in repd.iterrows()}
    return updater.finalize_row(updater.qc_process(dict(row), idx, []))


def test_battery_psa_name_and_revenue_mechanism():
    out = _run_qc(_qc_row(Sector="Battery", Project_name="Saticoy", Power_sales_arrangement="Not applicable"))
    assert out["Power_sales_arrangement"] == "Not Applicable"                 # HR20
    assert out["Project_name"] == "Saticoy BESS"                              # HR2
    assert out["Revenue_mechanism"] == "Yet to be determined/Unavailable"     # SOP 7


def test_formats_and_plain_letters():
    out = _run_qc(_qc_row(Show_SPV="True", SPV="Pr\u00f6sen \u2013 Gr\u00f6ditz",
                          Financing_type="['Debt', 'Equity']"))
    assert out["Show_SPV"] == "Yes" and out["SPV"] == "Prosen - Groditz"
    assert out["Show_analyst_comment"] in ("Yes", "No")


def test_reported_investment_clears_stale_estimate_and_comment():
    out = _run_qc(_qc_row(**{"Project_investment_status": "Reported", "Project_investment_($M)___Reported": "70",
                             "Project_investment_($M)___Estimated": "50",
                             "Analyst_comments": "Estimated Project Investment:\n-- old\n\nOther:\n-- keep"}))
    assert out["Project_investment_($M)___Estimated"] == ""
    assert "Estimated Project Investment" not in out["Analyst_comments"]
    assert "Other:\n-- keep" in out["Analyst_comments"]


def test_out_of_range_estimate_is_re_estimated_and_comment_matches():
    out = _run_qc(_qc_row(**{"Project_investment_status": "Estimated", "Project_investment_($M)___Estimated": "130"}))
    inv = float(out["Project_investment_($M)___Estimated"])
    assert 1.0 <= inv / 45 <= 1.2                                             # HR7 solar range
    assert f"${inv:,.0f} million" in out["Analyst_comments"]


def test_estimated_dates_are_month_end_and_fc_not_in_past():
    import datetime as dt
    from repd_qb import updater
    out = _run_qc(_qc_row(Project_life_cycle="Ready-to-build", EIA_approval_received="Yes",
                          Grid_connection_approved="Yes", Financing_secured="Yes - Partial",
                          **{"Financial_close_date_status": "Estimated", "Financial_close_(estimated)": "08/30/2026"}))
    fc = updater.parse_db_date(out["Financial_close_(estimated)"])
    assert fc > dt.date(2026, 11, 16) and fc == updater.end_of_month(fc)
    assert out["Project_life_cycle"] == "Ready-to-build"                      # REPD downgrade not applied


def test_mojibake_repaired_then_plain():
    from repd_qb.updater import plain_letters
    assert plain_letters("Crown Estate\u00e2\u20ac\u2122s") == "Crown Estate's"
    assert plain_letters("UK\u00c3\u00a2\u00e2\u201a\u00ac\u00e2\u201e\u00a2s") == "UK's"
    assert plain_letters("\u00a313.2 billion") == "\u00a313.2 billion"
    assert plain_letters("Orkney\u00ef\u00bf\u00bds") == "Orkney\ufffds"          # left for QC Flags


def test_rtb_with_past_reported_fc_is_kept():
    out = _run_qc(_qc_row(Project_life_cycle="Ready-to-build", EIA_approval_received="Yes",
                          Grid_connection_approved="Yes", Financing_secured="Yes - Partial",
                          **{"Financial_close_date_status": "Reported",
                             "Financial_close_date_(reported)": "09/30/2026",
                             "Operational_date_status": "Estimated", "Operational_date_(estimated)": "10/31/2026"}))
    from repd_qb import updater
    assert out["Project_life_cycle"] == "Ready-to-build"
    assert out["Financial_close_date_(reported)"] == "09/30/2026"
    assert updater.parse_db_date(out["Operational_date_(estimated)"]) > updater.PRESENT_DATE


def test_operational_without_cod_kept_as_is():
    from repd_qb import updater
    import datetime as dt
    updater.PRESENT_DATE = dt.date(2026, 11, 16)
    repd = pd.read_csv(FX / "repd.csv", dtype=str, keep_default_na=False)
    idx = {r["Ref ID"]: r for _, r in repd.iterrows()}
    fake = dict(repd.iloc[0])
    fake.update({"Ref ID": "9999", "Development Status": "Operational", "Operational": "", "Under Construction": ""})
    idx["9999"] = fake
    rid = "9999"
    row = _qc_row(State_data_plant_ID=rid, Project_life_cycle="Operational",
                  **{"Financial_close_date_status": "Estimated", "Financial_close_(estimated)": "06/30/2020",
                     "Operational_date_status": "Not reported, need more data to estimate",
                     "Project_investment_status": "Estimated", "Project_investment_($M)___Estimated": "50",
                     "Analyst_comments": "keep me"})
    out = updater.qc_process(dict(row), idx, [])
    for c in ("Financial_close_(estimated)", "Operational_date_status", "Project_investment_($M)___Estimated",
              "Analyst_comments"):
        assert out[c] == row[c]


def test_env_date_removed_when_future_or_eia_no():
    out = _run_qc(_qc_row(Environmental_approval_date="12/16/2026", EIA_approval_received="Yes"))
    assert out["Environmental_approval_date"] == ""
    from repd_qb import updater
    r = updater.finalize_row({"Sector": "Wind", "EIA_approval_received": "No",
                              "Environmental_approval_date": "05/07/2020"})
    assert r["Environmental_approval_date"] == ""


def test_reported_fc_after_reported_cod_is_removed_and_estimated():
    from repd_qb import updater
    import datetime as dt
    updater.PRESENT_DATE = dt.date(2026, 11, 16)
    repd = pd.read_csv(FX / "repd.csv", dtype=str, keep_default_na=False)
    idx = {r["Ref ID"]: r for _, r in repd.iterrows()}
    fake = dict(repd.iloc[0])
    fake.update({"Ref ID": "9998", "Development Status": "Operational", "Operational": "18/02/2016",
                 "Under Construction": "01/05/2016"})
    idx["9998"] = fake
    row = _qc_row(State_data_plant_ID="9998", Project_life_cycle="Operational", Sector="Solar")
    out = updater.finalize_row(updater.qc_process(dict(row), idx, []))
    assert out["Financial_close_date_(reported)"] == ""
    assert out["Financial_close_date_status"] == "Estimated"
    assert updater.parse_db_date(out["Financial_close_(estimated)"]) < dt.date(2016, 2, 18)


def test_awaiting_construction_stays_late_stage_with_dates_pushed_forward():
    from repd_qb import updater
    import datetime as dt
    updater.PRESENT_DATE = dt.date(2026, 11, 16)
    repd = pd.read_csv(FX / "repd.csv", dtype=str, keep_default_na=False)
    idx = {r["Ref ID"]: r for _, r in repd.iterrows()}
    fake = dict(repd.iloc[0])
    fake.update({"Ref ID": "9997", "Development Status": "Planning Permission Granted",
                 "Development Status (short)": "Awaiting Construction",
                 "Under Construction": "01/03/2025", "Operational": "01/06/2026"})
    idx["9997"] = fake
    row = _qc_row(State_data_plant_ID="9997", Project_life_cycle="Late-stage", Lead_QC="false",
                  EIA_approval_received="Yes",
                  **{"Financial_close_date_status": "Reported", "Financial_close_date_(reported)": "03/01/2025"})
    log = []
    out, _c, _t = updater.process_row(dict(row), idx, log)
    out = updater.finalize_row(out)
    assert out["Project_life_cycle"] == "Late-stage"
    fc = updater.parse_db_date(out["Financial_close_(estimated)"])
    op = updater.parse_db_date(out["Operational_date_(estimated)"])
    assert out["Financial_close_date_status"] == "Estimated" and fc > updater.PRESENT_DATE
    assert out["Operational_date_status"] == "Estimated" and op > fc
    assert out["Financial_close_date_(reported)"] == "" and out["Operational_date_(reported)"] == ""
