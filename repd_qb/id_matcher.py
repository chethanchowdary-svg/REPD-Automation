"""Match new REPD projects to DB UK projects that have no State_data_plant_ID (REPD Ref ID).

Priority (shortlist and GPT): location (Sub_location / State/Region / REPD county) 40,
sector (pre-filter), SPV / operator 25, project name 20, capacity 15.
Only "high" confidence is applied; one DB project takes one Ref ID.
"""
from __future__ import annotations

import csv
import json
import re
import time
from dataclasses import asdict, dataclass
from pathlib import Path

import pandas as pd
import requests
from rapidfuzz import fuzz

from . import config as CFG
from .updater import find_column

OPENAI_URL = "https://api.openai.com/v1/chat/completions"
STOP = set("""solar wind storage bess battery energy project projects farm park facility power
generation station ltd limited llp plc renewable renewables ess pv the and of co company hybrid
plant onshore offshore array scheme extension""".split())
ROMAN = {"i": 1, "ii": 2, "iii": 3, "iv": 4, "v": 5, "vi": 6, "vii": 7, "viii": 8, "ix": 9, "x": 10}
DIFFERENT_ASSET = re.compile(r"\b(repower\w*|extension|expansion|variation|phase\s*\w+|resubmission)\b", re.I)
W_LOC, W_SPV, W_NAME, W_CAP = 40, 25, 20, 15

SYSTEM_PROMPT = """You reconcile UK REPD (Renewable Energy Planning Database) entries with a renewable
energy project database. You get ONE REPD entry and up to five database projects that have no REPD
Ref ID yet. Decide whether exactly one candidate is the SAME physical project as the REPD entry.

Evaluate signals in this STRICT priority order:
1. Location (sub-location / county / region) -- STRONGEST signal. A clear location mismatch is almost
   always a disqualifier.
2. Sector -- must match (Solar / Wind / Battery / Hydro); already pre-filtered but confirm.
3. SPV / operator -- a name match here is strong supporting evidence.
4. Project name -- compare the distinctive part only, ignoring generic words such as Solar Farm,
   Wind Farm, BESS, Energy Storage, Park, Ltd, Limited.
5. Capacity (MW) -- supporting only; modest differences (< 20%) are acceptable.

NOT a match when the REPD entry is an extension, repower, variation, resubmission, or a different
phase/number (II vs I, Phase 2 vs Phase 1) of the candidate, or the sector differs.

Reply with JSON only:
{"match_project_id": "<Project ID of the candidate or null>",
 "confidence": "high" | "medium" | "low",
 "reason": "<one short sentence>"}
Use "high" only when location agrees AND the distinctive name or SPV clearly agrees AND nothing
suggests a different phase or asset."""


def norm_name(s: str) -> str:
    t = re.sub(r"[^a-z0-9 ]", " ", (s or "").lower())
    return " ".join(w for w in t.split() if w not in STOP)


def phase_numbers(s: str) -> set[int]:
    out = set()
    for w in re.sub(r"[^a-z0-9 ]", " ", (s or "").lower()).split():
        if w.isdigit() and int(w) < 20:
            out.add(int(w))
        elif w in ROMAN:
            out.add(ROMAN[w])
    return out


def _num(s):
    try:
        return float(re.sub(r"[^0-9.]", "", str(s)))
    except ValueError:
        return None


def _loc(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


@dataclass
class Decision:
    ref_id: str
    repd_name: str
    sector: str
    repd_mw: str
    repd_location: str
    project_id: str = ""
    db_name: str = ""
    db_mw: str = ""
    db_location: str = ""
    db_lifecycle: str = ""
    lead_qc: str = ""
    confidence: str = ""
    reason: str = ""
    method: str = ""
    candidates: int = 0
    applied: bool = False


class Matcher:
    def __init__(self, noid: pd.DataFrame, repd_idx: dict):
        cols = list(noid.columns)
        self.pid = find_column(cols, "Project ID", "Project_ID")
        self.qc = find_column(cols, "Lead QC")
        self.join = "State_data_plant_ID"
        d = noid[noid[self.join].astype(str).str.strip() == ""].copy() if self.join in noid else noid.copy()
        d["_n"] = d["Project_name"].map(norm_name)
        d["_s"] = d.get("SPV", "").map(norm_name) if "SPV" in d else ""
        d["_cap"] = d["Power_Capacity_(MW)"].map(_num)
        d["_locs"] = [{_loc(a), _loc(b)} - {""} for a, b in zip(d.get("Sub_location", ""), d.get("State/Region", ""))]
        self.noid, self.repd_idx = d, repd_idx

    def shortlist(self, new: dict, top=5, excluded=frozenset()):
        rr = self.repd_idx.get(new["State_data_plant_ID"], {})
        locs = {_loc(new.get("Sub_location")), _loc(new.get("State/Region")),
                _loc(rr.get("County")), _loc(rr.get("Region"))} - {""}
        rn, rs = norm_name(new.get("Project_name")), norm_name(new.get("SPV"))
        cap = _num(new.get("Power_Capacity_(MW)"))
        c = self.noid[(self.noid["Sector"] == new.get("Sector")) & (~self.noid[self.pid].isin(excluded))]
        out = []
        for i, d in c.iterrows():
            same_loc = bool(locs & d["_locs"]) or any(fuzz.ratio(a, b) >= 90 for a in locs for b in d["_locs"])
            ns = fuzz.token_set_ratio(rn, d["_n"]) if rn and d["_n"] else 0
            ss = max(fuzz.token_set_ratio(rs, d["_s"]) if rs and d["_s"] else 0,
                     fuzz.token_set_ratio(rs, d["_n"]) if rs and d["_n"] else 0)
            if not (same_loc or ns >= 85 or ss >= 90):
                continue
            capr = min(cap, d["_cap"]) / max(cap, d["_cap"]) if cap and d["_cap"] else 0
            score = (W_LOC if same_loc else 0) + W_SPV * ss / 100 + W_NAME * ns / 100 + W_CAP * capr
            if score >= 45:
                out.append((score, i))
        out.sort(reverse=True)
        return out[:top]

    def payload(self, new, cands):
        rr = self.repd_idx.get(new["State_data_plant_ID"], {})
        n = self.noid
        return {
            "repd_entry": {"ref_id": new["State_data_plant_ID"], "name": new.get("Project_name"),
                           "sector": new.get("Sector"), "MW": new.get("Power_Capacity_(MW)"),
                           "operator": new.get("SPV"), "sub_location": new.get("Sub_location"),
                           "state_region": new.get("State/Region"), "county": rr.get("County", ""),
                           "region": rr.get("Region", ""), "status": rr.get("Development Status", "")},
            "candidates": [{"project_id": n.at[i, self.pid], "name": n.at[i, "Project_name"],
                            "MW": n.at[i, "Power_Capacity_(MW)"], "sub_location": n.at[i, "Sub_location"],
                            "state_region": n.at[i, "State/Region"], "spv": n.at[i, "SPV"] if "SPV" in n else "",
                            "lifecycle": n.at[i, "Project_life_cycle"] if "Project_life_cycle" in n else ""}
                           for _, i in cands],
        }


def ask_gpt(payload: dict, api_key: str, model: str, retries: int = 4) -> dict:
    body = {"model": model, "response_format": {"type": "json_object"},
            "messages": [{"role": "system", "content": SYSTEM_PROMPT},
                         {"role": "user", "content": json.dumps(payload, default=str)}]}
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    for attempt in range(retries):
        r = requests.post(OPENAI_URL, headers=headers, json=body, timeout=120)
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(2 ** attempt * 3)
            continue
        if r.status_code != 200:
            raise RuntimeError(f"OpenAI API {r.status_code}: {r.text[:300]}")
        content = r.json()["choices"][0]["message"]["content"]
        return json.loads(re.sub(r"^```(json)?|```$", "", content.strip()).strip())
    raise RuntimeError("OpenAI API kept failing after retries")


def load_overrides(path: Path) -> dict[str, list[tuple[str, str]]]:
    out: dict[str, list[tuple[str, str]]] = {}
    if not path or not Path(path).exists():
        return out
    with open(path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            rid, dec = (r.get("Ref ID") or "").strip(), (r.get("Decision") or "").strip().lower()
            if rid and dec in ("match", "reject"):
                out.setdefault(rid, []).append(((r.get("Project ID") or "").strip(), dec))
    return out


def _guard(repd_name: str, db_name: str) -> str:
    if DIFFERENT_ASSET.search(repd_name or ""):
        return "REPD name suggests extension/repower/variation/phase"
    a, b = phase_numbers(repd_name), phase_numbers(db_name)
    return "phase/number differs" if a and b and a != b else ""


def match(new_rows: list[dict], noid: pd.DataFrame, repd_idx: dict, overrides_path: Path | None = None,
          api_key: str | None = None, model: str | None = None, max_calls: int | None = None, ask=None):
    """Return ({Ref ID: noid row index}, decisions, stats)."""
    ask = ask or ask_gpt
    api_key = CFG.OPENAI_API_KEY if api_key is None else api_key
    model = model or CFG.OPENAI_MODEL
    max_calls = max_calls or CFG.GPT_MAX_CALLS
    m = Matcher(noid, repd_idx)
    ov = load_overrides(overrides_path)
    stats = {"DB projects without Ref ID considered": len(m.noid), "GPT calls": 0}
    proposals, decisions = [], []

    for new in new_rows:
        rid = new["State_data_plant_ID"]
        dec = Decision(rid, new.get("Project_name", ""), new.get("Sector", ""), new.get("Power_Capacity_(MW)", ""),
                       " / ".join(x for x in (new.get("Sub_location"), new.get("State/Region")) if x))
        forced = [p for p, d in ov.get(rid, []) if d == "match" and p]
        rejected = {p for p, d in ov.get(rid, []) if d == "reject"}
        if "" in rejected:
            continue
        if forced:
            hit = m.noid.index[m.noid[m.pid] == forced[0]]
            if len(hit):
                dec.method, dec.confidence, dec.reason = "manual", "high", "Manual override"
                _fill(dec, m, hit[0])
                proposals.append((1000.0, dec, hit[0]))
                decisions.append(dec)
            continue
        cands = m.shortlist(new, excluded=rejected)
        if not cands:
            continue
        dec.candidates = len(cands)
        if not api_key or stats["GPT calls"] >= max_calls:
            dec.method = "none"
            dec.reason = "OPENAI_API_KEY not set" if not api_key else f"GPT call limit ({max_calls}) reached"
            _fill(dec, m, cands[0][1])
            decisions.append(dec)
            continue
        try:
            ans = ask(m.payload(new, cands), api_key, model)
            stats["GPT calls"] += 1
        except Exception as e:
            dec.method, dec.reason = "none", f"GPT error: {e}"[:300]
            decisions.append(dec)
            continue
        dec.method = "gpt"
        pid = str(ans.get("match_project_id") or "").strip()
        dec.confidence = str(ans.get("confidence") or "low").lower()
        dec.reason = str(ans.get("reason") or "")
        valid = {m.noid.at[i, m.pid]: i for _, i in cands}
        if pid in valid:
            i = valid[pid]
            _fill(dec, m, i)
            g = _guard(dec.repd_name, dec.db_name)
            if g and dec.confidence == "high":
                dec.confidence, dec.reason = "medium", dec.reason + f" [capped: {g}]"
            proposals.append((next(s for s, j in cands if j == i), dec, i))
        decisions.append(dec)

    applied, taken = {}, set()
    for score, dec, i in sorted(proposals, key=lambda x: -x[0]):
        if dec.confidence != "high":
            continue
        if i in taken:
            dec.reason += " [not applied: DB project already matched]"
            continue
        taken.add(i)
        applied[dec.ref_id] = i
        dec.applied = True
    stats.update({"Matches applied (moved to existing)": len(applied),
                  "Possible matches for review (medium)": sum(1 for d in decisions if d.project_id
                                                             and not d.applied and d.confidence == "medium"),
                  "GPT model": model if api_key else "not used (no OPENAI_API_KEY)"})
    return applied, decisions, stats, m


def _fill(dec: Decision, m: Matcher, i):
    n = m.noid
    dec.project_id = n.at[i, m.pid]
    dec.db_name = n.at[i, "Project_name"]
    dec.db_mw = n.at[i, "Power_Capacity_(MW)"]
    dec.db_location = " / ".join(x for x in (n.at[i, "Sub_location"], n.at[i, "State/Region"]) if x)
    dec.db_lifecycle = n.at[i, "Project_life_cycle"] if "Project_life_cycle" in n else ""
    dec.lead_qc = n.at[i, m.qc] if m.qc else ""


def decisions_table(decisions) -> list[dict]:
    names = {"ref_id": "REPD Ref ID", "repd_name": "REPD name", "sector": "Sector", "repd_mw": "REPD MW",
             "repd_location": "REPD location", "project_id": "Matched Project ID", "db_name": "DB project name",
             "db_mw": "DB MW", "db_location": "DB location", "db_lifecycle": "DB lifecycle",
             "lead_qc": "Lead QC", "confidence": "Confidence", "reason": "Reason", "method": "Method",
             "candidates": "Candidates", "applied": "Applied"}
    rows = []
    for d in sorted(decisions, key=lambda d: (not d.applied, d.confidence != "medium", d.ref_id)):
        x = asdict(d)
        x["applied"] = "Yes" if d.applied else "No"
        rows.append({names[k]: v for k, v in x.items()})
    return rows
