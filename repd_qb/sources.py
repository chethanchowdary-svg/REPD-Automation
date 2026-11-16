"""Quickbase report fetch (paginated) and REPD quarterly CSV download."""
from __future__ import annotations

import io

import pandas as pd
import requests
from bs4 import BeautifulSoup

from . import config as C

REPD_GOV_PAGE = ("https://www.gov.uk/government/publications/"
                 "renewable-energy-planning-database-quarterly-extract")


def fetch_quickbase(report_id: str, table_id: str = None, page: int = 1000) -> pd.DataFrame:
    """Run a Quickbase report and return all rows (values stringified like the original scripts)."""
    if not C.QB_USER_TOKEN:
        raise RuntimeError("QB_USER_TOKEN is empty. Add it as a GitHub secret.")
    headers = {"QB-Realm-Hostname": C.QB_REALM,
               "Authorization": f"QB-USER-TOKEN {C.QB_USER_TOKEN}",
               "User-Agent": "enerdatics-repd-automation",
               "Content-Type": "application/json"}
    rows, labels, skip, total = [], None, 0, None
    print(f"  Running Quickbase report {report_id} ...")
    while total is None or skip < total:
        r = requests.post(f"https://api.quickbase.com/v1/reports/{report_id}/run",
                          headers=headers, params={"tableId": table_id or C.QB_TABLE_ID,
                                                   "skip": skip, "top": page},
                          json={}, timeout=180)
        if not r.ok:
            raise RuntimeError(f"QB report {report_id} failed: {r.status_code} {r.text[:300]}")
        js = r.json()
        labels = labels or {str(f["id"]): f.get("label", str(f["id"])) for f in js.get("fields", [])}
        for rec in js.get("data", []):
            row = {}
            for fid, cell in rec.items():
                val = cell.get("value", "") if isinstance(cell, dict) else cell
                row[labels.get(str(fid), fid)] = "" if val is None else str(val)
            rows.append(row)
        meta = js.get("metadata", {})
        got = meta.get("numRecords", len(js.get("data", [])))
        total = meta.get("totalRecords", len(rows))
        if not got:
            break
        skip += got
    df = pd.DataFrame(rows, columns=list(labels.values()) if labels else None).fillna("")
    print(f"  Fetched {len(df):,} rows, {len(df.columns)} columns.")
    return df


def fetch_repd_csv(gov_page_url: str = REPD_GOV_PAGE) -> tuple[str, pd.DataFrame]:
    print(f"  Fetching REPD page: {gov_page_url}")
    page = requests.get(gov_page_url, timeout=30)
    if not page.ok:
        raise RuntimeError(f"Could not fetch GOV.UK page: {page.status_code}")
    soup = BeautifulSoup(page.text, "html.parser")
    links = [a["href"] for a in soup.find_all("a", href=True)
             if "assets.publishing.service.gov.uk" in a["href"] and a["href"].lower().endswith(".csv")]
    if not links:
        raise RuntimeError("No CSV link found on GOV.UK REPD page.")
    url, name = links[0], links[0].rsplit("/", 1)[-1]
    print(f"  Latest REPD: {name}")
    dl = requests.get(url, timeout=180)
    if not dl.ok:
        raise RuntimeError(f"REPD download failed: {dl.status_code}")
    return name, read_csv_bytes(dl.content)


def decode_mixed(raw: bytes) -> str:
    """UTF-8 per line, cp1252 only for the lines that are not valid UTF-8. Decoding a whole file as
    cp1252 because of one bad byte would garble every accented letter and dash in it."""
    if raw.startswith(b"\xef\xbb\xbf"):
        raw = raw[3:]
    out = []
    for line in raw.split(b"\n"):
        try:
            out.append(line.decode("utf-8"))
        except UnicodeDecodeError:
            out.append(line.decode("cp1252", errors="replace"))
    return "\n".join(out)


def read_csv_bytes(raw: bytes) -> pd.DataFrame:
    return pd.read_csv(io.StringIO(decode_mixed(raw)), dtype=str, keep_default_na=False)


def read_csv_path(path) -> pd.DataFrame:
    return read_csv_bytes(open(path, "rb").read())
