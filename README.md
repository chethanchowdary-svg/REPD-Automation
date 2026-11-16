# UK REPD to Quickbase (quarterly)

Downloads the latest REPD quarterly CSV from GOV.UK, pulls the UK projects from Quickbase, and builds
one workbook of Quickbase upload sheets. The rules are the ones from the Colab scripts
(`REPD_DB_update` and `REPD_new_projects`).

| Tab | What it is |
|---|---|
| Lifecycle Capacity / Site Area Changes / Other Changes | Non-QC upload sheets, changed rows only, changed cells light yellow |
| Lead QC Update | Upload sheet for Lead QC rows (replaces the old comment sheet) |
| New Projects | REPD Ref IDs not in the DB (Project_ID blank) |
| ID Matches (GPT) | New REPD projects checked against DB projects that have no State_data_plant_ID |
| Full DB / Change Log / Run Summary | Audit tabs |

## Setup (once)

1. Create a **private** repo and push this folder.
2. Settings > Secrets and variables > Actions > **Secrets**:
   `QB_USER_TOKEN`, `OPENCAGE_API_KEY`, `OPENAI_API_KEY`.
3. **Variables** (defaults built in): `QB_REPORT_ID` = 1000638, `OPENAI_MODEL` = gpt-4.1-mini.

## The one Quickbase report (1000638)

Filter: Country = United Kingdom, all rows (with and without State data plant ID, Lead QC included).
Rows with a Ref ID are updated; all Ref IDs decide what is new; rows without a Ref ID are the GPT
matching candidates. Columns (exact labels): Project ID, State data plant ID, Lead QC, Project name,
Sector, Offshore/Onshore, Power Capacity (MW), Project life cycle, State/Region, Sub-location,
Financial close date status, Financial close (estimated), Financial close date (reported),
Operational date status, Operational date (estimated), Operational date (reported),
Project investment status, Project investment ($M) - Estimated, Project investment ($M) - Reported,
EIA approval received, Grid connection approved, Financing secured, Financing type, PPA signed,
Power sales arrangement, Revenue mechanism, UK CFD Allocation Round (AR), Environmental approval date,
Site area (acres), Co-located storage, SPV, Show SPV, Analyst comments, Show analyst comment.
4. Actions > "REPD quarterly Quickbase update" > Run workflow. Output is the run's artifact (90 days).

Schedule: 06:00 UTC on 10 Jan / Apr / Jul / Oct. Edit the `cron` line to change it.

## Lead QC rows

- Capacity is **never changed**; a REPD difference is listed in QC Flags / Change Log only.
- Lifecycle only moves **up**. On a REPD downgrade the DB lifecycle is kept, but every other rule
  still runs (dates, flags, formats, comment).
- A reported investment, or an estimate inside the house-rule $/MW range, is kept; an estimate outside
  the range (or "Not reported" on a live project) is re-estimated.
- Only the "Estimated Project Investment" and "Estimated Development Timeline" paragraphs of the analyst
  comment are rebuilt (or removed when nothing is estimated); other analyst paragraphs are kept.

## House rules applied (Projects-Entry-Bot house_rules.md + SOP.md)

| Area | Rule |
|---|---|
| Dates (HR6) | Early FC = today + MW/50x8 (>= 6 mo); Late FC = today + MW/50x6 (>= 3 mo); In-construction FC = today - MW/50x6 (not after today); COD = FC + MW/50x6 (wind >= 6, others >= 3, <= 48 mo); operational FC = COD - MW/50x6 (<= 36 mo); estimates at month-end; reported dates are never overwritten by the gap checks |
| Lifecycle (HR10, SOP 6) | reported FC passed on a pre-construction row -> In construction; Late-stage with EIA = No and Grid = No -> Early-stage; Late / Ready-to-build approvals, financing and PPA are kept (researched), defaults only where the SOP fixes them |
| Investment (HR7) | nearest $1M (BESS 1 decimal); comment metric = value / MW |
| Comments (HR8) | template cases 2-5, "onshore/offshore wind", Show analyst comment = Yes only when something is estimated |
| Battery (HR2, HR20) | name gets " BESS"; Power sales arrangement = Not Applicable; Revenue mechanism defaults to Yet to be determined/Unavailable (New Projects now has the column) |
| Commercial (SOP 7) | blank PSA -> Yet to be determined/Unavailable (no longer cleared when PPA = No); CfD round adds CfD to the PSA list; Financing type Debt + Equity from construction on, other instruments kept |
| Formats (HR14) | Yes/No instead of True/False; multi-selects as `A;B`; mojibake repaired, accents -> plain letters |

**QC Flags tab**: what the rules cannot decide (duplicate REPD IDs / names, CfD capacity > project
capacity, broken characters, operational without a reported COD, COD > 4 years after a reported FC,
REPD capacity differences on Lead QC rows, ...).

## GPT matching (new REPD projects vs DB projects without a Ref ID)

1. Shortlist up to 5 DB projects locally: same sector, and same location (Sub_location / State/Region /
   REPD county) or a strong name / SPV match. Priority: location 40, SPV 25, name 20, MW 15.
2. GPT picks one or none. Only **high** confidence is applied; extension / repower / variation /
   phase wording or a different phase number caps it at medium. One DB project takes one Ref ID.
3. An applied match leaves New Projects: the DB row gets the Ref ID and runs through the normal rules.
4. `data/id_match_overrides.csv` (`Ref ID,Project ID,Decision` with `match` / `reject`) fixes decisions.

## Run locally

```bash
pip install -r requirements.txt
set QB_USER_TOKEN / OPENCAGE_API_KEY / OPENAI_API_KEY as environment variables
python -m repd_qb.main
python -m repd_qb.main --repd-csv repd.csv --db-csv db.csv   # offline
```
