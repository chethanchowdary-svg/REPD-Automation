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

- REPD capacity is applied (the old sheet only reported it).
- Lifecycle only moves **up** (a downgrade is logged, not applied).
- Investment fields are not re-estimated.
- Only the "Estimated Development Timeline" paragraph of the analyst comment is replaced.

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
