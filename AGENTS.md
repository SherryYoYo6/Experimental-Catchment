# AGENTS.md

## Project overview

**Experimental-Catchment** is a data-only hydrology research repository. It contains no application source code, build system, or test suite. The artifacts are CSV and XLSX datasets:

| Dataset | Files | Description |
|---|---|---|
| Penna Global Synthesis DB | `Penna_globalsynthesis_DB_Part1.csv`, `Penna_globalsynthesis_DB_Part2_update_2026_01.csv`, `Penna_globalsynthesis_DB_legend.txt` | Literature synthesis on runoff processes (~267 studies, ~1,194 catchment rows) |
| ECOI Data Availability Inventory | `Inventory_batch01.xlsx`–`Inventory_batch18.xlsx`, `Inventory_total.xlsx` | Watershed-level inventory of publicly available observational data (351 rows, 44 columns) |
| McMillan Process Database | `ProcessDatabase3-filtered.xlsx` | McMillan runoff-process figure/model metadata (351 rows) |
| McMillan ↔ Penna Match Table | `McMillan_Penna_Experimental_Match.xlsx` | Cross-dataset watershed matching (467 rows) |
| Merged Watershed View | `Inventory_total_merged_by_watershed.xlsx` | 243 deduplicated watersheds with audit log |

## Cursor Cloud specific instructions

### No services to start

This repository has no dev servers, databases, or background workers. There is nothing to `npm run dev`, `docker compose up`, or similar. "Running" the project means loading and inspecting the data files.

### Python tooling

Install dependencies from `requirements.txt` (already handled by the VM update script):

```bash
pip install -r requirements.txt
```

### CSV encoding

Penna CSV files use **latin-1** encoding (not UTF-8). Always pass `encoding="latin-1"` when reading them with pandas:

```python
df = pd.read_csv("Penna_globalsynthesis_DB_Part1.csv", encoding="latin-1")
```

### Quick data inspection

```bash
# Validate all files load correctly
python3 -c "
import glob, pandas as pd
for f in ['Penna_globalsynthesis_DB_Part1.csv','Penna_globalsynthesis_DB_Part2_update_2026_01.csv']:
    df = pd.read_csv(f, encoding='latin-1')
    print(f'{f}: {df.shape}')
for f in sorted(glob.glob('Inventory_batch*.xlsx')):
  df = pd.read_excel(f)
  print(f'{f}: {df.shape}')
"
```

### Lint / test / build

Not applicable — there is no source code, linter config, or test harness in this repository.
