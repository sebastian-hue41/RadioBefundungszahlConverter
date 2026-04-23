# RadioBefundungszahlConverter

A CLI tool that generates a **Facharzt-Auswertung** xlsx report for Radiology by cross-referencing a per-physician examination statistics file against a procedure-to-section mapping file.

---

## What it does

Given two input files:

| File | Content |
|---|---|
| `statistics.xlsx` | Raw examination log exported from the RIS — one row per individual Leistung, grouped by patient visit (IND1/IND2 indices) |
| `map.xlsx` | Lookup table mapping every procedure code (`Leist0`) to one or more radiology sections (e.g. MRT, CT Hals/Thorax, US gesamt) |

The tool produces a single output xlsx with:
- **Rows** — radiology sections from the map (e.g. MRT, US gesamt, Mamma gesamt)
- **Columns** — one per year found in the data, plus a **Gesamt** total
- **Cells** — count of examinations that belong to that section in that year
- Procedures not found in the map are appended as extra rows so nothing is silently lost
- A footer section notes any count discrepancies, merged map duplicates, and warnings

---

## Architecture

```
main.py              CLI entry point (argparse)
input_handler.py     File security validation + data loading (statistics & map)
logic.py             Pure processing: matching, aggregation, count verification
output_handler.py    Styled xlsx generation
```

The processing layer (`logic.py`) has no I/O and no side effects — it works purely on Python dicts. This makes it straightforward to wrap with a REST API in a future web integration without changing any business logic.

---

## Prerequisites

```bash
pip install -r requirements.txt
```

Requires Python 3.10+ and `openpyxl`.

---

## Usage

```bash
python main.py <statistics.xlsx> <map.xlsx> [options]
```

### Options

| Flag | Default | Description |
|---|---|---|
| `--output DIR` | `.` (current directory) | Directory where the output file is saved |
| `--include-underscore` | off | Include underscore-prefixed section columns from the map (`_CLIP`, `_TRENNER`, etc.) |

### Examples

```bash
# Basic usage
python main.py 2024_Statistik_KUNZ_A.xlsx mapped_leistungen.xlsx

# Save output to a specific folder
python main.py 2024_Statistik_KUNZ_A.xlsx mapped_leistungen.xlsx --output ./reports

# Include internal map columns
python main.py 2024_Statistik_KUNZ_A.xlsx mapped_leistungen.xlsx --include-underscore
```

### Output filename

The output file is named automatically from the statistics file metadata:

```
{Mitarbeiter}{Befunddatum}Auswertung.xlsx
```

Example: `KUNZ_A01042019-31122022Auswertung.xlsx`

If the metadata cannot be read, a timestamped generic name is used instead and a note is added to the output file.

---

## Input file requirements

### statistics.xlsx

Exported from the RIS (e.g. ORBIS MitarbeiterStatistik). Expected structure:

- **G6** — label `Befunddatum`, **H6** — date range string
- **G7** — label `Mitarbeiter`, **H7** — employee code
- **L8** — total Befunddokumente count (used as informational reference)
- **Row 10** — column headers including `DokDatum`, `Leistung`, `Leist-Kurz`, `IND1`, `IND2`
- **Rows 11+** — one row per individual Leistung; rows belonging to the same patient visit share the same `IND1` value and have sequential `IND2` values (1, 2, 3 …); only the first row in each group carries a `DokDatum`

The tool counts total Leistungen using an IND2-based group-end algorithm (when the next IND2 value is ≤ the current, the current row is the last in its group and its IND2 value is the group size).

### map.xlsx (Leistungen Map)

A mapping table. Expected structure:

- **B1** — must be `Leist0` (procedure code column)
- **F1** — must be `DL Gefäße` (first section column)
- **Column B** — procedure short codes matching `Leist-Kurz` in the statistics file
- **Columns F onwards** — section names; a cell value of `1` means the procedure belongs to that section

Duplicate procedure codes are resolved automatically: if both entries have section assignments, values are merged by taking the maximum per column. A warning is written to the output noting which codes were affected.

---

## Error handling

The tool is designed to be robust by default:

| Situation | Behaviour |
|---|---|
| H6 / H7 / L8 missing | Generic output filename + note in output file; processing continues |
| `Leist-Kurz` column missing | Fatal error — file is not a valid statistics export |
| `B1 ≠ Leist0` | Fatal error — map file is invalid |
| `F1 ≠ DL Gefäße` | Structured error (recoverable — admin can supply a corrected map) |
| Procedure not found in map | Appended as its own row in output; processing continues |
| Duplicate map entries, conflicting | Merged by max-per-column; noted in output footer |
| Duplicate map entries, identical | Silently deduplicated |
| Formula injection in cells | Fatal error — file rejected as potentially malicious |
| VBA macros in file | Fatal error — file rejected |
| File > 50 MB | Fatal error |
| Rows with unparseable dates | Skipped with a per-row warning; noted in output |
| Count mismatch (processed ≠ expected) | Non-fatal; noted in output footer |

---

## Security

All input files are treated as potentially malicious:

- File extension is validated (`.xlsx` only)
- The zip container is inspected — corrupt or disguised files are rejected
- VBA macro presence causes immediate rejection
- File size is capped at 50 MB
- Every string cell value is checked for formula injection characters (`=`, `+`, `-`, `@`, `|`, …)

---

## Roadmap

- [ ] REST API wrapper (FastAPI) — logic layer is already server-compatible
- [ ] Admin endpoint to upload a corrected map file
- [ ] Multi-physician batch processing
- [ ] Web frontend for file upload and report download
