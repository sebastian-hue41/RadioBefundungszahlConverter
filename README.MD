# RadioBefundungszahlConverter

A CLI tool that generates a **Facharzt-Auswertung** xlsx report for Radiology by cross-referencing a per-physician examination statistics file against a procedure-to-section mapping file, with optional colour-coded target comparison.

---

## What it does

Given two required input files (and one optional reference file):

| File | Required | Content |
|---|---|---|
| `statistics.xlsx` | Yes | Raw examination log exported from the RIS — one row per individual Leistung, grouped by patient visit (IND1/IND2 indices) |
| `map.xlsx` | Yes | Lookup table mapping every procedure code (`Leist0`) to one or more radiology sections (e.g. MRT, CT Hals/Thorax, US gesamt) |
| `reference_amount.xlsx` | No | Required examination counts per section — adds a colour-coded **Benötigt** column to the output |

The tool produces a single output xlsx with:
- **Rows** — radiology sections from the map (e.g. MRT, US gesamt, Mamma gesamt)
- **Columns** — one per year found in the data, plus a **Gesamt** total and (if reference data is present) a **Benötigt** target column
- **Cells** — count of examinations that belong to that section in that year
- **Colour coding** on the Gesamt cell showing how close the physician is to reaching the required amount
- Procedures not found in the map are appended as extra rows so nothing is silently lost
- A footer section notes any count discrepancies, merged map duplicates, and warnings

---

## How it works

### The two input files

**statistics.xlsx** is the raw export from your RIS. Each row is one billed procedure (Leistung). A single patient visit can produce multiple rows.

**map.xlsx** is a lookup table you maintain. It maps every procedure short code to one or more radiology sections and specifies how many times it counts there.

### The four key columns

| Column | Role |
|---|---|
| **Leist-Kurz** | Short procedure code — the key used to look up the section assignment in the map |
| **DokDatum** | Date of the examination visit — used to assign each procedure to a year |
| **IND1** | Patient/visit identifier — one value per patient visit |
| **IND2** | Position of this procedure within the visit (1, 2, 3 …) |

### Date carry-forward

The RIS only writes `DokDatum` on the **first row of each visit** (IND2 = 1). All subsequent rows in the same visit leave it blank. The loader carries the last seen date forward automatically, so every procedure from the same visit is counted in the same year.

### How IND2 computes the expected total

Rather than counting rows (which can be skipped for bad dates or missing codes), the tool derives an independent expected total from the IND2 sequence alone.

Walk the IND2 values in order. Whenever the next value is **≤ the current value** — or the list ends — the current row is the last in its group, and its IND2 value is exactly the group's size. Sum those values.

```
IND2 sequence:  1  2  3 │ 1  2 │ 1
                         ↑     ↑     ← group boundaries (next ≤ current or end)
Group sizes:       3       2     1   →  total = 6
```

After loading, the number of successfully processed rows is compared against this total. Any discrepancy means rows were silently dropped (bad date, missing Leist-Kurz) and is flagged in the output footer.

### Map lookup and section counting

For each loaded examination, the tool looks up its `Leist-Kurz` in the map:

- **Found** — for every section in the map entry where the weight `val ≥ 1`, add `val` to that section's counter for that year. A weight of `2` means one procedure counts twice — used for multi-region scans that are billed separately for each body region.
- **Not found** — the procedure is appended as its own row using its human-readable `Leistung` name, so nothing is silently lost.

### End-to-end summary

```
statistics.xlsx
  ├─ IND2 sequence  ──────────────────────────→  expected total (cross-check)
  ├─ DokDatum  ──→  carry-forward  ──────────→  year per row
  └─ Leist-Kurz  ─────────────────────┐
                                       ↓
map.xlsx:  Leist-Kurz → {section: weight, …}
                                       ↓
                        counts[section][year] += weight
                                       ↓
                        output.xlsx: section rows × year columns
```

---

## Architecture

```
main.py              CLI entry point (argparse)
input_handler.py     File security validation + data loading (statistics, map, reference)
logic.py             Pure processing: matching, aggregation, count verification
output_handler.py    Styled xlsx generation
```

The processing layer (`logic.py`) has no I/O and no side effects — it works purely on Python dicts.

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
| `--reference FILE` | `reference_amount.xlsx` (auto-detected in CWD) | Path to the reference file with required amounts per section |

### Examples

```bash
# Basic usage — auto-detects reference_amount.xlsx in the current directory
python main.py 2024_Statistik_DR_EXAMPLE.xlsx mapped_leistungen.xlsx

# Save output to a specific folder
python main.py 2024_Statistik_DR_EXAMPLE.xlsx mapped_leistungen.xlsx --output ./reports

# Specify a custom reference file
python main.py 2024_Statistik_DR_EXAMPLE.xlsx mapped_leistungen.xlsx --reference targets_2024.xlsx

# Include internal map columns
python main.py 2024_Statistik_DR_EXAMPLE.xlsx mapped_leistungen.xlsx --include-underscore
```

### Output filename

The output file is named automatically from the statistics file metadata:

```
{Mitarbeiter}{Befunddatum}Auswertung.xlsx
```

Example: `DR_EXAMPLE01042019-31122022Auswertung.xlsx`

If the metadata cannot be read, a timestamped generic name is used and a note is added to the output file.

---

## Input file requirements

### statistics.xlsx

Exported from the RIS (e.g. ORBIS MitarbeiterStatistik). Expected structure:

- **G6** — label `Befunddatum`, **H6** — date range string
- **G7** — label `Mitarbeiter`, **H7** — employee code
- **L8** — total Befunddokumente count (displayed for reference only)
- **Row 10** — column headers including `DokDatum`, `Leistung`, `Leist-Kurz`, `IND1`, `IND2`
- **Rows 11+** — one row per individual Leistung; rows belonging to the same patient visit share the same `IND1` value and have sequential `IND2` values (1, 2, 3 …); only the first row in each group carries a `DokDatum`

The tool counts total Leistungen using an IND2-based group-end algorithm: when the next IND2 value is ≤ the current, the current row ends its group and its IND2 value is the group size.

### map.xlsx (Leistungen Map)

A mapping table. Expected structure:

- **B1** — must be `Leist0` (procedure code column)
- **F1** — must be `DL Gefäße` (first section column)
- **Column B** — procedure short codes matching `Leist-Kurz` in the statistics file
- **Columns F onwards** — section names; a cell value of `1` means the procedure belongs to that section

Duplicate procedure codes are resolved automatically: if both entries have section assignments, values are merged by taking the maximum per column. A warning is written to the output noting which codes were affected.

### reference_amount.xlsx (optional)

A two-column file with required examination counts per section:

| Column A | Column B |
|---|---|
| Leistungsbereich | Benötigt |
| MRT | 3000 |
| CT Skelett | 4000 |
| CT Hals/Thorax | 4000 |
| … | … |

- Sections with an empty **Benötigt** cell are silently skipped
- The file is auto-detected as `reference_amount.xlsx` in the current working directory; override with `--reference`
- If the file is absent the **Benötigt** column is omitted from the output entirely

#### Colour coding

When reference data is present, the **Gesamt** cell for each section is colour-coded using Excel **conditional formatting** — the colour updates live when you edit a Benötigt value directly in the output file, without needing to regenerate it.

| Colour | Meaning |
|---|---|
| Green | ≥ 100 % of required — target met |
| Yellow | ≥ 80 % — close to target |
| Orange | ≥ 50 % — halfway |
| Red | < 50 % — well below target |

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
| `reference_amount.xlsx` not found | Benötigt column omitted; processing continues normally |
| Reference file fails security check | Logged as warning; Benötigt column omitted |
| Individual reference row unparseable | That row skipped; rest of reference data used |
| Section has no reference entry | Gesamt cell uncoloured; Benötigt cell left blank |

---

## Security

All input files — including the optional reference file — are treated as potentially malicious:

- File extension is validated (`.xlsx` only)
- The zip container is inspected — corrupt or disguised files are rejected
- VBA macro presence causes immediate rejection
- File size is capped at 50 MB
- Every string cell value is checked for formula injection characters (`=`, `+`, `-`, `@`, `|`, …)
- Reference amounts are bounds-checked (must be positive integers ≤ 1 000 000)
