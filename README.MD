# RadioBefundungszahlConverter

Generates a **Facharzt-Auswertung** xlsx report for Radiology by cross-referencing a per-physician examination statistics file against a procedure-to-section mapping file, with optional colour-coded target comparison.

Available as a **CLI tool** (master branch) and as an **iii web worker** (this branch).

---

## Table of contents

1. [What it does](#what-it-does)
2. [Package structure](#package-structure)
3. [CLI usage](#cli-usage)
4. [Web integration (iii)](#web-integration-iii)
   - [What is iii?](#what-is-iii)
   - [Prerequisites](#prerequisites-web)
   - [Running the worker](#running-the-worker)
   - [HTTP API](#http-api)
   - [Calling the API](#calling-the-api)
5. [Input file requirements](#input-file-requirements)
6. [Reference amounts and colour coding](#reference-amounts-and-colour-coding)
7. [Error handling](#error-handling)
8. [Security](#security)

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

The output xlsx is built **entirely in memory** — no files are written to disk on the server.

---

## Package structure

```
core/
  input_handler.py   Security validation + data loading (statistics, map, reference)
  logic.py           Pure processing: matching, aggregation, count verification
  output_handler.py  Styled xlsx generation (in-memory and disk variants)

cli/
  main.py            CLI entry point (argparse) — uses core/, writes to disk

web/
  worker.py          iii HTTP worker — uses core/, streams xlsx back to client

requirements.txt      CLI dependencies (openpyxl)
requirements-web.txt  Web worker dependencies (openpyxl + iii-sdk)
```

The `core/` layer has no I/O and no side effects. Both the CLI and the web worker call the same functions — the only difference is how files arrive (paths vs. in-memory bytes) and how the result leaves (saved to disk vs. streamed over HTTP).

---

## CLI usage

```bash
pip install -r requirements.txt
python -m cli.main <statistics.xlsx> <map.xlsx> [options]
```

### Options

| Flag | Default | Description |
|---|---|---|
| `--output DIR` | `.` (current directory) | Directory where the output file is saved |
| `--include-underscore` | off | Include `_`-prefixed section columns from the map |
| `--reference FILE` | `reference_amount.xlsx` (auto-detected in CWD) | Path to the reference file with required amounts per section |

### Examples

```bash
# Basic — auto-detects reference_amount.xlsx in the current directory
python -m cli.main 2024_Statistik_KUNZ_A.xlsx mapped_leistungen.xlsx

# Save to a specific folder
python -m cli.main 2024_Statistik_KUNZ_A.xlsx mapped_leistungen.xlsx --output ./reports

# Specify a custom reference file
python -m cli.main 2024_Statistik_KUNZ_A.xlsx mapped_leistungen.xlsx --reference targets_2024.xlsx
```

---

## Web integration (iii)

### What is iii?

[iii](https://github.com/iii-hq) is a lightweight backend infrastructure platform. You register Python functions as **workers** that connect to a central **iii engine** over WebSocket. The engine exposes the workers as HTTP endpoints and handles routing, load-balancing, and streaming — the worker code never opens a port itself.

```
Browser / API client
        │  POST /convert (multipart)
        ▼
  iii engine  ──WebSocket──►  web/worker.py
        │                          │
        │◄──── xlsx bytes ─────────┘
        ▼
  HTTP response (application/vnd.openxmlformats...)
```

The worker receives the raw HTTP request (headers + body) as structured objects and writes the response back through the same WebSocket connection. The engine handles TLS, keep-alive, and buffering.

### Prerequisites (web)

```bash
pip install -r requirements-web.txt
```

You also need a running iii engine. Follow the [iii setup guide](https://github.com/iii-hq) to start one, or point `III_ENGINE_URL` at an existing instance.

### Running the worker

```bash
# From the project root:
python -m web.worker
```

Environment variables:

| Variable | Default | Description |
|---|---|---|
| `III_ENGINE_URL` | `ws://localhost:49134` | WebSocket URL of the iii engine |
| `WORKER_NAME` | `befundungszahl-converter` | Display name shown in the iii console |

The worker registers itself with the engine on startup and logs a ready message:

```
INFO web.worker: Connecting to iii engine at ws://localhost:49134 ...
INFO web.worker: Worker 'befundungszahl-converter' ready — POST /convert
```

Once the worker is registered, the engine exposes the endpoint at:

```
POST http://<engine-host>/convert
```

### HTTP API

#### `POST /convert`

Upload the statistics and map files as `multipart/form-data`. The server processes them in memory and returns the output xlsx as a binary download — nothing is written to disk.

**Request fields:**

| Field | Type | Required | Description |
|---|---|---|---|
| `statistics` | file (`.xlsx`) | Yes | MitarbeiterStatistik export from the RIS |
| `map` | file (`.xlsx`) | Yes | Leistungen map file |
| `reference` | file (`.xlsx`) | No | Required amounts per section; adds colour-coded Benötigt column |

**Query parameters:**

| Parameter | Default | Description |
|---|---|---|
| `include_underscore` | `false` | Set to `true` to include `_`-prefixed section columns from the map |

**Responses:**

| Status | Content-Type | Description |
|---|---|---|
| `200` | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` | Output xlsx file as a binary stream |
| `400` | `application/json` | Bad request — missing required field or wrong content type |
| `413` | `application/json` | Upload too large (> 50 MB per file, > 110 MB total) |
| `422` | `application/json` | Validation failure — file is malformed, malicious, or structurally invalid |
| `500` | `application/json` | Unexpected server error |
| `503` | `application/json` | Processing timed out (> 60 s) |

All responses include `X-Content-Type-Options: nosniff` and `Cache-Control: no-store`.

Error responses have the shape:
```json
{ "error": "human-readable description" }
```

### Calling the API

#### curl

```bash
# Minimal — statistics + map only
curl -X POST http://localhost:<engine-port>/convert \
  -F "statistics=@2024_Statistik_KUNZ_A.xlsx" \
  -F "map=@mapped_leistungen.xlsx" \
  --output Auswertung.xlsx

# With reference amounts (Benötigt column + colour coding)
curl -X POST http://localhost:<engine-port>/convert \
  -F "statistics=@2024_Statistik_KUNZ_A.xlsx" \
  -F "map=@mapped_leistungen.xlsx" \
  -F "reference=@reference_amount.xlsx" \
  --output Auswertung.xlsx

# Include underscore-prefixed map columns
curl -X POST "http://localhost:<engine-port>/convert?include_underscore=true" \
  -F "statistics=@2024_Statistik_KUNZ_A.xlsx" \
  -F "map=@mapped_leistungen.xlsx" \
  --output Auswertung.xlsx
```

#### Python (requests)

```python
import requests

with (
    open("2024_Statistik_KUNZ_A.xlsx", "rb") as stats,
    open("mapped_leistungen.xlsx", "rb") as map_file,
    open("reference_amount.xlsx", "rb") as ref,
):
    response = requests.post(
        "http://localhost:<engine-port>/convert",
        files={
            "statistics": ("statistics.xlsx", stats, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "map":        ("map.xlsx",        map_file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "reference":  ("reference.xlsx",  ref,      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        },
    )
    response.raise_for_status()
    with open("Auswertung.xlsx", "wb") as out:
        out.write(response.content)
```

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

---

## Reference amounts and colour coding

### reference_amount.xlsx format

A two-column file with required examination counts per section:

| Column A | Column B |
|---|---|
| Leistungsbereich | Benötigt |
| MRT | 3000 |
| CT Skelett | 4000 |
| CT Hals/Thorax | 4000 |
| … | … |

- Sections with an empty **Benötigt** cell are silently skipped
- The **Benötigt** column only appears in the output when reference data is present
- If the file (or upload field) is absent, the column is omitted entirely

### Colour coding

When reference data is present, the **Gesamt** cell for each section is colour-coded using Excel **conditional formatting** — the colour updates live when the user edits a Benötigt value directly in the output file, without needing to regenerate it.

| Colour | Meaning |
|---|---|
| Green | ≥ 100 % of required — target met |
| Yellow | ≥ 80 % — close to target |
| Orange | ≥ 50 % — halfway |
| Red | < 50 % — well below target |

---

## Error handling

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
| File > 50 MB | Fatal error (CLI) / 413 response (web) |
| Rows with unparseable dates | Skipped with a per-row warning; noted in output |
| Count mismatch (processed ≠ expected) | Non-fatal; noted in output footer |
| Reference file not found / not uploaded | Benötigt column omitted; processing continues normally |
| Reference file fails security check | Logged as warning; Benötigt column omitted |
| Individual reference row unparseable | That row skipped; rest of reference data used |
| Section has no reference entry | Gesamt cell uncoloured; Benötigt cell left blank |
| Request body > 110 MB | 413 before file is read into memory |
| Processing timeout (> 60 s) | 503 response; worker continues serving other requests |

---

## Security

All input files — including the optional reference file — are treated as potentially malicious:

- File extension is validated (`.xlsx` only)
- The zip container is inspected — corrupt or disguised files are rejected
- Zip path-traversal entries (e.g. `../../etc/passwd`) are rejected
- Zip bombs are detected: per-entry compression ratio > 50:1 or total uncompressed size > 200 MB
- VBA macro presence causes immediate rejection
- File size is capped at 50 MB per file; the HTTP body is capped at 110 MB total
- Every string cell value is checked for formula injection characters (`=`, `+`, `-`, `@`, `|`, …)
- Reference amounts are bounds-checked (must be positive integers ≤ 1 000 000)
- Map files are limited to 500 section columns to prevent memory exhaustion
- All HTTP responses include `X-Content-Type-Options: nosniff` and `Cache-Control: no-store`
- Processing is time-limited to 60 seconds per request
