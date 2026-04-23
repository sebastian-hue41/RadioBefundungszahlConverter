"""
input_handler.py — file validation and loading for statistics.xlsx and map.xlsx.

All inputs are treated as potentially malicious. Validation is strict.
Designed for server-side use: no interactive prompts, structured errors only.
"""

import re
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl

# ── Constants ─────────────────────────────────────────────────────────────────

MAX_FILE_SIZE_MB = 50
MAX_DATA_ROWS = 500_000
ALLOWED_EXTENSIONS = {".xlsx"}

# Cells/values starting with these chars are formula injection attempts.
_FORMULA_PREFIXES = ("=", "+", "-", "@", "|", "\t", "\r")


# ── Custom exceptions ──────────────────────────────────────────────────────────

class FileValidationError(ValueError):
    """Raised when a file fails security or structural validation."""


class StatsValidationError(ValueError):
    """Raised when statistics.xlsx fails structural validation."""


class MapValidationError(ValueError):
    """
    Raised when map.xlsx fails structural validation.
    The `recoverable` flag signals that an admin supplying a corrected
    map file could resolve the issue (used for future API handling).
    """
    def __init__(self, message: str, recoverable: bool = False):
        super().__init__(message)
        self.recoverable = recoverable


# ── Shared security validation ─────────────────────────────────────────────────

def validate_xlsx_file(path: str) -> Path:
    """
    Validates that `path` points to a safe, well-formed xlsx file.
    Returns a resolved Path on success.
    Raises FileValidationError with a descriptive message on any problem.
    """
    try:
        p = Path(path).resolve()
    except Exception as exc:
        raise FileValidationError(f"Cannot resolve path {path!r}: {exc}") from exc

    if not p.exists():
        raise FileValidationError(f"File not found: {p}")

    if not p.is_file():
        raise FileValidationError(f"Path does not point to a file: {p}")

    if p.suffix.lower() not in ALLOWED_EXTENSIONS:
        raise FileValidationError(
            f"Unsupported file type '{p.suffix}' — only .xlsx files are accepted: {p.name}"
        )

    size_bytes = p.stat().st_size
    if size_bytes == 0:
        raise FileValidationError(f"File is empty (0 bytes): {p.name}")

    size_mb = size_bytes / (1024 * 1024)
    if size_mb > MAX_FILE_SIZE_MB:
        raise FileValidationError(
            f"File is too large ({size_mb:.1f} MB, limit is {MAX_FILE_SIZE_MB} MB): {p.name}"
        )

    # xlsx files are zip archives — verify the container and reject macros.
    try:
        with zipfile.ZipFile(p, "r") as zf:
            names = zf.namelist()
            if "xl/workbook.xml" not in names:
                raise FileValidationError(
                    f"File does not appear to be a valid xlsx (xl/workbook.xml missing): {p.name}"
                )
            if any(n.startswith("xl/vbaProject") for n in names):
                raise FileValidationError(
                    f"File contains VBA macros and cannot be processed safely: {p.name}"
                )
    except zipfile.BadZipFile as exc:
        raise FileValidationError(
            f"File is not a valid xlsx archive (corrupt or disguised file): {p.name}"
        ) from exc

    return p


def _check_cell_for_injection(value, context: str = "") -> None:
    """Raises FileValidationError if a string cell value looks like a formula injection."""
    if isinstance(value, str):
        stripped = value.strip()
        if stripped.startswith(_FORMULA_PREFIXES):
            raise FileValidationError(
                f"Potential formula injection detected{' in ' + context if context else ''}: {value!r}"
            )


# ── Statistics loader ──────────────────────────────────────────────────────────

def _search_metadata(ws, label: str, search_rows: int = 15):
    """
    Searches the first `search_rows` rows for a cell whose string value
    contains `label` (case-insensitive). Returns the value of the
    immediately adjacent cell to the right, or None if not found.
    """
    for row in ws.iter_rows(min_row=1, max_row=search_rows, values_only=True):
        for col_idx, cell_val in enumerate(row):
            if isinstance(cell_val, str) and label.lower() in cell_val.strip().lower():
                if col_idx + 1 < len(row):
                    neighbour = row[col_idx + 1]
                    if neighbour is not None and str(neighbour).strip():
                        return str(neighbour).strip()
    return None


def _parse_year(value) -> int | None:
    """Parses a year from a datetime object or common date string formats."""
    if isinstance(value, datetime):
        return value.year
    if hasattr(value, "year"):  # date objects
        return value.year
    if isinstance(value, str):
        v = value.strip()
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(v, fmt).year
            except ValueError:
                continue
    return None


def load_statistics(path: str) -> dict:
    """
    Loads and validates the statistics xlsx file.

    Returns a dict:
        mitarbeiter    str | None
        befunddatum    str | None
        total          int | None   (expected examination count from L8)
        leistungen     list[dict]   (one entry per data row)
        errors         list[str]    (non-fatal warnings accumulated during load)

    Raises StatsValidationError for fatal structural problems.
    Raises FileValidationError for file-level security problems.
    """
    p = validate_xlsx_file(path)
    errors: list[str] = []

    try:
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    except Exception as exc:
        raise StatsValidationError(f"Cannot open statistics file: {exc}") from exc

    try:
        ws = wb.active

        # ── Metadata: Mitarbeiter ──────────────────────────────────────────
        mitarbeiter = _search_metadata(ws, "Mitarbeiter")
        if mitarbeiter is None:
            errors.append(
                "WARNING: 'Mitarbeiter' label not found in metadata — "
                "a generic filename will be used for the output."
            )

        # ── Metadata: Befunddatum ──────────────────────────────────────────
        befunddatum = _search_metadata(ws, "Befunddatum")
        if befunddatum is None:
            errors.append(
                "WARNING: 'Befunddatum' label not found in metadata — "
                "a generic filename will be used for the output."
            )

        # ── Metadata: total count from L8 (row 8, col 12) ─────────────────
        total_examinations: int | None = None
        try:
            l8_raw = ws.cell(row=8, column=12).value
            if l8_raw is not None:
                total_examinations = int(l8_raw)
            else:
                # Fallback: search for a 'Gesamt' row in the first 15 rows
                for row in ws.iter_rows(min_row=1, max_row=15, values_only=True):
                    row_strings = [str(v) for v in row if v is not None]
                    if "Gesamt" in row_strings:
                        for v in reversed(row):
                            if isinstance(v, (int, float)) and v > 0:
                                total_examinations = int(v)
                                break
                        break
                if total_examinations is None:
                    errors.append(
                        "WARNING: Total examination count not found in L8 / 'Gesamt' row — "
                        "count verification will be skipped."
                    )
        except (TypeError, ValueError) as exc:
            errors.append(
                f"WARNING: Value in L8 cannot be read as a number ({exc}) — "
                "count verification will be skipped."
            )

        # ── Locate the data header row (contains 'Leist-Kurz') ────────────
        header_row_num: int | None = None
        col_map: dict[str, int] = {}  # column name → 1-based column index

        for row_num, row in enumerate(ws.iter_rows(min_row=1, max_row=50, values_only=True), 1):
            if any(v == "Leist-Kurz" for v in row):
                header_row_num = row_num
                for col_idx, v in enumerate(row, 1):
                    if v is not None and str(v).strip():
                        col_map[str(v).strip()] = col_idx
                break

        if header_row_num is None:
            raise StatsValidationError(
                "Column header 'Leist-Kurz' not found in the first 50 rows. "
                "This does not appear to be a valid statistics file."
            )

        # Ensure mandatory columns exist.
        for required in ("Leist-Kurz", "DokDatum"):
            if required not in col_map:
                raise StatsValidationError(
                    f"Required column '{required}' missing from header row {header_row_num}. "
                    "Please verify the statistics file."
                )

        leist_kurz_col = col_map["Leist-Kurz"]
        dok_datum_col = col_map["DokDatum"]
        leistung_col = col_map.get("Leistung")
        ind1_col = col_map.get("IND1")
        ind2_col = col_map.get("IND2")

        # ── Read examination rows ──────────────────────────────────────────
        leistungen: list[dict] = []
        current_date = None
        row_count = 0
        # Collect IND2 values from every non-empty row (including rows that will
        # later be skipped for missing LeistKurz or bad date).  Used after the loop
        # to compute total_leistungen via the IND2 group-end algorithm.
        ind2_sequence: list[int] = []

        for row_num, row in enumerate(
            ws.iter_rows(min_row=header_row_num + 1, values_only=True),
            header_row_num + 1,
        ):
            row = list(row)

            # Guard against rows shorter than expected.
            def _get(col_1based):
                idx = col_1based - 1
                return row[idx] if idx < len(row) else None

            # Skip completely empty rows.
            if all(v is None for v in row[:25]):
                continue

            row_count += 1
            if row_count > MAX_DATA_ROWS:
                raise StatsValidationError(
                    f"File exceeds maximum allowed row count ({MAX_DATA_ROWS:,}). "
                    "Please verify the file is not malformed."
                )

            dok_datum = _get(dok_datum_col)
            leist_kurz = _get(leist_kurz_col)
            leistung = _get(leistung_col) if leistung_col else None
            ind1 = _get(ind1_col) if ind1_col else None
            ind2 = _get(ind2_col) if ind2_col else None

            # Collect IND2 for the group-end total algorithm (done after the loop).
            if ind2 is not None and ind2_col is not None:
                try:
                    ind2_sequence.append(int(ind2))
                except (TypeError, ValueError):
                    pass

            # Carry forward date from the first row of a visit group.
            if dok_datum is not None:
                current_date = dok_datum

            if leist_kurz is None:
                ind_info = f" (IND1={ind1})" if ind1 is not None else f" row {row_num}"
                errors.append(
                    f"WARNING: Row {row_num}{ind_info} — 'Leist-Kurz' is empty, row skipped."
                )
                continue

            leist_kurz_str = str(leist_kurz).strip()
            leistung_str = str(leistung).strip() if leistung else None

            # Injection check on user-supplied string values.
            try:
                _check_cell_for_injection(leist_kurz_str, f"Leist-Kurz at row {row_num}")
                if leistung_str:
                    _check_cell_for_injection(leistung_str, f"Leistung at row {row_num}")
            except FileValidationError:
                raise  # re-raise as-is; caller treats FileValidationError as fatal

            year = _parse_year(current_date)
            if year is None:
                ind_info = f" (IND1={ind1})" if ind1 is not None else ""
                errors.append(
                    f"WARNING: Row {row_num}{ind_info} — cannot parse year from "
                    f"DokDatum {current_date!r}, row skipped."
                )
                continue

            leistungen.append(
                {
                    "leist_kurz": leist_kurz_str,
                    "leistung": leistung_str,
                    "year": year,
                    "ind1": ind1,
                    "row": row_num,
                }
            )

        if not leistungen:
            raise StatsValidationError(
                "No valid examination rows found in the statistics file. "
                "Check that the file contains data below the header row."
            )

        # ── Compute total_leistungen from IND2 sequence ────────────────────
        # Walk the IND2 sequence once.  When the next value is ≤ the current
        # value (or there is no next value), the current row is the last in its
        # group and its IND2 value equals the group's size — add it to the total.
        # This works without using IND1 at all.
        total_leistungen: int | None = None
        if ind2_sequence:
            total_leistungen = 0
            for i, val in enumerate(ind2_sequence):
                if i == len(ind2_sequence) - 1 or ind2_sequence[i + 1] <= val:
                    total_leistungen += val

    finally:
        wb.close()

    return {
        "mitarbeiter": mitarbeiter,
        "befunddatum": befunddatum,
        "total_documents": total_examinations,  # L8 = Befunddokumente (unique IND1), info only
        "total_leistungen": total_leistungen,   # computed from IND2 group-end algorithm
        "leistungen": leistungen,
        "errors": errors,
    }


# ── Map loader ─────────────────────────────────────────────────────────────────

def load_map(path: str, include_underscore_columns: bool = False) -> dict:
    """
    Loads and validates the map xlsx file.

    Parameters
    ----------
    include_underscore_columns : bool
        If False (default), columns whose header starts with '_' are excluded
        from the output sections (e.g. _CLIP, _TRENNER, _Zeit).

    Returns a dict:
        sections            list[str]         ordered section names
        leistungen          dict[str, dict]   leist0 → {kurztext, sections, row}
        merged_duplicates   list[str]         leist0 codes merged by max() across rows

    Raises MapValidationError for structural/content problems.
    Raises FileValidationError for file-level security problems.
    """
    p = validate_xlsx_file(path)

    try:
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    except Exception as exc:
        raise MapValidationError(f"Cannot open map file: {exc}") from exc

    try:
        ws = wb.active

        # Read the header row only.
        header_rows = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        if not header_rows:
            raise MapValidationError("Map file appears to be empty (no rows found).")
        header = list(header_rows[0])

        # ── Structural checks: B1 and F1 ──────────────────────────────────
        b1 = header[1] if len(header) > 1 else None
        if b1 != "Leist0":
            raise MapValidationError(
                f"Map validation failed: B1 is {b1!r}, expected 'Leist0'. "
                "Please provide a valid map file. "
                "(An admin can upload a corrected map file to resolve this.)",
                recoverable=False,
            )

        f1 = header[5] if len(header) > 5 else None
        if f1 != "DL Gefäße":
            raise MapValidationError(
                f"Map validation failed: F1 is {f1!r}, expected 'DL Gefäße'. "
                "The map file may use a different column layout. "
                "(An admin can upload a corrected map file to resolve this.)",
                recoverable=True,
            )

        # ── Collect section columns (F / index 5 onwards) ─────────────────
        sections: list[str] = []
        section_col_indices: list[int] = []  # 0-based

        for col_idx, col_name in enumerate(header[5:], 5):
            if col_name is None:
                continue
            name = str(col_name).strip()
            if not name:
                continue
            if not include_underscore_columns and name.startswith("_"):
                continue
            sections.append(name)
            section_col_indices.append(col_idx)

        if not sections:
            raise MapValidationError(
                "No section columns found starting from column F. "
                "Please verify the map file structure."
            )

        # ── Read leistung rows ─────────────────────────────────────────────
        # Track duplicates: leist0 → list of candidate entries
        candidates: dict[str, list[dict]] = {}
        row_count = 0

        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
            row = list(row)

            def _get_map(col_0based):
                return row[col_0based] if col_0based < len(row) else None

            leist0 = _get_map(1)  # column B
            if leist0 is None:
                continue

            leist0_str = str(leist0).strip()
            if not leist0_str:
                continue

            row_count += 1
            if row_count > MAX_DATA_ROWS:
                raise MapValidationError(
                    f"Map file exceeds maximum allowed row count ({MAX_DATA_ROWS:,})."
                )

            try:
                _check_cell_for_injection(leist0_str, f"Leist0 at map row {row_num}")
            except FileValidationError:
                raise

            kurztext_raw = _get_map(2)  # column C
            kurztext = str(kurztext_raw).strip() if kurztext_raw else ""

            section_values: dict[str, int] = {}
            for section_name, col_idx in zip(sections, section_col_indices):
                val = _get_map(col_idx)
                section_values[section_name] = 1 if val == 1 else 0

            entry = {
                "kurztext": kurztext,
                "sections": section_values,
                "row": row_num,
            }

            candidates.setdefault(leist0_str, []).append(entry)

        if not candidates:
            raise MapValidationError(
                "No leistung entries found in the map file. "
                "Please verify the file contains data below the header row."
            )

        # ── Resolve duplicates ─────────────────────────────────────────────
        leistungen: dict[str, dict] = {}
        merged_duplicates: list[str] = []

        for leist0_str, entries in candidates.items():
            if len(entries) == 1:
                leistungen[leist0_str] = entries[0]
                continue

            has_values = [e for e in entries if any(v == 1 for v in e["sections"].values())]

            if len(has_values) == 0:
                # None have any section assigned — keep first, no harm done.
                leistungen[leist0_str] = entries[0]

            elif len(has_values) == 1:
                # Exactly one entry has values — use it.
                leistungen[leist0_str] = has_values[0]

            else:
                # Multiple entries have values — merge by taking max() per column.
                reference_sections = has_values[0]["sections"]
                if all(e["sections"] == reference_sections for e in has_values[1:]):
                    # Identical assignments — no ambiguity, pick first.
                    leistungen[leist0_str] = has_values[0]
                else:
                    # Conflict: merge all entries by always taking the higher value.
                    merged_sections: dict[str, int] = {}
                    for section_name in sections:
                        merged_sections[section_name] = max(
                            e["sections"].get(section_name, 0) for e in has_values
                        )
                    rows_str = ", ".join(str(e["row"]) for e in entries)
                    leistungen[leist0_str] = {
                        "kurztext": has_values[0]["kurztext"],
                        "sections": merged_sections,
                        "row": rows_str,  # all source rows for traceability
                    }
                    merged_duplicates.append(leist0_str)

    finally:
        wb.close()

    return {
        "sections": sections,
        "leistungen": leistungen,
        "merged_duplicates": merged_duplicates,
    }
