"""
output_handler.py — xlsx generation from a processing result dict.

Two public entry points:
  build_xlsx_bytes(result) -> (filename, bytes)
      Builds the workbook entirely in memory. Used by the iii web worker —
      no files are ever written to disk on the server.
  save_output(result, output_dir) -> str
      Builds the workbook and saves it to disk. Used by the CLI.
"""

import io
import re
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── Palette ────────────────────────────────────────────────────────────────────
# One dark accent (header only), one near-white accent (Gesamt/Benötigt),
# and two text colours for severity. No fills in the footer.

_FONT         = "Calibri"

_HEADER_BG    = "1F3864"   # deep navy — header row only
_ACCENT_BG    = "EEF2F7"   # very light blue-gray — Gesamt + Benötigt columns
_SEP_C        = "A0AABB"   # medium gray — vertical separator before Gesamt
_GRID_C       = "D8DCE3"   # light gray  — horizontal row lines

_TEXT_BODY    = "1A1A2E"   # near-black  — section names and counts
_TEXT_MUTED   = "6B7280"   # gray        — footnotes / info lines
_TEXT_ERROR   = "C0392B"   # muted red   — errors
_TEXT_WARN    = "A0522D"   # muted amber — warnings and notes

# Achievement colours (soft pastels — used in conditional formatting rules)
_ACH_GREEN    = "D5F5E3"   # ≥ 100 %
_ACH_YELLOW   = "FCF3CF"   # ≥  80 %
_ACH_ORANGE   = "FAE5D3"   # ≥  50 %
_ACH_RED      = "FADBD8"   # <   50 %

# ── Shared style objects ───────────────────────────────────────────────────────

def _fill(hex_color: str) -> PatternFill:
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

# Data row: light bottom line only (open-grid look)
_ROW_BORDER = Border(bottom=Side(style="thin", color=_GRID_C))

# Gesamt / Benötigt: same bottom line + a medium left separator
_ACCENT_BORDER = Border(
    left=Side(style="medium", color=_SEP_C),
    bottom=Side(style="thin", color=_GRID_C),
)

# Header: slightly stronger bottom to close the header band
_HEADER_BORDER = Border(bottom=Side(style="medium", color=_SEP_C))


# ── Cell stylers ──────────────────────────────────────────────────────────────

def _header_cell(cell, value: str, align: str = "center") -> None:
    cell.value = value
    cell.font = Font(name=_FONT, bold=True, size=11, color="FFFFFF")
    cell.fill = _fill(_HEADER_BG)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    cell.border = _HEADER_BORDER


def _section_cell(cell, value: str) -> None:
    cell.value = value
    cell.font = Font(name=_FONT, size=10, color=_TEXT_BODY)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    cell.border = _ROW_BORDER


def _count_cell(cell, value: int) -> None:
    cell.value = value
    cell.font = Font(name=_FONT, size=10, color=_TEXT_BODY)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.border = _ROW_BORDER


def _accent_cell(cell, value, bold: bool = True) -> None:
    """Gesamt and Benötigt columns — subtle background, left separator."""
    cell.value = value
    cell.font = Font(name=_FONT, size=10, bold=bold, color=_TEXT_BODY)
    cell.fill = _fill(_ACCENT_BG)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.border = _ACCENT_BORDER


# ── Helpers ───────────────────────────────────────────────────────────────────

def _sections_to_suffix(sections: list) -> str:
    """['MRT Mamma', 'MRT Prostata'] → 'MrtMamma_MrtProstata'"""
    # Strip filesystem-invalid characters so a section name like 'MRT/CT' cannot
    # introduce a path separator into the filename.
    _safe = re.compile(r"[^\w]")
    return "_".join(
        "".join(_safe.sub("", word).capitalize() for word in s.split())
        for s in sections
    )


def _make_filename(
    mitarbeiter: "str | None",
    befunddatum: "str | None",
    section_suffix: "str | None" = None,
) -> str:
    if mitarbeiter and befunddatum:
        clean = lambda s: re.sub(r"[^\w\-]", "", s.replace(" ", ""))
        name = f"{clean(mitarbeiter)}{clean(befunddatum)}Auswertung"
    else:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        name = f"Auswertung_{ts}"
    if section_suffix:
        name = f"{name}_{section_suffix}"
    return f"{name}.xlsx"


def _add_achievement_cf(
    ws, gesamt_col: int, benoetigt_col: int, first_row: int, last_row: int
) -> None:
    """
    Conditional formatting on the Gesamt column.
    Colour updates live in Excel when the user edits Benötigt values.
    """
    try:
        g  = get_column_letter(gesamt_col)
        b  = get_column_letter(benoetigt_col)
        cf_range = f"{g}{first_row}:{g}{last_row}"
        gf = f"{g}{first_row}"
        bf = f"{b}{first_row}"
        for formula, color in (
            (f'AND({bf}<>"",{bf}>0,{gf}/{bf}>=1)',   _ACH_GREEN),
            (f'AND({bf}<>"",{bf}>0,{gf}/{bf}>=0.8)', _ACH_YELLOW),
            (f'AND({bf}<>"",{bf}>0,{gf}/{bf}>=0.5)', _ACH_ORANGE),
            (f'AND({bf}<>"",{bf}>0,{gf}/{bf}>0)',    _ACH_RED),
        ):
            ws.conditional_formatting.add(
                cf_range,
                FormulaRule(
                    formula=[formula],
                    fill=_fill(color),
                    stopIfTrue=True,
                ),
            )
    except Exception:
        pass  # CF is cosmetic — never let it break the file


def _set_column_widths(ws, n_years: int, gesamt_col: int, has_reference: bool) -> None:
    """
    Fixed widths: adaptive section-name column, compact fixed number columns.
    Section names: measure actual content, cap at 45.
    """
    max_section_len = max(
        (len(str(ws.cell(row=r, column=1).value or ""))
         for r in range(2, ws.max_row + 1)),
        default=20,
    )
    ws.column_dimensions["A"].width = min(max_section_len + 4, 45)

    for col in range(2, gesamt_col):                                   # year columns
        ws.column_dimensions[get_column_letter(col)].width = 11

    ws.column_dimensions[get_column_letter(gesamt_col)].width = 13    # Gesamt

    if has_reference:
        ws.column_dimensions[get_column_letter(gesamt_col + 1)].width = 13  # Benötigt


# ── Workbook builder ──────────────────────────────────────────────────────────

def _build_workbook(
    result: dict,
    reference_amounts: "dict[str, int] | None" = None,
) -> "tuple[str, openpyxl.Workbook]":
    """
    Build an openpyxl Workbook from a processing result dict.
    Returns (filename, workbook) — no I/O performed.
    """
    filter_secs: list = result.get("filter_sections", [])
    section_suffix = _sections_to_suffix(filter_secs) if filter_secs else None
    filename = _make_filename(
        result.get("mitarbeiter"),
        result.get("befunddatum"),
        section_suffix=section_suffix,
    )

    years:             list[int]       = result["years"]
    data:              dict[str, dict] = result["data"]
    errors:            list[str]       = result.get("errors", [])
    total_counted:     int             = result.get("total_counted", 0)
    total_leistungen:  int | None      = result.get("total_leistungen")
    total_documents:   int | None      = result.get("total_documents")
    merged_duplicates: list[str]       = result.get("merged_duplicates", [])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Auswertung"
    ws.freeze_panes = "B2"

    _ref:       dict[str, int] = reference_amounts or {}
    _ref_lower: dict[str, int] = {k.lower(): v for k, v in _ref.items()}
    has_reference = bool(_ref)

    def _lookup_required(section_name: str) -> "int | None":
        try:
            return _ref.get(section_name) or _ref_lower.get(section_name.lower())
        except Exception:
            return None

    gesamt_col    = len(years) + 2
    benoetigt_col = gesamt_col + 1

    col_headers = ["Leistungsbereich"] + [str(y) for y in years] + ["Gesamt"]
    if has_reference:
        col_headers.append("Benötigt")

    # ── Header row ────────────────────────────────────────────────────────────
    _header_cell(ws.cell(row=1, column=1), "Leistungsbereich", align="left")
    for col_idx, text in enumerate(col_headers[1:], 2):
        _header_cell(ws.cell(row=1, column=col_idx), text)
    ws.row_dimensions[1].height = 34

    # ── Data rows ─────────────────────────────────────────────────────────────
    for row_offset, (section, year_data) in enumerate(data.items(), 2):
        _section_cell(ws.cell(row=row_offset, column=1), section)

        for col_offset, year in enumerate(years, 2):
            _count_cell(ws.cell(row=row_offset, column=col_offset), year_data.get(year, 0))

        _accent_cell(ws.cell(row=row_offset, column=gesamt_col), year_data.get("Total", 0))

        if has_reference:
            required = _lookup_required(section)
            try:
                _accent_cell(
                    ws.cell(row=row_offset, column=benoetigt_col),
                    required,          # None leaves cell blank but keeps styling
                    bold=required is not None,
                )
            except Exception:
                pass

        ws.row_dimensions[row_offset].height = 20

    # Conditional formatting on Gesamt column (live colour in Excel)
    if has_reference and data:
        _add_achievement_cf(
            ws, gesamt_col, benoetigt_col, first_row=2, last_row=len(data) + 1
        )

    # ── Column widths ─────────────────────────────────────────────────────────
    _set_column_widths(ws, len(years), gesamt_col, has_reference)

    # ── Footer ────────────────────────────────────────────────────────────────
    next_row = len(data) + 3   # one blank row separates data from footer

    # Count summary
    lbl = ws.cell(row=next_row, column=1)
    lbl.value = "Gezählte Leistungen:"
    lbl.font = Font(name=_FONT, size=10, bold=True, color=_TEXT_BODY)

    cnt = ws.cell(row=next_row, column=2)
    cnt.value = total_counted
    cnt.font = Font(name=_FONT, size=10, bold=True, color=_TEXT_BODY)
    cnt.alignment = Alignment(horizontal="right")

    # Befunddokumente (L8) — informational only, not a count of Leistungen
    if total_documents is not None:
        next_row += 1
        doc = ws.cell(row=next_row, column=1)
        doc.value = (
            f"Befunddokumente (L8):   {total_documents:,}"
            "   —   Anzahl Befundberichte (IND1), nicht mit Leistungen vergleichbar"
        )
        doc.font = Font(name=_FONT, size=9, italic=True, color=_TEXT_MUTED)
        ws.merge_cells(
            start_row=next_row, start_column=1,
            end_row=next_row, end_column=len(col_headers),
        )

    # Count mismatch — text colour only, no background fill
    if total_leistungen is not None and total_counted != total_leistungen:
        skipped = total_leistungen - total_counted
        next_row += 1
        mm = ws.cell(row=next_row, column=1)
        mm.value = (
            f"ABWEICHUNG: IND2-Algorithmus erwartet {total_leistungen:,} Leistungen "
            f"— {skipped:,} Zeilen übersprungen"
        )
        mm.font = Font(name=_FONT, size=10, bold=True, color=_TEXT_ERROR)
        ws.merge_cells(
            start_row=next_row, start_column=1,
            end_row=next_row, end_column=len(col_headers),
        )
        ws.row_dimensions[next_row].height = 22

    # ── Merged duplicates notice ──────────────────────────────────────────────
    if merged_duplicates:
        next_row += 2
        hl = ws.cell(row=next_row, column=1)
        hl.value = (
            "Hinweis: Doppelte Leistungsbezeichnungen "
            "— Werte wurden per Maximum zusammengeführt"
        )
        hl.font = Font(name=_FONT, size=10, bold=True, color=_TEXT_WARN)
        ws.merge_cells(
            start_row=next_row, start_column=1,
            end_row=next_row, end_column=len(col_headers),
        )

        for code in merged_duplicates:
            next_row += 1
            c = ws.cell(row=next_row, column=1)
            c.value = f"    •  {code}"
            c.font = Font(name=_FONT, size=10, color=_TEXT_WARN)

    # ── Errors / warnings section ─────────────────────────────────────────────
    if errors:
        next_row += 2
        lbl = ws.cell(row=next_row, column=1)
        lbl.value = "Hinweise und Fehler"
        lbl.font = Font(name=_FONT, size=10, bold=True, color=_TEXT_BODY)

        for error_text in errors:
            next_row += 1
            cell = ws.cell(row=next_row, column=1)
            cell.value = error_text
            cell.alignment = Alignment(wrap_text=True)
            ws.row_dimensions[next_row].height = 28

            if error_text.startswith("COUNT MISMATCH") or error_text.startswith("ERROR"):
                cell.font = Font(name=_FONT, size=9, color=_TEXT_ERROR)
            elif error_text.startswith("WARNING"):
                cell.font = Font(name=_FONT, size=9, color=_TEXT_WARN)
            else:
                cell.font = Font(name=_FONT, size=9, color=_TEXT_MUTED)

            ws.merge_cells(
                start_row=next_row, start_column=1,
                end_row=next_row, end_column=len(col_headers),
            )

    return filename, wb


# ── Public entry points ───────────────────────────────────────────────────────

def build_xlsx_bytes(
    result: dict,
    reference_amounts: "dict[str, int] | None" = None,
) -> "tuple[str, bytes]":
    """
    Build the xlsx entirely in memory and return (filename, raw_bytes).

    No files are written to disk — intended for the iii web worker where
    the bytes are streamed directly back to the HTTP client.
    """
    filename, wb = _build_workbook(result, reference_amounts=reference_amounts)
    buf = io.BytesIO()
    wb.save(buf)
    return filename, buf.getvalue()


def save_output(
    result: dict,
    output_dir: str = ".",
    reference_amounts: "dict[str, int] | None" = None,
) -> str:
    """
    Build the xlsx and save it to `output_dir` on disk.

    Returns the absolute path of the saved file.
    Raises IOError if the directory cannot be created or the file cannot be written.
    """
    filename, wb = _build_workbook(result, reference_amounts=reference_amounts)

    try:
        out_dir = Path(output_dir).resolve()
        out_dir.mkdir(parents=True, exist_ok=True)
        output_path = out_dir / filename
    except Exception as exc:
        raise IOError(f"Cannot prepare output directory '{output_dir}': {exc}") from exc

    try:
        wb.save(output_path)
    except Exception as exc:
        raise IOError(f"Failed to save output file to '{output_path}': {exc}") from exc

    return str(output_path)
