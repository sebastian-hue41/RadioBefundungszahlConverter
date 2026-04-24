"""Pytest configuration and test helpers."""

import io
import sys
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl
import pytest

# Add project root to sys.path so tests can import project modules
sys.path.insert(0, str(Path(__file__).parent.parent))


# ── Workbook builders ──────────────────────────────────────────────────────────

def make_stats_wb(
    rows: list,
    mitarbeiter: str = "KUNZ_A",
    befunddatum: str = "01.01.2024-31.12.2024",
    total_l8: int | None = None,
) -> openpyxl.Workbook:
    """
    Create a statistics workbook with the given structure.

    rows: list of (ind1, ind2, dok_datum, leistung, leist_kurz) tuples
    dok_datum can be None (carry forward from previous row)
    """
    wb = openpyxl.Workbook()
    ws = wb.active

    # Metadata rows
    if mitarbeiter:
        ws["A1"] = "Mitarbeiter:"
        ws["B1"] = mitarbeiter
    if befunddatum:
        ws["A2"] = "Befunddatum:"
        ws["B2"] = befunddatum

    # L8 total count
    if total_l8 is not None:
        ws.cell(row=8, column=12).value = total_l8

    # Header row at row 10
    ws["A10"] = "IND1"
    ws["B10"] = "IND2"
    ws["C10"] = "DokDatum"
    ws["D10"] = "Leistung"
    ws["E10"] = "Leist-Kurz"

    # Data rows starting at row 11
    for i, (ind1, ind2, dok_datum, leistung, leist_kurz) in enumerate(rows, 11):
        ws.cell(row=i, column=1).value = ind1
        ws.cell(row=i, column=2).value = ind2
        ws.cell(row=i, column=3).value = dok_datum
        ws.cell(row=i, column=4).value = leistung
        ws.cell(row=i, column=5).value = leist_kurz

    return wb


def make_map_wb(
    codes: list,
    sections: list[str] | None = None,
    b1: str = "Leist0",
) -> openpyxl.Workbook:
    """
    Create a map workbook.

    codes: list of (leist0, kurztext, {section: 1/0, ...}) tuples
    sections: list of section names (defaults to ["DL Gefäße", "MRT", "CT"])
    """
    if sections is None:
        sections = ["DL Gefäße", "MRT", "CT"]

    wb = openpyxl.Workbook()
    ws = wb.active

    # Header row: B1 = Leist0, F1+ = sections
    ws.cell(row=1, column=2).value = b1
    for i, sec in enumerate(sections, 6):  # column 6 = F
        ws.cell(row=1, column=i).value = sec

    # Data rows
    for row_i, (leist0, kurztext, sec_vals) in enumerate(codes, 2):
        ws.cell(row=row_i, column=2).value = leist0
        ws.cell(row=row_i, column=3).value = kurztext
        for j, sec in enumerate(sections, 6):
            ws.cell(row=row_i, column=j).value = sec_vals.get(sec, 0)

    return wb


def make_reference_wb(
    amounts: dict[str, int],
    include_header: bool = True,
) -> openpyxl.Workbook:
    """
    Create a reference amounts workbook.

    amounts: {section_name: required_count}
    """
    wb = openpyxl.Workbook()
    ws = wb.active

    row = 1
    if include_header:
        ws.cell(row=row, column=1).value = "Leistungsbereich"
        ws.cell(row=row, column=2).value = "Benötigt"
        row += 1

    for name, amount in amounts.items():
        ws.cell(row=row, column=1).value = name
        ws.cell(row=row, column=2).value = amount
        row += 1

    return wb


def save_wb(wb: openpyxl.Workbook, path: Path) -> str:
    """Save workbook to path and return the path string."""
    wb.save(path)
    return str(path)


# ── Adversarial file builders ──────────────────────────────────────────────────

def make_corrupt_xlsx_bytes() -> bytes:
    """Create a file that's not a zip but has .xlsx extension."""
    return b"this is not a zip file but claims to be an xlsx"


def make_zip_without_workbook_xml() -> bytes:
    """Create a valid zip but without xl/workbook.xml."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("xl/not_workbook.xml", "<xml/>")
        zf.writestr("some_file.txt", "content")
    return buf.getvalue()


def inject_vba_into_xlsx(xlsx_bytes: bytes) -> bytes:
    """Take valid xlsx bytes and inject a VBA macro file."""
    src = io.BytesIO(xlsx_bytes)
    dst = io.BytesIO()
    with zipfile.ZipFile(src, "r") as zin:
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                zout.writestr(item, zin.read(item.filename))
            # Inject fake VBA project
            zout.writestr("xl/vbaProject.bin", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1fake vba")
    return dst.getvalue()


# ── Result/data builders for logic and output tests ─────────────────────────────

def make_stats_data(
    examinations: list[dict],
    errors: list[str] | None = None,
    total_leistungen: int | None = None,
    total_documents: int | None = None,
    mitarbeiter: str | None = "TEST",
    befunddatum: str | None = "2024",
) -> dict:
    """Build a stats_data dict as returned by input_handler.load_statistics()."""
    return {
        "errors": errors or [],
        "leistungen": examinations,
        "total_leistungen": total_leistungen,
        "total_documents": total_documents,
        "mitarbeiter": mitarbeiter,
        "befunddatum": befunddatum,
    }


def make_map_data(
    sections: list[str],
    leistungen: dict,
    merged_duplicates: list[str] | None = None,
) -> dict:
    """Build a map_data dict as returned by input_handler.load_map()."""
    return {
        "sections": sections,
        "leistungen": leistungen,
        "merged_duplicates": merged_duplicates or [],
    }


def make_exam(
    leist_kurz: str,
    year: int,
    leistung: str | None = None,
) -> dict:
    """Build an exam entry for stats_data."""
    return {
        "leist_kurz": leist_kurz,
        "leistung": leistung,
        "year": year,
        "ind1": 1,
        "row": 11,
    }


def make_leistung_entry(sections_dict: dict[str, int]) -> dict:
    """Build a leistung entry for map_data."""
    return {
        "kurztext": "",
        "sections": sections_dict,
        "row": 2,
    }


def make_result(
    sections: list[str],
    year_counts: dict[str, dict[int, int]],
    years: list[int],
    errors: list[str] | None = None,
    mitarbeiter: str | None = "TEST",
    befunddatum: str | None = "2024",
    total_counted: int = 10,
    total_leistungen: int | None = 10,
    total_documents: int | None = None,
    merged_duplicates: list[str] | None = None,
) -> dict:
    """
    Build a result dict as returned by logic.process().

    year_counts: {section_name: {year: count, ...}}
    """
    data = {}
    for section in sections:
        counts = {}
        for year in years:
            counts[year] = year_counts.get(section, {}).get(year, 0)
        counts["Total"] = sum(counts[y] for y in years)
        data[section] = counts

    return {
        "data": data,
        "years": years,
        "errors": errors or [],
        "total_counted": total_counted,
        "total_leistungen": total_leistungen,
        "total_documents": total_documents,
        "mitarbeiter": mitarbeiter,
        "befunddatum": befunddatum,
        "merged_duplicates": merged_duplicates or [],
    }


@pytest.fixture
def tmp_xlsx(tmp_path):
    """Return a function to save workbooks to tmp_path."""
    def _save(wb, name="test.xlsx"):
        path = tmp_path / name
        wb.save(path)
        return str(path)
    return _save
