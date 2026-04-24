"""
test_file_to_memory.py

Verifies that every value written into an xlsx file is read back exactly as
expected — correct field, correct type, correct content.

These tests exist because the counting tests only check row counts/totals.
A bug that reads the wrong column, misparses a date, or drops a field value
would pass the counting tests but be caught here.
"""

import openpyxl
import pytest

from input_handler import load_statistics, load_map
from .helpers import make_stats_wb, make_map_wb, save_wb


# ── Statistics: individual field values ───────────────────────────────────────

class TestStatisticsFieldValues:
    """Each field in the loaded exam dict must exactly match the file content."""

    def test_leist_kurz_preserved(self, tmp_path):
        """Leist-Kurz codes are read back exactly as written."""
        codes = ["MRT", "CT001", "US_Abdomen", "DOPPELGANGER", "X42"]
        rows = [(i, 1, "15.03.2024", f"Desc {c}", c) for i, c in enumerate(codes, 1)]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))

        loaded_codes = [e["leist_kurz"] for e in result["leistungen"]]
        assert loaded_codes == codes, f"Codes diverged: {loaded_codes} != {codes}"

    def test_leistung_description_preserved(self, tmp_path):
        """Leistungsbezeichnung is read back exactly as written."""
        rows = [
            (1, 1, "15.03.2024", "MR Kopf nativ",        "MRT"),
            (2, 1, "15.03.2024", "CT Thorax mit KM",      "CT"),
            (3, 1, "15.03.2024", "Sonographie Abdomen",   "US"),
        ]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))

        descs = [e["leistung"] for e in result["leistungen"]]
        assert descs[0] == "MR Kopf nativ"
        assert descs[1] == "CT Thorax mit KM"
        assert descs[2] == "Sonographie Abdomen"

    def test_year_parsed_from_german_date_string(self, tmp_path):
        """German date strings (DD.MM.YYYY) are parsed to the correct year."""
        rows = [
            (1, 1, "01.01.2022", "L", "MRT"),
            (2, 1, "15.06.2023", "L", "CT"),
            (3, 1, "31.12.2025", "L", "US"),
        ]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))

        years = [e["year"] for e in result["leistungen"]]
        assert years == [2022, 2023, 2025]

    def test_year_parsed_from_iso_date_string(self, tmp_path):
        """ISO date strings (YYYY-MM-DD) are parsed to the correct year."""
        wb = make_stats_wb([])
        ws = wb.active
        ws["A10"] = "IND1"
        ws["B10"] = "IND2"
        ws["C10"] = "DokDatum"
        ws["D10"] = "Leistung"
        ws["E10"] = "Leist-Kurz"
        ws.cell(row=11, column=1).value = 1
        ws.cell(row=11, column=2).value = 1
        ws.cell(row=11, column=3).value = "2024-07-15"
        ws.cell(row=11, column=4).value = "L"
        ws.cell(row=11, column=5).value = "MRT"

        result = load_statistics(save_wb(wb, tmp_path / "s.xlsx"))
        assert result["leistungen"][0]["year"] == 2024

    def test_year_parsed_from_datetime_object(self, tmp_path):
        """datetime values in DokDatum cells are parsed to the correct year."""
        from datetime import datetime
        wb = make_stats_wb([])
        ws = wb.active
        ws["A10"] = "IND1"
        ws["B10"] = "IND2"
        ws["C10"] = "DokDatum"
        ws["D10"] = "Leistung"
        ws["E10"] = "Leist-Kurz"
        ws.cell(row=11, column=1).value = 1
        ws.cell(row=11, column=2).value = 1
        ws.cell(row=11, column=3).value = datetime(2026, 3, 20)
        ws.cell(row=11, column=4).value = "L"
        ws.cell(row=11, column=5).value = "CT"

        result = load_statistics(save_wb(wb, tmp_path / "s.xlsx"))
        assert result["leistungen"][0]["year"] == 2026

    def test_date_carryforward_applies_correct_year(self, tmp_path):
        """Rows without DokDatum inherit the year from the previous row in the group."""
        rows = [
            (1, 1, "20.06.2023", "L", "MRT"),
            (1, 2, None,         "L", "CT"),
            (1, 3, None,         "L", "US"),
        ]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))

        for exam in result["leistungen"]:
            assert exam["year"] == 2023, (
                f"Expected year 2023 for {exam['leist_kurz']}, got {exam['year']}"
            )

    def test_year_resets_for_new_group(self, tmp_path):
        """Date resets correctly when a new group with a different year starts."""
        rows = [
            (1, 1, "15.03.2024", "L", "MRT"),
            (1, 2, None,         "L", "CT"),   # inherits 2024
            (2, 1, "20.06.2025", "L", "US"),   # new group, new year
            (2, 2, None,         "L", "PET"),  # inherits 2025
        ]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))

        years = [e["year"] for e in result["leistungen"]]
        assert years == [2024, 2024, 2025, 2025]

    def test_ind1_preserved(self, tmp_path):
        """IND1 values are read back correctly."""
        rows = [
            (100, 1, "15.03.2024", "L", "MRT"),
            (200, 1, "15.03.2024", "L", "CT"),
            (300, 1, "15.03.2024", "L", "US"),
        ]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))

        ind1_values = [e["ind1"] for e in result["leistungen"]]
        assert ind1_values == [100, 200, 300]

    def test_row_order_preserved(self, tmp_path):
        """Rows are returned in file order."""
        codes = ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH"]
        rows = [(i, 1, "15.03.2024", f"D{c}", c) for i, c in enumerate(codes, 1)]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))

        loaded = [e["leist_kurz"] for e in result["leistungen"]]
        assert loaded == codes

    def test_whitespace_stripped_from_leist_kurz(self, tmp_path):
        """Leading/trailing whitespace in Leist-Kurz is stripped."""
        wb = make_stats_wb([])
        ws = wb.active
        ws["A10"] = "IND1"; ws["B10"] = "IND2"; ws["C10"] = "DokDatum"
        ws["D10"] = "Leistung"; ws["E10"] = "Leist-Kurz"
        ws.cell(row=11, column=1).value = 1
        ws.cell(row=11, column=2).value = 1
        ws.cell(row=11, column=3).value = "15.03.2024"
        ws.cell(row=11, column=4).value = "L"
        ws.cell(row=11, column=5).value = "  MRT  "  # extra spaces

        result = load_statistics(save_wb(wb, tmp_path / "s.xlsx"))
        assert result["leistungen"][0]["leist_kurz"] == "MRT"

    def test_column_order_independence(self, tmp_path):
        """Loader works correctly when columns are in a non-standard order."""
        # Reorder: Leist-Kurz first, DokDatum last
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Mitarbeiter:"; ws["B1"] = "REORDER_TEST"
        # Header: different column order
        ws["A10"] = "Leist-Kurz"   # col 1
        ws["B10"] = "IND1"         # col 2
        ws["C10"] = "IND2"         # col 3
        ws["D10"] = "Leistung"     # col 4
        ws["E10"] = "DokDatum"     # col 5
        # Data row
        ws.cell(row=11, column=1).value = "REORDERED"
        ws.cell(row=11, column=2).value = 1
        ws.cell(row=11, column=3).value = 1
        ws.cell(row=11, column=4).value = "Desc"
        ws.cell(row=11, column=5).value = "15.03.2024"

        result = load_statistics(save_wb(wb, tmp_path / "s.xlsx"))
        assert result["leistungen"][0]["leist_kurz"] == "REORDERED"
        assert result["leistungen"][0]["year"] == 2024

    def test_100_distinct_codes_all_present(self, tmp_path):
        """100 distinct codes are all read back with correct identity."""
        codes = [f"CODE_{i:03d}" for i in range(100)]
        rows = [(i, 1, "15.03.2024", f"Desc {c}", c) for i, c in enumerate(codes, 1)]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))

        loaded = {e["leist_kurz"] for e in result["leistungen"]}
        missing = set(codes) - loaded
        extra   = loaded - set(codes)
        assert not missing, f"Codes missing from memory: {missing}"
        assert not extra,   f"Unexpected codes in memory: {extra}"


# ── Statistics: metadata fields ───────────────────────────────────────────────

class TestStatisticsMetadata:
    """Mitarbeiter, Befunddatum, and L8 must round-trip correctly from file."""

    def test_mitarbeiter_read_correctly(self, tmp_path):
        """Mitarbeiter metadata is read back exactly."""
        rows = [(1, 1, "15.03.2024", "L", "MRT")]
        result = load_statistics(
            save_wb(make_stats_wb(rows, mitarbeiter="POEHLS_B"), tmp_path / "s.xlsx")
        )
        assert result["mitarbeiter"] == "POEHLS_B"

    def test_befunddatum_read_correctly(self, tmp_path):
        """Befunddatum metadata is read back exactly."""
        rows = [(1, 1, "15.03.2024", "L", "MRT")]
        result = load_statistics(
            save_wb(
                make_stats_wb(rows, befunddatum="01.01.2024-31.12.2024"),
                tmp_path / "s.xlsx",
            )
        )
        assert result["befunddatum"] == "01.01.2024-31.12.2024"

    def test_l8_total_documents_read_correctly(self, tmp_path):
        """L8 (total documents count) is read back correctly."""
        rows = [(1, 1, "15.03.2024", "L", "MRT")]
        result = load_statistics(
            save_wb(make_stats_wb(rows, total_l8=42), tmp_path / "s.xlsx")
        )
        assert result["total_documents"] == 42

    def test_metadata_search_finds_label_in_any_column(self, tmp_path):
        """Metadata search works even if label is not in column A."""
        wb = openpyxl.Workbook()
        ws = wb.active
        # Label in col C, value in col D
        ws["C1"] = "Mitarbeiter:"
        ws["D1"] = "SHIFTED_MA"
        ws["C2"] = "Befunddatum:"
        ws["D2"] = "2024"
        ws["A10"] = "IND1"; ws["B10"] = "IND2"; ws["C10"] = "DokDatum"
        ws["D10"] = "Leistung"; ws["E10"] = "Leist-Kurz"
        ws.cell(row=11, column=1).value = 1
        ws.cell(row=11, column=2).value = 1
        ws.cell(row=11, column=3).value = "15.03.2024"
        ws.cell(row=11, column=4).value = "L"
        ws.cell(row=11, column=5).value = "MRT"

        result = load_statistics(save_wb(wb, tmp_path / "s.xlsx"))
        assert result["mitarbeiter"] == "SHIFTED_MA"
        assert result["befunddatum"] == "2024"


# ── Map: field values and section assignments ──────────────────────────────────

class TestMapFieldValues:
    """Every value in the map dict must exactly match the file content."""

    def test_leist0_codes_read_correctly(self, tmp_path):
        """Leist0 codes are read back exactly as written."""
        codes = ["MRT", "CT001", "US_AB", "DUMMY99"]
        wb = make_map_wb([(c, f"Desc {c}", {"DL Gefäße": 1, "MRT": 0, "CT": 0})
                          for c in codes])
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))

        for code in codes:
            assert code in result["leistungen"], f"Code {code!r} missing from map"

    def test_section_value_1_read_as_1(self, tmp_path):
        """Section values of 1 are read as integer 1."""
        wb = make_map_wb(
            [("MRT", "MR", {"DL Gefäße": 1, "MRT": 1, "CT": 0})]
        )
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))

        entry = result["leistungen"]["MRT"]["sections"]
        assert entry["DL Gefäße"] == 1
        assert entry["MRT"] == 1

    def test_section_value_0_read_as_0(self, tmp_path):
        """Section values of 0 (or blank) are read as integer 0."""
        wb = make_map_wb(
            [("MRT", "MR", {"DL Gefäße": 0, "MRT": 1, "CT": 0})]
        )
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))

        entry = result["leistungen"]["MRT"]["sections"]
        assert entry["DL Gefäße"] == 0
        assert entry["CT"] == 0

    def test_section_names_read_correctly(self, tmp_path):
        """Section column headers are read back exactly."""
        sections = ["DL Gefäße", "MRT gesamt", "CT Hals/Thorax", "US gesamt"]
        wb = make_map_wb(
            [("MRT", "MR", {s: 1 for s in sections})],
            sections=sections,
        )
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))

        for sec in sections:
            assert sec in result["sections"], f"Section {sec!r} missing"

    def test_section_order_preserved(self, tmp_path):
        """Sections appear in the same order as in the file."""
        sections = ["DL Gefäße", "Alpha", "Beta", "Gamma", "Delta"]
        wb = make_map_wb(
            [("MRT", "MR", {s: 1 for s in sections})],
            sections=sections,
        )
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))
        assert result["sections"] == sections

    def test_multiple_codes_all_loaded(self, tmp_path):
        """All leistung codes in the file appear in the map dict."""
        codes = [f"CODE{i}" for i in range(20)]
        wb = make_map_wb(
            [(c, f"D{c}", {"DL Gefäße": 1, "MRT": 0, "CT": 0}) for c in codes]
        )
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))

        for code in codes:
            assert code in result["leistungen"], f"Code {code!r} missing from map"
        assert len(result["leistungen"]) == len(codes)

    def test_kurztext_read_correctly(self, tmp_path):
        """Kurztext (column C) is read back correctly."""
        wb = make_map_wb(
            [("MRT", "MR Tomographie Schädel", {"DL Gefäße": 1, "MRT": 1, "CT": 0})]
        )
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))
        assert result["leistungen"]["MRT"]["kurztext"] == "MR Tomographie Schädel"

    def test_multi_value_map_cells_preserved(self, tmp_path):
        """Values > 1 in section cells are stored as-is (multi-region billing weight)."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=2).value = "Leist0"
        ws.cell(row=1, column=6).value = "DL Gefäße"
        ws.cell(row=1, column=7).value = "MRT"
        ws.cell(row=2, column=2).value = "MRT"
        ws.cell(row=2, column=6).value = 2    # val=2 → counts as 2 (multi-region billing)
        ws.cell(row=2, column=7).value = 1

        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))
        assert result["leistungen"]["MRT"]["sections"]["DL Gefäße"] == 2
        assert result["leistungen"]["MRT"]["sections"]["MRT"] == 1

    def test_blank_map_cell_treated_as_zero(self, tmp_path):
        """Blank section cells are treated as 0."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=2).value = "Leist0"
        ws.cell(row=1, column=6).value = "DL Gefäße"
        ws.cell(row=1, column=7).value = "MRT"
        ws.cell(row=2, column=2).value = "CT"
        # Both section cells left blank (None)
        ws.cell(row=2, column=6).value = None
        ws.cell(row=2, column=7).value = None

        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))
        assert result["leistungen"]["CT"]["sections"]["DL Gefäße"] == 0
        assert result["leistungen"]["CT"]["sections"]["MRT"] == 0


# ── Reference amounts: value round-trip ───────────────────────────────────────

class TestReferenceAmountsValues:
    """Values in reference_amount.xlsx must be read back exactly."""

    def test_amounts_read_correctly(self, tmp_path):
        """Reference amounts are read back as the exact integers written."""
        from input_handler import load_reference_amounts
        from .helpers import make_reference_wb

        amounts = {"MRT": 3000, "CT Skelett": 4000, "US gesamt": 500}
        wb = make_reference_wb(amounts)
        result = load_reference_amounts(save_wb(wb, tmp_path / "ref.xlsx"))

        for name, expected in amounts.items():
            assert result[name] == expected, (
                f"Amount for {name!r}: expected {expected}, got {result.get(name)}"
            )

    def test_section_names_read_correctly(self, tmp_path):
        """Section names in reference file match exactly (case-sensitive)."""
        from input_handler import load_reference_amounts
        from .helpers import make_reference_wb

        names = ["MRT gesamt", "CT Hals/Thorax", "DL Gefäße", "US Abdomen"]
        amounts = {n: 1000 for n in names}
        wb = make_reference_wb(amounts)
        result = load_reference_amounts(save_wb(wb, tmp_path / "ref.xlsx"))

        for name in names:
            assert name in result, f"Name {name!r} missing from reference result"

    def test_float_amounts_truncated_to_int(self, tmp_path):
        """Float values (e.g. 3000.7) are truncated to int without error."""
        from input_handler import load_reference_amounts

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=1).value = "Section"
        ws.cell(row=1, column=2).value = "Betrag"
        ws.cell(row=2, column=1).value = "MRT"
        ws.cell(row=2, column=2).value = 3000.7

        result = load_reference_amounts(save_wb(wb, tmp_path / "ref.xlsx"))
        assert result["MRT"] == 3000

    def test_all_rows_read_not_just_first_or_last(self, tmp_path):
        """All rows in the reference file are loaded, not just first or last."""
        from input_handler import load_reference_amounts
        from .helpers import make_reference_wb

        amounts = {f"Section{i}": i * 100 for i in range(1, 21)}
        wb = make_reference_wb(amounts)
        result = load_reference_amounts(save_wb(wb, tmp_path / "ref.xlsx"))

        for name, expected in amounts.items():
            assert result.get(name) == expected, (
                f"Row for {name!r} not loaded correctly: "
                f"expected {expected}, got {result.get(name)}"
            )
