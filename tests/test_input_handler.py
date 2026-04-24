"""Tests for input_handler.py — file validation and data loading."""

import io
import zipfile
from pathlib import Path

import openpyxl
import pytest

from input_handler import (
    FileValidationError,
    MapValidationError,
    StatsValidationError,
    load_map,
    load_reference_amounts,
    load_statistics,
    validate_xlsx_file,
    MAX_FILE_SIZE_MB,
)
from .helpers import (
    make_corrupt_xlsx_bytes,
    make_map_wb,
    make_reference_wb,
    make_stats_wb,
    make_zip_without_workbook_xml,
    inject_vba_into_xlsx,
    save_wb,
)


class TestValidateXlsxFile:
    """Security validation tests."""

    def test_nonexistent_file(self):
        """Raises FileValidationError for missing file."""
        with pytest.raises(FileValidationError, match="not found"):
            validate_xlsx_file("/nonexistent/path/file.xlsx")

    def test_wrong_extension(self, tmp_path):
        """Raises FileValidationError for non-.xlsx extension."""
        f = tmp_path / "test.txt"
        f.write_text("content")
        with pytest.raises(FileValidationError, match="Unsupported file type"):
            validate_xlsx_file(str(f))

    def test_empty_file(self, tmp_path):
        """Raises FileValidationError for 0-byte file."""
        f = tmp_path / "empty.xlsx"
        f.write_bytes(b"")
        with pytest.raises(FileValidationError, match="empty"):
            validate_xlsx_file(str(f))

    def test_file_too_large(self, tmp_path, monkeypatch):
        """Raises FileValidationError when file > 50MB."""
        f = tmp_path / "big.xlsx"
        # Create a sparse file of 51MB (fast on most filesystems)
        with open(f, "wb") as fh:
            fh.seek(51 * 1024 * 1024 - 1)
            fh.write(b"\x00")
        with pytest.raises(FileValidationError, match="too large"):
            validate_xlsx_file(str(f))

    def test_corrupt_zip(self, tmp_path):
        """Raises FileValidationError for corrupt zip (not a valid xlsx)."""
        f = tmp_path / "corrupt.xlsx"
        f.write_bytes(make_corrupt_xlsx_bytes())
        with pytest.raises(FileValidationError, match="not a valid xlsx archive"):
            validate_xlsx_file(str(f))

    def test_missing_workbook_xml(self, tmp_path):
        """Raises FileValidationError when xl/workbook.xml is missing."""
        f = tmp_path / "no_workbook.xlsx"
        f.write_bytes(make_zip_without_workbook_xml())
        with pytest.raises(FileValidationError, match="xl/workbook.xml missing"):
            validate_xlsx_file(str(f))

    def test_vba_macros_rejected(self, tmp_path):
        """Raises FileValidationError when VBA macros are present."""
        wb = openpyxl.Workbook()
        buf = io.BytesIO()
        wb.save(buf)
        f = tmp_path / "vba.xlsx"
        f.write_bytes(inject_vba_into_xlsx(buf.getvalue()))
        with pytest.raises(FileValidationError, match="VBA macros"):
            validate_xlsx_file(str(f))

    def test_valid_xlsx(self, tmp_path):
        """Returns Path for valid .xlsx file."""
        wb = openpyxl.Workbook()
        f = tmp_path / "valid.xlsx"
        save_wb(wb, f)
        result = validate_xlsx_file(str(f))
        assert isinstance(result, Path)
        assert result.exists()


class TestLoadStatistics:
    """Tests for statistics file loading and IND2 algorithm."""

    def test_load_statistics_valid(self, tmp_path):
        """Load a valid statistics file."""
        rows = [
            (1, 1, "15.03.2024", "MR Kopf", "MRT"),
            (1, 2, None, "CT Thorax", "CT"),
        ]
        wb = make_stats_wb(rows, mitarbeiter="KUNZ_A", befunddatum="01.01.2024")
        path = save_wb(wb, tmp_path / "stats.xlsx")

        result = load_statistics(path)
        assert result["mitarbeiter"] == "KUNZ_A"
        assert result["befunddatum"] == "01.01.2024"
        assert len(result["leistungen"]) == 2
        assert result["total_leistungen"] == 2
        assert result["leistungen"][0]["year"] == 2024

    def test_load_statistics_missing_leist_kurz_column(self, tmp_path):
        """Raises StatsValidationError when Leist-Kurz column is missing."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A10"] = "IND1"
        ws["B10"] = "IND2"
        ws["C10"] = "DokDatum"
        # Missing Leist-Kurz
        path = save_wb(wb, tmp_path / "bad_header.xlsx")

        with pytest.raises(StatsValidationError, match="Leist-Kurz"):
            load_statistics(path)

    def test_load_statistics_no_data_rows(self, tmp_path):
        """Raises StatsValidationError when no valid data rows exist."""
        wb = make_stats_wb([], mitarbeiter="TEST")  # Empty rows
        path = save_wb(wb, tmp_path / "no_data.xlsx")

        with pytest.raises(StatsValidationError, match="No valid examination rows"):
            load_statistics(path)

    def test_load_statistics_missing_mitarbeiter(self, tmp_path):
        """Missing Mitarbeiter metadata — warning added, processing continues."""
        rows = [(1, 1, "15.03.2024", "Proc", "MRT")]
        wb = make_stats_wb(rows)
        ws = wb.active
        ws.delete_rows(1, 1)  # Remove mitarbeiter row
        path = save_wb(wb, tmp_path / "no_ma.xlsx")

        result = load_statistics(path)
        assert result["mitarbeiter"] is None
        assert any("Mitarbeiter" in e for e in result["errors"])

    def test_load_statistics_missing_befunddatum(self, tmp_path):
        """Missing Befunddatum metadata — warning added, processing continues."""
        rows = [(1, 1, "15.03.2024", "Proc", "MRT")]
        wb = make_stats_wb(rows, befunddatum=None)
        path = save_wb(wb, tmp_path / "no_bd.xlsx")

        result = load_statistics(path)
        assert result["befunddatum"] is None
        assert any("Befunddatum" in e for e in result["errors"])

    def test_load_statistics_ind2_algorithm_simple(self, tmp_path):
        """IND2 algorithm with single group [1,2] → total=2."""
        rows = [
            (1, 1, "15.03.2024", "MR", "MRT"),
            (1, 2, None, "CT", "CT"),  # IND2=2 signals end of group
        ]
        wb = make_stats_wb(rows)
        path = save_wb(wb, tmp_path / "stats.xlsx")

        result = load_statistics(path)
        assert result["total_leistungen"] == 2
        assert len(result["leistungen"]) == 2

    def test_load_statistics_ind2_algorithm_multiple_groups(self, tmp_path):
        """IND2 algorithm with multiple groups [1,2,1,2,3] → total=5."""
        rows = [
            (1, 1, "15.03.2024", "MR", "MRT"),
            (1, 2, None, "CT", "CT"),
            (2, 1, "20.03.2024", "US", "US"),
            (2, 2, None, "MR", "MRT"),
            (2, 3, None, "CT", "CT"),
        ]
        wb = make_stats_wb(rows)
        path = save_wb(wb, tmp_path / "stats.xlsx")

        result = load_statistics(path)
        assert result["total_leistungen"] == 5
        assert len(result["leistungen"]) == 5

    def test_load_statistics_ind2_algorithm_single_item(self, tmp_path):
        """IND2 algorithm with single item [1] → total=1."""
        rows = [(1, 1, "15.03.2024", "MR", "MRT")]
        wb = make_stats_wb(rows)
        path = save_wb(wb, tmp_path / "stats.xlsx")

        result = load_statistics(path)
        assert result["total_leistungen"] == 1

    def test_load_statistics_ind2_algorithm_descending_boundaries(self, tmp_path):
        """IND2 algorithm with [3,1,2] → total=5."""
        rows = [
            (1, 3, "15.03.2024", "MR", "MRT"),  # Next is 1 ≤ 3 → end, add 3
            (2, 1, "20.03.2024", "CT", "CT"),   # Next is 2 > 1 → continue
            (2, 2, None, "US", "US"),           # Last → end, add 2
        ]
        wb = make_stats_wb(rows)
        path = save_wb(wb, tmp_path / "stats.xlsx")

        result = load_statistics(path)
        assert result["total_leistungen"] == 5

    def test_load_statistics_ind2_algorithm_identical_sequence(self, tmp_path):
        """IND2 algorithm with [1,1,1] → total=3."""
        rows = [
            (1, 1, "15.03.2024", "MR", "MRT"),
            (2, 1, "15.03.2024", "CT", "CT"),
            (3, 1, "15.03.2024", "US", "US"),
        ]
        wb = make_stats_wb(rows)
        path = save_wb(wb, tmp_path / "stats.xlsx")

        result = load_statistics(path)
        assert result["total_leistungen"] == 3

    def test_load_statistics_date_carryforward(self, tmp_path):
        """Date carries forward from first row of group to subsequent rows."""
        rows = [
            (1, 1, "15.03.2024", "MR", "MRT"),
            (1, 2, None, "CT", "CT"),
        ]
        wb = make_stats_wb(rows)
        path = save_wb(wb, tmp_path / "stats.xlsx")

        result = load_statistics(path)
        # Both rows should have year 2024 (second row inherited date)
        assert result["leistungen"][0]["year"] == 2024
        assert result["leistungen"][1]["year"] == 2024

    def test_load_statistics_bad_date_skipped(self, tmp_path):
        """Row with unparseable date is skipped with warning."""
        rows = [
            (1, 1, "not-a-date", "MR", "MRT"),
            (2, 1, "15.03.2024", "CT", "CT"),
        ]
        wb = make_stats_wb(rows)
        path = save_wb(wb, tmp_path / "stats.xlsx")

        result = load_statistics(path)
        assert len(result["leistungen"]) == 1  # First row skipped
        assert result["leistungen"][0]["leist_kurz"] == "CT"
        assert any("cannot parse year" in e for e in result["errors"])

    def test_load_statistics_null_leist_kurz_skipped(self, tmp_path):
        """Row with missing Leist-Kurz is skipped with warning."""
        rows = [
            (1, 1, "15.03.2024", "MR", None),
            (2, 1, "15.03.2024", "CT", "CT"),
        ]
        wb = make_stats_wb(rows)
        path = save_wb(wb, tmp_path / "stats.xlsx")

        result = load_statistics(path)
        assert len(result["leistungen"]) == 1
        assert any("Leist-Kurz" in e and "empty" in e for e in result["errors"])

    def test_load_statistics_formula_cell_neutralized(self, tmp_path):
        """Formula in Leist-Kurz reads as None with data_only=True (formulas neutralized)."""
        # openpyxl data_only=True returns None for formula cells — row is skipped.
        # A single-row file with only a formula row → no valid rows → StatsValidationError.
        rows = [(1, 1, "15.03.2024", "MR", "=MALWARE()")]
        wb = make_stats_wb(rows)
        path = save_wb(wb, tmp_path / "stats.xlsx")

        # Formula → None → row skipped → no valid rows → StatsValidationError
        with pytest.raises(StatsValidationError, match="No valid examination rows"):
            load_statistics(path)

    def test_load_statistics_count_mismatch_detection(self, tmp_path):
        """load_statistics reports discrepancy; COUNT MISMATCH is emitted by logic.process()."""
        rows = [
            (1, 1, "15.03.2024", "MR", "MRT"),
            (1, 2, "99.99.9999", "CT", "CT"),  # Unparseable date → row skipped
        ]
        wb = make_stats_wb(rows, total_l8=None)
        path = save_wb(wb, tmp_path / "stats.xlsx")

        result = load_statistics(path)
        # IND2 sequence [1, 2] → total_leistungen = 2
        assert result["total_leistungen"] == 2
        # Only one row processed (second skipped due to bad date)
        assert len(result["leistungen"]) == 1
        # input_handler adds a "cannot parse year" warning (not COUNT MISMATCH)
        assert any("cannot parse year" in e for e in result["errors"])

    def test_load_statistics_multiple_years(self, tmp_path):
        """Examinations from multiple years are correctly parsed."""
        rows = [
            (1, 1, "15.03.2024", "MR", "MRT"),
            (2, 1, "20.06.2025", "CT", "CT"),
        ]
        wb = make_stats_wb(rows)
        path = save_wb(wb, tmp_path / "stats.xlsx")

        result = load_statistics(path)
        assert result["leistungen"][0]["year"] == 2024
        assert result["leistungen"][1]["year"] == 2025


class TestLoadMap:
    """Tests for map file loading."""

    def test_load_map_valid(self, tmp_path):
        """Load a valid map file."""
        codes = [
            ("MRT", "MR Tomographie", {"DL Gefäße": 1, "MRT": 1, "CT": 0}),
            ("CT", "CT Scan", {"DL Gefäße": 0, "MRT": 0, "CT": 1}),
        ]
        wb = make_map_wb(codes)
        path = save_wb(wb, tmp_path / "map.xlsx")

        result = load_map(path)
        assert "MRT" in result["leistungen"]
        assert result["leistungen"]["MRT"]["sections"]["MRT"] == 1
        assert "DL Gefäße" in result["sections"]

    def test_load_map_b1_wrong(self, tmp_path):
        """Raises MapValidationError when B1 ≠ 'Leist0'."""
        codes = [("MRT", "desc", {"DL Gefäße": 1})]
        wb = make_map_wb(codes, b1="WrongHeader")
        path = save_wb(wb, tmp_path / "map.xlsx")

        with pytest.raises(MapValidationError, match="B1"):
            load_map(path)

    def test_load_map_f1_wrong(self, tmp_path):
        """Raises MapValidationError when F1 ≠ 'DL Gefäße'."""
        codes = [("MRT", "desc", {"WrongName": 1})]
        wb = make_map_wb(codes, sections=["WrongName", "MRT"])
        path = save_wb(wb, tmp_path / "map.xlsx")

        with pytest.raises(MapValidationError, match="F1"):
            load_map(path)

    def test_load_map_no_sections(self, tmp_path):
        """Raises MapValidationError when no sections found after F."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=2).value = "Leist0"
        ws.cell(row=1, column=6).value = "DL Gefäße"
        # Add underscore columns only (excluded by default → no usable sections)
        ws.cell(row=1, column=7).value = "_CLIP"
        ws.cell(row=1, column=8).value = "_TRENNER"
        ws.cell(row=2, column=2).value = "MRT"
        path = save_wb(wb, tmp_path / "map.xlsx")

        # DL Gefäße itself is valid, but no non-underscore sections after exclusion
        # The map would need _only_ underscore sections (besides DL Gefäße) to trigger this.
        # Easier: use include_underscore_columns=False and only have underscore columns.
        wb2 = openpyxl.Workbook()
        ws2 = wb2.active
        ws2.cell(row=1, column=2).value = "Leist0"
        ws2.cell(row=1, column=6).value = "_ONLY_UNDERSCORE"
        ws2.cell(row=2, column=2).value = "MRT"
        path2 = save_wb(wb2, tmp_path / "map2.xlsx")

        with pytest.raises(MapValidationError):
            load_map(path2, include_underscore_columns=False)

    def test_load_map_duplicate_no_values(self, tmp_path):
        """Duplicate leist0 with no section values keeps first."""
        codes = [
            ("MRT", "desc1", {"DL Gefäße": 0, "MRT": 0}),
            ("MRT", "desc2", {"DL Gefäße": 0, "MRT": 0}),
        ]
        wb = make_map_wb(codes)
        path = save_wb(wb, tmp_path / "map.xlsx")

        result = load_map(path)
        assert len(result["merged_duplicates"]) == 0  # Not considered a conflict

    def test_load_map_duplicate_one_has_values(self, tmp_path):
        """Duplicate leist0 where only one has values — that one is used."""
        codes = [
            ("MRT", "no_values", {"DL Gefäße": 0, "MRT": 0}),
            ("MRT", "has_values", {"DL Gefäße": 0, "MRT": 1}),
        ]
        wb = make_map_wb(codes)
        path = save_wb(wb, tmp_path / "map.xlsx")

        result = load_map(path)
        assert result["leistungen"]["MRT"]["sections"]["MRT"] == 1
        assert len(result["merged_duplicates"]) == 0

    def test_load_map_duplicate_conflicting_merged(self, tmp_path):
        """Duplicate leist0 with conflicting values — merged by max."""
        codes = [
            ("MRT", "v1", {"DL Gefäße": 1, "MRT": 1, "CT": 0}),
            ("MRT", "v2", {"DL Gefäße": 0, "MRT": 0, "CT": 1}),
        ]
        wb = make_map_wb(codes)
        path = save_wb(wb, tmp_path / "map.xlsx")

        result = load_map(path)
        entry = result["leistungen"]["MRT"]
        assert entry["sections"]["DL Gefäße"] == 1  # max(1, 0)
        assert entry["sections"]["MRT"] == 1        # max(1, 0)
        assert entry["sections"]["CT"] == 1         # max(0, 1)
        assert "MRT" in result["merged_duplicates"]

    def test_load_map_duplicate_identical(self, tmp_path):
        """Duplicate leist0 with identical assignments — no merge warning."""
        codes = [
            ("MRT", "v1", {"DL Gefäße": 1, "MRT": 1}),
            ("MRT", "v2", {"DL Gefäße": 1, "MRT": 1}),
        ]
        wb = make_map_wb(codes, sections=["DL Gefäße", "MRT"])
        path = save_wb(wb, tmp_path / "map.xlsx")

        result = load_map(path)
        assert len(result["merged_duplicates"]) == 0

    def test_load_map_underscore_columns_excluded(self, tmp_path):
        """Underscore-prefixed section columns excluded by default."""
        sections = ["DL Gefäße", "_CLIP", "MRT"]
        codes = [("MRT", "desc", {s: 1 for s in sections})]
        wb = make_map_wb(codes, sections=sections)
        path = save_wb(wb, tmp_path / "map.xlsx")

        result = load_map(path, include_underscore_columns=False)
        assert "_CLIP" not in result["sections"]
        assert "MRT" in result["sections"]

    def test_load_map_underscore_columns_included(self, tmp_path):
        """Underscore-prefixed sections included with flag."""
        sections = ["DL Gefäße", "_CLIP", "MRT"]
        codes = [("MRT", "desc", {s: 1 for s in sections})]
        wb = make_map_wb(codes, sections=sections)
        path = save_wb(wb, tmp_path / "map.xlsx")

        result = load_map(path, include_underscore_columns=True)
        assert "_CLIP" in result["sections"]

    def test_load_map_formula_injection_leist0(self, tmp_path):
        """Formula in Leist0 is treated as None (data_only=True neutralizes it)."""
        # Note: With data_only=True, formula cells read as None.
        # This is intentional security - prevents formula execution.
        codes = [("=MALWARE()", "desc", {"DL Gefäße": 1})]
        wb = make_map_wb(codes)
        path = save_wb(wb, tmp_path / "map.xlsx")

        # Formula becomes None, so no leistung entry exists
        with pytest.raises(MapValidationError, match="No leistung entries"):
            load_map(path)


class TestLoadReferenceAmounts:
    """Tests for reference amounts loading."""

    def test_load_reference_valid(self, tmp_path):
        """Load valid reference amounts."""
        amounts = {"MRT": 3000, "CT": 4000}
        wb = make_reference_wb(amounts)
        path = save_wb(wb, tmp_path / "ref.xlsx")

        result = load_reference_amounts(path)
        assert result["MRT"] == 3000
        assert result["CT"] == 4000

    def test_load_reference_nonexistent(self):
        """Nonexistent file returns {} with warning."""
        result = load_reference_amounts("/nonexistent/path.xlsx")
        assert result == {}

    def test_load_reference_header_detection_string(self, tmp_path):
        """First row with string in col B treated as header."""
        wb = make_reference_wb({"MRT": 3000}, include_header=True)
        path = save_wb(wb, tmp_path / "ref.xlsx")

        result = load_reference_amounts(path)
        assert result["MRT"] == 3000  # Header skipped

    def test_load_reference_header_detection_none(self, tmp_path):
        """First row with None in col B treated as header."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=1).value = "Leistungsbereich"
        ws.cell(row=1, column=2).value = None
        ws.cell(row=2, column=1).value = "MRT"
        ws.cell(row=2, column=2).value = 3000
        path = save_wb(wb, tmp_path / "ref.xlsx")

        result = load_reference_amounts(path)
        assert result["MRT"] == 3000

    def test_load_reference_none_amounts_skipped(self, tmp_path):
        """Rows with None amount are skipped."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=1).value = "Section"
        ws.cell(row=1, column=2).value = "Amt"
        ws.cell(row=2, column=1).value = "MRT"
        ws.cell(row=2, column=2).value = None
        ws.cell(row=3, column=1).value = "CT"
        ws.cell(row=3, column=2).value = 4000
        path = save_wb(wb, tmp_path / "ref.xlsx")

        result = load_reference_amounts(path)
        assert "MRT" not in result
        assert result["CT"] == 4000

    def test_load_reference_negative_skipped(self, tmp_path):
        """Negative amounts are skipped."""
        wb = make_reference_wb({})
        ws = wb.active
        ws.cell(row=2, column=1).value = "MRT"
        ws.cell(row=2, column=2).value = -1000
        path = save_wb(wb, tmp_path / "ref.xlsx")

        result = load_reference_amounts(path)
        assert result == {}

    def test_load_reference_too_large_skipped(self, tmp_path):
        """Amounts > 1,000,000 are skipped."""
        wb = make_reference_wb({})
        ws = wb.active
        ws.cell(row=2, column=1).value = "MRT"
        ws.cell(row=2, column=2).value = 2_000_000
        path = save_wb(wb, tmp_path / "ref.xlsx")

        result = load_reference_amounts(path)
        assert result == {}

    def test_load_reference_zero_skipped(self, tmp_path):
        """Zero amounts are skipped."""
        wb = make_reference_wb({})
        ws = wb.active
        ws.cell(row=2, column=1).value = "MRT"
        ws.cell(row=2, column=2).value = 0
        path = save_wb(wb, tmp_path / "ref.xlsx")

        result = load_reference_amounts(path)
        assert result == {}

    def test_load_reference_formula_injection_name_skipped(self, tmp_path):
        """Formula injection in section name skips that row, rest processed."""
        wb = make_reference_wb({})
        ws = wb.active
        ws.cell(row=2, column=1).value = "=MALWARE()"
        ws.cell(row=2, column=2).value = 3000
        ws.cell(row=3, column=1).value = "CT"
        ws.cell(row=3, column=2).value = 4000
        path = save_wb(wb, tmp_path / "ref.xlsx")

        result = load_reference_amounts(path)
        assert "=MALWARE()" not in result
        assert result["CT"] == 4000

    def test_load_reference_vba_file_rejected(self, tmp_path):
        """VBA file fails validation, returns {}."""
        wb = openpyxl.Workbook()
        buf = io.BytesIO()
        wb.save(buf)
        f = tmp_path / "vba_ref.xlsx"
        f.write_bytes(inject_vba_into_xlsx(buf.getvalue()))

        result = load_reference_amounts(str(f))
        assert result == {}

    def test_load_reference_corrupt_xlsx_returns_empty(self, tmp_path):
        """Corrupt xlsx file returns {}."""
        f = tmp_path / "corrupt_ref.xlsx"
        f.write_bytes(make_corrupt_xlsx_bytes())

        result = load_reference_amounts(str(f))
        assert result == {}
