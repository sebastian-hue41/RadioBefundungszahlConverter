"""Tests for core/output_handler.py — xlsx generation, styling, and in-memory build."""

import io
from pathlib import Path

import openpyxl
import pytest

from core.output_handler import build_xlsx_bytes, save_output
from .helpers import make_result


class TestSaveOutput:
    """save_output — disk-based xlsx generation (reused from main)."""

    def test_save_output_creates_file(self, tmp_path):
        result = make_result(["MRT"], {"MRT": {2024: 10}}, [2024])
        path = save_output(result, output_dir=str(tmp_path))
        assert Path(path).exists()
        assert Path(path).parent == tmp_path

    def test_save_output_filename_from_metadata(self, tmp_path):
        result = make_result(
            ["MRT"], {"MRT": {2024: 10}}, [2024],
            mitarbeiter="KUNZ_A", befunddatum="01012024-31122024",
        )
        path = save_output(result, output_dir=str(tmp_path))
        filename = Path(path).name
        assert "KUNZ_A" in filename
        assert "01012024" in filename
        assert "Auswertung" in filename

    def test_save_output_filename_generic_when_missing(self, tmp_path):
        result = make_result(
            ["MRT"], {"MRT": {2024: 10}}, [2024],
            mitarbeiter=None, befunddatum=None,
        )
        path = save_output(result, output_dir=str(tmp_path))
        filename = Path(path).name
        assert "Auswertung_" in filename
        assert filename.endswith(".xlsx")

    def test_save_output_year_columns_present(self, tmp_path):
        result = make_result(["MRT"], {"MRT": {2024: 5, 2025: 3}}, [2024, 2025])
        path = save_output(result, output_dir=str(tmp_path))
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        assert "2024" in header
        assert "2025" in header
        assert "Gesamt" in header

    def test_save_output_gesamt_column_correct_totals(self, tmp_path):
        result = make_result(["MRT"], {"MRT": {2024: 5, 2025: 3}}, [2024, 2025])
        path = save_output(result, output_dir=str(tmp_path))
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        gesamt_idx = header.index("Gesamt") + 1
        assert ws.cell(2, gesamt_idx).value == 8

    def test_save_output_without_reference_no_benoetigt_column(self, tmp_path):
        result = make_result(["MRT"], {"MRT": {2024: 10}}, [2024])
        path = save_output(result, output_dir=str(tmp_path), reference_amounts=None)
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        assert "Benötigt" not in header

    def test_save_output_with_reference_has_benoetigt_column(self, tmp_path):
        result = make_result(["MRT"], {"MRT": {2024: 10}}, [2024])
        path = save_output(result, output_dir=str(tmp_path), reference_amounts={"MRT": 3000})
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        assert "Benötigt" in header

    def test_save_output_benoetigt_values_written(self, tmp_path):
        result = make_result(
            ["MRT", "CT"],
            {"MRT": {2024: 10}, "CT": {2024: 5}},
            [2024],
        )
        path = save_output(result, output_dir=str(tmp_path), reference_amounts={"MRT": 3000, "CT": 4000})
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        benoetigt_idx = header.index("Benötigt") + 1
        assert ws.cell(2, benoetigt_idx).value == 3000
        assert ws.cell(3, benoetigt_idx).value == 4000

    def test_save_output_reference_lookup_case_insensitive(self, tmp_path):
        result = make_result(["MRT"], {"MRT": {2024: 10}}, [2024])
        path = save_output(result, output_dir=str(tmp_path), reference_amounts={"mrt": 3000})
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        benoetigt_idx = header.index("Benötigt") + 1
        assert ws.cell(2, benoetigt_idx).value == 3000

    def test_save_output_no_cf_without_reference(self, tmp_path):
        result = make_result(["MRT"], {"MRT": {2024: 10}}, [2024])
        path = save_output(result, output_dir=str(tmp_path), reference_amounts=None)
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        assert len(list(ws.conditional_formatting)) == 0

    def test_save_output_cf_rules_present_with_reference(self, tmp_path):
        result = make_result(
            ["MRT", "CT"],
            {"MRT": {2024: 10}, "CT": {2024: 5}},
            [2024],
        )
        path = save_output(result, output_dir=str(tmp_path), reference_amounts={"MRT": 3000, "CT": 4000})
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        assert len(list(ws.conditional_formatting)) > 0

    def test_save_output_footer_count_summary(self, tmp_path):
        result = make_result(["MRT"], {"MRT": {2024: 10}}, [2024], total_counted=10)
        path = save_output(result, output_dir=str(tmp_path))
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        found = any(
            "Gezählte" in str(v or "")
            for row in ws.iter_rows(values_only=True)
            for v in row
        )
        assert found

    def test_save_output_footer_mismatch_warning(self, tmp_path):
        result = make_result(
            ["MRT"], {"MRT": {2024: 10}}, [2024],
            total_counted=10, total_leistungen=15,
        )
        path = save_output(result, output_dir=str(tmp_path))
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        found = any(
            "ABWEICHUNG" in str(v or "")
            for row in ws.iter_rows(values_only=True)
            for v in row
        )
        assert found

    def test_save_output_errors_section(self, tmp_path):
        result = make_result(
            ["MRT"], {"MRT": {2024: 10}}, [2024],
            errors=["WARNING: Test error 1", "ERROR: Test error 2"],
        )
        path = save_output(result, output_dir=str(tmp_path))
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        full_text = " ".join(
            str(v or "") for row in ws.iter_rows(values_only=True) for v in row
        )
        assert "Test error 1" in full_text
        assert "Test error 2" in full_text

    def test_save_output_merged_duplicates_notice(self, tmp_path):
        result = make_result(
            ["MRT"], {"MRT": {2024: 10}}, [2024],
            merged_duplicates=["CODE1", "CODE2"],
        )
        path = save_output(result, output_dir=str(tmp_path))
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        full_text = " ".join(
            str(v or "") for row in ws.iter_rows(values_only=True) for v in row
        )
        assert "CODE1" in full_text
        assert "CODE2" in full_text

    def test_save_output_multiple_sections(self, tmp_path):
        result = make_result(
            ["MRT", "CT", "US"],
            {"MRT": {2024: 100, 2025: 50}, "CT": {2024: 200, 2025: 150}, "US": {2024: 300, 2025: 250}},
            [2024, 2025],
        )
        path = save_output(result, output_dir=str(tmp_path))
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        first_col = [ws.cell(r, 1).value for r in range(1, 10)]
        assert "MRT" in first_col
        assert "CT" in first_col
        assert "US" in first_col

    def test_save_output_zero_counts_displayed(self, tmp_path):
        result = make_result(
            ["MRT", "CT"],
            {"MRT": {2024: 10}, "CT": {2024: 0}},
            [2024],
        )
        path = save_output(result, output_dir=str(tmp_path))
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        assert ws.cell(3, 3).value == 0

    def test_save_output_handles_missing_reference_for_section(self, tmp_path):
        result = make_result(
            ["MRT", "CT"],
            {"MRT": {2024: 10}, "CT": {2024: 5}},
            [2024],
        )
        path = save_output(result, output_dir=str(tmp_path), reference_amounts={"MRT": 3000})
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        benoetigt_idx = header.index("Benötigt") + 1
        assert ws.cell(2, benoetigt_idx).value == 3000
        # CT has no reference — cell should be blank (None)
        assert ws.cell(3, benoetigt_idx).value is None


# ── TestBuildXlsxBytes (web-specific) ─────────────────────────────────────────

class TestBuildXlsxBytes:
    """build_xlsx_bytes — in-memory xlsx generation for the web worker."""

    def test_returns_filename_and_bytes(self):
        result = make_result(["MRT"], {"MRT": {2024: 10}}, [2024])
        filename, data = build_xlsx_bytes(result)
        assert isinstance(filename, str)
        assert isinstance(data, bytes)
        assert filename.endswith(".xlsx")

    def test_bytes_are_valid_xlsx(self):
        result = make_result(["MRT"], {"MRT": {2024: 10}}, [2024])
        _, data = build_xlsx_bytes(result)
        wb = openpyxl.load_workbook(io.BytesIO(data))
        assert wb.active is not None

    def test_filename_uses_metadata(self):
        result = make_result(
            ["MRT"], {"MRT": {2024: 10}}, [2024],
            mitarbeiter="KUNZ_A", befunddatum="01012024-31122024",
        )
        filename, _ = build_xlsx_bytes(result)
        assert "KUNZ_A" in filename
        assert "Auswertung" in filename

    def test_filename_generic_when_no_metadata(self):
        result = make_result(
            ["MRT"], {"MRT": {2024: 10}}, [2024],
            mitarbeiter=None, befunddatum=None,
        )
        filename, _ = build_xlsx_bytes(result)
        assert "Auswertung_" in filename

    def test_with_reference_has_benoetigt_column(self):
        result = make_result(["MRT"], {"MRT": {2024: 10}}, [2024])
        _, data = build_xlsx_bytes(result, reference_amounts={"MRT": 3000})
        wb = openpyxl.load_workbook(io.BytesIO(data))
        ws = wb.active
        header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        assert "Benötigt" in header

    def test_content_matches_save_output(self, tmp_path):
        """build_xlsx_bytes and save_output produce equivalent cell content."""
        result = make_result(
            ["MRT", "CT"],
            {"MRT": {2024: 100, 2025: 50}, "CT": {2024: 200, 2025: 150}},
            [2024, 2025],
            mitarbeiter="DOC", befunddatum="2024",
        )
        ref = {"MRT": 3000, "CT": 4000}

        _, data = build_xlsx_bytes(result, reference_amounts=ref)
        disk_path = save_output(result, output_dir=str(tmp_path), reference_amounts=ref)

        wb_mem  = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        wb_disk = openpyxl.load_workbook(disk_path, data_only=True)

        ws_mem  = wb_mem.active
        ws_disk = wb_disk.active

        # Compare all data rows cell by cell
        for r in range(1, len(result["data"]) + 3):
            for c in range(1, ws_mem.max_column + 1):
                assert ws_mem.cell(r, c).value == ws_disk.cell(r, c).value, \
                    f"Mismatch at row={r} col={c}"

    def test_no_disk_write(self, tmp_path, monkeypatch):
        """build_xlsx_bytes must not write to disk."""
        written = []
        original_save = openpyxl.Workbook.save

        def _spy_save(self, filename):
            if not isinstance(filename, io.BytesIO):
                written.append(filename)
            return original_save(self, filename)

        monkeypatch.setattr(openpyxl.Workbook, "save", _spy_save)

        result = make_result(["MRT"], {"MRT": {2024: 10}}, [2024])
        build_xlsx_bytes(result)
        assert written == [], f"build_xlsx_bytes wrote to disk: {written}"
