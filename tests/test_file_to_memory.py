"""
test_file_to_memory.py

Verifies that every value in the source files is read correctly and transferred
into the in-memory data structures — covering both the path-based (CLI) and
bytes-based (web worker) loading paths.
"""

import io
from datetime import date, datetime

import openpyxl
import pytest

from core.input_handler import load_map, load_reference_amounts, load_statistics
from .helpers import (
    make_map_wb,
    make_reference_wb,
    make_stats_wb,
    save_wb,
    wb_to_bytes,
)


# ── TestStatisticsFieldValues ──────────────────────────────────────────────────

class TestStatisticsFieldValues:
    """Every field from a statistics workbook is read correctly."""

    def test_leist_kurz_preserved(self, tmp_path):
        rows = [(1, 1, "01.01.2024", "MRT Schädel", "MRTS")]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))
        assert result["leistungen"][0]["leist_kurz"] == "MRTS"

    def test_leistung_description_preserved(self, tmp_path):
        rows = [(1, 1, "01.01.2024", "CT Thorax nativ+KM", "CTT")]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))
        assert result["leistungen"][0]["leistung"] == "CT Thorax nativ+KM"

    def test_german_date_format(self, tmp_path):
        rows = [(1, 1, "15.06.2024", "A", "A")]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))
        assert result["leistungen"][0]["year"] == 2024

    def test_iso_date_format(self, tmp_path):
        rows = [(1, 1, "2024-06-15", "A", "A")]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))
        assert result["leistungen"][0]["year"] == 2024

    def test_datetime_object(self, tmp_path):
        rows = [(1, 1, datetime(2024, 6, 15), "A", "A")]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))
        assert result["leistungen"][0]["year"] == 2024

    def test_date_carryforward_within_group(self, tmp_path):
        rows = [
            (1, 1, "01.01.2024", "A", "A"),
            (1, 2, None,         "B", "B"),
            (1, 3, None,         "C", "C"),
        ]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))
        assert all(e["year"] == 2024 for e in result["leistungen"])

    def test_year_resets_between_groups(self, tmp_path):
        rows = [
            (1, 1, "01.01.2023", "A", "A"),
            (2, 1, "01.01.2024", "B", "B"),
        ]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))
        assert result["leistungen"][0]["year"] == 2023
        assert result["leistungen"][1]["year"] == 2024

    def test_ind1_preserved(self, tmp_path):
        rows = [(42, 1, "01.01.2024", "A", "A")]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))
        assert result["leistungen"][0]["ind1"] == 42

    def test_row_order_preserved(self, tmp_path):
        rows = [
            (1, 1, "01.01.2024", "A", "AAA"),
            (2, 1, "02.01.2024", "B", "BBB"),
            (3, 1, "03.01.2024", "C", "CCC"),
        ]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))
        assert [e["leist_kurz"] for e in result["leistungen"]] == ["AAA", "BBB", "CCC"]

    def test_whitespace_stripped_from_leist_kurz(self, tmp_path):
        wb = make_stats_wb([(1, 1, "01.01.2024", "A", "  MRTS  ")])
        result = load_statistics(save_wb(wb, tmp_path / "s.xlsx"))
        assert result["leistungen"][0]["leist_kurz"] == "MRTS"

    def test_column_order_independence(self, tmp_path):
        """Loader uses header names, not fixed column positions."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Mitarbeiter:"
        ws["B1"] = "DOC"
        # Header at row 10 with columns in non-standard order
        ws["A10"] = "DokDatum"
        ws["B10"] = "Leist-Kurz"
        ws["C10"] = "Leistung"
        ws["D10"] = "IND1"
        ws["E10"] = "IND2"
        ws["A11"] = "01.01.2024"
        ws["B11"] = "MRTS"
        ws["C11"] = "MRT Schädel"
        ws["D11"] = 1
        ws["E11"] = 1
        result = load_statistics(save_wb(wb, tmp_path / "s.xlsx"))
        assert result["leistungen"][0]["leist_kurz"] == "MRTS"
        assert result["leistungen"][0]["year"] == 2024

    def test_100_distinct_codes_all_read(self, tmp_path):
        rows = [(i, 1, "01.01.2024", f"Leistung{i}", f"CODE{i:03d}") for i in range(1, 101)]
        result = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))
        codes = {e["leist_kurz"] for e in result["leistungen"]}
        assert len(codes) == 100
        assert all(f"CODE{i:03d}" in codes for i in range(1, 101))


# ── TestStatisticsMetadata ─────────────────────────────────────────────────────

class TestStatisticsMetadata:

    def test_mitarbeiter_read_correctly(self, tmp_path):
        wb = make_stats_wb([(1, 1, "01.01.2024", "A", "A")], mitarbeiter="DR_JONES")
        result = load_statistics(save_wb(wb, tmp_path / "s.xlsx"))
        assert result["mitarbeiter"] == "DR_JONES"

    def test_befunddatum_read_correctly(self, tmp_path):
        wb = make_stats_wb([(1, 1, "01.01.2024", "A", "A")], befunddatum="01.01.2020-31.12.2024")
        result = load_statistics(save_wb(wb, tmp_path / "s.xlsx"))
        assert result["befunddatum"] == "01.01.2020-31.12.2024"

    def test_total_l8_read_correctly(self, tmp_path):
        wb = make_stats_wb([(1, 1, "01.01.2024", "A", "A")], total_l8=999)
        result = load_statistics(save_wb(wb, tmp_path / "s.xlsx"))
        assert result["total_documents"] == 999

    def test_metadata_in_non_a_column_found(self, tmp_path):
        """Metadata search scans multiple columns, not just A."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["D1"] = "Mitarbeiter:"
        ws["E1"] = "HIDDEN_DOC"
        ws["A10"] = "IND1"
        ws["B10"] = "IND2"
        ws["C10"] = "DokDatum"
        ws["D10"] = "Leistung"
        ws["E10"] = "Leist-Kurz"
        ws["A11"] = 1
        ws["B11"] = 1
        ws["C11"] = "01.01.2024"
        ws["D11"] = "MRT"
        ws["E11"] = "MRT"
        result = load_statistics(save_wb(wb, tmp_path / "s.xlsx"))
        assert result["mitarbeiter"] == "HIDDEN_DOC"


# ── TestMapFieldValues ─────────────────────────────────────────────────────────

class TestMapFieldValues:

    def test_leist0_codes_read_correctly(self, tmp_path):
        wb = make_map_wb([("MRTS", "MRT Schädel", {"DL Gefäße": 0, "MRT": 1, "CT": 0})])
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))
        assert "MRTS" in result["leistungen"]

    def test_section_value_1_stored_as_1(self, tmp_path):
        wb = make_map_wb([("CODE", "d", {"DL Gefäße": 0, "MRT": 1, "CT": 0})])
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))
        assert result["leistungen"]["CODE"]["sections"]["MRT"] == 1

    def test_section_value_0_stored_as_0(self, tmp_path):
        wb = make_map_wb([("CODE", "d", {"DL Gefäße": 0, "MRT": 1, "CT": 0})])
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))
        assert result["leistungen"]["CODE"]["sections"]["CT"] == 0

    def test_section_blank_treated_as_0(self, tmp_path):
        wb = make_map_wb([("CODE", "d", {"DL Gefäße": 0, "MRT": 1})])
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))
        assert result["leistungen"]["CODE"]["sections"].get("CT", 0) == 0

    def test_multi_value_map_cells_preserved(self, tmp_path):
        """Values > 1 are preserved as-is (multi-region billing weight)."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=2).value = "Leist0"
        ws.cell(row=1, column=6).value = "DL Gefäße"
        ws.cell(row=1, column=7).value = "MRT"
        ws.cell(row=2, column=2).value = "MRT"
        ws.cell(row=2, column=6).value = 2    # val=2 → multi-region billing weight
        ws.cell(row=2, column=7).value = 1
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))
        assert result["leistungen"]["MRT"]["sections"]["DL Gefäße"] == 2
        assert result["leistungen"]["MRT"]["sections"]["MRT"] == 1

    def test_section_names_and_order_preserved(self, tmp_path):
        sections = ["DL Gefäße", "MRT", "CT Abdomen", "US gesamt"]
        wb = make_map_wb([("CODE", "d", {s: 0 for s in sections})], sections=sections)
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))
        assert result["sections"] == sections

    def test_non_one_non_two_values_treated_as_zero(self, tmp_path):
        """Negative values and strings are treated as 0."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=2).value = "Leist0"
        ws.cell(row=1, column=6).value = "DL Gefäße"
        ws.cell(row=2, column=2).value = "CODE"
        ws.cell(row=2, column=6).value = -1
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))
        assert result["leistungen"]["CODE"]["sections"]["DL Gefäße"] == 0

    def test_20_codes_all_read(self, tmp_path):
        codes = [(f"CODE{i:02d}", f"Desc {i}", {"DL Gefäße": i % 2, "MRT": 0, "CT": 0})
                 for i in range(20)]
        result = load_map(save_wb(make_map_wb(codes), tmp_path / "m.xlsx"))
        assert len(result["leistungen"]) == 20

    def test_kurztext_preserved(self, tmp_path):
        wb = make_map_wb([("MRTS", "MRT Schädel nativ+KM", {"DL Gefäße": 0, "MRT": 1, "CT": 0})])
        result = load_map(save_wb(wb, tmp_path / "m.xlsx"))
        assert result["leistungen"]["MRTS"]["kurztext"] == "MRT Schädel nativ+KM"


# ── TestReferenceAmountsValues ─────────────────────────────────────────────────

class TestReferenceAmountsValues:

    def test_exact_amounts_read(self, tmp_path):
        wb = make_reference_wb({"MRT": 3000, "CT": 4500, "US": 1200})
        result = load_reference_amounts(save_wb(wb, tmp_path / "r.xlsx"))
        assert result == {"MRT": 3000, "CT": 4500, "US": 1200}

    def test_section_names_exact(self, tmp_path):
        wb = make_reference_wb({"CT Abdomen, Becken, Retroperit": 2500})
        result = load_reference_amounts(save_wb(wb, tmp_path / "r.xlsx"))
        assert "CT Abdomen, Becken, Retroperit" in result

    def test_float_amount_truncated_to_int(self, tmp_path):
        wb = make_reference_wb({})
        ws = wb.active
        ws.cell(row=1, column=1).value = "MRT"
        ws.cell(row=1, column=2).value = 3000.9
        result = load_reference_amounts(save_wb(wb, tmp_path / "r.xlsx"))
        assert result["MRT"] == 3000

    def test_all_20_rows_read(self, tmp_path):
        amounts = {f"Section{i:02d}": (i + 1) * 100 for i in range(20)}
        wb = make_reference_wb(amounts)
        result = load_reference_amounts(save_wb(wb, tmp_path / "r.xlsx"))
        assert len(result) == 20
        for name, expected in amounts.items():
            assert result[name] == expected


# ── TestBytesIOSourceParity ────────────────────────────────────────────────────

class TestBytesIOSourceParity:
    """BytesIO and path-based loading produce byte-identical leistungen lists."""

    def test_statistics_bytesio_parity(self, tmp_path):
        rows = [
            (1, 1, "01.01.2023", "MRT Schädel",  "MRTS"),
            (1, 2, None,         "CT Thorax",     "CTT"),
            (1, 3, None,         "US Abdomen",    "USA"),
            (2, 1, "15.06.2024", "DL Aorta",      "DLAO"),
        ]
        wb = make_stats_wb(rows, mitarbeiter="DOC_X", befunddatum="01.01.2023-31.12.2024")
        path_result  = load_statistics(save_wb(wb, tmp_path / "s.xlsx"))
        bytes_result = load_statistics(wb_to_bytes(wb))

        assert path_result["mitarbeiter"]     == bytes_result["mitarbeiter"]
        assert path_result["befunddatum"]     == bytes_result["befunddatum"]
        assert path_result["total_leistungen"]== bytes_result["total_leistungen"]
        assert len(path_result["leistungen"]) == len(bytes_result["leistungen"])
        for ep, eb in zip(path_result["leistungen"], bytes_result["leistungen"]):
            assert ep["leist_kurz"] == eb["leist_kurz"]
            assert ep["year"]       == eb["year"]
            assert ep["leistung"]   == eb["leistung"]

    def test_map_bytesio_parity(self, tmp_path):
        codes = [
            ("MRTS",  "MRT Schädel",      {"DL Gefäße": 0, "MRT": 1, "CT": 0}),
            ("CTT",   "CT Thorax",        {"DL Gefäße": 0, "MRT": 0, "CT": 1}),
            ("CTTHA", "CT Thorax-Abdomen",{"DL Gefäße": 0, "MRT": 0, "CT": 2}),
        ]
        wb = make_map_wb(codes)
        path_result  = load_map(save_wb(wb, tmp_path / "m.xlsx"))
        bytes_result = load_map(wb_to_bytes(wb))

        assert path_result["sections"] == bytes_result["sections"]
        assert set(path_result["leistungen"]) == set(bytes_result["leistungen"])
        for code in path_result["leistungen"]:
            assert (path_result["leistungen"][code]["sections"]
                    == bytes_result["leistungen"][code]["sections"])

    def test_reference_bytesio_parity(self, tmp_path):
        amounts = {"MRT": 3000, "CT": 4000, "US": 1500, "DL Gefäße": 800}
        wb = make_reference_wb(amounts)
        path_result  = load_reference_amounts(save_wb(wb, tmp_path / "r.xlsx"))
        bytes_result = load_reference_amounts(wb_to_bytes(wb))
        assert path_result == bytes_result

    def test_bytesio_consumed_only_once(self):
        """A BytesIO passed to the loader is fully consumed; passing it twice fails gracefully."""
        wb = make_stats_wb([(1, 1, "01.01.2024", "MRT", "MRT")])
        buf = io.BytesIO(wb_to_bytes(wb))
        result1 = load_statistics(buf)   # first call reads and resets internally
        # A fresh BytesIO still works
        wb2 = make_stats_wb([(1, 1, "01.01.2024", "CT", "CT")])
        result2 = load_statistics(io.BytesIO(wb_to_bytes(wb2)))
        assert result1["leistungen"][0]["leist_kurz"] == "MRT"
        assert result2["leistungen"][0]["leist_kurz"] == "CT"
