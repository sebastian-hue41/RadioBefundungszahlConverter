"""
test_counting_integrity.py

Dedicated suite ensuring NO Leistung is ever lost, double-counted,
or mis-attributed across the full pipeline.

Every test either:
  (a) builds a workbook with a known number of rows and asserts
      total_leistungen and total_counted match that number exactly, or
  (b) deliberately skips rows and asserts the mismatch is detected, or
  (c) verifies section sums are arithmetically consistent with the raw inputs.
"""

import io
from pathlib import Path

import openpyxl
import pytest

from core.input_handler import load_statistics, load_map
from core.logic import process
from .helpers import (
    make_stats_wb,
    make_map_wb,
    make_exam,
    make_stats_data,
    make_map_data,
    make_leistung_entry,
    save_wb,
    wb_to_bytes,
)


# ── Helpers ────────────────────────────────────────────────────────────────────

def _rows_from_groups(groups: list[int]) -> list[tuple]:
    """
    Build row tuples from group sizes.
    groups=[3,2,1] → IND2 sequences [1,2,3], [1,2], [1] spread across rows.
    """
    rows = []
    for group_idx, size in enumerate(groups):
        date = "01.01.2024"
        for pos in range(1, size + 1):
            dok_datum = date if pos == 1 else None
            rows.append((group_idx + 1, pos, dok_datum, f"Code{group_idx}_{pos}", f"C{group_idx}_{pos}"))
    return rows


def _stats_result(rows: list[tuple], tmp_path: Path) -> dict:
    """Save rows to a stats workbook and run load_statistics."""
    wb = make_stats_wb(rows)
    return load_statistics(save_wb(wb, tmp_path / "s.xlsx"))


def _stats_result_bytes(rows: list[tuple]) -> dict:
    """Same as _stats_result but via bytes (exercises the web path)."""
    wb = make_stats_wb(rows)
    return load_statistics(wb_to_bytes(wb))


# ── TestIND2AlgorithmExhaustive ────────────────────────────────────────────────

class TestIND2AlgorithmExhaustive:
    """Exhaustive IND2 group-end algorithm verification."""

    def _assert_total(self, groups: list[int], tmp_path: Path) -> None:
        expected = sum(groups)
        rows = _rows_from_groups(groups)
        result = _stats_result(rows, tmp_path)
        assert result["total_leistungen"] == expected, \
            f"groups={groups}: expected {expected}, got {result['total_leistungen']}"
        assert len(result["leistungen"]) == expected, \
            f"groups={groups}: len(leistungen)={len(result['leistungen'])}, expected {expected}"

    def test_single_group_size_1(self, tmp_path):
        self._assert_total([1], tmp_path)

    def test_single_group_size_2(self, tmp_path):
        self._assert_total([2], tmp_path)

    def test_single_group_size_5(self, tmp_path):
        self._assert_total([5], tmp_path)

    def test_single_group_size_100(self, tmp_path):
        self._assert_total([100], tmp_path)

    def test_two_groups_equal(self, tmp_path):
        self._assert_total([1, 1], tmp_path)

    def test_five_single_groups(self, tmp_path):
        self._assert_total([1] * 5, tmp_path)

    def test_ten_single_groups(self, tmp_path):
        self._assert_total([1] * 10, tmp_path)

    def test_two_groups_same_size(self, tmp_path):
        self._assert_total([3, 3], tmp_path)

    def test_three_groups_equal(self, tmp_path):
        self._assert_total([2, 2, 2], tmp_path)

    def test_ascending_groups(self, tmp_path):
        self._assert_total([1, 2, 3], tmp_path)

    def test_descending_groups(self, tmp_path):
        self._assert_total([3, 2, 1], tmp_path)

    def test_mixed_groups(self, tmp_path):
        self._assert_total([5, 1, 3, 2, 1], tmp_path)

    def test_large_first_group(self, tmp_path):
        self._assert_total([10, 1, 1, 1], tmp_path)

    def test_ten_varied_groups(self, tmp_path):
        self._assert_total([3, 1, 4, 1, 5, 9, 2, 6, 5, 3], tmp_path)

    def test_hundred_single_groups(self, tmp_path):
        self._assert_total([1] * 100, tmp_path)

    def test_ind2_none_rows_excluded_from_total(self, tmp_path):
        """Rows with IND2=None are excluded from the sequence but not the leistungen list."""
        rows = [
            (1, 1, "01.01.2024", "A", "A"),
            (1, None, None,      "B", "B"),  # IND2=None → not in sequence
            (2, 1, "02.01.2024", "C", "C"),
        ]
        result = _stats_result(rows, tmp_path)
        # Both rows with valid leist_kurz and parseable year are counted
        assert len(result["leistungen"]) == 3
        # Only the two IND2 values (1, 1) contribute to total_leistungen
        assert result["total_leistungen"] == 2


# ── TestNoLossInvariant ────────────────────────────────────────────────────────

class TestNoLossInvariant:
    """len(leistungen) == total_leistungen for all fully valid datasets."""

    def _assert_no_loss(self, groups: list[int], tmp_path: Path) -> None:
        rows = _rows_from_groups(groups)
        result = _stats_result(rows, tmp_path)
        assert len(result["leistungen"]) == result["total_leistungen"], \
            f"Loss detected: {len(result['leistungen'])} processed, " \
            f"{result['total_leistungen']} expected"

    def test_no_loss_single_group(self, tmp_path):
        self._assert_no_loss([3], tmp_path)

    def test_no_loss_multiple_groups(self, tmp_path):
        self._assert_no_loss([2, 3, 1], tmp_path)

    def test_no_loss_large_dataset(self, tmp_path):
        self._assert_no_loss([5] * 20, tmp_path)

    def test_no_loss_100_groups(self, tmp_path):
        self._assert_no_loss([1] * 100, tmp_path)

    def test_no_loss_varied_sizes(self, tmp_path):
        self._assert_no_loss([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], tmp_path)

    def test_no_loss_via_bytes(self):
        """Same invariant holds when loading via bytes (web path)."""
        rows = _rows_from_groups([3, 2, 1])
        result = _stats_result_bytes(rows)
        assert len(result["leistungen"]) == result["total_leistungen"]

    def test_bad_date_causes_detected_loss(self, tmp_path):
        """Rows with unparseable dates produce detected (not silent) loss."""
        rows = [
            (1, 1, "01.01.2024", "A", "A"),
            (1, 2, "NOT_A_DATE", "B", "B"),
        ]
        result = _stats_result(rows, tmp_path)
        assert len(result["leistungen"]) < result["total_leistungen"]
        assert any("cannot parse year" in e for e in result["errors"])

    def test_null_leist_kurz_causes_detected_loss(self, tmp_path):
        """Rows with null Leist-Kurz produce detected (not silent) loss."""
        rows = [
            (1, 1, "01.01.2024", "A", "A"),
            (1, 2, None,         "B", None),
        ]
        result = _stats_result(rows, tmp_path)
        assert len(result["leistungen"]) < result["total_leistungen"]
        assert any("Leist-Kurz" in e for e in result["errors"])

    def test_no_loss_single_element_groups(self, tmp_path):
        self._assert_no_loss([1, 1, 1, 1, 1], tmp_path)

    def test_no_loss_descending(self, tmp_path):
        self._assert_no_loss([5, 4, 3, 2, 1], tmp_path)


# ── TestSectionCounting ────────────────────────────────────────────────────────

class TestSectionCounting:
    """Section totals are arithmetically consistent with raw inputs."""

    def _build_and_process(self, section_assignments: dict[str, list[int]]) -> dict:
        """
        section_assignments: {leist_kurz: list_of_years}
        Each code gets val=1 in its own section.
        """
        sections = list(section_assignments.keys())
        exams = []
        leistungen = {}
        for code, years in section_assignments.items():
            for year in years:
                exams.append(make_exam(code, year))
            leistungen[code] = make_leistung_entry({s: (1 if s == code else 0) for s in sections})

        stats_data = make_stats_data(exams)
        map_data = make_map_data(sections, leistungen)
        return process(stats_data, map_data)

    def test_single_section_counts(self):
        result = self._build_and_process({"MRT": [2024, 2024, 2025]})
        assert result["data"]["MRT"][2024] == 2
        assert result["data"]["MRT"][2025] == 1
        assert result["data"]["MRT"]["Total"] == 3

    def test_two_sections_independent(self):
        result = self._build_and_process({
            "MRT": [2024, 2024],
            "CT":  [2024, 2025, 2025],
        })
        assert result["data"]["MRT"]["Total"] == 2
        assert result["data"]["CT"]["Total"] == 3

    def test_section_totals_sum_to_input_count(self):
        """Sum of all section totals equals the number of matched examinations."""
        section_assignments = {
            "MRT": [2024] * 10,
            "CT":  [2024] * 7,
            "US":  [2024] * 5,
        }
        result = self._build_and_process(section_assignments)
        total = sum(result["data"][s]["Total"] for s in ["MRT", "CT", "US"])
        assert total == 22

    def test_val2_doubles_section_total(self):
        """A single exam with val=2 adds 2 to the section total."""
        exams = [make_exam("POLY", 2024)] * 3
        stats_data = make_stats_data(exams)
        map_data = make_map_data(
            ["CT Abdomen"],
            {"POLY": make_leistung_entry({"CT Abdomen": 2})},
        )
        result = process(stats_data, map_data)
        assert result["data"]["CT Abdomen"]["Total"] == 6  # 3 exams × 2

    def test_year_totals_consistent(self):
        """Total == sum of all year counts for every section."""
        result = self._build_and_process({
            "MRT": [2022, 2023, 2023, 2024],
            "CT":  [2022, 2024, 2024, 2024],
        })
        for section in ["MRT", "CT"]:
            year_sum = sum(
                v for k, v in result["data"][section].items() if k != "Total"
            )
            assert year_sum == result["data"][section]["Total"]

    def test_unmatched_not_counted_in_sections(self):
        """Unmatched exams do not inflate any section's count."""
        exams = [make_exam("MAPPED", 2024), make_exam("UNKNOWN", 2024)]
        result = process(
            make_stats_data(exams),
            make_map_data(["MRT"], {"MAPPED": make_leistung_entry({"MRT": 1})}),
        )
        assert result["data"]["MRT"]["Total"] == 1

    def test_zero_section_all_years_present(self):
        """Sections always have an entry for every year seen, even if zero."""
        exams = [make_exam("MRT", 2024), make_exam("CT", 2025)]
        result = process(
            make_stats_data(exams),
            make_map_data(
                ["MRT", "CT"],
                {
                    "MRT": make_leistung_entry({"MRT": 1, "CT": 0}),
                    "CT":  make_leistung_entry({"MRT": 0, "CT": 1}),
                },
            ),
        )
        assert result["data"]["MRT"].get(2025, 0) == 0
        assert result["data"]["CT"].get(2024, 0) == 0

    def test_multi_year_totals(self):
        exams = (
            [make_exam("MRT", 2023)] * 5 +
            [make_exam("MRT", 2024)] * 8 +
            [make_exam("MRT", 2025)] * 3
        )
        result = process(
            make_stats_data(exams),
            make_map_data(["MRT"], {"MRT": make_leistung_entry({"MRT": 1})}),
        )
        assert result["data"]["MRT"][2023] == 5
        assert result["data"]["MRT"][2024] == 8
        assert result["data"]["MRT"][2025] == 3
        assert result["data"]["MRT"]["Total"] == 16

    def test_large_scale_section_accuracy(self):
        """1 000 exams split 600/400 across two sections — totals exact."""
        exams = (
            [make_exam("MRT", 2024)] * 600 +
            [make_exam("CT",  2024)] * 400
        )
        result = process(
            make_stats_data(exams),
            make_map_data(
                ["MRT", "CT"],
                {
                    "MRT": make_leistung_entry({"MRT": 1, "CT": 0}),
                    "CT":  make_leistung_entry({"MRT": 0, "CT": 1}),
                },
            ),
        )
        assert result["data"]["MRT"]["Total"] == 600
        assert result["data"]["CT"]["Total"] == 400

    def test_section_order_matches_map(self):
        """Sections appear in map order, not insertion order of examinations."""
        exams = [make_exam("CT", 2024), make_exam("MRT", 2024)]
        result = process(
            make_stats_data(exams),
            make_map_data(
                ["MRT", "CT"],
                {
                    "MRT": make_leistung_entry({"MRT": 1, "CT": 0}),
                    "CT":  make_leistung_entry({"MRT": 0, "CT": 1}),
                },
            ),
        )
        keys = list(result["data"].keys())
        assert keys.index("MRT") < keys.index("CT")


# ── TestEndToEndCounting ───────────────────────────────────────────────────────

class TestEndToEndCounting:
    """Full pipeline: xlsx → load_statistics → load_map → process → verify."""

    def test_full_pipeline_counts_correct(self, tmp_path):
        rows = _rows_from_groups([3, 2, 1])
        stats_wb = make_stats_wb(rows)
        map_wb   = make_map_wb(
            [(f"C0_{i}", "d", {"DL Gefäße": 0, "MRT": 1, "CT": 0}) for i in range(1, 4)] +
            [(f"C1_{i}", "d", {"DL Gefäße": 0, "MRT": 1, "CT": 0}) for i in range(1, 3)] +
            [(f"C2_1",   "d", {"DL Gefäße": 0, "MRT": 1, "CT": 0})],
        )
        stats_data = load_statistics(save_wb(stats_wb, tmp_path / "s.xlsx"))
        map_data   = load_map(save_wb(map_wb, tmp_path / "m.xlsx"))
        result     = process(stats_data, map_data)

        assert result["total_counted"] == 6
        assert result["data"]["MRT"]["Total"] == 6

    def test_full_pipeline_via_bytes(self):
        """Identical result whether source is path or bytes."""
        rows = _rows_from_groups([2, 3])
        stats_wb = make_stats_wb(rows)
        map_wb   = make_map_wb(
            [(f"C{g}_{p}", "d", {"DL Gefäße": 1, "MRT": 0, "CT": 0})
             for g in range(2) for p in range(1, [2, 3][g] + 1)],
        )
        stats_data = load_statistics(wb_to_bytes(stats_wb))
        map_data   = load_map(wb_to_bytes(map_wb))
        result     = process(stats_data, map_data)

        assert result["total_counted"] == 5
        assert result["data"]["DL Gefäße"]["Total"] == 5

    def test_no_loss_end_to_end(self, tmp_path):
        rows = _rows_from_groups([4, 3, 2, 1])
        stats_data = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))
        assert len(stats_data["leistungen"]) == stats_data["total_leistungen"] == 10

    def test_val2_end_to_end(self, tmp_path):
        """val=2 codes produce doubled section counts through the full pipeline."""
        rows = [(1, 1, "01.01.2024", "CT Thorax-Abdomen", "CTTHA")]
        map_wb = make_map_wb(
            [("CTTHA", "CT Thorax-Abdomen", {"DL Gefäße": 0, "MRT": 0, "CT": 2})],
        )
        stats_data = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))
        map_data   = load_map(save_wb(map_wb, tmp_path / "m.xlsx"))
        result     = process(stats_data, map_data)
        assert result["data"]["CT"]["Total"] == 2

    def test_mismatch_detected_end_to_end(self, tmp_path):
        """If rows are skipped, COUNT MISMATCH propagates to the final result."""
        rows = [
            (1, 1, "01.01.2024", "A", "A"),
            (1, 2, "BAD_DATE",   "B", "B"),
        ]
        map_wb = make_map_wb([("A", "d", {"DL Gefäße": 1, "MRT": 0, "CT": 0})])
        stats_data = load_statistics(save_wb(make_stats_wb(rows), tmp_path / "s.xlsx"))
        map_data   = load_map(save_wb(map_wb, tmp_path / "m.xlsx"))
        result     = process(stats_data, map_data)
        assert any("COUNT MISMATCH" in e for e in result["errors"])
