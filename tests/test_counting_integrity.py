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

from pathlib import Path
import openpyxl
import pytest

from input_handler import load_statistics, load_map
from logic import process
from .helpers import (
    make_stats_wb,
    make_map_wb,
    make_exam,
    make_stats_data,
    make_map_data,
    make_leistung_entry,
    save_wb,
)


# ── Helpers ────────────────────────────────────────────────────────────────────

def _stats_result(rows, tmp_path, **kwargs):
    """Shortcut: build stats workbook → load → return result dict."""
    wb = make_stats_wb(rows, **kwargs)
    path = save_wb(wb, tmp_path / "stats.xlsx")
    return load_statistics(path)


def _process_result(rows, sections_map, tmp_path, **kwargs):
    """
    End-to-end shortcut: stats file + map dict → process() result.

    sections_map: {leist_kurz: {section_name: 1/0, ...}}
    Returns logic.process() result dict.
    """
    stat_rows = rows
    wb = make_stats_wb(stat_rows, **kwargs)
    path = save_wb(wb, tmp_path / "stats.xlsx")
    stats_data = load_statistics(path)

    sections = sorted({s for m in sections_map.values() for s in m})
    leistungen = {
        code: make_leistung_entry(mapping)
        for code, mapping in sections_map.items()
    }
    map_data = make_map_data(sections, leistungen)
    return process(stats_data, map_data)


# ── IND2 algorithm: exhaustive pattern coverage ────────────────────────────────

class TestIND2AlgorithmExhaustive:
    """
    Every conceivable IND2 sequence pattern.

    Convention for test rows:
        (ind1, ind2, dok_datum, leistung, leist_kurz)
    Groups always start with "15.03.2024"; subsequent rows in the same
    group have DokDatum=None (carry-forward).
    """

    def _rows_from_groups(self, groups: list[int]) -> list[tuple]:
        """
        Build row list from group sizes.
        groups = [3, 2, 1] → three groups of sizes 3, 2, 1.
        """
        rows = []
        ind1 = 1
        for size in groups:
            for ind2 in range(1, size + 1):
                dok = "15.03.2024" if ind2 == 1 else None
                rows.append((ind1, ind2, dok, "Leistung", "MRT"))
            ind1 += 1
        return rows

    def _assert_total(self, groups: list[int], tmp_path: Path):
        """Build file with the given group sizes; assert total matches sum(groups)."""
        expected = sum(groups)
        rows = self._rows_from_groups(groups)
        result = _stats_result(rows, tmp_path)
        # total_leistungen: IND2 algorithm expected count
        assert result["total_leistungen"] == expected, (
            f"IND2 algorithm wrong for groups {groups}: "
            f"expected {expected}, got {result['total_leistungen']}"
        )
        # len(leistungen): rows actually loaded — must match
        assert len(result["leistungen"]) == expected, (
            f"Loaded row count wrong for groups {groups}: "
            f"expected {expected}, got {len(result['leistungen'])}"
        )

    def test_single_item_group(self, tmp_path):
        """[1] → 1."""
        self._assert_total([1], tmp_path)

    def test_two_item_group(self, tmp_path):
        """[1,2] → 2."""
        self._assert_total([2], tmp_path)

    def test_five_item_group(self, tmp_path):
        """Single group of 5 → 5."""
        self._assert_total([5], tmp_path)

    def test_large_group_100(self, tmp_path):
        """Single group of 100 → 100."""
        self._assert_total([100], tmp_path)

    def test_two_single_item_groups(self, tmp_path):
        """[1],[1] → 2 (two groups each of size 1)."""
        self._assert_total([1, 1], tmp_path)

    def test_five_single_item_groups(self, tmp_path):
        """Five groups of size 1 → 5."""
        self._assert_total([1, 1, 1, 1, 1], tmp_path)

    def test_ten_single_item_groups(self, tmp_path):
        """Ten groups of size 1 → 10."""
        self._assert_total([1] * 10, tmp_path)

    def test_two_equal_groups(self, tmp_path):
        """Two groups of size 3 → 6."""
        self._assert_total([3, 3], tmp_path)

    def test_three_equal_groups(self, tmp_path):
        """Three groups of size 2 → 6."""
        self._assert_total([2, 2, 2], tmp_path)

    def test_ascending_group_sizes(self, tmp_path):
        """Groups [1, 2, 3] → 6."""
        self._assert_total([1, 2, 3], tmp_path)

    def test_descending_group_sizes(self, tmp_path):
        """Groups [3, 2, 1] → 6."""
        self._assert_total([3, 2, 1], tmp_path)

    def test_mixed_group_sizes(self, tmp_path):
        """Groups [5, 1, 3, 2, 1] → 12."""
        self._assert_total([5, 1, 3, 2, 1], tmp_path)

    def test_large_then_small_groups(self, tmp_path):
        """Groups [10, 1, 1, 1] → 13."""
        self._assert_total([10, 1, 1, 1], tmp_path)

    def test_ten_groups_varied(self, tmp_path):
        """10 groups of varying sizes — total must be exact."""
        groups = [3, 1, 5, 2, 1, 4, 1, 6, 2, 3]
        self._assert_total(groups, tmp_path)

    def test_hundred_groups_of_one(self, tmp_path):
        """100 single-item groups → 100."""
        self._assert_total([1] * 100, tmp_path)

    def test_ind2_none_rows_not_counted(self, tmp_path):
        """Rows where IND2 is None are excluded from the algorithm (not counted)."""
        rows = [
            (1, 1, "15.03.2024", "MR", "MRT"),
            (1, None, None, "US", "US"),  # IND2=None → excluded from sequence
            (1, 2, None, "CT", "CT"),
        ]
        wb = make_stats_wb(rows)
        path = save_wb(wb, tmp_path / "stats.xlsx")
        result = load_statistics(path)
        # IND2 sequence is [1, 2]; None row skipped in sequence
        # Algorithm: [1, 2] → single group of 2 → total_leistungen = 2
        assert result["total_leistungen"] == 2
        # The None-IND2 row itself is still a valid exam row (leist_kurz + date present)
        assert len(result["leistungen"]) == 3
        # Mismatch: IND2 expects 2 but 3 rows were loaded — detected by process()
        # Verify by running through process
        from logic import process
        from .helpers import make_map_data, make_leistung_entry
        map_data = make_map_data(
            ["MRT"], {"MRT": make_leistung_entry({"MRT": 1})}
        )
        proc = process(result, map_data)
        assert any("COUNT MISMATCH" in e for e in proc["errors"])


# ── No-loss invariant: total_counted == total_leistungen ──────────────────────

class TestNoLossInvariant:
    """
    For every valid file: total_counted must equal total_leistungen.
    A discrepancy means rows were silently dropped — catastrophic.
    """

    def _assert_no_loss(self, rows: list, tmp_path: Path):
        """Assert no rows are lost in loading."""
        result = _stats_result(rows, tmp_path)
        n_loaded = len(result["leistungen"])
        assert n_loaded == result["total_leistungen"], (
            f"LOSS DETECTED: loaded={n_loaded} "
            f"!= total_leistungen={result['total_leistungen']}. "
            f"Errors: {result['errors']}"
        )

    def test_no_loss_single_row(self, tmp_path):
        """Single row: no loss."""
        rows = [(1, 1, "15.03.2024", "MR", "MRT")]
        self._assert_no_loss(rows, tmp_path)

    def test_no_loss_single_large_group(self, tmp_path):
        """One group of 10: no loss."""
        rows = [(1, i, "15.03.2024" if i == 1 else None, "L", "MRT")
                for i in range(1, 11)]
        self._assert_no_loss(rows, tmp_path)

    def test_no_loss_50_single_groups(self, tmp_path):
        """50 single-item groups: no loss."""
        rows = [(i, 1, "15.03.2024", "L", "MRT") for i in range(1, 51)]
        self._assert_no_loss(rows, tmp_path)

    def test_no_loss_mixed_groups(self, tmp_path):
        """Mixed groups (sizes 1–5): no loss."""
        rows = []
        ind1 = 1
        for size in [1, 2, 3, 1, 5, 4, 1, 2]:
            for ind2 in range(1, size + 1):
                rows.append((ind1, ind2, "15.03.2024" if ind2 == 1 else None, "L", "MRT"))
            ind1 += 1
        self._assert_no_loss(rows, tmp_path)

    def test_no_loss_multiple_codes(self, tmp_path):
        """Multiple Leist-Kurz codes: no loss."""
        codes = ["MRT", "CT", "US", "XRAY", "PET", "SONO"]
        rows = []
        ind1 = 1
        for i, code in enumerate(codes * 10):
            rows.append((ind1, 1, "15.03.2024", "desc", code))
            ind1 += 1
        self._assert_no_loss(rows, tmp_path)

    def test_no_loss_multiple_years(self, tmp_path):
        """Rows from different years: no loss."""
        rows = [
            (1, 1, "15.03.2024", "L", "MRT"),
            (2, 1, "20.06.2025", "L", "CT"),
            (3, 1, "01.01.2026", "L", "US"),
        ]
        self._assert_no_loss(rows, tmp_path)

    def test_no_loss_date_carryforward(self, tmp_path):
        """Date carry-forward in groups: no loss."""
        rows = [
            (1, 1, "15.03.2024", "L", "MRT"),
            (1, 2, None,         "L", "CT"),
            (1, 3, None,         "L", "US"),
            (2, 1, "20.06.2025", "L", "MRT"),
            (2, 2, None,         "L", "CT"),
        ]
        self._assert_no_loss(rows, tmp_path)

    def test_no_loss_large_dataset(self, tmp_path):
        """1000 rows across 200 groups of 5: no loss."""
        rows = []
        ind1 = 1
        for _ in range(200):
            for ind2 in range(1, 6):
                rows.append((ind1, ind2, "15.03.2024" if ind2 == 1 else None, "L", "MRT"))
            ind1 += 1
        self._assert_no_loss(rows, tmp_path)
        result = _stats_result(rows, tmp_path)
        assert len(result["leistungen"]) == 1000

    def test_loss_detected_when_bad_date_row_present(self, tmp_path):
        """Loss IS reported when a row is skipped due to unparseable date."""
        rows = [
            (1, 1, "15.03.2024", "L", "MRT"),
            (1, 2, "BADDATE",    "L", "CT"),   # skipped → loss
        ]
        wb = make_stats_wb(rows)
        path = save_wb(wb, tmp_path / "stats.xlsx")
        result = load_statistics(path)
        # IND2 algorithm: [1, 2] → expects 2; only 1 loaded
        assert result["total_leistungen"] == 2
        assert len(result["leistungen"]) == 1
        # The row-level warning is in errors; COUNT MISMATCH appears after process()
        assert any("cannot parse year" in e for e in result["errors"])

    def test_loss_detected_when_null_leist_kurz_present(self, tmp_path):
        """Loss IS reported when a row is skipped due to missing Leist-Kurz."""
        rows = [
            (1, 1, "15.03.2024", "L", "MRT"),
            (1, 2, None,         "L", None),   # Leist-Kurz missing → skipped
        ]
        wb = make_stats_wb(rows)
        path = save_wb(wb, tmp_path / "stats.xlsx")
        result = load_statistics(path)
        assert result["total_leistungen"] == 2
        assert len(result["leistungen"]) == 1
        assert any("Leist-Kurz" in e and "empty" in e for e in result["errors"])


# ── Section counting integrity ─────────────────────────────────────────────────

class TestSectionCounting:
    """
    Section totals must be arithmetically consistent with the input exams.
    """

    def test_each_exam_counted_in_its_section(self, tmp_path):
        """Every matched exam increments exactly its section(s)."""
        rows = [
            (1, 1, "15.03.2024", "L", "MRT"),
            (2, 1, "15.03.2024", "L", "MRT"),
            (3, 1, "15.03.2024", "L", "CT"),
        ]
        result = _process_result(
            rows,
            {"MRT": {"MRT": 1, "CT": 0}, "CT": {"MRT": 0, "CT": 1}},
            tmp_path,
        )
        assert result["data"]["MRT"][2024] == 2
        assert result["data"]["CT"][2024] == 1
        assert result["total_counted"] == 3

    def test_multi_section_exam_counted_in_all_sections(self, tmp_path):
        """One exam in multiple sections increments each section by 1."""
        rows = [(1, 1, "15.03.2024", "L", "COMBO")]
        result = _process_result(
            rows,
            {"COMBO": {"MRT": 1, "CT": 1, "US": 1}},
            tmp_path,
        )
        assert result["data"]["MRT"][2024] == 1
        assert result["data"]["CT"][2024] == 1
        assert result["data"]["US"][2024] == 1
        # total_counted is still 1 (one exam)
        assert result["total_counted"] == 1

    def test_multi_section_does_not_inflate_total_counted(self, tmp_path):
        """Multi-section mapping never inflates total_counted."""
        rows = [(i, 1, "15.03.2024", "L", "COMBO") for i in range(1, 11)]
        result = _process_result(
            rows,
            {"COMBO": {"MRT": 1, "CT": 1, "US": 1}},
            tmp_path,
        )
        assert result["total_counted"] == 10
        # Each section also gets 10
        assert result["data"]["MRT"]["Total"] == 10
        assert result["data"]["CT"]["Total"] == 10
        assert result["data"]["US"]["Total"] == 10

    def test_unmatched_exam_still_counted(self, tmp_path):
        """Unmatched procedure counted in total_counted and as its own row."""
        rows = [(1, 1, "15.03.2024", "Unknown Proc", "UNMAPPED")]
        result = _process_result(
            rows,
            {"KNOWN": {"MRT": 1}},
            tmp_path,
        )
        assert result["total_counted"] == 1
        assert "Unknown Proc" in result["data"]
        assert result["data"]["Unknown Proc"][2024] == 1

    def test_mixed_matched_unmatched_total_correct(self, tmp_path):
        """Mix of matched and unmatched: total_counted covers both."""
        rows = [
            (1, 1, "15.03.2024", "L",  "MRT"),
            (2, 1, "15.03.2024", "L",  "MRT"),
            (3, 1, "15.03.2024", "L2", "UNMAPPED1"),
            (4, 1, "15.03.2024", "L3", "UNMAPPED2"),
        ]
        result = _process_result(
            rows,
            {"MRT": {"MRT": 1}},
            tmp_path,
        )
        assert result["total_counted"] == 4
        assert result["data"]["MRT"][2024] == 2

    def test_section_total_equals_sum_of_year_counts(self, tmp_path):
        """Each section's 'Total' == sum of its per-year counts."""
        rows = [
            (1, 1, "15.03.2024", "L", "MRT"),
            (2, 1, "15.03.2024", "L", "MRT"),
            (3, 1, "20.06.2025", "L", "MRT"),
        ]
        result = _process_result(
            rows,
            {"MRT": {"MRT": 1}},
            tmp_path,
        )
        yr_sum = sum(result["data"]["MRT"][yr] for yr in result["years"])
        assert result["data"]["MRT"]["Total"] == yr_sum == 3

    def test_year_counts_sum_to_section_total(self, tmp_path):
        """For each section: sum of per-year counts == Total."""
        rows = [
            (1, 1, "15.03.2024", "L", "MRT"),
            (2, 1, "20.06.2025", "L", "MRT"),
            (3, 1, "01.01.2026", "L", "MRT"),
            (4, 1, "15.03.2024", "L", "CT"),
            (5, 1, "20.06.2025", "L", "CT"),
        ]
        result = _process_result(
            rows,
            {"MRT": {"MRT": 1, "CT": 0}, "CT": {"MRT": 0, "CT": 1}},
            tmp_path,
        )
        for section in ["MRT", "CT"]:
            yr_sum = sum(result["data"][section][yr] for yr in result["years"])
            total = result["data"][section]["Total"]
            assert yr_sum == total, f"Section {section}: year sum {yr_sum} != Total {total}"

    def test_all_sections_receive_correct_count(self, tmp_path):
        """10 distinct sections, 5 exams each, all counts correct."""
        n_sections = 10
        n_per_section = 5
        sections = [f"SEC{i}" for i in range(n_sections)]
        rows = []
        ind1 = 1
        for sec in sections:
            for _ in range(n_per_section):
                rows.append((ind1, 1, "15.03.2024", "L", sec))
                ind1 += 1

        sections_map = {sec: {s: (1 if s == sec else 0) for s in sections}
                        for sec in sections}
        result = _process_result(rows, sections_map, tmp_path)

        assert result["total_counted"] == n_sections * n_per_section
        for sec in sections:
            assert result["data"][sec]["Total"] == n_per_section, (
                f"Section {sec}: expected {n_per_section}, got {result['data'][sec]['Total']}"
            )

    def test_zero_section_assignment_not_counted(self, tmp_path):
        """A procedure with val=0 for a section does NOT count toward that section."""
        rows = [(1, 1, "15.03.2024", "L", "MRT")]
        result = _process_result(
            rows,
            {"MRT": {"MRT": 1, "CT": 0}},
            tmp_path,
        )
        assert result["data"]["MRT"][2024] == 1
        assert result["data"]["CT"][2024] == 0

    def test_100_exams_all_same_section(self, tmp_path):
        """100 exams all in one section: section total == 100."""
        rows = [(i, 1, "15.03.2024", "L", "MRT") for i in range(1, 101)]
        result = _process_result(
            rows,
            {"MRT": {"MRT": 1}},
            tmp_path,
        )
        assert result["total_counted"] == 100
        assert result["data"]["MRT"]["Total"] == 100


# ── Integration: file → load → process → count verification ───────────────────

class TestEndToEndCounting:
    """
    Full pipeline tests: xlsx on disk → load_statistics → process → verify.
    """

    def test_e2e_single_group(self, tmp_path):
        """End-to-end: single IND2 group of 3, all counted."""
        rows = [
            (1, 1, "15.03.2024", "MR Kopf",   "MRT"),
            (1, 2, None,         "CT Thorax",  "CT"),
            (1, 3, None,         "US Abdomen", "US"),
        ]
        wb = make_stats_wb(rows, mitarbeiter="KUNZ_A", befunddatum="2024")
        map_wb = make_map_wb(
            [
                ("MRT", "MR",  {"DL Gefäße": 0, "MRT": 1, "CT": 0}),
                ("CT",  "CT",  {"DL Gefäße": 0, "MRT": 0, "CT": 1}),
                ("US",  "US",  {"DL Gefäße": 1, "MRT": 0, "CT": 0}),
            ]
        )
        stats_path = save_wb(wb, tmp_path / "stats.xlsx")
        map_path   = save_wb(map_wb, tmp_path / "map.xlsx")

        stats_data = load_statistics(stats_path)
        map_data   = load_map(map_path)
        result     = process(stats_data, map_data)

        assert result["total_counted"] == 3
        assert result["total_leistungen"] == 3
        assert not any("COUNT MISMATCH" in e for e in result["errors"])
        assert result["data"]["MRT"]["Total"] == 1
        assert result["data"]["CT"]["Total"] == 1

    def test_e2e_multiple_groups_no_loss(self, tmp_path):
        """End-to-end: 5 groups of 2 = 10 exams, all counted."""
        rows = []
        ind1 = 1
        for _ in range(5):
            rows += [
                (ind1, 1, "15.03.2024", "MR Kopf",  "MRT"),
                (ind1, 2, None,         "CT Thorax", "CT"),
            ]
            ind1 += 1
        wb     = make_stats_wb(rows)
        map_wb = make_map_wb(
            [
                ("MRT", "MR", {"DL Gefäße": 0, "MRT": 1}),
                ("CT",  "CT", {"DL Gefäße": 0, "MRT": 0}),
            ],
            sections=["DL Gefäße", "MRT"],
        )
        stats_path = save_wb(wb,     tmp_path / "stats.xlsx")
        map_path   = save_wb(map_wb, tmp_path / "map.xlsx")

        stats_data = load_statistics(stats_path)
        map_data   = load_map(map_path)
        result     = process(stats_data, map_data)

        assert result["total_counted"] == 10
        assert result["total_leistungen"] == 10
        assert result["data"]["MRT"]["Total"] == 5

    def test_e2e_unmatched_procedure_not_lost(self, tmp_path):
        """End-to-end: unmatched procedure appears in output and in total_counted."""
        rows = [
            (1, 1, "15.03.2024", "Known Proc",   "MRT"),
            (2, 1, "15.03.2024", "Unknown Proc",  "UNKNOWN"),
        ]
        wb     = make_stats_wb(rows)
        map_wb = make_map_wb([("MRT", "MR", {"DL Gefäße": 1, "MRT": 1})])
        stats_path = save_wb(wb,     tmp_path / "stats.xlsx")
        map_path   = save_wb(map_wb, tmp_path / "map.xlsx")

        stats_data = load_statistics(stats_path)
        map_data   = load_map(map_path)
        result     = process(stats_data, map_data)

        assert result["total_counted"] == 2
        assert "Unknown Proc" in result["data"]
        assert result["data"]["Unknown Proc"]["Total"] == 1

    def test_e2e_multiple_years_counted_per_year(self, tmp_path):
        """End-to-end: exams across years counted in correct year buckets."""
        rows = [
            (1, 1, "15.03.2024", "L", "MRT"),
            (2, 1, "15.03.2024", "L", "MRT"),
            (3, 1, "20.06.2025", "L", "MRT"),
        ]
        wb     = make_stats_wb(rows)
        map_wb = make_map_wb([("MRT", "MR", {"DL Gefäße": 1, "MRT": 1})])
        stats_path = save_wb(wb,     tmp_path / "stats.xlsx")
        map_path   = save_wb(map_wb, tmp_path / "map.xlsx")

        stats_data = load_statistics(stats_path)
        map_data   = load_map(map_path)
        result     = process(stats_data, map_data)

        assert result["total_counted"] == 3
        assert result["data"]["MRT"][2024] == 2
        assert result["data"]["MRT"][2025] == 1
        assert result["data"]["MRT"]["Total"] == 3

    def test_e2e_duplicate_map_does_not_lose_exams(self, tmp_path):
        """Duplicate map entries (merged) must not cause exam rows to be lost."""
        rows = [
            (1, 1, "15.03.2024", "L", "DUP"),
            (2, 1, "15.03.2024", "L", "DUP"),
        ]
        # Two conflicting map entries for the same code
        wb_map = make_map_wb(
            [
                ("DUP", "v1", {"DL Gefäße": 1, "MRT": 1, "CT": 0}),
                ("DUP", "v2", {"DL Gefäße": 0, "MRT": 0, "CT": 1}),
            ]
        )
        stats_path = save_wb(make_stats_wb(rows), tmp_path / "stats.xlsx")
        map_path   = save_wb(wb_map,              tmp_path / "map.xlsx")

        stats_data = load_statistics(stats_path)
        map_data   = load_map(map_path)
        result     = process(stats_data, map_data)

        assert result["total_counted"] == 2
        # After merge by max, DUP belongs to all three sections
        assert result["data"]["DL Gefäße"]["Total"] == 2
        assert result["data"]["MRT"]["Total"] == 2
        assert result["data"]["CT"]["Total"] == 2
        assert "DUP" in result["merged_duplicates"]
