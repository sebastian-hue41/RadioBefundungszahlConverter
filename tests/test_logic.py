"""Tests for core/logic.py — core processing and counting logic."""

import pytest

from core.logic import process
from .helpers import (
    make_stats_data,
    make_map_data,
    make_exam,
    make_leistung_entry,
)


class TestProcessCounting:

    def test_process_basic_single_section(self):
        """Basic aggregation: single examination in single section."""
        exams = [make_exam("MRT", 2024)]
        result = process(
            make_stats_data(exams),
            make_map_data(["MRT"], {"MRT": make_leistung_entry({"MRT": 1})}),
        )
        assert result["data"]["MRT"][2024] == 1
        assert result["data"]["MRT"]["Total"] == 1

    def test_process_multiple_years(self):
        """Examinations from multiple years counted separately."""
        exams = [make_exam("MRT", 2024), make_exam("MRT", 2025), make_exam("MRT", 2024)]
        result = process(
            make_stats_data(exams),
            make_map_data(["MRT"], {"MRT": make_leistung_entry({"MRT": 1})}),
        )
        assert result["data"]["MRT"][2024] == 2
        assert result["data"]["MRT"][2025] == 1
        assert result["data"]["MRT"]["Total"] == 3

    def test_process_multiple_sections(self):
        """Single examination can belong to multiple sections simultaneously."""
        exams = [make_exam("COMBO", 2024)]
        result = process(
            make_stats_data(exams),
            make_map_data(
                ["MRT", "CT"],
                {"COMBO": make_leistung_entry({"MRT": 1, "CT": 1})},
            ),
        )
        assert result["data"]["MRT"][2024] == 1
        assert result["data"]["CT"][2024] == 1

    def test_process_val2_counted_twice(self):
        """Map value 2 → examination counted as 2 in that section."""
        exams = [make_exam("MULTIREG", 2024)]
        result = process(
            make_stats_data(exams),
            make_map_data(
                ["CT Abdomen"],
                {"MULTIREG": make_leistung_entry({"CT Abdomen": 2})},
            ),
        )
        assert result["data"]["CT Abdomen"][2024] == 2
        assert result["data"]["CT Abdomen"]["Total"] == 2

    def test_process_val3_counted_three_times(self):
        """Map value 3 → examination counted as 3 in that section."""
        exams = [make_exam("POLY", 2024)]
        result = process(
            make_stats_data(exams),
            make_map_data(
                ["CT Skelett"],
                {"POLY": make_leistung_entry({"CT Skelett": 3})},
            ),
        )
        assert result["data"]["CT Skelett"][2024] == 3

    def test_process_unmatched_procedures_appended(self):
        """Unmatched leistung gets its own row at the end."""
        exams = [make_exam("UNKNOWN_CODE", 2024, leistung="Unknown Procedure")]
        result = process(
            make_stats_data(exams),
            make_map_data(["MRT"], {}),
        )
        assert "Unknown Procedure" in result["data"]

    def test_process_unmatched_one_note_per_unique_code(self):
        """Each unique unmatched code produces exactly one NOTE error."""
        exams = [make_exam("UNKNOWN", 2024)] * 5
        result = process(make_stats_data(exams), make_map_data(["MRT"], {}))
        notes = [e for e in result["errors"] if e.startswith("NOTE")]
        assert len(notes) == 1

    def test_process_unmatched_uses_leistung_name_if_available(self):
        """Display label uses Leistungsbezeichnung when available."""
        exams = [make_exam("CODE", 2024, leistung="Human Readable Name")]
        result = process(make_stats_data(exams), make_map_data(["MRT"], {}))
        assert "Human Readable Name" in result["data"]
        assert "CODE" not in result["data"]

    def test_process_years_sorted(self):
        """Years in result are sorted ascending."""
        exams = [make_exam("MRT", 2025), make_exam("MRT", 2023), make_exam("MRT", 2024)]
        result = process(
            make_stats_data(exams),
            make_map_data(["MRT"], {"MRT": make_leistung_entry({"MRT": 1})}),
        )
        assert result["years"] == [2023, 2024, 2025]

    def test_process_total_per_section(self):
        """Section totals equal sum of year counts."""
        exams = [make_exam("MRT", 2024)] * 3 + [make_exam("MRT", 2025)] * 2
        result = process(
            make_stats_data(exams),
            make_map_data(["MRT"], {"MRT": make_leistung_entry({"MRT": 1})}),
        )
        assert result["data"]["MRT"]["Total"] == 5

    def test_process_zero_counts_for_unused_sections(self):
        """Sections with no examinations show zeros, not missing keys."""
        exams = [make_exam("MRT", 2024)]
        result = process(
            make_stats_data(exams),
            make_map_data(
                ["MRT", "CT"],
                {"MRT": make_leistung_entry({"MRT": 1, "CT": 0})},
            ),
        )
        assert result["data"]["CT"][2024] == 0
        assert result["data"]["CT"]["Total"] == 0

    def test_process_count_match_no_error(self):
        """No COUNT MISMATCH error when totals match."""
        exams = [make_exam("MRT", 2024)] * 3
        result = process(
            make_stats_data(exams, total_leistungen=3),
            make_map_data(["MRT"], {"MRT": make_leistung_entry({"MRT": 1})}),
        )
        assert not any("COUNT MISMATCH" in e for e in result["errors"])

    def test_process_count_mismatch_detected(self):
        """COUNT MISMATCH error added when totals differ."""
        exams = [make_exam("MRT", 2024)] * 3
        result = process(
            make_stats_data(exams, total_leistungen=5),
            make_map_data(["MRT"], {"MRT": make_leistung_entry({"MRT": 1})}),
        )
        assert any("COUNT MISMATCH" in e for e in result["errors"])

    def test_process_total_counted_correct(self):
        """total_counted == len(examinations)."""
        exams = [make_exam("MRT", 2024)] * 7
        result = process(
            make_stats_data(exams),
            make_map_data(["MRT"], {"MRT": make_leistung_entry({"MRT": 1})}),
        )
        assert result["total_counted"] == 7

    def test_process_metadata_passthrough(self):
        """mitarbeiter and befunddatum are passed through unchanged."""
        exams = [make_exam("MRT", 2024)]
        result = process(
            make_stats_data(exams, mitarbeiter="DR_SMITH", befunddatum="2024"),
            make_map_data(["MRT"], {"MRT": make_leistung_entry({"MRT": 1})}),
        )
        assert result["mitarbeiter"] == "DR_SMITH"
        assert result["befunddatum"] == "2024"

    def test_process_errors_accumulated(self):
        """Errors from stats_data are carried into the result."""
        exams = [make_exam("MRT", 2024)]
        result = process(
            make_stats_data(exams, errors=["WARNING: something"]),
            make_map_data(["MRT"], {"MRT": make_leistung_entry({"MRT": 1})}),
        )
        assert any("WARNING: something" in e for e in result["errors"])

    def test_process_merged_duplicates_passthrough(self):
        """merged_duplicates from map_data are passed through."""
        exams = [make_exam("MRT", 2024)]
        result = process(
            make_stats_data(exams),
            make_map_data(
                ["MRT"],
                {"MRT": make_leistung_entry({"MRT": 1})},
                merged_duplicates=["DUPE1"],
            ),
        )
        assert "DUPE1" in result["merged_duplicates"]

    def test_process_large_volume(self):
        """1 000 examinations processed without error."""
        exams = [make_exam("MRT", 2024 + (i % 3)) for i in range(1000)]
        result = process(
            make_stats_data(exams),
            make_map_data(["MRT"], {"MRT": make_leistung_entry({"MRT": 1})}),
        )
        assert result["data"]["MRT"]["Total"] == 1000

    def test_process_empty_years_section_ordering(self):
        """Sections from map come first; unmatched appended at end."""
        exams = [make_exam("MRT", 2024), make_exam("UNKNOWN", 2024)]
        result = process(
            make_stats_data(exams),
            make_map_data(["MRT", "CT"], {"MRT": make_leistung_entry({"MRT": 1, "CT": 0})}),
        )
        keys = list(result["data"].keys())
        assert keys.index("MRT") < keys.index("CT")
        assert keys[-1] == "UNKNOWN"
