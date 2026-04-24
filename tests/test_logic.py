"""Tests for logic.py — core processing and counting logic."""

import pytest

from logic import process
from .helpers import (
    make_stats_data,
    make_map_data,
    make_exam,
    make_leistung_entry,
)


class TestProcessCounting:
    """Tests for examination counting and aggregation."""

    def test_process_basic_single_section(self):
        """Basic aggregation: single examination in single section."""
        exams = [make_exam("MRT", 2024)]
        stats_data = make_stats_data(exams)
        map_data = make_map_data(
            ["MRT"],
            {"MRT": make_leistung_entry({"MRT": 1})},
        )

        result = process(stats_data, map_data)
        assert result["data"]["MRT"][2024] == 1
        assert result["data"]["MRT"]["Total"] == 1

    def test_process_multiple_years(self):
        """Examinations from multiple years counted separately."""
        exams = [
            make_exam("MRT", 2024),
            make_exam("MRT", 2025),
            make_exam("MRT", 2024),
        ]
        stats_data = make_stats_data(exams)
        map_data = make_map_data(
            ["MRT"],
            {"MRT": make_leistung_entry({"MRT": 1})},
        )

        result = process(stats_data, map_data)
        assert result["data"]["MRT"][2024] == 2
        assert result["data"]["MRT"][2025] == 1
        assert result["data"]["MRT"]["Total"] == 3

    def test_process_multiple_sections(self):
        """Single examination can belong to multiple sections."""
        exams = [make_exam("MRT", 2024)]
        stats_data = make_stats_data(exams)
        map_data = make_map_data(
            ["DL Gefäße", "MRT"],
            {"MRT": make_leistung_entry({"DL Gefäße": 1, "MRT": 1})},
        )

        result = process(stats_data, map_data)
        assert result["data"]["DL Gefäße"][2024] == 1
        assert result["data"]["MRT"][2024] == 1

    def test_process_unmatched_procedures_appended(self):
        """Unmatched procedures added as separate rows at end."""
        exams = [
            make_exam("MRT", 2024),
            make_exam("UNMAPPED", 2024),
        ]
        stats_data = make_stats_data(exams)
        map_data = make_map_data(
            ["MRT"],
            {"MRT": make_leistung_entry({"MRT": 1})},
        )

        result = process(stats_data, map_data)
        sections_list = list(result["data"].keys())
        assert sections_list[0] == "MRT"  # Map sections first
        assert "UNMAPPED" in sections_list  # Unmatched appended

    def test_process_unmatched_one_note_per_unique_code(self):
        """One NOTE added per unique unmatched code, even if multiple rows."""
        exams = [
            make_exam("UNMAPPED", 2024),
            make_exam("UNMAPPED", 2024),  # Duplicate
            make_exam("UNMAPPED", 2025),  # Different year but same code
        ]
        stats_data = make_stats_data(exams)
        map_data = make_map_data(["MRT"], {"MRT": make_leistung_entry({"MRT": 1})})

        result = process(stats_data, map_data)
        unmapped_notes = [e for e in result["errors"] if "UNMAPPED" in e]
        assert len(unmapped_notes) == 1  # Only one NOTE for the unique code

    def test_process_unmatched_uses_leistung_name_if_available(self):
        """Unmatched procedure uses Leistung name if available, else leist_kurz."""
        exams = [
            make_exam("CODE1", 2024, leistung="Full Procedure Name"),
            make_exam("CODE2", 2024, leistung=None),
        ]
        stats_data = make_stats_data(exams)
        map_data = make_map_data(["MRT"], {"MRT": make_leistung_entry({"MRT": 1})})

        result = process(stats_data, map_data)
        # CODE1 should have the leistung name as display label
        assert "Full Procedure Name" in result["data"]
        assert result["data"]["Full Procedure Name"][2024] == 1
        # CODE2 should use leist_kurz
        assert "CODE2" in result["data"]

    def test_process_years_sorted(self):
        """Years list is sorted."""
        exams = [
            make_exam("MRT", 2026),
            make_exam("MRT", 2024),
            make_exam("MRT", 2025),
        ]
        stats_data = make_stats_data(exams)
        map_data = make_map_data(
            ["MRT"],
            {"MRT": make_leistung_entry({"MRT": 1})},
        )

        result = process(stats_data, map_data)
        assert result["years"] == [2024, 2025, 2026]

    def test_process_total_per_section(self):
        """Each section has a 'Total' key with sum of year counts."""
        exams = [
            make_exam("MRT", 2024),
            make_exam("MRT", 2024),
            make_exam("MRT", 2025),
        ]
        stats_data = make_stats_data(exams)
        map_data = make_map_data(
            ["MRT"],
            {"MRT": make_leistung_entry({"MRT": 1})},
        )

        result = process(stats_data, map_data)
        assert result["data"]["MRT"]["Total"] == 3

    def test_process_zero_counts_for_unused_sections(self):
        """Sections with no matching examinations have 0 counts."""
        exams = [make_exam("MRT", 2024)]
        stats_data = make_stats_data(exams)
        map_data = make_map_data(
            ["MRT", "CT", "US"],
            {
                "MRT": make_leistung_entry({"MRT": 1, "CT": 0, "US": 0}),
            },
        )

        result = process(stats_data, map_data)
        assert result["data"]["CT"][2024] == 0
        assert result["data"]["US"][2024] == 0
        assert result["data"]["CT"]["Total"] == 0

    def test_process_count_match_no_error(self):
        """When total_counted == total_leistungen, no mismatch error."""
        exams = [make_exam("MRT", 2024)]
        stats_data = make_stats_data(
            exams,
            total_leistungen=1,  # Matches len(exams)
        )
        map_data = make_map_data(
            ["MRT"],
            {"MRT": make_leistung_entry({"MRT": 1})},
        )

        result = process(stats_data, map_data)
        assert not any("COUNT MISMATCH" in e for e in result["errors"])

    def test_process_count_mismatch_detected(self):
        """Count mismatch when total_counted ≠ total_leistungen."""
        exams = [make_exam("MRT", 2024)]
        stats_data = make_stats_data(
            exams,
            total_leistungen=5,  # Expected 5, but only 1 processed
        )
        map_data = make_map_data(
            ["MRT"],
            {"MRT": make_leistung_entry({"MRT": 1})},
        )

        result = process(stats_data, map_data)
        assert any("COUNT MISMATCH" in e for e in result["errors"])

    def test_process_total_counted_correct(self):
        """total_counted equals length of examinations."""
        exams = [
            make_exam("MRT", 2024),
            make_exam("CT", 2024),
            make_exam("US", 2025),
        ]
        stats_data = make_stats_data(exams)
        map_data = make_map_data(
            ["MRT", "CT", "US"],
            {
                "MRT": make_leistung_entry({"MRT": 1}),
                "CT": make_leistung_entry({"CT": 1}),
                "US": make_leistung_entry({"US": 1}),
            },
        )

        result = process(stats_data, map_data)
        assert result["total_counted"] == 3

    def test_process_metadata_passthrough(self):
        """Mitarbeiter, befunddatum, etc. passed through."""
        exams = [make_exam("MRT", 2024)]
        stats_data = make_stats_data(
            exams,
            mitarbeiter="KUNZ_A",
            befunddatum="01.01.2024-31.12.2024",
        )
        map_data = make_map_data(
            ["MRT"],
            {"MRT": make_leistung_entry({"MRT": 1})},
        )

        result = process(stats_data, map_data)
        assert result["mitarbeiter"] == "KUNZ_A"
        assert result["befunddatum"] == "01.01.2024-31.12.2024"

    def test_process_errors_accumulated(self):
        """Errors from stats_data included in output errors."""
        exams = [make_exam("MRT", 2024)]
        stats_data = make_stats_data(
            exams,
            errors=["Original error 1", "Original error 2"],
        )
        map_data = make_map_data(
            ["MRT"],
            {"MRT": make_leistung_entry({"MRT": 1})},
        )

        result = process(stats_data, map_data)
        assert "Original error 1" in result["errors"]
        assert "Original error 2" in result["errors"]

    def test_process_merged_duplicates_passthrough(self):
        """Merged duplicates from map_data passed through."""
        exams = [make_exam("MRT", 2024)]
        stats_data = make_stats_data(exams)
        map_data = make_map_data(
            ["MRT"],
            {"MRT": make_leistung_entry({"MRT": 1})},
            merged_duplicates=["CODE1", "CODE2"],
        )

        result = process(stats_data, map_data)
        assert result["merged_duplicates"] == ["CODE1", "CODE2"]

    def test_process_large_volume(self):
        """Process large volume of examinations without error."""
        # Create 1000 examinations across 10 years
        exams = []
        for year in range(2015, 2025):
            for i in range(100):
                exams.append(make_exam("MRT", year))

        stats_data = make_stats_data(exams, total_leistungen=1000)
        map_data = make_map_data(
            ["MRT"],
            {"MRT": make_leistung_entry({"MRT": 1})},
        )

        result = process(stats_data, map_data)
        assert result["total_counted"] == 1000
        assert result["data"]["MRT"]["Total"] == 1000
        assert len(result["years"]) == 10

    def test_process_empty_years_section_ordering(self):
        """Sections with zero counts still appear in map order."""
        exams = []  # No examinations
        stats_data = make_stats_data(exams, total_leistungen=0)
        map_data = make_map_data(
            ["Alpha", "Beta", "Gamma"],
            {
                "DUMMY": make_leistung_entry({"Alpha": 1, "Beta": 1, "Gamma": 1}),
            },
        )

        result = process(stats_data, map_data)
        sections_list = list(result["data"].keys())
        assert sections_list == ["Alpha", "Beta", "Gamma"]
