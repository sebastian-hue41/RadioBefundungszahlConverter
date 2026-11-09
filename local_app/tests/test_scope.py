"""Tests for local_app.scope: which department's Leistungen an evaluation includes."""

import copy

import pytest

from core import input_handler
from core.logic import process
from local_app.errors import AppError
from local_app.scope import apply_scope, is_nrad_code
from tests.helpers import make_leistung_entry, make_map_data, make_stats_data

from .helpers import stats_bytes


def exam(code, ind1=1, year=2025):
    return {"leist_kurz": code, "leistung": f"Text {code}", "year": year, "ind1": ind1, "row": 11}


ALLRAD_ROWS = [exam("RMRSCHALNA", 1), exam("KKRABILA", 2), exam("FMA2ED", 3)]
NRAD_ROWS = [exam("NMRSCHGN", 4), exam("NCTHSN", 5)]
MIXED = [ALLRAD_ROWS[0], NRAD_ROWS[0], ALLRAD_ROWS[1], NRAD_ROWS[1], ALLRAD_ROWS[2]]


def stats_of(rows, **overrides):
    stats = make_stats_data(
        rows,
        total_leistungen=len(rows),
        total_documents=len({r["ind1"] for r in rows}),
        mitarbeiter="KUNZ_A",
        befunddatum="2024",
    )
    stats.update(overrides)
    return stats


def ids(scoped):
    return [r["ind1"] for r in scoped["leistungen"]]


# ── Which codes are NRAD ───────────────────────────────────────────────────────

class TestIsNradCode:

    @pytest.mark.parametrize("code", ["NMRGAOB", "NCTHSN", "NKRTHIL", "NDSAC", "NXXZEITD", "NSCBGA", "N"])
    def test_codes_starting_with_n(self, code):
        assert is_nrad_code(code)

    @pytest.mark.parametrize("code", [
        "RMRSCHALNA", "KKRABILA", "FMA2ED", "_trenner_", "ZKRTH2E", "APCNGKQ", "RANVABS", "GUSMS", "",
    ])
    def test_other_codes(self, code):
        assert not is_nrad_code(code)

    def test_the_rule_is_case_sensitive_like_the_ris_codes(self):
        assert not is_nrad_code("nmrgaob")

    @pytest.mark.parametrize("value", [None, 5, b"NMR", ["N"]])
    def test_anything_that_is_not_a_string_is_not_nrad(self, value):
        assert not is_nrad_code(value)       # str(None) would start with "N"


# ── Selecting departments ──────────────────────────────────────────────────────

class TestApplyScope:

    def test_allrad_only_keeps_the_other_rows_in_order(self):
        # MIXED is ordered 1, 4, 2, 5, 3 — the Allrad rows keep that relative order
        assert ids(apply_scope(stats_of(MIXED), True, False)) == [1, 2, 3]

    def test_nrad_only_keeps_the_n_rows_in_order(self):
        assert ids(apply_scope(stats_of(MIXED), False, True)) == [4, 5]

    def test_the_two_selections_are_disjoint_and_make_up_all_rows(self):
        stats = stats_of(MIXED)
        allrad = apply_scope(stats, True, False)["leistungen"]
        nrad = apply_scope(stats, False, True)["leistungen"]
        assert not [r for r in allrad if r in nrad]
        assert sorted(allrad + nrad, key=lambda r: r["ind1"]) == sorted(MIXED, key=lambda r: r["ind1"])

    def test_both_keeps_every_row_and_every_total(self):
        stats = stats_of(MIXED)
        scoped = apply_scope(stats, True, True)
        assert scoped == stats
        assert scoped["leistungen"] is not stats["leistungen"]      # a copy: callers may change it

    def test_neither_is_error_e301(self):
        with pytest.raises(AppError) as info:
            apply_scope(stats_of(MIXED), False, False)
        assert info.value.code == "E301"

    @pytest.mark.parametrize("allrad, nrad", [(1, 0), (0, 1), ("yes", ""), (True, None)])
    def test_truthy_values_work_like_booleans(self, allrad, nrad):
        scoped = apply_scope(stats_of(MIXED), allrad, nrad)
        assert ids(scoped) == ([1, 2, 3] if allrad else [4, 5])

    def test_everything_else_is_carried_over(self):
        stats = stats_of(MIXED, errors=["WARNING: x"])
        scoped = apply_scope(stats, False, True)
        assert scoped["mitarbeiter"] == "KUNZ_A" and scoped["befunddatum"] == "2024"
        assert scoped["errors"] == ["WARNING: x"]

    def test_nothing_is_changed_in_the_input(self):
        stats = stats_of(MIXED, errors=["WARNING: x"])
        before = copy.deepcopy(stats)
        for allrad, nrad in [(True, False), (False, True), (True, True)]:
            scoped = apply_scope(stats, allrad, nrad)
            scoped["errors"].append("AUSWAHL: a note")
            scoped["leistungen"].clear()
        assert stats == before


class TestTotals:

    def test_rows_removed_lower_the_expected_count_by_exactly_that_many(self):
        assert apply_scope(stats_of(MIXED), True, False)["total_leistungen"] == 3
        assert apply_scope(stats_of(MIXED), False, True)["total_leistungen"] == 2

    def test_rows_skipped_while_loading_stay_visible(self):
        # 6 rows in the file, 5 loaded, 1 skipped: the difference survives the selection
        scoped = apply_scope(stats_of(MIXED, total_leistungen=6), False, True)
        assert scoped["total_leistungen"] == 3 and len(scoped["leistungen"]) == 2

    def test_nothing_removed_nothing_changed(self):
        stats = stats_of(NRAD_ROWS, total_leistungen=7, total_documents=99)
        scoped = apply_scope(stats, False, True)
        assert scoped["total_leistungen"] == 7 and scoped["total_documents"] == 99

    def test_the_expected_count_is_never_negative(self):
        assert apply_scope(stats_of(MIXED, total_leistungen=1), True, False)["total_leistungen"] == 0

    def test_unknown_expected_count_stays_unknown(self):
        assert apply_scope(stats_of(MIXED, total_leistungen=None), True, False)["total_leistungen"] is None

    def test_reports_are_counted_again_from_the_distinct_ind1_of_the_remaining_rows(self):
        rows = [exam("NMRSCHGN", 1), exam("RMRSCHALNA", 1), exam("RMRSCHALNA", 2), exam("RKRTHPA", 2), exam("RKRTHPA", 3)]
        stats = stats_of(rows, total_documents=3)
        assert apply_scope(stats, False, True)["total_documents"] == 1
        assert apply_scope(stats, True, False)["total_documents"] == 3
        assert apply_scope(stats, True, True)["total_documents"] == 3

    def test_a_row_without_ind1_makes_the_report_count_unknown(self):
        rows = [exam("NMRSCHGN", 1), exam("NCTHSN", None), exam("RMRSCHALNA", 2)]
        assert apply_scope(stats_of(rows, total_documents=2), False, True)["total_documents"] is None

    def test_removed_rows_without_ind1_do_not_matter(self):
        rows = [exam("NMRSCHGN", 1), exam("RMRSCHALNA", None)]
        assert apply_scope(stats_of(rows, total_documents=2), False, True)["total_documents"] == 1

    def test_a_file_without_l8_does_not_get_a_report_count_made_up(self):
        assert apply_scope(stats_of(MIXED, total_documents=None), False, True)["total_documents"] is None


class TestNothingLeft:

    def test_nrad_only_on_an_allrad_statistics(self):
        with pytest.raises(AppError) as info:
            apply_scope(stats_of(ALLRAD_ROWS), False, True)
        assert info.value.code == "E303"
        assert info.value.title == "Diese Statistik enthält keine NRAD-Leistungen."
        assert info.value.slot is None

    def test_allrad_only_on_an_nrad_statistics(self):
        with pytest.raises(AppError) as info:
            apply_scope(stats_of(NRAD_ROWS), True, False)
        assert info.value.code == "E303"
        assert info.value.title == "Diese Statistik enthält keine Allrad-Leistungen."

    def test_both_on_an_empty_statistics(self):
        with pytest.raises(AppError) as info:
            apply_scope(stats_of([]), True, True)
        assert info.value.code == "E303" and "gewählten Bereiche" in info.value.title


# ── Together with the real loaders and core's process() ────────────────────────

SECTIONS = ["MRT"]
MAP = make_map_data(
    SECTIONS,
    {"RMRSCHALNA": make_leistung_entry({"MRT": 1}), "NMRSCHGN": make_leistung_entry({"MRT": 1})},
)
THREE_ROWS = [
    (1, 1, "01.01.2024", "Schädel nativ NRAD", "NMRSCHGN"),
    (1, 2, "01.01.2024", "Schädel nativ", "RMRSCHALNA"),
    (2, 1, "01.01.2024", "Schädel nativ", "RMRSCHALNA"),
]


def mismatch(result):
    return [e for e in result["errors"] if e.startswith("COUNT MISMATCH")]


class TestWithProcess:

    def load(self, rows, **kwargs):
        return input_handler.load_statistics(stats_bytes(rows, **kwargs))

    def test_the_loader_counts_what_the_file_holds(self):
        stats = self.load(THREE_ROWS, total_l8=2)
        assert stats["total_leistungen"] == 3 and len(stats["leistungen"]) == 3

    @pytest.mark.parametrize("allrad, nrad, counted, documents", [
        (True, False, 2, 2),
        (False, True, 1, 1),
        (True, True, 3, 2),
    ])
    def test_no_selection_causes_a_count_mismatch(self, allrad, nrad, counted, documents):
        result = process(apply_scope(self.load(THREE_ROWS, total_l8=2), allrad, nrad), MAP)
        assert mismatch(result) == []
        assert result["total_counted"] == counted
        assert result["total_documents"] == documents

    def test_without_adjusting_the_expected_count_the_mismatch_would_appear(self):
        # The reason apply_scope lowers total_leistungen: the plain filter alone would be flagged.
        stats = self.load(THREE_ROWS, total_l8=2)
        filtered = dict(stats, leistungen=[r for r in stats["leistungen"] if is_nrad_code(r["leist_kurz"])])
        assert len(mismatch(process(filtered, MAP))) == 1

    def test_a_row_that_was_really_skipped_is_still_reported(self):
        rows = THREE_ROWS + [(3, 1, "01.01.2024", "ohne Code", None)]
        stats = self.load(rows, total_l8=3)
        assert stats["total_leistungen"] == 4 and len(stats["leistungen"]) == 3
        for allrad, nrad, left in [(True, False, 2), (False, True, 1), (True, True, 3)]:
            (message,) = mismatch(process(apply_scope(stats, allrad, nrad), MAP))
            assert f"{left} successfully processed (1 rows skipped)" in message

    def test_the_other_departments_codes_are_not_reported_as_missing_from_the_map(self):
        stats = self.load(THREE_ROWS, total_l8=2)
        stats["leistungen"].append(exam("NXXZEITD", 9))           # an N code the Allrad map does not know
        stats["total_leistungen"] += 1
        unscoped = process(stats, MAP)
        assert any("NXXZEITD" in e for e in unscoped["errors"])
        allrad_only = process(apply_scope(stats, True, False), MAP)
        assert not any("NXXZEITD" in e for e in allrad_only["errors"])

    def test_counts_per_section_follow_the_selection(self):
        stats = self.load(THREE_ROWS, total_l8=2)
        totals = {
            (a, n): process(apply_scope(stats, a, n), MAP)["data"]["MRT"]["Total"]
            for a, n in [(True, False), (False, True), (True, True)]
        }
        assert totals == {(True, False): 2, (False, True): 1, (True, True): 3}


# ── Two maps: the swapped-file guard and the merge ─────────────────────────────

import random  # noqa: E402

from local_app.errors import ALLRAD, NRAD  # noqa: E402
from local_app.scope import check_department, merge_maps  # noqa: E402

SECTIONS = ["MRT", "CT", "Rö Skelett"]


def entry(**values):
    """A map entry assigning the given section values; the others are 0."""
    return make_leistung_entry({name: values.get(name, 0) for name in SECTIONS})


def allrad_map(**extra):
    leistungen = {
        "RMRSCHALNA": entry(MRT=1),
        "RCTTHO": entry(CT=1),
        "RRSKELETT": entry(**{"Rö Skelett": 1}),
        **extra,
    }
    return make_map_data(list(SECTIONS), leistungen)


def nrad_map(**extra):
    leistungen = {
        "NMRSCHGN": entry(MRT=2),
        "NCTHSN": entry(CT=1),
        **extra,
    }
    return make_map_data(list(SECTIONS), leistungen)


class TestCheckDepartment:

    def test_the_right_maps_pass(self):
        check_department(allrad_map(), ALLRAD)
        check_department(nrad_map(), NRAD)

    def test_an_allrad_map_with_a_few_n_codes_passes(self):
        check_department(allrad_map(NMRGAOB=entry(MRT=1), NCTHSN=entry(CT=1)), ALLRAD)

    def test_an_nrad_map_with_a_few_other_codes_passes(self):
        check_department(nrad_map(RMRSCHALNA=entry(MRT=1)), NRAD)

    def test_an_nrad_map_in_the_allrad_field(self):
        with pytest.raises(AppError) as info:
            check_department(nrad_map(), ALLRAD)
        assert info.value.code == "E220" and info.value.slot == ALLRAD
        assert "2 von 2" in info.value.detail and "NRAD-Map" in info.value.detail

    def test_an_allrad_map_in_the_nrad_field(self):
        with pytest.raises(AppError) as info:
            check_department(allrad_map(), NRAD)
        assert info.value.code == "E220" and info.value.slot == NRAD
        assert "0 von 3" in info.value.detail and "Allrad-Map" in info.value.detail

    @pytest.mark.parametrize("n, total, ok_for", [
        (0, 10, ALLRAD), (4, 10, ALLRAD), (49, 100, ALLRAD),      # below half
        (5, 10, NRAD), (6, 10, NRAD), (10, 10, NRAD),             # half or more
    ])
    def test_the_limit_is_half_of_the_codes(self, n, total, ok_for):
        leistungen = {("N" if i < n else "R") + f"CODE{i}": entry(MRT=1) for i in range(total)}
        data = make_map_data(list(SECTIONS), leistungen)
        other = NRAD if ok_for == ALLRAD else ALLRAD
        check_department(data, ok_for)
        with pytest.raises(AppError) as info:
            check_department(data, other)
        assert info.value.code == "E220" and info.value.slot == other

    def test_the_first_letter_decides_like_everywhere_else(self):
        data = make_map_data(list(SECTIONS), {"nmr": entry(MRT=1), "nct": entry(CT=1)})     # lower case: not NRAD
        check_department(data, ALLRAD)

    def test_a_slot_that_is_not_a_map_slot(self):
        for slot in ("statistics", "elsewhere", None):
            with pytest.raises(ValueError):
                check_department(allrad_map(), slot)

    def test_the_map_is_not_changed(self):
        data = nrad_map()
        before = copy.deepcopy(data)
        check_department(data, NRAD)
        assert data == before


class TestMergeMaps:

    def test_the_nrad_entry_replaces_the_allrad_entry_of_the_same_n_code(self):
        allrad = allrad_map(NMRSCHGN=entry(MRT=1))               # the Allrad map values it 1, the NRAD map 2
        merged = merge_maps(allrad, nrad_map())
        assert merged["leistungen"]["NMRSCHGN"]["sections"]["MRT"] == 2

    def test_other_allrad_entries_are_kept(self):
        merged = merge_maps(allrad_map(), nrad_map())
        assert {"RMRSCHALNA", "RCTTHO", "RRSKELETT"} <= set(merged["leistungen"])
        assert merged["leistungen"]["RMRSCHALNA"] == allrad_map()["leistungen"]["RMRSCHALNA"]

    def test_all_nrad_entries_are_added(self):
        merged = merge_maps(allrad_map(), nrad_map())
        assert {"NMRSCHGN", "NCTHSN"} <= set(merged["leistungen"])

    def test_an_n_code_only_the_allrad_map_has_is_not_carried_over(self):
        # N codes belong to the NRAD map; otherwise "Allrad + NRAD = both" would not hold exactly.
        merged = merge_maps(allrad_map(NZUNRAD=entry(MRT=1)), nrad_map())
        assert "NZUNRAD" not in merged["leistungen"]

    def test_a_non_n_code_only_the_nrad_map_has_is_not_carried_over(self):
        merged = merge_maps(allrad_map(), nrad_map(RIMNRAD=entry(MRT=1)))
        assert "RIMNRAD" not in merged["leistungen"]

    def test_the_sections_are_those_of_the_allrad_map_in_its_order(self):
        reordered = make_map_data(["CT", "Rö Skelett", "MRT"], nrad_map()["leistungen"])
        assert merge_maps(allrad_map(), reordered)["sections"] == SECTIONS

    def test_the_inputs_are_not_changed_and_the_result_shares_nothing_with_them(self):
        allrad, nrad = allrad_map(NMRSCHGN=entry(MRT=1)), nrad_map()
        before = copy.deepcopy((allrad, nrad))
        merged = merge_maps(allrad, nrad)
        assert (allrad, nrad) == before
        merged["leistungen"]["RMRSCHALNA"]["sections"]["MRT"] = 99
        merged["leistungen"]["NMRSCHGN"]["sections"]["MRT"] = 99
        merged["sections"].append("X")
        assert (allrad, nrad) == before

    def test_merged_duplicates_are_combined_without_repeats(self):
        allrad = allrad_map()
        allrad["merged_duplicates"] = ["RMRSCHALNA", "RCTTHO", "RMRSCHALNA"]
        nrad = nrad_map()
        nrad["merged_duplicates"] = ["NCTHSN", "NMRSCHGN"]
        assert merge_maps(allrad, nrad)["merged_duplicates"] == ["RMRSCHALNA", "RCTTHO", "NCTHSN", "NMRSCHGN"]

    def test_a_duplicate_of_a_code_that_was_not_carried_over_is_not_listed(self):
        allrad = allrad_map(NZUNRAD=entry(MRT=1))
        allrad["merged_duplicates"] = ["NZUNRAD", "RCTTHO"]
        assert merge_maps(allrad, nrad_map())["merged_duplicates"] == ["RCTTHO"]

    def test_maps_without_a_merged_duplicates_key(self):
        allrad, nrad = allrad_map(), nrad_map()
        del allrad["merged_duplicates"], nrad["merged_duplicates"]
        assert merge_maps(allrad, nrad)["merged_duplicates"] == []

    @pytest.mark.parametrize("nrad_sections", [
        ["MRT", "CT"],                              # one fewer
        ["MRT", "CT", "Rö Skelett", "Neu"],         # one more
        ["MRT", "CT", "Rö Skeletts"],               # one renamed
        [],
    ])
    def test_different_section_lists_are_refused(self, nrad_sections):
        with pytest.raises(AppError) as info:
            merge_maps(allrad_map(), make_map_data(nrad_sections, {"NMRSCHGN": entry(MRT=2)}))
        assert info.value.code == "E221" and info.value.slot is None

    def test_the_detail_names_the_sections_that_differ(self):
        with pytest.raises(AppError) as info:
            merge_maps(allrad_map(), make_map_data(["MRT", "CT", "Neu"], {"NMRSCHGN": entry(MRT=2)}))
        assert "Nur in der Allrad-Map: 'Rö Skelett'" in info.value.detail
        assert "Nur in der NRAD-Map: 'Neu'" in info.value.detail

    def test_a_long_difference_is_shortened(self):
        many = [f"S{i}" for i in range(12)]
        with pytest.raises(AppError) as info:
            merge_maps(make_map_data(many, {"RX": {"kurztext": "", "sections": {}, "row": 2}}),
                       make_map_data(["MRT"], {"NX": {"kurztext": "", "sections": {}, "row": 2}}))
        assert "und 7 weitere" in info.value.detail

    def test_the_result_works_in_process(self):
        merged = merge_maps(allrad_map(), nrad_map())
        stats = stats_of([exam("RMRSCHALNA", 1), exam("NMRSCHGN", 2), exam("NCTHSN", 3), exam("UNBEKANNT", 4)])
        result = process(apply_scope(stats, True, True), merged)
        assert result["data"]["MRT"]["Total"] == 1 + 2
        assert result["data"]["CT"]["Total"] == 1
        assert "Text UNBEKANNT" in result["data"]
        assert not any(msg.startswith("COUNT MISMATCH") for msg in result["errors"])


def random_map(rng, sections, codes):
    return make_map_data(list(sections), {
        code: make_leistung_entry({s: rng.choice([0, 0, 1, 2]) for s in sections}) for code in codes
    })


class TestAllradPlusNradIsBoth:
    """The central promise of the checkboxes, tested on many random maps and statistics."""

    @pytest.mark.parametrize("seed", range(60))
    def test_every_section_and_year_adds_up(self, seed):
        rng = random.Random(seed)
        sections = [f"S{i}" for i in range(rng.randint(1, 6))]
        allrad_codes = [f"R{i}" for i in range(rng.randint(3, 12))] + [f"N{i}" for i in range(rng.randint(0, 5))]
        nrad_codes = [f"N{i}" for i in range(rng.randint(2, 9))]
        allrad, nrad = random_map(rng, sections, allrad_codes), random_map(rng, sections, nrad_codes)
        pool = [f"R{i}" for i in range(15)] + [f"N{i}" for i in range(12)]      # some codes are in no map
        rows = [exam(rng.choice(pool), ind1=i, year=rng.choice([2023, 2024, 2025])) for i in range(rng.randint(1, 60))]
        stats = stats_of(rows)

        both = process(apply_scope(stats, True, True), merge_maps(allrad, nrad))
        parts = []
        for include_allrad, include_nrad, map_data in ((True, False, allrad), (False, True, nrad)):
            try:
                parts.append(process(apply_scope(stats, include_allrad, include_nrad), map_data))
            except AppError as exc:
                assert exc.code == "E303"           # this selection has no rows; it adds nothing
        assert both["total_counted"] == sum(p["total_counted"] for p in parts)
        for name in both["data"]:
            for year in both["years"] + ["Total"]:
                assert both["data"][name].get(year, 0) == sum(p["data"].get(name, {}).get(year, 0) for p in parts), (name, year)
        assert {name for p in parts for name in p["data"]} <= set(both["data"])
