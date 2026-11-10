"""Tests for local_app.server: bundle_dir() and evaluate() (the work of one upload, without HTTP)."""

import io
import os
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import openpyxl
import pytest

from core import input_handler, output_handler
from core.input_handler import load_reference_data
from local_app import server
from local_app.errors import ALLRAD, NRAD, SLOT_LABELS, STATISTICS, AppError
from local_app.server import Evaluation, bundle_dir, evaluate
from tests.helpers import make_reference_wb, wb_to_bytes

from .helpers import force_text, map_bytes, stats_bytes
from .test_security import HOSTILE, series

SECTIONS = ["DL Gefäße", "MRT", "CT"]
# Allrad map: MRT 1 for RMRT, CT 1 for RCT, and an entry for the N code NMRT that the NRAD map values differently.
ALLRAD_CODES = [("RMRT", "MRT Schädel", {"MRT": 1}), ("RCT", "CT Thorax", {"CT": 1}), ("NMRT", "N MRT (Allrad)", {"MRT": 1})]
NRAD_CODES = [("NMRT", "N MRT", {"MRT": 2}), ("NCT", "N CT", {"CT": 1})]


def exams(*codes):
    return [(i, 1, "01.01.2024", f"Text {code}", code) for i, code in enumerate(codes, 1)]


MIXED = exams("RMRT", "RMRT", "RCT", "NMRT", "NCT", "NCT", "NCT")       # 3 Allrad rows, 4 N rows


def make_files(rows=MIXED, allrad=True, nrad=True, **stats_kwargs):
    files = {STATISTICS: stats_bytes(rows, total_l8=len(rows), **stats_kwargs)}
    if allrad:
        files[ALLRAD] = map_bytes(ALLRAD_CODES, sections=SECTIONS)
    if nrad:
        files[NRAD] = map_bytes(NRAD_CODES, sections=SECTIONS)
    return files


REFERENCE = wb_to_bytes(make_reference_wb({"MRT": 100}))


def sheet(evaluation_or_bytes):
    data = evaluation_or_bytes.data if isinstance(evaluation_or_bytes, Evaluation) else evaluation_or_bytes
    return openpyxl.load_workbook(io.BytesIO(data)).active


def totals(evaluation) -> dict:
    """section -> Gesamt, read back from the Excel."""
    ws = sheet(evaluation)
    header = [c.value for c in ws[1]]
    gesamt = header.index("Gesamt")
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None or str(row[0]).startswith("Gezählte"):
            break
        out[row[0]] = row[gesamt]
    return out


def notes(evaluation) -> list[str]:
    """The lines under "Hinweise und Fehler" in the Excel."""
    column = [c.value for c in sheet(evaluation)["A"]]
    if "Hinweise und Fehler" not in column:
        return []
    return [str(v) for v in column[column.index("Hinweise und Fehler") + 1:] if v is not None]


def codes_of(evaluation) -> list[str]:
    return [h["code"] for h in evaluation.summary["hints"]]


# ── Each selection ─────────────────────────────────────────────────────────────

class TestSelections:

    def test_allrad_only(self):
        result = evaluate(make_files(nrad=False))
        assert totals(result)["MRT"] == 2 and totals(result)["CT"] == 1          # the N rows are not counted

    def test_nrad_only(self):
        result = evaluate(make_files(allrad=False))
        assert totals(result)["MRT"] == 2 and totals(result)["CT"] == 3          # NMRT is worth 2 in the NRAD map

    def test_both_add_up_with_the_nrad_entry_winning_for_n_codes(self):
        result = evaluate(make_files())
        assert totals(result)["MRT"] == 2 + 2 and totals(result)["CT"] == 1 + 3
        # NMRT is worth 1 in the Allrad map and 2 in the NRAD map; the NRAD map wins

    def test_allrad_plus_nrad_is_both_in_every_section(self):
        both = totals(evaluate(make_files()))
        a, n = totals(evaluate(make_files(nrad=False))), totals(evaluate(make_files(allrad=False)))
        assert {s: a.get(s, 0) + n.get(s, 0) for s in both} == both

    def test_the_counted_leistungen_follow_the_selection(self):
        assert evaluate(make_files(nrad=False)).summary["total_counted"] == 3
        assert evaluate(make_files(allrad=False)).summary["total_counted"] == 4
        assert evaluate(make_files()).summary["total_counted"] == 7

    def test_the_summary_names_the_selection(self):
        assert evaluate(make_files(nrad=False)).summary["selection"] == [ALLRAD]
        assert evaluate(make_files(allrad=False)).summary["selection_label"] == "NRAD"
        both = evaluate(make_files()).summary
        assert both["selection"] == [ALLRAD, NRAD] and both["selection_label"] == "Allrad + NRAD"

    def test_an_n_code_only_the_allrad_map_has_shows_as_not_in_map_when_both_are_chosen(self):
        files = make_files(exams("NZUNRAD", "RMRT"))
        extra = [("NZUNRAD", "nur Allrad", {"MRT": 5}), ("RX1", "x", {"CT": 1}), ("RX2", "x", {"CT": 1})]    # stays an Allrad map
        files[ALLRAD] = map_bytes(ALLRAD_CODES + extra, sections=SECTIONS)
        result = evaluate(files)
        assert totals(result)["MRT"] == 1 and "Text NZUNRAD" in totals(result)
        assert "W201" in codes_of(result)

    def test_the_result_is_an_evaluation_with_bytes_and_a_name(self):
        result = evaluate(make_files())
        assert isinstance(result, Evaluation) and isinstance(result.data, bytes) and result.data[:2] == b"PK"
        assert result.summary["filename"] == result.filename

    def test_the_summary_carries_what_the_page_shows(self):
        summary = evaluate(make_files(mitarbeiter="ZUFALL_N", befunddatum="01.04.2024-31.12.9999")).summary
        assert summary["mitarbeiter"] == "ZUFALL_N" and summary["befunddatum"] == "01.04.2024-31.12.9999"
        assert set(summary) == {"filename", "mitarbeiter", "befunddatum", "total_counted", "selection", "selection_label", "hints", "preview"}


def excel_table(evaluation):
    """(header, rows) of the main table, read back from the finished Excel."""
    ws = sheet(evaluation)
    rows = [list(row) for row in ws.iter_rows(values_only=True)]
    header = [v for v in rows[0]]
    while header and header[-1] is None:
        header.pop()
    body = []
    for row in rows[1:]:
        if row[0] is None or str(row[0]).startswith("Gezählte"):
            break
        body.append(row[: len(header)])
    return header, body


class TestPreview:
    """The table the page shows before the download is the table of the Excel itself."""

    @pytest.mark.parametrize("allrad, nrad", [(True, False), (False, True), (True, True)])
    def test_it_is_the_table_of_the_excel_for_every_selection(self, allrad, nrad):
        result = evaluate(make_files(allrad=allrad, nrad=nrad), reference_data=load_reference_data(REFERENCE))
        header, body = excel_table(result)
        assert result.summary["preview"]["columns"] == header
        assert result.summary["preview"]["rows"] == [list(row) for row in body]

    def test_the_total_and_the_target_column_are_named_by_their_position(self):
        preview = evaluate(make_files(), reference_data=load_reference_data(REFERENCE)).summary["preview"]
        assert preview["columns"][preview["total_column"]] == "Gesamt" and preview["columns"][preview["target_column"]] == "Benötigt"
        assert preview["columns"][0] == "Leistungsbereich" and preview["columns"][1] == "2024"

    def test_without_reference_data_there_is_no_target_column(self):
        preview = evaluate(make_files()).summary["preview"]
        assert preview["target_column"] is None and "Benötigt" not in preview["columns"] and preview["total_column"] is not None

    def test_numbers_stay_numbers_and_blanks_are_none(self):
        preview = evaluate(make_files(), reference_data=load_reference_data(REFERENCE)).summary["preview"]
        rows = {row[0]: row for row in preview["rows"]}
        assert rows["MRT"][1] == 4 and rows["MRT"][preview["total_column"]] == 4 and isinstance(rows["MRT"][1], int)
        assert all(row[preview["target_column"]] is None or isinstance(row[preview["target_column"]], int) for row in preview["rows"])

    def test_a_target_from_the_reference_file_arrives_next_to_the_total(self):
        preview = evaluate(make_files(), reference_data=load_reference_data(REFERENCE)).summary["preview"]
        assert {row[0]: row[preview["target_column"]] for row in preview["rows"]}["MRT"] == 100          # REFERENCE: {"MRT": 100}

    def test_a_combination_row_of_the_reference_file_is_in_it(self):
        reference = load_reference_data(wb_to_bytes(make_reference_wb({"MRT": 100, "CT": 50}, combinations={"MRT und CT": ["MRT", "CT"]})))
        preview = evaluate(make_files(), reference_data=reference).summary["preview"]
        assert {row[0]: row[preview["total_column"]] for row in preview["rows"]}["MRT und CT"] == 4 + 4

    def test_rows_that_are_not_in_the_map_are_in_it_as_in_the_excel(self):
        preview = evaluate(make_files(exams("RMRT", "RUNBEKANNT"), nrad=False)).summary["preview"]
        assert "Text RUNBEKANNT" in [row[0] for row in preview["rows"]]

    def test_nothing_below_the_table_is_in_it(self):
        text = json.dumps(evaluate(make_files(), names={ALLRAD: "geheim.xlsx", NRAD: "auch.xlsx"}, reference_data=load_reference_data(REFERENCE)).summary["preview"])
        for needle in ("Gezählte", "AUSWAHL", "Hinweise und Fehler", "geheim", "auch.xlsx", "Befunddokumente"):
            assert needle not in text

    def test_a_section_name_with_markup_stays_plain_text(self):
        evil = '<img src=x onerror="alert(1)">'
        files = make_files(nrad=False)
        files[ALLRAD] = map_bytes(ALLRAD_CODES, sections=["DL Gefäße", evil, "CT"])
        result = evaluate(files)
        assert evil in [row[0] for row in result.summary["preview"]["rows"]]
        assert all(isinstance(v, (str, int, float, type(None))) for row in result.summary["preview"]["rows"] for v in row)

    def test_a_long_name_is_clipped(self):
        files = make_files(nrad=False)
        files[ALLRAD] = map_bytes(ALLRAD_CODES, sections=["DL Gefäße", "M" * 500, "CT"])
        names = [row[0] for row in evaluate(files).summary["preview"]["rows"]]
        assert max(len(n) for n in names) <= 100

    def test_the_number_of_rows_is_limited(self, monkeypatch):
        monkeypatch.setattr(server, "_PREVIEW_MAX_ROWS", 2)
        assert len(evaluate(make_files()).summary["preview"]["rows"]) == 2

    def test_the_number_of_columns_is_limited(self, monkeypatch):
        monkeypatch.setattr(server, "_PREVIEW_MAX_COLUMNS", 2)
        preview = evaluate(make_files()).summary["preview"]
        assert len(preview["columns"]) == 2 and preview["total_column"] is None and all(len(row) == 2 for row in preview["rows"])

    def test_it_does_not_change_the_download(self):
        result = evaluate(make_files(), reference_data=load_reference_data(REFERENCE))
        assert openpyxl.load_workbook(io.BytesIO(result.data)).active.max_row > 5 and result.data[:2] == b"PK"

    def test_it_is_small(self):
        assert len(json.dumps(evaluate(make_files(), reference_data=load_reference_data(REFERENCE)).summary)) < 20_000


class TestFileNames:

    @pytest.mark.parametrize("allrad, nrad, suffix", [(True, False, "_Allrad"), (False, True, "_NRAD"), (True, True, "_Allrad_NRAD")])
    def test_the_suffix_follows_the_selection(self, allrad, nrad, suffix):
        name = evaluate(make_files(allrad=allrad, nrad=nrad)).filename
        assert name == f"KUNZ_A01012024-31122024Auswertung{suffix}.xlsx"

    def test_the_name_is_safe(self):
        name = evaluate(make_files(mitarbeiter='../../Ev<il>"x', befunddatum="01.01.2024")).filename
        assert name.endswith("_Allrad_NRAD.xlsx") and all(c.isalnum() or c in " .-_()" for c in name)

    def test_without_staff_name_and_date_the_name_is_general(self):
        result = evaluate(make_files(mitarbeiter=None, befunddatum=None))
        assert result.filename.startswith("Auswertung_") and result.filename.endswith("_Allrad_NRAD.xlsx")
        assert result.summary["mitarbeiter"] is None and result.summary["befunddatum"] is None

    def test_a_name_from_core_is_cleaned_again(self, monkeypatch):
        real = server.build_xlsx_bytes
        monkeypatch.setattr(server, "build_xlsx_bytes", lambda *a, **k: ('../ev<il>"x.xlsx', real(*a, **k)[1]))
        name = evaluate(make_files(nrad=False)).filename
        assert name == "evilx_Allrad.xlsx"


# ── The note on the sheet and the hints ────────────────────────────────────────

class TestNotesAndHints:

    def test_the_selection_and_the_map_names_are_noted_on_the_sheet(self):
        result = evaluate(make_files(), names={ALLRAD: "mapped_leistungen_2611.xlsx", NRAD: "mapped_leistungen_nrad.xlsx"})
        note = [n for n in notes(result) if n.startswith("AUSWAHL:")][0]
        assert "Allrad + NRAD" in note
        assert "Allrad-Map „mapped_leistungen_2611.xlsx“" in note and "NRAD-Map „mapped_leistungen_nrad.xlsx“" in note

    def test_only_the_chosen_maps_are_named(self):
        note = [n for n in notes(evaluate(make_files(nrad=False), names={ALLRAD: "a.xlsx", NRAD: "n.xlsx"})) if n.startswith("AUSWAHL:")][0]
        assert "„a.xlsx“" in note and "n.xlsx" not in note

    def test_without_names_the_note_says_only_which_maps(self):
        note = [n for n in notes(evaluate(make_files())) if n.startswith("AUSWAHL:")][0]
        assert "Allrad-Map; NRAD-Map." in note and "„" not in note

    def test_unsafe_characters_in_names_are_stripped(self):
        result = evaluate(make_files(nrad=False), names={ALLRAD: '../ev<il>"x.xlsx'})
        text = " ".join(notes(result))
        assert "„evilx.xlsx“" in text and "<" not in text and "../" not in text

    def test_a_name_that_looks_like_a_formula_stays_text(self):
        result = evaluate(make_files(nrad=False), names={ALLRAD: "=HYPERLINK(\"http://evil.example\",\"x\")"})
        assert "evil.example" not in " ".join(notes(result)) or "=" not in " ".join(notes(result))
        assert all(isinstance(c.value, str) or c.value is None or isinstance(c.value, (int, float)) for row in sheet(result).iter_rows() for c in row)
        assert not any(c.data_type == "f" for row in sheet(result).iter_rows() for c in row)

    def test_the_selection_note_is_not_a_hint(self):
        assert not any(h["code"] == "W900" for h in evaluate(make_files()).summary["hints"])

    def test_w401_reports_the_rows_left_out(self):
        allrad_only = evaluate(make_files(nrad=False))
        assert {"code": "W401", "text": "4 Leistungen der anderen Abteilung wurden nicht ausgewertet.", "count": 4} in allrad_only.summary["hints"]
        assert any("4 Leistungen der anderen Abteilung" in n for n in notes(allrad_only))
        nrad_only = evaluate(make_files(allrad=False))
        assert {"code": "W401", "text": "3 Leistungen der anderen Abteilung wurden nicht ausgewertet.", "count": 3} in nrad_only.summary["hints"]

    def test_w401_is_absent_when_both_are_chosen(self):
        both = evaluate(make_files())
        assert "W401" not in codes_of(both)
        assert not any("anderen Abteilung" in n for n in notes(both))

    def test_w401_is_singular_for_one_row(self):
        files = make_files(exams("RMRT", "NMRT"), nrad=False)
        assert evaluate(files).summary["hints"][-1]["text"] == "1 Leistung der anderen Abteilung wurde nicht ausgewertet."

    def test_no_count_mismatch_is_reported_for_the_rows_left_out_on_purpose(self):
        result = evaluate(make_files(nrad=False))
        assert "W104" not in codes_of(result)
        assert not any(n.startswith("COUNT MISMATCH") for n in notes(result))

    def test_a_row_that_was_really_skipped_is_still_reported(self):
        rows = exams("RMRT", "NMRT") + [(3, 1, "01.01.2024", "ohne Code", None)]
        result = evaluate(make_files(rows, nrad=False))
        assert "W101" in codes_of(result) or "W104" in codes_of(result)

    def test_codes_missing_in_the_map_are_a_hint(self):
        result = evaluate(make_files(exams("RMRT", "RUNBEKANNT", "RANDERE"), nrad=False))
        assert {"code": "W201", "text": "2 Leistungscodes fehlen in der Map (eigene Zeilen in der Excel).", "count": 2} in result.summary["hints"]

    def test_the_hints_are_sorted_by_code(self):
        codes = codes_of(evaluate(make_files(exams("RMRT", "NMRT", "RUNBEKANNT"), nrad=False)))
        assert codes == sorted(codes) and len(codes) >= 3

    def test_a_message_nobody_translated_is_sorted_in_too(self, monkeypatch):
        real = server.process

        def with_unknown_message(stats, map_data):
            result = real(stats, map_data)
            result["errors"].append("a message nobody translated")
            return result

        monkeypatch.setattr(server, "process", with_unknown_message)
        codes = codes_of(evaluate(make_files(nrad=False)))                  # W900 (from the messages) before W301 and W401 (added here)
        assert "W900" in codes and codes == sorted(codes)


class TestReference:

    def test_benoetigt_appears_with_reference_data(self):
        result = evaluate(make_files(), reference_data=load_reference_data(REFERENCE))
        assert "Benötigt" in [c.value for c in sheet(result)[1]]
        assert not {"W301", "W302"} & set(codes_of(result))

    def test_without_reference_data_w301_is_reported_and_the_evaluation_works(self):
        for reference in (None, {}):
            result = evaluate(make_files(), reference_data=reference)
            assert "Benötigt" not in [c.value for c in sheet(result)[1]]
            assert "W301" in codes_of(result) and totals(result)["MRT"] == 4

    def test_a_reference_file_that_was_rejected_gives_w301_and_its_reason_is_on_the_sheet(self):
        reference = load_reference_data(b"not a workbook")
        result = evaluate(make_files(), reference_data=reference)
        assert "W301" in codes_of(result) and "W302" not in codes_of(result)
        assert "Hinweis: Probleme mit der Referenzdatei" in [c.value for c in sheet(result)["A"]]

    def test_a_partly_usable_reference_gives_w302_and_keeps_the_column(self):
        wb = make_reference_wb({"MRT": 100, "CT": -5})
        reference = load_reference_data(wb_to_bytes(wb))
        assert reference["problems"]
        result = evaluate(make_files(), reference_data=reference)
        assert "W302" in codes_of(result) and "W301" not in codes_of(result)
        assert "Benötigt" in [c.value for c in sheet(result)[1]]

    def test_no_path_or_file_content_in_the_hint(self):
        reference = load_reference_data("/geheim/verzeichnis/reference_amount.xlsx")
        texts = " ".join(h["text"] for h in evaluate(make_files(), reference_data=reference).summary["hints"])
        assert "geheim" not in texts and "reference_amount" not in texts


# ── Failing files ──────────────────────────────────────────────────────────────

class TestFailingFiles:

    @pytest.mark.parametrize("slot", [STATISTICS, ALLRAD, NRAD])
    def test_not_an_excel_file_in_each_slot(self, slot):
        files = make_files()
        files[slot] = b"this is not a workbook"
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == series(slot) + "01" and info.value.slot == slot

    @pytest.mark.parametrize("slot", [STATISTICS, ALLRAD, NRAD])
    def test_an_empty_file_in_each_slot(self, slot):
        files = make_files()
        files[slot] = b""
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == series(slot) + "02" and info.value.slot == slot

    def test_a_map_in_the_statistics_slot(self):
        files = make_files()
        files[STATISTICS] = files[ALLRAD]
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == "E110" and info.value.slot == STATISTICS

    @pytest.mark.parametrize("slot", [ALLRAD, NRAD])
    def test_the_statistics_in_a_map_slot(self, slot):
        files = make_files()
        files[slot] = files[STATISTICS]
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == "E210" and info.value.slot == slot

    def test_the_first_problem_in_the_order_of_the_fields_is_reported(self):
        files = make_files()
        files[STATISTICS], files[ALLRAD], files[NRAD] = b"x", b"y", b"z"
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.slot == STATISTICS

    def test_a_statistics_file_without_rows(self):
        with pytest.raises(AppError) as info:
            evaluate(make_files(rows=[]))
        assert info.value.code == "E112" and info.value.slot == STATISTICS

    def test_a_map_without_codes(self):
        files = make_files()
        files[NRAD] = map_bytes([], sections=SECTIONS)
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == "E211" and info.value.slot == NRAD

    def test_swapped_maps(self):
        files = make_files()
        files[ALLRAD], files[NRAD] = map_bytes(NRAD_CODES, sections=SECTIONS), map_bytes(ALLRAD_CODES, sections=SECTIONS)
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == "E220" and info.value.slot == ALLRAD

    def test_a_single_map_in_the_wrong_field(self):
        files = make_files(nrad=False)
        files[ALLRAD] = map_bytes(NRAD_CODES, sections=SECTIONS)
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == "E220"

    def test_maps_with_different_sections(self):
        files = make_files()
        files[NRAD] = map_bytes(NRAD_CODES, sections=["DL Gefäße", "MRT", "Andere"])
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == "E221"

    def test_a_selection_without_rows(self):
        with pytest.raises(AppError) as info:
            evaluate(make_files(exams("NMRT", "NCT"), nrad=False))
        assert info.value.code == "E303" and "Allrad-Leistungen" in info.value.title
        with pytest.raises(AppError) as info:
            evaluate(make_files(exams("RMRT"), allrad=False))
        assert info.value.code == "E303" and "NRAD-Leistungen" in info.value.title

    def test_no_statistics(self):
        files = make_files()
        del files[STATISTICS]
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == "E302" and info.value.slot == STATISTICS

    def test_no_map(self):
        with pytest.raises(AppError) as info:
            evaluate(make_files(allrad=False, nrad=False))
        assert info.value.code == "E301"

    def test_a_slot_with_none_counts_as_missing(self):
        files = make_files()
        files[ALLRAD] = None
        assert evaluate(files).summary["selection"] == [NRAD]

    def test_every_failure_has_a_message_a_hint_and_a_slot_label(self):
        files = make_files()
        files[NRAD] = b"x"
        with pytest.raises(AppError) as info:
            evaluate(files)
        shown = info.value.to_dict()
        assert shown["title"] and shown["hint"] and shown["slot_label"] == SLOT_LABELS[NRAD]


# ── Harmful files ──────────────────────────────────────────────────────────────

class TestHarmfulFiles:

    @pytest.mark.parametrize("slot", [STATISTICS, ALLRAD, NRAD])
    @pytest.mark.parametrize("name", HOSTILE)
    def test_every_attack_is_stopped_with_its_code_in_every_slot(self, name, slot):
        make, suffix = HOSTILE[name]
        files = make_files()
        files[slot] = make(lambda: files[slot])
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == series(slot) + suffix and info.value.slot == slot

    @pytest.mark.parametrize("slot", [STATISTICS, ALLRAD, NRAD])
    @pytest.mark.parametrize("name", HOSTILE)
    def test_a_hostile_file_is_refused_before_any_file_is_parsed(self, name, slot, monkeypatch):
        parsed = []                                              # recorded, not raised: evaluate turns loader errors into AppErrors
        for loader in ("load_statistics", "load_map"):
            monkeypatch.setattr(server, loader, lambda *a, _real=getattr(server, loader), _n=loader, **k: (parsed.append(_n), _real(*a, **k))[1])
        make, _ = HOSTILE[name]
        files = make_files()
        files[slot] = make(lambda: files[slot])
        with pytest.raises(AppError):
            evaluate(files)
        assert parsed == [], f"parsed {parsed} although a file was hostile"

    def test_a_zip_bomb_is_stopped_at_once(self):
        make, _ = HOSTILE["zip bomb"]
        files = make_files()
        files[STATISTICS] = make(None)
        start = time.perf_counter()
        with pytest.raises(AppError):
            evaluate(files)
        assert time.perf_counter() - start < 5

    @pytest.mark.parametrize("header", ['=HYPERLINK("http://evil.example","klick")', "=1+1", "+cmd|' /C calc'!A0", "@SUM(1)", "-2+3"])
    def test_a_map_with_a_formula_header_never_reaches_the_output(self, header):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=2).value = "Leist0"
        for i, name in enumerate(["DL Gefäße", header, "CT"], 6):
            force_text(ws.cell(row=1, column=i), name)
        ws.cell(row=2, column=2).value = "RMRT"
        ws.cell(row=2, column=7).value = 1
        files = make_files(nrad=False)
        files[ALLRAD] = wb_to_bytes(wb)
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == "E205" and info.value.slot == ALLRAD
        assert "evil.example" not in info.value.title

    def test_the_second_line_of_defence_when_the_loader_check_is_bypassed(self, monkeypatch):
        monkeypatch.setattr(input_handler, "_check_cell_for_injection", lambda *args, **kwargs: None)
        files = make_files(nrad=False)
        files[ALLRAD] = map_with_formula_header()
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == "E903"

    def test_the_third_line_of_defence_when_both_core_checks_are_bypassed(self, monkeypatch):
        monkeypatch.setattr(input_handler, "_check_cell_for_injection", lambda *args, **kwargs: None)
        monkeypatch.setattr(output_handler, "_assert_text_only", lambda wb: None)
        files = make_files(nrad=False)
        files[ALLRAD] = map_with_formula_header()
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == "E903" and "Formel" in info.value.detail

    def test_html_in_the_staff_name_stays_plain_text_in_the_summary_and_the_sheet(self):
        evil = "<img src=x onerror=alert(1)>"
        result = evaluate(make_files(mitarbeiter=evil))
        assert result.summary["mitarbeiter"] == evil           # inert: the page inserts it as text
        assert all(c.isalnum() or c in " .-_()" for c in result.filename)

    def test_long_texts_from_the_file_are_clipped_in_the_summary(self):
        assert len(evaluate(make_files(mitarbeiter="M" * 5000)).summary["mitarbeiter"]) <= 100


def map_with_formula_header() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=2).value = "Leist0"
    for i, name in enumerate(["DL Gefäße", '=HYPERLINK("http://evil.example","klick")', "CT"], 6):
        force_text(ws.cell(row=1, column=i), name)
    ws.cell(row=2, column=2).value = "RMRT"
    ws.cell(row=2, column=7).value = 1
    return wb_to_bytes(wb)


# ── Nothing is left behind ─────────────────────────────────────────────────────

def openpyxl_temp_files() -> set[str]:
    return {name for name in os.listdir(tempfile.gettempdir()) if name.startswith("openpyxl") or ".openpyxl." in name}


class TestNothingIsLeftBehind:

    def test_the_working_folder_stays_empty(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        evaluate(make_files(), reference_data=load_reference_data(REFERENCE))
        assert os.listdir(tmp_path) == []

    def test_no_temp_file_of_openpyxl_is_left(self):
        before = openpyxl_temp_files()
        evaluate(make_files())
        with pytest.raises(AppError):
            evaluate(make_files(allrad=False, nrad=False))
        assert openpyxl_temp_files() == before

    def test_the_bytes_that_were_passed_in_are_not_changed(self):
        files = make_files()
        before = dict(files)
        copy = {slot: bytearray(data) for slot, data in files.items()}
        evaluate(files)
        evaluate(copy)
        assert files == before and all(bytes(copy[slot]) == before[slot] for slot in files)

    def test_the_temp_file_openpyxl_writes_holds_no_patient_name(self, monkeypatch):
        seen = []
        real_remove = os.remove

        def spy(path, *args, **kwargs):
            if "openpyxl" in os.path.basename(str(path)):
                seen.append(Path(path).read_bytes())
            return real_remove(path, *args, **kwargs)

        # Build the statistics first: saving a workbook here would write (and remove) a temp file of its own.
        wb = openpyxl.load_workbook(io.BytesIO(stats_bytes(exams("RMRT", "NMRT"), total_l8=2)))
        wb.active.cell(row=11, column=6).value = "Mustermann, Maria"      # a column the evaluation never reads
        buffer = io.BytesIO()
        wb.save(buffer)
        files = make_files()
        files[STATISTICS] = buffer.getvalue()
        planted = openpyxl.load_workbook(io.BytesIO(files[STATISTICS])).active.cell(row=11, column=6).value
        assert planted == "Mustermann, Maria"                              # the name really is in the upload
        monkeypatch.setattr(os, "remove", spy)
        result = evaluate(files)
        with zipfile.ZipFile(io.BytesIO(result.data)) as zf:                # ... and nowhere in the Excel
            assert all(b"Mustermann" not in zf.read(entry) for entry in zf.namelist())
        assert seen, "the spy saw no temp file, so this test proves nothing"
        assert all(b"Mustermann" not in content for content in seen)
        assert any(b"MRT" in content for content in seen)                 # it holds the result, nothing more


# ── Unexpected errors ──────────────────────────────────────────────────────────

class TestUnexpectedErrors:

    def test_a_bug_becomes_e999_without_its_message(self, monkeypatch):
        def broken(*args, **kwargs):
            raise RuntimeError("Mustermann, Maria hat Befund 4711")
        monkeypatch.setattr(server, "process", broken)
        with pytest.raises(AppError) as info:
            evaluate(make_files())
        assert info.value.code == "E999"
        assert "RuntimeError" in info.value.detail and "Mustermann" not in info.value.detail and "4711" not in info.value.detail

    def test_any_exception_while_reading_a_file_gives_its_file_code_never_e999(self, monkeypatch):
        def broken(*args, **kwargs):
            raise KeyError("Mustermann")
        monkeypatch.setattr(server, "load_map", broken)
        with pytest.raises(AppError) as info:
            evaluate(make_files())
        assert info.value.code == "E200" and info.value.slot == ALLRAD

    def test_a_memory_error_while_reading_is_a_file_error_too(self, monkeypatch):
        def broken(*args, **kwargs):
            raise MemoryError()
        monkeypatch.setattr(server, "load_statistics", broken)
        with pytest.raises(AppError) as info:
            evaluate(make_files())
        assert info.value.code == "E100" and info.value.slot == STATISTICS

    def test_unknown_slots_are_a_programming_error(self):
        files = make_files()
        files["../etc/passwd"] = b"x"
        with pytest.raises(ValueError):
            evaluate(files)

    @pytest.mark.parametrize("value", ["text", 5, [b"x"]])
    def test_content_that_is_not_bytes_is_a_programming_error(self, value):
        files = make_files()
        files[STATISTICS] = value
        with pytest.raises(ValueError):
            evaluate(files)


# ── bundle_dir ─────────────────────────────────────────────────────────────────

class TestBundleDir:

    def test_it_is_the_project_root_normally(self):
        root = bundle_dir()
        assert (root / "core" / "input_handler.py").is_file() and (root / "local_app" / "server.py").is_file()

    def test_it_is_the_unpack_folder_when_frozen(self, monkeypatch, tmp_path):
        monkeypatch.setattr(sys, "frozen", True, raising=False)
        monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path), raising=False)
        assert bundle_dir() == tmp_path

    def test_frozen_without_an_unpack_folder_falls_back_to_the_root(self, monkeypatch):
        monkeypatch.setattr(sys, "frozen", True, raising=False)
        monkeypatch.delattr(sys, "_MEIPASS", raising=False)
        assert (bundle_dir() / "core").is_dir()

    def test_an_unpack_folder_without_frozen_is_ignored(self, monkeypatch, tmp_path):
        monkeypatch.delattr(sys, "frozen", raising=False)
        monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path), raising=False)
        assert bundle_dir() != tmp_path


# ── split_upload ───────────────────────────────────────────────────────────────

import random  # noqa: E402
from urllib.parse import urlencode  # noqa: E402

from local_app.server import MAX_BODY_BYTES, split_upload  # noqa: E402


def binary(n, seed):
    """n bytes of everything a file may hold, including NUL, 0xFF, 'PK' and a boundary-like text."""
    rng = random.Random(seed)
    return bytes(rng.randrange(256) for _ in range(n)) + b"PK\x03\x04\x00\xff--boundary\r\n"


S_BYTES, A_BYTES, N_BYTES = binary(300, 1), binary(1000, 2), binary(77, 3)


def upload(stats=S_BYTES, allrad=A_BYTES, nrad=N_BYTES, **extra):
    """(body, query) the way the page sends it: the files back to back, the sizes in the query."""
    body, params = b"", {}
    for key, data in (("stats", stats), ("allrad", allrad), ("nrad", nrad)):
        if data is not None:
            body += data
            params[key] = str(len(data))
    params.update(extra)
    return body, urlencode(params)


def refused(body, query):
    with pytest.raises(AppError) as info:
        split_upload(body, query)
    return info.value


class TestSplitUpload:

    def test_both_maps(self):
        files, names = split_upload(*upload())
        assert files == {STATISTICS: S_BYTES, ALLRAD: A_BYTES, NRAD: N_BYTES} and names == {}

    def test_allrad_only(self):
        files, _ = split_upload(*upload(nrad=None))
        assert files == {STATISTICS: S_BYTES, ALLRAD: A_BYTES}

    def test_nrad_only(self):
        files, _ = split_upload(*upload(allrad=None))
        assert files == {STATISTICS: S_BYTES, NRAD: N_BYTES}

    def test_binary_content_is_preserved_exactly(self):
        for seed in range(20):
            stats, allrad, nrad = binary(random.Random(seed).randrange(0, 500), seed), binary(50, seed + 100), binary(0, seed + 200)
            files, _ = split_upload(*upload(stats, allrad, nrad))
            assert files == {STATISTICS: stats, ALLRAD: allrad, NRAD: nrad}

    def test_the_slices_are_bytes_and_in_the_order_of_the_body_not_of_the_query(self):
        body, _ = upload()
        files, _ = split_upload(body, f"nrad={len(N_BYTES)}&allrad={len(A_BYTES)}&stats={len(S_BYTES)}")
        assert files[STATISTICS] == S_BYTES and files[ALLRAD] == A_BYTES and files[NRAD] == N_BYTES
        assert all(type(data) is bytes for data in files.values())

    def test_an_empty_file_is_passed_on_so_that_the_scan_can_name_it(self):
        files, _ = split_upload(*upload(allrad=b""))
        assert files[ALLRAD] == b""
        with pytest.raises(AppError) as info:
            evaluate({**make_files(), ALLRAD: files[ALLRAD]})
        assert info.value.code == "E202" and info.value.slot == ALLRAD

    def test_the_query_may_start_with_a_question_mark_and_be_a_mapping(self):
        body, query = upload()
        assert split_upload(body, "?" + query)[0] == split_upload(body, query)[0]
        assert split_upload(body, {"stats": str(len(S_BYTES)), "allrad": str(len(A_BYTES)), "nrad": str(len(N_BYTES))})[0][STATISTICS] == S_BYTES


class TestMissingParts:

    def test_no_statistics(self):
        error = refused(*upload(stats=None))
        assert error.code == "E302" and error.slot == STATISTICS

    def test_no_map(self):
        error = refused(*upload(allrad=None, nrad=None))
        assert error.code == "E302" and error.slot is None

    def test_nothing_at_all(self):
        assert refused(b"", "").code == "E302"

    def test_a_size_without_a_value_is_invalid_not_missing(self):
        assert refused(*upload(stats=None), ) .code == "E302"
        assert refused(S_BYTES, "stats=").code == "E304"


class TestSizesThatAreNotRight:

    @pytest.mark.parametrize("size", ["-1", "1e3", "", " 5", "5 ", "+5", "0x10", "5.0", "1_000", "٣٠٠", "２", "abc", "5;6", "NaN", "inf", "9" * 30, "0" * 13 + "1"])
    def test_a_size_that_is_not_a_plain_number(self, size):
        body, _ = upload(nrad=None)
        for param in ("stats", "allrad"):
            params = {"stats": str(len(S_BYTES)), "allrad": str(len(A_BYTES))}
            params[param] = size
            error = refused(body, urlencode(params))
            assert error.code == "E304" and "Ungültige Größenangabe" in error.detail

    def test_sizes_that_add_up_to_less_than_the_body(self):
        body, _ = upload()
        assert refused(body, f"stats={len(S_BYTES)}&allrad={len(A_BYTES)}&nrad={len(N_BYTES) - 1}").code == "E304"

    def test_sizes_that_add_up_to_more_than_the_body(self):
        body, _ = upload()
        assert refused(body, f"stats={len(S_BYTES)}&allrad={len(A_BYTES)}&nrad={len(N_BYTES) + 1}").code == "E304"

    def test_a_body_cut_off_in_transit(self):
        body, query = upload()
        error = refused(body[:-5], query)
        assert error.code == "E304" and f"{len(body):,}" in error.detail and f"{len(body) - 5:,}" in error.detail

    def test_a_body_with_extra_bytes(self):
        body, query = upload()
        assert refused(body + b"x", query).code == "E304"

    def test_the_same_size_sent_twice(self):
        body, _ = upload()
        error = refused(body, f"stats={len(S_BYTES)}&stats={len(S_BYTES)}&allrad={len(A_BYTES)}&nrad={len(N_BYTES)}")
        assert error.code == "E304" and "mehrfach" in error.detail

    def test_the_same_name_sent_twice(self):
        body, query = upload()
        assert refused(body, query + "&allrad_name=a&allrad_name=b").code == "E304"

    def test_too_many_parameters(self):
        body, query = upload()
        assert refused(body, query + "&" + "&".join(f"x{i}=1" for i in range(100))).code == "E304"

    def test_no_failure_returns_anything(self):
        body, query = upload()
        for broken in (body[:-1], body + b"!"):
            with pytest.raises(AppError):
                split_upload(broken, query)


class TestSizeLimit:

    def test_a_body_over_the_limit_is_refused_first(self, monkeypatch):
        monkeypatch.setattr(server, "MAX_BODY_BYTES", 100)
        error = refused(b"x" * 101, "this is not even a query")
        assert error.code == "E305"

    def test_a_body_at_the_limit_is_accepted(self, monkeypatch):
        body, query = upload()
        monkeypatch.setattr(server, "MAX_BODY_BYTES", len(body))
        assert split_upload(body, query)[0][STATISTICS] == S_BYTES

    def test_the_limit_is_three_files_at_the_file_limit_plus_a_margin(self):
        assert 3 * input_handler.MAX_FILE_SIZE_MB * 1024 * 1024 < MAX_BODY_BYTES < 4 * input_handler.MAX_FILE_SIZE_MB * 1024 * 1024


class TestParametersThatAreNotOurs:

    @pytest.mark.parametrize("extra", [
        {"evil": "../../etc/passwd"}, {"statistics": "5"}, {"STATS": "5"}, {"slot": "../x", "path": "/etc/passwd"},
        {"__class__": "x"}, {"stats_name": "x"}, {"": "1"}, {"a b": "c"},
    ])
    def test_unknown_parameters_are_ignored(self, extra):
        body, query = upload()
        assert split_upload(body, query + "&" + urlencode(extra))[0] == {STATISTICS: S_BYTES, ALLRAD: A_BYTES, NRAD: N_BYTES}

    def test_the_result_only_ever_has_our_slots(self):
        body, query = upload()
        files, names = split_upload(body, query + "&evil_name=x&slot=../x&allrad_name=ok.xlsx")
        assert set(files) <= {STATISTICS, ALLRAD, NRAD} and set(names) <= {ALLRAD, NRAD}

    def test_a_name_for_a_map_that_was_not_sent_is_dropped(self):
        body, query = upload(nrad=None)
        assert split_upload(body, query + "&nrad_name=n.xlsx&allrad_name=a.xlsx")[1] == {ALLRAD: "a.xlsx"}


class TestNames:

    def test_names_are_cleaned(self):
        body, query = upload()
        _, names = split_upload(body, query + "&" + urlencode({"allrad_name": '../ev<il>"x.xlsx', "nrad_name": "map\x00name\n.xlsx"}))
        assert names == {ALLRAD: "evilx.xlsx", NRAD: "mapname.xlsx"}

    def test_umlauts_and_spaces_survive(self):
        body, query = upload()
        _, names = split_upload(body, query + "&" + urlencode({"allrad_name": "Map Würzburg (neu).xlsx"}))
        assert names[ALLRAD] == "Map Würzburg (neu).xlsx"

    def test_a_long_name_is_cut(self):
        body, query = upload()
        _, names = split_upload(body, query + "&" + urlencode({"allrad_name": "m" * 100_000 + ".xlsx"}))
        assert len(names[ALLRAD]) <= 80

    def test_only_the_start_of_a_huge_name_is_looked_at(self, monkeypatch):
        seen = []
        real = server.safe_name
        monkeypatch.setattr(server, "safe_name", lambda text, *a, **k: (seen.append(len(text)), real(text, *a, **k))[1])
        body, query = upload()
        split_upload(body, query + "&" + urlencode({"allrad_name": "m" * 200_000}))
        assert seen and max(seen) <= 300

    def test_a_name_that_is_nothing_but_unsafe_characters(self):
        body, query = upload()
        assert split_upload(body, query + "&allrad_name=%2F%2F%2F")[1] == {ALLRAD: "Datei"}

    def test_an_empty_name_is_no_name(self):
        body, query = upload()
        assert split_upload(body, query + "&allrad_name=")[1] == {}


class TestRoundTrip:

    def test_split_then_evaluate_gives_the_same_result_as_evaluating_directly(self):
        direct = make_files()
        body = direct[STATISTICS] + direct[ALLRAD] + direct[NRAD]
        query = urlencode({"stats": len(direct[STATISTICS]), "allrad": len(direct[ALLRAD]), "nrad": len(direct[NRAD]),
                           "allrad_name": "a.xlsx", "nrad_name": "n.xlsx"})
        files, names = split_upload(body, query)
        assert files == direct
        via_upload = evaluate(files, names, load_reference_data(REFERENCE))
        assert totals(via_upload) == totals(evaluate(direct, names, load_reference_data(REFERENCE)))
        assert any("„a.xlsx“" in note for note in notes(via_upload))

    def test_a_cut_off_upload_never_reaches_evaluate(self):
        direct = make_files()
        body = direct[STATISTICS] + direct[ALLRAD]
        with pytest.raises(AppError) as info:
            split_upload(body[:-10], urlencode({"stats": len(direct[STATISTICS]), "allrad": len(direct[ALLRAD])}))
        assert info.value.code == "E304"


# ── The HTTP server (a real one, on a free port, in a thread) ──────────────────

import http.client  # noqa: E402
import json  # noqa: E402
import socket  # noqa: E402
import threading  # noqa: E402

from local_app.server import _Handler, create_server  # noqa: E402


def port_refuses_connections(port, seconds=15):
    """
    True once nobody listens on `port` any more. Only a refusal counts: a reset or a timeout is what a full,
    unattended backlog looks like. Each attempt may take a few seconds because Windows needs about two seconds
    to refuse a connection (Linux and macOS do it at once).
    """
    deadline = time.time() + seconds
    while time.time() < deadline:
        try:
            socket.create_connection(("127.0.0.1", port), timeout=3).close()
        except ConnectionRefusedError:
            return True
        except OSError:
            pass
        time.sleep(0.05)
    return False


class Response:
    def __init__(self, status, headers, body):
        self.status, self.headers, self.body = status, headers, body

    @property
    def json(self):
        return json.loads(self.body)

    @property
    def error(self):
        return self.json["error"]


def call(app, method, path, body=None, headers=None, host="own", length="auto", timeout=10):
    """One request. host='own' sends our address, a string sends that, False sends none."""
    conn = http.client.HTTPConnection("127.0.0.1", app.port, timeout=timeout)
    try:
        conn.putrequest(method, path, skip_host=True, skip_accept_encoding=True)
        if host is not False:
            conn.putheader("Host", f"127.0.0.1:{app.port}" if host == "own" else host)
        for name, value in (headers or {}).items():
            conn.putheader(name, value)
        if length == "auto":
            if body is not None:
                conn.putheader("Content-Length", str(len(body)))
        elif length is not None:
            conn.putheader("Content-Length", str(length))
        conn.endheaders()
        if body:
            conn.send(body)
        response = conn.getresponse()
        return Response(response.status, {k.lower(): v for k, v in response.getheaders()}, response.read())
    finally:
        conn.close()


def at(app, route=""):
    return f"/{app.token}/{route}"


def post_files(app, files, names=None, **kwargs):
    order = [(STATISTICS, "stats"), (ALLRAD, "allrad"), (NRAD, "nrad")]
    body = b"".join(files[slot] for slot, _ in order if slot in files)
    params = {param: len(files[slot]) for slot, param in order if slot in files}
    for slot, name in (names or {}).items():
        params[f"{slot}_name"] = name
    return call(app, "POST", at(app, "evaluate") + "?" + urlencode(params), body, **kwargs)


@pytest.fixture
def app():
    server_ = create_server(load_reference_data(REFERENCE))
    thread = threading.Thread(target=server_.serve_forever, kwargs={"poll_interval": 0.02}, daemon=True)
    thread.start()
    server_.thread = thread
    yield server_
    server_.shutdown()
    server_.server_close()
    thread.join(5)


SECURITY_HEADERS = {
    "x-content-type-options": "nosniff", "cache-control": "no-store", "referrer-policy": "no-referrer",
    "x-frame-options": "DENY", "cross-origin-resource-policy": "same-origin", "cross-origin-opener-policy": "same-origin",
}


def assert_secure_headers(response):
    for name, value in SECURITY_HEADERS.items():
        assert response.headers.get(name) == value, (name, response.headers)
    policy = response.headers["content-security-policy"]
    assert "default-src 'none'" in policy and "frame-ancestors 'none'" in policy
    assert "python" not in response.headers.get("server", "").lower()


class TestBinding:

    def test_it_listens_on_127_0_0_1_only(self, app):
        assert app.server_address[0] == "127.0.0.1" and app.port > 0
        assert app.url == f"http://127.0.0.1:{app.port}/{app.token}/"

    def test_the_token_is_long_random_and_new_for_every_server(self):
        servers = [create_server() for _ in range(5)]
        try:
            tokens = {s.token for s in servers}
            assert len(tokens) == 5 and all(len(t) >= 30 for t in tokens)
            assert len({s.port for s in servers}) == 5
        finally:
            for s in servers:
                s.server_close()

    def test_the_port_cannot_be_shared(self, app):
        assert app.allow_reuse_address is False
        with pytest.raises(OSError):
            create_server(port=app.port)

    def test_a_stuck_request_cannot_keep_the_program_alive_and_a_stalled_client_is_dropped(self, app):
        assert app.daemon_threads is True and 0 < _Handler.timeout <= 120

    def test_a_page_without_the_nonce_placeholder_is_refused(self):
        with pytest.raises(ValueError):
            create_server(page_html="<html><script>alert(1)</script></html>")


class TestWhatIsRefused:

    def test_everything_works_with_our_own_address(self, app):
        assert call(app, "GET", at(app)).status == 200

    @pytest.mark.parametrize("path", ["/", "/evaluate", "/download", "/reset", "/shutdown", "/favicon.ico", "/index.html", "//", "/static/index.html"])
    def test_without_the_token(self, app, path):
        response = call(app, "GET", path)
        assert response.status == 404 and response.body == b""
        assert_secure_headers(response)

    @pytest.mark.parametrize("mutate", [
        lambda t: t[:-1] + ("a" if t[-1] != "a" else "b"), lambda t: t[:-1], lambda t: t + "x", lambda t: "", lambda t: t.upper(),
        lambda t: t[1:], lambda t: "../" + t, lambda t: t + "/" + t, lambda t: "%" + t[1:],
    ])
    def test_with_a_wrong_token(self, app, mutate):
        for method, route in (("GET", ""), ("GET", "download"), ("POST", "evaluate"), ("POST", "reset"), ("POST", "shutdown")):
            response = call(app, method, f"/{mutate(app.token)}/{route}", b"x" if method == "POST" else None)
            assert response.status == 404 and response.body == b""
        assert call(app, "GET", at(app)).status == 200          # shutdown with a wrong token did nothing

    def test_the_token_without_a_trailing_slash(self, app):
        assert call(app, "GET", f"/{app.token}").status == 404

    @pytest.mark.parametrize("route", ["x", "evaluate/", "Evaluate", "/", "../", "%65valuate", "download/x", "evaluate%00", "index.html", "static/index.html", "download?x=1/.."])
    def test_unknown_routes_under_the_right_token(self, app, route):
        response = call(app, "GET", at(app, route))
        assert response.status == 404

    def test_a_query_and_a_fragment_do_not_change_the_route(self, app):
        assert call(app, "GET", at(app) + "?a=b").status == 200
        assert call(app, "GET", at(app) + "#frag").status == 200

    @pytest.mark.parametrize("method, route", [("GET", "evaluate"), ("GET", "reset"), ("GET", "shutdown"), ("POST", "download"), ("POST", ""),
                                               ("PUT", "evaluate"), ("DELETE", "download"), ("PATCH", ""), ("HEAD", ""), ("HEAD", "download"),
                                               ("OPTIONS", ""), ("OPTIONS", "evaluate")])
    def test_the_wrong_method_for_a_route(self, app, method, route):
        response = call(app, method, at(app, route), b"" if method in ("PUT", "PATCH") else None)
        assert response.status == 405 and response.body == b""
        assert call(app, "GET", at(app)).status == 200          # and nothing happened

    def test_a_method_nobody_knows(self, app):
        response = call(app, "FOO", at(app))
        assert response.status == 501 and response.json["error"]["code"] == "E304"

    @pytest.mark.parametrize("host", [
        "evil.example", "127.0.0.1", "127.0.0.1:1", "localhost", "localhost.evil.example:{port}", "127.0.0.1:{port}.evil.example",
        "evil.example:{port}", "[::1]:{port}", "0.0.0.0:{port}", "", "127.0.0.1:{port}, evil.example",
    ])
    def test_a_wrong_host_even_with_the_right_token(self, app, host):
        response = call(app, "GET", at(app), host=host.format(port=app.port))
        assert response.status == 403 and response.body == b""
        assert_secure_headers(response)

    def test_no_host_header_at_all(self, app):
        assert call(app, "GET", at(app), host=False).status == 403

    def test_a_wrong_host_for_every_route(self, app):
        for method, route in (("GET", ""), ("GET", "download"), ("POST", "evaluate"), ("POST", "reset"), ("POST", "shutdown")):
            assert call(app, method, at(app, route), b"x" if method == "POST" else None, host="evil.example").status == 403
        assert call(app, "GET", at(app)).status == 200

    @pytest.mark.parametrize("host", ["localhost:{port}", "LOCALHOST:{port}", "127.0.0.1:{port}"])
    def test_the_accepted_spellings_of_our_own_address(self, app, host):
        assert call(app, "GET", at(app), host=host.format(port=app.port)).status == 200

    @pytest.mark.parametrize("origin", [
        "http://evil.example", "null", "https://127.0.0.1:{port}", "http://127.0.0.1", "http://127.0.0.1:{other}", "http://localhost.evil.example:{port}",
        "http://127.0.0.1:{port}.evil.example", "file://", "",
    ])
    def test_a_foreign_origin_on_every_route(self, app, origin):
        origin = origin.format(port=app.port, other=app.port + 1)
        for method, route in (("GET", ""), ("GET", "download"), ("POST", "evaluate"), ("POST", "reset"), ("POST", "shutdown")):
            response = call(app, method, at(app, route), b"x" if method == "POST" else None, headers={"Origin": origin})
            assert response.status == 403 and response.body == b"", (method, route)
        assert call(app, "GET", at(app)).status == 200          # the foreign shutdown did nothing

    @pytest.mark.parametrize("origin", ["http://127.0.0.1:{port}", "http://localhost:{port}", "HTTP://LOCALHOST:{port}"])
    def test_our_own_origin_is_fine(self, app, origin):
        assert call(app, "POST", at(app, "reset"), headers={"Origin": origin.format(port=app.port)}).status == 200

    @pytest.mark.parametrize("site, allowed", [("cross-site", False), ("same-site", False), ("same-origin", True), ("none", True)])
    def test_the_browser_says_where_the_request_comes_from(self, app, site, allowed):
        response = call(app, "POST", at(app, "reset"), headers={"Sec-Fetch-Site": site})
        assert (response.status == 200) is allowed and (allowed or response.status == 403)

    def test_a_refused_request_changes_nothing(self, app):
        assert post_files(app, make_files()).status == 200
        for headers, host in (({"Origin": "http://evil.example"}, "own"), ({}, "evil.example")):
            assert call(app, "POST", at(app, "reset"), headers=headers, host=host).status == 403
        assert call(app, "GET", at(app, "download")).status == 200

    def test_the_token_is_compared_in_constant_time(self, app, monkeypatch):
        calls = []
        real = server.hmac.compare_digest
        monkeypatch.setattr(server.hmac, "compare_digest", lambda a, b: (calls.append((a, b)), real(a, b))[1])
        call(app, "GET", at(app))
        assert calls and calls[0][1] == app.token.encode()


class TestSecurityHeaders:

    def test_every_kind_of_response_has_them(self, app):
        assert post_files(app, make_files()).status == 200
        responses = [
            call(app, "GET", at(app)), call(app, "GET", at(app, "download")), call(app, "POST", at(app, "reset")),
            post_files(app, make_files()), post_files(app, {STATISTICS: b"x", ALLRAD: b"y"}), call(app, "GET", "/nope"),
            call(app, "GET", at(app), host="evil.example"), call(app, "PUT", at(app)), call(app, "FOO", at(app)),
            call(app, "POST", at(app, "evaluate"), length=None),
        ]
        assert {r.status for r in responses} >= {200, 404, 403, 405, 422, 400, 501}
        for response in responses:
            assert_secure_headers(response)

    def test_the_server_does_not_name_itself(self, app):
        assert "python" not in call(app, "GET", at(app)).headers["server"].lower()


class TestPage:

    def test_the_page_is_served_with_a_policy_that_names_its_nonce(self, app):
        response = call(app, "GET", at(app))
        assert response.status == 200 and response.headers["content-type"] == "text/html; charset=utf-8"
        policy = response.headers["content-security-policy"]
        nonce = re.search(r"script-src 'nonce-([^']+)'", policy).group(1)
        assert f"style-src 'nonce-{nonce}'" in policy
        assert f'<script nonce="{nonce}">' in response.body.decode() and f'<style nonce="{nonce}">' in response.body.decode()
        assert "{{NONCE}}" not in response.body.decode()

    def test_every_response_has_a_fresh_nonce(self, app):
        nonces = {re.search(r"nonce-([^']+)", call(app, "GET", at(app)).headers["content-security-policy"]).group(1) for _ in range(5)}
        assert len(nonces) == 5 and all(len(n) >= 20 for n in nonces)

    def test_the_policy_allows_nothing_external_and_nothing_inline(self, app):
        policy = call(app, "GET", at(app)).headers["content-security-policy"]
        directives = dict(part.strip().split(" ", 1) for part in policy.split(";") if part.strip())
        assert directives["default-src"] == "'none'" and directives["connect-src"] == "'self'"
        assert directives["frame-ancestors"] == "'none'" and directives["base-uri"] == "'none'" and directives["form-action"] == "'none'"
        for forbidden in ("unsafe-inline", "unsafe-eval", "*", "http:", "https:", "data:", "blob:"):
            assert forbidden not in policy
        assert set(directives) == {"default-src", "script-src", "style-src", "connect-src", "base-uri", "form-action", "frame-ancestors"}

    def test_other_responses_have_the_strict_policy_without_a_nonce(self, app):
        policy = call(app, "GET", at(app, "download")).headers["content-security-policy"]
        assert policy == "default-src 'none'; frame-ancestors 'none'"

    def test_a_page_given_by_the_caller_is_used(self):
        s = create_server(page_html="<p nonce='{{NONCE}}'>eigene Seite</p>")
        thread = threading.Thread(target=s.serve_forever, kwargs={"poll_interval": 0.02}, daemon=True)
        thread.start()
        try:
            assert b"eigene Seite" in call(s, "GET", at(s)).body
        finally:
            s.shutdown(); s.server_close(); thread.join(5)

    def test_the_bundled_page_exists_and_carries_the_placeholder(self):
        text = (bundle_dir() / "local_app" / "static" / "index.html").read_text(encoding="utf-8")
        assert text.count("{{NONCE}}") >= 2 and "<html" in text


import re  # noqa: E402


class TestEvaluateAndDownload:

    def test_the_summary_comes_back_as_json(self, app):
        response = post_files(app, make_files(), {ALLRAD: "mapped_leistungen_2611.xlsx", NRAD: "mapped_leistungen_nrad.xlsx"})
        assert response.status == 200 and response.headers["content-type"] == "application/json"
        expected = evaluate(make_files(), {ALLRAD: "mapped_leistungen_2611.xlsx", NRAD: "mapped_leistungen_nrad.xlsx"},
                            load_reference_data(REFERENCE)).summary
        assert response.json == {"ok": True, "summary": expected}
        assert response.json["summary"]["selection_label"] == "Allrad + NRAD"

    def test_the_download_is_a_valid_excel_file_with_the_right_headers(self, app):
        evaluation = post_files(app, make_files(nrad=False))
        download = call(app, "GET", at(app, "download"))
        assert download.status == 200 and download.headers["content-type"] == server._XLSX
        assert download.headers["content-disposition"].startswith("attachment; ")
        assert_secure_headers(download)
        assert download.body[:2] == b"PK" and int(download.headers["content-length"]) == len(download.body)
        ws = openpyxl.load_workbook(io.BytesIO(download.body)).active
        assert ws.cell(row=1, column=1).value == "Leistungsbereich"
        assert evaluation.json["summary"]["filename"] in download.headers["content-disposition"]

    def test_a_staff_name_with_umlauts_downloads_with_a_correct_file_name(self, app):
        post_files(app, make_files(mitarbeiter="Müller_Jürgen", befunddatum="01.01.2024-31.12.2024"))
        download = call(app, "GET", at(app, "download"))
        disposition = download.headers["content-disposition"]
        assert "filename*=UTF-8''M%C3%BCller_J%C3%BCrgen01012024-31122024Auswertung_Allrad_NRAD.xlsx" in disposition
        assert 'filename="M_ller_J_rgen01012024-31122024Auswertung_Allrad_NRAD.xlsx"' in disposition
        disposition.encode("ascii")                                        # a header must be plain ASCII
        assert openpyxl.load_workbook(io.BytesIO(download.body)).active.max_row > 1

    def test_html_in_the_data_comes_back_as_inert_json_text(self, app):
        evil = "<img src=x onerror=alert(1)>"
        response = post_files(app, make_files(mitarbeiter=evil))
        assert response.headers["content-type"] == "application/json" and response.headers["x-content-type-options"] == "nosniff"
        assert response.json["summary"]["mitarbeiter"] == evil
        assert response.body.decode("ascii")                               # plain ASCII JSON, nothing the browser could read as markup

    def test_characters_beyond_ascii_are_escaped_in_the_json(self, app):
        response = post_files(app, make_files(mitarbeiter="Müller_Jürgen ✓"))
        body = response.body.decode("ascii")                                 # raises if anything non-ASCII is in the body
        assert "\\u00fc" in body and response.json["summary"]["mitarbeiter"] == "Müller_Jürgen ✓"

    def test_the_map_names_reach_the_sheet(self, app):
        post_files(app, make_files(nrad=False), {ALLRAD: "Map Würzburg (neu).xlsx"})
        ws = openpyxl.load_workbook(io.BytesIO(call(app, "GET", at(app, "download")).body)).active
        assert any("„Map Würzburg (neu).xlsx“" in str(c.value) for row in ws.iter_rows() for c in row if c.value)

    def test_the_reference_data_of_the_server_is_used(self, app):
        post_files(app, make_files())
        assert "Benötigt" in [c.value for c in openpyxl.load_workbook(io.BytesIO(call(app, "GET", at(app, "download")).body)).active[1]]

    def test_without_reference_data_the_hint_says_so(self):
        s = create_server()
        thread = threading.Thread(target=s.serve_forever, kwargs={"poll_interval": 0.02}, daemon=True)
        thread.start()
        try:
            hints = [h["code"] for h in post_files(s, make_files()).json["summary"]["hints"]]
            assert "W301" in hints
        finally:
            s.shutdown(); s.server_close(); thread.join(5)

    def test_download_before_any_evaluation(self, app):
        response = call(app, "GET", at(app, "download"))
        assert response.status == 404 and response.error["code"] == "E901" and response.error["hint"]

    def test_download_after_reset(self, app):
        post_files(app, make_files())
        assert call(app, "GET", at(app, "download")).status == 200
        assert call(app, "POST", at(app, "reset")).json == {"ok": True}
        assert call(app, "GET", at(app, "download")).error["code"] == "E901"

    def test_a_second_evaluation_replaces_the_first(self, app):
        post_files(app, make_files(nrad=False))
        post_files(app, make_files(allrad=False))
        disposition = call(app, "GET", at(app, "download")).headers["content-disposition"]
        assert "_NRAD.xlsx" in disposition and "_Allrad" not in disposition

    def test_a_failed_evaluation_leaves_no_old_result(self, app):
        assert post_files(app, make_files()).status == 200
        failed = post_files(app, {STATISTICS: b"not a workbook", ALLRAD: map_bytes(ALLRAD_CODES, sections=SECTIONS)})
        assert failed.status == 422
        assert call(app, "GET", at(app, "download")).error["code"] == "E901"

    def test_a_request_that_is_refused_before_it_is_read_also_leaves_no_old_result(self, app):
        post_files(app, make_files())
        call(app, "POST", at(app, "evaluate"), length=None)
        assert call(app, "GET", at(app, "download")).status == 404

    @pytest.mark.parametrize("slot", [STATISTICS, ALLRAD, NRAD])
    def test_a_validation_error_comes_back_with_everything_the_page_shows(self, app, slot):
        files = make_files()
        files[slot] = b"this is not a workbook"
        response = post_files(app, files)
        assert response.status == 422 and response.json["ok"] is False
        error = response.error
        assert set(error) == {"code", "title", "hint", "detail", "slot", "slot_label"}
        assert error["code"] == series(slot) + "01" and error["slot"] == slot and error["slot_label"] == SLOT_LABELS[slot]
        assert error["title"] and error["hint"]

    @pytest.mark.parametrize("name", ["internal entity", "billion laughs", "zip bomb", "macros", "UTF-16 sheet"])
    def test_the_attacks_are_refused_over_http_too(self, app, name):
        make, suffix = HOSTILE[name]
        files = make_files()
        files[NRAD] = make(lambda: files[NRAD])
        response = post_files(app, files)
        assert response.status == 422 and response.error["code"] == "E2" + suffix and response.error["slot"] == NRAD

    def test_selection_errors(self, app):
        assert post_files(app, make_files(exams("RMRT"), allrad=False)).error["code"] == "E303"
        assert post_files(app, {STATISTICS: make_files()[STATISTICS]}).error["code"] == "E302"
        assert post_files(app, {ALLRAD: map_bytes(ALLRAD_CODES, sections=SECTIONS)}).error["slot"] == STATISTICS

    def test_an_unexpected_error_gives_e999_with_only_type_and_place(self, app, monkeypatch):
        def broken(*args, **kwargs):
            raise RuntimeError("Mustermann, Maria, Befund 4711")
        monkeypatch.setattr(server, "evaluate", broken)
        response = post_files(app, make_files())
        assert response.status == 500 and response.error["code"] == "E999"
        assert "RuntimeError" in response.error["detail"] and "Mustermann" not in response.body.decode() and "4711" not in response.body.decode()

    def test_a_bug_outside_evaluate_is_also_e999(self, app, monkeypatch):
        monkeypatch.setattr(server, "split_upload", lambda *a, **k: 1 / 0)
        response = post_files(app, make_files())
        assert response.status == 500 and response.error["code"] == "E999" and "ZeroDivisionError" in response.error["detail"]

    def test_a_bug_in_a_route_is_e999_and_the_server_goes_on(self, app, monkeypatch):
        monkeypatch.setattr(app.state, "get_result", lambda: (_ for _ in ()).throw(KeyError("Mustermann")))
        response = call(app, "GET", at(app, "download"))
        assert response.status == 500 and response.error["code"] == "E999" and "Mustermann" not in response.body.decode()
        monkeypatch.undo()
        assert call(app, "GET", at(app)).status == 200

    def test_the_server_keeps_no_upload(self, app):
        post_files(app, make_files())
        held = app.state.get_result()
        assert isinstance(held, Evaluation) and set(vars(app.state)) == {"reference_data", "page_html", "busy", "_lock", "_result", "_closed", "idle_seconds", "_auto_stop_seconds", "_clock", "_last_activity", "expired"}
        assert not any(isinstance(v, (bytes, bytearray)) and len(v) > 0 and v is not held.data for v in vars(held).values())


class TestContentLength:

    def test_a_missing_length(self, app):
        response = call(app, "POST", at(app, "evaluate") + "?stats=1&allrad=1", length=None)
        assert response.status == 400 and response.error["code"] == "E304"

    @pytest.mark.parametrize("value", ["abc", "-5", "1e3", "5 5", "", "+5", "0x1F", "9" * 13, "5.0", "1,000", "NaN"])
    def test_an_invalid_length_is_refused_without_reading_the_body(self, app, value):
        # No body is sent. A server that tried to read one would wait here until the client's timeout.
        response = call(app, "POST", at(app, "evaluate") + "?stats=1&allrad=1", length=value, timeout=3)
        assert response.status == 400 and response.error["code"] == "E304"

    def test_digits_of_other_scripts_are_not_a_length(self, app):
        with socket.create_connection(("127.0.0.1", app.port), timeout=3) as sock:     # http.client cannot even send this
            sock.sendall(f"POST /{app.token}/evaluate HTTP/1.0\r\nHost: 127.0.0.1:{app.port}\r\nContent-Length: ".encode() + "١٢".encode() + b"\r\n\r\n")
            data = b""
            while chunk := sock.recv(65536):                      # headers and body may arrive separately: read until the server closes
                data += chunk
        assert data.startswith(b"HTTP/1.0 400") and b'"E304"' in data

    def test_a_client_that_sends_expect_100_continue_and_the_body_at_once_is_answered(self, app):
        # The server answers as HTTP/1.0, which has no "100 Continue": such a client just sends its body.
        files = make_files(nrad=False)
        body = files[STATISTICS] + files[ALLRAD]
        with socket.create_connection(("127.0.0.1", app.port), timeout=5) as sock:
            sock.sendall(f"POST /{app.token}/evaluate?stats={len(files[STATISTICS])}&allrad={len(files[ALLRAD])} HTTP/1.1\r\n"
                         f"Host: 127.0.0.1:{app.port}\r\nExpect: 100-continue\r\nContent-Length: {len(body)}\r\n\r\n".encode() + body)
            data = b""
            while chunk := sock.recv(65536):
                data += chunk
        assert data.startswith(b"HTTP/1.0 200") and b"100 Continue" not in data

    def test_two_different_lengths(self, app):
        conn = http.client.HTTPConnection("127.0.0.1", app.port, timeout=3)
        conn.putrequest("POST", at(app, "evaluate"), skip_host=True)
        for name, value in (("Host", f"127.0.0.1:{app.port}"), ("Content-Length", "5"), ("Content-Length", "6")):
            conn.putheader(name, value)
        conn.endheaders()
        assert conn.getresponse().status == 400

    def test_chunked_bodies_are_not_accepted(self, app):
        response = call(app, "POST", at(app, "evaluate"), headers={"Transfer-Encoding": "chunked"}, length=None, timeout=3)
        assert response.status == 400 and response.error["code"] == "E304"

    def test_a_length_together_with_a_transfer_encoding_is_refused(self, app):
        files = make_files(nrad=False)
        body = files[STATISTICS] + files[ALLRAD]
        query = f"?stats={len(files[STATISTICS])}&allrad={len(files[ALLRAD])}"
        response = call(app, "POST", at(app, "evaluate") + query, body, headers={"Transfer-Encoding": "chunked"})
        assert response.status == 400 and "Teilen" in response.error["detail"]
        assert call(app, "POST", at(app, "evaluate") + query, body).status == 200          # the same request without it is fine

    def test_an_oversized_length_is_refused_without_reading_the_body(self, app):
        response = call(app, "POST", at(app, "evaluate") + "?stats=1&allrad=1", length=server.MAX_BODY_BYTES + 1, timeout=3)
        assert response.status == 413 and response.error["code"] == "E305"
        assert call(app, "POST", at(app, "evaluate"), length=10 ** 11, timeout=3).status == 413
        assert call(app, "POST", at(app, "evaluate"), length=10 ** 12, timeout=3).status == 400      # 13 digits: not a length at all

    def test_a_length_at_the_limit_is_not_refused_for_its_size(self, app, monkeypatch):
        monkeypatch.setattr(server, "MAX_BODY_BYTES", 50)
        response = call(app, "POST", at(app, "evaluate") + "?stats=50&allrad=0", b"x" * 50)
        assert response.status != 413
        assert call(app, "POST", at(app, "evaluate"), b"x" * 51).status == 413

    def test_the_length_counts_the_bytes_that_follow(self, app):
        body = b"".join(make_files(nrad=False).values())
        response = call(app, "POST", at(app, "evaluate") + f"?stats=10&allrad={len(body) - 10}", body)
        assert response.status == 422 or response.status == 400            # a statistics file of 10 bytes: refused

    def test_sizes_in_the_query_that_do_not_add_up(self, app):
        files = make_files(nrad=False)
        body = files[STATISTICS] + files[ALLRAD]
        response = call(app, "POST", at(app, "evaluate") + f"?stats={len(files[STATISTICS])}&allrad={len(files[ALLRAD]) + 1}", body)
        assert response.status == 400 and response.error["code"] == "E304"

    def test_an_empty_body(self, app):
        response = call(app, "POST", at(app, "evaluate"), b"")
        assert response.status == 400 and response.error["code"] == "E302"

    def test_a_connection_that_ends_before_the_body_is_complete(self, app):
        with socket.create_connection(("127.0.0.1", app.port), timeout=5) as sock:
            sock.sendall(f"POST /{app.token}/evaluate?stats=100&allrad=0 HTTP/1.0\r\nHost: 127.0.0.1:{app.port}\r\nContent-Length: 100\r\n\r\n".encode() + b"x" * 10)
            sock.shutdown(socket.SHUT_WR)
            data = b""
            while chunk := sock.recv(65536):
                data += chunk
        assert data.startswith(b"HTTP/1.0 400") and b'"E304"' in data
        assert call(app, "GET", at(app)).status == 200

    def test_a_body_shorter_than_announced_is_caught_by_the_length_even_if_the_query_matches_it(self, app):
        # The query declares exactly the 10 bytes that arrive, so only the Content-Length check can notice.
        with socket.create_connection(("127.0.0.1", app.port), timeout=5) as sock:
            sock.sendall(f"POST /{app.token}/evaluate?stats=10&allrad=0 HTTP/1.0\r\nHost: 127.0.0.1:{app.port}\r\nContent-Length: 100\r\n\r\n".encode() + b"x" * 10)
            sock.shutdown(socket.SHUT_WR)
            data = b""
            while chunk := sock.recv(65536):
                data += chunk
        assert data.startswith(b"HTTP/1.0 400")
        assert json.loads(data.split(b"\r\n\r\n", 1)[1])["error"]["detail"] == "Angekündigt: 100 Bytes, empfangen: 10."

    def test_a_client_that_stalls_is_dropped_with_an_answer(self, app, monkeypatch):
        monkeypatch.setattr(server._Handler, "timeout", 0.3)
        with socket.create_connection(("127.0.0.1", app.port), timeout=5) as sock:         # sends a part of the body and then nothing
            sock.sendall(f"POST /{app.token}/evaluate?stats=100&allrad=0 HTTP/1.0\r\nHost: 127.0.0.1:{app.port}\r\nContent-Length: 100\r\n\r\n".encode() + b"x" * 10)
            data = b""
            while chunk := sock.recv(65536):
                data += chunk
        assert data.startswith(b"HTTP/1.0 400") and b'"E304"' in data and b"zu langsam" in data
        assert call(app, "GET", at(app)).status == 200


class TestNothingIsWrittenToTheConsole:

    def test_not_for_good_requests_bad_requests_and_broken_connections(self, app, capfd):
        post_files(app, make_files())
        call(app, "GET", at(app, "download"))
        call(app, "GET", "/nope")
        call(app, "POST", at(app, "evaluate"), b"x")
        for raw in (b"GARBAGE\r\n\r\n", b"GET / HTTP/1.0\r\n" + b"X: " + b"a" * 70000 + b"\r\n\r\n", b"POST /x HTTP/9.9\r\n\r\n", b"\x00\xff\xfe" * 100, b""):
            with socket.create_connection(("127.0.0.1", app.port), timeout=0.5) as sock:   # some inputs never get an answer: a stalled client
                sock.sendall(raw)
                try:
                    sock.recv(65536)
                except OSError:
                    pass
        with socket.create_connection(("127.0.0.1", app.port), timeout=5) as sock:      # a client that vanishes mid-request
            sock.sendall(f"POST /{app.token}/evaluate?stats=1000&allrad=0 HTTP/1.0\r\nHost: 127.0.0.1:{app.port}\r\nContent-Length: 1000\r\n\r\n".encode() + b"x" * 10)
        call(app, "GET", at(app))                                                       # the server still answers
        captured = capfd.readouterr()
        assert captured.err == "" and captured.out == ""

    def test_the_server_survives_garbage(self, app):
        for raw in (b"GARBAGE\r\n\r\n", b"\x00" * 1000, b"GET\r\n\r\n"):
            with socket.create_connection(("127.0.0.1", app.port), timeout=0.5) as sock:   # some inputs never get an answer: a stalled client
                sock.sendall(raw)
                try:
                    sock.recv(65536)
                except OSError:
                    pass
        assert call(app, "GET", at(app)).status == 200


class BlockedEvaluation:
    """Replaces evaluate(): signals that it started, waits to be released, then does the real work."""

    def __init__(self, monkeypatch):
        self.entered, self.release = threading.Event(), threading.Event()
        real = server.evaluate

        def blocked(*args, **kwargs):
            self.entered.set()
            self.release.wait(10)
            return real(*args, **kwargs)

        monkeypatch.setattr(server, "evaluate", blocked)

    def start(self, app):
        answer = []
        thread = threading.Thread(target=lambda: answer.append(post_files(app, make_files())), daemon=True)
        thread.start()
        assert self.entered.wait(5), "the evaluation did not start"
        return thread, answer


class TestOneEvaluationAtATime:

    def test_a_second_evaluation_is_refused_at_once(self, app, monkeypatch):
        blocked = BlockedEvaluation(monkeypatch)
        thread, answer = blocked.start(app)
        try:
            start = time.perf_counter()
            second = post_files(app, make_files())
            assert time.perf_counter() - start < 2 and not answer                    # at once, while the first is still running
            assert second.status == 409 and second.error["code"] == "E307" and second.error["hint"]
        finally:
            blocked.release.set()
            thread.join(10)
        assert answer[0].status == 200 and answer[0].json["ok"] is True
        assert call(app, "GET", at(app, "download")).status == 200               # the first one finished normally
        assert post_files(app, make_files()).status == 200                       # and the lock is free again

    def test_the_other_routes_answer_while_an_evaluation_runs(self, app, monkeypatch):
        blocked = BlockedEvaluation(monkeypatch)
        thread, answer = blocked.start(app)
        try:
            assert call(app, "GET", at(app)).status == 200
            assert call(app, "GET", at(app, "download")).status == 404           # no result while one is being made
            assert call(app, "POST", at(app, "reset")).status == 200
        finally:
            blocked.release.set()
            thread.join(10)

    def test_the_lock_is_released_after_every_kind_of_failure(self, app, monkeypatch):
        assert post_files(app, {STATISTICS: b"x", ALLRAD: b"y"}).status == 422
        assert call(app, "POST", at(app, "evaluate"), length=None).status == 400
        monkeypatch.setattr(server, "evaluate", lambda *a, **k: 1 / 0)
        assert post_files(app, make_files()).status == 500
        monkeypatch.undo()
        assert post_files(app, make_files()).status == 200


class TestShutdown:

    @staticmethod
    def port_is_closed(app):
        return port_refuses_connections(app.port)

    def test_it_answers_stops_serving_closes_the_port_and_clears_the_result(self, app):
        post_files(app, make_files())
        assert app.state.get_result() is not None
        response = call(app, "POST", at(app, "shutdown"))
        assert response.status == 200 and response.json == {"ok": True}
        assert_secure_headers(response)
        app.thread.join(5)
        assert not app.thread.is_alive()
        assert self.port_is_closed(app)
        assert app.state.get_result() is None

    def test_a_shutdown_without_a_result(self, app):
        assert call(app, "POST", at(app, "shutdown")).status == 200
        app.thread.join(5)
        assert not app.thread.is_alive()

    def test_it_does_not_wait_for_an_evaluation_that_is_stuck(self, app, monkeypatch):
        blocked = BlockedEvaluation(monkeypatch)
        thread, answer = blocked.start(app)
        start = time.perf_counter()
        response = call(app, "POST", at(app, "shutdown"))
        assert response.status == 200 and time.perf_counter() - start < 2
        app.thread.join(5)
        assert not app.thread.is_alive() and self.port_is_closed(app)
        assert app.state.get_result() is None
        blocked.release.set()                                                   # the stuck evaluation now finishes ...
        thread.join(10)
        assert app.state.get_result() is None                                   # ... and its result is dropped

    def test_a_refused_shutdown_does_not_stop_the_server(self, app):
        for kwargs in ({"headers": {"Origin": "http://evil.example"}}, {"host": "evil.example"}):
            assert call(app, "POST", at(app, "shutdown"), **kwargs).status == 403
        assert call(app, "GET", at(app, "shutdown")).status == 405
        assert call(app, "POST", f"/wrong/shutdown").status == 404
        assert app.thread.is_alive() and call(app, "GET", at(app)).status == 200

    def test_after_the_shutdown_nothing_is_served(self, app):
        call(app, "POST", at(app, "shutdown"))
        app.thread.join(5)
        assert self.port_is_closed(app)
        with pytest.raises(OSError):
            call(app, "GET", at(app), timeout=2)

    def test_serve_forever_returns_so_the_program_can_end(self):
        s = create_server()
        done = threading.Event()
        thread = threading.Thread(target=lambda: (s.serve_forever(poll_interval=0.02), done.set()), daemon=True)
        thread.start()
        call(s, "POST", at(s, "shutdown"))
        assert done.wait(5)


# ── Stopping by itself after a while without use ───────────────────────────────

def wrong_token(token):
    """A token that differs from `token` in its last character, whatever that character is."""
    return token[:-1] + ("a" if token[-1] != "a" else "b")


class Clock:
    """A clock the tests move by hand, so that thirty minutes take no time."""

    def __init__(self):
        self.now = 1000.0

    def __call__(self):
        return self.now

    def advance(self, seconds):
        self.now += seconds


def start(idle_seconds=60, clock=None):
    srv = create_server(load_reference_data(REFERENCE), idle_seconds=idle_seconds, clock=clock or time.monotonic)
    thread = threading.Thread(target=srv.serve_forever, kwargs={"poll_interval": 0.02}, daemon=True)
    thread.start()
    srv.thread = thread
    return srv


def stop_server(srv):
    srv.shutdown()
    srv.server_close()
    srv.thread.join(5)


@pytest.fixture
def timed():
    """(server that stops by itself after 60 'seconds', the clock that drives it)"""
    clock = Clock()
    srv = start(60, clock)
    yield srv, clock
    stop_server(srv)


def status(srv):
    return call(srv, "GET", at(srv, "status")).json


def stopped(srv, seconds=5):
    srv.thread.join(seconds)
    return not srv.thread.is_alive()


class TestTheClockOfTheProgram:

    def test_the_time_left_runs_down_and_using_it_starts_it_again(self):
        clock = Clock()
        state = server._State(None, "x{{NONCE}}", 100, clock)
        assert state.remaining() == 100
        clock.advance(30)
        assert state.remaining() == 70
        state.mark_used()
        assert state.remaining() == 100
        clock.advance(500)
        assert state.remaining() == 0                                   # never negative

    def test_without_a_limit_it_never_runs_out(self):
        for none in (None, 0):
            state = server._State(None, "x", none, Clock())
            assert state.remaining() is None and state.idle_seconds is None

    def test_a_negative_limit_is_refused(self):
        with pytest.raises(ValueError):
            create_server(idle_seconds=-1)

    def test_switching_it_off_and_on(self):
        clock = Clock()
        state = server._State(None, "x", 100, clock)
        state.set_auto_stop(False)
        clock.advance(10_000)
        assert state.remaining() is None
        state.set_auto_stop(True)
        assert state.remaining() == 100                                  # on again with the time it was started with

    def test_started_without_a_limit_it_can_be_switched_on_with_the_default(self):
        state = server._State(None, "x", None, Clock())
        state.set_auto_stop(True)
        assert state.idle_seconds == server.IDLE_SECONDS == 1800


class TestStatusAndActivity:

    def test_status_says_how_much_time_is_left(self, timed):
        srv, clock = timed
        assert status(srv) == {"ok": True, "enabled": True, "limit": 60, "remaining": 60}

    def test_asking_does_not_use_the_time_up_or_extend_it(self, timed):
        srv, clock = timed
        clock.advance(20)
        assert status(srv)["remaining"] == 40
        clock.advance(10)
        assert status(srv)["remaining"] == 30                            # asking twice changed nothing

    def test_a_limit_that_is_not_a_whole_number_of_seconds_is_rounded_up_too(self):
        srv = start(1.5)
        try:
            reply = status(srv)
            assert reply["limit"] == 2 and isinstance(reply["limit"], int) and isinstance(reply["remaining"], int)
        finally:
            stop_server(srv)

    def test_the_time_is_rounded_up_to_whole_seconds(self, timed):
        srv, clock = timed
        clock.advance(0.2)
        assert status(srv)["remaining"] == 60 and isinstance(status(srv)["remaining"], int)

    @pytest.mark.parametrize("method, route, body", [
        ("GET", "", None), ("GET", "download", None), ("POST", "reset", b""), ("POST", "activity", b""),
        ("POST", "autostop?on=1", b""), ("POST", "evaluate", b"x"),
    ])
    def test_every_request_that_passed_the_checks_counts_as_use(self, timed, method, route, body):
        srv, clock = timed
        clock.advance(30)
        assert status(srv)["remaining"] == 30
        call(srv, method, at(srv, route), body)
        assert status(srv)["remaining"] == 60

    def test_the_activity_request_answers_with_the_new_time(self, timed):
        srv, clock = timed
        clock.advance(40)
        assert call(srv, "POST", at(srv, "activity")).json["remaining"] == 60

    def test_the_wrong_token_is_always_wrong(self):
        for token in ("abcx", "abca", "abcb"):
            assert wrong_token(token) != token

    @pytest.mark.parametrize("what", ["wrong token", "wrong host", "foreign origin", "cross-site", "wrong method", "unknown route", "no token"])
    def test_a_refused_request_never_counts_as_use(self, timed, what):
        srv, clock = timed
        clock.advance(30)
        calls = {
            "wrong token": lambda: call(srv, "POST", f"/{wrong_token(srv.token)}/activity"),
            "wrong host": lambda: call(srv, "POST", at(srv, "activity"), host="evil.example"),
            "foreign origin": lambda: call(srv, "POST", at(srv, "activity"), headers={"Origin": "http://evil.example"}),
            "cross-site": lambda: call(srv, "POST", at(srv, "activity"), headers={"Sec-Fetch-Site": "cross-site"}),
            "wrong method": lambda: call(srv, "GET", at(srv, "activity")),
            "unknown route": lambda: call(srv, "POST", at(srv, "nope")),
            "no token": lambda: call(srv, "POST", "/activity"),
        }
        assert calls[what]().status in (403, 404, 405)
        assert status(srv)["remaining"] == 30                            # so no other site can keep the program running

    def test_status_and_the_others_carry_the_security_headers(self, timed):
        srv, _ = timed
        for route, method in (("status", "GET"), ("activity", "POST"), ("autostop?on=1", "POST")):
            assert_secure_headers(call(srv, method, at(srv, route)))

    def test_no_idle_limit_means_no_remaining_time(self):
        srv = start(None)
        try:
            assert status(srv) == {"ok": True, "enabled": False, "limit": None, "remaining": None}
        finally:
            stop_server(srv)


class TestItStopsByItself:

    def test_not_before_the_time_is_up(self, timed):
        srv, clock = timed
        clock.advance(59)
        time.sleep(0.3)
        assert srv.thread.is_alive() and srv.state.expired is False

    def test_when_the_time_is_up_it_stops_closes_the_port_and_forgets_the_result(self, timed):
        srv, clock = timed
        post_files(srv, make_files())
        assert srv.state.get_result() is not None
        clock.advance(61)
        assert stopped(srv)
        assert srv.state.expired is True and srv.state.get_result() is None
        assert port_refuses_connections(srv.port)

    def test_use_keeps_it_going(self, timed):
        srv, clock = timed
        for _ in range(6):
            clock.advance(50)
            time.sleep(0.1)
            assert srv.thread.is_alive()
            assert call(srv, "POST", at(srv, "activity")).status == 200
        clock.advance(61)
        assert stopped(srv)

    def test_only_a_real_use_keeps_it_going(self, timed):
        srv, clock = timed
        clock.advance(40)
        call(srv, "POST", at(srv, "activity"), host="evil.example")      # refused
        status(srv)                                                       # asking
        clock.advance(25)
        assert stopped(srv)

    def test_never_while_an_evaluation_runs_but_afterwards(self, timed, monkeypatch):
        srv, clock = timed
        blocked = BlockedEvaluation(monkeypatch)
        thread, answer = blocked.start(srv)
        clock.advance(600)                                                # long past the limit
        time.sleep(0.3)
        assert srv.thread.is_alive() and not srv.state.expired
        blocked.release.set()
        thread.join(10)
        assert answer[0].status == 200
        assert stopped(srv)                                               # the evaluation finished: now it stops

    def test_it_stops_only_once(self, timed, monkeypatch):
        srv, clock = timed
        calls = []
        monkeypatch.setattr(srv, "shutdown_soon", lambda: calls.append(1))
        clock.advance(100)
        time.sleep(0.4)
        assert calls == [1]

    def test_with_the_real_clock(self):
        srv = start(1.0)
        try:
            assert srv.thread.is_alive()
            assert stopped(srv, 5)
        finally:
            srv.server_close()

    def test_with_the_real_clock_use_keeps_it_going(self):
        srv = start(1.0)
        try:
            end = time.time() + 2.5
            while time.time() < end:
                assert call(srv, "POST", at(srv, "activity")).status == 200
                time.sleep(0.3)
            assert srv.thread.is_alive()
            assert stopped(srv, 5)
        finally:
            srv.server_close()


class TestSwitchingItOff:

    def test_off_stays_off_however_long_nobody_uses_it(self, timed):
        srv, clock = timed
        reply = call(srv, "POST", at(srv, "autostop?on=0")).json
        assert reply == {"ok": True, "enabled": False, "limit": None, "remaining": None}
        clock.advance(100_000)
        time.sleep(0.3)
        assert srv.thread.is_alive() and status(srv)["enabled"] is False

    def test_on_again_starts_the_time_again(self, timed):
        srv, clock = timed
        call(srv, "POST", at(srv, "autostop?on=0"))
        clock.advance(500)
        assert call(srv, "POST", at(srv, "autostop?on=1")).json == {"ok": True, "enabled": True, "limit": 60, "remaining": 60}
        clock.advance(61)
        assert stopped(srv)

    def test_a_program_started_without_a_limit_gets_the_default_when_switched_on(self):
        srv = start(None)
        try:
            assert call(srv, "POST", at(srv, "autostop?on=1")).json == {"ok": True, "enabled": True, "limit": 1800, "remaining": 1800}
        finally:
            stop_server(srv)

    @pytest.mark.parametrize("query", ["", "?on=2", "?on=yes", "?on=", "?off=1", "?on=0&on=1x", "?" + "&".join(f"a{i}=1" for i in range(20))])
    def test_anything_but_0_or_1_is_refused_and_changes_nothing(self, timed, query):
        srv, _ = timed
        response = call(srv, "POST", at(srv, "autostop") + query)
        assert response.status == 400 and response.error["code"] == "E304"
        assert status(srv)["enabled"] is True

    def test_the_switch_needs_post_the_token_and_our_own_origin(self, timed):
        srv, _ = timed
        assert call(srv, "GET", at(srv, "autostop?on=0")).status == 405
        assert call(srv, "POST", f"/wrong/autostop?on=0").status == 404
        assert call(srv, "POST", at(srv, "autostop?on=0"), headers={"Origin": "http://evil.example"}).status == 403
        assert call(srv, "POST", at(srv, "autostop?on=0"), host="evil.example").status == 403
        assert status(srv)["enabled"] is True
