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
        assert set(summary) == {"filename", "mitarbeiter", "befunddatum", "total_counted", "selection", "selection_label", "hints"}


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
