"""Tests for local_app.errors: catalog, translating core/ errors, grouping hints."""

import ast
from pathlib import Path

import openpyxl
import pytest

from core import input_handler
from core.logic import process
from core.output_handler import OutputValidationError, build_xlsx_bytes
from local_app.errors import (
    ALLRAD,
    CATALOG,
    HINTS,
    MAPS,
    NRAD,
    SELECTION_NOTE_PREFIX,
    SLOT_LABELS,
    STATISTICS,
    AppError,
    classify_core_message,
    classify_hint,
    file_error,
    from_core_error,
    from_output_error,
    hint,
    summarize_hints,
    unexpected_error,
)
from tests.helpers import (
    inject_vba_into_xlsx,
    make_corrupt_xlsx_bytes,
    make_map_wb,
    make_path_traversal_xlsx_bytes,
    make_result,
    make_stats_wb,
    make_zip_bomb_xlsx_bytes,
    make_zip_without_workbook_xml,
    wb_to_bytes,
)

from .helpers import DEFAULT_CODES, DEFAULT_ROWS, SHEET, force_text, map_bytes, patch_zip, stats_bytes, with_doctype

SLOTS = [STATISTICS, ALLRAD, NRAD, MAPS]
MAP_FIELDS = [ALLRAD, NRAD, MAPS]

EXPECTED_CODES = {
    "E100", "E101", "E102", "E103", "E104", "E105", "E110", "E111", "E112", "E113",
    "E200", "E201", "E202", "E203", "E204", "E205", "E210", "E211", "E212", "E213", "E214", "E220", "E221", "E222",
    "E301", "E302", "E303", "E304", "E305", "E306", "E307",
    "E901", "E902", "E903", "E999",
}
EXPECTED_HINTS = {"W101", "W102", "W103", "W104", "W105", "W201", "W202", "W301", "W302", "W401", "W900"}


def series(slot):
    """'E1' for the statistics file, 'E2' for a map."""
    return "E1" if slot == STATISTICS else "E2"


def raised(slot, data):
    """Run the real loader of `slot` on `data` and translate what it raises."""
    loader = input_handler.load_statistics if slot == STATISTICS else input_handler.load_map
    with pytest.raises(Exception) as info:
        loader(data)
    return from_core_error(info.value, slot)


# ── Catalog ────────────────────────────────────────────────────────────────────

class TestCatalog:

    def test_codes_are_exactly_the_planned_ones(self):
        assert set(CATALOG) == EXPECTED_CODES

    def test_hint_codes_are_exactly_the_planned_ones(self):
        assert set(HINTS) == EXPECTED_HINTS

    def test_every_code_has_a_title_and_a_what_to_do_text(self):
        for code, entry in CATALOG.items():
            assert entry.code == code
            assert entry.title.strip().endswith("."), code
            assert entry.hint.strip().endswith((".", "?")), code

    @pytest.mark.parametrize("n", range(6))
    def test_statistics_and_map_share_the_file_level_titles(self, n):
        assert CATALOG[f"E10{n}"].title == CATALOG[f"E20{n}"].title

    def test_size_limit_is_taken_from_core(self):
        assert f"{input_handler.MAX_FILE_SIZE_MB} MB" in CATALOG["E103"].title
        assert "Zeitraum" in CATALOG["E103"].hint        # a statistics can be exported shorter
        assert "Zeitraum" not in CATALOG["E203"].hint    # a map cannot


class TestAppError:

    def test_to_dict_has_everything_the_page_needs(self):
        err = AppError("E210", NRAD, "B1 is None")
        assert err.to_dict() == {
            "code": "E210",
            "title": "Das ist keine Leistungsmap.",
            "hint": CATALOG["E210"].hint,
            "detail": "B1 is None",
            "slot": "nrad",
            "slot_label": "NRAD-Map",
        }

    def test_error_without_a_file_has_no_slot(self):
        d = AppError("E301").to_dict()
        assert d["slot"] is None and d["slot_label"] is None

    def test_message_names_code_and_title(self):
        assert str(AppError("E999")) == "E999: Unerwarteter Fehler."

    def test_unknown_code_fails_loudly(self):
        with pytest.raises(KeyError):
            AppError("E000")

    def test_unknown_slot_fails_loudly(self):
        with pytest.raises(ValueError):
            AppError("E101", "elsewhere")

    def test_long_detail_is_cut(self):
        err = AppError("E100", STATISTICS, "x" * 5000)
        assert len(err.detail) == 500 and err.detail.endswith("…")

    def test_e303_names_the_department(self):
        assert AppError("E303", dept="NRAD-Leistungen").title == "Diese Statistik enthält keine NRAD-Leistungen."
        assert "gewählten Bereiche" in AppError("E303").title

    def test_slot_labels_cover_all_slots(self):
        assert set(SLOT_LABELS) == set(SLOTS)


class TestUnexpectedError:

    def test_shows_type_and_place_but_not_the_message(self):
        try:
            raise ValueError("Patient Mustermann")
        except ValueError as exc:
            err = unexpected_error(exc)
        assert err.code == "E999"
        assert "ValueError" in err.detail
        assert "test_errors.py" in err.detail
        assert "Mustermann" not in err.detail

    def test_exception_that_was_never_raised(self):
        assert unexpected_error(RuntimeError("never raised")).detail == "RuntimeError"


# ── Translating what the real loaders raise ────────────────────────────────────

FILE_PROBLEMS = {
    "corrupt": (make_corrupt_xlsx_bytes, "01"),
    "no_workbook_xml": (make_zip_without_workbook_xml, "01"),
    "empty": (lambda: b"", "02"),
    "macros": (lambda: inject_vba_into_xlsx(stats_bytes()), "04"),
    "path_traversal": (make_path_traversal_xlsx_bytes, "05"),
    "zip_bomb": (make_zip_bomb_xlsx_bytes, "05"),
}


class TestFileLevelProblems:

    @pytest.mark.parametrize("slot", SLOTS)
    @pytest.mark.parametrize("name", FILE_PROBLEMS)
    def test_bad_file_gets_the_code_of_its_problem(self, slot, name):
        build, suffix = FILE_PROBLEMS[name]
        err = raised(slot, build())
        assert err.code == series(slot) + suffix
        assert err.slot == slot
        assert err.detail

    @pytest.mark.parametrize("slot", SLOTS)
    def test_oversized_file(self, slot, monkeypatch):
        monkeypatch.setattr(input_handler, "MAX_FILE_SIZE_MB", 0)
        err = raised(slot, stats_bytes())
        assert err.code == series(slot) + "03"
        assert "too large" in err.detail

    def test_formula_text_as_a_statistics_code(self):
        wb = make_stats_wb(DEFAULT_ROWS)
        force_text(wb.active["E11"], "=1+1")
        err = raised(STATISTICS, wb_to_bytes(wb))
        assert err.code == "E105"
        assert "formula injection" in err.detail

    def test_formula_text_as_a_leistung_name_is_only_a_hint(self):
        # core leaves such a name out (the code stands in for it) instead of refusing the file
        wb = make_stats_wb(DEFAULT_ROWS, total_l8=1)
        force_text(wb.active["D11"], "=1+1")
        stats = input_handler.load_statistics(wb_to_bytes(wb))
        assert [h["code"] for h in summarize_hints(stats["errors"])] == ["W105"]

    @pytest.mark.parametrize("slot", MAP_FIELDS)
    def test_two_section_columns_with_one_name(self, slot):
        wb = make_map_wb(DEFAULT_CODES, sections=["DL Gefäße", "MRT", "MRT"])
        err = raised(slot, wb_to_bytes(wb))
        assert err.code == "E213" and err.slot == slot and "columns G and H" in err.detail

    @pytest.mark.parametrize("slot", MAP_FIELDS)
    @pytest.mark.parametrize("value", [1.5, -1, "x"])
    def test_a_section_value_that_is_no_whole_number(self, slot, value):
        wb = make_map_wb(DEFAULT_CODES)
        wb.active["G2"] = value
        err = raised(slot, wb_to_bytes(wb))
        assert err.code == "E214" and err.slot == slot and "G2 = " in err.detail

    @pytest.mark.parametrize("slot", [ALLRAD, NRAD])
    def test_formula_text_as_a_map_code(self, slot):
        wb = make_map_wb(DEFAULT_CODES)
        force_text(wb.active["B2"], "=EVIL")
        assert raised(slot, wb_to_bytes(wb)).code == "E205"

    @pytest.mark.parametrize("slot, builder", [(STATISTICS, stats_bytes), (ALLRAD, map_bytes), (NRAD, map_bytes)])
    def test_hidden_xml_entity_is_a_security_rejection(self, slot, builder):
        err = raised(slot, with_doctype(builder(), '<!DOCTYPE worksheet [<!ENTITY x "y">]>'))
        assert err.code == series(slot) + "05" and "XML entity" in err.detail

    @pytest.mark.parametrize("slot, builder", [(STATISTICS, stats_bytes), (ALLRAD, map_bytes), (NRAD, map_bytes)])
    def test_utf16_xml_is_a_security_rejection(self, slot, builder):
        err = raised(slot, patch_zip(builder(), SHEET, lambda d: d.decode("utf-8").encode("utf-16")))
        assert err.code == series(slot) + "05" and "XML encoding" in err.detail

    @pytest.mark.parametrize("slot", [ALLRAD, NRAD])
    def test_formula_text_as_a_map_section_name(self, slot):
        wb = make_map_wb(DEFAULT_CODES)
        force_text(wb.active.cell(row=1, column=7), "=EVIL(1)")       # G1, the "MRT" section
        err = raised(slot, wb_to_bytes(wb))
        assert err.code == "E205" and "section name" in err.detail

    @pytest.mark.parametrize("slot", [ALLRAD, NRAD])
    def test_cached_formula_value_as_a_map_section_name(self, slot):
        # a real formula whose stored result looks like a formula: the value is what gets read
        data = patch_zip(map_bytes(), SHEET, lambda d: d.replace(
            b'<c r="G1" t="inlineStr"><is><t>MRT</t></is></c>',
            b'<c r="G1" t="str"><f>1+1</f><v>=EVIL(1)</v></c>',
        ))
        err = raised(slot, data)
        assert err.code == "E205" and "formula injection" in err.detail

    @pytest.mark.parametrize("slot, builder", [(STATISTICS, stats_bytes), (ALLRAD, map_bytes), (NRAD, map_bytes)])
    def test_damaged_sheet_means_not_a_valid_excel_file(self, slot, builder):
        damaged = patch_zip(builder(), SHEET, lambda d: d[: len(d) // 2])
        err = raised(slot, damaged)
        assert err.code == series(slot) + "01"
        assert "damaged or malformed content" in err.detail and "ParseError" in err.detail

    @pytest.mark.parametrize("slot", SLOTS)
    def test_unknown_error_is_could_not_be_read_with_the_original_text(self, slot):
        err = from_core_error(RuntimeError("boom"), slot)
        assert err.code == series(slot) + "00"
        assert err.detail == "RuntimeError: boom"

    def test_unknown_message_of_a_core_error_is_could_not_be_read(self):
        err = from_core_error(input_handler.FileValidationError("Something new"), ALLRAD)
        assert err.code == "E200" and err.detail == "Something new"

    def test_app_errors_pass_through_unchanged(self):
        original = AppError("E105", NRAD, "x")
        assert from_core_error(original, STATISTICS) is original

    def test_unknown_slot_fails_loudly(self):
        with pytest.raises(ValueError):
            from_core_error(RuntimeError("x"), "elsewhere")

    def test_a_statistics_problem_in_a_map_slot_falls_back_to_could_not_be_read(self):
        assert file_error(ALLRAD, "no_rows").code == "E200"
        assert file_error(STATISTICS, "not_map").code == "E100"


class TestStructureProblems:

    def test_map_in_the_statistics_slot(self):
        err = raised(STATISTICS, map_bytes())
        assert err.code == "E110" and err.slot == STATISTICS

    @pytest.mark.parametrize("slot", [ALLRAD, NRAD])
    def test_statistics_in_a_map_slot(self, slot):
        err = raised(slot, stats_bytes())
        assert err.code == "E210" and err.slot == slot

    def test_map_with_another_column_layout(self):
        assert raised(ALLRAD, map_bytes(sections=["MRT", "CT"])).code == "E210"

    def test_statistics_without_a_required_column(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A10"], ws["B10"] = "IND1", "Leist-Kurz"
        ws["A11"], ws["B11"] = 1, "MRTS"
        err = raised(STATISTICS, wb_to_bytes(wb))
        assert err.code == "E111" and "DokDatum" in err.detail

    def test_statistics_without_rows(self):
        assert raised(STATISTICS, stats_bytes(rows=[])).code == "E112"

    def test_statistics_with_too_many_rows(self, monkeypatch):
        monkeypatch.setattr(input_handler, "MAX_DATA_ROWS", 1)
        rows = DEFAULT_ROWS + [(2, 1, "01.01.2024", "MRT Schädel", "MRTS")]
        assert raised(STATISTICS, stats_bytes(rows=rows)).code == "E113"

    def test_map_without_leistungen(self):
        assert raised(NRAD, map_bytes(codes=[])).code == "E211"

    def test_map_without_any_row(self):
        assert raised(NRAD, wb_to_bytes(openpyxl.Workbook())).code == "E211"

    def test_map_with_too_many_columns(self, monkeypatch):
        monkeypatch.setattr(input_handler, "MAX_MAP_COLUMNS", 1)
        assert raised(ALLRAD, map_bytes()).code == "E212"

    def test_map_with_too_many_rows(self, monkeypatch):
        monkeypatch.setattr(input_handler, "MAX_DATA_ROWS", 1)
        codes = DEFAULT_CODES + [("OTHER", "x", {"MRT": 1})]
        assert raised(ALLRAD, map_bytes(codes=codes)).code == "E212"


# ── Contract: everything core/ can say is translated ───────────────────────────

CORE = Path(input_handler.__file__).parent
VALIDATION_ERRORS = {"FileValidationError", "StatsValidationError", "MapValidationError"}
# Only raised when a file is read by path (CLI); the browser always sends bytes.
PATH_ONLY = ("Cannot resolve path", "File not found", "Path does not point to a file")


def render(node):
    """The text of a message expression with every {placeholder} shown as <X>; None if not literal."""
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    if isinstance(node, ast.JoinedStr):
        return "".join(p.value if isinstance(p, ast.Constant) else "<X>" for p in node.values)
    return None


def raised_messages():
    tree = ast.parse((CORE / "input_handler.py").read_text(encoding="utf-8"))
    found = []
    for node in ast.walk(tree):
        if (isinstance(node, ast.Raise) and isinstance(node.exc, ast.Call)
                and getattr(node.exc.func, "id", None) in VALIDATION_ERRORS):
            found.append((node.lineno, render(node.exc.args[0]) if node.exc.args else None))
    return found


def appended_warnings():
    found = []
    for filename in ("input_handler.py", "logic.py"):
        tree = ast.parse((CORE / filename).read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
                    and node.func.attr == "append" and isinstance(node.func.value, ast.Name)
                    and node.func.value.id == "errors" and node.args):
                found.append((f"{filename}:{node.lineno}", render(node.args[0])))
    return found


class TestContractWithCore:

    def test_every_raise_in_core_is_translated(self):
        found = raised_messages()
        assert len(found) >= 20, "found too few raise statements — is the extraction broken?"
        for lineno, message in found:
            assert message is not None, f"input_handler.py:{lineno}: message is not a string literal"
            if message.startswith(PATH_ONLY):
                continue
            assert classify_core_message(message) is not None, (
                f"input_handler.py:{lineno}: no translation for {message!r} — add it to local_app/errors.py"
            )

    def test_the_path_only_list_is_not_stale(self):
        messages = [m for _, m in raised_messages() if m]
        for prefix in PATH_ONLY:
            assert any(m.startswith(prefix) for m in messages), f"{prefix!r} is not raised any more"

    def test_every_warning_in_core_is_translated(self):
        found = appended_warnings()
        assert len(found) >= 8, "found too few warnings — is the extraction broken?"
        for where, message in found:
            assert message is not None, f"{where}: message is not a string literal"
            assert classify_hint(message) is not None, (
                f"{where}: no hint code for {message!r} — add it to local_app/errors.py"
            )

    @pytest.mark.parametrize("slot", SLOTS)
    def test_every_key_leads_to_a_catalog_code(self, slot):
        keys = {key for _, key in [
            ("", "unreadable"), ("", "invalid"), ("", "empty"), ("", "toolarge"), ("", "macros"),
            ("", "security"), ("", "not_stats"), ("", "column"), ("", "no_rows"), ("", "too_many_rows"),
            ("", "not_map"), ("", "map_empty"), ("", "map_big"),
        ]}
        for key in keys:
            assert file_error(slot, key).code in CATALOG

    def test_classification_keys_cover_all_prefixes(self):
        from local_app.errors import _CORE_PREFIXES
        for _, key in _CORE_PREFIXES:
            for slot in SLOTS:
                assert file_error(slot, key).code in CATALOG

    def test_core_raises_only_one_kind_of_error_for_a_refused_result(self):
        # One mapping (from_output_error) covers the output side; a new exception type
        # there must be handled in local_app/errors.py as well.
        tree = ast.parse((CORE / "output_handler.py").read_text(encoding="utf-8"))
        kinds = {getattr(n.exc.func, "id", None) for n in ast.walk(tree)
                 if isinstance(n, ast.Raise) and isinstance(n.exc, ast.Call)}
        assert kinds == {"OutputValidationError", "IOError"}       # IOError: the CLI cannot write the file

    def test_text_from_a_file_cannot_steer_the_classification(self):
        # A cell value that imitates another message does not change the verdict.
        message = "Potential formula injection detected in Leistung at row 11: '=File is empty'"
        assert classify_core_message(message) == "security"


class TestRefusedResult:

    def test_a_result_with_a_formula_is_e903_without_a_file(self):
        with pytest.raises(OutputValidationError) as info:
            build_xlsx_bytes(make_result(["=SUM(1)"], {"=SUM(1)": {2024: 1}}, [2024]))
        err = from_output_error(info.value)
        assert err.code == "E903" and err.slot is None
        assert "formula" in err.detail and "Auswertung!A" in err.detail

    def test_the_detail_never_contains_the_content_of_the_cell(self):
        with pytest.raises(OutputValidationError) as info:
            build_xlsx_bytes(make_result(["=SECRET(1)"], {"=SECRET(1)": {2024: 1}}, [2024]))
        assert "SECRET" not in from_output_error(info.value).detail


# ── Hints ──────────────────────────────────────────────────────────────────────

def evaluate(rows, codes, **stats_kwargs):
    """Run loaders and process() the way the app will, return the result."""
    stats = input_handler.load_statistics(stats_bytes(rows, **stats_kwargs))
    return process(stats, input_handler.load_map(map_bytes(codes)))


def by_code(hints):
    return {h["code"]: h for h in hints}


class TestHints:

    def test_real_messages_of_one_evaluation(self):
        rows = [
            (1, 1, "01.01.2024", "a", "MRTS"),
            (1, 2, "01.01.2024", "b", "ZZ1"),       # not in the map
            (2, 1, "01.01.2024", "c", "ZZ2"),       # not in the map
            (2, 2, "01.01.2024", "d", None),        # no Leist-Kurz: row skipped
        ]
        result = evaluate(rows, DEFAULT_CODES + [("NOSEC", "x", {})], total_l8=2)
        hints = by_code(summarize_hints(result["errors"]))
        assert hints["W101"]["count"] == 1
        assert hints["W201"]["count"] == 2
        assert hints["W104"]["count"] is None
        assert "W202" not in hints and "W103" not in hints

    def test_missing_metadata_and_count_give_one_hint_each(self):
        result = evaluate(DEFAULT_ROWS, DEFAULT_CODES, mitarbeiter="", befunddatum="")
        hints = by_code(summarize_hints(result["errors"]))
        assert set(hints) == {"W102", "W103"}      # two messages for the metadata become one hint

    def test_unassigned_leistungen_are_counted_by_the_number_in_the_message(self):
        rows = [(i, 1, "01.01.2024", "x", "NOSEC") for i in range(1, 4)]
        result = evaluate(rows, [("NOSEC", "x", {})], total_l8=3)
        hints = by_code(summarize_hints(result["errors"]))
        assert hints["W202"]["count"] == 3
        assert hints["W202"]["text"] == "3 Leistungen haben keine Zuordnung (am Ende der Excel aufgelistet)."

    def test_clean_evaluation_has_no_hints(self):
        result = evaluate(DEFAULT_ROWS, DEFAULT_CODES, total_l8=1)
        assert summarize_hints(result["errors"]) == []

    def test_grouping_sorts_by_code_and_sums_counts(self):
        messages = [
            "NOTE: Leistung 'A' ('a') not found in map — added as a separate row in the output.",
            "WARNING: Row 12 (IND1=1) — 'Leist-Kurz' is empty, row skipped.",
            "WARNING: Row 13 row 13 — cannot parse year from DokDatum 'x', row skipped.",
            "NOTE: Leistung 'B' ('b') not found in map — added as a separate row in the output.",
        ]
        hints = summarize_hints(messages)
        assert [h["code"] for h in hints] == ["W101", "W201"]
        assert hints[0]["text"] == "2 Zeilen wurden übersprungen."
        assert hints[1]["text"] == "2 Leistungscodes fehlen in der Map (eigene Zeilen in der Excel)."

    def test_thousands_use_the_german_separator(self):
        message = "NICHT ZUGEORDNET: 2,320 Leistungen (67 Codes) sind in der Map keinem Leistungsbereich zugeordnet"
        (h,) = summarize_hints([message])
        assert h["count"] == 2320
        assert h["text"] == "2.320 Leistungen haben keine Zuordnung (am Ende der Excel aufgelistet)."

    def test_singular_and_plural(self):
        assert hint("W101", 1)["text"] == "1 Zeile wurde übersprungen."
        assert hint("W101", 5)["text"] == "5 Zeilen wurden übersprungen."
        assert hint("W401", 1)["text"] == "1 Leistung der anderen Abteilung wurde nicht ausgewertet."
        assert hint("W401", 1234)["text"] == "1.234 Leistungen der anderen Abteilung wurden nicht ausgewertet."

    def test_uncounted_hint_has_one_text(self):
        assert hint("W301") == {
            "code": "W301",
            "text": "Soll-Werte konnten nicht geladen werden; die Spalte „Benötigt“ fehlt.",
            "count": None,
        }

    def test_selection_note_is_not_a_hint(self):
        assert summarize_hints([f"{SELECTION_NOTE_PREFIX} Bereiche: Allrad; Map: map.xlsx"]) == []

    def test_names_that_look_like_formulas_are_counted_per_code(self):
        rows = [(1, 1, "01.01.2024", "-a", "MRTS"), (2, 1, "01.01.2024", "-a", "MRTS"), (3, 1, "01.01.2024", "@b", "ZZ1")]
        result = evaluate(rows, DEFAULT_CODES, total_l8=3)
        hints = by_code(summarize_hints(result["errors"]))
        assert hints["W105"]["count"] == 2 and hints["W105"]["text"].startswith("Bei 2 Leistungscodes")
        assert hints["W201"]["count"] == 1                    # ZZ1 is not in the map: its own row, named by its code

    def test_a_missing_code_with_a_taken_name_is_still_one_missing_code(self):
        # core adds the code to the row name then; the note has another ending but is the same hint
        rows = [(1, 1, "01.01.2024", "MRT", "ZZ1"), (2, 1, "01.01.2024", "MRT", "ZZ2")]
        result = evaluate(rows, DEFAULT_CODES, total_l8=2)
        assert by_code(summarize_hints(result["errors"]))["W201"]["count"] == 2

    def test_unknown_message_is_not_lost(self):
        hints = summarize_hints(["something nobody translated", "and another one"])
        assert hints == [hint("W900", 2)]
        assert hints[0]["text"] == "2 weitere Hinweise (siehe Excel, „Hinweise und Fehler“)."

    def test_no_messages(self):
        assert summarize_hints([]) == []
