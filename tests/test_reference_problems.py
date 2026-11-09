"""
A reference file that cannot be used, or only in part, must never make the
"Benötigt" column vanish without a word:

  - load_reference_data reports what it could not use ("problems"),
  - the result sheet lists those problems,
  - the CLI warns, the web worker logs them and lets them reach every result.

The loader still never raises: the reference file stays optional.
"""

import asyncio
import io
import logging
import sys
import zipfile
from unittest.mock import MagicMock

import openpyxl
import pytest

for _mod in ("iii", "iii.types"):          # web/worker.py imports the iii framework
    sys.modules.setdefault(_mod, MagicMock())

from cli import main as cli_main                                   # noqa: E402
from core import input_handler                                     # noqa: E402
from core.input_handler import load_map, load_reference_data, load_statistics   # noqa: E402
from core.logic import process                                     # noqa: E402
from core.output_handler import build_xlsx_bytes                   # noqa: E402
from web import worker                                             # noqa: E402

from .helpers import (                                             # noqa: E402
    assert_reference_unusable,
    make_corrupt_xlsx_bytes,
    make_map_wb,
    make_reference_wb,
    make_stats_wb,
    wb_to_bytes,
)
from .test_hardening import _upload                                # noqa: E402

SHEET = "xl/worksheets/sheet1.xml"
ROWS = [(1, 1, "01.01.2024", "MRT Schädel", "MRTS")]
CODES = [("MRTS", "MRT Schädel", {"MRT": 1})]
NOTE_HEADER = "Hinweis: Probleme mit der Referenzdatei"


# ── Builders ───────────────────────────────────────────────────────────────────

def _stats() -> bytes:
    return wb_to_bytes(make_stats_wb(ROWS, total_l8=1))


def _map() -> bytes:
    return wb_to_bytes(make_map_wb(CODES))


def _good_reference() -> bytes:
    return wb_to_bytes(make_reference_wb({"MRT": 100}))


def _reference(*rows) -> bytes:
    """A reference workbook with a header row and the given (name, amount[, combination]) rows."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Leistungsbereich", "Benötigt", "Includiert"])
    for row in rows:
        ws.append(list(row))
    for cells in ws.iter_rows():
        for cell in cells:
            if isinstance(cell.value, str):
                cell.data_type = "s"             # text, even where it starts with "="
    return wb_to_bytes(wb)


def _with_dtd(data: bytes) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            content = zin.read(info.filename)
            if info.filename == SHEET:
                content = content.replace(b"<worksheet", b'<!DOCTYPE worksheet [<!ENTITY x "y">]><worksheet', 1)
            zout.writestr(info.filename, content)
    return out.getvalue()


def _texts(xlsx: bytes) -> list[str]:
    ws = openpyxl.load_workbook(io.BytesIO(xlsx)).active
    return [str(c.value) for row in ws.iter_rows() for c in row if c.value is not None]


def _evaluate(reference_data) -> bytes:
    result = process(load_statistics(_stats()), load_map(_map()))
    return build_xlsx_bytes(result, reference_data=reference_data)[1]


# ── The loader ─────────────────────────────────────────────────────────────────

class TestLoaderReportsProblems:

    def test_a_complete_file_has_no_problems(self):
        data = load_reference_data(_reference(("MRT", 100), ("CT", 50), ("Kombi", None, "MRT;CT"), ("Leer", 0)))
        assert data["problems"] == []
        assert data["amounts"] == {"MRT": 100, "CT": 50}

    def test_the_real_shape_with_all_columns_is_clean(self):
        assert load_reference_data(_good_reference())["problems"] == []

    # ── the file cannot be used at all ───────────────────────────────────────

    def test_a_missing_file(self, tmp_path):
        assert_reference_unusable(load_reference_data(str(tmp_path / "gone.xlsx")), "wurde nicht gefunden")

    def test_a_missing_file_given_as_a_path_object(self, tmp_path):
        assert_reference_unusable(load_reference_data(tmp_path / "gone.xlsx"), "wurde nicht gefunden")

    def test_a_directory_is_not_a_file(self, tmp_path):
        assert_reference_unusable(load_reference_data(str(tmp_path)), "wurde von der Datei- und Sicherheitsprüfung abgelehnt")

    @pytest.mark.parametrize("name, data", [
        ("entity declaration", lambda: _with_dtd(_good_reference())),
        ("not a zip", make_corrupt_xlsx_bytes),
        ("empty", lambda: b""),
    ])
    def test_a_file_that_the_validation_refuses(self, name, data):
        assert_reference_unusable(load_reference_data(data()), "wurde von der Datei- und Sicherheitsprüfung abgelehnt")

    def test_a_file_that_the_validation_refuses_on_disk(self, tmp_path):
        path = tmp_path / "reference_amount.xlsx"
        path.write_bytes(_with_dtd(_good_reference()))
        assert_reference_unusable(load_reference_data(str(path)), "wurde von der Datei- und Sicherheitsprüfung abgelehnt")

    def test_a_file_that_cannot_be_opened(self, monkeypatch):
        def boom(*args, **kwargs):
            raise OSError("disk on fire")
        monkeypatch.setattr(input_handler.openpyxl, "load_workbook", boom)
        assert_reference_unusable(load_reference_data(_good_reference()), "ist beschädigt oder nicht lesbar")

    def test_a_workbook_without_a_sheet(self, monkeypatch):
        class NoSheet:
            active = None
            def close(self): pass
        monkeypatch.setattr(input_handler.openpyxl, "load_workbook", lambda *a, **k: NoSheet())
        assert_reference_unusable(load_reference_data(_good_reference()), "enthält kein Tabellenblatt")

    def test_a_file_that_fails_while_it_is_read(self, monkeypatch):
        class Broken:
            @property
            def active(self):
                raise RuntimeError("damaged")
            def close(self): pass
        monkeypatch.setattr(input_handler.openpyxl, "load_workbook", lambda *a, **k: Broken())
        assert_reference_unusable(load_reference_data(_good_reference()), "ist beschädigt oder nicht vollständig lesbar")

    def test_an_unexpected_validation_error(self, monkeypatch):
        def boom(*args, **kwargs):
            raise MemoryError("no")
        monkeypatch.setattr(input_handler, "_resolve_source", boom)
        assert_reference_unusable(load_reference_data(_good_reference()), "konnte nicht geprüft werden")

    def test_the_loader_still_never_raises(self, monkeypatch):
        for bad in (b"", b"x" * 50, make_corrupt_xlsx_bytes(), "/nonexistent.xlsx", "\x00"):
            data = load_reference_data(bad)
            assert set(data) == {"amounts", "combinations", "order", "problems"}

    # ── nothing leaks into the text ──────────────────────────────────────────

    def test_no_server_path_in_the_problem(self, tmp_path):
        secret = tmp_path / "geheimes-verzeichnis"
        for source in (str(secret / "reference_amount.xlsx"), str(tmp_path)):
            text = " ".join(load_reference_data(source)["problems"])
            assert "geheimes" not in text and str(tmp_path) not in text

    def test_no_file_name_or_exception_text_in_a_rejection(self):
        text = " ".join(load_reference_data(_with_dtd(_good_reference()))["problems"])
        assert "reference_amount.xlsx" not in text and "DTD" not in text and "entity" not in text.lower()

    # ── some rows cannot be used ─────────────────────────────────────────────

    def test_a_negative_amount(self):
        data = load_reference_data(_reference(("MRT", 100), ("CT", -5)))
        assert data["amounts"] == {"MRT": 100}
        assert data["problems"] == ["Zeile 3 („CT“): Der Wert in „Benötigt“ ist negativ — es wird kein Zielwert angezeigt."]

    def test_zero_means_no_required_amount_and_is_not_a_problem(self):
        data = load_reference_data(_reference(("MRT", 100), ("CT", 0)))
        assert data["problems"] == []

    def test_an_implausibly_large_amount(self):
        data = load_reference_data(_reference(("MRT", 100), ("CT", 5_000_000)))
        assert data["amounts"] == {"MRT": 100}
        assert len(data["problems"]) == 1 and "größer als 1,000,000" in data["problems"][0] and "Zeile 3" in data["problems"][0]

    def test_an_amount_that_is_not_a_number(self):
        data = load_reference_data(_reference(("MRT", 100), ("CT", "viele")))
        assert data["amounts"] == {"MRT": 100}
        assert data["problems"] == ["Zeile 3 („CT“): Der Wert in „Benötigt“ ist keine Zahl — es wird kein Zielwert angezeigt."]
        assert "viele" not in data["problems"][0]

    @pytest.mark.parametrize("value", ["inf", "nan", "1e999"])
    def test_amounts_that_break_the_conversion(self, value):
        data = load_reference_data(_reference(("MRT", 100), ("CT", value)))
        assert data["amounts"] == {"MRT": 100}
        assert len(data["problems"]) == 1 and data["problems"][0].startswith("Zeile 3")

    @pytest.mark.parametrize("name", ['=HYPERLINK("http://evil.example","x")', "+cmd|' /C calc'!A0", "@SUM(1)"])
    def test_a_section_name_that_looks_like_a_formula(self, name):
        data = load_reference_data(_reference(("MRT", 100), (name, 7)))
        assert data["amounts"] == {"MRT": 100}
        assert data["problems"] == ["Zeile 3: Der Bereichsname sieht wie eine Formel aus — Zeile übersprungen."]
        assert "evil" not in data["problems"][0] and "calc" not in data["problems"][0]

    def test_a_row_with_an_amount_but_no_name(self):
        data = load_reference_data(_reference(("MRT", 100), (None, 7), ("  ", None, "MRT;CT")))
        assert data["amounts"] == {"MRT": 100}
        assert data["problems"] == [
            "Zeile 3: Es steht kein Bereichsname in Spalte A — Zeile übersprungen.",
            "Zeile 4: Es steht kein Bereichsname in Spalte A — Zeile übersprungen.",
        ]

    def test_a_row_with_only_a_note_further_right_is_not_a_problem(self):
        data = load_reference_data(_reference(("MRT", 100), (None, None, None, "Notiz", "mehr")))
        assert data["problems"] == []

    def test_a_long_name_is_shortened_in_the_note(self):
        data = load_reference_data(_reference(("M" * 200, -1)))
        assert len(data["problems"][0]) < 200 and "M" * 61 not in data["problems"][0] and "..." in data["problems"][0]

    def test_too_many_rows(self, monkeypatch):
        monkeypatch.setattr(input_handler, "_MAX_REFERENCE_ROWS", 3)
        data = load_reference_data(_reference(*[(f"S{i}", i + 1) for i in range(6)]))
        assert list(data["amounts"]) == ["S0", "S1", "S2"]
        assert data["problems"] == ["Mehr als 3 Zeilen — der Rest der Datei wurde nicht gelesen."]

    def test_a_file_without_any_amount(self):
        data = load_reference_data(_reference(("MRT", None), ("CT", 0)))
        assert data["amounts"] == {}
        assert data["problems"] == ["Die Referenzdatei enthält keine Zielwerte (Spalte B) — die Spalte „Benötigt“ fehlt."]

    def test_a_header_only_file(self):
        assert len(load_reference_data(_reference())["problems"]) == 1

    def test_the_good_rows_are_kept_next_to_the_bad_ones(self):
        data = load_reference_data(_reference(("A", 1), ("B", -1), ("C", "x"), ("D", 4), ("E", 5_000_000), ("F", 6)))
        assert data["amounts"] == {"A": 1, "D": 4, "F": 6}
        assert [p.split(":")[0] for p in data["problems"]] == ["Zeile 3 („B“)", "Zeile 4 („C“)", "Zeile 6 („E“)"]

    def test_many_problems_are_listed_up_to_a_limit_and_then_counted(self, monkeypatch):
        monkeypatch.setattr(input_handler, "_MAX_REFERENCE_PROBLEMS", 3)
        data = load_reference_data(_reference(("OK", 1), *[(f"S{i}", -1) for i in range(7)]))
        assert len(data["problems"]) == 4
        assert data["problems"][-1] == "… und 4 weitere Probleme (nicht einzeln aufgeführt)."

    def test_a_hostile_file_cannot_flood_the_list(self):
        data = load_reference_data(_reference(*[(f"S{i}", -1) for i in range(3000)]))
        assert len(data["problems"]) <= input_handler._MAX_REFERENCE_PROBLEMS + 2

    def test_the_problems_also_reach_the_log(self, caplog):
        with caplog.at_level(logging.WARNING, logger="core.input_handler"):
            load_reference_data(_reference(("MRT", 100), ("CT", -5)))
        assert "non-positive value" in caplog.text

    def test_the_old_amounts_wrapper_is_unchanged(self):
        assert input_handler.load_reference_amounts(_reference(("MRT", 100), ("CT", -5))) == {"MRT": 100}


# ── The result sheet ───────────────────────────────────────────────────────────

class TestResultSheet:

    def test_a_rejected_file_is_explained_on_the_sheet(self):
        texts = _texts(_evaluate(load_reference_data(_with_dtd(_good_reference()))))
        assert NOTE_HEADER in texts
        assert any("wurde von der Datei- und Sicherheitsprüfung abgelehnt" in t for t in texts)
        assert "Benötigt" not in texts

    def test_partial_problems_are_listed_and_the_good_rows_still_show(self):
        texts = _texts(_evaluate(load_reference_data(_reference(("MRT", 100), ("CT", -5)))))
        assert NOTE_HEADER in texts
        assert any(t.startswith("    •  Zeile 3 („CT“)") for t in texts)
        assert "Benötigt" in texts

    def test_a_clean_file_adds_no_note(self):
        assert NOTE_HEADER not in _texts(_evaluate(load_reference_data(_good_reference())))

    def test_no_reference_data_adds_no_note(self):
        for reference_data in (None, {}, {"amounts": {}, "combinations": {}, "order": []}):
            assert NOTE_HEADER not in _texts(_evaluate(reference_data))

    def test_a_clean_result_is_byte_identical_with_and_without_the_new_field(self):
        base = load_reference_data(_good_reference())
        old_shape = {k: v for k, v in base.items() if k != "problems"}
        assert _texts(_evaluate(base)) == _texts(_evaluate(old_shape))

    def test_a_note_that_looks_like_a_formula_stays_text(self):
        data = {"amounts": {}, "combinations": {}, "order": [], "problems": ["=1+1", '=HYPERLINK("http://evil.example","x")']}
        xlsx = _evaluate(data)
        with zipfile.ZipFile(io.BytesIO(xlsx)) as zf:
            assert b"<f>" not in zf.read(SHEET) and b"<hyperlink" not in zf.read(SHEET)
        assert "    •  =1+1" in _texts(xlsx)

    def test_a_long_note_is_cut(self):
        data = {"amounts": {}, "combinations": {}, "order": [], "problems": ["x" * 5000]}
        assert max(len(t) for t in _texts(_evaluate(data))) < 400

    @pytest.mark.parametrize("problems", ["text", 5, {"a": 1}, [None, "", 0]])
    def test_a_malformed_problems_field_is_ignored(self, problems):
        data = {"amounts": {}, "combinations": {}, "order": [], "problems": problems}
        assert NOTE_HEADER not in _texts(_evaluate(data))


# ── The CLI ────────────────────────────────────────────────────────────────────

class TestCli:

    @pytest.fixture
    def workdir(self, tmp_path, monkeypatch):
        (tmp_path / "stats.xlsx").write_bytes(_stats())
        (tmp_path / "map.xlsx").write_bytes(_map())
        monkeypatch.chdir(tmp_path)
        return tmp_path

    def _run(self, workdir, monkeypatch, capsys, *extra):
        monkeypatch.setattr(sys, "argv", ["cli", "stats.xlsx", "map.xlsx", "--output", str(workdir / "out"), *extra])
        code = cli_main.main()
        captured = capsys.readouterr()
        outputs = list((workdir / "out").glob("*.xlsx"))
        return code, captured, _texts(outputs[0].read_bytes()) if outputs else None

    def test_no_reference_file_at_all_is_a_normal_run(self, workdir, monkeypatch, capsys):
        code, captured, texts = self._run(workdir, monkeypatch, capsys)
        assert code == 0
        assert "No reference file found" in captured.out
        assert "WARNING" not in captured.err
        assert NOTE_HEADER not in texts

    def test_a_good_default_file_is_a_normal_run(self, workdir, monkeypatch, capsys):
        (workdir / "reference_amount.xlsx").write_bytes(_good_reference())
        code, captured, texts = self._run(workdir, monkeypatch, capsys)
        assert code == 0 and "WARNING" not in captured.err and "1 section(s)" in captured.out
        assert NOTE_HEADER not in texts and "Benötigt" in texts

    def test_a_rejected_default_file_is_reported(self, workdir, monkeypatch, capsys):
        (workdir / "reference_amount.xlsx").write_bytes(_with_dtd(_good_reference()))
        code, captured, texts = self._run(workdir, monkeypatch, capsys)
        assert code == 0                                               # the reference stays optional
        assert "WARNING (reference): Die Referenzdatei wurde von der Datei- und Sicherheitsprüfung abgelehnt" in captured.err
        assert "No reference amounts loaded" in captured.out
        assert NOTE_HEADER in texts and "Benötigt" not in texts

    def test_a_named_file_that_is_missing_is_reported(self, workdir, monkeypatch, capsys):
        code, captured, texts = self._run(workdir, monkeypatch, capsys, "--reference", "nirgends.xlsx")
        assert code == 0
        assert "WARNING (reference): Die Referenzdatei wurde nicht gefunden" in captured.err
        assert NOTE_HEADER in texts
        assert "nirgends" not in " ".join(texts)

    def test_a_partly_usable_file_is_reported_row_by_row(self, workdir, monkeypatch, capsys):
        (workdir / "ref.xlsx").write_bytes(_reference(("MRT", 100), ("CT", -5)))
        code, captured, texts = self._run(workdir, monkeypatch, capsys, "--reference", "ref.xlsx")
        assert code == 0
        assert "WARNING (reference): Zeile 3 („CT“)" in captured.err
        assert "Benötigt" in texts and NOTE_HEADER in texts


# ── The web worker ─────────────────────────────────────────────────────────────

class TestWebWorker:

    def _load(self, monkeypatch, path):
        monkeypatch.setenv("REFERENCE_FILE", str(path))
        return worker._load_server_reference()

    def test_a_good_file_is_loaded_quietly(self, tmp_path, monkeypatch, caplog):
        path = tmp_path / "reference_amount.xlsx"
        path.write_bytes(_good_reference())
        with caplog.at_level(logging.WARNING):
            data = self._load(monkeypatch, path)
        assert data["problems"] == [] and data["amounts"] == {"MRT": 100}
        assert not [r for r in caplog.records if r.levelno >= logging.WARNING and "eference" in r.getMessage()]

    @pytest.mark.parametrize("content", [lambda: _with_dtd(_good_reference()), make_corrupt_xlsx_bytes, None])
    def test_a_rejected_or_missing_file_is_logged_as_an_error_and_kept_as_a_problem(self, tmp_path, monkeypatch, caplog, content):
        path = tmp_path / "reference_amount.xlsx"
        if content:
            path.write_bytes(content())
        with caplog.at_level(logging.ERROR):
            data = self._load(monkeypatch, path)
        assert data["amounts"] == {} and len(data["problems"]) == 1
        assert [r for r in caplog.records if r.levelno == logging.ERROR and "problem" in r.getMessage()]
        assert data["problems"][0] in [r.getMessage().strip() for r in caplog.records if r.levelno == logging.ERROR]

    def test_when_loading_itself_fails_the_result_still_says_so(self, monkeypatch, caplog):
        def boom(path):
            raise RuntimeError("unexpected")
        monkeypatch.setattr(worker, "load_reference_data", boom)
        with caplog.at_level(logging.ERROR):
            data = worker._load_server_reference()
        assert data["amounts"] == {} and len(data["problems"]) == 1
        assert "Die Referenzdatei konnte nicht geladen werden" in data["problems"][0]

    def test_every_result_carries_the_problems(self, tmp_path, monkeypatch):
        path = tmp_path / "reference_amount.xlsx"
        path.write_bytes(_with_dtd(_good_reference()))
        monkeypatch.setattr(worker, "_MAP_DATA", load_map(_map()))
        monkeypatch.setattr(worker, "_REFERENCE_DATA", self._load(monkeypatch, path))
        for _ in range(2):
            response, body = _upload(_stats())
            assert response.code == 200
            texts = _texts(body)
            assert NOTE_HEADER in texts and "Benötigt" not in texts

    def test_a_good_reference_gives_the_same_result_as_before(self, tmp_path, monkeypatch):
        path = tmp_path / "reference_amount.xlsx"
        path.write_bytes(_good_reference())
        monkeypatch.setattr(worker, "_MAP_DATA", load_map(_map()))
        monkeypatch.setattr(worker, "_REFERENCE_DATA", self._load(monkeypatch, path))
        response, body = _upload(_stats())
        texts = _texts(body)
        assert response.code == 200 and "Benötigt" in texts and NOTE_HEADER not in texts
