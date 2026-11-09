"""
Tests for three hardening measures:
  - the loaders read only the columns they use (a tiny sheet with one far-away
    cell must not keep them busy),
  - damaged content in a file is reported as a validation error, not as an
    internal error,
  - the finished workbook is checked once more: no formula, no link.
"""

import asyncio
import io
import json
import re
import sys
import time
import zipfile
from unittest.mock import MagicMock

import openpyxl
import pytest

for _mod in ("iii", "iii.types"):          # web/worker.py imports the iii framework
    sys.modules.setdefault(_mod, MagicMock())

from cli import main as cli_main                                   # noqa: E402
from core import input_handler                                     # noqa: E402
from core.input_handler import (                                   # noqa: E402
    FileValidationError,
    MapValidationError,
    StatsValidationError,
    load_map,
    load_reference_data,
    load_statistics,
)
from core import output_handler                                    # noqa: E402
from core.logic import process                                     # noqa: E402
from core.output_handler import (                                  # noqa: E402
    OutputValidationError,
    _assert_text_only,
    _build_workbook,
    build_xlsx_bytes,
    save_output,
)
from web import worker                                             # noqa: E402

from .helpers import (                                             # noqa: E402
    assert_reference_unusable,
    make_map_wb,
    make_multipart_body,
    make_reference_wb,
    make_result,
    make_stats_wb,
    wb_to_bytes,
)

SHEET = "xl/worksheets/sheet1.xml"
ROWS = [(1, 1, "01.01.2024", "MRT Schädel", "MRTS")]
CODES = [("MRTS", "MRT Schädel", {"MRT": 1})]


# ── Builders ───────────────────────────────────────────────────────────────────

def _stats() -> bytes:
    return wb_to_bytes(make_stats_wb(ROWS, total_l8=1))


def _map() -> bytes:
    return wb_to_bytes(make_map_wb(CODES))


def _reference() -> bytes:
    return wb_to_bytes(make_reference_wb({"MRT": 100}))


def _patch_zip(data: bytes, name: str, fn) -> bytes:
    """A copy of the zip in which entry `name` has the content fn(old content)."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        found = False
        for info in zin.infolist():
            content = zin.read(info.filename)
            if info.filename == name:
                found = True
                patched = fn(content)
                assert patched != content, f"patching {name} changed nothing"
                content = patched
            zout.writestr(info.filename, content)
        assert found, f"{name} is not in the archive"
    return out.getvalue()


def _without_entry(data: bytes, name: str) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename != name:
                zout.writestr(info.filename, zin.read(info.filename))
    return out.getvalue()


def _with_far_cell(data: bytes, row: int, col: str) -> bytes:
    """A copy of the workbook with one more cell far away; the file stays a few KB."""
    far = f'<row r="{row}"><c r="{col}{row}" t="n"><v>1</v></c></row>'.encode()

    def patch(d):
        d = d.replace(b"</sheetData>", far + b"</sheetData>", 1)
        return re.sub(rb'<dimension ref="[^"]*"', f'<dimension ref="A1:{col}{row}"'.encode(), d, count=1)

    return _patch_zip(data, SHEET, patch)


def _truncated(data: bytes) -> bytes:
    """The sheet cut in half: well-formed zip, broken XML."""
    return _patch_zip(data, SHEET, lambda d: d[: len(d) // 2])


# ── Stall protection ───────────────────────────────────────────────────────────

class TestOnlyTheUsedColumnsAreRead:
    """Without a column limit, a far-away cell makes the reader invent a million padded rows (15 s)."""

    def test_far_away_cell_does_not_keep_the_statistics_loader_busy(self):
        data = _with_far_cell(_stats(), 1_048_576, "XFD")
        start = time.perf_counter()
        result = load_statistics(data)
        assert time.perf_counter() - start < 8       # about 0.5 s; 15 s before the limit
        assert [e["leist_kurz"] for e in result["leistungen"]] == ["MRTS"]

    def test_far_away_cell_does_not_keep_the_map_loader_busy(self):
        data = _with_far_cell(_map(), 1_048_576, "XFD")
        start = time.perf_counter()
        result = load_map(data)
        assert time.perf_counter() - start < 8
        assert list(result["leistungen"]) == ["MRTS"]

    def test_far_away_cell_does_not_keep_the_reference_loader_busy(self):
        data = _with_far_cell(_reference(), 1_048_576, "XFD")
        start = time.perf_counter()
        result = load_reference_data(data)
        assert time.perf_counter() - start < 8
        assert result["amounts"] == {"MRT": 100}

    def test_the_unbounded_loops_ask_for_a_limited_width(self, monkeypatch):
        from openpyxl.worksheet._read_only import ReadOnlyWorksheet
        calls = []
        original = ReadOnlyWorksheet.iter_rows

        def spy(self, *args, **kwargs):
            calls.append(kwargs)
            return original(self, *args, **kwargs)

        monkeypatch.setattr(ReadOnlyWorksheet, "iter_rows", spy)
        load_statistics(_stats())
        load_map(_map())
        load_reference_data(_reference())
        unbounded = [kw for kw in calls if "max_row" not in kw]
        assert len(unbounded) >= 3                                  # statistics, map and reference data loops
        assert all(kw.get("max_col", 10**6) <= 60 for kw in unbounded), unbounded

    def test_cells_beyond_the_used_columns_are_ignored_and_not_fatal(self):
        wb = make_stats_wb(ROWS, total_l8=1)
        wb.active["Z11"] = "column 26"
        wb.active["AN11"] = "column 40"
        stats = load_statistics(wb_to_bytes(wb))
        assert stats["leistungen"] == load_statistics(_stats())["leistungen"]

        wb = make_map_wb(CODES)
        wb.active["AZ2"] = "column 52"
        assert load_map(wb_to_bytes(wb))["leistungen"] == load_map(_map())["leistungen"]

    def test_columns_further_right_than_the_default_are_still_read_when_they_are_used(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        # Like the real exports, the row has a value in column A — a row is "empty" when
        # its first 25 columns are, which is how the loader has always judged it.
        for col, header in ((1, "Patienten Name"), (27, "IND1"), (28, "IND2"), (29, "DokDatum"), (30, "Leistung"), (31, "Leist-Kurz")):
            ws.cell(row=10, column=col).value = header
        for col, value in ((1, "Muster"), (27, 1), (28, 1), (29, "01.01.2024"), (30, "MRT Schädel"), (31, "MRTS")):
            ws.cell(row=11, column=col).value = value
        stats = load_statistics(wb_to_bytes(wb))
        assert [(e["leist_kurz"], e["leistung"], e["year"], e["ind1"]) for e in stats["leistungen"]] == [
            ("MRTS", "MRT Schädel", 2024, 1)
        ]
        assert stats["total_leistungen"] == 1

    def test_ordinary_files_give_the_same_result_as_before(self):
        stats = load_statistics(_stats())
        assert stats["leistungen"] == [{"leist_kurz": "MRTS", "leistung": "MRT Schädel", "year": 2024, "ind1": 1, "row": 11}]
        assert stats["total_leistungen"] == 1 and stats["total_documents"] == 1


# ── Damaged files ──────────────────────────────────────────────────────────────

class TestDamagedContent:

    def test_cut_off_sheet_in_a_statistics_file(self):
        with pytest.raises(StatsValidationError, match=r"Cannot read statistics file \(damaged or malformed content\): ParseError"):
            load_statistics(_truncated(_stats()))

    def test_cut_off_sheet_in_a_map(self):
        with pytest.raises(MapValidationError, match=r"Cannot read map file \(damaged or malformed content\): ParseError") as info:
            load_map(_truncated(_map()))
        assert info.value.recoverable is False

    def test_missing_sheet_part(self):
        with pytest.raises(StatsValidationError):
            load_statistics(_without_entry(_stats(), SHEET))
        with pytest.raises(MapValidationError):
            load_map(_without_entry(_map(), SHEET))

    def test_the_original_exception_is_kept_as_the_cause(self):
        with pytest.raises(StatsValidationError) as info:
            load_statistics(_truncated(_stats()))
        assert info.value.__cause__ is not None

    def test_the_reference_loader_still_never_raises(self):
        assert_reference_unusable(load_reference_data(_truncated(_reference())), "ist beschädigt oder nicht vollständig lesbar")

    def test_our_own_errors_are_passed_on_unchanged(self):
        # a statistics file that is no statistics file
        with pytest.raises(StatsValidationError, match="Column header 'Leist-Kurz' not found"):
            load_statistics(_map())
        # a map with another column layout stays "recoverable"
        with pytest.raises(MapValidationError, match="F1 is") as info:
            load_map(wb_to_bytes(make_map_wb(CODES, sections=["MRT", "CT"])))
        assert info.value.recoverable is True
        # formula text stays a FileValidationError
        wb = make_stats_wb(ROWS, total_l8=1)
        wb.active["D11"].value = "=1+1"
        wb.active["D11"].data_type = "s"
        with pytest.raises(FileValidationError, match="formula injection"):
            load_statistics(wb_to_bytes(wb))


class _Writer:
    def __init__(self):
        self.chunks = []
        self.stream = self

    async def write(self, data):
        self.chunks.append(data)

    def end(self, data):
        self.chunks.append(data)

    def close(self):
        pass


class _Response:
    def __init__(self):
        self.code = None
        self.writer = _Writer()

    async def status(self, code):
        self.code = code

    async def headers(self, headers):
        pass


class _Request:
    method = "POST"

    def __init__(self, body, content_type):
        self.headers = {"content-type": content_type, "content-length": str(len(body))}
        self.request_body = self
        self._body = body

    async def read_all(self):
        return self._body


def _upload(statistics: bytes):
    body, content_type = make_multipart_body({"statistics": statistics})
    response = _Response()
    asyncio.run(worker.handle_convert(_Request(body, content_type), response))
    return response, b"".join(response.writer.chunks)


class TestWhatTheUserSees:

    @pytest.fixture(autouse=True)
    def server_files(self, monkeypatch):
        monkeypatch.setattr(worker, "_MAP_DATA", load_map(_map()))
        monkeypatch.setattr(worker, "_REFERENCE_DATA", {})

    def test_web_valid_upload_still_works(self):
        response, body = _upload(_stats())
        assert response.code == 200
        assert openpyxl.load_workbook(io.BytesIO(body)).active.max_row > 1

    def test_web_damaged_upload_is_422_with_a_readable_message(self):
        response, body = _upload(_truncated(_stats()))
        assert response.code == 422
        assert "Cannot read statistics file (damaged or malformed content)" in json.loads(body)["error"]

    def test_cli_damaged_file_is_a_clear_error_line(self, tmp_path, monkeypatch, capsys):
        stats, map_ = tmp_path / "stats.xlsx", tmp_path / "map.xlsx"
        stats.write_bytes(_truncated(_stats()))
        map_.write_bytes(_map())
        monkeypatch.setattr(sys, "argv", ["cli", str(stats), str(map_), "--output", str(tmp_path / "out")])
        assert cli_main.main() == 1
        err = capsys.readouterr().err
        assert "ERROR (statistics): Cannot read statistics file (damaged or malformed content)" in err
        assert "UNEXPECTED" not in err

    def test_cli_damaged_map_is_a_clear_error_line(self, tmp_path, monkeypatch, capsys):
        stats, map_ = tmp_path / "stats.xlsx", tmp_path / "map.xlsx"
        stats.write_bytes(_stats())
        map_.write_bytes(_truncated(_map()))
        monkeypatch.setattr(sys, "argv", ["cli", str(stats), str(map_), "--output", str(tmp_path / "out")])
        assert cli_main.main() == 1
        assert "ERROR (map): Cannot read map file (damaged or malformed content)" in capsys.readouterr().err

    def test_cli_valid_files_still_work(self, tmp_path, monkeypatch):
        stats, map_ = tmp_path / "stats.xlsx", tmp_path / "map.xlsx"
        stats.write_bytes(_stats())
        map_.write_bytes(_map())
        monkeypatch.setattr(sys, "argv", ["cli", str(stats), str(map_), "--output", str(tmp_path / "out")])
        assert cli_main.main() == 0
        assert len(list((tmp_path / "out").glob("*.xlsx"))) == 1


# ── Final check of the result ──────────────────────────────────────────────────

def _result_with_section(name: str) -> dict:
    return make_result([name], {name: {2024: 3}}, [2024])


class TestFinalCheck:

    def test_a_formula_in_the_result_is_not_delivered(self):
        with pytest.raises(OutputValidationError, match=r"formula in Auswertung!A\d+"):
            build_xlsx_bytes(_result_with_section("=HYPERLINK(\"http://evil.example\",\"x\")"))

    def test_nothing_is_written_to_disk_either(self, tmp_path):
        with pytest.raises(OutputValidationError):
            save_output(_result_with_section("=SUM(1)"), output_dir=str(tmp_path))
        assert list(tmp_path.iterdir()) == []

    def test_a_link_in_the_result_is_not_delivered(self):
        wb = openpyxl.Workbook()
        wb.active["A1"] = "click"
        wb.active["A1"].hyperlink = "http://evil.example"
        with pytest.raises(OutputValidationError, match="link"):
            _assert_text_only(wb)

    def test_normal_results_pass_with_and_without_reference_data(self):
        result = make_result(["MRT gesamt", "CT"], {"MRT gesamt": {2024: 3}}, [2024])
        build_xlsx_bytes(result)
        _, data = build_xlsx_bytes(result, reference_data={"amounts": {"MRT gesamt": 5}, "combinations": {}, "order": ["MRT gesamt"]})
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            assert b"<cfRule" in zf.read(SHEET)       # the colour rules use formulas, which are fine

    def test_the_check_does_not_change_the_workbook(self, monkeypatch):
        # Build without the check, then count the cells before and after running it.
        monkeypatch.setattr(output_handler, "_assert_text_only", lambda wb: None)
        _, wb = _build_workbook(make_result(["MRT gesamt"], {"MRT gesamt": {2024: 3}}, [2024]))
        cells = len(wb.active._cells)
        _assert_text_only(wb)
        assert len(wb.active._cells) == cells

    def test_the_error_is_a_value_error_with_a_message_naming_the_cell(self):
        with pytest.raises(ValueError, match="potential formula injection"):
            build_xlsx_bytes(_result_with_section("=1+1"))

    def test_second_line_of_defence_when_the_input_check_is_bypassed(self, monkeypatch):
        # Without the section-name check a formula header would reach the output; the final check stops it.
        wb = make_map_wb(CODES)
        cell = wb.active.cell(row=1, column=7)
        cell.value, cell.data_type = "=EVIL(1)", "s"
        monkeypatch.setattr(input_handler, "_check_cell_for_injection", lambda *args, **kwargs: None)
        map_data = load_map(wb_to_bytes(wb))
        assert "=EVIL(1)" in map_data["sections"]
        result = process(load_statistics(_stats()), map_data)
        with pytest.raises(OutputValidationError, match="formula"):
            build_xlsx_bytes(result)
