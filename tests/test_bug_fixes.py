"""
Tests for the findings of the bug and misuse hunt (2026-11-10). Each class names the
finding; every test here failed, or showed the silent behaviour, before the fix.

  - a Leistung missing from the map was counted into a section of the same name,
    and two missing codes with one name shared a row,
  - a map with two section columns of one name lost the first column,
  - a map value such as 1.9 or -1 was counted as 1 or 0 without a word,
  - one formula-like Leistung text refused the whole statistics file,
  - a report without a DokDatum took the date of the report before it,
  - a very long Mitarbeiter made the file name too long to save,
  - the CLI replaced an earlier result of the same name without a word,
  - the web worker parsed on its event loop, enforced its size limit only through
    Content-Length and sent the download name unencoded.
"""

import asyncio
import io
import json
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from unittest.mock import MagicMock

import openpyxl
import pytest

for _mod in ("iii", "iii.types"):          # web/worker.py imports the iii framework
    sys.modules.setdefault(_mod, MagicMock())

from cli import main as cli_main                                   # noqa: E402
from core import output_handler                                    # noqa: E402
from core.input_handler import (                                   # noqa: E402
    FileValidationError,
    MapValidationError,
    load_map,
    load_statistics,
)
from core.logic import process, unique_row_label                   # noqa: E402
from core.output_handler import _make_filename, build_xlsx_bytes, save_output   # noqa: E402
from web import worker                                             # noqa: E402

from .helpers import (                                             # noqa: E402
    make_map_wb,
    make_multipart_body,
    make_reference_wb,
    make_stats_wb,
    wb_to_bytes,
)

SECTIONS = ["DL Gefäße", "MRT", "CT", "Drainage"]
MAP_CODES = [
    ("MRTS", "MRT Schädel", {"MRT": 1}),
    ("CTT", "CT Thorax", {"CT": 1}),
    ("DRN", "Drainage", {"Drainage": 1}),
]


def _map(codes=MAP_CODES, sections=SECTIONS) -> dict:
    return load_map(wb_to_bytes(make_map_wb(codes, sections=sections)))


def _stats(rows, **kwargs) -> dict:
    return load_statistics(wb_to_bytes(make_stats_wb(rows, **kwargs)))


def _rows(*items):
    """(leistung, code, count) → statistics rows, one report each, all in 2024."""
    rows, ind1 = [], 0
    for leistung, code, count in items:
        for _ in range(count):
            ind1 += 1
            rows.append((ind1, 1, "01.03.2024", leistung, code))
    return rows


def _sheet(result, reference_data=None) -> dict:
    """Row label → (Gesamt, Benötigt or None) of the main table of the built workbook."""
    _, data = build_xlsx_bytes(result, reference_data=reference_data)
    ws = openpyxl.load_workbook(io.BytesIO(data)).active
    head = [c.value for c in ws[1]]
    total_col = head.index("Gesamt")
    needed_col = head.index("Benötigt") if "Benötigt" in head else None
    table = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            break
        assert row[0] not in table, f"row {row[0]!r} appears twice"
        table[row[0]] = (row[total_col], row[needed_col] if needed_col is not None else None)
    return table


# ── 1. Missing codes stay apart from sections and from each other ─────────────

class TestMissingCodesKeepTheirOwnRow:

    def test_a_missing_code_named_like_a_section_is_not_counted_into_it(self):
        # The real case: code RXXDR "Drainage" (not in the NRAD map) and the section "Drainage".
        result = process(_stats(_rows(("Drainage", "DRN", 2), ("Drainage", "RXXDR", 5))), _map())
        assert result["data"]["Drainage"]["Total"] == 2                 # was 7
        assert result["data"]["Drainage (RXXDR)"]["Total"] == 5
        assert result["unmatched_rows"] == {"Drainage (RXXDR)": "RXXDR"}
        note = next(e for e in result["errors"] if "RXXDR" in e)
        assert note.startswith("NOTE: Leistung 'RXXDR' ('Drainage') not found in map")
        assert "separate row 'Drainage (RXXDR)'" in note

    def test_the_comparison_ignores_case_like_the_target_lookup(self):
        result = process(_stats(_rows(("drainage", "RXXDR", 3))), _map())
        assert result["data"]["Drainage"]["Total"] == 0
        assert result["data"]["drainage (RXXDR)"]["Total"] == 3

    def test_two_missing_codes_with_one_name_get_a_row_each(self):
        result = process(_stats(_rows(("Spezial", "XA", 2), ("Spezial", "XB", 3), ("spezial", "XC", 1))), _map())
        assert result["data"]["Spezial (XA)"]["Total"] == 2
        assert result["data"]["Spezial (XB)"]["Total"] == 3
        assert result["data"]["spezial (XC)"]["Total"] == 1
        assert "Spezial" not in result["data"]

    def test_one_code_with_two_texts_is_counted_once_completely(self):
        # The second text used to open a row that was never written: its count was lost.
        result = process(_stats(_rows(("Text eins", "XA", 2), ("Text zwei", "XA", 3))), _map())
        assert result["data"]["Text eins"]["Total"] == 5
        assert "Text zwei" not in result["data"]

    def test_a_code_without_text_that_equals_a_section_name(self):
        result = process(_stats(_rows((None, "CT", 4), ("CT Thorax", "CTT", 1))), _map())
        assert result["data"]["CT"]["Total"] == 1
        assert result["data"]["CT (nicht in der Map)"]["Total"] == 4

    def test_a_name_that_is_taken_even_with_its_code_gets_a_number(self):
        taken = {"mrt", "mrt (x)"}
        assert unique_row_label("MRT", "X", taken) == "MRT (X) [2]"
        assert "mrt (x) [2]" in taken
        assert unique_row_label("Frei", "Y", taken) == "Frei"

    def test_a_name_that_is_free_is_kept_and_the_note_is_unchanged(self):
        result = process(_stats(_rows(("Sonderleistung", "XS", 2))), _map())
        assert result["data"]["Sonderleistung"]["Total"] == 2
        assert ("NOTE: Leistung 'XS' ('Sonderleistung') not found in map — "
                "added as a separate row in the output.") in result["errors"]

    def test_every_counted_row_ends_up_in_exactly_one_row(self):
        items = [("MRT Schädel", "MRTS", 3), ("MRT", "XM", 4), ("MRT", "XN", 2), ("Drainage", "RXXDR", 1),
                 ("CT Thorax", "CTT", 5), (None, "MRT", 6)]
        result = process(_stats(_rows(*items)), _map())
        assert sum(row["Total"] for row in result["data"].values()) == sum(n for *_, n in items)
        assert len(result["unmatched_rows"]) == 4

    def test_the_workbook_shows_the_rows_apart(self):
        result = process(_stats(_rows(("Drainage", "DRN", 2), ("Drainage", "RXXDR", 5))), _map())
        table = _sheet(result)
        assert table["Drainage"][0] == 2 and table["Drainage (RXXDR)"][0] == 5


class TestMissingCodesAndTheReferenceFile:

    def _reference(self, amounts, combinations=None, order=None):
        from core.input_handler import load_reference_data
        return load_reference_data(wb_to_bytes(make_reference_wb(amounts, combinations=combinations, order=order)))

    def test_a_missing_code_named_like_a_reference_row_gets_no_target(self):
        # "Punktion" is a reference row but no section of this map.
        result = process(_stats(_rows(("Punktion", "XP", 3))), _map())
        table = _sheet(result, self._reference({"MRT": 100, "Punktion": 50}))
        assert "Punktion" not in table
        assert table["Punktion (XP)"] == (3, None)

    def test_a_missing_code_is_never_summed_into_a_combination_row(self):
        reference = self._reference({"MRT": 100}, combinations={"MRT plus": ["MRT", "Extra"]},
                                    order=["MRT", "MRT plus"])
        result = process(_stats(_rows(("MRT Schädel", "MRTS", 2), ("Extra", "XE", 7))), _map())
        table = _sheet(result, reference)
        assert "MRT plus" not in table                       # component "Extra" is no section: dropped
        assert table["Extra (XE)"][0] == 7

    def test_a_missing_code_named_like_a_combination_row_is_not_dropped(self):
        reference = self._reference({"MRT": 100}, combinations={"Bildgebung": ["MRT", "CT"]},
                                    order=["MRT", "CT", "Bildgebung"])
        result = process(_stats(_rows(("MRT Schädel", "MRTS", 2), ("Bildgebung", "XB", 4))), _map())
        table = _sheet(result, reference)
        assert table["Bildgebung"][0] == 2                   # the combination: MRT + CT
        assert table["Bildgebung (XB)"][0] == 4              # used to vanish behind it

    def test_without_a_collision_nothing_is_renamed(self):
        result = process(_stats(_rows(("Sonderleistung", "XS", 2))), _map())
        table = _sheet(result, self._reference({"MRT": 100}))
        assert table["Sonderleistung"] == (2, None)
        assert table["MRT"] == (0, 100)


# ── 2. Duplicate section names ────────────────────────────────────────────────

class TestDuplicateSectionNames:

    def _map_with_header(self, names, values, include_underscore=False):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=2).value = "Leist0"
        for offset, name in enumerate(names):
            ws.cell(row=1, column=6 + offset).value = name
            ws.cell(row=2, column=6 + offset).value = values[offset]
        ws.cell(row=2, column=2).value = "CODE"
        return load_map(wb_to_bytes(wb), include_underscore_columns=include_underscore)

    def test_two_columns_with_one_name_refuse_the_map(self):
        with pytest.raises(MapValidationError, match=r"Duplicate section name in map: 'MRT' .* columns G and I"):
            self._map_with_header(["DL Gefäße", "MRT", "CT", "MRT"], [0, 1, 0, 0])

    def test_names_that_differ_only_in_surrounding_spaces_are_the_same_name(self):
        with pytest.raises(MapValidationError, match="Duplicate section name"):
            self._map_with_header(["DL Gefäße", "MRT", "MRT "], [0, 1, 1])

    def test_ignored_underscore_columns_may_share_a_name(self):
        result = self._map_with_header(["DL Gefäße", "_x", "_x"], [1, 2, 3])
        assert result["sections"] == ["DL Gefäße"]
        with pytest.raises(MapValidationError, match="Duplicate section name"):
            self._map_with_header(["DL Gefäße", "_x", "_x"], [1, 2, 3], include_underscore=True)

    def test_a_long_duplicate_name_is_shortened_in_the_message(self):
        name = "N" * 5000
        with pytest.raises(MapValidationError) as info:
            self._map_with_header(["DL Gefäße", name, name], [0, 1, 1])
        assert len(str(info.value)) < 300


# ── 3. Map values must be whole numbers ───────────────────────────────────────

class TestMapValues:

    def _load(self, *values):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=2).value = "Leist0"
        ws.cell(row=1, column=6).value = "DL Gefäße"
        for row, value in enumerate(values, 2):
            ws.cell(row=row, column=2).value = f"C{row}"
            ws.cell(row=row, column=6).value = value
        return load_map(wb_to_bytes(wb))

    @pytest.mark.parametrize("value, expected", [
        (None, 0), (0, 0), (1, 1), (2, 2), (3, 3), (1.0, 1), (2.0, 2), (0.0, 0),
        ("1", 1), (" 2 ", 2), ("", 0), ("   ", 0),
    ])
    def test_accepted(self, value, expected):
        assert self._load(value)["leistungen"]["C2"]["sections"]["DL Gefäße"] == expected

    @pytest.mark.parametrize("value", [
        1.9, 0.5, 1.5, -1, -0.5, -2.0, "x", "1.5", "1e3", "-1", "١", "9" * 20, True, False,
        datetime(2024, 1, 1),
    ])
    def test_refused_with_the_cell(self, value):
        with pytest.raises(MapValidationError, match=r"Invalid value in map: .* F2 = "):
            self._load(value)

    def test_the_message_names_five_cells_and_counts_the_rest(self):
        with pytest.raises(MapValidationError) as info:
            self._load(*([1.5] * 8))
        text = str(info.value)
        assert "Found 8 other cell(s)" in text and "F2 = 1.5" in text and "F6 = 1.5" in text
        assert "F7" not in text and "and 3 more" in text

    def test_a_long_text_value_is_shortened_in_the_message(self):
        with pytest.raises(MapValidationError) as info:
            self._load("x" * 10_000)
        assert len(str(info.value)) < 400

    def test_ignored_underscore_columns_are_not_checked(self):
        wb = make_map_wb([("CODE", "d", {"MRT": 1})], sections=["DL Gefäße", "MRT", "_Zeit"])
        wb.active.cell(row=2, column=8).value = 12.5             # minutes, not a count
        assert load_map(wb_to_bytes(wb))["sections"] == ["DL Gefäße", "MRT"]
        with pytest.raises(MapValidationError, match="Invalid value"):
            load_map(wb_to_bytes(wb), include_underscore_columns=True)


# ── 4. A formula-like Leistung text no longer refuses the file ────────────────

class TestFormulaLikeLeistungText:

    @pytest.mark.parametrize("text", ["-Zuschlag", "+ Kontrast", "@Sonder", "=1+1", "|Pipe"])
    def test_the_file_is_read_and_the_text_left_out(self, text):
        rows = _rows(("MRT Schädel", "MRTS", 2), (text, "MRTS", 3), (text, "XQ", 1))
        wb = make_stats_wb(rows)
        for row in range(11, 11 + len(rows)):                   # keep "=…" a text cell
            wb.active.cell(row=row, column=4).data_type = "s"
        stats = load_statistics(wb_to_bytes(wb))
        result = process(stats, _map())
        assert result["data"]["MRT"]["Total"] == 5               # the code is what counts
        assert result["data"]["XQ"]["Total"] == 1                # the code stands in for the text
        assert all(row["leistung"] in (None, "MRT Schädel") for row in stats["leistungen"])
        warnings = [e for e in stats["errors"] if "looks like a formula" in e]
        assert len(warnings) == 2
        assert "'MRTS'" in warnings[0] and "3 row(s), first at row 13" in warnings[0]
        assert text not in " ".join(stats["errors"])
        build_xlsx_bytes(result)                                  # no formula reaches the result

    def test_a_formula_like_code_still_refuses_the_file(self):
        with pytest.raises(FileValidationError, match="formula injection"):
            _stats(_rows(("MRT", "-MRTS", 1)))

    def test_a_long_formula_like_code_is_shortened_in_the_message(self):
        with pytest.raises(FileValidationError) as info:
            _stats(_rows(("MRT", "-" + "A" * 10_000, 1)))
        assert len(str(info.value)) < 200


# ── 5. A report without a DokDatum ────────────────────────────────────────────

class TestDateOfANewReport:

    def test_a_new_report_without_a_date_does_not_borrow_the_previous_one(self):
        rows = [
            (1, 1, "01.12.2023", "MRT Schädel", "MRTS"),
            (1, 2, None, "MRT Schädel", "MRTS"),         # same report: carries 2023
            (2, 1, None, "CT Thorax", "CTT"),            # new report, no date
            (2, 2, None, "CT Thorax", "CTT"),
            (3, 1, "05.01.2024", "CT Thorax", "CTT"),
        ]
        stats = _stats(rows)
        assert [(r["leist_kurz"], r["year"]) for r in stats["leistungen"]] == [("MRTS", 2023), ("MRTS", 2023), ("CTT", 2024)]
        skipped = [e for e in stats["errors"] if "cannot parse year" in e]
        assert len(skipped) == 2 and "(IND1=2)" in skipped[0]
        result = process(stats, _map())
        assert any(e.startswith("COUNT MISMATCH") for e in result["errors"])

    def test_without_an_ind1_column_the_date_is_still_carried(self):
        wb = make_stats_wb([(None, 1, "01.12.2023", "MRT Schädel", "MRTS"), (None, 1, None, "CT Thorax", "CTT")])
        wb.active["A10"] = None                                  # no IND1 column
        stats = load_statistics(wb_to_bytes(wb))
        assert [r["year"] for r in stats["leistungen"]] == [2023, 2023]


# ── 6. File names ─────────────────────────────────────────────────────────────

class TestFileName:

    def test_a_normal_name_is_unchanged(self):
        assert _make_filename("KUNZ_A", "01.01.2024-31.12.2024") == "KUNZ_A01012024-31122024Auswertung.xlsx"

    @pytest.mark.parametrize("mitarbeiter", ["A" * 300, "ä" * 300, "漢" * 300, "😀" * 300])
    def test_a_very_long_name_fits_the_file_system(self, mitarbeiter, tmp_path):
        name = _make_filename(mitarbeiter, "1" * 300)
        assert len(name.encode("utf-8")) <= 255 - len("_999")
        name.encode("utf-8").decode("utf-8")                     # no character cut in half
        result = process(_stats(_rows(("MRT Schädel", "MRTS", 1)), mitarbeiter=mitarbeiter), _map())
        path = save_output(result, output_dir=str(tmp_path))
        assert Path(path).is_file()


# ── 7. The CLI does not replace an earlier result ─────────────────────────────

class TestNoSilentOverwrite:

    def _result(self, count):
        return process(_stats(_rows(("MRT Schädel", "MRTS", count))), _map())

    def _total(self, path):
        ws = openpyxl.load_workbook(path).active
        return next(r[2] for r in ws.iter_rows(values_only=True) if r[0] == "MRT")

    def test_a_second_result_of_the_same_name_gets_a_number(self, tmp_path):
        first = save_output(self._result(1), output_dir=str(tmp_path))
        second = save_output(self._result(2), output_dir=str(tmp_path))
        third = save_output(self._result(3), output_dir=str(tmp_path))
        assert Path(second).name == Path(first).stem + "_2.xlsx"
        assert Path(third).name == Path(first).stem + "_3.xlsx"
        assert [self._total(p) for p in (first, second, third)] == [1, 2, 3]

    def test_overwrite_replaces_it(self, tmp_path):
        first = save_output(self._result(1), output_dir=str(tmp_path))
        again = save_output(self._result(2), output_dir=str(tmp_path), overwrite=True)
        assert again == first and self._total(first) == 2
        assert len(list(tmp_path.iterdir())) == 1

    def test_a_failed_write_leaves_no_half_file(self, tmp_path, monkeypatch):
        class Broken(io.BytesIO):
            def write(self, data):
                raise OSError("disk full")

        def open_then_fail(path, mode):
            Path(path).touch()                     # the file exists now, as after a real open
            return Broken()

        monkeypatch.setattr(output_handler, "open", open_then_fail, raising=False)
        with pytest.raises(IOError, match="disk full"):
            save_output(self._result(1), output_dir=str(tmp_path))
        assert list(tmp_path.iterdir()) == []

    def test_the_cli_keeps_the_first_file_and_says_so(self, tmp_path, monkeypatch, capsys):
        stats, map_ = tmp_path / "s.xlsx", tmp_path / "m.xlsx"
        stats.write_bytes(wb_to_bytes(make_stats_wb(_rows(("MRT Schädel", "MRTS", 1)))))
        map_.write_bytes(wb_to_bytes(make_map_wb(MAP_CODES, sections=SECTIONS)))
        out = tmp_path / "out"
        argv = ["cli", str(stats), str(map_), "--output", str(out)]
        monkeypatch.setattr(sys, "argv", argv)
        assert cli_main.main() == 0
        assert "already exists" not in capsys.readouterr().out
        assert cli_main.main() == 0
        printed = capsys.readouterr().out
        assert "_2.xlsx" in printed and "already exists there and was kept" in printed
        assert len(list(out.iterdir())) == 2
        monkeypatch.setattr(sys, "argv", argv + ["--overwrite"])
        assert cli_main.main() == 0
        assert "already exists" not in capsys.readouterr().out
        assert len(list(out.iterdir())) == 2


# ── 8. Web worker ─────────────────────────────────────────────────────────────

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
        self.sent_headers = {}
        self.writer = _Writer()

    async def status(self, code):
        self.code = code

    async def headers(self, headers):
        self.sent_headers.update(headers)

    def json(self):
        return json.loads(b"".join(self.writer.chunks))


class _ChunkedBody:
    """A request body like the SDK's ChannelReader: async iteration over chunks."""

    def __init__(self, body: bytes, chunk: int = 64 * 1024):
        self._body, self._chunk = body, chunk
        self.chunks_read = 0
        self.closed = False

    async def __aiter__(self):
        for offset in range(0, len(self._body), self._chunk):
            self.chunks_read += 1
            yield self._body[offset: offset + self._chunk]

    async def close_async(self):
        self.closed = True


class _Request:
    method = "POST"

    def __init__(self, body: bytes, content_type: str, content_length: bool = True):
        self.headers = {"content-type": content_type}
        if content_length:
            self.headers["content-length"] = str(len(body))
        self.request_body = _ChunkedBody(body)


def _stats_bytes(count=1, mitarbeiter="KUNZ_A"):
    return wb_to_bytes(make_stats_wb(_rows(("MRT Schädel", "MRTS", count)), mitarbeiter=mitarbeiter))


async def _convert(statistics: bytes, content_length: bool = True):
    body, content_type = make_multipart_body({"statistics": statistics})
    request, response = _Request(body, content_type, content_length), _Response()
    await worker.handle_convert(request, response)
    return request, response


class TestWebWorker:

    @pytest.fixture(autouse=True)
    def server_files(self, monkeypatch):
        monkeypatch.setattr(worker, "_MAP_DATA", _map())
        monkeypatch.setattr(worker, "_REFERENCE_DATA", {})

    def test_a_normal_upload_still_works(self):
        _, response = asyncio.run(_convert(_stats_bytes(3)))
        assert response.code == 200
        assert openpyxl.load_workbook(io.BytesIO(b"".join(response.writer.chunks))).active["A2"].value

    def test_without_content_length_the_limit_stops_the_reading(self, monkeypatch):
        monkeypatch.setattr(worker, "_MAX_BODY_BYTES", 1024 * 1024)
        body = b"x" * (20 * 1024 * 1024)
        request, response = _Request(body, "multipart/form-data; boundary=x", content_length=False), _Response()
        asyncio.run(worker.handle_convert(request, response))
        assert response.code == 413
        assert request.request_body.chunks_read == 17          # 1 MiB / 64 KiB + 1, not 320
        assert request.request_body.closed

    def test_a_wrong_content_length_does_not_help(self, monkeypatch):
        monkeypatch.setattr(worker, "_MAX_BODY_BYTES", 1024 * 1024)
        request = _Request(b"x" * (3 * 1024 * 1024), "multipart/form-data; boundary=x")
        request.headers["content-length"] = "10"
        response = _Response()
        asyncio.run(worker.handle_convert(request, response))
        assert response.code == 413

    def test_a_body_at_the_limit_is_read(self):
        assert asyncio.run(worker._read_body(_ChunkedBody(b"a" * 1000, chunk=100), 1000)) == b"a" * 1000
        assert asyncio.run(worker._read_body(_ChunkedBody(b"a" * 1001, chunk=100), 1000)) is None

    def test_a_reader_without_chunks_is_still_limited(self):
        class Whole:
            def __init__(self, body):
                self.body = body

            async def read_all(self):
                return self.body

        assert asyncio.run(worker._read_body(Whole(b"abc"), 3)) == b"abc"
        assert asyncio.run(worker._read_body(Whole(b"abcd"), 3)) is None

    def test_the_event_loop_keeps_running_while_a_file_is_parsed(self, monkeypatch):
        real = worker.load_statistics

        def slow(source):
            time.sleep(0.6)
            return real(source)

        monkeypatch.setattr(worker, "load_statistics", slow)

        async def scenario():
            ticks = 0

            async def ticker():
                nonlocal ticks
                while True:
                    await asyncio.sleep(0.01)
                    ticks += 1

            task = asyncio.create_task(ticker())
            _, response = await _convert(_stats_bytes())
            task.cancel()
            return response, ticks

        response, ticks = asyncio.run(scenario())
        assert response.code == 200
        assert ticks >= 20                                       # was 0: the loop stood still

    def test_the_timeout_answers_while_a_file_is_still_parsed(self, monkeypatch):
        monkeypatch.setattr(worker, "_PROCESSING_TIMEOUT_S", 0.3)
        release = threading.Event()

        def stuck(source):
            release.wait(5)
            raise RuntimeError("too late")

        monkeypatch.setattr(worker, "load_statistics", stuck)

        async def scenario():
            started = time.monotonic()
            _, response = await _convert(_stats_bytes())
            answered = time.monotonic() - started
            release.set()                      # let the abandoned conversion end
            return response, answered

        response, answered = asyncio.run(scenario())
        assert response.code == 503
        assert answered < 2.5                  # was: only after the parsing had finished

    def test_at_most_two_conversions_at_a_time(self, monkeypatch):
        real, running, peak, lock = worker.load_statistics, [0], [0], threading.Lock()

        def counted(source):
            with lock:
                running[0] += 1
                peak[0] = max(peak[0], running[0])
            time.sleep(0.2)
            with lock:
                running[0] -= 1
            return real(source)

        monkeypatch.setattr(worker, "load_statistics", counted)

        async def scenario():
            return await asyncio.gather(*(_convert(_stats_bytes()) for _ in range(5)))

        results = asyncio.run(scenario())
        assert [response.code for _, response in results] == [200] * 5
        assert peak[0] == 2

    def test_the_download_name_keeps_its_umlauts(self):
        _, response = asyncio.run(_convert(_stats_bytes(mitarbeiter="MÜLLER_Ä")))
        header = response.sent_headers["Content-Disposition"]
        assert header.startswith('attachment; filename="MULLER_A')
        assert "filename*=UTF-8''M%C3%9CLLER_%C3%84" in header
        header.encode("ascii")                                    # a header must be plain ASCII

    def test_the_fallback_name_is_never_empty_and_never_breaks_the_header(self):
        assert worker._content_disposition("漢字.xlsx").startswith('attachment; filename="__.xlsx"')
        assert worker._content_disposition("").startswith('attachment; filename="Auswertung.xlsx"')
        assert '"' not in worker._content_disposition('a"b\r\n.xlsx').split("filename=")[1].split(";")[0][1:-1]

    def test_a_long_error_message_is_shortened(self):
        response = _Response()
        asyncio.run(worker._send_json_error(response, 422, "E" * 5000))
        assert len(response.json()["error"]) == 500

    def test_a_refused_upload_quotes_only_a_short_piece_of_the_cell(self):
        wb = make_stats_wb(_rows(("MRT", "-" + "B" * 20_000, 1)))
        _, response = asyncio.run(_convert(wb_to_bytes(wb)))
        assert response.code == 422 and len(response.json()["error"]) < 200
