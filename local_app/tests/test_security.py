"""
Tests for local_app.security: core's validation with the codes of a slot, the
check of the finished Excel, and safe file names.

How core itself handles hostile files (every kind of DTD, encodings, chunk
boundaries, …) is tested in tests/test_xml_safety.py and tests/test_hardening.py.
Here it is only checked that each kind of refusal reaches the user with the right code.
"""

import io
import zipfile

import openpyxl
import pytest

from core import input_handler, output_handler
from core.logic import process
from core.output_handler import build_xlsx_bytes
from local_app.errors import ALLRAD, NRAD, SLOT_LABELS, STATISTICS, AppError
from local_app.security import assert_no_formulas, safe_name, scan_xlsx
from tests.helpers import (
    inject_vba_into_xlsx,
    make_corrupt_xlsx_bytes,
    make_map_wb,
    make_path_traversal_xlsx_bytes,
    make_zip_bomb_xlsx_bytes,
    make_zip_without_workbook_xml,
    wb_to_bytes,
)

from .helpers import DEFAULT_CODES, SHEET, add_zip_entry, force_text, map_bytes, patch_zip, stats_bytes, with_doctype

# (slot, builder of a valid file for that slot)
SLOT_BUILDERS = [(STATISTICS, stats_bytes), (ALLRAD, map_bytes), (NRAD, map_bytes)]


def series(slot):
    """'E1' for the statistics file, 'E2' for a map."""
    return "E1" if slot == STATISTICS else "E2"


def billion_laughs():
    entities = '<!ENTITY a0 "lol">'
    for i in range(1, 10):
        body = f"&a{i - 1};" * 10
        entities += f'<!ENTITY a{i} "{body}">'
    return f"<!DOCTYPE worksheet [{entities}]>"


def damaged_entry(_build):
    """A zip whose entry has the right size but a wrong checksum."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_STORED) as zf:
        zf.writestr("xl/workbook.xml", b"<workbook/>")
    return buffer.getvalue().replace(b"<workbook/>", b"<workbook//")


def encrypted_entry(_build):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as zf:
        zf.writestr("xl/workbook.xml", b"<workbook/>")
    data = bytearray(buffer.getvalue())
    data[data.find(b"PK\x03\x04") + 6] |= 0x01             # general purpose flag: encrypted
    data[data.find(b"PK\x01\x02") + 8] |= 0x01
    return bytes(data)


# name: (builds the hostile file from the valid file of the slot, suffix of the expected code)
HOSTILE = {
    "internal entity": (lambda build: with_doctype(build(), '<!DOCTYPE worksheet [<!ENTITY x "EINGESCHLEUST">]>'), "05"),
    "external entity": (lambda build: with_doctype(build(), '<!DOCTYPE worksheet [<!ENTITY x SYSTEM "file:///etc/passwd">]>'), "05"),
    "billion laughs": (lambda build: with_doctype(build(), billion_laughs()), "05"),
    "DTD in an unusual entry": (lambda build: add_zip_entry(build(), "xl/extra.dat", b"<!DOCTYPE x>"), "05"),
    "UTF-16 sheet": (lambda build: patch_zip(build(), SHEET, lambda d: d.decode("utf-8").encode("utf-16")), "05"),
    "declared ISO-8859-1": (lambda build: add_zip_entry(
        build(), "xl/extra.xml", b'<?xml version="1.0" encoding="ISO-8859-1"?><root/>'), "05"),
    "zip bomb": (lambda build: make_zip_bomb_xlsx_bytes(), "05"),
    "path traversal": (lambda build: make_path_traversal_xlsx_bytes(), "05"),
    "macros": (lambda build: inject_vba_into_xlsx(build()), "04"),
    "not a zip": (lambda build: make_corrupt_xlsx_bytes(), "01"),
    "no workbook.xml": (lambda build: make_zip_without_workbook_xml(), "01"),
    "damaged entry": (damaged_entry, "01"),
    "encrypted entry": (encrypted_entry, "01"),
    "empty": (lambda build: b"", "02"),
}


# ── scan_xlsx ──────────────────────────────────────────────────────────────────

class TestScanXlsx:

    @pytest.mark.parametrize("slot, build", SLOT_BUILDERS)
    @pytest.mark.parametrize("name", HOSTILE)
    def test_a_hostile_file_gets_the_code_of_its_problem_and_of_its_slot(self, name, slot, build):
        make, suffix = HOSTILE[name]
        with pytest.raises(AppError) as info:
            scan_xlsx(make(build), slot)
        err = info.value
        assert err.code == series(slot) + suffix
        assert err.slot == slot
        assert SLOT_LABELS[slot] in err.detail            # core names the file by its slot

    @pytest.mark.parametrize("name", HOSTILE)
    def test_it_refuses_exactly_what_core_refuses(self, name):
        data = HOSTILE[name][0](stats_bytes)
        with pytest.raises(input_handler.FileValidationError):
            input_handler.validate_xlsx_bytes(data)
        with pytest.raises(AppError):
            scan_xlsx(data, STATISTICS)

    @pytest.mark.parametrize("slot, build", SLOT_BUILDERS)
    def test_a_file_over_the_size_limit(self, slot, build, monkeypatch):
        monkeypatch.setattr(input_handler, "MAX_FILE_SIZE_MB", 0)
        with pytest.raises(AppError) as info:
            scan_xlsx(build(), slot)
        assert info.value.code == series(slot) + "03"

    @pytest.mark.parametrize("slot, build", SLOT_BUILDERS)
    def test_a_file_that_unpacks_to_too_much(self, slot, build, monkeypatch):
        monkeypatch.setattr(input_handler, "_MAX_UNCOMPRESSED_MB", 0)
        with pytest.raises(AppError) as info:
            scan_xlsx(build(), slot)
        assert info.value.code == series(slot) + "05"

    @pytest.mark.parametrize("slot, build", SLOT_BUILDERS)
    def test_valid_files_pass(self, slot, build):
        scan_xlsx(build(), slot)

    def test_a_valid_file_with_the_parts_excel_adds_passes(self):
        data = add_zip_entry(
            stats_bytes(), "userCustomization/customUI.xml",
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><customUI xmlns="x"/>',
        )
        data = add_zip_entry(data, "xl/printerSettings/printerSettings1.bin", bytes(range(256)) * 20)
        data = add_zip_entry(data, "xl/_rels/extra.xml.rels", b'<?xml version="1.0"?><Relationships/>')
        scan_xlsx(data, STATISTICS)

    def test_an_unknown_slot_fails_loudly_even_for_a_valid_file(self):
        with pytest.raises(ValueError):
            scan_xlsx(stats_bytes(), "elsewhere")


# ── assert_no_formulas ─────────────────────────────────────────────────────────

def map_with_header(name):
    wb = make_map_wb(DEFAULT_CODES)
    force_text(wb.active.cell(row=1, column=7), name)         # G1, the "MRT" section
    return wb_to_bytes(wb)


def real_output(**kwargs):
    stats = input_handler.load_statistics(stats_bytes())
    result = process(stats, input_handler.load_map(map_with_header(kwargs.get("header", "MRT"))))
    return build_xlsx_bytes(result, reference_data=kwargs.get("reference"))[1]


class TestAssertNoFormulas:

    def test_a_normal_evaluation_passes_even_with_conditional_formatting(self):
        data = real_output(reference={"amounts": {"MRT": 5}, "combinations": {}, "order": ["MRT"]})
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            assert b"<cfRule" in zf.read(SHEET)                # the colour rules use <formula>, which is fine
        assert_no_formulas(data)

    def test_a_formula_cell_is_rejected(self):
        wb = openpyxl.Workbook()
        wb.active["A1"] = "=1+1"
        with pytest.raises(AppError) as info:
            assert_no_formulas(wb_to_bytes(wb))
        assert info.value.code == "E903" and "Formel" in info.value.detail

    @pytest.mark.parametrize("element", [
        b'<f t="shared" ref="A1:A2" si="0">1+1</f>',           # a shared formula with its attributes
        b'<f t="shared" si="0"/>',                              # its self-closing follower
        b'<f/>',                                                # self-closing without attributes
        b'<f t="array" ref="A1">1+1</f>',
        b'<f\tt="shared" si="0"/>',                             # other white space after the name
        b'<f\nt="shared" si="0"/>',
    ])
    def test_every_form_of_a_formula_element_is_rejected(self, element):
        wb = openpyxl.Workbook()
        wb.active["A1"] = "=1+1"
        data = patch_zip(wb_to_bytes(wb), SHEET, lambda d: d.replace(b"<f>1+1</f>", element))
        with pytest.raises(AppError) as info:
            assert_no_formulas(data)
        assert info.value.code == "E903" and "Formel" in info.value.detail

    def test_a_formula_on_a_later_sheet_is_found(self):
        wb = openpyxl.Workbook()
        wb.active["A1"] = "harmlos"
        wb.create_sheet("zwei")["B2"] = "=2+2"
        with pytest.raises(AppError) as info:
            assert_no_formulas(wb_to_bytes(wb))
        assert info.value.code == "E903" and "sheet2.xml" in info.value.detail

    def test_a_damaged_entry_is_refused(self):
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_STORED) as zf:
            zf.writestr(SHEET, b"<worksheet><sheetData/></worksheet>")
        data = buffer.getvalue().replace(b"<sheetData/>", b"<sheetData//")      # same size, wrong checksum
        with pytest.raises(AppError) as info:
            assert_no_formulas(data)
        assert info.value.code == "E903" and "nicht lesbar" in info.value.detail

    def test_a_link_is_rejected(self):
        wb = openpyxl.Workbook()
        wb.active["A1"] = "klick"
        wb.active["A1"].hyperlink = "http://evil.example"
        with pytest.raises(AppError) as info:
            assert_no_formulas(wb_to_bytes(wb))
        assert info.value.code == "E903" and "Link" in info.value.detail

    @pytest.mark.parametrize("part", [
        "xl/vbaProject.bin", "xl/externalLinks/externalLink1.xml", "xl/embeddings/oleObject1.bin",
    ])
    def test_forbidden_parts_are_rejected(self, part):
        data = add_zip_entry(real_output(), part, b"x")
        with pytest.raises(AppError) as info:
            assert_no_formulas(data)
        assert info.value.code == "E903" and part in info.value.detail

    def test_macros_injected_the_core_way(self):
        with pytest.raises(AppError):
            assert_no_formulas(inject_vba_into_xlsx(real_output()))

    def test_not_a_zip(self):
        with pytest.raises(AppError) as info:
            assert_no_formulas(b"nope")
        assert info.value.code == "E903"

    def test_the_third_line_of_defence_when_both_checks_of_core_are_bypassed(self, monkeypatch):
        # core refuses a formula header while loading (1) and refuses to build a result that
        # holds a formula (2). With both switched off the bytes check still stops it (3).
        monkeypatch.setattr(input_handler, "_check_cell_for_injection", lambda *args, **kwargs: None)
        monkeypatch.setattr(output_handler, "_assert_text_only", lambda wb: None)
        hostile = real_output(header='=HYPERLINK("http://evil.example","klick")')
        with pytest.raises(AppError) as info:
            assert_no_formulas(hostile)
        assert info.value.code == "E903"

    def test_core_builds_no_result_from_a_formula_header_even_if_the_loader_check_is_bypassed(self, monkeypatch):
        monkeypatch.setattr(input_handler, "_check_cell_for_injection", lambda *args, **kwargs: None)
        with pytest.raises(output_handler.OutputValidationError):
            real_output(header='=HYPERLINK("http://evil.example","klick")')


# ── safe_name ──────────────────────────────────────────────────────────────────

class TestSafeName:

    @pytest.mark.parametrize("raw, expected", [
        ("mapped_leistungen_nrad.xlsx", "mapped_leistungen_nrad.xlsx"),
        ("Map Würzburg (neu).xlsx", "Map Würzburg (neu).xlsx"),
        ("../../etc/passwd", "etcpasswd"),
        ("C:\\Users\\x\\map.xlsx", "CUsersxmap.xlsx"),
        ('a"b<c>d.xlsx', "abcd.xlsx"),
        ("map\x00name\n.xlsx", "mapname.xlsx"),
        ("a\u202eb.xlsx", "ab.xlsx"),                          # right-to-left override
        ("Karte – Kopie.xlsx", "Karte Kopie.xlsx"),
        ("=cmd|' /C calc'!A0", "cmd C calcA0"),
        ("  ..-hidden.xlsx", "hidden.xlsx"),
        ("emoji \U0001F600.xlsx", "emoji .xlsx"),
        ("A\u0308.xlsx", "\u00c4.xlsx"),                        # decomposed Ä becomes one character
    ])
    def test_cleaning(self, raw, expected):
        assert safe_name(raw) == expected

    @pytest.mark.parametrize("raw", ["", None, "///", "\x00\x01", "...", "---"])
    def test_never_empty(self, raw):
        assert safe_name(raw) == "Datei"

    def test_long_names_are_cut_and_keep_the_extension(self):
        name = safe_name("m" * 300 + ".xlsx")
        assert len(name) == 80 and name.endswith(".xlsx")
        assert len(safe_name("m" * 300)) == 80

    def test_result_only_has_safe_characters(self):
        for raw in ['"><script>alert(1)</script>.xlsx', "a/b\\c:d*e?f|g.xlsx", "x\r\ny.xlsx"]:
            assert all(c.isalnum() or c in " .-_()" for c in safe_name(raw)), raw
