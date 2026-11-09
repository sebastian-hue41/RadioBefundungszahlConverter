"""Tests for the XML content scan and the section-name check in core/input_handler.py."""

import io
import zipfile

import pytest

from core import input_handler
from core.input_handler import (
    FileValidationError,
    load_map,
    load_reference_data,
    load_statistics,
    validate_xlsx_bytes,
    validate_xlsx_file,
)

from .helpers import (
    make_map_wb,
    make_reference_wb,
    make_stats_wb,
    make_zip_bomb_xlsx_bytes,
    wb_to_bytes,
)

SHEET = "xl/worksheets/sheet1.xml"
ROWS = [(1, 1, "01.01.2024", "MRT Schädel", "MRTS")]
CODES = [("MRTS", "MRT Schädel", {"MRT": 1})]
EMPTY_REFERENCE = {"amounts": {}, "combinations": {}, "order": []}


# ── Builders ───────────────────────────────────────────────────────────────────

def _stats() -> bytes:
    return wb_to_bytes(make_stats_wb(ROWS))


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


def _add_entry(data: bytes, name: str, content: bytes) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info.filename, zin.read(info.filename))
        zout.writestr(name, content)
    return out.getvalue()


def _with_doctype(data: bytes, dtd: str) -> bytes:
    """A copy of the workbook whose sheet starts with the given DOCTYPE."""
    return _patch_zip(data, SHEET, lambda d: d.replace(b"<worksheet", dtd.encode() + b"<worksheet", 1))


def _force_text(cell, value: str) -> None:
    """Store `value` as plain text even when it starts with '=' (openpyxl would make it a formula)."""
    cell.value = value
    cell.data_type = "s"


def _billion_laughs() -> str:
    entities = '<!ENTITY a0 "lol">'
    for i in range(1, 10):
        body = f"&a{i - 1};" * 10
        entities += f'<!ENTITY a{i} "{body}">'
    return f"<!DOCTYPE worksheet [{entities}]>"


ATTACKS = {
    "internal_entity": '<!DOCTYPE worksheet [<!ENTITY x "EINGESCHLEUST">]>',
    "billion_laughs": _billion_laughs(),
    "external_entity": '<!DOCTYPE worksheet [<!ENTITY x SYSTEM "file:///etc/passwd">]>',
    "doctype_only": "<!DOCTYPE worksheet>",
    "lowercase": '<!doctype worksheet [<!entity x "y">]>',
    "mixed_case": '<!DocType worksheet [<!EnTiTy x "y">]>',
    "attlist": "<!DOCTYPE worksheet [<!ATTLIST worksheet a CDATA #IMPLIED>]>",
}


# ── DTD and entity declarations ────────────────────────────────────────────────

class TestDtdAndEntities:

    @pytest.mark.parametrize("attack", ATTACKS)
    @pytest.mark.parametrize("build, load", [(_stats, load_statistics), (_map, load_map)])
    def test_entity_attacks_are_refused_before_any_loader_parses_the_file(self, attack, build, load):
        data = _with_doctype(build(), ATTACKS[attack])
        with pytest.raises(FileValidationError, match="XML entity"):
            validate_xlsx_bytes(data)
        with pytest.raises(FileValidationError, match="XML entity"):
            load(data)

    def test_the_message_names_the_zip_entry(self):
        with pytest.raises(FileValidationError, match=SHEET):
            validate_xlsx_bytes(_with_doctype(_stats(), ATTACKS["internal_entity"]), "stats.xlsx")

    def test_entity_text_can_no_longer_be_injected_into_a_leistung(self):
        # Before the scan the parser resolved &x; and "EINGESCHLEUST" became a Leistung name.
        data = _patch_zip(_stats(), SHEET, lambda d: d.replace(
            b"<worksheet", ATTACKS["internal_entity"].encode() + b"<worksheet", 1
        ).replace("MRT Schädel".encode(), b"&x;"))
        with pytest.raises(FileValidationError, match="XML entity"):
            load_statistics(data)

    def test_the_reference_loader_still_never_raises(self):
        assert load_reference_data(_with_doctype(_reference(), ATTACKS["internal_entity"])) == EMPTY_REFERENCE

    def test_files_read_by_path_are_covered_too(self, tmp_path):
        path = tmp_path / "hostile.xlsx"
        path.write_bytes(_with_doctype(_stats(), ATTACKS["external_entity"]))
        with pytest.raises(FileValidationError, match="XML entity"):
            validate_xlsx_file(str(path))
        with pytest.raises(FileValidationError, match="XML entity"):
            load_statistics(str(path))

    @pytest.mark.parametrize("name, content", [
        ("xl/extra.dat", ATTACKS["internal_entity"].encode()),
        ("docProps/notes.txt", b"harmless text " + ATTACKS["external_entity"].encode()),
        ("xl/extra.dat", ATTACKS["internal_entity"].encode("utf-16-le")),
        ("xl/extra.dat", ATTACKS["internal_entity"].encode("utf-16-be")),
        ("xl/extra.dat", ATTACKS["lowercase"].encode()),
    ])
    def test_a_dtd_in_an_unusual_zip_entry_is_refused(self, name, content):
        with pytest.raises(FileValidationError, match="XML entity"):
            validate_xlsx_bytes(_add_entry(_stats(), name, content))

    def test_marker_that_straddles_two_chunks_is_found(self, monkeypatch):
        monkeypatch.setattr(input_handler, "_SCAN_CHUNK", 16)
        for content in (
            b" " * 10 + b"<!DOCTYPE x>",
            (" " * 5 + "<!DOCTYPE x>").encode("utf-16-le"),
        ):
            with pytest.raises(FileValidationError, match="XML entity"):
                validate_xlsx_bytes(_add_entry(_stats(), "xl/extra.dat", content))

    def test_chunk_overlap_causes_no_false_alarm(self, monkeypatch):
        monkeypatch.setattr(input_handler, "_SCAN_CHUNK", 16)
        content = b" " * 10 + b"<!-- ok -->" + b"<![CDATA[x]]>" + b"<!-"
        validate_xlsx_bytes(_add_entry(_stats(), "xl/extra.dat", content))

    def test_a_zip_bomb_is_reported_before_anything_is_unpacked(self):
        with pytest.raises(FileValidationError, match="compression ratio"):
            validate_xlsx_bytes(make_zip_bomb_xlsx_bytes())


# ── XML encoding ───────────────────────────────────────────────────────────────

class TestXmlEncoding:

    @pytest.mark.parametrize("encoding", ["utf-16", "utf-16-le", "utf-16-be"])
    @pytest.mark.parametrize("dtd", ["", ATTACKS["internal_entity"]], ids=["no_dtd", "with_dtd"])
    def test_utf16_xml_is_refused_even_without_a_dtd(self, encoding, dtd):
        def recode(raw):
            return raw.decode("utf-8").replace("<worksheet", dtd + "<worksheet", 1).encode(encoding)

        with pytest.raises(FileValidationError, match="XML encoding"):
            validate_xlsx_bytes(_patch_zip(_stats(), SHEET, recode))

    @pytest.mark.parametrize("declaration, accepted", [
        ('<?xml version="1.0" encoding="UTF-8"?>', True),
        ("<?xml version='1.0' encoding='utf-8' standalone='yes'?>", True),
        ('<?xml version="1.0" encoding="utf8"?>', True),
        ('<?xml version="1.0"?>', True),
        ("", True),
        ('<?xml version="1.0" encoding="ISO-8859-1"?>', False),
        ('<?xml version="1.0" encoding="US-ASCII"?>', False),
        ('<?xml version="1.0" encoding="windows-1252"?>', False),
        ('<?xml version="1.0" encoding="UTF-16"?>', False),
        ('<?xml version="1.0" encoding="IBM037"?>', False),
    ])
    def test_only_utf8_may_be_declared(self, declaration, accepted):
        data = _add_entry(_stats(), "xl/extra.xml", (declaration + "<root/>").encode())
        if accepted:
            validate_xlsx_bytes(data)
        else:
            with pytest.raises(FileValidationError, match="XML encoding"):
                validate_xlsx_bytes(data)

    def test_utf8_byte_order_mark_is_fine(self):
        part = b'\xef\xbb\xbf<?xml version="1.0" encoding="UTF-8"?><root/>'
        validate_xlsx_bytes(_add_entry(_stats(), "xl/extra.xml", part))

    def test_binary_parts_that_are_not_xml_are_not_judged_by_their_bytes(self):
        validate_xlsx_bytes(_add_entry(_stats(), "xl/printerSettings/printerSettings1.bin", bytes(range(256)) * 20))


# ── Files that must keep working ───────────────────────────────────────────────

class TestCleanFiles:

    def test_valid_files_still_load(self):
        assert len(load_statistics(_stats())["leistungen"]) == 1
        assert "MRT" in load_map(_map())["sections"]
        assert load_reference_data(_reference())["amounts"] == {"MRT": 100}

    def test_files_with_the_parts_excel_adds_still_validate(self):
        data = _add_entry(
            _stats(), "userCustomization/customUI.xml",
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><customUI xmlns="x"/>',
        )
        data = _add_entry(data, "xl/_rels/extra.xml.rels", b'<?xml version="1.0"?><Relationships/>')
        validate_xlsx_bytes(data)

    def test_entries_that_cannot_be_read_mean_an_invalid_archive(self):
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_STORED) as zf:
            zf.writestr("xl/workbook.xml", b"<workbook/>")
        damaged = buffer.getvalue().replace(b"<workbook/>", b"<workbook//")   # same size, wrong CRC
        with pytest.raises(FileValidationError, match="not a valid xlsx archive"):
            validate_xlsx_bytes(damaged)

    def test_encrypted_entries_mean_an_invalid_archive(self):
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as zf:
            zf.writestr("xl/workbook.xml", b"<workbook/>")
        data = bytearray(buffer.getvalue())
        data[data.find(b"PK\x03\x04") + 6] |= 0x01             # general purpose flag: encrypted
        data[data.find(b"PK\x01\x02") + 8] |= 0x01
        with pytest.raises(FileValidationError, match="not a valid xlsx archive"):
            validate_xlsx_bytes(bytes(data))


# ── Formulas in section names ──────────────────────────────────────────────────

def _map_with_header(name: str) -> bytes:
    wb = make_map_wb(CODES)
    _force_text(wb.active.cell(row=1, column=7), name)         # G1, the "MRT" section
    return wb_to_bytes(wb)


class TestSectionNames:

    def test_formula_text_as_section_name_is_refused(self):
        with pytest.raises(FileValidationError, match="formula injection.*section name"):
            load_map(_map_with_header('=HYPERLINK("http://evil.example","klick")'))

    def test_cached_formula_value_as_section_name_is_refused(self):
        data = _patch_zip(_map(), SHEET, lambda d: d.replace(
            b'<c r="G1" t="inlineStr"><is><t>MRT</t></is></c>',
            b'<c r="G1" t="str"><f>1+1</f><v>=EVIL(1)</v></c>',
        ))
        with pytest.raises(FileValidationError, match="formula injection"):
            load_map(data)

    @pytest.mark.parametrize("name", ["=x", "+x", "-x", "@x", "|x", "  =x"])
    def test_every_formula_prefix_is_refused(self, name):
        with pytest.raises(FileValidationError, match="formula injection"):
            load_map(_map_with_header(name))

    def test_normal_section_names_pass(self):
        names = ["DL Gefäße", "Rö Skelett", "MRT gesamt", "Roe/CT/DL Abdomen, Becken, Retroperit", "US FKDS"]
        data = wb_to_bytes(make_map_wb(CODES, sections=names))
        assert load_map(data)["sections"] == names

    def test_underscore_columns_keep_working_with_and_without_the_flag(self):
        data = wb_to_bytes(make_map_wb(CODES, sections=["DL Gefäße", "MRT", "_CLIP", "_Zeit"]))
        assert load_map(data)["sections"] == ["DL Gefäße", "MRT"]
        assert load_map(data, include_underscore_columns=True)["sections"] == ["DL Gefäße", "MRT", "_CLIP", "_Zeit"]
