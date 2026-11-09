"""Test helpers for local_app: workbook bytes and in-memory zip patching."""

import io
import zipfile

from tests.helpers import make_map_wb, make_stats_wb, wb_to_bytes

SHEET = "xl/worksheets/sheet1.xml"

DEFAULT_ROWS = [(1, 1, "01.01.2024", "MRT Schädel", "MRTS")]
DEFAULT_CODES = [("MRTS", "MRT Schädel", {"MRT": 1})]


def stats_bytes(rows=None, **kwargs) -> bytes:
    """A valid statistics workbook (one MRT row unless `rows` is given)."""
    return wb_to_bytes(make_stats_wb(DEFAULT_ROWS if rows is None else rows, **kwargs))


def map_bytes(codes=None, sections=None) -> bytes:
    """A valid map workbook (one MRT code unless `codes` is given)."""
    return wb_to_bytes(make_map_wb(DEFAULT_CODES if codes is None else codes, sections=sections))


def force_text(cell, value: str) -> None:
    """Store `value` as plain text even when it starts with '=' (openpyxl would make it a formula)."""
    cell.value = value
    cell.data_type = "s"


def patch_zip(data: bytes, name: str, fn) -> bytes:
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


def add_zip_entry(data: bytes, name: str, content: bytes) -> bytes:
    """A copy of the zip with one more entry."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info.filename, zin.read(info.filename))
        zout.writestr(name, content)
    return out.getvalue()


def with_doctype(data: bytes, dtd: str) -> bytes:
    """A copy of the workbook whose sheet starts with the given DOCTYPE."""
    return patch_zip(data, SHEET, lambda d: d.replace(b"<worksheet", dtd.encode() + b"<worksheet", 1))
