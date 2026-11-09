"""
security.py — what the app adds on top of core's checks.

core/ validates every uploaded xlsx before it parses a byte of it
(core/input_handler.py): the zip container, macros, zip bombs, path traversal,
DTD / XML entity declarations, UTF-8-only XML, and formula text in cells and in
the map's section names. Its loaders refuse damaged content, and
core/output_handler.py refuses a result that would hold a formula or a link.
This module keeps only what is specific to the app:

  * scan_xlsx()          core's validation of an upload, reported with the error
                         codes of the file's slot
  * assert_no_formulas() a last look at the bytes of the finished Excel (core
                         checks the workbook object; this checks the file that
                         is delivered)
  * safe_name()          cleans names that arrive with the request

Everything works on bytes in memory; nothing here touches the file system.
"""

from __future__ import annotations

import io
import re
import unicodedata
import zipfile
import zlib

from core.input_handler import FileValidationError, validate_xlsx_bytes

from .errors import SLOT_LABELS, AppError, from_core_error

# ── Uploads ────────────────────────────────────────────────────────────────────

def scan_xlsx(data: bytes, slot: str) -> None:
    """
    Validate an uploaded xlsx exactly as core does before it parses anything, and
    report a refusal with the error codes of `slot` (E1xx statistics, E2xx map).

    Nothing is checked here that core does not check: the loaders run the same
    validation again when they read the file, so the result is the same either way.
    This only lets a caller refuse one upload before it touches the others. The
    file's slot name ends up in the technical detail where core names the file.

    Raises AppError.
    """
    if slot not in SLOT_LABELS:
        raise ValueError(f"unknown file slot: {slot!r}")
    try:
        validate_xlsx_bytes(data, SLOT_LABELS[slot])
    except FileValidationError as exc:
        raise from_core_error(exc, slot) from exc


# ── The finished Excel ─────────────────────────────────────────────────────────

# What reading a damaged zip entry raises.
_UNREADABLE = (
    zipfile.BadZipFile, zipfile.LargeZipFile, RuntimeError, NotImplementedError, EOFError, zlib.error,
)
_FORMULA_ELEMENT = re.compile(rb"<f[\s>/]")   # <f>, <f t="shared" …>; not <formula> (conditional formats)
_FORBIDDEN_PARTS = ("xl/vbaproject", "xl/externallinks/", "xl/embeddings/", "xl/activex/", "xl/ctrlprops/")


def assert_no_formulas(xlsx_bytes: bytes) -> None:
    """
    Last line of defence: the finished Excel holds text and numbers only. Raises
    E903 if it contains a formula, a link, macros, external links or embedded objects.
    Normal evaluations contain none of these.
    """
    try:
        archive = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    except zipfile.BadZipFile as exc:
        raise AppError("E903", detail="Das Ergebnis ist keine gültige Excel-Datei.") from exc
    try:
        with archive:
            for info in archive.infolist():
                name = info.filename.lower()
                if name.startswith(_FORBIDDEN_PARTS):
                    raise AppError("E903", detail=f"Unzulässiger Bestandteil: {info.filename}")
                if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                    sheet = archive.read(info)
                    if _FORMULA_ELEMENT.search(sheet):
                        raise AppError("E903", detail=f"Formel in {info.filename}")
                    if b"<hyperlink" in sheet:
                        raise AppError("E903", detail=f"Link in {info.filename}")
    except _UNREADABLE as exc:
        raise AppError("E903", detail=f"Das Ergebnis ist nicht lesbar: {type(exc).__name__}") from exc


# ── Names from the request ─────────────────────────────────────────────────────

_UNSAFE_NAME = re.compile(r"[^\w .()\-]")


def safe_name(text, max_len: int = 80) -> str:
    """
    A file name reduced to letters, digits, space, '.', '-', '_' and parentheses,
    at most `max_len` characters. Slashes, quotes, '<', control characters and
    the like are dropped, so the name is safe as plain text anywhere. Never empty.
    """
    cleaned = _UNSAFE_NAME.sub("", unicodedata.normalize("NFC", str(text or "")))
    cleaned = re.sub(r" {2,}", " ", cleaned).strip(" .-")
    if len(cleaned) > max_len:
        suffix = ".xlsx" if cleaned.lower().endswith(".xlsx") else ""
        stem = cleaned[: len(cleaned) - len(suffix)]
        cleaned = stem[: max_len - len(suffix)].rstrip(" .-") + suffix
    return cleaned or "Datei"
