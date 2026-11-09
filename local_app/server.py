"""
server.py — what the local app does with an upload.

evaluate() turns the uploaded files into the finished Excel; it is plain Python on
bytes, so everything it does can be tested without a web server. Order of work:

  1. every file is validated before any of them is parsed (scan_xlsx: core's checks
     with the error codes of the file's slot),
  2. each file is loaded with its own slot, so a problem is reported for the file
     that has it,
  3. the department guard, the merge of both maps (if both are chosen) and the
     narrowing of the statistics to the chosen departments (scope.py),
  4. core's process() and build_xlsx_bytes(); then a last look at the bytes of the
     result (security.assert_no_formulas).

Everything stays in memory. The only thing openpyxl writes while it builds the Excel is
a temporary file with the result sheet (counts and Leistung names, never anything
from the statistics beyond that); it deletes it right away.
"""

from __future__ import annotations

import re
import sys
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import parse_qsl

from core.input_handler import MAX_FILE_SIZE_MB, load_map, load_statistics
from core.logic import process
from core.output_handler import OutputValidationError, build_xlsx_bytes

from .errors import (
    ALLRAD,
    MAP_SLOTS,
    NRAD,
    SELECTION_NOTE_PREFIX,
    SLOT_LABELS,
    STATISTICS,
    AppError,
    from_core_error,
    from_output_error,
    hint,
    summarize_hints,
    unexpected_error,
)
from .scope import apply_scope, check_department, merge_maps
from .security import assert_no_formulas, safe_name, scan_xlsx

_SELECTION_LABELS = {ALLRAD: "Allrad", NRAD: "NRAD"}
_MAX_SHOWN = 100   # longest text from a file (staff name, dates) passed on to the page

# The upload is one request: the files back to back, their sizes in the URL.
#   ?stats=<bytes>&allrad=<bytes>&nrad=<bytes>&allrad_name=<text>&nrad_name=<text>
# The maps that are present decide which departments are evaluated.
MAX_BODY_BYTES = 3 * MAX_FILE_SIZE_MB * 1024 * 1024 + 1024 * 1024   # three files at the limit, plus a margin
_SIZE_PARAMS = {STATISTICS: "stats", ALLRAD: "allrad", NRAD: "nrad"}   # in the order of the body
_NAME_PARAMS = {ALLRAD: "allrad_name", NRAD: "nrad_name"}
_SIZE = re.compile(r"[0-9]{1,12}")                                    # ASCII digits only: no sign, no exponent
_MAX_PARAMS = 30
_MAX_RAW_NAME = 300


def bundle_dir() -> Path:
    """
    The folder with the files shipped with the app (the page, the reference file):
    the project root normally, the folder PyInstaller unpacks to when frozen.
    """
    if getattr(sys, "frozen", False):
        unpacked = getattr(sys, "_MEIPASS", None)
        if unpacked:
            return Path(unpacked)
    return Path(__file__).resolve().parent.parent


# ── Evaluation ─────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class Evaluation:
    """The result of one evaluation: the download and what the page shows about it."""
    filename: str
    data: bytes
    summary: dict


def _as_bytes(slot: str, value) -> bytes:
    if not isinstance(value, (bytes, bytearray, memoryview)):
        raise ValueError(f"the content of {slot!r} must be bytes, not {type(value).__name__}")
    return bytes(value)


def _shown(text) -> str | None:
    """A text from a file for the page: clipped; the page inserts it as plain text."""
    if text is None:
        return None
    text = str(text)
    return text if len(text) <= _MAX_SHOWN else text[: _MAX_SHOWN - 1] + "…"


def _file_name(core_name: str, selection: list[str]) -> str:
    """core's name for the Excel with the selection appended: ..._NRAD.xlsx"""
    suffix = "_" + "_".join(_SELECTION_LABELS[slot] for slot in selection)
    stem = safe_name(core_name)
    if stem.lower().endswith(".xlsx"):
        stem = stem[: -len(".xlsx")]
    return f"{stem}{suffix}.xlsx"


def _named(slot: str, names: Mapping[str, str]) -> str:
    """'NRAD-Map „datei.xlsx“' for the note on the sheet; the name is cleaned first."""
    given = names.get(slot)
    return f"{SLOT_LABELS[slot]} „{safe_name(given)}“" if given else SLOT_LABELS[slot]


def _reference_hint(reference_data: dict | None) -> list[dict]:
    """W301 when there is no usable reference (no 'Benötigt' column), W302 when only some rows were unusable."""
    if not reference_data or not reference_data.get("amounts"):
        return [hint("W301")]
    if reference_data.get("problems"):
        return [hint("W302")]
    return []


def evaluate(
    files: Mapping[str, bytes],
    names: Mapping[str, str] | None = None,
    reference_data: dict | None = None,
) -> Evaluation:
    """
    Evaluate the uploaded files.

    `files` maps a slot (STATISTICS, ALLRAD, NRAD) to the content of the file. The maps
    that are present are the departments that are evaluated: Allrad, NRAD or both.
    `names` maps a slot to the file name the user chose (only used as text in a note
    on the sheet, after safe_name). `reference_data` is core's load_reference_data
    result; without it the sheet has no "Benötigt" column (hint W301).

    Raises AppError for everything the user can fix and for what was refused for
    safety reasons; anything nobody foresaw becomes E999 with only the error type and
    place (never the message, it could hold file content). Nothing is returned for a
    failed evaluation.
    """
    unknown = set(files) - {STATISTICS, *MAP_SLOTS}
    if unknown:
        raise ValueError(f"unknown file slots: {sorted(unknown)!r}")
    contents = {slot: _as_bytes(slot, value) for slot, value in files.items() if value is not None}
    try:
        return _evaluate(contents, names or {}, reference_data)
    except AppError:
        raise
    except Exception as exc:
        raise unexpected_error(exc) from exc


def _evaluate(files: dict[str, bytes], names: Mapping[str, str], reference_data: dict | None) -> Evaluation:
    maps = [slot for slot in MAP_SLOTS if slot in files]
    if STATISTICS not in files:
        raise AppError("E302", STATISTICS, "Es wurde keine Mitarbeiterstatistik übergeben.")
    if not maps:
        raise AppError("E301")

    # 1. Validate all files before any of them is parsed.
    for slot in (STATISTICS, *maps):
        scan_xlsx(files[slot], slot)

    # 2. Load each file with its own slot, so a problem is reported for the right file.
    try:
        stats = load_statistics(files[STATISTICS])
    except Exception as exc:
        raise from_core_error(exc, STATISTICS) from exc
    loaded = {}
    for slot in maps:
        try:
            loaded[slot] = load_map(files[slot])
        except Exception as exc:
            raise from_core_error(exc, slot) from exc

    # 3. Right maps in the right fields, merged if both; the statistics narrowed to the selection.
    for slot in maps:
        check_department(loaded[slot], slot)
    map_data = merge_maps(loaded[ALLRAD], loaded[NRAD]) if len(maps) == 2 else loaded[maps[0]]
    scoped = apply_scope(stats, ALLRAD in maps, NRAD in maps)
    removed = len(stats["leistungen"]) - len(scoped["leistungen"])

    label = " + ".join(_SELECTION_LABELS[slot] for slot in maps)
    chosen = "; ".join(_named(slot, names) for slot in maps)
    scoped["errors"].append(f"{SELECTION_NOTE_PREFIX} Ausgewertet wurde: {label}. {chosen}.")
    if removed:
        scoped["errors"].append(f"{SELECTION_NOTE_PREFIX} {hint('W401', removed)['text']}")

    # 4. Evaluate, build the Excel, look at the finished bytes once more.
    result = process(scoped, map_data)
    try:
        core_name, data = build_xlsx_bytes(result, reference_data=reference_data or None)
    except OutputValidationError as exc:
        raise from_output_error(exc) from exc
    assert_no_formulas(data)

    hints = summarize_hints(result["errors"]) + _reference_hint(reference_data)
    if removed:
        hints.append(hint("W401", removed))
    filename = _file_name(core_name, maps)
    summary = {
        "filename": filename,
        "mitarbeiter": _shown(result.get("mitarbeiter")),
        "befunddatum": _shown(result.get("befunddatum")),
        "total_counted": result["total_counted"],
        "selection": list(maps),
        "selection_label": label,
        "hints": sorted(hints, key=lambda h: h["code"]),
    }
    return Evaluation(filename, data, summary)


# ── Splitting an upload ────────────────────────────────────────────────────────

def _one_value(params: list[tuple[str, str]], name: str) -> str | None:
    """The value of parameter `name`; None when it is absent; E304 when it was sent twice."""
    values = [value for key, value in params if key == name]
    if len(values) > 1:
        raise AppError("E304", detail=f"Der Parameter „{name}“ wurde mehrfach gesendet.")
    return values[0] if values else None


def _size(params: list[tuple[str, str]], name: str) -> int | None:
    text = _one_value(params, name)
    if text is None:
        return None
    if not _SIZE.fullmatch(text):
        raise AppError("E304", detail=f"Ungültige Größenangabe für „{name}“: {text[:20]!r}")
    return int(text)


def split_upload(body: bytes, query: "str | Mapping[str, str]") -> tuple[dict[str, bytes], dict[str, str]]:
    """
    Cut the body of an upload into its files, by the sizes the query declares.

    Returns (files, names) as evaluate() takes them: files maps STATISTICS and each map
    that was sent to its content, names maps a map slot to the cleaned file name (only
    if one was sent). The sizes must be whole numbers and add up to the body exactly.
    Parameters that are not ours are ignored; nothing from the query is ever used as a
    slot or a path.

    Raises AppError E305 for a body over MAX_BODY_BYTES, E302 when the statistics or
    every map is missing, E304 for sizes that are invalid, sent twice or do not add up.
    """
    if len(body) > MAX_BODY_BYTES:
        raise AppError("E305", detail=f"{len(body):,} Bytes empfangen; erlaubt sind {MAX_BODY_BYTES:,}.")
    if isinstance(query, str):
        try:
            params = parse_qsl(query.lstrip("?"), keep_blank_values=True, max_num_fields=_MAX_PARAMS)
        except ValueError as exc:
            raise AppError("E304", detail="Zu viele Parameter in der Anfrage.") from exc
    else:
        params = [(str(key), str(value)) for key, value in query.items()]

    sizes = {slot: _size(params, name) for slot, name in _SIZE_PARAMS.items()}
    if sizes[STATISTICS] is None:
        raise AppError("E302", STATISTICS, "Die Größe der Mitarbeiterstatistik fehlt in der Anfrage.")
    if sizes[ALLRAD] is None and sizes[NRAD] is None:
        raise AppError("E302", detail="Es wurde keine Map gesendet.")
    declared = sum(size for size in sizes.values() if size is not None)
    if declared != len(body):
        raise AppError("E304", detail=f"Angekündigt: {declared:,} Bytes, empfangen: {len(body):,}.")

    files, names, start = {}, {}, 0
    for slot, size in sizes.items():
        if size is None:
            continue
        files[slot] = body[start:start + size]
        start += size
        raw_name = _one_value(params, _NAME_PARAMS[slot]) if slot in _NAME_PARAMS else None
        if raw_name:
            names[slot] = safe_name(raw_name[:_MAX_RAW_NAME])
    return files, names
