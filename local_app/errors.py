"""
errors.py — error and hint codes shown to the user of the local app.

core/ raises English messages and stays untouched, so this module translates
them into a German message, a what-to-do text and a fixed code. The page shows
the code; support gets the original text as technical detail.

    E1xx  statistics file         E2xx  map file (E220/E221: both maps together)
    E3xx  request and selection   E9xx  the program itself
    W1xx  statistics hints        W2xx  map hints        W3xx  reference hints
    W4xx  selection hints         W9xx  other hints

The contract tests in tests/test_errors.py read every message core/ can produce
and fail if one is not translated here, so a message changed on `main` is
noticed at the next `git merge main`.
"""

from __future__ import annotations

import re
import traceback
import zipfile
import zlib
from dataclasses import dataclass
from pathlib import PurePath
from xml.etree.ElementTree import ParseError

from openpyxl.utils.exceptions import InvalidFileException

from core.input_handler import (
    MAX_FILE_SIZE_MB,
    FileValidationError,
    MapValidationError,
    StatsValidationError,
)

# ── File slots ─────────────────────────────────────────────────────────────────
# Every uploaded file belongs to one slot; it decides E1xx (statistics) or E2xx
# (map) and which field the page marks.

STATISTICS = "statistics"
ALLRAD = "allrad"
NRAD = "nrad"
MAPS = "maps"                 # the one field of the automatic mode, for maps of either department
MAP_SLOTS = (ALLRAD, NRAD)    # the departments; an upload always names one of them per map
SLOT_LABELS = {
    STATISTICS: "Mitarbeiterstatistik",
    ALLRAD: "Allrad-Map",
    NRAD: "NRAD-Map",
    MAPS: "Leistungsmap",
}


def _kind(slot: str) -> str:
    if slot == STATISTICS:
        return "stats"
    if slot in MAP_SLOTS or slot == MAPS:
        return "map"
    raise ValueError(f"unknown file slot: {slot!r}")


# ── Error catalog ──────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class Entry:
    code: str
    title: str   # what happened
    hint: str    # what to do


CATALOG: dict[str, Entry] = {}


def _add(code: str, title: str, hint: str) -> None:
    CATALOG[code] = Entry(code, title, hint)


# Problems with the file itself: E10x for the statistics, E20x for a map.
# (n, title, hint for the statistics, hint for a map — None means the same)
_FILE_LEVEL = (
    (0, "Die Datei konnte nicht gelesen werden.",
        "Bitte den Code und die technischen Details melden.", None),
    (1, "Die Datei ist keine gültige Excel-Datei.",
        "Bitte als .xlsx ohne Passwortschutz speichern oder neu exportieren.", None),
    (2, "Die Datei ist leer.",
        "Bitte die richtige Datei wählen.", None),
    (3, f"Die Datei ist größer als {MAX_FILE_SIZE_MB} MB.",
        "Bitte die richtige Datei wählen oder einen kürzeren Zeitraum exportieren.",
        "Bitte die richtige Datei wählen."),
    (4, "Die Datei enthält Makros.",
        "Bitte als normale .xlsx ohne Makros speichern.", None),
    (5, "Die Datei wurde aus Sicherheitsgründen abgelehnt.",
        "Bitte neu exportieren. Bei Wiederholung den Code melden.", None),
)
for _n, _title, _stats_hint, _map_hint in _FILE_LEVEL:
    _add(f"E10{_n}", _title, _stats_hint)
    _add(f"E20{_n}", _title, _map_hint or _stats_hint)

# Structure of the statistics file.
_add("E110", "Das ist keine Mitarbeiterstatistik.",
     "Bitte den Export aus dem RIS wählen. Wurde versehentlich eine Map gewählt?")
_add("E111", "In der Statistik fehlt eine Pflichtspalte.",
     "Bitte den Export prüfen (der Spaltenname steht in den technischen Details).")
_add("E112", "Die Statistik enthält keine auswertbaren Leistungen.",
     "Bitte Export und Zeitraum prüfen.")
_add("E113", "Die Statistik hat zu viele Zeilen.",
     "Bitte einen kürzeren Zeitraum exportieren.")

# Structure of a map, and the two maps together.
_add("E210", "Das ist keine Leistungsmap.",
     "Bitte die richtige Map wählen. Wurde versehentlich die Statistik gewählt?")
_add("E211", "Die Map enthält keine Leistungen.",
     "Bitte die richtige Map wählen.")
_add("E212", "Die Map hat zu viele Zeilen oder Spalten.",
     "Bitte die richtige Map wählen.")
_add("E213", "Zwei Spalten der Map haben denselben Namen.",
     "Bitte in der Map jedem Leistungsbereich einen eigenen Spaltennamen geben "
     "(Name und Spalten stehen in den technischen Details).")
_add("E214", "Die Map enthält Werte, die keine ganzen Zahlen sind.",
     "In den Spalten der Leistungsbereiche sind nur leere Zellen und ganze Zahlen (0, 1, 2 …) "
     "erlaubt. Die Zellen stehen in den technischen Details.")
_add("E220", "Allrad- und NRAD-Map sind vertauscht.",
     "Bitte die beiden Dateien tauschen.")
_add("E221", "Allrad- und NRAD-Map passen nicht zusammen.",
     "Bitte zusammengehörige Versionen wählen.")
_add("E222", "Es wurden zwei Maps derselben Art gewählt.",
     "Bitte je eine Allrad-Map und eine NRAD-Map wählen.")

# Request and selection.
_add("E301", "Kein Bereich ausgewählt.",
     "Bitte „Alles“, „Nur NRAD-Codes“ oder „Nur Allrad-Codes“ wählen.")
_add("E302", "Es fehlt eine Datei.",
     "Bitte die markierte Datei wählen.")
_add("E303", "Diese Statistik enthält keine {dept}.",
     "Bitte bei „Was auswerten?“ eine andere Auswahl treffen, zum Beispiel „Alles“.")
_add("E304", "Die Übertragung war unvollständig.",
     "Bitte noch einmal versuchen.")
_add("E305", "Die Dateien sind zusammen zu groß.",
     "Bitte die richtigen Dateien wählen.")
_add("E306", "Die Datei konnte im Browser nicht gelesen werden.",
     "Sie liegt eventuell nur in der Cloud oder wurde inzwischen geändert. "
     "Bitte lokal speichern und neu wählen.")
_add("E307", "Es läuft bereits eine Auswertung.",
     "Bitte kurz warten oder das Programm neu starten.")

# The program itself.
_add("E901", "Das Ergebnis ist nicht mehr vorhanden.",
     "Bitte erneut auswerten.")
_add("E902", "Die Verbindung zum Programm ist unterbrochen.",
     "Bitte das Programm neu starten.")
_add("E903", "Die Auswertung wurde aus Sicherheitsgründen nicht erstellt.",
     "Bitte die Eingabedateien prüfen und den Code melden.")
_add("E999", "Unerwarteter Fehler.",
     "Bitte noch einmal versuchen; sonst den Code und die technischen Details melden.")

_DEFAULT_DEPT = "Leistungen für die gewählten Bereiche"
_MAX_DETAIL = 500


def _clip(text: str, limit: int = _MAX_DETAIL) -> str:
    text = str(text)
    return text if len(text) <= limit else text[: limit - 1] + "…"


class AppError(Exception):
    """
    An error with a fixed code, ready to show to the user.

    `slot` names the file field it belongs to (or None), `detail` is the
    technical text for support, `params` fill placeholders in the title
    (E303: dept="NRAD-Leistungen").
    """

    def __init__(self, code: str, slot: str | None = None, detail: str = "", **params: str):
        spec = CATALOG[code]   # an unknown code is a programming error — fail loudly
        if slot is not None and slot not in SLOT_LABELS:
            raise ValueError(f"unknown file slot: {slot!r}")
        title = spec.title
        if "{dept}" in title:
            title = title.replace("{dept}", params.get("dept", _DEFAULT_DEPT))
        self.code = code
        self.slot = slot
        self.title = title
        self.hint = spec.hint
        self.detail = _clip(detail)
        super().__init__(f"{code}: {title}")

    def to_dict(self) -> dict:
        return {
            "code": self.code,
            "title": self.title,
            "hint": self.hint,
            "detail": self.detail,
            "slot": self.slot,
            "slot_label": SLOT_LABELS.get(self.slot) if self.slot else None,
        }


# ── Translating core/ errors ───────────────────────────────────────────────────

# Which problem a file has, independent of statistics or map. The code follows
# from the slot: series 1 for the statistics, series 2 for a map.
_FILE_KEYS = {"unreadable": 0, "invalid": 1, "empty": 2, "toolarge": 3, "macros": 4, "security": 5}
_STATS_KEYS = {"not_stats": "E110", "column": "E111", "no_rows": "E112", "too_many_rows": "E113"}
_MAP_KEYS = {"not_map": "E210", "map_empty": "E211", "map_big": "E212", "map_names": "E213", "map_values": "E214"}


def _code_for(key: str, kind: str) -> str:
    if key in _FILE_KEYS:
        return f"E{1 if kind == 'stats' else 2}0{_FILE_KEYS[key]}"
    if kind == "stats" and key in _STATS_KEYS:
        return _STATS_KEYS[key]
    if kind == "map" and key in _MAP_KEYS:
        return _MAP_KEYS[key]
    return _code_for("unreadable", kind)   # a statistics problem in a map slot or vice versa


def file_error(slot: str, key: str, detail: str = "") -> AppError:
    """The AppError for problem `key` (see _FILE_KEYS and friends) of the file in `slot`."""
    return AppError(_code_for(key, _kind(slot)), slot, detail)


# Prefixes of the messages core/input_handler.py raises, in matching order. Every
# one of them starts with fixed text, so text taken from a file (cell values,
# zip entry names) inside a message cannot steer the classification.
_CORE_PREFIXES = (
    ("Zip entry contains a suspicious path", "security"),
    ("File does not appear to be a valid xlsx", "invalid"),
    ("File contains VBA macros", "macros"),
    ("File total uncompressed content", "security"),
    ("File is not a valid xlsx archive", "invalid"),
    ("Potential formula injection", "security"),
    ("Potential XML entity attack", "security"),
    ("Unsupported XML encoding", "security"),
    ("File is empty", "empty"),
    ("File is too large", "toolarge"),
    ("Unsupported file type", "invalid"),
    ("Cannot open statistics file", "invalid"),
    ("Cannot open map file", "invalid"),
    ("Cannot read statistics file", "invalid"),   # damaged content found while reading
    ("Cannot read map file", "invalid"),
    ("Column header 'Leist-Kurz' not found", "not_stats"),
    ("Required column", "column"),
    ("File exceeds maximum allowed row count", "too_many_rows"),
    ("No valid examination rows", "no_rows"),
    ("Map file appears to be empty", "map_empty"),
    ("No leistung entries", "map_empty"),
    ("Map validation failed", "not_map"),
    ("No section columns", "not_map"),
    ("Map file contains", "map_big"),
    ("Map file exceeds maximum allowed row count", "map_big"),
    ("Duplicate section name in map", "map_names"),
    ("Invalid value in map", "map_values"),
)

# Exceptions a damaged file makes the parser raise instead of a message of ours.
_INVALID_TYPES =(ParseError, zipfile.BadZipFile, zlib.error, EOFError, InvalidFileException)


def classify_core_message(text: str) -> str | None:
    """The problem key for a message raised by core/input_handler.py, or None if unknown."""
    if text.startswith("Zip entry ") and "extreme compression ratio" in text:
        return "security"
    for prefix, key in _CORE_PREFIXES:
        if text.startswith(prefix):
            return key
    return None


def from_core_error(exc: BaseException, slot: str) -> AppError:
    """
    Translate whatever the loader of `slot` raised into an AppError.

    Errors core/ raises on purpose are recognised by their message; a parser
    error from a damaged file means "not a valid Excel file"; anything else is
    "could not be read" with the original text kept as detail. Never E999: the
    file is the only thing that was being touched.
    """
    if isinstance(exc, AppError):
        return exc
    text = str(exc)
    if isinstance(exc, (FileValidationError, StatsValidationError, MapValidationError)):
        key = classify_core_message(text) or "unreadable"
        detail = text
    else:
        key = "invalid" if isinstance(exc, _INVALID_TYPES) else "unreadable"
        detail = f"{type(exc).__name__}: {text}"
    return file_error(slot, key, detail)


def from_output_error(exc: BaseException) -> AppError:
    """
    E903 for a result core/output_handler.py refused to deliver because it would have
    held a formula or a link. The message names the cell, never its content.
    """
    return AppError("E903", detail=str(exc))


def unexpected_error(exc: BaseException) -> AppError:
    """
    E999 for something nobody foresaw. The detail holds the error type and where
    it happened, never its message: that could contain file content.
    """
    where = ""
    frames = traceback.extract_tb(exc.__traceback__) if exc.__traceback__ else []
    if frames:
        last = frames[-1]
        where = f" in {PurePath(last.filename).name}:{last.lineno} ({last.name})"
    return AppError("E999", detail=f"{type(exc).__name__}{where}")


# ── Hints ──────────────────────────────────────────────────────────────────────
# Warnings that do not stop the evaluation. Each is (singular, plural); {n} is the
# count. Hints that are not counted have the same text twice.

SELECTION_NOTE_PREFIX = "AUSWAHL:"   # the note naming the departments and maps; not a hint


def _same(text: str) -> tuple[str, str]:
    return (text, text)


HINTS: dict[str, tuple[str, str]] = {
    "W101": ("1 Zeile wurde übersprungen.", "{n} Zeilen wurden übersprungen."),
    "W102": _same("Mitarbeiter oder Befunddatum fehlen; der Dateiname ist allgemein."),
    "W103": _same("Die Zählerprüfung war nicht möglich."),
    "W104": _same("Die Anzahl der Leistungen weicht vom Erwarteten ab."),
    "W105": ("Bei 1 Leistungscode sieht der Leistungsname wie eine Formel aus; in der Excel steht der Code statt des Namens.",
             "Bei {n} Leistungscodes sieht der Leistungsname wie eine Formel aus; in der Excel steht der Code statt des Namens."),
    "W201": ("1 Leistungscode fehlt in der Map (eigene Zeile in der Excel).",
             "{n} Leistungscodes fehlen in der Map (eigene Zeilen in der Excel)."),
    "W202": ("1 Leistung hat keine Zuordnung (am Ende der Excel aufgelistet).",
             "{n} Leistungen haben keine Zuordnung (am Ende der Excel aufgelistet)."),
    "W301": _same("Soll-Werte konnten nicht geladen werden; die Spalte „Benötigt“ fehlt."),
    "W302": _same("Einzelne Zeilen der Referenzdatei konnten nicht verwendet werden "
                  "(siehe Excel, „Hinweis: Probleme mit der Referenzdatei“)."),
    "W401": ("1 Leistung der anderen Abteilung wurde nicht ausgewertet.",
             "{n} Leistungen der anderen Abteilung wurden nicht ausgewertet."),
    "W900": ("1 weiterer Hinweis (siehe Excel, „Hinweise und Fehler“).",
             "{n} weitere Hinweise (siehe Excel, „Hinweise und Fehler“)."),
}
_COUNTED = {"W101", "W105", "W201", "W202", "W401", "W900"}
_UNASSIGNED_COUNT = re.compile(r"NICHT ZUGEORDNET:\s*([\d.,]+)")


def _de(n: int) -> str:
    """1234 -> '1.234'"""
    return f"{n:,}".replace(",", ".")


def hint(code: str, count: int | None = None) -> dict:
    """The hint `code` as {code, text, count}; `count` fills {n} of counted hints."""
    singular, plural = HINTS[code]
    text = singular if count is None or count == 1 else plural.format(n=_de(count))
    return {"code": code, "text": text, "count": count}


def classify_hint(text: str) -> str | None:
    """The hint code for a warning produced by core/input_handler.py or core/logic.py."""
    if text.startswith(("WARNING: 'Mitarbeiter' label", "WARNING: 'Befunddatum' label")):
        return "W102"
    if text.startswith(("WARNING: Total examination count", "WARNING: Value in L8")):
        return "W103"
    if text.startswith("WARNING: Row ") and "row skipped" in text:
        return "W101"
    if text.startswith("WARNING: The Leistung text of ") and "looks like a formula" in text:
        return "W105"
    if text.startswith("NOTE: Leistung ") and "not found in map" in text:
        return "W201"
    if text.startswith("NICHT ZUGEORDNET:"):
        return "W202"
    if text.startswith("COUNT MISMATCH"):
        return "W104"
    return None


def _hint_count(code: str, text: str) -> int:
    if code == "W202":   # one message names how many Leistungen are unassigned
        match = _UNASSIGNED_COUNT.match(text)
        if match:
            digits = re.sub(r"\D", "", match.group(1))
            if digits:
                return int(digits)
    return 1


def summarize_hints(messages) -> list[dict]:
    """
    Group the warnings of one evaluation into hints, sorted by code.

    `messages` is the `errors` list of the result. The selection note is skipped;
    a message nobody translated counts towards W900 so that nothing disappears.
    """
    counts: dict[str, int] = {}
    for message in messages:
        text = str(message)
        if text.startswith(SELECTION_NOTE_PREFIX):
            continue
        code = classify_hint(text) or "W900"
        counts[code] = counts.get(code, 0) + _hint_count(code, text)
    return [hint(code, counts[code] if code in _COUNTED else None) for code in sorted(counts)]
