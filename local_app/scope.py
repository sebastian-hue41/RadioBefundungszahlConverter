"""
scope.py — which department's Leistungen an evaluation includes.

Allrad (general radiology) and NRAD (neuroradiology) work in the same RIS, but the
NRAD codes all start with "N" and have their own map. The user chooses everything,
only the NRAD codes or only the Allrad codes; this module narrows the loaded
statistics to that selection before core's process() runs, so core/ needs no
change, and tells which department a map belongs to.

Pure functions on the dicts core/input_handler.py returns; no I/O.
"""

from __future__ import annotations

import copy

from .errors import ALLRAD, NRAD, SLOT_LABELS, AppError

NRAD_PREFIX = "N"
_SHARE_LIMIT = 0.5   # a map is an NRAD map when at least half of its codes start with "N"
_NAMES_SHOWN = 5     # section names listed in a technical detail


def is_nrad_code(code) -> bool:
    """True for a Leist-Kurz code of the neuroradiology (case-sensitive, like the RIS codes)."""
    return isinstance(code, str) and code.startswith(NRAD_PREFIX)


def _reports(rows: list[dict]) -> int | None:
    """The number of distinct reports (IND1) among `rows`; None when a row has no IND1."""
    ids = [row.get("ind1") for row in rows]
    if any(i is None for i in ids):
        return None
    return len(set(ids))


def apply_scope(stats: dict, include_allrad: bool, include_nrad: bool) -> dict:
    """
    Narrow `stats` (the result of load_statistics) to the selected departments and
    return it as a new dict; the input is not changed.

    Rows of the other department are dropped. The totals the Excel reports follow
    the rows that are left:
      * total_leistungen drops by the number of rows removed. That keeps core's
        IND2 check honest: it still flags rows that were really skipped while
        loading, and no longer flags the ones this function removed on purpose.
      * total_documents (Befundberichte) is counted again from the distinct IND1
        of the remaining rows, or becomes None when a row lacks IND1. It stays
        None when the file had no L8 value to begin with.
    With both departments selected nothing is removed and nothing changes.

    Raises AppError E301 when neither is selected and E303 when nothing is left.
    """
    include_allrad, include_nrad = bool(include_allrad), bool(include_nrad)
    if not (include_allrad or include_nrad):
        raise AppError("E301")

    rows = stats["leistungen"]
    both = include_allrad and include_nrad
    kept = list(rows) if both else [r for r in rows if is_nrad_code(r["leist_kurz"]) == include_nrad]
    if not kept:
        params = {} if both else {"dept": "NRAD-Leistungen" if include_nrad else "Allrad-Leistungen"}
        raise AppError("E303", **params)

    scoped = dict(stats)
    scoped["leistungen"] = kept
    scoped["errors"] = list(stats.get("errors", []))   # the caller adds notes to its own copy
    removed = len(rows) - len(kept)
    if removed:
        if stats.get("total_leistungen") is not None:
            scoped["total_leistungen"] = max(stats["total_leistungen"] - removed, 0)
        if stats.get("total_documents") is not None:
            scoped["total_documents"] = _reports(kept)
    return scoped


# ── Two maps ───────────────────────────────────────────────────────────────────

def department_of(map_data: dict) -> tuple[str, int, int]:
    """
    Which department a loaded map belongs to: (ALLRAD or NRAD, its N codes, all its codes).

    An NRAD map is recognised by its codes: at least half of them start with "N" (the
    real NRAD map: all 115). An Allrad map has few (the real one: 52 of 3,166), so a
    handful of N codes in an Allrad map is fine. The automatic mode of the page sorts
    the maps with this rule, and check_department refuses a map in the wrong field by it.
    """
    codes = list(map_data["leistungen"])
    n_codes = sum(1 for code in codes if is_nrad_code(code))
    share = n_codes / len(codes) if codes else 0.0
    return (NRAD if share >= _SHARE_LIMIT else ALLRAD), n_codes, len(codes)


def check_department(map_data: dict, slot: str) -> None:
    """
    Refuse a map that was put into the wrong field.

    The maps can be assigned by hand, and a swapped pair would silently give nearly
    empty results: the NRAD codes would not be found in the Allrad map and the other
    way round. The department of a map follows from its codes (department_of).

    Raises AppError E220 for `slot` when the map belongs to the other field.
    """
    if slot not in (ALLRAD, NRAD):
        raise ValueError(f"not a map slot: {slot!r}")
    department, n_codes, total = department_of(map_data)
    if department == slot:
        return
    looks_like = "NRAD" if department == NRAD else "Allrad"
    raise AppError(
        "E220", slot,
        f"{n_codes} von {total} Codes der Datei beginnen mit „{NRAD_PREFIX}“; "
        f"sie sieht wie eine {looks_like}-Map aus, steht aber im Feld „{SLOT_LABELS[slot]}“.",
    )


def _only_in(names: list[str], other: list[str]) -> str:
    known = set(other)
    missing = [name for name in names if name not in known]
    shown = ", ".join(repr(name) for name in missing[:_NAMES_SHOWN])
    return shown + (f" und {len(missing) - _NAMES_SHOWN} weitere" if len(missing) > _NAMES_SHOWN else "")


def merge_maps(allrad: dict, nrad: dict) -> dict:
    """
    The map for an evaluation of both departments, built in memory from the two
    loaded maps (core's load_map results). Neither input is changed.

    The department of a Leistung follows from its code, as in apply_scope: the NRAD
    map supplies every code that starts with "N", the Allrad map all the others.
    That keeps "Allrad + NRAD = both" exact for every section and year: an N code the
    NRAD map does not know is reported as "not in map", with one or both ticked,
    even when the Allrad map has an entry for it.

    Raises AppError E221 when the two maps do not have the same sections.
    """
    if set(allrad["sections"]) != set(nrad["sections"]):
        raise AppError(
            "E221",
            detail=(
                f"Nur in der Allrad-Map: {_only_in(allrad['sections'], nrad['sections']) or '–'}. "
                f"Nur in der NRAD-Map: {_only_in(nrad['sections'], allrad['sections']) or '–'}."
            ),
        )
    leistungen = {}
    for code, entry in allrad["leistungen"].items():
        if not is_nrad_code(code):
            leistungen[code] = copy.deepcopy(entry)
    for code, entry in nrad["leistungen"].items():
        if is_nrad_code(code):
            leistungen[code] = copy.deepcopy(entry)
    merged = [*allrad.get("merged_duplicates", []), *nrad.get("merged_duplicates", [])]
    return {
        "sections": list(allrad["sections"]),
        "leistungen": leistungen,
        "merged_duplicates": [code for code in dict.fromkeys(merged) if code in leistungen],
    }
