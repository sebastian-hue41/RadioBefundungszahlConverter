"""
logic.py — core processing: matches statistics examinations against the map
and aggregates counts by section and year.

Designed to be server-compatible: pure functions, no I/O, no side effects.
"""

from collections import defaultdict


def process(stats_data: dict, map_data: dict) -> dict:
    """
    Match every examination in `stats_data` against `map_data` and aggregate
    counts per section per year.

    Parameters
    ----------
    stats_data : dict
        Output of input_handler.load_statistics().
    map_data : dict
        Output of input_handler.load_map().

    Returns
    -------
    dict with keys:
        data            dict[str, dict]   section → {year: count, 'Total': count}
        years           list[int]         sorted list of years found
        errors          list[str]         all warnings/errors (load + processing)
        total_counted   int               total examination rows processed
        total_expected  int | None        expected count from L8 (may be None)
        mitarbeiter     str | None
        befunddatum     str | None
    """
    errors: list[str] = list(stats_data.get("errors", []))

    sections: list[str] = map_data["sections"]
    map_leistungen: dict[str, dict] = map_data["leistungen"]
    # total_leistungen = expected count from the IND2 group-end algorithm.
    # total_documents  = L8 value = Befunddokumente (unique IND1) — informational only.
    total_leistungen: int | None = stats_data.get("total_leistungen")
    total_documents: int | None = stats_data.get("total_documents")
    examinations: list[dict] = stats_data["leistungen"]

    # raw_counts[key][year] = count
    # key is either a section name (from map) or a display label (unmatched leistungen)
    raw_counts: dict[str, dict[int, int]] = defaultdict(lambda: defaultdict(int))

    years_seen: set[int] = set()
    unmatched: dict[str, str] = {}  # leist_kurz → display label (collected once per unique code)

    for exam in examinations:
        leist_kurz: str = exam["leist_kurz"]
        year: int = exam["year"]
        years_seen.add(year)

        if leist_kurz in map_leistungen:
            map_entry = map_leistungen[leist_kurz]
            # An examination can belong to multiple sections simultaneously.
            for section_name, val in map_entry["sections"].items():
                if val >= 1:
                    raw_counts[section_name][year] += val
        else:
            # Not found in map — use Leistungsbezeichnung as the display key,
            # fall back to the code itself when the name is absent.
            display_label = exam.get("leistung") or leist_kurz
            raw_counts[display_label][year] += 1

            if leist_kurz not in unmatched:
                unmatched[leist_kurz] = display_label
                errors.append(
                    f"NOTE: Leistung '{leist_kurz}' ('{display_label}') not found in map — "
                    "added as a separate row in the output."
                )

    years: list[int] = sorted(years_seen)
    total_counted: int = len(examinations)

    # ── Build the final ordered section table ─────────────────────────────────
    # Sections from the map come first (in map order), unmatched appended at end.
    data: dict[str, dict] = {}

    for section in sections:
        year_counts = {yr: raw_counts[section].get(yr, 0) for yr in years}
        year_counts["Total"] = sum(year_counts.values())
        data[section] = year_counts

    for leist_kurz, display_label in unmatched.items():
        year_counts = {yr: raw_counts[display_label].get(yr, 0) for yr in years}
        year_counts["Total"] = sum(year_counts.values())
        data[display_label] = year_counts

    # ── Count verification ─────────────────────────────────────────────────────
    # Compare processed rows against the IND2-derived expected total.
    # A mismatch means rows were dropped (missing LeistKurz, unparseable date, etc.).
    # total_documents (L8) counts unique IND1 (Befunddokumente / reports) — a
    # different unit — so it is NOT used here; passed through for display only.
    if total_leistungen is not None and total_counted != total_leistungen:
        skipped = total_leistungen - total_counted
        errors.append(
            f"COUNT MISMATCH: IND2 algorithm expects {total_leistungen:,} Leistungen; "
            f"{total_counted:,} successfully processed ({skipped:,} rows skipped). "
            "Skipped rows are listed as warnings above. "
            "This discrepancy is noted at the bottom of the output file."
        )

    return {
        "data": data,
        "years": years,
        "errors": errors,
        "total_counted": total_counted,
        "total_leistungen": total_leistungen,
        "total_documents": total_documents,
        "mitarbeiter": stats_data.get("mitarbeiter"),
        "befunddatum": stats_data.get("befunddatum"),
        "merged_duplicates": map_data.get("merged_duplicates", []),
    }
