"""
logic.py — core processing: matches statistics examinations against the map
and aggregates counts by section and year.

Designed to be server-compatible: pure functions, no I/O, no side effects.
"""

from collections import defaultdict


def process(
    stats_data: dict,
    map_data: dict,
    filter_sections: "list[str] | None" = None,
) -> dict:
    """
    Match every examination in `stats_data` against `map_data` and aggregate
    counts per section per year.

    Parameters
    ----------
    stats_data : dict
        Output of input_handler.load_statistics().
    map_data : dict
        Output of input_handler.load_map().
    filter_sections : list[str] | None
        Canonical section names to include (from validate_section_filter).
        None or [] → all sections evaluated (default behaviour).
        Non-empty  → only those sections included; errors, unmatched leistungen,
                     and footer notices are suppressed for a clean output.

    Returns
    -------
    dict with keys:
        data              dict[str, dict]   section → {year: count, 'Total': count}
        years             list[int]         sorted list of years found
        errors            list[str]         warnings/errors (empty when filtered)
        total_counted     int               total examination rows processed
        total_leistungen  int | None        IND2-algorithm total (None when filtered)
        total_documents   int | None        L8 count (None when filtered)
        mitarbeiter       str | None
        befunddatum       str | None
        filter_sections   list[str]         active filter (empty = no filter)
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
    data: dict[str, dict] = {}

    if filter_sections:
        # Filtered mode: only requested sections; no unmatched rows; no errors.
        for section in filter_sections:
            year_counts = {yr: raw_counts[section].get(yr, 0) for yr in years}
            year_counts["Total"] = sum(year_counts.values())
            data[section] = year_counts
        output_errors: list[str] = []
        output_total_leistungen = None
        output_total_documents = None
        output_merged_duplicates: list[str] = []
    else:
        # Full mode: map sections first (in map order), then unmatched appended.
        for section in sections:
            year_counts = {yr: raw_counts[section].get(yr, 0) for yr in years}
            year_counts["Total"] = sum(year_counts.values())
            data[section] = year_counts

        for leist_kurz, display_label in unmatched.items():
            year_counts = {yr: raw_counts[display_label].get(yr, 0) for yr in years}
            year_counts["Total"] = sum(year_counts.values())
            data[display_label] = year_counts

        # Count verification: compare processed rows against the IND2-derived total.
        # total_documents (L8) counts unique IND1 (Befunddokumente) — a different
        # unit — so it is NOT used here; passed through for display only.
        if total_leistungen is not None and total_counted != total_leistungen:
            skipped = total_leistungen - total_counted
            errors.append(
                f"COUNT MISMATCH: IND2 algorithm expects {total_leistungen:,} Leistungen; "
                f"{total_counted:,} successfully processed ({skipped:,} rows skipped). "
                "Skipped rows are listed as warnings above. "
                "This discrepancy is noted at the bottom of the output file."
            )

        output_errors = errors
        output_total_leistungen = total_leistungen
        output_total_documents = total_documents
        output_merged_duplicates = map_data.get("merged_duplicates", [])

    return {
        "data": data,
        "years": years,
        "errors": output_errors,
        "total_counted": total_counted,
        "total_leistungen": output_total_leistungen,
        "total_documents": output_total_documents,
        "mitarbeiter": stats_data.get("mitarbeiter"),
        "befunddatum": stats_data.get("befunddatum"),
        "merged_duplicates": output_merged_duplicates,
        "filter_sections": list(filter_sections) if filter_sections else [],
    }
