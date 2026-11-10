"""
logic.py — core processing: matches statistics examinations against the map
and aggregates counts by section and year.

Designed to be server-compatible: pure functions, no I/O, no side effects.
"""

from collections import Counter, defaultdict


def unique_row_label(base: str, code: str, taken: "set[str]") -> str:
    """
    A row label for a Leistung that is not in the map, different from every label
    in `taken` (compared case-insensitively, as the Benötigt lookup does).

    `base` is used as it is when it is free; otherwise the code is added, and a
    number if even that is taken. The label is added to `taken`.
    """
    label = base
    if label.casefold() in taken:
        label = f"{base} ({code})" if base != code else f"{code} (nicht in der Map)"
    candidate, number = label, 2
    while candidate.casefold() in taken:
        candidate = f"{label} [{number}]"
        number += 1
    taken.add(candidate.casefold())
    return candidate


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
        unassigned      list[dict]        map entries hit by the statistics but
                                          assigned to no section — each entry is
                                          {leist_kurz, kurztext, count}, sorted
                                          by count descending
        unmatched_rows  dict[str, str]    row label → Leist-Kurz for the rows of
                                          Leistungen not found in the map (they
                                          follow the section rows in `data`)

    A Leistung that is not in the map gets a row of its own, counted per code. Its
    label is its Leistung text (the code when there is none). When that text equals
    a section name or the text of another missing code, the code is added to the
    label, so its count never lands in a section row or in another code's row.
    """
    errors: list[str] = list(stats_data.get("errors", []))

    sections: list[str] = map_data["sections"]
    map_leistungen: dict[str, dict] = map_data["leistungen"]
    # total_leistungen = expected count from the IND2 group-end algorithm.
    # total_documents  = L8 value = Befunddokumente (unique IND1) — informational only.
    total_leistungen: int | None = stats_data.get("total_leistungen")
    total_documents: int | None = stats_data.get("total_documents")
    examinations: list[dict] = stats_data["leistungen"]

    # raw_counts[section][year] = count, for the sections of the map only.
    raw_counts: dict[str, dict[int, int]] = defaultdict(lambda: defaultdict(int))
    # Leistungen not found in the map, kept apart from the sections and per code:
    # leist_kurz → {year: count}, and leist_kurz → its Leistung text (the first one seen).
    unmatched_counts: dict[str, dict[int, int]] = defaultdict(lambda: defaultdict(int))
    unmatched_text: dict[str, str] = {}

    years_seen: set[int] = set()
    # leist_kurz → occurrence count for map entries with no section assignment.
    # These examinations would otherwise vanish silently: they are in the map
    # (so not reported as unmatched) but contribute to no section.
    unassigned_counts: dict[str, int] = {}

    for exam in examinations:
        leist_kurz: str = exam["leist_kurz"]
        year: int = exam["year"]
        years_seen.add(year)

        if leist_kurz in map_leistungen:
            map_entry = map_leistungen[leist_kurz]
            # An examination can belong to multiple sections simultaneously.
            assigned = False
            for section_name, val in map_entry["sections"].items():
                if val >= 1:
                    raw_counts[section_name][year] += val
                    assigned = True
            if not assigned:
                unassigned_counts[leist_kurz] = unassigned_counts.get(leist_kurz, 0) + 1
        else:
            # Not found in map — counted under its own code. The row is labelled
            # with the Leistungsbezeichnung, or the code itself when that is absent.
            unmatched_counts[leist_kurz][year] += 1
            if leist_kurz not in unmatched_text:
                unmatched_text[leist_kurz] = exam.get("leistung") or leist_kurz

    years: list[int] = sorted(years_seen)
    total_counted: int = len(examinations)

    # ── Row labels for the Leistungen not in the map ──────────────────────────
    # A label never equals a section name or another missing code's label, so no
    # count can be added to a row it does not belong to.
    taken: set[str] = {section.casefold() for section in sections}
    text_uses = Counter(text.casefold() for text in unmatched_text.values())
    unmatched_rows: dict[str, str] = {}             # row label → leist_kurz
    for leist_kurz, text in unmatched_text.items():
        if text_uses[text.casefold()] > 1:
            taken.add(text.casefold())              # shared text: every one of them gets its code
        label = unique_row_label(text, leist_kurz, taken)
        unmatched_rows[label] = leist_kurz
        if label == text:
            errors.append(
                f"NOTE: Leistung '{leist_kurz}' ('{text}') not found in map — "
                "added as a separate row in the output."
            )
        else:
            errors.append(
                f"NOTE: Leistung '{leist_kurz}' ('{text}') not found in map — "
                f"added as a separate row '{label}' in the output (its name is also "
                "the name of a section or of another missing Leistung)."
            )

    # ── Build the final ordered section table ─────────────────────────────────
    # Sections from the map come first (in map order), unmatched appended at end.
    data: dict[str, dict] = {}

    for section in sections:
        year_counts = {yr: raw_counts[section].get(yr, 0) for yr in years}
        year_counts["Total"] = sum(year_counts.values())
        data[section] = year_counts

    for label, leist_kurz in unmatched_rows.items():
        year_counts = {yr: unmatched_counts[leist_kurz].get(yr, 0) for yr in years}
        year_counts["Total"] = sum(year_counts.values())
        data[label] = year_counts

    # ── Unassigned map entries ─────────────────────────────────────────────────
    # Sorted by volume so the biggest losses surface first.
    unassigned: list[dict] = [
        {
            "leist_kurz": code,
            "kurztext": map_leistungen[code].get("kurztext", ""),
            "count": count,
        }
        for code, count in sorted(unassigned_counts.items(), key=lambda kv: -kv[1])
    ]
    if unassigned:
        total_unassigned = sum(u["count"] for u in unassigned)
        errors.append(
            f"NICHT ZUGEORDNET: {total_unassigned:,} Leistungen "
            f"({len(unassigned)} Codes) sind in der Map keinem Leistungsbereich "
            "zugeordnet und wurden in keiner Bereichszeile gezählt. "
            "Details am Ende der Ausgabedatei."
        )

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
        "unassigned": unassigned,
        "unmatched_rows": unmatched_rows,
    }
