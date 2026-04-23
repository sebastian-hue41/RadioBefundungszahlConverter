"""
main.py — CLI entry point for the RadioBefundungszahlConverter.

Usage:
    python main.py <statistics.xlsx> <map.xlsx> [options]

Options:
    --output DIR            Directory to save the output file (default: current directory)
    --include-underscore    Include underscore-prefixed map columns in output

The processing logic (logic.py) and loaders (input_handler.py) are kept
independent of this CLI so they can be called directly from a future API layer.
"""

import argparse
import sys

from core.input_handler import (
    FileValidationError,
    MapValidationError,
    StatsValidationError,
    load_map,
    load_statistics,
)
from core.logic import process
from core.output_handler import save_output


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="RadioBefundungszahlConverter",
        description=(
            "Generates a Facharzt-Auswertung xlsx by comparing a radiology "
            "statistics file against a leistungen map."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  python main.py statistics.xlsx map.xlsx
  python main.py statistics.xlsx map.xlsx --output ./output
  python main.py statistics.xlsx map.xlsx --include-underscore
        """,
    )
    parser.add_argument("statistics", help="Path to the statistics .xlsx file")
    parser.add_argument("map", help="Path to the map .xlsx file")
    parser.add_argument(
        "--output",
        default=".",
        metavar="DIR",
        help="Output directory for the generated file (default: current directory)",
    )
    parser.add_argument(
        "--include-underscore",
        action="store_true",
        help="Include underscore-prefixed section columns from the map (e.g. _CLIP, _TRENNER)",
    )
    return parser


def main() -> int:
    """
    Entry point. Returns an exit code:
        0 — success (output file created)
        1 — fatal error (no output file created)
    """
    parser = _build_parser()
    args = parser.parse_args()

    # ── Load statistics ────────────────────────────────────────────────────────
    print(f"[1/3] Loading statistics: {args.statistics}")
    try:
        stats_data = load_statistics(args.statistics)
    except FileValidationError as exc:
        print(f"  ERROR (file): {exc}", file=sys.stderr)
        return 1
    except StatsValidationError as exc:
        print(f"  ERROR (statistics): {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"  UNEXPECTED ERROR while loading statistics: {exc}", file=sys.stderr)
        return 1

    rows = len(stats_data["leistungen"])
    print(f"  -> {rows:,} examination rows loaded")
    if stats_data["mitarbeiter"]:
        print(f"  -> Mitarbeiter : {stats_data['mitarbeiter']}")
    if stats_data["befunddatum"]:
        print(f"  -> Befunddatum : {stats_data['befunddatum']}")
    if stats_data["total_documents"] is not None:
        print(f"  -> Befunddokumente (L8): {stats_data['total_documents']:,}")
    for warning in stats_data["errors"]:
        print(f"  ! {warning}")

    # ── Load map ───────────────────────────────────────────────────────────────
    print(f"\n[2/3] Loading map: {args.map}")
    try:
        map_data = load_map(args.map, include_underscore_columns=args.include_underscore)
    except FileValidationError as exc:
        print(f"  ERROR (file): {exc}", file=sys.stderr)
        return 1
    except MapValidationError as exc:
        print(f"  ERROR (map): {exc}", file=sys.stderr)
        if exc.recoverable:
            print(
                "  -> This error may be resolved by an admin uploading a corrected map file.",
                file=sys.stderr,
            )
        return 1
    except Exception as exc:
        print(f"  UNEXPECTED ERROR while loading map: {exc}", file=sys.stderr)
        return 1

    print(f"  -> {len(map_data['leistungen']):,} leistung entries loaded")
    print(f"  -> {len(map_data['sections'])} section columns")

    # ── Process ────────────────────────────────────────────────────────────────
    print("\n[3/3] Processing...")
    try:
        result = process(stats_data, map_data)
    except Exception as exc:
        print(f"  UNEXPECTED ERROR during processing: {exc}", file=sys.stderr)
        return 1

    if result["errors"]:
        print("  Notices:")
        for msg in result["errors"]:
            print(f"    ! {msg}")

    # ── Save output ────────────────────────────────────────────────────────────
    try:
        out_path = save_output(result, output_dir=args.output)
    except IOError as exc:
        print(f"\n  ERROR saving output: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"\n  UNEXPECTED ERROR while saving output: {exc}", file=sys.stderr)
        return 1

    print(f"\nDone. Output saved to:\n  {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
