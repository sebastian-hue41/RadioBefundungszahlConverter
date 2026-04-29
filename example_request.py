"""
example_request.py — template for calling the RadioBefundungszahlConverter web API.

Call flow:
  1. Parse arguments — statistics file (required), reference file (optional),
     section filter (optional).
  2. Validate that the local files exist before sending anything.
  3. Build a multipart/form-data POST to POST /convert:
       statistics  — the MitarbeiterStatistik xlsx (required)
       reference   — required amounts per section (optional)
       sections    — JSON array of Leistungsbereich names to include (optional)
  4. On success (HTTP 200): read the filename from the Content-Disposition
     response header and save the xlsx to the current directory.
  5. On error: print the server error body and exit non-zero.

Usage:
    python example_request.py <statistics.xlsx> [reference.xlsx] [--sections SECTION ...]

Arguments:
    statistics      Path to the MitarbeiterStatistik xlsx file (required)
    reference       Path to the reference amounts xlsx file (optional positional)
    --sections      One or more Leistungsbereich names to include (optional).
                    Names are matched case-insensitively against the server-side map.
                    When set, only those sections appear in the output and the
                    selected names are appended to the output filename.

Environment:
    III_ENGINE_URL  Base URL of the iii engine (default: http://localhost:49134)

Output filename:
    The server sets the filename in the Content-Disposition response header.
    Without a filter:  DR_EXAMPLE2024Auswertung.xlsx
    With a filter:     DR_EXAMPLE2024Auswertung_MrtMamma_MrtProstata.xlsx
    The script saves the file to the current directory using that name.

HTTP errors:
    400  Bad request — missing field, invalid sections JSON, or unknown section name
    413  Upload too large (> 50 MB per file)
    422  File validation failed (corrupt, malicious, or structurally invalid)
    500  Unexpected server error
    503  Processing timed out on the server side
"""

import argparse
import json
import os
import re
import sys

import requests

# ── Configuration ──────────────────────────────────────────────────────────────

CONVERT_URL = os.environ.get("III_ENGINE_URL", "http://localhost:49134").rstrip("/") + "/convert"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ── Argument parsing ───────────────────────────────────────────────────────────

parser = argparse.ArgumentParser(
    description="Call the RadioBefundungszahlConverter web API and save the result.",
    formatter_class=argparse.RawDescriptionHelpFormatter,
    epilog="""
examples:
  python example_request.py statistics.xlsx
  python example_request.py statistics.xlsx reference_amount.xlsx
  python example_request.py statistics.xlsx --sections "MRT Mamma" "MRT Prostata"
  python example_request.py statistics.xlsx reference_amount.xlsx --sections "CT Abdomen"
    """,
)
parser.add_argument("statistics", help="Path to the MitarbeiterStatistik xlsx file")
parser.add_argument(
    "reference",
    nargs="?",
    default=None,
    help="Path to the reference amounts xlsx file (optional)",
)
parser.add_argument(
    "--sections",
    nargs="+",
    metavar="SECTION",
    help=(
        "One or more Leistungsbereich names to include (case-insensitive). "
        "When set, only those sections appear in the output."
    ),
)
args = parser.parse_args()


# ── Local file validation (fast-fail before sending anything) ──────────────────

if not os.path.isfile(args.statistics):
    print(f"Error: statistics file not found: {args.statistics}", file=sys.stderr)
    sys.exit(1)

if args.reference and not os.path.isfile(args.reference):
    print(f"Error: reference file not found: {args.reference}", file=sys.stderr)
    sys.exit(1)


# ── Build and send request ─────────────────────────────────────────────────────

# Each entry is a (filename_hint, file_object, mime_type) tuple.
# The server uses the MIME type to validate the upload — correct MIME avoids 400.
files = {
    "statistics": (os.path.basename(args.statistics), open(args.statistics, "rb"), XLSX_MIME),
}
if args.reference:
    files["reference"] = (os.path.basename(args.reference), open(args.reference, "rb"), XLSX_MIME)

# sections is sent as a JSON-encoded byte string, not as a file upload.
if args.sections:
    files["sections"] = (None, json.dumps(args.sections).encode("utf-8"), "application/json")

print(f"Sending {args.statistics!r} to {CONVERT_URL} ...")
if args.sections:
    print(f"  sections filter: {args.sections}")

try:
    # timeout=120 gives the server its full 60 s processing budget plus headroom.
    response = requests.post(CONVERT_URL, files=files, timeout=120)
finally:
    # Close file descriptors regardless of outcome to prevent resource leaks.
    for entry in files.values():
        if isinstance(entry, tuple) and len(entry) >= 2 and hasattr(entry[1], "close"):
            entry[1].close()


# ── Handle response ────────────────────────────────────────────────────────────

if response.status_code != 200:
    # Server returns {"error": "..."} on all non-200 responses.
    print(f"Error {response.status_code}: {response.text}", file=sys.stderr)
    sys.exit(1)

# Extract the filename from Content-Disposition; fall back to a generic name.
# Header format: attachment; filename="DR_EXAMPLE2024Auswertung_MrtMamma.xlsx"
disposition = response.headers.get("Content-Disposition", "")
match = re.search(r'filename="([^"]+)"', disposition)
filename = match.group(1) if match else "Auswertung.xlsx"

output_path = os.path.join(".", filename)
with open(output_path, "wb") as f:
    f.write(response.content)

print(f"Saved: {output_path}")
