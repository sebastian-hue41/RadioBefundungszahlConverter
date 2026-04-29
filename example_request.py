"""
example_request.py — template for calling the RadioBefundungszahlConverter web API.

Usage:
    python example_request.py <statistics.xlsx> [reference.xlsx]

Environment:
    III_ENGINE_URL   Base URL of the iii engine (default: http://localhost:49134)
"""

import os
import re
import sys

import requests

# ── Configuration ──────────────────────────────────────────────────────────────

CONVERT_URL = os.environ.get("III_ENGINE_URL", "http://localhost:49134").rstrip("/") + "/convert"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ── Argument validation ────────────────────────────────────────────────────────

if len(sys.argv) < 2:
    print("Usage: python example_request.py <statistics.xlsx> [reference.xlsx]")
    sys.exit(1)

stats_path = sys.argv[1]
ref_path   = sys.argv[2] if len(sys.argv) > 2 else None

if not os.path.isfile(stats_path):
    print(f"Error: statistics file not found: {stats_path}")
    sys.exit(1)

if ref_path and not os.path.isfile(ref_path):
    print(f"Error: reference file not found: {ref_path}")
    sys.exit(1)


# ── Build and send request ─────────────────────────────────────────────────────

files = {
    "statistics": (os.path.basename(stats_path), open(stats_path, "rb"), XLSX_MIME),
}
if ref_path:
    files["reference"] = (os.path.basename(ref_path), open(ref_path, "rb"), XLSX_MIME)

print(f"Sending {stats_path!r} to {CONVERT_URL} ...")

try:
    response = requests.post(CONVERT_URL, files=files, timeout=120)
finally:
    for _, fobj, _ in files.values():
        fobj.close()


# ── Handle response ────────────────────────────────────────────────────────────

if response.status_code != 200:
    print(f"Error {response.status_code}: {response.text}")
    sys.exit(1)

# Extract filename from Content-Disposition header, fall back to generic name.
disposition = response.headers.get("Content-Disposition", "")
match = re.search(r'filename="([^"]+)"', disposition)
filename = match.group(1) if match else "Auswertung.xlsx"

output_path = os.path.join(".", filename)
with open(output_path, "wb") as f:
    f.write(response.content)

print(f"Saved: {output_path}")
