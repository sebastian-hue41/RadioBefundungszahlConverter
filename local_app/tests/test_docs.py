"""
Tests for the two documents of the local app: README.md (short, how to start and use it, what a map must look like)
and ERROR_CODES.md (every code of the catalog). Each statement the README makes about the files is checked against what
the program really does, so the text cannot drift from the code.

Refresh ERROR_CODES.md after a change of local_app/errors.py:  python -m local_app.tests.test_docs --write
"""

import io
import re
import sys
from pathlib import Path

import openpyxl
import pytest

from core import input_handler
from local_app import main as app_main
from local_app.errors import ALLRAD, CATALOG, HINTS, NRAD, STATISTICS, AppError
from local_app.server import evaluate
from tests.helpers import make_map_wb, wb_to_bytes

from .helpers import map_bytes, stats_bytes
from .test_server import ALLRAD_CODES, MIXED, NRAD_CODES, SECTIONS, exams, make_files, totals

APP = Path(app_main.__file__).resolve().parent
ROOT = APP.parent
README = (APP / "README.md").read_text(encoding="utf-8")
ERROR_CODES_PATH = APP / "ERROR_CODES.md"


def render_error_codes() -> str:
    lines = [
        "# Error codes and hints",
        "",
        "Every message of the app has a code. If something goes wrong, the page shows what happened, what to do and the code",
        "(\"Fehlercode E210\"). Quote the code, and the \"Technische Details\" under it, when you ask for help.",
        "",
        "This file is generated from `local_app/errors.py`; a test fails if the two differ. To refresh it after a change run",
        "`python -m local_app.tests.test_docs --write`.",
        "",
        "- **E1xx**: the statistics file. **E2xx**: a map file (E220/E221: the two maps together). **E3xx**: the request and the selection.",
        "  **E9xx**: the program itself. **W…**: hints; the evaluation still worked and the Excel was made.",
        "",
        "## Errors (nothing was evaluated)",
        "",
        "| Code | What happened | What to do |",
        "|---|---|---|",
    ]
    for code, entry in CATALOG.items():
        lines.append(f"| {code} | {entry.title} | {entry.hint} |")
    lines += ["", "## Hints (the Excel was made)", "", "| Code | Text |", "|---|---|"]
    for code, (singular, plural) in HINTS.items():
        text = singular if singular == plural else f"{singular} / {plural}"
        lines.append(f"| {code} | {text.replace('{n}', 'N')} |")
    return "\n".join(lines) + "\n"


def section(title: str) -> str:
    """The part of the README below the heading that starts with `title`, up to the next heading of the same level."""
    match = re.search(rf"^## {re.escape(title)}.*?$(.*?)(?=^## |\Z)", README, re.S | re.M)
    assert match, title
    return match.group(1)


def table(evaluation) -> dict:
    return totals(evaluation)


# ── The error code file ────────────────────────────────────────────────────────

class TestErrorCodes:

    def test_it_is_exactly_what_the_catalog_says(self):
        assert ERROR_CODES_PATH.read_text(encoding="utf-8") == render_error_codes(), \
            "ERROR_CODES.md is out of date: run  python -m local_app.tests.test_docs --write"

    def test_every_code_and_every_hint_is_in_it(self):
        text = ERROR_CODES_PATH.read_text(encoding="utf-8")
        for code, entry in CATALOG.items():
            assert f"| {code} | {entry.title} | {entry.hint} |" in text
        for code in HINTS:
            assert f"| {code} |" in text

    def test_it_names_no_code_that_does_not_exist(self):
        mentioned = set(re.findall(r"^\| ([EW]\d{3}) \|", ERROR_CODES_PATH.read_text(encoding="utf-8"), re.M))
        assert mentioned == set(CATALOG) | set(HINTS)

    def test_the_readme_points_to_it(self):
        assert "[ERROR_CODES.md](ERROR_CODES.md)" in README and ERROR_CODES_PATH.is_file()


# ── The README is short and says what it must ──────────────────────────────────

class TestReadmeIsShortAndComplete:

    def test_it_is_short(self):
        assert len(README.splitlines()) <= 100 and len(README.split()) <= 900

    def test_how_to_start_and_how_to_use_come_first(self):
        headings = re.findall(r"^## (.*)$", README, re.M)
        assert headings[:2] == ["1. Start it", "2. Use it"]

    def test_it_says_a_map_is_needed_and_what_it_must_look_like(self):
        text = section("3. The files you need")
        assert "You need a map file" in text and "not part of the program" in text and "you choose it yourself" in text
        for needle in ("B1 must be `Leist0`", "F1 must be `DL Gefäße`", "Column B", "From column F on", "`1`", "`2`", "`_`", ".xlsx", "E220", "E221"):
            assert needle in text, needle

    def test_it_says_what_the_statistics_file_must_have(self):
        text = section("3. The files you need")
        for needle in ("Mitarbeiter", "Befunddatum", "IND1", "IND2", "DokDatum", "Leistung", "Leist-Kurz"):
            assert needle in text, needle

    def test_the_limits_it_names_are_the_real_ones(self):
        assert set(re.findall(r"(\d+) MB", README)) == {str(input_handler.MAX_FILE_SIZE_MB)}        # every size it names is the real limit
        assert README.count(f"{input_handler.MAX_FILE_SIZE_MB} MB") == 2                              # for the map and for the statistics
        assert f"{input_handler.MAX_DATA_ROWS:,} rows" in README

    def test_the_buttons_and_labels_it_names_are_on_the_page(self):
        page = (APP / "static" / "index.html").read_text(encoding="utf-8")
        for label in ("Mitarbeiterstatistik", "Was auswerten?", "Alles", "Nur NRAD-Codes", "Nur Allrad-Codes", "Leistungsmaps",
                      "Maps selbst zuordnen", "Auswerten", "Excel herunterladen", "Weitere Statistik hochladen", "Server beenden",
                      "Hinweise anzeigen"):
            assert f"**{label}**" in README or label in README, label
            assert label in page, label
        for bold in re.findall(r"\*\*([^*…]+)\*\*", section("2. Use it")):
            assert bold in page or bold.strip() in page, bold

    def test_the_download_names_it_gives_are_the_ones_the_workflow_makes(self):
        workflow = (ROOT / ".github" / "workflows" / "build-local.yml").read_text(encoding="utf-8")
        for name in re.findall(r"`(befundungszahlen-[\w.-]+)`", README):
            assert name in workflow, name
        assert len(re.findall(r"`(befundungszahlen-[\w.-]+)`", README)) == 3

    def test_every_command_it_shows_is_exactly_the_one_that_works(self):
        assert (ROOT / "requirements.txt").is_file() and (ROOT / "requirements-dev.txt").is_file()
        assert (APP / "requirements-build.txt").is_file() and (APP / "radiobefund_local.spec").is_file()
        assert (APP / "tests" / "smoke_binary.py").is_file() and callable(app_main.main)
        block_lines = {line.strip() for block in re.findall(r"```bash\n(.*?)```", README, re.S) for line in block.splitlines()}
        inline = set(re.findall(r"`([^`\n]+)`", README))
        for line in ("pip install -r requirements.txt", "python -m local_app.main", "pip install -r local_app/requirements-build.txt",
                     "pyinstaller --noconfirm --clean --distpath ../dist --workpath ../work local_app/radiobefund_local.spec",
                     "python local_app/tests/smoke_binary.py ../dist/befundungszahlen"):
            assert line in block_lines, line
        for command in ("pip install -r requirements-dev.txt", "python -m pytest", "tar -xzf befundungszahlen-*.tar.gz", "./befundungszahlen"):
            assert command in inline, command

    def test_the_build_command_is_the_one_of_the_workflow(self):
        workflow = (ROOT / ".github" / "workflows" / "build-local.yml").read_text(encoding="utf-8")
        for flag in ("--noconfirm", "--clean", "--distpath", "--workpath", "local_app/radiobefund_local.spec"):
            assert flag in README and flag in workflow, flag

    def test_it_says_where_this_code_lives(self):
        assert "branch `local`" in README and "never merged into `main`" in README


# ── What the README says about the files is what the program does ──────────────

def codes():
    return [("RMRT", "MRT Schädel", {"MRT": 1}), ("RCT", "CT Thorax", {"CT": 1})]


def evaluate_with(map_data, rows=None, **kwargs):
    rows = rows or exams("RMRT", "RCT")
    return evaluate({"statistics": stats_bytes(rows, total_l8=len(rows)), "allrad": map_data}, **kwargs)


class TestTheMapRules:

    def test_a_map_built_exactly_as_the_readme_describes_is_accepted(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["B1"], ws["F1"], ws["G1"], ws["H1"] = "Leist0", "DL Gefäße", "MRT", "CT"          # B1, F1, then one column per section
        for row, (code, name, values) in enumerate(codes(), 2):
            ws.cell(row=row, column=2).value, ws.cell(row=row, column=3).value = code, name   # column B: the code, column C: its name
            ws.cell(row=row, column=7).value, ws.cell(row=row, column=8).value = values.get("MRT"), values.get("CT")
        result = evaluate_with(wb_to_bytes(wb))
        assert table(result)["MRT"] == 1 and table(result)["CT"] == 1

    @pytest.mark.parametrize("b1", ["Code", "leist0", "", None])
    def test_b1_must_be_leist0(self, b1):
        with pytest.raises(AppError) as info:
            evaluate_with(wb_to_bytes(make_map_wb(codes(), b1=b1)))
        assert info.value.code == "E210" and info.value.slot == ALLRAD

    @pytest.mark.parametrize("first", ["DL Gefaesse", "MRT", ""])
    def test_f1_must_be_dl_gefaesse(self, first):
        with pytest.raises(AppError) as info:
            evaluate_with(map_bytes(codes(), sections=[first, "MRT", "CT"] if first != "MRT" else ["MRT", "CT", "X"]))
        assert info.value.code == "E210"

    def test_the_value_is_how_often_a_code_counts_and_zero_or_empty_not_at_all(self):
        map_data = map_bytes([("RA", "a", {"MRT": 1}), ("RB", "b", {"MRT": 2}), ("RC", "c", {"MRT": 0}), ("RD", "d", {})], sections=["DL Gefäße", "MRT", "CT"])
        result = evaluate_with(map_data, exams("RA", "RB", "RC", "RD"))
        assert table(result)["MRT"] == 1 + 2

    def test_a_code_may_count_in_several_sections(self):
        result = evaluate_with(map_bytes([("RX", "x", {"MRT": 1, "CT": 1})], sections=["DL Gefäße", "MRT", "CT"]), exams("RX"))
        assert table(result)["MRT"] == 1 and table(result)["CT"] == 1

    def test_columns_that_start_with_an_underscore_are_ignored(self):
        result = evaluate_with(map_bytes([("RX", "x", {"MRT": 1, "_CLIP": 1})], sections=["DL Gefäße", "MRT", "_CLIP"]), exams("RX"))
        assert "_CLIP" not in table(result) and table(result)["MRT"] == 1

    @pytest.mark.parametrize("start", ["=", "+", "-", "@", "|"])
    def test_names_and_codes_must_not_start_with_the_listed_characters(self, start):
        for sections, code in ((["DL Gefäße", f"{start}Bereich", "CT"], "RX"), (["DL Gefäße", "MRT", "CT"], f"{start}RX")):
            wb = make_map_wb([(code, "x", {"MRT": 1})], sections=sections)
            for cell in list(wb.active[1]) + list(wb.active["B"]):
                if isinstance(cell.value, str) and cell.value.startswith(start):
                    cell.data_type = "s"
            with pytest.raises(AppError) as info:
                evaluate_with(wb_to_bytes(wb), exams("RX"))
            assert info.value.code == "E205", (start, sections, code)

    def test_a_code_listed_twice_is_merged_with_the_higher_value_and_noted_in_the_excel(self):
        map_data = map_bytes([("RX", "x", {"MRT": 1}), ("RX", "x", {"MRT": 2})], sections=["DL Gefäße", "MRT", "CT"])
        result = evaluate_with(map_data, exams("RX"))
        ws = openpyxl.load_workbook(io.BytesIO(result.data)).active
        assert table(result)["MRT"] == 2
        assert any("Doppelte Leistungsbezeichnungen" in str(c.value) for row in ws.iter_rows() for c in row if c.value)

    def test_a_code_missing_from_the_map_gets_its_own_row(self):
        result = evaluate_with(map_bytes(codes()), exams("RMRT", "RUNBEKANNT"))
        assert "Text RUNBEKANNT" in table(result)

    def test_a_section_name_becomes_a_row_of_the_excel(self):
        assert "Mein Leistungsbereich" in table(evaluate_with(map_bytes([("RX", "x", {"Mein Leistungsbereich": 1})], sections=["DL Gefäße", "Mein Leistungsbereich"]), exams("RX")))


class TestTwoMapsAndTheSelection:

    def test_the_table_of_the_readme(self):
        files = make_files()                                              # 3 codes without N, 4 starting with N
        assert evaluate({k: v for k, v in files.items() if k != NRAD}).summary["total_counted"] == 3
        assert evaluate({k: v for k, v in files.items() if k != ALLRAD}).summary["total_counted"] == 4
        assert evaluate(files).summary["total_counted"] == 7
        # both: the NRAD map decides for the N codes (NMRT is worth 1 in the Allrad map and 2 in the NRAD map)
        assert table(evaluate(files))["MRT"] == 2 + 2

    def test_a_swapped_pair_is_refused_with_e220(self):
        files = make_files()
        files[ALLRAD], files[NRAD] = files[NRAD], files[ALLRAD]
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == "E220"

    def test_two_maps_need_the_same_sections_e221(self):
        files = make_files()
        files[NRAD] = map_bytes(NRAD_CODES, sections=["DL Gefäße", "MRT", "Andere"])
        with pytest.raises(AppError) as info:
            evaluate(files)
        assert info.value.code == "E221"

    def test_one_map_alone_does_not_need_the_other_sections(self):
        files = make_files(nrad=False)
        assert evaluate(files).summary["selection_label"] == "Allrad"


class TestTheStatisticsFile:

    def test_a_statistics_file_as_the_readme_describes_it_is_accepted(self):
        rows = exams("RMRT", "RCT")
        assert evaluate({"statistics": stats_bytes(rows, total_l8=2, mitarbeiter="KUNZ_A", befunddatum="01.01.2024-31.12.2024"), "allrad": map_bytes(codes())}).summary["mitarbeiter"] == "KUNZ_A"

    def test_without_the_leist_kurz_header_it_is_not_a_statistics_file(self):
        wb = openpyxl.load_workbook(io.BytesIO(stats_bytes(exams("RMRT"))))
        wb.active["E10"] = "Etwas anderes"
        with pytest.raises(AppError) as info:
            evaluate({"statistics": wb_to_bytes(wb), "allrad": map_bytes(codes())})
        assert info.value.code == "E110" and info.value.slot == STATISTICS


if __name__ == "__main__":
    if "--write" in sys.argv:
        ERROR_CODES_PATH.write_text(render_error_codes(), encoding="utf-8")
        print("written", ERROR_CODES_PATH)
    else:
        sys.exit(__doc__)
