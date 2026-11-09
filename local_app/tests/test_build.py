"""Static checks of what the single-file program is built from: the spec and the pinned build requirements.
That the finished program works is checked by smoke_binary.py (not collected by pytest: it needs the built program)."""

import re
from pathlib import Path

import pytest

from local_app import main as app_main
from local_app import server

APP = Path(server.__file__).resolve().parent
ROOT = APP.parent
SPEC = (APP / "radiobefund_local.spec").read_text(encoding="utf-8")
BUILD_REQUIREMENTS = (APP / "requirements-build.txt").read_text(encoding="utf-8")
PINNED = re.compile(r"^(?P<name>[A-Za-z0-9][A-Za-z0-9_.-]*)==(?P<version>[0-9][0-9A-Za-z.]*)(?:\s*;.*)?$")


def requirement_lines(text):
    return [line.strip() for line in text.splitlines() if line.strip() and not line.lstrip().startswith("#")]


def datas():
    """[(source expression, destination)] of the Analysis(datas=[...]) in the spec."""
    block = re.search(r"datas=\[(.*?)\n    \]", SPEC, re.S).group(1)
    return re.findall(r'\(os\.path\.join\(ROOT, ([^)]*)\), (os\.path\.join\([^)]*\)|"[^"]*")\)', block)


class TestSpec:

    def test_it_is_valid_python(self):
        compile(SPEC, "radiobefund_local.spec", "exec")

    def test_one_file_program_from_main_py_with_the_project_root_on_the_path(self):
        assert 'os.path.join(ROOT, "local_app", "main.py")' in SPEC and "pathex=[ROOT]" in SPEC
        assert "ROOT = os.path.dirname(SPECPATH)" in SPEC
        assert "COLLECT(" not in SPEC                                                  # one file, not a folder
        assert 'name="befundungszahlen"' in SPEC

    def test_upx_is_off_and_the_console_stays_for_now(self):
        assert "upx=False" in SPEC and "console=True" in SPEC and "upx=True" not in SPEC

    def test_only_the_page_and_the_reference_file_are_shipped_as_data(self):
        found = datas()
        assert [source for source, _ in found] == ['"local_app", "static", "index.html"', '"reference_amount.xlsx"']
        analysis = SPEC.split("a = Analysis")[1]
        assert "mapped_leistungen" not in analysis and "MA_Statistik" not in analysis
        assert analysis.count(".xlsx") == analysis.count("reference_amount.xlsx") == 1             # no other workbook, by any name

    def test_the_data_ends_up_where_the_program_looks_for_it(self):
        (page_source, page_dest), (reference_source, reference_dest) = datas()
        assert page_dest == 'os.path.join("local_app", "static")'                     # server._load_page: bundle/local_app/static/index.html
        assert (server.bundle_dir() / "local_app" / "static" / "index.html").is_file()
        assert reference_dest == '"."' and reference_source == f'"{app_main.REFERENCE_NAME}"'   # main.reference_path: bundle/reference_amount.xlsx
        assert app_main.reference_path().name == "reference_amount.xlsx"

    def test_every_file_it_names_exists(self):
        assert (ROOT / "local_app" / "main.py").is_file() and (ROOT / "local_app" / "static" / "index.html").is_file()
        assert (ROOT / app_main.REFERENCE_NAME).is_file()

    def test_the_web_worker_its_framework_and_the_tests_stay_out(self):
        excludes = re.search(r"excludes=\[(.*?)\]", SPEC).group(1)
        for module in ("web", "iii", "tests", "pytest", "tkinter"):
            assert f'"{module}"' in excludes
        assert "hiddenimports=[]" in SPEC

    def test_the_program_does_not_import_what_it_excludes(self):
        sources = "\n".join(p.read_text(encoding="utf-8") for p in (APP / "main.py", APP / "server.py", APP / "scope.py", APP / "security.py", APP / "errors.py"))
        assert not re.search(r"^\s*(from|import)\s+(web|iii|tests|pytest|tkinter)\b", sources, re.M)


class TestBuildRequirements:

    def test_everything_is_pinned_to_one_exact_version(self):
        lines = [line for line in requirement_lines(BUILD_REQUIREMENTS) if not line.startswith("-r ")]
        assert lines
        for line in lines:
            assert PINNED.match(line), line

    def test_it_includes_the_runtime_requirements_and_pyinstaller(self):
        assert "-r ../requirements.txt" in requirement_lines(BUILD_REQUIREMENTS)
        names = {PINNED.match(line)["name"] for line in requirement_lines(BUILD_REQUIREMENTS) if PINNED.match(line)}
        assert {"pyinstaller", "pyinstaller-hooks-contrib", "altgraph", "packaging", "setuptools"} <= names

    def test_the_referenced_requirements_file_exists_and_is_the_pinned_one(self):
        runtime = requirement_lines((ROOT / "requirements.txt").read_text(encoding="utf-8"))
        assert runtime and all(PINNED.match(line) for line in runtime)

    def test_a_package_that_is_also_in_the_runtime_requirements_has_the_same_version(self):
        runtime = {PINNED.match(l)["name"].lower(): PINNED.match(l)["version"] for l in requirement_lines((ROOT / "requirements.txt").read_text(encoding="utf-8"))}
        for line in requirement_lines(BUILD_REQUIREMENTS):
            match = PINNED.match(line)
            if match and match["name"].lower() in runtime:
                assert match["version"] == runtime[match["name"].lower()]

    def test_the_macos_only_package_is_marked_as_such(self):
        assert re.search(r'^macholib==[\d.]+; sys_platform == "darwin"$', BUILD_REQUIREMENTS, re.M)

    def test_the_file_is_not_hidden_by_the_ignore_rule_for_txt_files(self):
        assert "!requirements-build.txt" in (APP / ".gitignore").read_text(encoding="utf-8")


class TestTheBuildTools:

    @pytest.mark.parametrize("name", ["smoke_binary.py", "browser_pass.py"])
    def test_the_scripts_that_pytest_does_not_collect_are_valid_python(self, name):
        compile((APP / "tests" / name).read_text(encoding="utf-8"), name, "exec")

    def test_the_smoke_test_cannot_hang_a_build(self):
        text = (APP / "tests" / "smoke_binary.py").read_text(encoding="utf-8")
        assert "watchdog(240)" in text and "os._exit(2)" in text

    def test_the_smoke_test_checks_the_things_the_plan_names(self):
        text = (APP / "tests" / "smoke_binary.py").read_text(encoding="utf-8")
        for needle in ("reference_amount.xlsx is the only workbook inside", "_MEI", "the port is closed", "exits with code 0", "E101", "E105", "E205", "E220",
                       "reset works", "Allrad + NRAD", "plain ASCII", "stderr"):
            assert needle in text, needle
