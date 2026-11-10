"""
Static checks of .github/workflows/build-local.yml: what may trigger it, what the token may do, which
actions it uses (pinned to exact commits), and that its commands are the ones verified on a developer
machine. A real run on GitHub is the final proof (step 15 of the plan); until then local/dry runs.
"""

import re
from pathlib import Path

import pytest

from local_app import main as app_main

ROOT = Path(app_main.__file__).resolve().parent.parent
WORKFLOW_PATH = ROOT / ".github" / "workflows" / "build-local.yml"
TEXT = WORKFLOW_PATH.read_text(encoding="utf-8")
SPEC = (ROOT / "local_app" / "radiobefund_local.spec").read_text(encoding="utf-8")

ALLOWED_ACTIONS = {"actions/checkout", "actions/setup-python", "actions/upload-artifact"}


def code_lines():
    """The workflow without comment lines."""
    return [line for line in TEXT.splitlines() if not line.lstrip().startswith("#")]


def uses():
    return re.findall(r"^\s*(?:-\s*)?uses:\s*(\S+)(?:\s+#\s*(\S+))?", TEXT, re.M)


def load():
    yaml = pytest.importorskip("yaml")
    return yaml.safe_load(TEXT)


def steps():
    return load()["jobs"]["build"]["steps"]


class TestWhenItRuns:

    def test_only_a_push_to_the_branch_local(self):
        data = load()
        trigger = data.get("on", data.get(True))                  # YAML 1.1 reads a bare `on` as the boolean True
        assert trigger == {"push": {"branches": ["local"]}}

    def test_nothing_else_can_start_it(self):
        text = "\n".join(code_lines())
        for trigger in ("pull_request", "pull_request_target", "workflow_dispatch", "workflow_run", "schedule", "repository_dispatch", "issue_comment", "release"):
            assert not re.search(rf"^\s*{trigger}\s*:", text, re.M), trigger

    def test_a_newer_push_replaces_a_running_build_of_the_same_branch(self):
        concurrency = load()["concurrency"]
        assert concurrency["cancel-in-progress"] is True and "github.ref" in concurrency["group"]

    def test_the_file_is_where_github_requires_it(self):
        assert WORKFLOW_PATH.is_file() and WORKFLOW_PATH.parent == ROOT / ".github" / "workflows"


class TestWhatItMayDo:

    def test_the_token_can_only_read_the_repository(self):
        data = load()
        assert data["permissions"] == {"contents": "read"}
        assert "permissions" not in data["jobs"]["build"]                      # no job widens it

    def test_no_secrets_and_no_token_are_used(self):
        text = "\n".join(code_lines())
        assert "secrets." not in text and "GITHUB_TOKEN" not in text and "github.token" not in text
        assert not re.search(r"^\s*token\s*:", text, re.M)

    def test_only_githubs_own_actions_each_pinned_to_a_full_commit(self):
        found = uses()
        assert {name.split("@")[0] for name, _ in found} == ALLOWED_ACTIONS
        for reference, tag in found:
            name, _, sha = reference.partition("@")
            assert re.fullmatch(r"[0-9a-f]{40}", sha), reference
            assert re.fullmatch(r"v\d+\.\d+\.\d+", tag or ""), f"{reference} has no version comment"

    def test_no_action_comes_from_a_path_a_docker_image_or_a_branch(self):
        for reference, _ in uses():
            assert reference.startswith("actions/") and "docker://" not in reference and "./" not in reference

    def test_the_checkout_does_not_keep_the_credentials(self):
        checkout = next(step for step in steps() if step.get("uses", "").startswith("actions/checkout"))
        assert checkout["with"] == {"persist-credentials": False}

    def test_nothing_downloads_and_runs_a_script(self):
        text = "\n".join(code_lines())
        assert not re.search(r"\bcurl\b|\bwget\b|Invoke-WebRequest|\|\s*(ba)?sh\b|\biex\b", text)

    def test_every_install_is_pinned(self):
        installs = [line.strip() for line in code_lines() if "pip install" in line]
        assert installs
        for line in installs:
            tokens = line.split("pip install", 1)[1].split()
            requirement_file = tokens[tokens.index("-r") + 1]
            packages = [token for token in tokens if not token.startswith("-") and token != requirement_file]
            assert requirement_file == "local_app/requirements-build.txt"
            assert all(re.fullmatch(r"[A-Za-z][\w.-]*==\d+(\.\d+)*", package) for package in packages), packages        # exact versions only
            assert [package.split("==")[0] for package in packages] == ["pytest", "pyyaml"]

    def test_the_pytest_pin_is_the_one_of_the_projects_dev_requirements(self):
        pinned = re.search(r"pytest==([\d.]+)", TEXT).group(1)
        assert f"pytest=={pinned}" in (ROOT / "requirements-dev.txt").read_text(encoding="utf-8")

    def test_the_build_needs_no_secret_data_and_stays_inside_the_runner_temp_folder(self):
        run = "\n".join(step.get("run", "") for step in steps())
        assert "mapped_leistungen" not in TEXT and "MA_Statistik" not in TEXT
        assert '--distpath "$RUNNER_TEMP/dist"' in run and '--workpath "$RUNNER_TEMP/work"' in run


class TestWhatItBuilds:

    def test_three_systems_each_on_an_explicit_runner_image(self):
        matrix = load()["jobs"]["build"]["strategy"]["matrix"]["include"]
        assert [entry["runner"] for entry in matrix] == ["windows-2025", "macos-15", "ubuntu-24.04"]
        assert not any("latest" in entry["runner"] for entry in matrix)
        assert load()["jobs"]["build"]["strategy"]["fail-fast"] is False       # one system failing must not hide the others

    def test_every_entry_names_its_program_and_artifact(self):
        matrix = load()["jobs"]["build"]["strategy"]["matrix"]["include"]
        assert {entry["program"] for entry in matrix} == {"befundungszahlen", "befundungszahlen.exe"}
        assert len({entry["artifact"] for entry in matrix}) == 3
        for entry in matrix:
            assert set(entry) == {"label", "runner", "program", "artifact", "deselect"}
            assert entry["program"] == ("befundungszahlen.exe" if entry["runner"].startswith("windows") else "befundungszahlen")
            assert entry["artifact"].endswith(".exe" if entry["program"].endswith(".exe") else ".tar.gz")

    def test_the_program_name_is_the_one_of_the_spec(self):
        assert re.search(r'name="(\w+)"', SPEC).group(1) == "befundungszahlen"

    def test_a_skipped_test_is_named_and_nothing_is_skipped_now(self):
        for entry in load()["jobs"]["build"]["strategy"]["matrix"]["include"]:
            assert entry["deselect"] == "" or re.fullmatch(r"(--deselect \S+ ?)+", entry["deselect"]), entry["deselect"]
        assert [entry["deselect"] for entry in load()["jobs"]["build"]["strategy"]["matrix"]["include"]] == ["", "", ""]

    def test_python_3_14(self):
        setup = next(step for step in steps() if step.get("uses", "").startswith("actions/setup-python"))
        assert setup["with"] == {"python-version": "3.14"}

    def test_the_steps_in_their_order(self):
        names = [step["name"] for step in steps()]
        assert names == ["Check out the branch", "Set up Python", "Install the pinned build requirements and pytest", "Run all tests",
                         "Build the program", "Smoke test the finished program", "Package the program", "Keep the program as a build artifact"]

    def test_all_tests_run_and_a_failure_stops_the_job(self):
        run = next(step for step in steps() if step["name"] == "Run all tests")["run"].strip()
        assert run == "python -m pytest -q -p no:cacheprovider $DESELECT"          # the whole suite: tests/ and local_app/tests/
        assert "continue-on-error" not in TEXT

    def test_the_build_is_the_one_verified_on_a_developer_machine(self):
        run = " ".join(next(step for step in steps() if step["name"] == "Build the program")["run"].split())
        assert run == ('python -m PyInstaller --noconfirm --clean --distpath "$RUNNER_TEMP/dist" --workpath "$RUNNER_TEMP/work" '
                       "local_app/radiobefund_local.spec")

    def test_the_smoke_test_runs_against_the_finished_program(self):
        run = next(step for step in steps() if step["name"] == "Smoke test the finished program")["run"].strip()
        assert run == 'python local_app/tests/smoke_binary.py "$RUNNER_TEMP/dist/$PROGRAM"'

    def test_a_tar_file_keeps_the_executable_bit_on_macos_and_linux(self):
        run = next(step for step in steps() if step["name"] == "Package the program")["run"]
        assert 'tar -czf "$RUNNER_TEMP/out/$ARTIFACT" -C "$RUNNER_TEMP/dist" "$PROGRAM"' in run
        assert '*.tar.gz' in run and 'cp "$RUNNER_TEMP/dist/$PROGRAM" "$RUNNER_TEMP/out/$ARTIFACT"' in run

    def test_the_upload_takes_the_packaged_file_as_it_is_and_fails_if_it_is_missing(self):
        upload = next(step for step in steps() if step.get("uses", "").startswith("actions/upload-artifact"))
        assert upload["with"] == {"name": "${{ matrix.artifact }}", "path": "${{ runner.temp }}/out/${{ matrix.artifact }}",
                                  "archive": False, "if-no-files-found": "error", "retention-days": 30}

    def test_the_commands_run_in_bash_on_every_system(self):
        assert load()["defaults"]["run"]["shell"] == "bash"

    def test_matrix_values_reach_the_commands_through_the_environment_not_through_text_substitution(self):
        data = load()
        assert data["jobs"]["build"]["env"] == {"PROGRAM": "${{ matrix.program }}", "ARTIFACT": "${{ matrix.artifact }}", "DESELECT": "${{ matrix.deselect }}"}
        assert not [step for step in steps() if "run" in step and "${{" in step["run"]]

    def test_there_is_a_time_limit(self):
        assert load()["jobs"]["build"]["timeout-minutes"] <= 60
