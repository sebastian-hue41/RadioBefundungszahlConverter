"""The requirements files pin exact versions, and the same package has the same version in every file."""

import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
FILES = ["requirements.txt", "requirements-web.txt", "requirements-dev.txt"]
PINNED = re.compile(r"^(?P<name>[A-Za-z0-9][A-Za-z0-9_.-]*)==(?P<version>[0-9][0-9A-Za-z.]*)$")


def requirement_lines(name):
    lines = (ROOT / name).read_text(encoding="utf-8").splitlines()
    return [line.strip() for line in lines if line.strip() and not line.lstrip().startswith("#")]


@pytest.mark.parametrize("name", FILES)
def test_every_requirement_is_pinned_to_one_exact_version(name):
    lines = requirement_lines(name)
    assert lines, f"{name} lists no packages"
    for line in lines:
        assert PINNED.match(line), f"{name}: {line!r} is not of the form package==version"


def test_a_package_has_the_same_version_in_every_file():
    seen = {}
    for name in FILES:
        for line in requirement_lines(name):
            match = PINNED.match(line)
            package = re.sub(r"[-_.]+", "-", match["name"]).lower()
            seen.setdefault(package, {})[name] = match["version"]
    for package, versions in seen.items():
        assert len(set(versions.values())) == 1, f"{package} differs between files: {versions}"


def test_the_runtime_packages_are_in_the_cli_and_web_files():
    for name in ("requirements.txt", "requirements-web.txt"):
        assert {"openpyxl", "et_xmlfile"} <= {PINNED.match(line)["name"] for line in requirement_lines(name)}
