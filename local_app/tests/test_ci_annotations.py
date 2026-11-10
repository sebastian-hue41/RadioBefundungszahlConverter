"""Tests for the script that turns pytest failures into GitHub annotations."""

import subprocess
import sys
from pathlib import Path

import pytest

from local_app.tests import ci_annotations as ci

ROOT = Path(ci.__file__).resolve().parents[2]

REPORT = """<?xml version="1.0" encoding="utf-8"?>
<testsuites><testsuite name="pytest" tests="4" failures="2" errors="1">
<testcase classname="tests.test_a.TestX" name="test_ok" time="0.1"/>
<testcase classname="tests.test_a.TestX" name="test_bad" time="0.1"><failure message="assert 1 == 2&#10;  + where 1 = f()">long text</failure></testcase>
<testcase classname="local_app.tests.test_b" name="test_broken[param]" time="0.1"><error message="fixture 'x' not found"/></testcase>
<testcase classname="local_app.tests.test_c" name="test_text" time="0.1"><failure>no message attribute, 100% sure</failure></testcase>
</testsuite></testsuites>
"""


def write(tmp_path, text=REPORT):
    path = tmp_path / "junit.xml"
    path.write_text(text, encoding="utf-8")
    return path


class TestFailures:

    def test_failures_and_errors_are_found_and_passes_are_not(self, tmp_path):
        found = ci.failures(write(tmp_path))
        assert [name for name, _ in found] == ["tests.test_a.TestX::test_bad", "local_app.tests.test_b::test_broken[param]", "local_app.tests.test_c::test_text"]

    def test_the_message_is_on_one_line_and_cut(self, tmp_path):
        assert ci.failures(write(tmp_path))[0][1] == "assert 1 == 2 + where 1 = f()"
        long = REPORT.replace("assert 1 == 2&#10;  + where 1 = f()", "x" * 1000)
        assert len(ci.failures(write(tmp_path, long))[0][1]) == ci.MESSAGE

    def test_the_text_is_used_when_there_is_no_message(self, tmp_path):
        assert ci.failures(write(tmp_path))[2][1] == "no message attribute, 100% sure"

    @pytest.mark.parametrize("content", [None, "", "not xml", "<testsuites>"])
    def test_a_missing_or_broken_report_means_no_failures(self, tmp_path, content):
        path = tmp_path / "x.xml"
        if content is not None:
            path.write_text(content)
        assert ci.failures(path) == []

    def test_a_report_without_failures(self, tmp_path):
        assert ci.failures(write(tmp_path, '<testsuites><testsuite><testcase classname="a" name="b"/></testsuite></testsuites>')) == []


class TestAnnotations:

    def test_nothing_to_report_prints_nothing(self):
        assert ci.annotations([]) == []

    def test_a_summary_and_one_annotation_per_failed_test(self, tmp_path):
        lines = ci.annotations(ci.failures(write(tmp_path)))
        assert lines[0] == "::error title=3 test(s) failed::tests.test_a.TestX::test_bad | local_app.tests.test_b::test_broken[param] | local_app.tests.test_c::test_text"
        assert lines[1] == "::error title=tests.test_a.TestX::test_bad::assert 1 == 2 + where 1 = f()"
        assert len(lines) == 4

    def test_the_characters_the_syntax_needs_are_escaped(self):
        lines = ci.annotations([("t::a", "100% sure\nnew line\r")])
        assert "100%25 sure%0Anew line%0D" in lines[1] and "\n" not in "".join(lines)

    def test_a_long_list_is_cut_and_counted(self):
        found = [(f"t::{i}", "m") for i in range(60)]
        lines = ci.annotations(found)
        assert lines[0].endswith("| ... and 20 more") and lines[0].count(" | ") == ci.LISTED
        assert len(lines) == 1 + ci.DETAILED

    def test_escape(self):
        assert ci.escape("a%b\nc\rd") == "a%25b%0Ac%0Dd"


class TestAsAScript:

    def test_it_prints_the_annotations(self, tmp_path):
        done = subprocess.run([sys.executable, "-B", str(ROOT / "local_app" / "tests" / "ci_annotations.py"), str(write(tmp_path))],
                              cwd=tmp_path, capture_output=True, text=True, timeout=30)
        assert done.returncode == 0 and done.stdout.startswith("::error title=3 test(s) failed::") and done.stderr == ""

    def test_it_prints_nothing_and_succeeds_without_a_report(self, tmp_path):
        done = subprocess.run([sys.executable, "-B", str(ROOT / "local_app" / "tests" / "ci_annotations.py"), str(tmp_path / "none.xml")],
                              cwd=tmp_path, capture_output=True, text=True, timeout=30)
        assert done.returncode == 0 and done.stdout == ""

    def test_without_an_argument_it_explains_itself(self):
        done = subprocess.run([sys.executable, "-B", str(ROOT / "local_app" / "tests" / "ci_annotations.py")], capture_output=True, text=True, timeout=30)
        assert done.returncode != 0 and "GitHub annotations" in (done.stdout + done.stderr)
