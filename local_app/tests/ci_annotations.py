"""
Turns the failures in a pytest JUnit report into GitHub annotations, so that they show on the page of a workflow run
(the full log needs a login with admin rights). Used by .github/workflows/build-local.yml; not collected by pytest.

    python local_app/tests/ci_annotations.py REPORT.xml

Prints nothing and exits 0 if there is no report or no failure. It only prints: it changes no file.
"""
import sys
import xml.etree.ElementTree as ET

LISTED = 40            # test names in the summary line
DETAILED = 10          # tests with their own annotation (GitHub shows at most 10 per step)
MESSAGE = 300          # characters of each failure message


def escape(text):
    """The escaping the workflow command syntax needs."""
    return str(text).replace("%", "%25").replace("\r", "%0D").replace("\n", "%0A")


def failures(path):
    """[(test id, failure message)] of the report; [] if the file is missing or unreadable."""
    try:
        root = ET.parse(path).getroot()
    except (OSError, ET.ParseError):
        return []
    found = []
    for case in root.iter("testcase"):
        problem = case.find("failure")
        if problem is None:
            problem = case.find("error")
        if problem is not None:
            name = f"{case.get('classname', '')}::{case.get('name', '')}".strip(":")
            found.append((name, " ".join((problem.get("message") or problem.text or "").split())[:MESSAGE]))
    return found


def annotations(found):
    if not found:
        return []
    names = " | ".join(name for name, _ in found[:LISTED]) + (f" | ... and {len(found) - LISTED} more" if len(found) > LISTED else "")
    lines = [f"::error title={escape(str(len(found)) + ' test(s) failed')}::{escape(names)}"]
    for name, message in found[:DETAILED]:
        lines.append(f"::error title={escape(name)}::{escape(message)}")
    return lines


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    for line in annotations(failures(sys.argv[1])):
        print(line)
