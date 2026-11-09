"""
Static guard: the app's own code never writes files and never runs dynamic code.

The app only reads what the browser sends and hands the finished Excel back, so
there is no reason for it to touch the file system. This test is the tripwire
that keeps it that way (and keeps the confidential maps safe from any change the
generated code could make). It looks at local_app/*.py, not at the tests.
"""

import ast
import re
from pathlib import Path

import pytest

import local_app

APP_DIR = Path(local_app.__file__).parent

FORBIDDEN_IMPORTS = {"shutil", "tempfile", "pickle", "marshal", "subprocess"}
FORBIDDEN_NAMES = {"eval", "exec"}
FORBIDDEN_METHODS = {
    "write_text", "write_bytes", "unlink", "rename", "touch", "mkdir", "rmdir",
    "save", "truncate", "chmod", "symlink_to", "hardlink_to",
}
FORBIDDEN_OS = {
    "remove", "unlink", "rename", "replace", "mkdir", "makedirs", "rmdir", "removedirs",
    "truncate", "chmod", "chown", "link", "symlink", "mkfifo", "utime", "write", "open",
    "fdopen", "system", "popen", "fork", "startfile",
}
FORBIDDEN_OS_PREFIXES = ("spawn", "exec")
MODE = re.compile(r"[rwxabtU+]{1,4}")


def _literal(node):
    return node.value if isinstance(node, ast.Constant) and isinstance(node.value, str) else None


def _is_write_mode(node):
    value = _literal(node)
    return value is not None and MODE.fullmatch(value) is not None and any(c in value for c in "wax+")


def _mode_arg(call, position):
    if len(call.args) > position:
        return call.args[position]
    for keyword in call.keywords:
        if keyword.arg == "mode":
            return keyword.value
    return None


def find_violations(source):
    """Everything in `source` the app must not do, as readable strings."""
    found = []
    os_names = set()
    for node in ast.walk(ast.parse(source)):
        line = getattr(node, "lineno", 0)

        if isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name.split(".")[0] in FORBIDDEN_IMPORTS:
                    found.append(f"line {line}: import {alias.name}")
                if alias.name == "os":
                    os_names.add(alias.asname or "os")
        elif isinstance(node, ast.ImportFrom):
            module = (node.module or "").split(".")[0]
            if module in FORBIDDEN_IMPORTS:
                found.append(f"line {line}: from {node.module} import …")
            if node.module == "os":
                for alias in node.names:
                    if alias.name in FORBIDDEN_OS or alias.name.startswith(FORBIDDEN_OS_PREFIXES):
                        found.append(f"line {line}: from os import {alias.name}")

        elif isinstance(node, ast.Call):
            func = node.func
            if isinstance(func, ast.Name):
                if func.id in FORBIDDEN_NAMES:
                    found.append(f"line {line}: {func.id}()")
                elif func.id == "open":
                    mode = _mode_arg(node, 1)
                    if mode is not None and (_literal(mode) is None or _is_write_mode(mode)):
                        found.append(f"line {line}: open() with a write or unknown mode")
                elif func.id == "ZipFile" and _is_write_mode(_mode_arg(node, 1)):
                    found.append(f"line {line}: ZipFile in write mode")
            elif isinstance(func, ast.Attribute):
                base = func.value.id if isinstance(func.value, ast.Name) else None
                if func.attr in FORBIDDEN_METHODS:
                    found.append(f"line {line}: .{func.attr}()")
                elif base in os_names and (func.attr in FORBIDDEN_OS or func.attr.startswith(FORBIDDEN_OS_PREFIXES)):
                    found.append(f"line {line}: os.{func.attr}()")
                elif func.attr == "open":
                    # io.open(file, mode) vs. Path.open(mode); zf.open(name) has no mode at all
                    mode = _mode_arg(node, 1 if base in ("io", "builtins") else 0)
                    if mode is not None and _is_write_mode(mode):
                        found.append(f"line {line}: .open() in write mode")
                elif func.attr == "ZipFile" and _is_write_mode(_mode_arg(node, 1)):
                    found.append(f"line {line}: ZipFile in write mode")
    return found


@pytest.mark.parametrize("source", [
    'open("x", "w")', 'open("x", mode="ab")', 'open("x", "r+")', "open(name, flag)",
    "import shutil", "from shutil import copy", "import tempfile", "import pickle",
    "import marshal", "import subprocess",
    'eval("1")', 'exec("x = 1")',
    'p.write_text("x")', 'p.write_bytes(b"x")', "p.unlink()", 'p.rename("x")', "p.touch()",
    "p.mkdir()", 'wb.save("x.xlsx")',
    'import os\nos.remove("x")', 'import os as o\no.makedirs("x")', "from os import remove",
    'import os\nos.system("ls")', 'import os\nos.execv("a", [])', 'import os\nos.spawnl(1, "a")',
    'p.open("w")', 'p.open(mode="wb")', 'io.open("x", "w")',
    'zipfile.ZipFile("x.zip", "w")', 'ZipFile("x.zip", mode="a")',
])
def test_guard_flags(source):
    assert find_violations(source), source


@pytest.mark.parametrize("source", [
    'open("x")', 'open("x", "rb")', "p.read_bytes()", 'p.open("rb")', 'zf.open("xl/workbook.xml")',
    "zf.open(info)", "zipfile.ZipFile(io.BytesIO(data))", 'import os\nos.path.join("a", "b")',
    's.replace("a", "b")', "import re\nre.sub('a', 'b', s)", "import zipfile", "import traceback",
    "PurePath(x).name", 'from os import path',
])
def test_guard_allows(source):
    assert find_violations(source) == [], source


def test_the_app_code_is_clean():
    files = [p for p in APP_DIR.rglob("*.py") if "tests" not in p.relative_to(APP_DIR).parts]
    assert len(files) >= 2, "found too few files — is the scan broken?"
    problems = {
        str(p.relative_to(APP_DIR)): find_violations(p.read_text(encoding="utf-8")) for p in files
    }
    assert {name: found for name, found in problems.items() if found} == {}
