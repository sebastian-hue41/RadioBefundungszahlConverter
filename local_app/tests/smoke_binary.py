"""
Smoke test of the built single-file program (not collected by pytest).

    python local_app/tests/smoke_binary.py PATH_TO_PROGRAM

Runs the finished program, not the source: starts it with --no-browser from an empty folder and talks to it
over HTTP the way the page does. Everything is made up in memory (no map, no statistics, no patient data).

  * archive: the reference file is the only workbook inside; neither the web worker, the iii framework, the
    tests nor a map are in it (needs PyInstaller to be importable; otherwise this part is skipped and said so)
  * the page loads with its policy; all three selections evaluate and the downloads are valid Excel files;
    the bundled reference file is used; reset works
  * a corrupt file, a file with an XML entity, a map with a formula as a section name and swapped maps each
    come back with their error code; a wrong Host and a wrong token are refused
  * shutdown: answer, exit code 0, the port is closed, the folder the program unpacked itself to (_MEI...) is gone,
    the console output is plain ASCII and stderr is empty

Exit code 0 if every check passed, 1 otherwise. Needs openpyxl.
"""
import http.client
import io
import json
import os
import queue
import socket
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from urllib.parse import urlencode

import openpyxl

SECTIONS = ["DL Gefäße", "MRT gesamt", "CT"]       # "MRT gesamt" is a section of the bundled reference file
ALLRAD_CODES = [("RMRT", "MRT", {"MRT gesamt": 1}), ("RCT", "CT", {"CT": 1}), ("NMRT", "N MRT (Allrad)", {"MRT gesamt": 1})]
NRAD_CODES = [("NMRT", "N MRT", {"MRT gesamt": 2}), ("NCT", "N CT", {"CT": 1})]
ROWS = ["RMRT", "RMRT", "RCT", "NMRT", "NCT", "NCT", "NCT"]          # 3 Allrad rows, 4 N rows

results = []


def ascii_only(text):
    """Printed text is plain ASCII, so that no console code page (Windows) can stop the test."""
    return str(text).encode("ascii", "replace").decode("ascii")


def check(name, ok, detail=""):
    results.append((name, bool(ok)))
    print(ascii_only(("PASS " if ok else "FAIL ") + name + (f"   [{detail}]" if detail and not ok else "")), flush=True)


# ── Made-up workbooks ──────────────────────────────────────────────────────────

def workbook_bytes(wb):
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def statistics(codes=ROWS, name="SMOKE_T"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"], ws["B1"], ws["A2"], ws["B2"] = "Mitarbeiter:", name, "Befunddatum:", "01.01.2024-31.12.2024"
    ws.cell(row=8, column=12).value = len(codes)
    for column, title in enumerate(["IND1", "IND2", "DokDatum", "Leistung", "Leist-Kurz"], 1):
        ws.cell(row=10, column=column).value = title
    for i, code in enumerate(codes, 11):
        for column, value in enumerate([i - 10, 1, "01.01.2024", f"Text {code}", code], 1):
            ws.cell(row=i, column=column).value = value
    return workbook_bytes(wb)


def leistungsmap(codes, sections=SECTIONS, header=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=2).value = "Leist0"
    for column, section in enumerate(sections, 6):
        ws.cell(row=1, column=column).value = section
    if header:                                              # a formula as a section name, stored as plain text
        ws.cell(row=1, column=7).value = header
        ws.cell(row=1, column=7).data_type = "s"
    for i, (code, text, values) in enumerate(codes, 2):
        ws.cell(row=i, column=2).value, ws.cell(row=i, column=3).value = code, text
        for column, section in enumerate(sections, 6):
            ws.cell(row=i, column=column).value = values.get(section, 0)
    return workbook_bytes(wb)


def with_doctype(data):
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            content = zin.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                content = content.replace(b"<worksheet", b'<!DOCTYPE worksheet [<!ENTITY x "EINGESCHLEUST">]><worksheet', 1)
            zout.writestr(info.filename, content)
    return out.getvalue()


# ── The archive ────────────────────────────────────────────────────────────────

def check_archive(path):
    try:
        from PyInstaller.archive.readers import CArchiveReader
    except ImportError:
        print("SKIP archive check: PyInstaller is not importable here", flush=True)
        return
    reader = CArchiveReader(path)
    names = sorted(name.replace("\\", "/") for name in reader.toc)
    modules = sorted(reader.open_embedded_archive("PYZ.pyz").toc)
    check("archive: reference_amount.xlsx is the only workbook inside", [n for n in names if n.lower().endswith((".xlsx", ".xls", ".xlsm"))] == ["reference_amount.xlsx"], str(names))
    check("archive: the page is inside", "local_app/static/index.html" in names)
    check("archive: no map and no statistics file", not [n for n in names if "mapped_leistungen" in n or "MA_Statistik" in n])
    check("archive: neither the web worker, iii, the tests nor pytest nor tkinter", not [m for m in modules if m.split(".")[0] in ("web", "iii", "tests", "pytest", "tkinter", "_tkinter")])
    check("archive: our code and openpyxl are inside", {"core.input_handler", "core.logic", "core.output_handler", "local_app.server", "local_app.scope"} <= set(modules) and any(m.startswith("openpyxl") for m in modules))


# ── Running it ─────────────────────────────────────────────────────────────────

def unpack_folders():
    return {name for name in os.listdir(tempfile.gettempdir()) if name.startswith("_MEI")}


class Program:
    def __init__(self, path, cwd):
        self.before = unpack_folders()
        self.proc = subprocess.Popen([path, "--no-browser"], cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        self.lines, self.printed = queue.Queue(), []
        threading.Thread(target=lambda: [self.lines.put(line) for line in self.proc.stdout], daemon=True).start()
        self.url, end = None, time.time() + 40
        while self.url is None and time.time() < end and self.proc.poll() is None:
            try:
                line = self.lines.get(timeout=0.5)
            except queue.Empty:
                continue
            self.printed.append(line)
            if b"running at: " in line:
                self.url = line.split(b"running at: ")[1].strip().decode("ascii")
        host, _, path = (self.url or "").removeprefix("http://").partition("/")
        self.host, self.path = host, path.strip("/")                   # the token, without the slashes of the address
        self.unpacked = unpack_folders() - self.before

    def call(self, method, route="", body=None, query="", host=None, token=None, headers=None):
        conn = http.client.HTTPConnection(self.host, timeout=60)
        conn.putrequest(method, f"/{self.path if token is None else token}/{route}{query}", skip_host=True, skip_accept_encoding=True)
        conn.putheader("Host", host or self.host)
        for name, value in (headers or {}).items():
            conn.putheader(name, value)
        if body is not None:
            conn.putheader("Content-Length", str(len(body)))
        conn.endheaders()
        if body:
            conn.send(body)
        response = conn.getresponse()
        data = response.read()
        conn.close()
        return response.status, {k.lower(): v for k, v in response.getheaders()}, data

    def evaluate(self, files, names=None):
        order = [("statistics", "stats"), ("allrad", "allrad"), ("nrad", "nrad")]
        body = b"".join(files[slot] for slot, _ in order if slot in files)
        params = {param: len(files[slot]) for slot, param in order if slot in files}
        params.update({f"{slot}_name": name for slot, name in (names or {}).items()})
        status, headers, data = self.call("POST", "evaluate", body, "?" + urlencode(params), headers={"Origin": f"http://{self.host}"})
        return status, json.loads(data)


def totals(xlsx_bytes):
    ws = openpyxl.load_workbook(io.BytesIO(xlsx_bytes)).active
    header = [c.value for c in ws[1]]
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None or str(row[0]).startswith("Gezählte"):
            break
        out[row[0]] = row[header.index("Gesamt")]
    return header, out


def run(path):
    print(ascii_only(f"== program: {path} ({os.path.getsize(path) / 1e6:.1f} MB)"), flush=True)
    check_archive(path)
    cwd = tempfile.mkdtemp(prefix="local_app_smoke_")
    program = Program(path, cwd)
    try:
        check("starts and prints its address", bool(program.url), b"".join(program.printed).decode("ascii", "replace")[:300])
        if not program.url:
            return
        check("it unpacked itself to a temporary _MEI folder", len(program.unpacked) >= 1, str(program.unpacked))
        stats, allrad, nrad = statistics(), leistungsmap(ALLRAD_CODES), leistungsmap(NRAD_CODES)

        status, headers, page = program.call("GET")
        text = page.decode("utf-8")
        check("the page loads with its policy and nonce", status == 200 and "Befundungszahlen auswerten" in text and "nonce-" in headers.get("content-security-policy", "") and "{{" not in text)
        check("the catalog is in the page", '"E210"' in text and '"W301"' in text)

        expected = {"Allrad": ({"allrad": allrad}, "_Allrad.xlsx", {"MRT gesamt": 2, "CT": 1}),
                    "NRAD": ({"nrad": nrad}, "_NRAD.xlsx", {"MRT gesamt": 2, "CT": 3}),
                    "Allrad + NRAD": ({"allrad": allrad, "nrad": nrad}, "_Allrad_NRAD.xlsx", {"MRT gesamt": 4, "CT": 4})}
        for label, (maps, suffix, numbers) in expected.items():
            status, reply = program.evaluate({"statistics": stats, **maps}, {slot: f"{slot}.xlsx" for slot in maps})
            ok = status == 200 and reply["ok"] and reply["summary"]["selection_label"] == label
            check(f"evaluates {label}", ok, str(reply)[:200])
            if not ok:
                continue
            check(f"{label}: no hint says the reference file is missing", "W301" not in [h["code"] for h in reply["summary"]["hints"]])
            status, headers, xlsx = program.call("GET", "download")
            header, table = totals(xlsx)
            check(f"{label}: the download is a valid Excel with the right numbers and name",
                  status == 200 and xlsx[:2] == b"PK" and zipfile.ZipFile(io.BytesIO(xlsx)).testzip() is None and all(table.get(k) == v for k, v in numbers.items())
                  and suffix in headers["content-disposition"], f"{status} {table} {headers.get('content-disposition')}")
            check(f"{label}: the bundled reference file is used (Benötigt column)", "Benötigt" in header)
            preview = reply["summary"].get("preview") or {}
            check(f"{label}: the preview shown before the download is the table of the Excel",
                  preview.get("columns") == [h for h in header if h is not None] and len(preview.get("rows", [])) == len(table)
                  and {row[0]: row[preview["columns"].index("Gesamt")] for row in preview["rows"]} == table, str(preview)[:200])
        status, headers, data = program.call("POST", "reset")
        check("reset works: the result is gone", status == 200 and program.call("GET", "download")[0] == 404 and json.loads(program.call("GET", "download")[2])["error"]["code"] == "E901")

        for label, files, code, slot in (
            ("a corrupt file", {"statistics": b"this is not a workbook", "allrad": allrad}, "E101", "statistics"),
            ("a file with an XML entity", {"statistics": with_doctype(stats), "allrad": allrad}, "E105", "statistics"),
            ("a map with a formula as a section name", {"statistics": stats, "allrad": leistungsmap(ALLRAD_CODES, header='=HYPERLINK("http://evil.example","x")')}, "E205", "allrad"),
            ("swapped maps", {"statistics": stats, "allrad": nrad, "nrad": allrad}, "E220", "allrad"),
        ):
            status, reply = program.evaluate(files)
            check(f"{label} is refused with {code} for the {slot} field", status == 422 and reply["error"]["code"] == code and reply["error"]["slot"] == slot, str(reply)[:200])
        check("a failed evaluation leaves no result behind", program.call("GET", "download")[0] == 404)
        check("a wrong Host is refused", program.call("GET", host="evil.example")[0] == 403)
        check("a foreign Origin is refused", program.call("POST", "reset", headers={"Origin": "http://evil.example"})[0] == 403)
        check("a wrong token is a 404", program.call("GET", token="wrong")[0] == 404)
        check("a malformed length is refused", program.call("POST", "evaluate", headers={"Content-Length": "abc"}, body=None)[0] == 400)

        status, headers, data = program.call("POST", "shutdown")
        check("shutdown answers 200", status == 200 and json.loads(data) == {"ok": True})
        try:
            code = program.proc.wait(20)
        except subprocess.TimeoutExpired:
            code = None
        check("it exits with code 0", code == 0, str(code))
        port = int(program.host.rsplit(":", 1)[1])
        closed, deadline = False, time.time() + 15
        while not closed and time.time() < deadline:        # Windows needs a moment to refuse a connection
            try:
                socket.create_connection(("127.0.0.1", port), timeout=3).close()
            except ConnectionRefusedError:
                closed = True
            except OSError:
                pass
            if not closed:
                time.sleep(0.2)
        check("the port is closed", closed)
        deadline = time.time() + 15
        while program.unpacked & unpack_folders() and time.time() < deadline:      # Windows deletes the folder a little later
            time.sleep(0.2)
        check("the temporary _MEI folder is gone", not (program.unpacked & unpack_folders()), str(program.unpacked & unpack_folders()))
        out = b"".join(program.printed) + b"".join(list(program.lines.queue))
        check("the console output is plain ASCII", all(byte < 128 for byte in out) and b"has stopped" in out)
        check("nothing was written to stderr", program.proc.stderr.read() == b"")
        check("the working folder stayed empty", os.listdir(cwd) == [])
    finally:
        if program.proc.poll() is None:
            program.proc.kill()
        program.proc.wait(10)                                           # (the pipes are not closed by hand: a reader thread may still hold one)
        if not os.listdir(cwd):
            os.rmdir(cwd)


def watchdog(seconds):
    """A smoke test must never hang a build: after `seconds` everything is stopped and the test fails."""
    def stop():
        print(f"FAIL the smoke test did not finish within {seconds} s", flush=True)
        os._exit(2)
    timer = threading.Timer(seconds, stop)
    timer.daemon = True
    timer.start()


if __name__ == "__main__":
    if len(sys.argv) != 2 or not os.path.isfile(sys.argv[1]):
        sys.exit(__doc__)
    watchdog(240)
    run(os.path.abspath(sys.argv[1]))
    failed = [name for name, ok in results if not ok]
    print(ascii_only(f"\n{len(results) - len(failed)} of {len(results)} checks passed" + (f"; FAILED: {failed}" if failed else "")))
    sys.exit(1 if failed else 0)
