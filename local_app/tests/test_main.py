"""Tests for local_app.main: starting, the browser, stopping, and what is printed."""

import http.client
import json
import os
import queue
import socket
import subprocess
import sys
import threading
import time
from pathlib import Path

import pytest

from local_app import main as app_main
from local_app import server
from local_app.errors import ALLRAD, NRAD, STATISTICS

from .test_server import ALLRAD_CODES, NRAD_CODES, SECTIONS, make_files, post_files

ROOT = Path(app_main.__file__).resolve().parent.parent


def own_request(srv, method, route=""):
    conn = http.client.HTTPConnection("127.0.0.1", srv.port, timeout=10)
    conn.request(method, f"/{srv.token}/{route}", headers={"Host": f"127.0.0.1:{srv.port}"})
    response = conn.getresponse()
    body = response.read()
    conn.close()
    return response.status, body


def run_app(*argv, opener=None, ready=None, **kwargs):
    """main() in a thread; returns (thread, result, started) where started holds the server once it is up."""
    started, result = queue.Queue(), []

    def on_ready(srv):
        started.put(srv)
        if ready:
            ready(srv)

    def work():
        result.append(app_main.main(list(argv), opener=opener or (lambda url, new=0: True), on_ready=on_ready, **kwargs))

    thread = threading.Thread(target=work, daemon=True)
    thread.start()
    return thread, result, started


def stop(thread, result, srv):
    own_request(srv, "POST", "shutdown")
    thread.join(10)
    assert not thread.is_alive() and result == [0]


# ── The real thing: a program started the way a user starts it ─────────────────

class TestAsAProgram:

    def test_it_starts_serves_the_page_and_stops_with_code_0(self):
        env = {**os.environ, "PYTHONDONTWRITEBYTECODE": "1", "PYTHONIOENCODING": "ascii"}
        proc = subprocess.Popen([sys.executable, "-B", "-m", "local_app.main", "--no-browser"], cwd=ROOT, env=env,
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        lines = queue.Queue()
        threading.Thread(target=lambda: [lines.put(line) for line in proc.stdout], daemon=True).start()
        try:
            printed, url = [], None
            deadline = time.time() + 20
            while time.time() < deadline and url is None:
                try:
                    line = lines.get(timeout=0.5)
                except queue.Empty:
                    continue
                printed.append(line)
                if b"running at: " in line:
                    url = line.split(b"running at: ")[1].strip().decode("ascii")
            assert url, f"no address printed: {printed}"
            host, _, path = url.removeprefix("http://").partition("/")
            conn = http.client.HTTPConnection(host, timeout=10)
            conn.request("GET", "/" + path)
            page = conn.getresponse()
            assert page.status == 200 and b"Befundungszahlen" in page.read()
            assert "nonce-" in page.getheader("Content-Security-Policy")
            conn = http.client.HTTPConnection(host, timeout=10)
            conn.request("POST", "/" + path + "shutdown")
            assert conn.getresponse().status == 200
            assert proc.wait(15) == 0
            printed += list(lines.queue)
            assert all(byte < 128 for line in printed for byte in line), "the console output must be ASCII"
            assert any(b"has stopped" in line for line in printed)
            assert proc.stderr.read() == b""
        finally:
            if proc.poll() is None:
                proc.kill()
            proc.stdout.close()
            proc.stderr.close()

    def test_it_also_runs_as_a_plain_script_the_way_the_packaged_program_starts(self):
        # PyInstaller runs the file as a script, where relative imports fail; the project root is on the path via PYTHONPATH.
        env = {**os.environ, "PYTHONDONTWRITEBYTECODE": "1", "PYTHONPATH": str(ROOT)}
        helped = subprocess.run([sys.executable, "-B", str(ROOT / "local_app" / "main.py"), "--help"], cwd=ROOT.parent, env=env,
                                capture_output=True, timeout=30)
        assert helped.returncode == 0 and b"--no-browser" in helped.stdout, helped.stderr

    def test_core_log_lines_stay_off_the_console_even_without_a_reference_file(self):
        # In-process, pytest's own log handling would hide a missing silencer, so this runs a real process.
        code = ("import pathlib, sys, local_app.main as m; m.reference_path = lambda: pathlib.Path('nirgends') / 'reference_amount.xlsx'; "
                "sys.exit(m.main(['--no-browser']))")
        proc = subprocess.Popen([sys.executable, "-B", "-c", code], cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"})
        try:
            url = None
            for line in proc.stdout:
                if b"running at: " in line:
                    url = line.split(b"running at: ")[1].strip().decode("ascii")
                    break
            assert url
            host, _, path = url.removeprefix("http://").partition("/")
            conn = http.client.HTTPConnection(host, timeout=10)
            conn.request("POST", "/" + path + "shutdown")
            assert conn.getresponse().status == 200
            assert proc.wait(15) == 0
            assert proc.stderr.read() == b"", "core's log lines reached the console"
        finally:
            if proc.poll() is None:
                proc.kill()
            proc.stdout.close()
            proc.stderr.close()

    def test_help_and_wrong_arguments(self):
        env = {**os.environ, "PYTHONDONTWRITEBYTECODE": "1"}
        helped = subprocess.run([sys.executable, "-B", "-m", "local_app.main", "--help"], cwd=ROOT, env=env, capture_output=True, timeout=30)
        assert helped.returncode == 0 and b"--no-browser" in helped.stdout
        wrong = subprocess.run([sys.executable, "-B", "-m", "local_app.main", "--nope"], cwd=ROOT, env=env, capture_output=True, timeout=30)
        assert wrong.returncode == 2

    def test_the_program_needs_neither_the_web_worker_nor_its_framework(self):
        code = "import local_app.main, sys; bad = [m for m in ('web', 'iii', 'tests') if m in sys.modules]; sys.exit(1 if bad else 0)"
        assert subprocess.run([sys.executable, "-B", "-c", code], cwd=ROOT, timeout=30,
                              env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"}).returncode == 0


# ── Starting and stopping in this process ──────────────────────────────────────

class TestRun:

    def test_it_serves_until_the_shutdown_route_and_then_returns_0(self, capsys):
        thread, result, started = run_app("--no-browser")
        srv = started.get(timeout=10)
        assert own_request(srv, "GET")[0] == 200
        stop(thread, result, srv)
        with pytest.raises(OSError):
            socket.create_connection(("127.0.0.1", srv.port), timeout=0.5).close()       # the port is closed again

    def test_it_prints_the_address_and_how_to_stop(self, capsys):
        thread, result, started = run_app("--no-browser")
        srv = started.get(timeout=10)
        stop(thread, result, srv)
        out = capsys.readouterr().out
        assert f"Befundungszahlen app is running at: {srv.url}" in out
        assert "Server beenden" in out and "Ctrl+C" in out and "has stopped" in out

    def test_the_bundled_reference_file_is_used(self):
        seen = []
        thread, result, started = run_app("--no-browser", ready=lambda srv: seen.append(srv.state.reference_data))
        stop(thread, result, started.get(timeout=10))
        assert seen[0]["amounts"] and seen[0]["problems"] == []

    def test_a_page_and_evaluation_work_end_to_end(self):
        thread, result, started = run_app("--no-browser")
        srv = started.get(timeout=10)
        reply = post_files(srv, make_files())
        assert reply.status == 200 and reply.json["summary"]["selection_label"] == "Allrad + NRAD"
        assert "W301" not in [h["code"] for h in reply.json["summary"]["hints"]]          # the bundled reference is in use
        stop(thread, result, srv)

    def test_ctrl_c_stops_the_app_cleanly(self, monkeypatch, capsys):
        def interrupted(self, *args, **kwargs):
            raise KeyboardInterrupt
        monkeypatch.setattr(server.AppServer, "serve_forever", interrupted)
        servers = []                                          # keeps the server alive: only close() may free its port, not garbage collection
        assert app_main.main(["--no-browser"], on_ready=servers.append) == 0
        out = capsys.readouterr().out
        assert "Stopped (Ctrl+C)" in out and "has stopped" in out
        with pytest.raises(ConnectionRefusedError):
            socket.create_connection(("127.0.0.1", servers[0].port), timeout=0.5).close()

    def test_nothing_goes_to_stderr(self, capfd):
        thread, result, started = run_app("--no-browser")
        stop(thread, result, started.get(timeout=10))
        assert capfd.readouterr().err == ""


class TestWithoutAReferenceFile:

    @pytest.fixture
    def no_reference(self, monkeypatch, tmp_path):
        monkeypatch.setattr(app_main, "reference_path", lambda: tmp_path / "missing" / "reference_amount.xlsx")

    def test_the_app_starts_and_evaluates_and_says_what_is_missing(self, no_reference, capsys):
        thread, result, started = run_app("--no-browser")
        srv = started.get(timeout=10)
        reply = post_files(srv, make_files())
        assert reply.status == 200 and "W301" in [h["code"] for h in reply.json["summary"]["hints"]]
        stop(thread, result, srv)
        assert "Reference file: not fully usable (1 problem(s))" in capsys.readouterr().out

    def test_the_reason_is_on_the_sheet_and_no_path_is_printed(self, no_reference, capsys, tmp_path):
        thread, result, started = run_app("--no-browser")
        srv = started.get(timeout=10)
        post_files(srv, make_files())
        conn = http.client.HTTPConnection("127.0.0.1", srv.port, timeout=10)
        conn.request("GET", f"/{srv.token}/download", headers={"Host": f"127.0.0.1:{srv.port}"})
        import io
        import openpyxl
        cells = [str(c.value) for row in openpyxl.load_workbook(io.BytesIO(conn.getresponse().read())).active.iter_rows() for c in row if c.value]
        stop(thread, result, srv)
        assert "Hinweis: Probleme mit der Referenzdatei" in cells
        assert str(tmp_path) not in capsys.readouterr().out and not any(str(tmp_path) in c for c in cells)

    def test_a_reference_file_that_is_rejected_does_not_stop_the_app(self, monkeypatch, tmp_path):
        bad = tmp_path / "reference_amount.xlsx"
        bad.write_bytes(b"this is not a workbook")
        monkeypatch.setattr(app_main, "reference_path", lambda: bad)
        thread, result, started = run_app("--no-browser")
        srv = started.get(timeout=10)
        assert post_files(srv, make_files()).status == 200
        stop(thread, result, srv)


class TestCouldNotStart:

    def test_a_missing_page_is_a_clear_message_and_exit_code_1(self, monkeypatch, capsys):
        monkeypatch.setattr(server, "_load_page", lambda: (_ for _ in ()).throw(FileNotFoundError("index.html")))
        assert app_main.main(["--no-browser"]) == 1
        out = capsys.readouterr()
        assert "Could not start the app: FileNotFoundError" in out.out and "Traceback" not in out.out + out.err

    def test_a_port_that_cannot_be_bound(self, monkeypatch, capsys):
        monkeypatch.setattr(app_main, "create_server", lambda reference: (_ for _ in ()).throw(OSError("no ports")))
        assert app_main.main(["--no-browser"]) == 1
        assert "Could not start the app: OSError" in capsys.readouterr().out

    def test_a_page_without_a_nonce_placeholder(self, monkeypatch, capsys):
        monkeypatch.setattr(server, "_load_page", lambda: "<html>no placeholder</html>")
        assert app_main.main(["--no-browser"]) == 1
        assert "Could not start the app: ValueError" in capsys.readouterr().out


# ── The browser ────────────────────────────────────────────────────────────────

class TestBrowser:

    def test_it_is_given_the_address_once_in_a_thread_of_its_own(self):
        calls, threads = [], []
        done = threading.Event()

        def opener(url, new=0):
            calls.append((url, new))
            threads.append(threading.current_thread().name)
            assert threading.current_thread().daemon                       # a browser that never returns must not keep the program alive
            done.set()
            return True

        thread, result, started = run_app(opener=opener)
        srv = started.get(timeout=10)
        assert done.wait(5)
        assert calls == [(srv.url, 2)] and threads == ["open-browser"]
        stop(thread, result, srv)

    def test_it_is_not_called_with_no_browser(self):
        calls = []
        thread, result, started = run_app("--no-browser", opener=lambda *a, **k: calls.append(a))
        srv = started.get(timeout=10)
        time.sleep(0.2)
        stop(thread, result, srv)
        assert calls == []

    def test_a_browser_that_blocks_does_not_delay_the_app(self):
        release = threading.Event()
        thread, result, started = run_app(opener=lambda url, new=0: release.wait(10))
        srv = started.get(timeout=10)
        try:
            start = time.perf_counter()
            assert own_request(srv, "GET")[0] == 200 and time.perf_counter() - start < 2
        finally:
            release.set()
        stop(thread, result, srv)

    @pytest.mark.parametrize("opener", [lambda url, new=0: False, lambda url, new=0: (_ for _ in ()).throw(RuntimeError("no display"))])
    def test_a_browser_that_cannot_be_opened_is_not_fatal_and_the_user_is_told(self, opener, capsys):
        thread, result, started = run_app(opener=opener)
        srv = started.get(timeout=10)
        deadline = time.time() + 5
        out = ""
        while time.time() < deadline and "Could not open a browser" not in out:
            out += capsys.readouterr().out
            time.sleep(0.05)
        stop(thread, result, srv)
        assert "Could not open a browser. Please open the address above in one." in out

    def test_the_library_path_is_restored_before_the_browser_is_started(self, monkeypatch):
        order = []
        monkeypatch.setattr(app_main, "restore_library_path", lambda: order.append("restore"))
        done = threading.Event()
        thread, result, started = run_app(opener=lambda url, new=0: (order.append("open"), done.set())[1])
        srv = started.get(timeout=10)
        assert done.wait(5)
        stop(thread, result, srv)
        assert order == ["restore", "open"]

    def test_the_library_path_is_not_touched_with_no_browser(self, monkeypatch):
        calls = []
        monkeypatch.setattr(app_main, "restore_library_path", lambda: calls.append(1))
        thread, result, started = run_app("--no-browser")
        stop(thread, result, started.get(timeout=10))
        assert calls == []


# ── The library path of a frozen Linux build ───────────────────────────────────

class TestRestoreLibraryPath:

    def test_it_is_put_back_to_the_original(self):
        env = {"LD_LIBRARY_PATH": "/tmp/_MEI123:/usr/lib", "LD_LIBRARY_PATH_ORIG": "/opt/lib"}
        app_main.restore_library_path(env, frozen=True, platform="linux")
        assert env["LD_LIBRARY_PATH"] == "/opt/lib"

    def test_it_is_removed_when_there_was_none(self):
        env = {"LD_LIBRARY_PATH": "/tmp/_MEI123"}
        app_main.restore_library_path(env, frozen=True, platform="linux")
        assert "LD_LIBRARY_PATH" not in env

    def test_it_is_removed_when_the_original_was_empty(self):
        env = {"LD_LIBRARY_PATH": "/tmp/_MEI123", "LD_LIBRARY_PATH_ORIG": ""}
        app_main.restore_library_path(env, frozen=True, platform="linux")
        assert "LD_LIBRARY_PATH" not in env                 # an empty value would mean "the current folder" on some systems

    def test_without_an_original_the_other_variables_stay(self):
        env = {"LD_LIBRARY_PATH": "/x", "LD_LIBRARY_PATH_ORIG": "", "PATH": "/bin", "HOME": "/home/u"}
        app_main.restore_library_path(env, frozen=True, platform="linux")
        assert env == {"LD_LIBRARY_PATH_ORIG": "", "PATH": "/bin", "HOME": "/home/u"}

    def test_nothing_to_do_when_it_was_never_set(self):
        env = {"PATH": "/usr/bin"}
        app_main.restore_library_path(env, frozen=True, platform="linux")
        assert env == {"PATH": "/usr/bin"}

    @pytest.mark.parametrize("frozen, platform", [(False, "linux"), (True, "darwin"), (True, "win32"), (False, "win32"), (False, "darwin")])
    def test_it_is_left_alone_when_not_frozen_or_not_linux(self, frozen, platform):
        env = {"LD_LIBRARY_PATH": "/tmp/_MEI123", "LD_LIBRARY_PATH_ORIG": "/opt/lib"}
        before = dict(env)
        app_main.restore_library_path(env, frozen=frozen, platform=platform)
        assert env == before

    def test_other_variables_are_not_touched(self):
        env = {"LD_LIBRARY_PATH": "/x", "LD_LIBRARY_PATH_ORIG": "/y", "PATH": "/bin", "HOME": "/home/u"}
        app_main.restore_library_path(env, frozen=True, platform="linux")
        assert env["PATH"] == "/bin" and env["HOME"] == "/home/u" and env["LD_LIBRARY_PATH_ORIG"] == "/y"

    def test_the_real_environment_is_the_default_and_a_normal_run_does_not_change_it(self, monkeypatch):
        monkeypatch.delattr(sys, "frozen", raising=False)
        monkeypatch.setenv("LD_LIBRARY_PATH", "/tmp/normal")
        app_main.restore_library_path()
        assert os.environ["LD_LIBRARY_PATH"] == "/tmp/normal"

    def test_the_platform_is_read_from_sys(self, monkeypatch):
        monkeypatch.setattr(sys, "frozen", True, raising=False)
        monkeypatch.setattr(sys, "platform", "darwin")
        monkeypatch.setenv("LD_LIBRARY_PATH", "/tmp/_MEI9")
        app_main.restore_library_path()
        assert os.environ["LD_LIBRARY_PATH"] == "/tmp/_MEI9"

    def test_frozen_is_read_from_sys_and_the_real_environment_is_changed(self, monkeypatch):
        monkeypatch.setattr(sys, "frozen", True, raising=False)
        monkeypatch.setattr(sys, "platform", "linux")
        monkeypatch.setenv("LD_LIBRARY_PATH", "/tmp/_MEI9")
        monkeypatch.setenv("LD_LIBRARY_PATH_ORIG", "/orig")
        app_main.restore_library_path()
        assert os.environ["LD_LIBRARY_PATH"] == "/orig"


# ── What is printed ────────────────────────────────────────────────────────────

class TestConsoleOutput:

    def test_say_turns_everything_into_ascii(self, capsys):
        app_main.say("Müller – „Benötigt“ ✓")
        out = capsys.readouterr().out
        assert out.isascii() and out.strip() == "M?ller ? ?Ben?tigt? ?"

    def test_without_a_console_nothing_happens(self, monkeypatch):
        monkeypatch.setattr(sys, "stdout", None)
        app_main.say("hello")                                # a windowless build has no stdout

    def test_a_broken_console_does_not_stop_the_app(self, monkeypatch):
        class Broken:
            def write(self, text):
                raise OSError("closed")

            def flush(self):
                raise OSError("closed")

        monkeypatch.setattr(sys, "stdout", Broken())
        app_main.say("hello")

    def test_a_whole_run_prints_only_ascii_even_without_a_reference(self, monkeypatch, capsys, tmp_path):
        monkeypatch.setattr(app_main, "reference_path", lambda: tmp_path / "nirgends.xlsx")
        thread, result, started = run_app("--no-browser")
        stop(thread, result, started.get(timeout=10))
        out = capsys.readouterr().out
        assert out.isascii() and "Reference file: not fully usable" in out and "?" not in out

    def test_a_run_without_a_console_works(self, monkeypatch):
        monkeypatch.setattr(sys, "stdout", None)
        thread, result, started = run_app("--no-browser")
        srv = started.get(timeout=10)
        assert own_request(srv, "GET")[0] == 200
        stop(thread, result, srv)


class TestReferenceLoading:

    def test_the_default_location_is_the_bundle(self):
        assert app_main.reference_path() == server.bundle_dir() / "reference_amount.xlsx"
        assert app_main.reference_path().is_file()

    def test_the_location_does_not_depend_on_the_working_folder(self, monkeypatch, tmp_path):
        monkeypatch.chdir(tmp_path)
        assert app_main.reference_path() == server.bundle_dir() / "reference_amount.xlsx" and app_main.reference_path().is_file()

    def test_in_a_frozen_program_it_is_looked_for_in_the_unpack_folder(self, monkeypatch, tmp_path):
        monkeypatch.setattr(sys, "frozen", True, raising=False)
        monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path), raising=False)
        assert app_main.reference_path() == tmp_path / "reference_amount.xlsx"

    def test_a_loaded_reference_is_reported_with_its_size(self, capsys):
        data = app_main.load_reference(app_main.reference_path())
        assert data["amounts"] and f"{len(data['amounts'])} required amounts loaded" in capsys.readouterr().out

    def test_a_missing_file_gives_the_empty_structure_and_a_reason(self, tmp_path):
        data = app_main.load_reference(tmp_path / "gone.xlsx")
        assert data["amounts"] == {} and len(data["problems"]) == 1
