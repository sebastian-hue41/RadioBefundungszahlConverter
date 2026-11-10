"""
main.py — start the local app.

    python -m local_app.main [--no-browser]

Starts the server on a free port of 127.0.0.1, prints the address, opens it in the
default browser and serves until the page's "Server beenden" button or Ctrl+C.
The same file is the entry point of the PyInstaller build.

Everything printed here is plain ASCII: a Windows console with a legacy code page
must not be able to break the start, and a windowless build has no console at all.
"""

from __future__ import annotations

import argparse
import logging
import math
import os
import sys
import threading
import webbrowser
from collections.abc import Callable, MutableMapping
from pathlib import Path

# Absolute imports: PyInstaller runs this file as a script, where relative ones fail.
from core.input_handler import load_reference_data
from local_app import server as server_module
from local_app.server import AppServer, bundle_dir, create_server

REFERENCE_NAME = "reference_amount.xlsx"


def say(text: str) -> None:
    """Print one line of plain ASCII (anything else becomes '?'). Without a console (stdout is None) print does nothing."""
    try:
        print(text.encode("ascii", "replace").decode("ascii"), flush=True)
    except (OSError, ValueError):       # a closed or broken console must never stop the app
        pass


def reference_path() -> Path:
    """The reference file shipped with the app (the only xlsx inside the program)."""
    return bundle_dir() / REFERENCE_NAME


def load_reference(path: Path) -> dict:
    """
    Core's reference data for `path`. Never fatal: without a usable file the Excel has no
    "Benoetigt" column and says why (core puts the reasons on the sheet; the page shows hint W301).
    """
    data = load_reference_data(path)
    if data["amounts"] and not data["problems"]:
        say(f"Reference file: {len(data['amounts'])} required amounts loaded.")
    else:
        say(f"Reference file: not fully usable ({len(data['problems'])} problem(s)); "
            "the Excel will say what is missing.")
    return data


def restore_library_path(
    env: MutableMapping | None = None, *, frozen: bool | None = None, platform: str | None = None,
) -> None:
    """
    In a frozen Linux build, undo PyInstaller's change of LD_LIBRARY_PATH.

    The bootloader points it at the unpacked program so that its own libraries are found.
    Programs we start inherit that and can crash on the wrong libraries; the default browser
    is such a program. PyInstaller keeps the original in LD_LIBRARY_PATH_ORIG: put it back,
    or remove the variable if there was none (an empty value is removed too: on some systems
    it means "the current folder"). Anything else is left alone: not frozen, not Linux.
    """
    env = os.environ if env is None else env
    frozen = bool(getattr(sys, "frozen", False)) if frozen is None else frozen
    platform = sys.platform if platform is None else platform
    if not frozen or not platform.startswith("linux"):
        return
    original = env.get("LD_LIBRARY_PATH_ORIG")
    if original:
        env["LD_LIBRARY_PATH"] = original
    else:
        env.pop("LD_LIBRARY_PATH", None)


def open_in_browser(url: str, opener: Callable[..., object]) -> None:
    """Open `url` in a thread of its own (some systems block here); say so if that did not work."""
    def work() -> None:
        try:
            opened = opener(url, new=2)
        except Exception:               # no browser, no display, a broken launcher: not fatal
            opened = False
        if opened is False:
            say("Could not open a browser. Please open the address above in one.")

    threading.Thread(target=work, name="open-browser", daemon=True).start()


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m local_app.main",
        description="Local app: evaluate a staff statistics file against the Leistungen maps in the browser.",
    )
    parser.add_argument("--no-browser", action="store_true", help="do not open the browser; only print the address")
    parser.add_argument("--idle-minutes", type=float, default=server_module.IDLE_SECONDS / 60, metavar="N",
                        help="stop by itself after N minutes without use (default 30; 0 = never)")
    return parser


def main(
    argv: list[str] | None = None,
    *,
    opener: Callable[..., object] = webbrowser.open,
    on_ready: Callable[[AppServer], None] | None = None,
) -> int:
    """
    Run the app until it is stopped. Returns the exit code: 0 after a normal stop (the button
    or Ctrl+C), 1 if it could not start. `opener` and `on_ready(server)` exist for the tests.
    """
    parser = _parser()
    args = parser.parse_args(argv)
    if not math.isfinite(args.idle_minutes) or args.idle_minutes < 0:
        parser.error("--idle-minutes must be a number, 0 or more")
    # core logs what it could not use, with file paths. Nothing of that belongs on the console of the app:
    # the reasons are on the sheet of the Excel and in the hints of the page.
    if not any(isinstance(handler, logging.NullHandler) for handler in logging.getLogger("core").handlers):
        logging.getLogger("core").addHandler(logging.NullHandler())

    reference = load_reference(reference_path())
    try:
        server = create_server(reference, idle_seconds=args.idle_minutes * 60)
    except (OSError, ValueError) as exc:
        say(f"Could not start the app: {type(exc).__name__}. Is the program complete? "
            "Please report this message.")
        return 1

    say("")
    say(f"Befundungszahlen app is running at: {server.url}")
    say("Keep this window open while you work. To stop, use the 'Server beenden' button or press Ctrl+C.")
    if args.idle_minutes:
        say(f"It also stops by itself after {args.idle_minutes:g} minutes without use.")
    say("")
    if not args.no_browser:
        restore_library_path()
        open_in_browser(server.url, opener)
    try:
        if on_ready is not None:
            on_ready(server)
        server.serve_forever()
    except KeyboardInterrupt:
        say("Stopped (Ctrl+C).")
    finally:
        server.server_close()
    if server.state.expired:
        say(f"The app stopped by itself after {args.idle_minutes:g} minutes without use. You can close this window.")
    else:
        say("The app has stopped. You can close this window.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
