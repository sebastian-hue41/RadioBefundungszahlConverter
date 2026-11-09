"""
server.py — what the local app does with an upload.

evaluate() turns the uploaded files into the finished Excel; it is plain Python on
bytes, so everything it does can be tested without a web server. Order of work:

  1. every file is validated before any of them is parsed (scan_xlsx: core's checks
     with the error codes of the file's slot),
  2. each file is loaded with its own slot, so a problem is reported for the file
     that has it,
  3. the department guard, the merge of both maps (if both are chosen) and the
     narrowing of the statistics to the chosen departments (scope.py),
  4. core's process() and build_xlsx_bytes(); then a last look at the bytes of the
     result (security.assert_no_formulas).

Everything stays in memory. The only thing openpyxl writes while it builds the Excel is
a temporary file with the result sheet (counts and Leistung names, never anything
from the statistics beyond that); it deletes it right away.
"""

from __future__ import annotations

import hmac
import json
import re
import secrets
import socketserver
import sys
import threading
from collections.abc import Mapping
from dataclasses import dataclass
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qsl, quote

from core.input_handler import MAX_FILE_SIZE_MB, load_map, load_statistics
from core.logic import process
from core.output_handler import OutputValidationError, build_xlsx_bytes

from . import __version__
from .errors import (
    ALLRAD,
    CATALOG,
    HINTS,
    MAP_SLOTS,
    NRAD,
    SELECTION_NOTE_PREFIX,
    SLOT_LABELS,
    STATISTICS,
    AppError,
    from_core_error,
    from_output_error,
    hint,
    summarize_hints,
    unexpected_error,
)
from .scope import apply_scope, check_department, merge_maps
from .security import assert_no_formulas, safe_name, scan_xlsx

_SELECTION_LABELS = {ALLRAD: "Allrad", NRAD: "NRAD"}
_MAX_SHOWN = 100   # longest text from a file (staff name, dates) passed on to the page

# The upload is one request: the files back to back, their sizes in the URL.
#   ?stats=<bytes>&allrad=<bytes>&nrad=<bytes>&allrad_name=<text>&nrad_name=<text>
# The maps that are present decide which departments are evaluated.
MAX_BODY_BYTES = 3 * MAX_FILE_SIZE_MB * 1024 * 1024 + 1024 * 1024   # three files at the limit, plus a margin
_SIZE_PARAMS = {STATISTICS: "stats", ALLRAD: "allrad", NRAD: "nrad"}   # in the order of the body
_NAME_PARAMS = {ALLRAD: "allrad_name", NRAD: "nrad_name"}
_SIZE = re.compile(r"[0-9]{1,12}")                                    # ASCII digits only: no sign, no exponent
_MAX_PARAMS = 30
_MAX_RAW_NAME = 300


def bundle_dir() -> Path:
    """
    The folder with the files shipped with the app (the page, the reference file):
    the project root normally, the folder PyInstaller unpacks to when frozen.
    """
    if getattr(sys, "frozen", False):
        unpacked = getattr(sys, "_MEIPASS", None)
        if unpacked:
            return Path(unpacked)
    return Path(__file__).resolve().parent.parent


# ── Evaluation ─────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class Evaluation:
    """The result of one evaluation: the download and what the page shows about it."""
    filename: str
    data: bytes
    summary: dict


def _as_bytes(slot: str, value) -> bytes:
    if not isinstance(value, (bytes, bytearray, memoryview)):
        raise ValueError(f"the content of {slot!r} must be bytes, not {type(value).__name__}")
    return bytes(value)


def _shown(text) -> str | None:
    """A text from a file for the page: clipped; the page inserts it as plain text."""
    if text is None:
        return None
    text = str(text)
    return text if len(text) <= _MAX_SHOWN else text[: _MAX_SHOWN - 1] + "…"


def _file_name(core_name: str, selection: list[str]) -> str:
    """core's name for the Excel with the selection appended: ..._NRAD.xlsx"""
    suffix = "_" + "_".join(_SELECTION_LABELS[slot] for slot in selection)
    stem = safe_name(core_name)
    if stem.lower().endswith(".xlsx"):
        stem = stem[: -len(".xlsx")]
    return f"{stem}{suffix}.xlsx"


def _named(slot: str, names: Mapping[str, str]) -> str:
    """'NRAD-Map „datei.xlsx“' for the note on the sheet; the name is cleaned first."""
    given = names.get(slot)
    return f"{SLOT_LABELS[slot]} „{safe_name(given)}“" if given else SLOT_LABELS[slot]


def _reference_hint(reference_data: dict | None) -> list[dict]:
    """W301 when there is no usable reference (no 'Benötigt' column), W302 when only some rows were unusable."""
    if not reference_data or not reference_data.get("amounts"):
        return [hint("W301")]
    if reference_data.get("problems"):
        return [hint("W302")]
    return []


def evaluate(
    files: Mapping[str, bytes],
    names: Mapping[str, str] | None = None,
    reference_data: dict | None = None,
) -> Evaluation:
    """
    Evaluate the uploaded files.

    `files` maps a slot (STATISTICS, ALLRAD, NRAD) to the content of the file. The maps
    that are present are the departments that are evaluated: Allrad, NRAD or both.
    `names` maps a slot to the file name the user chose (only used as text in a note
    on the sheet, after safe_name). `reference_data` is core's load_reference_data
    result; without it the sheet has no "Benötigt" column (hint W301).

    Raises AppError for everything the user can fix and for what was refused for
    safety reasons; anything nobody foresaw becomes E999 with only the error type and
    place (never the message, it could hold file content). Nothing is returned for a
    failed evaluation.
    """
    unknown = set(files) - {STATISTICS, *MAP_SLOTS}
    if unknown:
        raise ValueError(f"unknown file slots: {sorted(unknown)!r}")
    contents = {slot: _as_bytes(slot, value) for slot, value in files.items() if value is not None}
    try:
        return _evaluate(contents, names or {}, reference_data)
    except AppError:
        raise
    except Exception as exc:
        raise unexpected_error(exc) from exc


def _evaluate(files: dict[str, bytes], names: Mapping[str, str], reference_data: dict | None) -> Evaluation:
    maps = [slot for slot in MAP_SLOTS if slot in files]
    if STATISTICS not in files:
        raise AppError("E302", STATISTICS, "Es wurde keine Mitarbeiterstatistik übergeben.")
    if not maps:
        raise AppError("E301")

    # 1. Validate all files before any of them is parsed.
    for slot in (STATISTICS, *maps):
        scan_xlsx(files[slot], slot)

    # 2. Load each file with its own slot, so a problem is reported for the right file.
    try:
        stats = load_statistics(files[STATISTICS])
    except Exception as exc:
        raise from_core_error(exc, STATISTICS) from exc
    loaded = {}
    for slot in maps:
        try:
            loaded[slot] = load_map(files[slot])
        except Exception as exc:
            raise from_core_error(exc, slot) from exc

    # 3. Right maps in the right fields, merged if both; the statistics narrowed to the selection.
    for slot in maps:
        check_department(loaded[slot], slot)
    map_data = merge_maps(loaded[ALLRAD], loaded[NRAD]) if len(maps) == 2 else loaded[maps[0]]
    scoped = apply_scope(stats, ALLRAD in maps, NRAD in maps)
    removed = len(stats["leistungen"]) - len(scoped["leistungen"])

    label = " + ".join(_SELECTION_LABELS[slot] for slot in maps)
    chosen = "; ".join(_named(slot, names) for slot in maps)
    scoped["errors"].append(f"{SELECTION_NOTE_PREFIX} Ausgewertet wurde: {label}. {chosen}.")
    if removed:
        scoped["errors"].append(f"{SELECTION_NOTE_PREFIX} {hint('W401', removed)['text']}")

    # 4. Evaluate, build the Excel, look at the finished bytes once more.
    result = process(scoped, map_data)
    try:
        core_name, data = build_xlsx_bytes(result, reference_data=reference_data or None)
    except OutputValidationError as exc:
        raise from_output_error(exc) from exc
    assert_no_formulas(data)

    hints = summarize_hints(result["errors"]) + _reference_hint(reference_data)
    if removed:
        hints.append(hint("W401", removed))
    filename = _file_name(core_name, maps)
    summary = {
        "filename": filename,
        "mitarbeiter": _shown(result.get("mitarbeiter")),
        "befunddatum": _shown(result.get("befunddatum")),
        "total_counted": result["total_counted"],
        "selection": list(maps),
        "selection_label": label,
        "hints": sorted(hints, key=lambda h: h["code"]),
    }
    return Evaluation(filename, data, summary)


# ── Splitting an upload ────────────────────────────────────────────────────────

def _one_value(params: list[tuple[str, str]], name: str) -> str | None:
    """The value of parameter `name`; None when it is absent; E304 when it was sent twice."""
    values = [value for key, value in params if key == name]
    if len(values) > 1:
        raise AppError("E304", detail=f"Der Parameter „{name}“ wurde mehrfach gesendet.")
    return values[0] if values else None


def _size(params: list[tuple[str, str]], name: str) -> int | None:
    text = _one_value(params, name)
    if text is None:
        return None
    if not _SIZE.fullmatch(text):
        raise AppError("E304", detail=f"Ungültige Größenangabe für „{name}“: {text[:20]!r}")
    return int(text)


def split_upload(body: bytes, query: "str | Mapping[str, str]") -> tuple[dict[str, bytes], dict[str, str]]:
    """
    Cut the body of an upload into its files, by the sizes the query declares.

    Returns (files, names) as evaluate() takes them: files maps STATISTICS and each map
    that was sent to its content, names maps a map slot to the cleaned file name (only
    if one was sent). The sizes must be whole numbers and add up to the body exactly.
    Parameters that are not ours are ignored; nothing from the query is ever used as a
    slot or a path.

    Raises AppError E305 for a body over MAX_BODY_BYTES, E302 when the statistics or
    every map is missing, E304 for sizes that are invalid, sent twice or do not add up.
    """
    if len(body) > MAX_BODY_BYTES:
        raise AppError("E305", detail=f"{len(body):,} Bytes empfangen; erlaubt sind {MAX_BODY_BYTES:,}.")
    if isinstance(query, str):
        try:
            params = parse_qsl(query.lstrip("?"), keep_blank_values=True, max_num_fields=_MAX_PARAMS)
        except ValueError as exc:
            raise AppError("E304", detail="Zu viele Parameter in der Anfrage.") from exc
    else:
        params = [(str(key), str(value)) for key, value in query.items()]

    sizes = {slot: _size(params, name) for slot, name in _SIZE_PARAMS.items()}
    if sizes[STATISTICS] is None:
        raise AppError("E302", STATISTICS, "Die Größe der Mitarbeiterstatistik fehlt in der Anfrage.")
    if sizes[ALLRAD] is None and sizes[NRAD] is None:
        raise AppError("E302", detail="Es wurde keine Map gesendet.")
    declared = sum(size for size in sizes.values() if size is not None)
    if declared != len(body):
        raise AppError("E304", detail=f"Angekündigt: {declared:,} Bytes, empfangen: {len(body):,}.")

    files, names, start = {}, {}, 0
    for slot, size in sizes.items():
        if size is None:
            continue
        files[slot] = body[start:start + size]
        start += size
        raw_name = _one_value(params, _NAME_PARAMS[slot]) if slot in _NAME_PARAMS else None
        if raw_name:
            names[slot] = safe_name(raw_name[:_MAX_RAW_NAME])
    return files, names


# ── The HTTP server ────────────────────────────────────────────────────────────
#
# All routes sit under a random per-launch prefix: /<token>/
#
#   GET  /<token>/           the page
#   POST /<token>/evaluate   the upload (see split_upload); JSON summary or error
#   GET  /<token>/download   the latest Excel
#   POST /<token>/reset      forget the latest result
#   POST /<token>/shutdown   forget the result, answer, stop the server
#
# Anything else is a 404. The server listens on 127.0.0.1 only. Other web pages the
# user has open cannot use it: the prefix is secret, the Host header must be our own
# address (DNS rebinding) and a foreign Origin is refused.

_LENGTH = re.compile(r"[0-9]{1,12}")
_PAGE_PLACEHOLDER = "{{NONCE}}"
_CATALOG_PLACEHOLDER = "{{CATALOG}}"
_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_STATUS = {"E301": 400, "E302": 400, "E304": 400, "E305": 413, "E307": 409, "E901": 404, "E999": 500}
_DEFAULT_ERROR_STATUS = 422

# Sent with every response.
_SECURITY_HEADERS = (
    ("X-Content-Type-Options", "nosniff"),
    ("Cache-Control", "no-store"),
    ("Referrer-Policy", "no-referrer"),
    ("X-Frame-Options", "DENY"),
    ("Cross-Origin-Resource-Policy", "same-origin"),
    ("Cross-Origin-Opener-Policy", "same-origin"),
)
# The page may run its own script and style (marked with the nonce) and talk to this
# server; nothing else: no external script, style, image or connection, no framing.
_PAGE_POLICY = (
    "default-src 'none'; script-src 'nonce-{nonce}'; style-src 'nonce-{nonce}'; connect-src 'self'; "
    "base-uri 'none'; form-action 'none'; frame-ancestors 'none'"
)
_DATA_POLICY = "default-src 'none'; frame-ancestors 'none'"


def catalog_json() -> str:
    """
    Everything the page needs to word its own errors and hints, as JSON that is safe inside a
    <script> element: the characters < > & are written as \\uXXXX, so no text of the catalog
    can close the element or open another one.
    """
    data = {
        "errors": {code: {"title": entry.title, "hint": entry.hint} for code, entry in CATALOG.items()},
        "hints": {code: list(texts) for code, texts in HINTS.items()},
        "slots": dict(SLOT_LABELS),
        "limits": {"file_bytes": MAX_FILE_SIZE_MB * 1024 * 1024, "body_bytes": MAX_BODY_BYTES},
        "version": __version__,
    }
    text = json.dumps(data, ensure_ascii=True, separators=(",", ":"))
    return text.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")


def _load_page() -> str:
    return (bundle_dir() / "local_app" / "static" / "index.html").read_text(encoding="utf-8")


class _State:
    """What the server remembers: the latest result (one slot) and whether it is shutting down."""

    def __init__(self, reference_data: dict | None, page_html: str):
        self.reference_data = reference_data
        self.page_html = page_html
        self.busy = threading.Lock()          # held while an evaluation runs; a second one is refused at once
        self._lock = threading.Lock()         # protects the result slot
        self._result: Evaluation | None = None
        self._closed = False

    def set_result(self, evaluation: Evaluation | None) -> None:
        with self._lock:
            if not self._closed:              # a result that finishes after shutdown is dropped
                self._result = evaluation

    def get_result(self) -> Evaluation | None:
        with self._lock:
            return self._result

    def close(self) -> None:
        with self._lock:
            self._closed = True
            self._result = None


class AppServer(ThreadingHTTPServer):
    """The local app's server. `url` is the address to open; `token` is its secret prefix."""

    daemon_threads = True                     # a stuck evaluation must not keep the program alive
    allow_reuse_address = False               # on Windows this would let another program share the port

    def __init__(self, port: int, state: _State):
        self.state = state
        self.token = secrets.token_urlsafe(24)
        super().__init__(("127.0.0.1", port), _Handler)
        self.port = self.server_address[1]
        self.allowed_hosts = {f"127.0.0.1:{self.port}", f"localhost:{self.port}"}
        self.allowed_origins = {f"http://{host}" for host in self.allowed_hosts}
        self.url = f"http://127.0.0.1:{self.port}/{self.token}/"

    def server_bind(self) -> None:
        socketserver.TCPServer.server_bind(self)   # HTTPServer's own bind looks up the host name, which can be slow
        self.server_name, self.server_port = "127.0.0.1", self.server_address[1]

    def handle_error(self, request, client_address) -> None:
        pass                                  # nothing is printed: a windowless build has no stderr

    def shutdown_soon(self) -> None:
        """Stop serving and close the port, from a thread of its own (shutdown() waits for the serve loop)."""
        def stop():
            self.shutdown()
            self.server_close()
        threading.Thread(target=stop, daemon=True).start()


def create_server(reference_data: dict | None = None, port: int = 0, page_html: str | None = None) -> AppServer:
    """
    The server for one run of the app, listening on 127.0.0.1 (port 0: a free one).
    Call serve_forever() on it; open `url`. `page_html` replaces the bundled page (tests).
    """
    html = _load_page() if page_html is None else page_html
    if _PAGE_PLACEHOLDER not in html:
        raise ValueError(f"the page has no {_PAGE_PLACEHOLDER} for the nonce of its script and style")
    return AppServer(port, _State(reference_data, html))


class _Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.0"             # one request per connection: no keep-alive, nothing left unread
    timeout = 60                              # seconds a client may stall before the connection is dropped
    server_version = "local"
    sys_version = ""

    # ── Quiet ──────────────────────────────────────────────────────────────
    def log_message(self, format, *args):     # noqa: A002 (name fixed by the base class)
        pass

    def version_string(self) -> str:
        return "local"

    # ── Answers ────────────────────────────────────────────────────────────
    def _send(self, status: int, body: bytes = b"", content_type: str = "text/plain; charset=utf-8",
              policy: str = _DATA_POLICY, extra: tuple = ()) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Content-Security-Policy", policy)
        for name, value in (*_SECURITY_HEADERS, *extra):
            self.send_header(name, value)
        self.end_headers()
        if body:
            self.wfile.write(body)

    def _json(self, status: int, payload: dict) -> None:
        self._send(status, json.dumps(payload, ensure_ascii=True).encode("ascii"), "application/json")

    def _app_error(self, error: AppError) -> None:
        self._json(_STATUS.get(error.code, _DEFAULT_ERROR_STATUS), {"ok": False, "error": error.to_dict()})

    def send_error(self, code, message=None, explain=None):
        # Errors found while parsing the request line or the headers (too long, malformed, ...).
        error = AppError("E304", detail=f"Die Anfrage war ungültig (HTTP {code}).")
        self._json(code, {"ok": False, "error": error.to_dict()})

    # ── Which requests are accepted ────────────────────────────────────────
    def _own_origin(self) -> bool:
        server = self.server
        if (self.headers.get("Host") or "").lower() not in server.allowed_hosts:
            return False
        origin = self.headers.get("Origin")
        if origin is not None and origin.lower() not in server.allowed_origins:
            return False                      # also "null" and every other site
        site = self.headers.get("Sec-Fetch-Site")
        return site is None or site in ("same-origin", "none")

    def _route(self) -> tuple[str, str] | None:
        """(route, query) when the path is /<token>/<route>; None for anything else."""
        target = self.path.split("#", 1)[0]
        path, _, query = target.partition("?")
        if not path.startswith("/"):
            return None
        segment, slash, rest = path[1:].partition("/")
        if not slash or not hmac.compare_digest(segment.encode("utf-8", "replace"), self.server.token.encode("ascii")):
            return None
        return "/" + rest, query

    _ROUTES = {
        "/": {"GET": "_page"},
        "/download": {"GET": "_download"},
        "/evaluate": {"POST": "_evaluate"},
        "/reset": {"POST": "_reset"},
        "/shutdown": {"POST": "_shutdown"},
    }

    def _dispatch(self) -> None:
        try:
            if not self._own_origin():
                return self._send(403)
            found = self._route()
            methods = self._ROUTES.get(found[0]) if found else None
            if methods is None:
                return self._send(404)
            action = methods.get(self.command)
            if action is None:
                return self._send(405)
            getattr(self, action)(found[1])
        except OSError:
            pass                              # the client went away or stalled; nobody is left to answer
        except Exception as exc:              # a bug: the type and place, never the message
            try:
                self._app_error(unexpected_error(exc))
            except OSError:
                pass

    do_GET = do_POST = do_PUT = do_DELETE = do_PATCH = do_HEAD = do_OPTIONS = _dispatch

    # ── Routes ─────────────────────────────────────────────────────────────
    def _page(self, query: str) -> None:
        nonce = secrets.token_urlsafe(18)
        html = self.server.state.page_html.replace(_CATALOG_PLACEHOLDER, catalog_json()).replace(_PAGE_PLACEHOLDER, nonce)
        self._send(200, html.encode("utf-8"), "text/html; charset=utf-8", _PAGE_POLICY.format(nonce=nonce))

    def _download(self, query: str) -> None:
        evaluation = self.server.state.get_result()
        if evaluation is None:
            return self._app_error(AppError("E901"))
        fallback = re.sub(r"[^A-Za-z0-9._() -]", "_", evaluation.filename)
        disposition = f"attachment; filename=\"{fallback}\"; filename*=UTF-8''{quote(evaluation.filename, safe='')}"
        self._send(200, evaluation.data, _XLSX, extra=(("Content-Disposition", disposition),))

    def _reset(self, query: str) -> None:
        self.server.state.set_result(None)
        self._json(200, {"ok": True})

    def _shutdown(self, query: str) -> None:
        self.server.state.close()
        self._json(200, {"ok": True})
        self.wfile.flush()
        self.server.shutdown_soon()

    def _content_length(self) -> int:
        if self.headers.get("Transfer-Encoding") is not None:
            raise AppError("E304", detail="Die Anfrage darf nicht in Teilen gesendet werden.")
        values = self.headers.get_all("Content-Length") or []
        if len(values) != 1 or not _LENGTH.fullmatch(values[0].strip()):
            raise AppError("E304", detail="Die Längenangabe der Anfrage fehlt oder ist ungültig.")
        length = int(values[0])
        if length > MAX_BODY_BYTES:
            raise AppError("E305", detail=f"{length:,} Bytes angekündigt; erlaubt sind {MAX_BODY_BYTES:,}.")
        return length

    def _read_body(self, length: int) -> bytes:
        try:
            body = self.rfile.read(length)
        except OSError as exc:                # a stalled or dropped connection
            raise AppError("E304", detail="Die Verbindung wurde unterbrochen oder war zu langsam.") from exc
        if len(body) != length:
            raise AppError("E304", detail=f"Angekündigt: {length:,} Bytes, empfangen: {len(body):,}.")
        return body

    def _evaluate(self, query: str) -> None:
        state = self.server.state
        if not state.busy.acquire(blocking=False):
            return self._app_error(AppError("E307"))      # answered at once, never queued
        try:
            state.set_result(None)            # whatever happens now, no older result stays behind
            try:
                body = self._read_body(self._content_length())
                files, names = split_upload(body, query)
                del body                      # the slices are copies; the uploads are dropped after this call
                evaluation = evaluate(files, names, state.reference_data)
                files = names = None
            except AppError as error:
                return self._app_error(error)
            except Exception as exc:
                return self._app_error(unexpected_error(exc))
            state.set_result(evaluation)
            self._json(200, {"ok": True, "summary": evaluation.summary})
        finally:
            state.busy.release()
