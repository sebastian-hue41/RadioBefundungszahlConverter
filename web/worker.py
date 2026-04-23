"""
web/worker.py — iii worker that exposes the converter as an HTTP endpoint.

Entrypoint:
    python -m web.worker          (from the project root)

Environment variables:
    III_ENGINE_URL   WebSocket URL of the iii engine (default: ws://localhost:49134)
    WORKER_NAME      Display name shown in the iii console   (default: befundungszahl-converter)

HTTP API:
    POST /convert
        Content-Type: multipart/form-data
        Fields:
            statistics   — the MitarbeiterStatistik .xlsx file
            map          — the Leistungen map .xlsx file
        Query params:
            include_underscore=true   (optional) include _-prefixed map columns

        Response 200: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
            Content-Disposition: attachment; filename="..."
        Response 400: application/json  {"error": "..."}  — bad request / missing fields
        Response 413: application/json  {"error": "..."}  — upload too large
        Response 422: application/json  {"error": "..."}  — validation failure
        Response 500: application/json  {"error": "..."}  — unexpected server error
        Response 503: application/json  {"error": "..."}  — processing timed out

The output xlsx is built entirely in memory — no files are written to disk.
"""

from __future__ import annotations

import asyncio
import email.parser
import email.policy
import io
import json
import logging
import os
import sys

# Allow running as  python -m web.worker  from the project root.
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from iii import InitOptions, register_worker, http
from iii.types import HttpRequest, HttpResponse

from core.input_handler import (
    FileValidationError,
    MapValidationError,
    MAX_FILE_SIZE_MB,
    StatsValidationError,
    load_map,
    load_statistics,
)
from core.logic import process
from core.output_handler import build_xlsx_bytes

log = logging.getLogger(__name__)

_XLSX_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
_CHUNK_SIZE = 64 * 1024  # 64 KB

# Two files at the per-file limit plus generous multipart framing overhead.
_MAX_BODY_BYTES = (2 * MAX_FILE_SIZE_MB + 10) * 1024 * 1024

# Maximum number of accepted multipart fields (statistics + map = 2; a few extra
# for leniency with browsers that add metadata fields).
_MAX_FIELDS = 10

# Wall-clock budget for the entire request (read + parse + process + write).
_PROCESSING_TIMEOUT_S = 60.0

# Security headers sent with every response.
_SECURITY_HEADERS = {
    "X-Content-Type-Options": "nosniff",
    "Cache-Control": "no-store",
}


# ── Multipart parsing ──────────────────────────────────────────────────────────

def _parse_multipart(body: bytes, content_type: str) -> dict[str, bytes]:
    """
    Parse a multipart/form-data body and return {field_name: file_bytes}.
    Uses the stdlib email parser — no extra dependencies.
    Silently ignores parts without a name or without a payload.
    """
    raw = f"Content-Type: {content_type}\r\n\r\n".encode() + body
    msg = email.parser.BytesParser(policy=email.policy.compat32).parsebytes(raw)

    fields: dict[str, bytes] = {}
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        disposition = part.get("Content-Disposition", "")
        name: str | None = None
        for segment in disposition.split(";"):
            segment = segment.strip()
            if segment.startswith("name="):
                name = segment[5:].strip().strip('"')
        if name:
            payload = part.get_payload(decode=True)
            if payload is not None:
                fields[name] = payload

    return fields


# ── Response helpers ───────────────────────────────────────────────────────────

async def _send_json_error(res: HttpResponse, status: int, message: str) -> None:
    body = json.dumps({"error": message}).encode()
    await res.status(status)
    await res.headers({
        **_SECURITY_HEADERS,
        "Content-Type": "application/json",
        "Content-Length": str(len(body)),
    })
    res.writer.stream.end(body)


# ── Core request logic (wrapped in a timeout by the outer handler) ─────────────

async def _process_request(req: HttpRequest, res: HttpResponse) -> None:
    # ── Method guard ──────────────────────────────────────────────────────────
    if getattr(req, "method", "POST").upper() != "POST":
        await _send_json_error(res, 405, "Method not allowed — use POST.")
        return

    # ── Content-Length pre-check (reject oversized uploads before reading) ────
    content_length_raw = req.headers.get("content-length", "")
    if isinstance(content_length_raw, list):
        content_length_raw = content_length_raw[0]
    if content_length_raw:
        try:
            declared_length = int(content_length_raw)
            if declared_length > _MAX_BODY_BYTES:
                await _send_json_error(
                    res, 413,
                    f"Upload too large ({declared_length // (1024 * 1024)} MB). "
                    f"Maximum combined upload size is {_MAX_BODY_BYTES // (1024 * 1024)} MB.",
                )
                return
        except ValueError:
            pass  # malformed header — let the read proceed and check afterwards

    # ── Read and size-check the request body ──────────────────────────────────
    body = await req.request_body.read_all()

    if len(body) > _MAX_BODY_BYTES:
        await _send_json_error(
            res, 413,
            f"Upload too large ({len(body) // (1024 * 1024)} MB). "
            f"Maximum combined upload size is {_MAX_BODY_BYTES // (1024 * 1024)} MB.",
        )
        return

    content_type = req.headers.get("content-type", "")
    if isinstance(content_type, list):
        content_type = content_type[0]

    if "multipart/form-data" not in content_type:
        await _send_json_error(
            res, 400,
            "Expected multipart/form-data with 'statistics' and 'map' fields.",
        )
        return

    # ── Parse multipart ───────────────────────────────────────────────────────
    fields = _parse_multipart(body, content_type)
    del body  # release the raw body from memory as early as possible

    if len(fields) > _MAX_FIELDS:
        await _send_json_error(
            res, 400,
            f"Too many form fields ({len(fields)}); at most {_MAX_FIELDS} are accepted.",
        )
        return

    stats_bytes = fields.get("statistics")
    map_bytes = fields.get("map")

    if not stats_bytes:
        await _send_json_error(res, 400, "Missing required field: 'statistics'.")
        return
    if not map_bytes:
        await _send_json_error(res, 400, "Missing required field: 'map'.")
        return

    # Per-field size check (defence-in-depth: validate_xlsx_bytes also checks,
    # but catching it here gives a cleaner 413 rather than a 422).
    per_field_limit = MAX_FILE_SIZE_MB * 1024 * 1024
    for field_name, field_bytes in (("statistics", stats_bytes), ("map", map_bytes)):
        if len(field_bytes) > per_field_limit:
            await _send_json_error(
                res, 413,
                f"Field '{field_name}' is too large "
                f"({len(field_bytes) // (1024 * 1024)} MB). "
                f"Maximum is {MAX_FILE_SIZE_MB} MB per file.",
            )
            return

    # ── Optional query params ─────────────────────────────────────────────────
    include_underscore_raw = req.query_params.get("include_underscore", "false")
    if isinstance(include_underscore_raw, list):
        include_underscore_raw = include_underscore_raw[0]
    # Treat only the literal string "true" as truthy — anything else is false.
    include_underscore = include_underscore_raw.strip().lower() == "true"

    # ── Core processing (no disk I/O) ─────────────────────────────────────────
    stats_data = load_statistics(io.BytesIO(stats_bytes))
    map_data = load_map(io.BytesIO(map_bytes), include_underscore_columns=include_underscore)
    result = process(stats_data, map_data)
    filename, xlsx_bytes = build_xlsx_bytes(result)

    # ── Stream the response ───────────────────────────────────────────────────
    await res.status(200)
    await res.headers({
        **_SECURITY_HEADERS,
        "Content-Type": _XLSX_CONTENT_TYPE,
        "Content-Disposition": f'attachment; filename="{filename}"',
        "Content-Length": str(len(xlsx_bytes)),
    })

    writer = res.writer
    for offset in range(0, len(xlsx_bytes), _CHUNK_SIZE):
        await writer.write(xlsx_bytes[offset : offset + _CHUNK_SIZE])
    writer.close()


# ── HTTP handler ───────────────────────────────────────────────────────────────

async def handle_convert(req: HttpRequest, res: HttpResponse) -> None:
    """
    POST /convert — receive two xlsx uploads, run the converter, stream result back.
    All error paths return JSON; the success path streams an xlsx.
    """
    try:
        await asyncio.wait_for(_process_request(req, res), timeout=_PROCESSING_TIMEOUT_S)

    except asyncio.TimeoutError:
        log.warning("handle_convert timed out after %.0fs", _PROCESSING_TIMEOUT_S)
        await _send_json_error(res, 503, "Request processing timed out.")

    except (FileValidationError, StatsValidationError) as exc:
        log.warning("Statistics validation error: %s", exc)
        await _send_json_error(res, 422, str(exc))

    except MapValidationError as exc:
        log.warning("Map validation error (recoverable=%s): %s", exc.recoverable, exc)
        await _send_json_error(res, 422, str(exc))

    except Exception:
        log.exception("Unexpected error in handle_convert")
        await _send_json_error(res, 500, "Internal server error.")


# ── Worker registration ────────────────────────────────────────────────────────

def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

    engine_url = os.environ.get("III_ENGINE_URL", "ws://localhost:49134")
    worker_name = os.environ.get("WORKER_NAME", "befundungszahl-converter")

    log.info("Connecting to iii engine at %s ...", engine_url)
    client = register_worker(engine_url, InitOptions(worker_name=worker_name))

    client.register_function("convert", http(handle_convert))
    client.register_trigger({
        "type": "http",
        "function_id": "convert",
        "config": {"api_path": "/convert", "http_method": "POST"},
    })

    log.info("Worker '%s' ready — POST /convert", worker_name)


if __name__ == "__main__":
    main()
