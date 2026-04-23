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
        Response 422: application/json  {"error": "..."}  — validation failure
        Response 500: application/json  {"error": "..."}  — unexpected server error

The output xlsx is built entirely in memory — no files are written to disk.
"""

from __future__ import annotations

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


# ── Multipart parsing ──────────────────────────────────────────────────────────

def _parse_multipart(body: bytes, content_type: str) -> dict[str, bytes]:
    """
    Parse a multipart/form-data body and return {field_name: file_bytes}.
    Uses the stdlib email parser — no extra dependencies.
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
    await res.headers({"Content-Type": "application/json", "Content-Length": str(len(body))})
    res.writer.stream.end(body)


# ── HTTP handler ───────────────────────────────────────────────────────────────

async def handle_convert(req: HttpRequest, res: HttpResponse) -> None:
    """
    POST /convert — receive two xlsx uploads, run the converter, stream result back.
    """
    try:
        # ── Read and parse the request body ───────────────────────────────────
        body = await req.request_body.read_all()

        content_type = req.headers.get("content-type", "")
        if isinstance(content_type, list):
            content_type = content_type[0]

        if "multipart/form-data" not in content_type:
            await _send_json_error(res, 400, "Expected multipart/form-data with 'statistics' and 'map' fields.")
            return

        fields = _parse_multipart(body, content_type)
        stats_bytes = fields.get("statistics")
        map_bytes = fields.get("map")

        if not stats_bytes:
            await _send_json_error(res, 400, "Missing required field: 'statistics'.")
            return
        if not map_bytes:
            await _send_json_error(res, 400, "Missing required field: 'map'.")
            return

        # ── Optional query params ──────────────────────────────────────────────
        include_underscore_raw = req.query_params.get("include_underscore", "false")
        if isinstance(include_underscore_raw, list):
            include_underscore_raw = include_underscore_raw[0]
        include_underscore = include_underscore_raw.lower() == "true"

        # ── Core processing (no disk I/O) ──────────────────────────────────────
        stats_data = load_statistics(io.BytesIO(stats_bytes))
        map_data = load_map(io.BytesIO(map_bytes), include_underscore_columns=include_underscore)
        result = process(stats_data, map_data)
        filename, xlsx_bytes = build_xlsx_bytes(result)

        # ── Stream the response ────────────────────────────────────────────────
        await res.status(200)
        await res.headers({
            "Content-Type": _XLSX_CONTENT_TYPE,
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Content-Length": str(len(xlsx_bytes)),
        })

        writer = res.writer
        for offset in range(0, len(xlsx_bytes), _CHUNK_SIZE):
            await writer.write(xlsx_bytes[offset : offset + _CHUNK_SIZE])
        writer.close()

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
