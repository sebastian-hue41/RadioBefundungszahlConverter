"""
test_web_worker.py

Tests for web-specific functionality:
  - _parse_multipart: pure-Python multipart/form-data parser
  - build_xlsx_bytes: in-memory xlsx pipeline
  - validate_xlsx_bytes: raw-bytes security validation
  - End-to-end: bytes → load → process → build_xlsx_bytes

The iii framework is mocked at the module level so worker.py can be imported
without a running iii engine.
"""

import io
import json
import sys
import zipfile
from unittest.mock import MagicMock

import openpyxl
import pytest

# ── Mock iii before importing web.worker ──────────────────────────────────────
for _mod in ("iii", "iii.types"):
    if _mod not in sys.modules:
        sys.modules[_mod] = MagicMock()

from web.worker import _parse_multipart               # noqa: E402
from core.input_handler import (                       # noqa: E402
    FileValidationError,
    load_map,
    load_reference_amounts,
    load_statistics,
    validate_xlsx_bytes,
)
from core.logic import process                         # noqa: E402
from core.output_handler import build_xlsx_bytes       # noqa: E402
from .helpers import (                                 # noqa: E402
    inject_vba_into_xlsx,
    make_map_wb,
    make_multipart_body,
    make_path_traversal_xlsx_bytes,
    make_reference_wb,
    make_result,
    make_stats_wb,
    make_zip_bomb_xlsx_bytes,
    save_wb,
    wb_to_bytes,
)


# ── TestParseMultipart ─────────────────────────────────────────────────────────

class TestParseMultipart:
    """_parse_multipart — stdlib multipart/form-data parser."""

    def test_single_field_extracted(self):
        body, ct = make_multipart_body({"statistics": b"xlsx-content"})
        fields = _parse_multipart(body, ct)
        assert fields["statistics"] == b"xlsx-content"

    def test_two_fields_extracted(self):
        body, ct = make_multipart_body({"statistics": b"stats", "map": b"mapdata"})
        fields = _parse_multipart(body, ct)
        assert fields["statistics"] == b"stats"
        assert fields["map"] == b"mapdata"

    def test_three_fields_including_reference(self):
        body, ct = make_multipart_body({
            "statistics": b"stats",
            "map":        b"mapdata",
            "reference":  b"refdata",
        })
        fields = _parse_multipart(body, ct)
        assert set(fields.keys()) == {"statistics", "map", "reference"}

    def test_empty_body_returns_empty_dict(self):
        boundary = "emptyboundary"
        body = f"--{boundary}--\r\n".encode()
        ct = f"multipart/form-data; boundary={boundary}"
        fields = _parse_multipart(body, ct)
        assert fields == {}

    def test_binary_payload_preserved(self):
        """Binary xlsx bytes are round-tripped without corruption."""
        wb = make_stats_wb([(1, 1, "01.01.2024", "MRT", "MRT")])
        original = wb_to_bytes(wb)
        body, ct = make_multipart_body({"statistics": original})
        fields = _parse_multipart(body, ct)
        assert fields["statistics"] == original

    def test_field_names_are_exact(self):
        """Field names are extracted without surrounding quotes or whitespace."""
        body, ct = make_multipart_body({"my_field": b"value"})
        fields = _parse_multipart(body, ct)
        assert "my_field" in fields
        assert '"my_field"' not in fields


# ── TestValidateXlsxBytesSecuritySuite ────────────────────────────────────────

class TestValidateXlsxBytesSecuritySuite:
    """Security checks specific to the bytes-upload path."""

    def test_path_traversal_dotdot_blocked(self):
        with pytest.raises(FileValidationError, match="suspicious path"):
            validate_xlsx_bytes(make_path_traversal_xlsx_bytes("../../../etc/passwd"))

    def test_path_traversal_windows_style_blocked(self):
        # Backslash path traversal
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("xl/workbook.xml", "<workbook/>")
            # zipfile normalises backslashes on Windows, use forward slash with ..
            zf.writestr("xl/../../evil.py", "bad")
        with pytest.raises(FileValidationError, match="suspicious path"):
            validate_xlsx_bytes(buf.getvalue())

    def test_path_traversal_absolute_blocked(self):
        with pytest.raises(FileValidationError, match="suspicious path"):
            validate_xlsx_bytes(make_path_traversal_xlsx_bytes("/absolute/path"))

    def test_zip_bomb_high_ratio_blocked(self):
        with pytest.raises(FileValidationError, match="compression ratio"):
            validate_xlsx_bytes(make_zip_bomb_xlsx_bytes())

    def test_vba_upload_blocked(self):
        wb = make_stats_wb([(1, 1, "01.01.2024", "MRT", "MRT")])
        with pytest.raises(FileValidationError, match="VBA"):
            validate_xlsx_bytes(inject_vba_into_xlsx(wb_to_bytes(wb)))

    def test_empty_upload_blocked(self):
        with pytest.raises(FileValidationError, match="empty"):
            validate_xlsx_bytes(b"")

    def test_truncated_zip_blocked(self):
        wb = make_stats_wb([(1, 1, "01.01.2024", "MRT", "MRT")])
        data = wb_to_bytes(wb)
        with pytest.raises(FileValidationError):
            validate_xlsx_bytes(data[:50])   # truncate to corrupt the zip

    def test_valid_upload_accepted(self):
        wb = make_stats_wb([(1, 1, "01.01.2024", "MRT", "MRT")])
        validate_xlsx_bytes(wb_to_bytes(wb))   # must not raise


# ── TestBuildXlsxBytesIntegration ─────────────────────────────────────────────

class TestBuildXlsxBytesIntegration:
    """End-to-end: bytes in → process → bytes out, entirely in memory."""

    def _run_pipeline(self, stats_rows, map_codes, ref_amounts=None):
        """Helper: load from bytes, process, return (result, xlsx_bytes)."""
        stats_bytes = wb_to_bytes(make_stats_wb(stats_rows))
        map_bytes   = wb_to_bytes(make_map_wb(map_codes))

        stats_data = load_statistics(io.BytesIO(stats_bytes))
        map_data   = load_map(io.BytesIO(map_bytes))
        result     = process(stats_data, map_data)

        ref = None
        if ref_amounts is not None:
            ref_bytes = wb_to_bytes(make_reference_wb(ref_amounts))
            ref = load_reference_amounts(io.BytesIO(ref_bytes))

        filename, xlsx_bytes = build_xlsx_bytes(result, reference_amounts=ref)
        return result, filename, xlsx_bytes

    def test_basic_in_memory_pipeline(self):
        rows = [(1, 1, "01.01.2024", "MRT Schädel", "MRTS")]
        codes = [("MRTS", "MRT Schädel", {"DL Gefäße": 0, "MRT": 1, "CT": 0})]
        result, filename, data = self._run_pipeline(rows, codes)
        assert isinstance(data, bytes)
        assert len(data) > 0
        wb = openpyxl.load_workbook(io.BytesIO(data))
        assert wb.active is not None

    def test_counts_correct_in_output_bytes(self):
        rows = [
            (1, 1, "01.01.2024", "MRT Schädel", "MRTS"),
            (2, 1, "01.01.2024", "MRT Schädel", "MRTS"),
            (3, 1, "01.01.2024", "MRT Schädel", "MRTS"),
        ]
        codes = [("MRTS", "MRT Schädel", {"DL Gefäße": 0, "MRT": 1, "CT": 0})]
        result, _, data = self._run_pipeline(rows, codes)

        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        ws = wb.active
        header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        gesamt_idx = header.index("Gesamt") + 1
        mrt_row = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value == "MRT")
        assert ws.cell(mrt_row, gesamt_idx).value == 3

    def test_reference_amounts_via_bytes(self):
        rows = [(1, 1, "01.01.2024", "MRT Schädel", "MRTS")]
        codes = [("MRTS", "MRT Schädel", {"DL Gefäße": 0, "MRT": 1, "CT": 0})]
        ref = {"MRT": 5000}
        _, _, data = self._run_pipeline(rows, codes, ref_amounts=ref)

        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        ws = wb.active
        header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        assert "Benötigt" in header
        benoetigt_idx = header.index("Benötigt") + 1
        mrt_row = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value == "MRT")
        assert ws.cell(mrt_row, benoetigt_idx).value == 5000

    def test_val2_propagates_through_full_bytes_pipeline(self):
        """val=2 codes double the section count end-to-end through in-memory loading."""
        rows = [
            (1, 1, "01.01.2024", "CT Thorax-Abdomen", "CTTHA"),
            (2, 1, "01.01.2024", "CT Thorax-Abdomen", "CTTHA"),
        ]
        codes = [("CTTHA", "CT Thorax-Abdomen", {"DL Gefäße": 0, "MRT": 0, "CT": 2})]
        result, _, data = self._run_pipeline(rows, codes)

        assert result["data"]["CT"]["Total"] == 4  # 2 exams × val 2

        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        ws = wb.active
        header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        gesamt_idx = header.index("Gesamt") + 1
        # Find CT row
        ct_row = next(
            r for r in range(2, ws.max_row + 1)
            if ws.cell(r, 1).value == "CT"
        )
        assert ws.cell(ct_row, gesamt_idx).value == 4

    def test_no_files_created_on_disk(self, tmp_path, monkeypatch):
        """The full in-memory pipeline writes nothing to the filesystem."""
        written = []
        original_save = openpyxl.Workbook.save

        def _spy(self, filename):
            if not isinstance(filename, io.BytesIO):
                written.append(str(filename))
            return original_save(self, filename)

        monkeypatch.setattr(openpyxl.Workbook, "save", _spy)

        rows = [(1, 1, "01.01.2024", "MRT", "MRTS")]
        codes = [("MRTS", "d", {"DL Gefäße": 0, "MRT": 1, "CT": 0})]
        self._run_pipeline(rows, codes)
        assert written == [], f"Unexpected disk writes: {written}"

    def test_multiple_years_in_output(self):
        rows = [
            (1, 1, "01.01.2023", "MRT", "MRTS"),
            (2, 1, "01.01.2024", "MRT", "MRTS"),
            (3, 1, "01.01.2025", "MRT", "MRTS"),
        ]
        codes = [("MRTS", "MRT Schädel", {"DL Gefäße": 0, "MRT": 1, "CT": 0})]
        _, _, data = self._run_pipeline(rows, codes)

        wb = openpyxl.load_workbook(io.BytesIO(data))
        ws = wb.active
        header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        assert "2023" in header
        assert "2024" in header
        assert "2025" in header
