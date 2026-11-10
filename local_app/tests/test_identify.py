"""
Tests for the automatic mode: identify_map() and the route /identify. The page sends every map the user
drops into its one field alone; the program says whether it is an Allrad or an NRAD map, or why it is no
usable map (E2xx for the field "maps"). The rule is the one of the swapped-map guard (scope.department_of).
"""

import threading
import time

import pytest

from core.input_handler import load_map, load_reference_data
from local_app import server
from local_app.errors import ALLRAD, MAPS, NRAD, SLOT_LABELS, STATISTICS, AppError
from local_app.scope import check_department, department_of
from local_app.server import MAX_MAP_BYTES, identify_map
from tests.helpers import inject_vba_into_xlsx

from .helpers import force_text, map_bytes, stats_bytes, with_doctype
from .test_server import (
    ALLRAD_CODES,
    MIXED,
    NRAD_CODES,
    REFERENCE,
    SECTIONS,
    BlockedEvaluation,
    assert_secure_headers,
    at,
    call,
    make_files,
    post_files,
)

ALLRAD_MAP = map_bytes(ALLRAD_CODES, sections=SECTIONS)
NRAD_MAP = map_bytes(NRAD_CODES, sections=SECTIONS)


def refused(data) -> AppError:
    with pytest.raises(AppError) as info:
        identify_map(data)
    return info.value


# ── The rule ───────────────────────────────────────────────────────────────────

class TestTheRule:

    def test_the_two_maps_are_told_apart(self):
        assert identify_map(ALLRAD_MAP) == {"department": ALLRAD, "label": "Allrad-Map", "codes": 3, "sections": 3}
        assert identify_map(NRAD_MAP) == {"department": NRAD, "label": "NRAD-Map", "codes": 2, "sections": 3}

    @pytest.mark.parametrize("n_codes, others, expected", [(1, 1, NRAD), (2, 1, NRAD), (1, 2, ALLRAD), (0, 5, ALLRAD), (5, 0, NRAD)])
    def test_an_nrad_map_has_at_least_half_n_codes(self, n_codes, others, expected):
        codes = [(f"N{i}", "n", {"MRT": 1}) for i in range(n_codes)] + [(f"R{i}", "r", {"MRT": 1}) for i in range(others)]
        assert identify_map(map_bytes(codes, sections=SECTIONS))["department"] == expected
        assert department_of(load_map(map_bytes(codes, sections=SECTIONS))) == (expected, n_codes, n_codes + others)

    @pytest.mark.parametrize("data", [ALLRAD_MAP, NRAD_MAP])
    def test_it_agrees_with_the_swapped_map_guard(self, data):
        department = identify_map(data)["department"]
        check_department(load_map(data), department)                    # the field it is sorted into is accepted
        with pytest.raises(AppError) as info:
            check_department(load_map(data), NRAD if department == ALLRAD else ALLRAD)
        assert info.value.code == "E220"

    def test_a_map_sorted_by_the_rule_evaluates(self):
        files = {STATISTICS: stats_bytes(MIXED, total_l8=len(MIXED))}
        for data in (NRAD_MAP, ALLRAD_MAP):                              # in any order
            files[identify_map(data)["department"]] = data
        assert server.evaluate(files).summary["selection"] == [ALLRAD, NRAD]


# ── What is refused, for the field "maps" ──────────────────────────────────────

class TestWhatIsNoMap:

    @pytest.mark.parametrize("data, code", [
        (stats_bytes(MIXED), "E210"),                                    # the statistics dropped among the maps
        (b"this is not an excel file", "E201"),
        (b"", "E202"),
        (b"x" * (MAX_MAP_BYTES + 1), "E203"),
        (inject_vba_into_xlsx(ALLRAD_MAP), "E204"),
        (with_doctype(ALLRAD_MAP, '<!DOCTYPE x [<!ENTITY e "boese">]>'), "E205"),
        (map_bytes(ALLRAD_CODES, sections=["DL Gefäße", "MRT", "MRT"]), "E213"),
    ])
    def test_the_code_of_the_problem_for_the_field_maps(self, data, code):
        error = refused(data)
        assert error.code == code and error.slot == MAPS
        assert error.to_dict()["slot_label"] == SLOT_LABELS[MAPS] == "Leistungsmap"

    def test_the_file_is_checked_first_and_named_after_its_field(self):
        # scan_xlsx runs before load_map: core names the file by the field, not by a made-up "map.xlsx"
        detail = refused(inject_vba_into_xlsx(ALLRAD_MAP)).detail
        assert detail.endswith(": Leistungsmap") and "map.xlsx" not in detail

    def test_the_limit_is_the_one_of_a_single_file(self):
        from core.input_handler import MAX_FILE_SIZE_MB
        assert MAX_MAP_BYTES == MAX_FILE_SIZE_MB * 1024 * 1024 == 50 * 1024 * 1024
        assert refused(b"x" * (MAX_MAP_BYTES + 1)).code == "E203"

    def test_a_value_that_is_no_whole_number(self):
        from tests.helpers import make_map_wb, wb_to_bytes
        wb = make_map_wb(ALLRAD_CODES, sections=SECTIONS)
        wb.active["G2"] = 1.5
        error = refused(wb_to_bytes(wb))
        assert error.code == "E214" and "G2 = 1.5" in error.detail

    def test_a_formula_as_a_section_name(self):
        from tests.helpers import make_map_wb, wb_to_bytes
        wb = make_map_wb(ALLRAD_CODES, sections=SECTIONS)
        force_text(wb.active["G1"], "=HYPERLINK(1)")
        assert refused(wb_to_bytes(wb)).code == "E205"

    def test_something_nobody_foresaw_is_e999_without_the_message(self, monkeypatch):
        def broken(map_data):
            raise KeyError("Patient Mustermann")
        monkeypatch.setattr(server, "department_of", broken)
        error = refused(ALLRAD_MAP)
        assert error.code == "E999" and "Mustermann" not in error.detail and "KeyError" in error.detail

    def test_only_bytes_are_taken(self):
        with pytest.raises(ValueError):
            identify_map("not bytes")


# ── The route ──────────────────────────────────────────────────────────────────

@pytest.fixture
def app():
    srv = server.create_server(load_reference_data(REFERENCE))
    thread = threading.Thread(target=srv.serve_forever, kwargs={"poll_interval": 0.02}, daemon=True)
    thread.start()
    yield srv
    srv.shutdown()
    srv.server_close()
    thread.join(5)


def identify(app, body, **kwargs):
    return call(app, "POST", at(app, "identify"), body, **kwargs)


class TestTheRoute:

    def test_it_answers_with_the_department(self, app):
        for data, department in ((ALLRAD_MAP, ALLRAD), (NRAD_MAP, NRAD)):
            response = identify(app, data)
            assert response.status == 200 and response.json["ok"] is True
            assert response.json["map"]["department"] == department and response.json["map"]["label"] == SLOT_LABELS[department]
            assert response.headers["content-type"] == "application/json"
            assert_secure_headers(response)

    def test_a_file_that_is_no_map_is_an_error_for_the_field_maps(self, app):
        response = identify(app, stats_bytes(MIXED))
        assert response.status == 422 and response.error["code"] == "E210" and response.error["slot"] == MAPS

    def test_a_file_over_the_limit_is_refused_before_it_is_read(self, app):
        started = time.perf_counter()
        response = identify(app, None, length=MAX_MAP_BYTES + 1)          # announced, never sent
        assert time.perf_counter() - started < 2
        assert response.error["code"] == "E203" and response.error["slot"] == MAPS

    def test_without_a_length_or_with_a_wrong_one(self, app):
        assert identify(app, None, length=None).error["code"] == "E304"
        assert identify(app, b"abc", length="-3").error["code"] == "E304"

    def test_only_post(self, app):
        assert call(app, "GET", at(app, "identify")).status == 405

    def test_another_site_cannot_use_it(self, app):
        assert identify(app, ALLRAD_MAP, headers={"Origin": "http://example.org"}).status == 403
        assert call(app, "POST", "/wrongtoken/identify", ALLRAD_MAP).status == 404

    def test_it_waits_for_no_evaluation_but_says_so(self, app, monkeypatch):
        blocked = BlockedEvaluation(monkeypatch)
        thread, answer = blocked.start(app)
        try:
            response = identify(app, ALLRAD_MAP)
            assert response.status == 409 and response.error["code"] == "E307"
        finally:
            blocked.release.set()
            thread.join(10)
        assert identify(app, ALLRAD_MAP).status == 200                     # free again

    def test_it_leaves_the_result_alone(self, app):
        assert post_files(app, make_files()).status == 200
        assert identify(app, NRAD_MAP).status == 200
        assert identify(app, b"broken").status == 422
        assert call(app, "GET", at(app, "download")).status == 200

    def test_it_counts_as_use(self, app):
        app.state._last_activity -= 100
        before = app.state.remaining()
        identify(app, ALLRAD_MAP)
        assert app.state.remaining() > before
