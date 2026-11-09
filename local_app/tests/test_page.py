"""
Tests for local_app/static/index.html: what the security policy would silently block, what the page
promises, and that it agrees with the server and the catalog. How it behaves in a browser is checked by
local_app/tests/browser_pass.py (needs Playwright and a browser, so pytest does not collect it).
"""

import json
import re
import shutil
import subprocess
from html.parser import HTMLParser

import pytest

from local_app import __version__, server
from local_app.errors import CATALOG, HINTS, SLOT_LABELS
from local_app.server import _Handler, bundle_dir, catalog_json, evaluate

from .test_server import REFERENCE, at, call, load_reference_data, make_files, threading

PAGE_PATH = bundle_dir() / "local_app" / "static" / "index.html"
PAGE = PAGE_PATH.read_text(encoding="utf-8")


class Parsed(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.elements = []                    # (tag, {attribute: value})
        self._stack, self.scripts, self.styles = [], [], []

    def handle_starttag(self, tag, attrs):
        self.elements.append((tag, dict(attrs)))
        self._stack.append(tag)

    def handle_endtag(self, tag):
        if self._stack and self._stack[-1] == tag:
            self._stack.pop()

    def handle_startendtag(self, tag, attrs):
        self.elements.append((tag, dict(attrs)))

    def handle_data(self, data):
        if self._stack and self._stack[-1] == "script":
            self.scripts.append(data)
        elif self._stack and self._stack[-1] == "style":
            self.styles.append(data)


def parse(html=PAGE):
    parser = Parsed()
    parser.feed(html)
    return parser


PARSED = parse()


def tags(name):
    return [attrs for tag, attrs in PARSED.elements if tag == name]


def main_script():
    return next(text for text in PARSED.scripts if "use strict" in text)


def served_page(srv):
    return call(srv, "GET", at(srv)).body.decode("utf-8")


@pytest.fixture
def app():
    srv = server.create_server(load_reference_data(REFERENCE))
    thread = threading.Thread(target=srv.serve_forever, kwargs={"poll_interval": 0.02}, daemon=True)
    thread.start()
    yield srv
    srv.shutdown()
    srv.server_close()
    thread.join(5)


# ── What the policy would block, and what must not be in the page ──────────────

class TestNothingThePolicyWouldBlock:

    def test_no_address_to_anywhere_so_it_works_offline_and_sends_nothing_elsewhere(self):
        assert not re.search(r"(?i)\b(https?|ftp|wss?)://", PAGE)
        assert "//cdn" not in PAGE and "@import" not in PAGE and "url(" not in PAGE

    def test_no_external_or_embedded_things(self):
        for forbidden in ("link", "img", "iframe", "frame", "object", "embed", "form-external", "base", "svg", "video", "audio", "source", "meta-refresh"):
            assert not tags(forbidden), forbidden
        assert all("src" not in attrs for tag, attrs in PARSED.elements)
        assert not any(tag == "meta" and attrs.get("http-equiv") for tag, attrs in PARSED.elements)

    @pytest.mark.parametrize("word", ["innerHTML", "outerHTML", "insertAdjacentHTML", "document.write", "createContextualFragment", "srcdoc", "setAttribute(\"style\"", "setAttribute('style'"])
    def test_nothing_that_inserts_markup(self, word):
        assert word not in PAGE

    def test_no_eval_and_no_code_from_strings(self):
        script = main_script()
        assert not re.search(r"\beval\s*\(", script) and not re.search(r"\bnew\s+Function\b", script)
        assert not re.search(r"\bFunction\s*\(", script)
        assert not re.search(r"set(Timeout|Interval)\s*\(\s*[\"'`]", script)             # a timer with a string is eval in disguise
        assert "importScripts" not in script and "import(" not in script and "javascript:" not in PAGE.lower()

    def test_no_inline_event_handlers(self):
        for tag, attrs in PARSED.elements:
            assert not [name for name in attrs if name.startswith("on")], (tag, attrs)

    def test_no_inline_style_attributes(self):
        # style-src has no 'unsafe-inline': a style attribute would be dropped silently
        assert not [attrs for tag, attrs in PARSED.elements if "style" in attrs]

    def test_the_script_does_not_set_styles_through_the_element(self):
        assert not re.search(r"\.style\.\w+\s*=", main_script()) and ".style.cssText" not in main_script()

    def test_every_script_and_style_carries_the_nonce(self):
        scripts, styles = tags("script"), tags("style")
        assert len(scripts) == 2 and len(styles) == 1
        for attrs in scripts + styles:
            assert attrs.get("nonce") == "{{NONCE}}"
        assert not [attrs for attrs in scripts if "src" in attrs]

    def test_the_catalog_is_a_data_block_not_code(self):
        block = next(attrs for attrs in tags("script") if attrs.get("id") == "catalog")
        assert block["type"] == "application/json"

    def test_the_page_is_not_huge(self):
        assert len(PAGE.encode("utf-8")) < 60_000

    def test_the_javascript_is_syntactically_valid(self, tmp_path):
        node = shutil.which("node")
        if node is None:
            pytest.skip("node is not installed")
        path = tmp_path / "page.js"
        path.write_text(main_script(), encoding="utf-8")
        done = subprocess.run([node, "--check", str(path)], capture_output=True, text=True, timeout=30)
        assert done.returncode == 0, done.stderr


# ── What the page promises ─────────────────────────────────────────────────────

class TestWhatThePageOffers:

    def test_it_is_german_and_has_one_heading_one(self):
        assert tags("html")[0]["lang"] == "de" and len(tags("h1")) == 1
        assert "Befundungszahlen auswerten" in PAGE

    def test_the_two_checkboxes_both_ticked_to_start_with(self):
        boxes = {attrs["id"]: attrs for attrs in tags("input") if attrs.get("type") == "checkbox"}
        assert set(boxes) == {"use-allrad", "use-nrad"}
        assert all("checked" in attrs for attrs in boxes.values())

    def test_three_file_pickers_that_only_offer_xlsx(self):
        pickers = [attrs for attrs in tags("input") if attrs.get("type") == "file"]
        assert {attrs["id"] for attrs in pickers} == {"file-statistics", "file-allrad", "file-nrad"}
        assert all(attrs.get("accept") == ".xlsx" for attrs in pickers)
        assert not any("multiple" in attrs for attrs in pickers)

    def test_all_the_buttons_exist_and_none_submits_by_accident(self):
        buttons = {attrs["id"]: attrs for attrs in tags("button")}
        assert set(buttons) == {"submit", "download", "again", "quit", "quit-yes", "quit-no"}
        assert all(attrs.get("type") in ("submit", "button") for attrs in buttons.values())
        assert [i for i, attrs in buttons.items() if attrs["type"] == "submit"] == ["submit"]
        assert "disabled" in buttons["submit"]                                  # nothing to evaluate before files are chosen

    def test_the_texts_the_plan_names(self):
        for text in ("Auswerten", "Excel herunterladen", "Weitere Statistik hochladen", "Server beenden", "Welche Bereiche?",
                     "Mitarbeiterstatistik", "Leistungsmaps", "Allrad-Map", "NRAD-Map", "Datei wählen oder hierher ziehen",
                     "Wird ausgewertet …", "Es fehlt noch: ", "Bitte mindestens einen Bereich ankreuzen.", "Technische Details",
                     "Fehlercode", "Auswertung fertig", "Das Programm wurde beendet."):
            assert text in PAGE, text

    def test_every_input_has_a_label(self):
        label_for = {attrs["for"] for attrs in tags("label") if "for" in attrs}
        for attrs in tags("input"):
            assert attrs["id"] in label_for or "aria-labelledby" in attrs, attrs
        # a label that names the file picker must point at it and at an element that exists
        ids = {attrs["id"] for _, attrs in PARSED.elements if "id" in attrs}
        for attrs in tags("input"):
            for target in attrs.get("aria-labelledby", "").split():
                assert target in ids, (attrs["id"], target)

    def test_a_file_picker_is_named_after_its_slot_first(self):
        # the label of the drop area alone would only say "chosen file ...": the slot name comes first
        for attrs in tags("input"):
            if attrs.get("type") == "file":
                slot = attrs["id"].removeprefix("file-")
                assert attrs["aria-labelledby"].split()[0] == f"h-{slot}", attrs

    def test_the_error_box_is_an_alert_that_can_take_focus(self):
        box = next(attrs for attrs in tags("div") if attrs.get("id") == "error")
        assert box["role"] == "alert" and box["tabindex"] == "-1" and "hidden" in box

    def test_the_status_next_to_the_button_is_announced(self):
        status = next(attrs for attrs in tags("p") if attrs.get("id") == "status")
        assert status["role"] == "status" and status["aria-live"] == "polite"

    def test_ids_are_unique(self):
        ids = [attrs["id"] for _, attrs in PARSED.elements if "id" in attrs]
        assert len(ids) == len(set(ids))

    def test_there_is_a_message_for_a_browser_without_javascript(self):
        assert tags("noscript") and "JavaScript" in PAGE

    def test_it_fits_small_windows_and_both_colour_schemes(self):
        assert any(attrs.get("name") == "viewport" and "width=device-width" in attrs["content"] for attrs in tags("meta"))
        css = "".join(PARSED.styles)
        assert "prefers-color-scheme: dark" in css and "max-width: 30rem" in css
        assert any(attrs.get("name") == "color-scheme" and attrs["content"] == "light dark" for attrs in tags("meta"))

    def test_the_focus_is_visible_everywhere(self):
        css = "".join(PARSED.styles)
        assert ":focus-within" in css and ":focus-visible" in css and "outline" in css


# ── The page and the server agree ──────────────────────────────────────────────

class TestPageAndServerAgree:

    def test_the_upload_parameters_are_the_ones_the_server_expects(self):
        script = main_script()
        assert set(server._SIZE_PARAMS.values()) == {"stats", "allrad", "nrad"}
        assert '"stats"' in script and 'slot === "statistics" ? "stats" : slot' in script
        assert 'slot + "_name"' in script
        assert set(server._NAME_PARAMS.values()) == {"allrad_name", "nrad_name"}

    def test_the_page_only_calls_routes_the_server_has(self):
        routes = {route.strip("/") for route in _Handler._ROUTES} - {""}
        called = set(re.findall(r'fetch\("([a-z]+)', main_script()))
        assert called == routes

    def test_the_methods_the_page_uses_are_the_ones_the_routes_accept(self):
        script = main_script()
        assert 'fetch("evaluate?" + query.toString(), { method: "POST"' in script
        assert 'fetch("reset", { method: "POST" })' in script and 'fetch("shutdown", { method: "POST" })' in script
        assert 'fetch("download")' in script
        for route, methods in _Handler._ROUTES.items():
            expected = "GET" if route in ("/", "/download") else "POST"
            assert methods == {expected: methods[expected]}

    def test_the_file_slots_are_the_servers_slots(self):
        assert {attrs["data-slot"] for attrs in tags("div") if "data-slot" in attrs and attrs.get("class") == "field"} == set(SLOT_LABELS)
        assert 'const MAPS = ["allrad", "nrad"]' in main_script()

    def test_every_element_the_script_looks_up_exists(self):
        ids = {attrs["id"] for _, attrs in PARSED.elements if "id" in attrs}
        script = main_script()
        used = set(re.findall(r'\$\("([\w-]+)"\)', script)) | set(re.findall(r'getElementById\("([\w-]+)"\)', script))
        assert used and used <= ids, used - ids
        for prefix in ("field-", "file-", "drop-", "text-", "sub-", "use-"):
            for slot in (SLOT_LABELS if prefix not in ("use-",) else ("allrad", "nrad")):
                assert prefix + slot in ids, prefix + slot

    def test_every_element_with_a_data_hook_in_the_script_exists_in_the_page(self):
        script = main_script()
        css_classes = set(re.findall(r'classList\.(?:add|remove|toggle|contains)\("([\w-]+)"', script)) | set(re.findall(r'\.className = "([\w-]+)"', script))
        defined = set(re.findall(r"\.([a-zA-Z][\w-]*)", "".join(PARSED.styles)))
        assert css_classes <= defined | {"over", "invalid", "has-file"}, css_classes - defined

    def test_the_codes_the_page_makes_up_itself_are_in_the_catalog(self):
        used = set(re.findall(r'appError\("(E\d{3})"', main_script())) | set(re.findall(r'"(E\d{3})"', main_script()))
        assert {"E301", "E302", "E306", "E902", "E103", "E203", "E305"} <= used
        assert used <= set(CATALOG), used - set(CATALOG)

    def test_the_fields_of_an_error_that_the_page_reads_exist(self):
        from local_app.errors import AppError
        shown = AppError("E210", "nrad", "details").to_dict()
        read = set(re.findall(r"\berror\.(\w+)", main_script())) - {"code"} | {"code"}
        assert read <= set(shown), read - set(shown)

    def test_the_fields_of_the_summary_that_the_page_reads_exist(self):
        summary = evaluate(make_files(), reference_data=load_reference_data(REFERENCE)).summary
        read = set(re.findall(r"\bsummary\.(\w+)", main_script())) - {"hints"}
        assert read and read <= set(summary), read - set(summary)
        assert "hints" in summary

    def test_the_hint_fields_the_page_reads_exist(self):
        from .test_server import exams
        hints = evaluate(make_files(exams("RMRT", "RUNBEKANNT"), nrad=False)).summary["hints"]
        assert hints and set(re.findall(r"\bentry\.(\w+)", main_script())) <= set(hints[0])


# ── The catalog in the served page ─────────────────────────────────────────────

def catalog_block(html):
    match = re.search(r'<script type="application/json" id="catalog" nonce="[^"]+">(.*?)</script>', html, re.S)
    assert match, "no catalog block in the page"
    return match.group(1)


class TestCatalog:

    def test_the_served_page_carries_the_whole_catalog(self, app):
        data = json.loads(catalog_block(served_page(app)))
        assert set(data) == {"errors", "hints", "slots", "limits", "version"}
        assert set(data["errors"]) == set(CATALOG)
        assert all(data["errors"][code] == {"title": entry.title, "hint": entry.hint} for code, entry in CATALOG.items())
        assert data["hints"] == {code: list(texts) for code, texts in HINTS.items()}
        assert data["slots"] == SLOT_LABELS
        assert data["limits"] == {"file_bytes": 50 * 1024 * 1024, "body_bytes": server.MAX_BODY_BYTES}
        assert data["version"] == __version__

    def test_the_block_has_no_character_that_could_end_it(self, app):
        block = catalog_block(served_page(app))
        assert not {"<", ">", "&"} & set(block)
        assert not {"<", ">", "&"} & set(catalog_json())

    def test_text_in_the_catalog_cannot_break_out_of_the_page(self, app, monkeypatch):
        hostile = '</script><script>alert(1)</script><img src=x onerror=alert(2)> & "quotes"  '
        monkeypatch.setitem(CATALOG, "E999", type(CATALOG["E999"])("E999", hostile, hostile))
        monkeypatch.setitem(HINTS, "W900", (hostile, hostile))
        html = served_page(app)
        parsed = parse(html)
        assert len(parsed.scripts) == 2 and not [tag for tag, _ in parsed.elements if tag == "img"]
        assert html.count("</script>") == 2 and "<img" not in html and "alert(1)" not in html.replace("\\u003cscript\\u003ealert(1)", "")
        block = catalog_block(html)
        assert block.isascii() and not {"<", ">", "&"} & set(block)                           # plain ASCII, nothing that could end the block
        assert json.loads(block)["errors"]["E999"]["title"] == hostile                         # and it still arrives intact

    def test_every_script_and_style_of_the_served_page_has_the_same_fresh_nonce_as_the_policy(self, app):
        response = call(app, "GET", at(app))
        nonce = re.search(r"script-src 'nonce-([^']+)'", response.headers["content-security-policy"]).group(1)
        html = response.body.decode("utf-8")
        assert re.findall(r'<(?:script|style)[^>]* nonce="([^"]+)"', html) == [nonce] * 3

    def test_the_version_label_is_the_one_in_the_catalog(self):
        assert '"Version " + catalog.version' in main_script()
        assert __version__ == "0.1.0"

    def test_a_page_without_the_catalog_placeholder_still_works(self):
        s = server.create_server(page_html="<p nonce='{{NONCE}}'>x</p>")
        s.server_close()

    def test_nothing_of_the_placeholders_is_left(self, app):
        html = served_page(app)
        assert "{{NONCE}}" not in html and "{{CATALOG}}" not in html
