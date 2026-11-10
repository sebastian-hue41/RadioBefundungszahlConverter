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

    def test_the_choice_what_to_evaluate_starts_with_everything(self):
        radios = [attrs for attrs in tags("input") if attrs.get("type") == "radio"]
        assert [attrs["id"] for attrs in radios] == ["scope-all", "scope-nrad", "scope-allrad"]
        assert {attrs["name"] for attrs in radios} == {"scope"}                     # one group: exactly one is chosen
        assert ["checked" in attrs for attrs in radios] == [True, False, False]
        script = main_script()
        assert 'scope: "all", auto: true' in script                               # "Alles" and the automatic mode at the start
        assert 'const SCOPES = { all: ["allrad", "nrad"], nrad: ["nrad"], allrad: ["allrad"] }' in script
        for scope, label in (("all", "Alles"), ("nrad", "Nur NRAD-Codes"), ("allrad", "Nur Allrad-Codes")):
            assert f'id="scope-{scope}" value="{scope}"' in PAGE and f">{label}</label>" in PAGE

    def test_the_only_checkbox_is_the_one_of_the_preview(self):
        boxes = {attrs["id"]: attrs for attrs in tags("input") if attrs.get("type") == "checkbox"}
        assert set(boxes) == {"hide-empty"}
        assert "checked" not in boxes["hide-empty"]                              # the preview shows every row of the Excel at first

    def test_alles_with_one_map_names_the_other_or_the_choice_that_needs_only_this_one(self):
        script = main_script()
        assert 'state.scope === "all" && present.length === 1 ? " – oder " + ONLY[present[0]] + " wählen"' in script
        assert 'const ONLY = { allrad: "„Nur Allrad-Codes“", nrad: "„Nur NRAD-Codes“" }' in script

    def test_the_automatic_mode_sends_every_map_alone_and_takes_all_or_none(self):
        script = main_script()
        assert 'fetch("identify", { method: "POST", body: content })' in script
        assert '!MAPS.includes(reply.map.department)' in script                    # only a department the page knows
        assert 'if (found[slot]) {\n          showError(appError("E222", "maps"' in script                    # two of one kind
        assert script.index("found[slot] = { file, map: reply.map }") < script.index("state.files[slot] = found[slot].file")   # stored only after all were recognised
        assert script.count("state.files[slot] =") == 2                                 # by hand (setFile) and after all were recognised, nowhere else

    def test_the_switch_between_automatic_and_by_hand_keeps_the_maps(self):
        assert '$("mode-switch").addEventListener("click", () => { if (state.busy) return; state.auto = !state.auto; clearError(); render(); });' in main_script()
        assert "Maps selbst zuordnen" in PAGE and "Maps automatisch erkennen" in main_script()

    def test_the_preview_is_a_labelled_scrollable_table_that_the_keyboard_can_reach(self):
        region = next(attrs for attrs in tags("div") if attrs.get("id") == "table-wrap")
        assert region["tabindex"] == "0" and region["role"] == "region" and region["aria-labelledby"] == "preview-title"
        assert [attrs["id"] for attrs in tags("table")] == ["preview-table"]
        assert any(attrs.get("for") == "hide-empty" for attrs in tags("label"))
        assert "Vorschau" in PAGE and "Leere Zeilen ausblenden" in PAGE

    def test_the_quit_button_is_red_and_the_confirmation_too(self):
        buttons = {attrs["id"]: attrs for attrs in tags("button")}
        assert "danger" in buttons["quit"]["class"].split() and "danger" in buttons["quit-yes"]["class"].split()
        css = "".join(PARSED.styles)
        assert re.search(r"\.danger\s*\{[^}]*background:\s*var\(--danger\)", css) and "--danger:" in css and ".danger:focus-visible" in css
        assert "danger" not in buttons["quit-no"].get("class", "")                 # 'Abbrechen' stays neutral

    def test_four_file_pickers_that_only_offer_xlsx_and_only_the_automatic_one_takes_several(self):
        pickers = {attrs["id"]: attrs for attrs in tags("input") if attrs.get("type") == "file"}
        assert set(pickers) == {"file-statistics", "file-maps", "file-allrad", "file-nrad"}
        assert all(attrs.get("accept") == ".xlsx" for attrs in pickers.values())
        assert [name for name, attrs in pickers.items() if "multiple" in attrs] == ["file-maps"]

    def test_all_the_buttons_exist_and_none_submits_by_accident(self):
        buttons = {attrs["id"]: attrs for attrs in tags("button")}
        assert set(buttons) == {"submit", "download", "again", "quit", "quit-yes", "quit-no", "idle-off", "idle-on", "idle-yes", "idle-no", "mode-switch"}
        assert all(attrs.get("type") in ("submit", "button") for attrs in buttons.values())
        assert [i for i, attrs in buttons.items() if attrs["type"] == "submit"] == ["submit"]
        assert "disabled" in buttons["submit"]                                  # nothing to evaluate before files are chosen

    def test_the_texts_the_plan_names(self):
        for text in ("Auswerten", "Excel herunterladen", "Weitere Statistik hochladen", "Server beenden", "Was auswerten?",
                     "Mitarbeiterstatistik", "Leistungsmaps", "Allrad-Map", "NRAD-Map", "Datei wählen oder hierher ziehen",
                     "Map-Dateien wählen oder hierher ziehen", "sie werden automatisch erkannt", "Map wird erkannt …",
                     "Wird ausgewertet …", "Es fehlt noch: ", "Technische Details",
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
            expected = "GET" if route in ("/", "/download", "/status") else "POST"
            assert methods == {expected: methods[expected]}

    def test_the_file_slots_are_the_servers_slots(self):
        assert {attrs["data-slot"] for attrs in tags("div") if "data-slot" in attrs and attrs.get("class") == "field"} == set(SLOT_LABELS)
        assert 'const MAPS = ["allrad", "nrad"]' in main_script()

    def test_every_element_the_script_looks_up_exists(self):
        ids = {attrs["id"] for _, attrs in PARSED.elements if "id" in attrs}
        script = main_script()
        used = set(re.findall(r'\$\("([\w-]+)"\)', script)) | set(re.findall(r'getElementById\("([\w-]+)"\)', script))
        assert used and used <= ids, used - ids
        for prefix in ("field-", "file-", "drop-", "text-", "sub-"):
            for slot in SLOT_LABELS:
                assert prefix + slot in ids, prefix + slot
        for scope in ("all", "nrad", "allrad"):
            assert "scope-" + scope in ids

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

    def test_the_fields_of_the_preview_that_the_page_reads_exist(self):
        preview = evaluate(make_files(), reference_data=load_reference_data(REFERENCE)).summary["preview"]
        read = set(re.findall(r"\bpreview\.(\w+)", main_script()))
        assert read == {"columns", "rows", "total_column", "target_column"} and read <= set(preview)

    def test_the_colours_of_the_total_follow_the_excel(self):
        from core import output_handler
        script, css = main_script(), "".join(PARSED.styles)
        assert "ratio >= 1 ?" in script and "ratio >= 0.8 ?" in script and "ratio >= 0.5 ?" in script        # the thresholds of _add_achievement_cf
        for formula in ("{gf}/{bf}>=1", "{gf}/{bf}>=0.8", "{gf}/{bf}>=0.5", "{gf}/{bf}>0"):
            assert formula in open(output_handler.__file__, encoding="utf-8").read()
        for name, colour in (("green", output_handler._ACH_GREEN), ("yellow", output_handler._ACH_YELLOW), ("red", output_handler._ACH_RED)):
            assert f"--ach-{name}: #{colour.lower()}" in css                                           # the light colours are the Excel's

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


# ── Stopping by itself: the counter and the two-click switch ───────────────────

def handler(element_id):
    """The source of the click handler of one button (every one is a single line)."""
    match = re.search(rf'\$\("{element_id}"\)\.addEventListener\("click", (.*)\);', main_script())
    assert match, element_id
    return match.group(1)


class TestStoppingByItself:

    def test_the_counter_the_warning_and_the_switch_are_in_the_page_and_start_hidden(self):
        elements = {attrs["id"]: attrs for _, attrs in PARSED.elements if "id" in attrs}
        for name in ("idle-box", "idle-confirm", "idle-on", "idle-warn"):
            assert "hidden" in elements[name], name
        assert "hidden" not in elements["idle-off"] and "hidden" not in elements["idle"]            # shown as soon as the box is
        assert elements["idle"]["role"] == "timer" and elements["idle-warn"]["role"] == "alert" and elements["idle-confirm"]["role"] == "alert"

    def test_the_texts_the_user_reads(self):
        for text in ("Beendet sich automatisch in", "Automatisches Beenden ausschalten", "Automatisches Beenden einschalten", "Automatisches Beenden ist aus",
                     "Das Programm beendet sich dann nicht mehr von selbst. Wirklich ausschalten?", "Ja, ausschalten", "Das Programm wurde automatisch beendet.",
                     "Das Programm beendet sich bald von selbst. Jede Eingabe verlängert die Zeit."):
            assert text in PAGE, text

    def test_switching_it_off_takes_two_clicks_the_first_only_asks(self):
        first = handler("idle-off")
        assert "fetch" not in first and "setAutoStop" not in first                      # the first click sends nothing ...
        assert '$("idle-confirm").hidden = false' in first and '$("idle-no").focus()' in first      # ... it asks, with the safe button focused
        assert "setAutoStop(false)" in handler("idle-yes")                              # the second click does it
        assert main_script().count("setAutoStop(false)") == 1                           # and nothing else can
        assert len(re.findall(r'fetch\("autostop', main_script())) == 1

    def test_cancelling_sends_nothing_and_switching_on_again_is_one_click(self):
        assert "fetch" not in handler("idle-no") and "setAutoStop" not in handler("idle-no")
        assert "setAutoStop(true)" in handler("idle-on")

    def test_the_confirmation_of_the_quit_button_and_this_one_never_show_together(self):
        assert '$("confirm").hidden = true' in handler("idle-off")
        assert '$("idle-confirm").hidden = true' in re.search(r'\$\("quit"\)\.addEventListener\("click", \(\) => \{.*?\n  \}\);', main_script(), re.S).group(0)

    def test_asking_for_the_time_left_is_only_a_read_the_use_is_reported_separately(self):
        script = main_script()
        assert 'fetch("status")' in script and 'fetch("status", {' not in script       # a plain GET: it must not count as use
        assert 'fetch("activity", { method: "POST" })' in script

    def test_the_events_that_count_as_use(self):
        events = re.search(r"for \(const type of \[(.*?)\]\) \{\s*window\.addEventListener\(type, activity", main_script(), re.S).group(1)
        assert {e.strip(' "') for e in events.split(",")} == {"pointerdown", "keydown", "wheel", "touchstart", "drop", "change", "scroll"}
        assert "mousemove" not in main_script()                                          # a mouse resting on the page is not use

    def test_the_fields_of_the_replies_the_page_reads_exist(self):
        from local_app.errors import AppError
        read = set(re.findall(r"\breply\.(\w+)", main_script()))
        reply = {"ok", "error", "enabled", "limit", "remaining", "summary", "map"}
        assert read <= reply, read - reply
        payload = set(evaluate_status_payload())
        assert {"enabled", "limit", "remaining"} <= payload and {"enabled", "limit", "remaining", "ok"} == payload

    def test_the_counter_shows_minutes_and_two_digit_seconds(self):
        assert 'Math.floor(seconds / 60) + ":" + String(seconds % 60).padStart(2, "0")' in main_script()      # 29:05, not 29:5

    def test_the_warning_comes_in_the_last_fifth_of_the_time_but_at_most_two_minutes(self):
        assert "Math.min(120, idle.limit * 0.2)" in main_script()

    def test_the_closed_by_itself_screen_names_the_time(self):
        assert 'Math.round((idle.limit || 0) / 60) + " Minuten ohne Eingabe' in main_script()

    def test_a_status_that_says_the_program_is_about_to_stop_closes_the_page_but_one_that_says_it_is_used_elsewhere_does_not(self):
        script = main_script()
        assert "reply.remaining > 2" in script and "closedByItself();" in script

    def test_the_counter_ticks_every_second_and_polls_the_program_only_every_fifteen(self):
        script = main_script()
        assert "setInterval(tickIdle, 1000)" in script and ">= 15000" in script and "idle.lastSent + 5000" in script


def evaluate_status_payload():
    srv = server.create_server()
    try:
        thread = threading.Thread(target=srv.serve_forever, kwargs={"poll_interval": 0.02}, daemon=True)
        thread.start()
        return call(srv, "GET", at(srv, "status")).json
    finally:
        srv.shutdown()
        srv.server_close()
        thread.join(5)
