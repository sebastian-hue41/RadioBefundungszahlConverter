"""
Browser pass for the page of the local app: drives the real program in a real browser and checks, one
by one, what the plan promises (a complete run, the line next to "Auswerten", drag and drop, the busy
state, every error code with its field, E306, hostile text, E902, quitting, light/dark, a small window,
keyboard only, and that the browser reports no Content-Security-Policy violation).

Not collected by pytest: it needs Playwright and a browser, which are not dependencies of the project.

    pip install playwright openpyxl        # in a scratch environment, not in the project's
    python local_app/tests/browser_pass.py [chromium|webkit|firefox] [--shots FOLDER]

Environment: BROWSER_EXECUTABLE = a Chromium-based browser to use instead of Playwright's own (for example
Brave or Chrome). Everything is generated here (made-up workbooks) and kept in a temporary folder.
"""
import atexit, base64, io, json, os, queue, re, shutil, subprocess, sys, tempfile, threading, time, zipfile
from contextlib import contextmanager
from pathlib import Path

ROOT = str(Path(__file__).resolve().parents[2])
sys.path.insert(0, ROOT)
WORK = tempfile.mkdtemp(prefix="local_app_browser_pass_")
atexit.register(shutil.rmtree, WORK, True)       # the made-up workbooks and the program's working folder
S = WORK                                   # screenshots go here unless --shots says otherwise (and go with the folder)
EXECUTABLE = os.environ.get("BROWSER_EXECUTABLE")


class App:
    """The program, started as a user starts it: python -m local_app.main --no-browser, from an empty folder."""

    def __init__(self):
        cwd = os.path.join(WORK, "cwd"); os.makedirs(cwd, exist_ok=True)
        self.proc = subprocess.Popen([sys.executable, "-B", "-m", "local_app.main", "--no-browser"], cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                     env={**os.environ, "PYTHONPATH": ROOT, "PYTHONDONTWRITEBYTECODE": "1"})
        self.lines = queue.Queue(); threading.Thread(target=lambda: [self.lines.put(l) for l in self.proc.stdout], daemon=True).start()
        self.url = None; end = time.time() + 20
        while self.url is None and time.time() < end:
            try: l = self.lines.get(timeout=0.5)
            except queue.Empty: continue
            if b"running at: " in l: self.url = l.split(b"running at: ")[1].strip().decode()
        assert self.url, "the program printed no address"

    def stop(self):
        if self.proc.poll() is None:
            self.proc.terminate()
            try: self.proc.wait(5)
            except Exception: self.proc.kill()


@contextmanager
def browser(p, engine):
    if engine == "chromium" and EXECUTABLE: b = p.chromium.launch(executable_path=EXECUTABLE, headless=True)
    else: b = getattr(p, engine).launch(headless=True)
    try: yield b
    finally: b.close()

import openpyxl
from playwright.sync_api import sync_playwright
from local_app.errors import ALLRAD, NRAD, STATISTICS
from local_app.tests.helpers import map_bytes, stats_bytes, with_doctype
from local_app.tests.test_server import ALLRAD_CODES, NRAD_CODES, SECTIONS, MIXED, exams


args = [a for a in sys.argv[1:] if not a.startswith("--")]
ENGINE = args[0] if args else "chromium"
if "--shots" in sys.argv: S = sys.argv[sys.argv.index("--shots") + 1]; os.makedirs(S, exist_ok=True)
F = f"{WORK}/files"; os.makedirs(F, exist_ok=True)
def put(name, data): open(f"{F}/{name}", "wb").write(data); return f"{F}/{name}"
rows = lambda r: dict(total_l8=len(r), mitarbeiter="ZUFALL_N", befunddatum="01.04.2024-31.12.2024")
stats = put("statistik.xlsx", stats_bytes(MIXED, **rows(MIXED)))
allrad = put("map_allrad.xlsx", map_bytes(ALLRAD_CODES, sections=SECTIONS))
nrad = put("map_nrad.xlsx", map_bytes(NRAD_CODES, sections=SECTIONS))
HINTROWS = MIXED + [(8, 1, "01.01.2024", "Text RUNBEKANNT", "RUNBEKANNT")]
stats_hint = put("statistik_hinweis.xlsx", stats_bytes(HINTROWS, **rows(HINTROWS)))
only_n = put("statistik_nur_N.xlsx", stats_bytes(exams("NMRT", "NCT"), **rows(exams("NMRT", "NCT"))))
evil = put("statistik_evil.xlsx", stats_bytes(MIXED, total_l8=len(MIXED), mitarbeiter='<img src=x onerror="window.__pwned=1"><b>fett</b>', befunddatum="01.01.2024-31.12.2024"))
other_sections = put("map_andere_bereiche.xlsx", map_bytes(NRAD_CODES, sections=["DL Gefäße", "MRT", "Andere"]))
targets = put("map_mit_zielen.xlsx", map_bytes([("RMRT", "MRT", {"MRT gesamt": 1}), ("RCT", "CT", {"CT": 1})], sections=["DL Gefäße", "MRT gesamt", "CT", "Mamma gesamt"]))
evil_section = put("map_boese_bereiche.xlsx", map_bytes(ALLRAD_CODES, sections=["DL Gefäße", '<img src=x onerror="window.__pwned=1">', "CT"]))
broken = put("kaputt.xlsx", b"this is not an excel file")
dtd = put("statistik_dtd.xlsx", with_doctype(stats_bytes(MIXED, **rows(MIXED)), '<!DOCTYPE worksheet [<!ENTITY x "EINGESCHLEUST">]>'))
big = f"{F}/gross.xlsx"
with open(big, "wb") as fh: fh.truncate(51 * 1024 * 1024)

results = []
def shot(page, path, full=False):
    if ENGINE != "webkit": page.screenshot(path=path, full_page=full)   # in WebKit Playwright injects a style element, which the CSP of the page blocks

def check(name, ok, extra=""):
    results.append((name, bool(ok))); print(("PASS " if ok else "FAIL ") + name + (f"   [{extra}]" if extra and not ok else ""))

INIT = """
window.__csp = []; window.__requests = [];
document.addEventListener('securitypolicyviolation', e => window.__csp.push(e.violatedDirective + ' ' + e.blockedURI));
const realFetch = window.fetch;
window.fetch = async (url, opts) => { window.__requests.push(String(url).split('?')[0]); if (String(url).startsWith('evaluate') && window.__slow) await new Promise(r => setTimeout(r, window.__slow)); return realFetch(url, opts); };
"""
def new_page(b, app, **kw):
    ctx = b.new_context(accept_downloads=True, **kw); page = ctx.new_page(); page.msgs = []
    page.on("console", lambda m: page.msgs.append((m.type, m.text))); page.on("pageerror", lambda e: page.msgs.append(("pageerror", str(e))))
    page.add_init_script(INIT); page.goto(app.url); page.wait_for_load_state("networkidle"); return page
def pick(page, s=None, a=None, n=None):
    for slot, path in (("statistics", s), ("allrad", a), ("nrad", n)):
        if path: page.set_input_files(f"#file-{slot}", path)
def run_eval(page, s=stats, a=allrad, n=nrad):
    pick(page, s, a, n); page.click("#submit")
    page.wait_for_selector("#view-result:not([hidden]), #error:not([hidden])", timeout=20000)
def csp(page): return page.evaluate("window.__csp")
def visible(page, sel): return page.locator(sel).is_visible()
def download(page, dest):
    with page.expect_download(timeout=15000) as d: page.click("#download")
    d.value.save_as(dest); return d.value.suggested_filename
def table(path):
    ws = openpyxl.load_workbook(path).active; head = [c.value for c in ws[1]]; g = head.index("Gesamt")
    return {r[0]: r[g] for r in ws.iter_rows(min_row=2, values_only=True) if r[0] and not str(r[0]).startswith("Gezählte")}

with sync_playwright() as p, browser(p, ENGINE) as b:
    print(f"== engine: {ENGINE} {b.version}")
    app = App()
    try:
        # 1. the complete run: choose files, Auswerten, download
        page = new_page(b, app, viewport={"width": 900, "height": 1000})
        check("first screen: button disabled, status names what is missing", page.locator("#submit").is_disabled() and page.inner_text("#status") == "Es fehlt noch: Mitarbeiterstatistik, Allrad-Map, NRAD-Map")
        pick(page, s=stats); check("after the statistics: status names the two maps", page.inner_text("#status") == "Es fehlt noch: Allrad-Map, NRAD-Map" and page.locator("#drop-statistics").evaluate("e => e.classList.contains('has-file')"))
        check("the chosen file shows its name with a check mark and its size", page.inner_text("#text-statistics") == "✓ statistik.xlsx" and "KB" in page.inner_text("#sub-statistics"))
        check("the picker is cleared so the same file can be chosen again", page.input_value("#file-statistics") == "")
        pick(page, a=allrad); check("one map still missing", page.inner_text("#status") == "Es fehlt noch: NRAD-Map" and page.locator("#submit").is_disabled())
        pick(page, n=nrad); check("all there: status says ready, button enabled", page.inner_text("#status") == "Bereit zum Auswerten." and page.locator("#submit").is_enabled())
        shot(page, f"{S}/shot_ready.png")
        page.click("#submit"); page.wait_for_selector("#view-result:not([hidden])", timeout=20000)
        check("result screen: title, who, what", "Auswertung fertig" in page.inner_text("#result-title") and page.inner_text("#result-who") == "ZUFALL_N · 01.04.2024 – 31.12.2024" and page.inner_text("#result-what") == "7 Leistungen · Allrad + NRAD", page.inner_text("#result-what"))
        check("the form is gone, the download button is the one main button", not visible(page, "#view-form") and visible(page, "#download"))
        check("no hints, no hints box", not visible(page, "#hints"))
        # preview before the download
        check("preview: a table is shown on the result screen, before the download", visible(page, "#preview") and visible(page, "#preview-table") and page.locator("#preview-body tr").count() >= 3)
        shown = page.evaluate("[...document.querySelectorAll('#preview-head th')].map(th => th.textContent)")
        check("preview: the columns are the ones of the Excel", shown[0] == "Leistungsbereich" and "Gesamt" in shown, str(shown))
        preview_rows = page.evaluate("[...document.querySelectorAll('#preview-body tr')].map(tr => [...tr.cells].map(c => c.textContent))")
        gesamt = shown.index("Gesamt")
        check("preview: it shows the counts (MRT 4 and CT 4 for both maps)", any(r[0] == "MRT" and r[gesamt] == "4" for r in preview_rows) and any(r[0] == "CT" and r[gesamt] == "4" for r in preview_rows), str(preview_rows[:4]))
        inside = page.evaluate("(() => { const w = document.getElementById('table-wrap').getBoundingClientRect(); return [...document.querySelectorAll('#preview-head th.total, #preview-head th.target')].map(th => { const r = th.getBoundingClientRect(); return r.left >= w.left - 1 && r.right <= w.right + 1; }); })()")
        check("preview: Gesamt and Benötigt are inside the visible part of the table, not cut off", len(inside) == 2 and all(inside), str(inside))
        check("preview: the first column stays visible when the table scrolls (sticky)", page.evaluate("getComputedStyle(document.querySelector('#preview-body th')).position") == "sticky")
        all_rows = page.locator("#preview-body tr:visible").count()
        page.check("#hide-empty")
        fewer = page.locator("#preview-body tr:visible").count()
        check("preview: 'Leere Zeilen ausblenden' hides rows without any value and says how many", fewer < all_rows and "leere Zeilen ausgeblendet" in page.inner_text("#preview-note") and not visible(page, "tr.empty-row"), f"{all_rows} -> {fewer}")
        page.uncheck("#hide-empty"); check("preview: unticking shows every row again", page.locator("#preview-body tr:visible").count() == all_rows)
        check("preview: the table can be reached and scrolled with the keyboard", page.evaluate("(() => { document.getElementById('table-wrap').focus(); return document.activeElement.id; })()") == "table-wrap")
        shot(page, f"{S}/shot_preview.png")
        # the red button
        quit_style = page.evaluate("(() => { const b = getComputedStyle(document.getElementById('quit')); return [b.backgroundColor, b.color]; })()")
        def rgb(text): return [int(x) for x in re.findall(r"\d+", text)[:3]]
        def luminance(c): return sum(w * (v / 255 / 12.92 if v / 255 <= 0.03928 else ((v / 255 + 0.055) / 1.055) ** 2.4) for w, v in zip((0.2126, 0.7152, 0.0722), c))
        bg, fg = rgb(quit_style[0]), rgb(quit_style[1]); ratio = (max(luminance(bg), luminance(fg)) + 0.05) / (min(luminance(bg), luminance(fg)) + 0.05)
        check("'Server beenden' is red (light mode) with a readable text", bg[0] > 150 and bg[1] < 90 and bg[2] < 90 and ratio >= 4.5, f"{quit_style} contrast {ratio:.1f}")
        box = page.locator("#quit").bounding_box()
        check("'Server beenden' is a real button, not a quiet link (at least 120 x 36 px) and inside the window", box["width"] >= 120 and box["height"] >= 36 and box["y"] + box["height"] <= 1000, str(box))
        shot(page, f"{S}/shot_result.png")
        name = download(page, f"{F}/out_both.xlsx"); t = table(f"{F}/out_both.xlsx")
        check("download: name has the selection and the umlaut-free fallback is fine", name.endswith("_Allrad_NRAD.xlsx") and name.startswith("ZUFALL_N"), name)
        check("download: a valid Excel with the right numbers", t.get("MRT") == 4 and t.get("CT") == 4, str(t))
        check("download: the page knows it was saved", visible(page, "#saved"))
        check("no CSP violation so far (including the download)", csp(page) == [], str(csp(page)))
        # 1b. hints
        pg = new_page(b, app); run_eval(pg, s=stats_hint)
        check("hints: the box says how many", visible(pg, "#hints") and pg.inner_text("#hints-summary") == "1 Hinweis anzeigen", pg.inner_text("#hints-summary"))
        check("hints: closed until asked for", not pg.evaluate("document.getElementById('hints').open") and not visible(pg, "#hints-list"))
        pg.click("#hints-summary"); check("hints: the list shows code and text", pg.inner_text("#hints-list li") == "W201" + "1 Leistungscode fehlt in der Map (eigene Zeile in der Excel).".replace("", "") or pg.inner_text("#hints-list .code") == "W201", pg.inner_text("#hints-list"))
        check("hints: text is the catalog text", "fehlt in der Map" in pg.inner_text("#hints-list")); shot(pg, f"{S}/shot_hints.png"); pg.context.close()
        # 1b2. many years in a narrow window: Gesamt and Benötigt stay visible while the years scroll
        pg = new_page(b, app, viewport={"width": 380, "height": 900}); pick(pg, stats, allrad, nrad); pg.click("#submit"); pg.wait_for_selector("#view-result:not([hidden])")
        pg.evaluate("document.getElementById('table-wrap').scrollLeft = 10000"); pg.evaluate("document.getElementById('table-wrap').scrollLeft = 0")
        pos = pg.evaluate("(() => { const w = document.getElementById('table-wrap').getBoundingClientRect(); return [...document.querySelectorAll('#preview-head th.total, #preview-head th.target')].map(th => { const r = th.getBoundingClientRect(); return r.left >= w.left - 1 && r.right <= w.right + 1; }); })()")
        check("preview: in a 380 px window Gesamt and Benötigt are visible too, and nothing makes the page scroll sideways", all(pos) and len(pos) == 2 and pg.evaluate("document.documentElement.scrollWidth <= window.innerWidth"), str(pos)); pg.context.close()
        # 1c. the colours of the Gesamt cell follow the share of the target
        pg = new_page(b, app); pick(pg, stats, targets); pg.uncheck("#use-nrad"); pg.click("#submit"); pg.wait_for_selector("#view-result:not([hidden])")
        gi = "[...document.querySelectorAll('#preview-head th')].findIndex(th => th.textContent === 'Gesamt')"
        cells = pg.evaluate("(() => { const g = " + gi + "; return [...document.querySelectorAll('#preview-body tr')].filter(tr => tr.cells[0].textContent === 'MRT gesamt').map(tr => ({cls: tr.cells[g].className, text: [...tr.cells].map(c => c.textContent), title: tr.cells[g].title, bg: getComputedStyle(tr.cells[g]).backgroundColor})); })()")
        check("preview: a total that is a small share of its target is shaded red, with the share in the tooltip", cells and "ach-red" in cells[0]["cls"] and cells[0]["title"] == "unter 1 % des Ziels" and cells[0]["bg"] != "rgba(0, 0, 0, 0)", str(cells))
        zero = pg.evaluate("(() => { const g = " + gi + "; return [...document.querySelectorAll('#preview-body tr')].filter(tr => tr.cells[0].textContent === 'Mamma gesamt').map(tr => tr.cells[g].className); })()")
        check("preview: a total of 0 is not shaded (as in the Excel)", zero and "ach-" not in zero[0], str(zero))
        wide = pg.evaluate("document.getElementById('card').classList.contains('wide')")
        check("the result screen is wider than the form (room for the table)", wide)
        check("preview: it says what the colours mean when there are targets", "grün ab 100 %" in pg.inner_text("#preview-note")); pg.context.close()
        # 1d. markup in a section name
        pg = new_page(b, app); pick(pg, stats, evil_section); pg.uncheck("#use-nrad"); pg.click("#submit"); pg.wait_for_selector("#view-result:not([hidden])")
        names = pg.evaluate("[...document.querySelectorAll('#preview-body th')].map(th => th.textContent)")
        check("preview: markup in a section name is shown as plain text and nothing runs", any("<img src=x" in n for n in names) and pg.locator("#preview img").count() == 0 and pg.evaluate("window.__pwned") is None and csp(pg) == [], str(names[:3])); pg.context.close()
        # 2. Weitere Statistik
        page.click("#again"); check("'Weitere Statistik' returns to the form", visible(page, "#view-form") and not visible(page, "#view-result"))
        check("... clears only the statistics file", page.inner_text("#text-statistics").startswith("Datei wählen") and page.inner_text("#text-allrad") == "✓ map_allrad.xlsx" and page.inner_text("#text-nrad") == "✓ map_nrad.xlsx")
        check("... and the server forgot the result", page.evaluate("fetch('download').then(r => r.status)") == 404)
        # 3. the three selections and their download names
        for label, use_a, use_n, suffix, expect in (("Allrad only", True, False, "_Allrad.xlsx", {"MRT": 2, "CT": 1}), ("NRAD only", False, True, "_NRAD.xlsx", {"MRT": 2, "CT": 3}), ("both", True, True, "_Allrad_NRAD.xlsx", {"MRT": 4, "CT": 4})):
            pg = new_page(b, app); pick(pg, s=stats, a=allrad, n=nrad)
            if not use_a: pg.uncheck("#use-allrad")
            if not use_n: pg.uncheck("#use-nrad")
            pg.click("#submit"); pg.wait_for_selector("#view-result:not([hidden])")
            nm = download(pg, f"{F}/out_{label.replace(' ', '_')}.xlsx"); t = table(f"{F}/out_{label.replace(' ', '_')}.xlsx")
            check(f"selection {label}: right file name and numbers", nm.endswith(suffix) and all(t.get(k) == v for k, v in expect.items()), f"{nm} {t}")
            check(f"selection {label}: no CSP violation", csp(pg) == [])
            pg.context.close()
        # 4. departments: statuses and hiding
        pg = new_page(b, app); pick(pg, s=stats, a=allrad, n=nrad)
        pg.uncheck("#use-nrad"); check("unticking NRAD hides its picker, status ready", not visible(pg, "#field-nrad") and visible(pg, "#field-allrad") and pg.inner_text("#status") == "Bereit zum Auswerten.")
        pg.uncheck("#use-allrad"); check("both unticked: status says so, button disabled, maps step hidden", pg.inner_text("#status") == "Bitte mindestens einen Bereich ankreuzen." and pg.locator("#submit").is_disabled() and not visible(pg, "#step-maps"))
        pg.check("#use-nrad"); check("re-ticking shows the picker again with the file still chosen", visible(pg, "#field-nrad") and pg.inner_text("#text-nrad") == "✓ map_nrad.xlsx")
        pg.context.close()
        pg = new_page(b, app); pick(pg, s=stats, n=nrad); pg.uncheck("#use-allrad")
        check("an unticked department needs no map", pg.inner_text("#status") == "Bereit zum Auswerten." and pg.locator("#submit").is_enabled()); pg.context.close()
        # 5. drag and drop
        def drop(page, sel, files, bubbles=True):
            payload = [{"name": os.path.basename(f), "b64": base64.b64encode(open(f, "rb").read()).decode()} for f in files]
            return page.evaluate("""([sel, payload]) => { const dt = new DataTransfer(); for (const f of payload) dt.items.add(new File([Uint8Array.from(atob(f.b64), c => c.charCodeAt(0))], f.name));
                const el = document.querySelector(sel); const out = {};
                el.dispatchEvent(new DragEvent('dragenter', {dataTransfer: dt, bubbles: true, cancelable: true}));
                const over = new DragEvent('dragover', {dataTransfer: dt, bubbles: true, cancelable: true}); el.dispatchEvent(over); out.overClass = el.classList.contains('over') || (el.closest('.drop') || {classList: {contains: () => false}}).classList.contains('over');
                const d = new DragEvent('drop', {dataTransfer: dt, bubbles: true, cancelable: true}); el.dispatchEvent(d); out.prevented = d.defaultPrevented; out.overAfter = el.classList.contains('over'); return out; }""", [sel, payload])
        pg = new_page(b, app); url0 = pg.url
        r = drop(pg, "#drop-statistics", [stats]); check("drag and drop onto the drop area chooses the file and highlights while over", pg.inner_text("#text-statistics") == "✓ statistik.xlsx" and r["overClass"] and not r["overAfter"], str(r))
        r = drop(pg, "#drop-allrad", [allrad, nrad]); check("two files dropped at once are refused with a hint, nothing chosen", pg.inner_text("#text-allrad").startswith("Datei wählen") and "nur eine Datei" in pg.inner_text("#sub-allrad"), pg.inner_text("#sub-allrad"))
        r = drop(pg, "h1", [nrad]); check("a file dropped beside the drop area does nothing and does not navigate", r["prevented"] and pg.url == url0 and pg.inner_text("#text-nrad").startswith("Datei wählen"), str(r))
        r = drop(pg, "body", [nrad]); check("... also on the page background", r["prevented"] and pg.url == url0)
        drop(pg, "#drop-allrad", [allrad]); drop(pg, "#drop-nrad", [nrad]); check("all three by drag and drop: ready", pg.inner_text("#status") == "Bereit zum Auswerten.")
        pg.click("#submit"); pg.wait_for_selector("#view-result:not([hidden])"); check("evaluation after drag and drop works", "Auswertung fertig" in pg.inner_text("#result-title")); pg.context.close()
        # 6. working state
        pg = new_page(b, app); pg.evaluate("window.__slow = 1500"); pick(pg, stats, allrad, nrad); pg.click("#submit")
        pg.wait_for_function("document.getElementById('submit').textContent.includes('Wird ausgewertet')", timeout=3000)
        check("while working: button says 'Wird ausgewertet …' and is disabled", pg.inner_text("#submit") == "Wird ausgewertet …" and pg.locator("#submit").is_disabled())
        check("... the form is locked", all(pg.locator(s).is_disabled() for s in ("#file-statistics", "#file-allrad", "#file-nrad", "#use-allrad", "#use-nrad")))
        check("... 'Server beenden' stays usable", pg.locator("#quit").is_enabled() and visible(pg, "#quit"))
        pg.locator("#submit").click(force=True); pg.keyboard.press("Enter")
        pg.wait_for_selector("#view-result:not([hidden])", timeout=20000)
        check("... nothing was sent twice", pg.evaluate("window.__requests.filter(u => u.startsWith('evaluate')).length") == 1); pg.context.close()
        # 7. errors: right code, right field, everything stays in place
        def error_case(label, code, marked, **files):
            pg = new_page(b, app); run_eval(pg, **files)
            shown = pg.inner_text("#error-code") if visible(pg, "#error") else None
            red = pg.evaluate("[...document.querySelectorAll('.field.invalid')].map(e => e.dataset.slot)")
            kept = pg.inner_text("#text-statistics").startswith("✓") and pg.inner_text("#text-allrad").startswith("✓") and pg.inner_text("#text-nrad").startswith("✓")
            check(f"error {label}: code {code}, field(s) marked {marked}, entries stay", shown == code and sorted(red) == sorted(marked) and kept and visible(pg, "#view-form"), f"{shown} {red} kept={kept}")
            check(f"error {label}: the box has what happened, what to do, the code and details are collapsed", bool(pg.inner_text("#error-title")) and bool(pg.inner_text("#error-hint")) and pg.get_attribute("#error", "role") == "alert")
            check(f"error {label}: the error box has focus and the field is aria-invalid", pg.evaluate("document.activeElement.id") == "error" and (not marked or pg.get_attribute(f"#file-{marked[0]}", "aria-invalid") == "true"))
            return pg
        pg = error_case("not an Excel file in the statistics field", "E101", ["statistics"], s=broken); pg.context.close()
        pg = error_case("map in the statistics field", "E110", ["statistics"], s=allrad); pg.context.close()
        pg = error_case("statistics in the NRAD field", "E210", ["nrad"], n=stats); pg.context.close()
        pg = error_case("not an Excel file in the Allrad field", "E201", ["allrad"], a=broken); pg.context.close()
        pg = error_case("swapped maps", "E220", ["allrad"], a=nrad, n=allrad); pg.context.close()
        pg = error_case("maps with different sections", "E221", ["allrad", "nrad"], n=other_sections); pg.context.close()
        pg = error_case("a file with an XML entity", "E105", ["statistics"], s=dtd); pg.context.close()
        pg = new_page(b, app); pick(pg, only_n, allrad, None); pg.uncheck("#use-nrad"); pg.click("#submit"); pg.wait_for_selector("#error:not([hidden])")
        check("error: a selection with no matching rows (E303) marks the departments and keeps the entries", pg.inner_text("#error-code") == "E303" and pg.evaluate("document.getElementById('step-departments').classList.contains('invalid')") and "Allrad-Leistungen" in pg.inner_text("#error-title") and pg.inner_text("#text-statistics").startswith("✓"))
        pg.check("#use-nrad"); check("... fixing the checkbox clears the mark", not pg.evaluate("document.getElementById('step-departments').classList.contains('invalid')") and not visible(pg, "#error")); pg.context.close()
        pg = new_page(b, app); run_eval(pg, s=broken); pick(pg, s=stats)
        check("an error goes away when the file it names is replaced, the rest stays", not visible(pg, "#error") and pg.inner_text("#text-allrad") == "✓ map_allrad.xlsx")
        pg.click("#submit"); pg.wait_for_selector("#view-result:not([hidden])"); check("... and the corrected run works", "Auswertung fertig" in pg.inner_text("#result-title")); pg.context.close()
        pg = new_page(b, app); pick(pg, big, allrad, nrad); pg.click("#submit"); pg.wait_for_selector("#error:not([hidden])")
        check("a file over the size limit is refused in the browser with E103, before any upload", pg.inner_text("#error-code") == "E103" and pg.evaluate("window.__requests.filter(u => u.startsWith('evaluate')).length") == 0 and pg.evaluate("document.getElementById('field-statistics').classList.contains('invalid')")); pg.context.close()
        # 8. unreadable file -> E306
        pg = new_page(b, app); gone = put("wird_geloescht.xlsx", open(stats, "rb").read()); pick(pg, gone, allrad, nrad); os.remove(gone); pg.click("#submit"); pg.wait_for_selector("#error:not([hidden])")
        check("a file that changed after choosing it gives E306 for that field, not a lost connection", pg.inner_text("#error-code") == "E306" and pg.evaluate("document.getElementById('field-statistics').classList.contains('invalid')") and pg.evaluate("window.__requests.filter(u => u.startsWith('evaluate')).length") == 0)
        check("... and the busy state is over", pg.locator("#submit").is_enabled() and pg.inner_text("#submit") == "Auswerten"); pg.context.close()
        # 9. hostile content
        pg = new_page(b, app); run_eval(pg, s=evil)
        who = pg.inner_text("#result-who")
        check("HTML in the staff name is shown as plain text and nothing ran", "<img src=x" in who and "<b>fett</b>" in who and pg.evaluate("window.__pwned") is None and pg.locator("#view-result img, #view-result b").count() == 0, who)
        check("... no CSP violation either", csp(pg) == []); pg.context.close()
        # 10. CSP / console over everything seen with a clean page
        pg = new_page(b, app); run_eval(pg); download(pg, f"{F}/x.xlsx")
        check("browser console shows no errors or warnings, no CSP violation", not [m for m in pg.msgs if m[0] in ("error", "warning", "pageerror")] and csp(pg) == [], str(pg.msgs)); pg.context.close()
        # 11. quitting
        pg = new_page(b, app); run_eval(pg)
        pg.click("#quit"); check("'Server beenden' asks first when the result was not downloaded", visible(pg, "#confirm") and pg.locator("#quit-no").evaluate("e => document.activeElement === e"))
        pg.click("#quit-no"); check("'Abbrechen' keeps everything", not visible(pg, "#confirm") and visible(pg, "#view-result") and app.proc.poll() is None)
        download(pg, f"{F}/y.xlsx"); pg.click("#quit")
        pg.wait_for_selector("#view-closed:not([hidden])", timeout=5000)
        check("after the download it asks nothing: message shown, form gone", "beendet" in pg.inner_text("#view-closed") and not visible(pg, "#view-form") and not visible(pg, "#footer"))
        t0 = time.time()
        while app.proc.poll() is None and time.time() - t0 < 10: time.sleep(0.1)
        check("... and the process is gone with exit code 0", app.proc.poll() == 0, str(app.proc.poll())); pg.context.close()
    finally:
        app.stop()
    # 12. E902 and the confirmed quit, each with its own program
    for case in ("e902", "confirmed quit"):
        app = App()
        try:
            pg = new_page(b, app); pick(pg, stats, allrad, nrad)
            if case == "e902":
                app.proc.terminate(); app.proc.wait(10); pg.click("#submit"); pg.wait_for_selector("#error:not([hidden])", timeout=10000)
                check("E902 appears when the program was stopped while the page is open", pg.inner_text("#error-code") == "E902" and "neu starten" in pg.inner_text("#error-hint") and pg.locator("#submit").is_enabled())
            else:
                pg.click("#submit"); pg.wait_for_selector("#view-result:not([hidden])"); pg.click("#quit"); pg.click("#quit-yes"); pg.wait_for_selector("#view-closed:not([hidden])", timeout=5000)
                t0 = time.time()
                while app.proc.poll() is None and time.time() - t0 < 10: time.sleep(0.1)
                check("'Ja, beenden' without a download ends the program too", app.proc.poll() == 0 and "beendet" in pg.inner_text("#view-closed"))
            pg.context.close()
        finally:
            app.stop()
    # 13. appearance and keyboard
    app = App()
    try:
        pg = new_page(b, app, color_scheme="dark", viewport={"width": 900, "height": 1000}); run_eval(pg)
        bg_dark = pg.evaluate("getComputedStyle(document.documentElement).backgroundColor"); shot(pg, f"{S}/shot_dark_result.png"); pg.context.close()
        pg = new_page(b, app, color_scheme="light"); bg_light = pg.evaluate("getComputedStyle(document.documentElement).backgroundColor"); pg.context.close()
        check("dark mode follows the system and differs from light mode", bg_dark != bg_light and bg_dark.startswith("rgb(17"), f"{bg_dark} vs {bg_light}")
        pg = new_page(b, app, color_scheme="dark", viewport={"width": 900, "height": 1000}); pick(pg, stats, allrad, nrad); shot(pg, f"{S}/shot_dark_form.png"); pg.context.close()
        pg = new_page(b, app, viewport={"width": 360, "height": 640}); pick(pg, stats, allrad, nrad)
        check("small window (360 px): no horizontal scrolling", pg.evaluate("document.documentElement.scrollWidth <= window.innerWidth"), pg.evaluate("[document.documentElement.scrollWidth, window.innerWidth]").__str__())
        shot(pg, f"{S}/shot_small.png", True)
        pg.click("#submit"); pg.wait_for_selector("#view-result:not([hidden])")
        check("small window: result fits too", pg.evaluate("document.documentElement.scrollWidth <= window.innerWidth")); shot(pg, f"{S}/shot_small_result.png", True); pg.context.close()
        # keyboard only (the file dialog itself cannot be driven here: files are set directly, everything else uses keys)
        pg = new_page(b, app); order = []
        for _ in range(8):
            pg.keyboard.press("Tab"); order.append(pg.evaluate("document.activeElement.id"))
        check("tab order follows the page: statistics, Allrad, NRAD, then the maps, then 'Server beenden'", order[:5] == ["file-statistics", "use-allrad", "use-nrad", "file-allrad", "file-nrad"] and "quit" in order, str(order))
        pg.focus("#file-statistics"); check("the focused drop area shows a visible focus ring", pg.evaluate("getComputedStyle(document.getElementById('drop-statistics')).outlineStyle") != "none")
        pg.keyboard.press("Tab"); pg.keyboard.press("Space"); check("a checkbox toggles with the space bar", not pg.is_checked("#use-allrad")); pg.keyboard.press("Space")
        pick(pg, stats, allrad, nrad); pg.focus("#use-nrad")
        for _ in range(3): pg.keyboard.press("Tab")
        check("with everything chosen the next tab stop after the picker is 'Auswerten'", pg.evaluate("document.activeElement.id") == "submit", pg.evaluate("document.activeElement.id"))
        pg.keyboard.press("Enter"); pg.wait_for_selector("#view-result:not([hidden])")
        check("Enter starts the evaluation and focus moves to the result", pg.evaluate("document.activeElement.id") == "result-title")
        pg.keyboard.press("Tab")
        with pg.expect_download(timeout=15000) as d: pg.keyboard.press("Enter")
        check("keyboard: Tab to 'Excel herunterladen', Enter downloads", d.value.suggested_filename.endswith(".xlsx") and pg.evaluate("document.activeElement.id") == "download")
        pg.context.close()
    finally:
        app.stop()

bad = [n for n, ok in results if not ok]
print(f"\n{len(results) - len(bad)} of {len(results)} checks passed" + (f"; FAILED: {bad}" if bad else ""))
sys.exit(1 if bad else 0)
