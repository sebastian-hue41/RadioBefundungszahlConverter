# Befundungszahlen – local app

Turns a staff statistics file from the RIS into the Excel evaluation, in your browser, on your own computer.
Nothing is sent anywhere and nothing is stored: the only file you get is the Excel you download.

## 1. Start it

**The ready-made program** (nothing to install)

1. Get the file for your system from the build: GitHub → **Actions** → the latest run of *Build local app* → **Artifacts**
   (`befundungszahlen-windows-x64.exe`, `befundungszahlen-macos-arm64.tar.gz` or `befundungszahlen-linux-x64.tar.gz`).
2. Windows: double-click the `.exe`. macOS / Linux: `tar -xzf befundungszahlen-*.tar.gz`, then `./befundungszahlen`.
3. A window prints an address and your browser opens it. **The first start takes up to 10 seconds.** If no browser opens,
   copy the printed address into one. Keep the window open while you work.

The program is not signed, so Windows ("More info" → "Run anyway") and macOS (right-click → "Open") warn on the first start.

**From the source** (Python 3.14):

```bash
pip install -r requirements.txt
python -m local_app.main
```

## 2. Use it

1. **Mitarbeiterstatistik**: choose the statistics file, or drag it onto the box.
2. **Welche Bereiche?**: tick **Allrad**, **NRAD** or both.
3. **Leistungsmaps**: choose the map file for each ticked department (see below).
4. **Auswerten**, look at the **Vorschau** (the table of the Excel, before you download it), then **Excel herunterladen**.
5. Next person: **Weitere Statistik hochladen** (the maps stay chosen until you reload the page).
6. When you are done: **Server beenden**.

The line next to **Auswerten** always says what is still missing. A red box says what went wrong and what to do.

The program stops by itself after 30 minutes without use. The time left is shown at the bottom; every click, key press or scroll
starts the 30 minutes again. To keep it running, use **Automatisches Beenden ausschalten** (it asks once more before it does it).
`--idle-minutes N` changes the time, `--idle-minutes 0` turns it off.

| You tick | Counted | Map used |
|---|---|---|
| Allrad | codes that do **not** start with `N` | Allrad map |
| NRAD | codes that start with `N` | NRAD map |
| both | everything | Allrad map, and the NRAD map for the `N` codes |

## 3. The files you need

**You need a map file. It is not part of the program** (it is confidential); you choose it yourself.

*Map file* (`.xlsx`, up to 50 MB, no password, no macros):

- Row 1 is the header. **B1 must be `Leist0`** and **F1 must be `DL Gefäße`**.
- **Column B**: the Leistungscode, the same as `Leist-Kurz` in the statistics. Column C: its name (optional).
- **From column F on, every column is one Leistungsbereich**, named in row 1; that name becomes a row of the Excel.
  Columns whose name starts with `_` are ignored. Codes and names must not start with `=`, `+`, `-`, `@` or `|`.
- In a section column: `1` = the code counts once for that section, `2` = twice, empty or `0` = not at all.
  A code may count in several sections.
- A code listed twice is merged (the higher value wins) and noted in the Excel. A code missing from the map gets its own row.
- **Two maps**: the Allrad map holds the Allrad codes, the NRAD map the `N` codes. Put each in its own field: a swapped
  pair is refused (E220). To use both together they need the same sections (E221 otherwise).

*Statistics file* (`.xlsx` from the RIS, up to 50 MB and 500,000 rows): labelled `Mitarbeiter` and `Befunddatum`, a header
row with `IND1`, `IND2`, `DokDatum`, `Leistung` and `Leist-Kurz` (row 10 in the RIS export) and one row per Leistung below it.

The Excel also gets a **Benötigt** column; its target values are built in, there is nothing to choose.

## 4. If something goes wrong

The red box shows a code such as `E210`. All codes and what to do: [ERROR_CODES.md](ERROR_CODES.md).
After a successful evaluation, **… Hinweise anzeigen** lists notes (for example codes that are missing from the map). They do not stop anything.

## For developers

- This code lives only on the branch `local` and is never merged into `main`; it uses `core/` and changes no file of it.
- Tests: `pip install -r requirements-dev.txt`, then `python -m pytest`.
- Build the single-file program (outside the synced folder, on the system it is meant for):

```bash
pip install -r local_app/requirements-build.txt
pyinstaller --noconfirm --clean --distpath ../dist --workpath ../work local_app/radiobefund_local.spec
python local_app/tests/smoke_binary.py ../dist/befundungszahlen
```

GitHub builds all three systems on every push to `local` (`.github/workflows/build-local.yml`).
Only the page and `reference_amount.xlsx` are inside the program; no map, no statistics.
