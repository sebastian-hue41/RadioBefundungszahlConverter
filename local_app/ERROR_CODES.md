# Error codes and hints

Every message of the app has a code. If something goes wrong, the page shows what happened, what to do and the code
("Fehlercode E210"). Quote the code, and the "Technische Details" under it, when you ask for help.

This file is generated from `local_app/errors.py`; a test fails if the two differ. To refresh it after a change run
`python -m local_app.tests.test_docs --write`.

- **E1xx**: the statistics file. **E2xx**: a map file (E220/E221: the two maps together). **E3xx**: the request and the selection.
  **E9xx**: the program itself. **W…**: hints; the evaluation still worked and the Excel was made.

## Errors (nothing was evaluated)

| Code | What happened | What to do |
|---|---|---|
| E100 | Die Datei konnte nicht gelesen werden. | Bitte den Code und die technischen Details melden. |
| E200 | Die Datei konnte nicht gelesen werden. | Bitte den Code und die technischen Details melden. |
| E101 | Die Datei ist keine gültige Excel-Datei. | Bitte als .xlsx ohne Passwortschutz speichern oder neu exportieren. |
| E201 | Die Datei ist keine gültige Excel-Datei. | Bitte als .xlsx ohne Passwortschutz speichern oder neu exportieren. |
| E102 | Die Datei ist leer. | Bitte die richtige Datei wählen. |
| E202 | Die Datei ist leer. | Bitte die richtige Datei wählen. |
| E103 | Die Datei ist größer als 50 MB. | Bitte die richtige Datei wählen oder einen kürzeren Zeitraum exportieren. |
| E203 | Die Datei ist größer als 50 MB. | Bitte die richtige Datei wählen. |
| E104 | Die Datei enthält Makros. | Bitte als normale .xlsx ohne Makros speichern. |
| E204 | Die Datei enthält Makros. | Bitte als normale .xlsx ohne Makros speichern. |
| E105 | Die Datei wurde aus Sicherheitsgründen abgelehnt. | Bitte neu exportieren. Bei Wiederholung den Code melden. |
| E205 | Die Datei wurde aus Sicherheitsgründen abgelehnt. | Bitte neu exportieren. Bei Wiederholung den Code melden. |
| E110 | Das ist keine Mitarbeiterstatistik. | Bitte den Export aus dem RIS wählen. Wurde versehentlich eine Map gewählt? |
| E111 | In der Statistik fehlt eine Pflichtspalte. | Bitte den Export prüfen (der Spaltenname steht in den technischen Details). |
| E112 | Die Statistik enthält keine auswertbaren Leistungen. | Bitte Export und Zeitraum prüfen. |
| E113 | Die Statistik hat zu viele Zeilen. | Bitte einen kürzeren Zeitraum exportieren. |
| E210 | Das ist keine Leistungsmap. | Bitte die richtige Map wählen. Wurde versehentlich die Statistik gewählt? |
| E211 | Die Map enthält keine Leistungen. | Bitte die richtige Map wählen. |
| E212 | Die Map hat zu viele Zeilen oder Spalten. | Bitte die richtige Map wählen. |
| E213 | Zwei Spalten der Map haben denselben Namen. | Bitte in der Map jedem Leistungsbereich einen eigenen Spaltennamen geben (Name und Spalten stehen in den technischen Details). |
| E214 | Die Map enthält Werte, die keine ganzen Zahlen sind. | In den Spalten der Leistungsbereiche sind nur leere Zellen und ganze Zahlen (0, 1, 2 …) erlaubt. Die Zellen stehen in den technischen Details. |
| E220 | Allrad- und NRAD-Map sind vertauscht. | Bitte die beiden Dateien tauschen. |
| E221 | Allrad- und NRAD-Map passen nicht zusammen. | Bitte zusammengehörige Versionen wählen. |
| E222 | Es wurden zwei Maps derselben Art gewählt. | Bitte je eine Allrad-Map und eine NRAD-Map wählen. |
| E301 | Kein Bereich ausgewählt. | Bitte „Alles“, „Nur NRAD-Codes“ oder „Nur Allrad-Codes“ wählen. |
| E302 | Es fehlt eine Datei. | Bitte die markierte Datei wählen. |
| E303 | Diese Statistik enthält keine {dept}. | Bitte bei „Was auswerten?“ eine andere Auswahl treffen, zum Beispiel „Alles“. |
| E304 | Die Übertragung war unvollständig. | Bitte noch einmal versuchen. |
| E305 | Die Dateien sind zusammen zu groß. | Bitte die richtigen Dateien wählen. |
| E306 | Die Datei konnte im Browser nicht gelesen werden. | Sie liegt eventuell nur in der Cloud oder wurde inzwischen geändert. Bitte lokal speichern und neu wählen. |
| E307 | Es läuft bereits eine Auswertung. | Bitte kurz warten oder das Programm neu starten. |
| E901 | Das Ergebnis ist nicht mehr vorhanden. | Bitte erneut auswerten. |
| E902 | Die Verbindung zum Programm ist unterbrochen. | Bitte das Programm neu starten. |
| E903 | Die Auswertung wurde aus Sicherheitsgründen nicht erstellt. | Bitte die Eingabedateien prüfen und den Code melden. |
| E999 | Unerwarteter Fehler. | Bitte noch einmal versuchen; sonst den Code und die technischen Details melden. |

## Hints (the Excel was made)

| Code | Text |
|---|---|
| W101 | 1 Zeile wurde übersprungen. / N Zeilen wurden übersprungen. |
| W102 | Mitarbeiter oder Befunddatum fehlen; der Dateiname ist allgemein. |
| W103 | Die Zählerprüfung war nicht möglich. |
| W104 | Die Anzahl der Leistungen weicht vom Erwarteten ab. |
| W105 | Bei 1 Leistungscode sieht der Leistungsname wie eine Formel aus; in der Excel steht der Code statt des Namens. / Bei N Leistungscodes sieht der Leistungsname wie eine Formel aus; in der Excel steht der Code statt des Namens. |
| W201 | 1 Leistungscode fehlt in der Map (eigene Zeile in der Excel). / N Leistungscodes fehlen in der Map (eigene Zeilen in der Excel). |
| W202 | 1 Leistung hat keine Zuordnung (am Ende der Excel aufgelistet). / N Leistungen haben keine Zuordnung (am Ende der Excel aufgelistet). |
| W301 | Soll-Werte konnten nicht geladen werden; die Spalte „Benötigt“ fehlt. |
| W302 | Einzelne Zeilen der Referenzdatei konnten nicht verwendet werden (siehe Excel, „Hinweis: Probleme mit der Referenzdatei“). |
| W401 | 1 Leistung der anderen Abteilung wurde nicht ausgewertet. / N Leistungen der anderen Abteilung wurden nicht ausgewertet. |
| W900 | 1 weiterer Hinweis (siehe Excel, „Hinweise und Fehler“). / N weitere Hinweise (siehe Excel, „Hinweise und Fehler“). |
