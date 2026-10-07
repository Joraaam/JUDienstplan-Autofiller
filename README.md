# Dienstplan Autofill

Liest den monatlichen **Dienstplan (PDF)** und füllt automatisch die **Arbeitszeiterfassung-Excel-Dateien** der Azubis aus.
Unterstützt **HC**, **Fiege (FACL)** und **LUG FRA**. Die Firma und die Schichtzeiten werden aus dem PDF erkannt.

## Ablauf
1. Dienstplan-PDF und die Excel-Dateien (`<Jahr>_Arbeitszeiterfassung_<Nachname>_<Vorname>.xlsx`) hochladen.
2. Zuordnung und Monat prüfen, ein Klick auf einen Tag ändert den Code.
3. Ausfüllen, dann einzeln oder als ZIP herunterladen.

Geschrieben werden nur die Zellen D (Kommt), E (Geht), J (Code), K (IST) und P (Bemerkung) der Tageszeilen 4-34.
Alle Formeln, Formatierungen und Dropdowns der Vorlage bleiben erhalten.

| Plan | Eintrag |
|---|---|
| Schicht (F, M, S, S3, S4, N, 1-7, WD) | Kommt/Geht laut Legende, Pause rechnet Excel |
| BS | Code `B`, IST fix 8:00 (bei jeder Firma) |
| U / UU | Code `U`, IST = tägliche SOLL-Zeit der Vorlage |
| K | Code `K`, IST = tägliche SOLL-Zeit der Vorlage |
| S/L (HC) | wie S, Bemerkung "Training" |
| DGR | nur "DGR" in den Bemerkungen, keine Zeiten |

## Lokal ausführen
```bash
pip install -r requirements.txt
python api/index.py        # oder start_local.bat
```
Öffnet http://127.0.0.1:5173.

## Windows-Programm (.exe) bauen
`build_exe.bat` ausführen. Ergebnis: `dist/Dienstplan-Autofill.exe`, läuft ohne Python auf anderen Windows-PCs.

## Projektstruktur
```
api/index.py        Flask-API (Vercel-Funktion und lokaler Start)
core/plan_parser.py Dienstplan-PDF lesen (Firma, Legende, Codes je Tag)
core/matching.py    Excel-Datei einer Zeile im Plan zuordnen
core/excel_fill.py  XLSX direkt im XML füllen (ohne Formeln anzufassen)
public/index.html   Oberfläche
vercel.json         Routing und Funktions-Einstellungen
```
