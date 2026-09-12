# DSK Film-Katalog

GUI-Tool, das einen Ordner (z.B. eine Festplatte / Filme-Sammlung) rekursiv
nach Video-Dateien durchsucht, aus dem Dateinamen Titel + Jahr extrahiert und
über die TMDb-API automatisch **Genre** und **Produktionsland** nachschlägt.

Ergebnis-Tabelle: **Nummer, Titel, Genre, Land**.

- Per Klick auf eine Spaltenüberschrift wird danach sortiert (nochmaliger Klick
  kehrt die Reihenfolge um).
- Export als **CSV, Excel (xlsx), PDF oder TXT** - das gewünschte Format wird
  einfach über die Dateiendung im Speichern-Dialog gewählt.

## Nutzung (Windows-exe)

1. Repo als neues GitHub-Repository anlegen (z.B. `dsk-filmkatalog`) und
   `main.py`, `.github/workflows/build.yml` hochladen.
2. Unter dem Reiter **Actions** den Workflow "Build DSK Film-Katalog exe"
   manuell starten (oder er läuft automatisch bei einem Push auf `main`).
3. Nach Abschluss unter **Artifacts** die Datei `DSK-Film-Katalog` herunterladen
   -> enthält `DSK-Film-Katalog.exe`.
4. Beim ersten Start: eigenen **TMDb API-Key** eintragen (wird lokal in
   `%USERPROFILE%\.dsk_filmkatalog.json` gespeichert) und den zu scannenden
   Ordner auswählen.

## Lokal testen (optional, mit Python)

```
pip install -r requirements.txt   # openpyxl + reportlab für Excel-/PDF-Export
python main.py
```

## Hinweise

- Erkannte Video-Endungen: mp4, mkv, avi, mov, wmv, m4v, ts, flv, webm
- Die Titel-Erkennung entfernt gängige Release-Tags (Auflösung, Codec, Gruppe,
  Sprache etc.) und schneidet ab dem erkannten Jahr ab.
- Scan läuft in einem Hintergrund-Thread, GUI bleibt bedienbar (inkl. Stop-Button).
- Aktuell nur Filme (keine Serien-Erkennung).
