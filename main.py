#!/usr/bin/env python3
"""
DSK Film-Katalog
-----------------
Durchsucht ein Verzeichnis (z.B. eine Festplatte / einen Filme-Ordner) nach
Video-Dateien, extrahiert aus den Dateinamen Titel (+ Jahr) und fragt bei
TMDb (themoviedb.org) automatisch Genre und Produktionsland ab.

Spalten der Ergebnistabelle: Nummer, Titel, Genre, Land
Per Klick auf eine Spaltenüberschrift kann danach sortiert werden.
Export als CSV, Excel (xlsx), PDF oder TXT möglich.

Design/Aufbau angelehnt an die anderen DSK-Tools (DSK Suite, DSK NVEncC Studio):
dunkles, metallisches Theme mit "by DSK"-Signatur, Einstellungen werden
lokal gespeichert, GUI bleibt während des Scans responsiv (Threading + Queue).
"""

import json
import os
import queue
import re
import threading
import tkinter as tk
import urllib.parse
import urllib.request
import csv
from tkinter import ttk, filedialog, messagebox

# Spaltenreihenfolge zentral definiert (Schlüssel, Anzeigename, Breite, Anker)
COLUMNS = [
    ("nummer", "Nummer", 70, "center"),
    ("titel", "Titel", 380, "w"),
    ("genre", "Genre", 220, "w"),
    ("land", "Land", 160, "w"),
]

# ---------------------------------------------------------------------------
# Konfiguration / Konstanten
# ---------------------------------------------------------------------------

APP_NAME = "DSK Film-Katalog"
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".dsk_filmkatalog.json")

VIDEO_EXTENSIONS = {
    ".mp4", ".mkv", ".avi", ".mov", ".wmv", ".m4v", ".ts", ".flv", ".webm",
}

TMDB_SEARCH_URL = "https://api.themoviedb.org/3/search/movie"
TMDB_DETAILS_URL = "https://api.themoviedb.org/3/movie/{id}"

# Farben - gleiches metallisches Dark-Theme wie bei DSK Suite / NVEncC Studio
COLOR_BG = "#1e1f22"
COLOR_BG_ALT = "#2a2b2e"
COLOR_FG = "#e6e6e6"
COLOR_ACCENT = "#8a8f98"
COLOR_ACCENT_BRIGHT = "#c9cdd3"
COLOR_HEADER = "#3a3c40"
COLOR_SIGNATURE = "#9ea3ab"

# Release-Tags, die beim Titel-Parsing entfernt werden
RELEASE_NOISE = [
    r"\b(1080p|720p|2160p|480p|4k|uhd|hdr10?|dv|hdtv|webrip|web-?dl|bluray|"
    r"bdrip|brrip|dvdrip|remux|x264|x265|h ?264|h ?265|hevc|avc|aac\d?|"
    r"dts(-?hd)?|ac3|truehd|atmos|repack|proper|extended|unrated|readnfo|"
    r"multi|german|dl|dubbed|subbed)\b",
]


# ---------------------------------------------------------------------------
# Titel-Parsing
# ---------------------------------------------------------------------------

def parse_title_and_year(filename: str):
    """Extrahiert (Titel, Jahr) aus einem Release-Dateinamen."""
    name = os.path.splitext(filename)[0]
    name = name.replace(".", " ").replace("_", " ")

    year = None
    year_match = re.search(r"\b(19\d{2}|20\d{2})\b", name)
    if year_match:
        year = year_match.group(1)
        # Alles ab dem Jahr abschneiden - danach kommen i.d.R. nur noch Release-Tags
        name = name[: year_match.start()]

    for pattern in RELEASE_NOISE:
        name = re.sub(pattern, "", name, flags=re.IGNORECASE)

    name = re.sub(r"[\[\](){}]", " ", name)
    name = re.sub(r"\s+", " ", name).strip(" -.")
    return name, year


# ---------------------------------------------------------------------------
# TMDb-Abfrage
# ---------------------------------------------------------------------------

class TmdbClient:
    def __init__(self, api_key: str):
        self.api_key = api_key

    def _get(self, url: str, params: dict):
        query = urllib.parse.urlencode(params)
        full_url = f"{url}?{query}"
        with urllib.request.urlopen(full_url, timeout=10) as response:
            return json.loads(response.read().decode("utf-8"))

    def lookup(self, title: str, year: str | None):
        params = {"api_key": self.api_key, "query": title, "language": "de-DE"}
        if year:
            params["year"] = year

        data = self._get(TMDB_SEARCH_URL, params)
        results = data.get("results") or []
        if not results and year:
            # Fallback ohne Jahr, falls das geparste Jahr falsch war
            params.pop("year", None)
            data = self._get(TMDB_SEARCH_URL, params)
            results = data.get("results") or []

        if not results:
            return None

        movie_id = results[0]["id"]
        details = self._get(
            TMDB_DETAILS_URL.format(id=movie_id),
            {"api_key": self.api_key, "language": "de-DE"},
        )

        genres = ", ".join(g["name"] for g in details.get("genres", []))
        countries = details.get("production_countries", [])
        country = ", ".join(c["name"] for c in countries) if countries else ""
        matched_title = details.get("title") or results[0].get("title") or title

        return {"title": matched_title, "genre": genres, "country": country}


# ---------------------------------------------------------------------------
# Dateisystem-Scan
# ---------------------------------------------------------------------------

def find_video_files(root_folder: str):
    for dirpath, _dirnames, filenames in os.walk(root_folder):
        for fname in filenames:
            if os.path.splitext(fname)[1].lower() in VIDEO_EXTENSIONS:
                yield os.path.join(dirpath, fname)


# ---------------------------------------------------------------------------
# Einstellungen
# ---------------------------------------------------------------------------

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_config(cfg: dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class FilmKatalogApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("980x620")
        self.minsize(780, 480)
        self.configure(bg=COLOR_BG)

        self.config_data = load_config()
        self.result_queue: "queue.Queue" = queue.Queue()
        self.scan_thread: threading.Thread | None = None
        self.stop_requested = False
        self.rows = []  # gesammelte Ergebnisse für CSV-Export

        self._build_style()
        self._build_widgets()
        self.after(100, self._poll_queue)

    # -- Styling -----------------------------------------------------------

    def _build_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("TFrame", background=COLOR_BG)
        style.configure("TLabel", background=COLOR_BG, foreground=COLOR_FG,
                         font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=COLOR_BG,
                         foreground=COLOR_ACCENT_BRIGHT,
                         font=("Segoe UI", 14, "bold"))
        style.configure("Sig.TLabel", background=COLOR_BG,
                         foreground=COLOR_SIGNATURE,
                         font=("Segoe UI", 9, "italic"))
        style.configure("TButton", background=COLOR_HEADER, foreground=COLOR_FG,
                         font=("Segoe UI", 10), padding=6, borderwidth=0)
        style.map("TButton", background=[("active", COLOR_ACCENT)])
        style.configure("TEntry", fieldbackground=COLOR_BG_ALT,
                         foreground=COLOR_FG, insertcolor=COLOR_FG)
        style.configure("Horizontal.TProgressbar", background=COLOR_ACCENT_BRIGHT,
                         troughcolor=COLOR_BG_ALT, borderwidth=0)

        style.configure("Treeview", background=COLOR_BG_ALT, fieldbackground=COLOR_BG_ALT,
                         foreground=COLOR_FG, rowheight=26, borderwidth=0,
                         font=("Segoe UI", 10))
        style.configure("Treeview.Heading", background=COLOR_HEADER,
                         foreground=COLOR_ACCENT_BRIGHT, font=("Segoe UI", 10, "bold"))
        style.map("Treeview", background=[("selected", COLOR_ACCENT)])

    # -- Aufbau --------------------------------------------------------------

    def _build_widgets(self):
        header = ttk.Frame(self)
        header.pack(fill="x", padx=16, pady=(14, 6))

        ttk.Label(header, text=APP_NAME, style="Header.TLabel").pack(side="left")
        ttk.Label(header, text="by DSK", style="Sig.TLabel").pack(side="right")

        # Ordner- und API-Key Zeile
        controls = ttk.Frame(self)
        controls.pack(fill="x", padx=16, pady=6)

        ttk.Label(controls, text="Ordner:").grid(row=0, column=0, sticky="w")
        self.folder_var = tk.StringVar(value=self.config_data.get("last_folder", ""))
        folder_entry = ttk.Entry(controls, textvariable=self.folder_var, width=60)
        folder_entry.grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(controls, text="Durchsuchen…", command=self._choose_folder).grid(
            row=0, column=2, padx=4)

        ttk.Label(controls, text="TMDb API-Key:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.api_key_var = tk.StringVar(value=self.config_data.get("api_key", ""))
        api_entry = ttk.Entry(controls, textvariable=self.api_key_var, width=60, show="•")
        api_entry.grid(row=1, column=1, sticky="ew", padx=6, pady=(8, 0))

        controls.columnconfigure(1, weight=1)

        # Buttons
        action_row = ttk.Frame(self)
        action_row.pack(fill="x", padx=16, pady=8)

        self.start_btn = ttk.Button(action_row, text="Scan starten", command=self._start_scan)
        self.start_btn.pack(side="left")
        self.stop_btn = ttk.Button(action_row, text="Stop", command=self._stop_scan, state="disabled")
        self.stop_btn.pack(side="left", padx=6)
        ttk.Button(action_row, text="Exportieren…", command=self._export).pack(side="left", padx=6)

        self.status_var = tk.StringVar(value="Bereit.")
        ttk.Label(action_row, textvariable=self.status_var).pack(side="right")

        # Fortschrittsbalken
        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", padx=16, pady=(0, 8))

        # Tabelle
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=16, pady=(0, 14))

        column_keys = [c[0] for c in COLUMNS]
        self.tree = ttk.Treeview(table_frame, columns=column_keys, show="headings")

        self.sort_state = {"column": None, "reverse": False}
        for key, label, width, anchor in COLUMNS:
            self.tree.heading(key, text=label, command=lambda k=key: self._sort_by(k))
            self.tree.column(key, width=width, anchor=anchor)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

    # -- Aktionen --------------------------------------------------------------

    def _choose_folder(self):
        folder = filedialog.askdirectory(title="Ordner mit Filmen auswählen")
        if folder:
            self.folder_var.set(folder)

    def _start_scan(self):
        folder = self.folder_var.get().strip()
        api_key = self.api_key_var.get().strip()

        if not folder or not os.path.isdir(folder):
            messagebox.showerror(APP_NAME, "Bitte einen gültigen Ordner auswählen.")
            return
        if not api_key:
            messagebox.showerror(APP_NAME, "Bitte einen TMDb API-Key eingeben.")
            return

        save_config({"last_folder": folder, "api_key": api_key})

        # Tabelle leeren
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.rows = []

        self.stop_requested = False
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.status_var.set("Suche Video-Dateien…")
        self.progress.configure(value=0)

        self.scan_thread = threading.Thread(
            target=self._scan_worker, args=(folder, api_key), daemon=True
        )
        self.scan_thread.start()

    def _stop_scan(self):
        self.stop_requested = True
        self.status_var.set("Stoppe…")

    def _sort_by(self, column: str):
        """Sortiert die Tabelle nach der angeklickten Spalte (auf-/absteigend umschaltbar)."""
        items = [(self.tree.set(item_id, column), item_id) for item_id in self.tree.get_children("")]

        if column == "nummer":
            items.sort(key=lambda t: int(t[0]) if t[0].isdigit() else 0)
        else:
            items.sort(key=lambda t: t[0].lower())

        reverse = self.sort_state["column"] == column and not self.sort_state["reverse"]
        if reverse:
            items.reverse()
        self.sort_state = {"column": column, "reverse": reverse}

        for position, (_value, item_id) in enumerate(items):
            self.tree.move(item_id, "", position)

        # Pfeil-Indikator in der Spaltenüberschrift aktualisieren
        arrow = " ▼" if reverse else " ▲"
        for key, label, _width, _anchor in COLUMNS:
            text = label + (arrow if key == column else "")
            self.tree.heading(key, text=text)

        # self.rows in die neue Reihenfolge bringen, damit der Export sortiert bleibt
        self.rows = [
            {k: self.tree.set(item_id, k) for k in ("nummer", "titel", "genre", "land")}
            for item_id in self.tree.get_children("")
        ]

    def _scan_worker(self, folder: str, api_key: str):
        client = TmdbClient(api_key)
        files = list(find_video_files(folder))
        total = len(files)
        self.result_queue.put(("total", total))

        for index, filepath in enumerate(files, start=1):
            if self.stop_requested:
                break

            filename = os.path.basename(filepath)
            title, year = parse_title_and_year(filename)

            genre, country, matched_title = "", "", title
            try:
                info = client.lookup(title, year)
                if info:
                    genre = info["genre"]
                    country = info["country"]
                    matched_title = info["title"]
                else:
                    genre, country = "(nicht gefunden)", ""
            except Exception as exc:  # Netzwerkfehler etc. - Zeile trotzdem anzeigen
                genre, country = f"(Fehler: {exc})", ""

            self.result_queue.put(("row", (index, matched_title, genre, country)))

        self.result_queue.put(("done", None))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.result_queue.get_nowait()
                if kind == "total":
                    self.progress.configure(maximum=max(payload, 1))
                    self.status_var.set(f"0 / {payload} Dateien verarbeitet")
                elif kind == "row":
                    index, title, genre, country = payload
                    self.tree.insert("", "end", values=(index, title, genre, country))
                    self.rows.append({"nummer": index, "titel": title, "genre": genre, "land": country})
                    self.progress.configure(value=index)
                    total = int(self.progress["maximum"])
                    self.status_var.set(f"{index} / {total} Dateien verarbeitet")
                elif kind == "done":
                    self.start_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")
                    if self.stop_requested:
                        self.status_var.set(f"Abgebrochen - {len(self.rows)} Filme erfasst.")
                    else:
                        self.status_var.set(f"Fertig - {len(self.rows)} Filme erfasst.")
        except queue.Empty:
            pass
        self.after(150, self._poll_queue)

    def _export(self):
        if not self.rows:
            messagebox.showinfo(APP_NAME, "Keine Daten zum Exportieren vorhanden.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[
                ("CSV-Datei", "*.csv"),
                ("Excel-Datei", "*.xlsx"),
                ("PDF-Datei", "*.pdf"),
                ("Text-Datei", "*.txt"),
            ],
            initialfile="filmkatalog.csv",
            title="Exportieren als…",
        )
        if not path:
            return

        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == ".xlsx":
                self._export_xlsx(path)
            elif ext == ".pdf":
                self._export_pdf(path)
            elif ext == ".txt":
                self._export_txt(path)
            else:
                self._export_csv(path)
        except ImportError as exc:
            messagebox.showerror(
                APP_NAME,
                f"Für dieses Format wird ein zusätzliches Python-Paket benötigt, "
                f"das nicht installiert ist:\n{exc}",
            )
            return
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"Export fehlgeschlagen:\n{exc}")
            return

        messagebox.showinfo(APP_NAME, f"Export abgeschlossen:\n{path}")

    def _export_csv(self, path: str):
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=["nummer", "titel", "genre", "land"])
            writer.writeheader()
            writer.writerows(self.rows)

    def _export_txt(self, path: str):
        headers = ["Nummer", "Titel", "Genre", "Land"]
        col_width = {
            "nummer": max(len(headers[0]), *(len(str(r["nummer"])) for r in self.rows)),
            "titel": max(len(headers[1]), *(len(r["titel"]) for r in self.rows)),
            "genre": max(len(headers[2]), *(len(r["genre"]) for r in self.rows)),
            "land": max(len(headers[3]), *(len(r["land"]) for r in self.rows)),
        }

        def fmt_row(nummer, titel, genre, land):
            return (
                f"{str(nummer).ljust(col_width['nummer'])}  "
                f"{titel.ljust(col_width['titel'])}  "
                f"{genre.ljust(col_width['genre'])}  "
                f"{land.ljust(col_width['land'])}"
            )

        with open(path, "w", encoding="utf-8") as f:
            f.write(fmt_row(*headers) + "\n")
            f.write("-" * (sum(col_width.values()) + 6) + "\n")
            for r in self.rows:
                f.write(fmt_row(r["nummer"], r["titel"], r["genre"], r["land"]) + "\n")

    def _export_xlsx(self, path: str):
        from openpyxl import Workbook
        from openpyxl.styles import Font

        wb = Workbook()
        ws = wb.active
        ws.title = "Filmkatalog"

        headers = ["Nummer", "Titel", "Genre", "Land"]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)

        for r in self.rows:
            ws.append([r["nummer"], r["titel"], r["genre"], r["land"]])

        widths = {"A": 10, "B": 45, "C": 28, "D": 22}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        wb.save(path)

    def _export_pdf(self, path: str):
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
        from reportlab.lib.units import cm

        doc = SimpleDocTemplate(path, pagesize=landscape(A4),
                                 leftMargin=1.5 * cm, rightMargin=1.5 * cm,
                                 topMargin=1.5 * cm, bottomMargin=1.5 * cm)

        data = [["Nummer", "Titel", "Genre", "Land"]]
        for r in self.rows:
            data.append([str(r["nummer"]), r["titel"], r["genre"], r["land"]])

        table = Table(data, colWidths=[2 * cm, 10 * cm, 6 * cm, 5 * cm], repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#3a3c40")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f0f0f0")]),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))

        doc.build([table])


if __name__ == "__main__":
    app = FilmKatalogApp()
    app.mainloop()
