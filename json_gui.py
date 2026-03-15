#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI zum Bearbeiten und Anlegen von Typen in Typdaten.json
"""

import json
import os
import shutil
import tkinter as tk
from datetime import datetime
from tkinter import ttk, messagebox

try:
    from PIL import Image, ImageDraw, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

JSON_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Typdaten.json")

# ─── Standardlimits (werden über Tab "Einrichter" angepasst) ────────────────
DEFAULT_LIMITS = {
    "Diavite": {"X": {"min": 0, "max": 2000}, "Y": {"min": 0, "max": 500}, "Z": {"min": 0, "max": 200}},
    "CaptureImage": {"X": {"min": 0, "max": 2000}, "Y": {"min": 0, "max": 500}, "Z": {"min": 0, "max": 200}},
}
LIMITS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Limits.json")

# ─── Referenzwerte für die Positionsberechnung ───────────────────────────────
# Der X-Achsen-Ursprung entspricht der ersten Zeile (Row-Richtung).
# Der Y-Achsen-Ursprung entspricht der ersten Spalte (Column-Richtung).
# WICHTIG: X-Achse = Zeilen-Richtung, Y-Achse = Spalten-Richtung.
GRID_ORIGIN_X = 935.7252861842105   # Start-X (erste Zeile)
GRID_ORIGIN_Y = 48.0                # Start-Y (erste Spalte)
DEFAULT_GRID_WIDTH = 177.5394276316  # Standard-Breite (X-Span, gerundet)
DEFAULT_GRID_HEIGHT = 162.0         # Standard-Höhe  (Y-Span)
DEFAULT_CAPTURE_Z = 17.83           # Fest definierte Z-Höhe für CaptureImage
DEFAULT_TILE_WIDTH = 120            # Standard-Einzelbild-Breite (px)
DEFAULT_TILE_HEIGHT = 120           # Standard-Einzelbild-Höhe  (px)


# ═══════════════════════════════════════════════════════════════════════════════
#  Hilfsfunktionen
# ═══════════════════════════════════════════════════════════════════════════════
def load_json():
    with open(JSON_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def _format_json(data):
    """Formatiert die JSON-Daten so, dass jeder Positions-Eintrag auf einer Zeile steht.

    Erzeugt Tab-basierte Einrückung auf drei Ebenen (Typ, Felder, Positions-Einträge).
    Jeder Eintrag im Positions-Array wird als kompakte Zeile mit Leerzeichen
    innerhalb der geschweiften Klammern ausgegeben, z.B.:
        { "PathPosAction": "CaptureImage", "Position": { "X": 1.0, "Y": 2.0, "Z": 3.0 }, ... }
    """
    lines = []
    lines.append("{")
    type_keys = list(data.keys())
    for ti, typ_key in enumerate(type_keys):
        typ = data[typ_key]
        lines.append(f'\t"{typ_key}": {{')
        lines.append(f'\t\t"Description": {json.dumps(typ["Description"], ensure_ascii=False)},')
        lines.append(f'\t\t"ImageGrid": {json.dumps(typ["ImageGrid"], ensure_ascii=False)},')
        lines.append('\t\t"Positions": [')
        positions = typ["Positions"]
        for pi, pos in enumerate(positions):
            entry = json.dumps(pos, ensure_ascii=False)
            # Leerzeichen innerhalb der geschweiften Klammern einfügen
            entry = entry.replace("{", "{ ").replace("}", " }")
            comma = "," if pi < len(positions) - 1 else ""
            lines.append(f"\t\t\t{entry}{comma}")
        lines.append("\t\t]")
        comma = "," if ti < len(type_keys) - 1 else ""
        lines.append(f"\t}}{comma}")
        if ti < len(type_keys) - 1:
            lines.append("")
    lines.append("}")
    return "\n".join(lines) + "\n"

def save_json(data):
    if os.path.exists(JSON_FILE):
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        backup_path = os.path.join(
            os.path.dirname(JSON_FILE),
            f"Typdaten_backup_{timestamp}.json",
        )
        shutil.copy2(JSON_FILE, backup_path)
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        f.write(_format_json(data))

def load_limits():
    if os.path.exists(LIMITS_FILE):
        with open(LIMITS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return DEFAULT_LIMITS

def save_limits(limits):
    with open(LIMITS_FILE, "w", encoding="utf-8") as f:
        json.dump(limits, f, indent="\t", ensure_ascii=False)

def parse_grid(grid_str):
    """'5x4' -> (rows=5, cols=4)"""
    parts = grid_str.lower().split("x")
    return int(parts[0]), int(parts[1])

def validate_position(pos, category, limits):
    """Prüft ob X/Y/Z innerhalb der Limits liegen. Gibt Fehlerliste zurück."""
    errors = []
    cat_limits = limits.get(category, {})
    for axis in ("X", "Y", "Z"):
        val = pos.get(axis, 0)
        lo = cat_limits.get(axis, {}).get("min", float("-inf"))
        hi = cat_limits.get(axis, {}).get("max", float("inf"))
        if not (lo <= val <= hi):
            errors.append(f"{category} {axis}={val} ausserhalb [{lo}, {hi}]")
    return errors

# ═══════════════════════════════════════════════════════════════════════════════
#  Haupt-App
# ═══════════════════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Typdaten Editor")
        self.geometry("1100x750")
        self.resizable(True, True)

        self.data = load_json()
        self.limits = load_limits()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=6)

        # --- Tab 1: Bearbeiten ---
        self.tab_edit = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_edit, text="Typ bearbeiten")
        self._build_edit_tab()

        # --- Tab 2: Neuer Typ ---
        self.tab_new = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_new, text="Neuen Typ anlegen")
        self._build_new_tab()

        # --- Tab 3: Neuer Typ (Eckpunkte) ---
        self.tab_corners = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_corners, text="Neuer Typ (Eckpunkte)")
        self._build_corners_tab()

        # --- Tab 4: Einrichter (Limits) ---
        self.tab_limits = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_limits, text="Einrichter – Limits")
        self._build_limits_tab()

    # ───────────────────────────────────────────────────────────────────────
    #  TAB 1 – Typ bearbeiten
    # ───────────────────────────────────────────────────────────────────────
    def _build_edit_tab(self):
        top = ttk.Frame(self.tab_edit)
        top.pack(fill="x", padx=8, pady=6)

        ttk.Label(top, text="Typ auswählen:").pack(side="left")
        self.edit_type_var = tk.StringVar()
        self.edit_combo = ttk.Combobox(top, textvariable=self.edit_type_var,
                                       values=sorted(self.data.keys()), state="readonly", width=20)
        self.edit_combo.pack(side="left", padx=6)
        self.edit_combo.bind("<<ComboboxSelected>>", self._on_type_selected)

        ttk.Button(top, text="Quit", command=self.destroy).pack(side="right")

        # Scrollbarer Bereich
        container = ttk.Frame(self.tab_edit)
        container.pack(fill="both", expand=True)

        self.edit_canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.edit_canvas.yview)
        self.edit_scroll_frame = ttk.Frame(self.edit_canvas)
        self.edit_scroll_frame.bind("<Configure>",
                                    lambda e: self.edit_canvas.configure(scrollregion=self.edit_canvas.bbox("all")))
        self.edit_canvas.create_window((0, 0), window=self.edit_scroll_frame, anchor="nw")
        self.edit_canvas.configure(yscrollcommand=scrollbar.set)
        self.edit_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.edit_canvas.bind_all("<MouseWheel>",
                                  lambda e: self.edit_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        # Speichern- und Löschen-Buttons
        btn_frame = ttk.Frame(self.tab_edit)
        btn_frame.pack(fill="x", padx=8, pady=6)
        ttk.Button(btn_frame, text="💾  Änderungen speichern", command=self._save_edit).pack(side="right")
        ttk.Button(btn_frame, text="🗑 Typ löschen", command=self._delete_type).pack(side="left")

        # Datencontainer
        self.edit_diavite_entries = []  # [(action_name, (x_var, y_var, z_var)), ...]
        self.edit_capture_entries = {}  # {(row, col): (x_var, y_var, z_var)}
        self.edit_prestart_entry = None
        self.edit_process_entry = None
        self.current_edit_type = None
        self.edit_imagegrid_var = tk.StringVar()

    def _on_type_selected(self, _event=None):
        typ_key = self.edit_type_var.get()
        if not typ_key:
            return
        typ = self.data[typ_key]
        self.current_edit_type = typ_key

        # Alten Inhalt löschen
        for w in self.edit_scroll_frame.winfo_children():
            w.destroy()
        self.edit_diavite_entries.clear()
        self.edit_capture_entries.clear()

        # --- Description & Grid ---
        info = ttk.LabelFrame(self.edit_scroll_frame, text="Info")
        info.pack(fill="x", padx=8, pady=4)
        ttk.Label(info, text=f"Description: {typ['Description']}").pack(anchor="w", padx=4, pady=2)
        grid_row = ttk.Frame(info)
        grid_row.pack(anchor="w", padx=4, pady=2)
        ttk.Label(grid_row, text="ImageGrid:").pack(side="left")
        self.edit_imagegrid_var = tk.StringVar(value=typ["ImageGrid"])
        grid_entry = ttk.Entry(grid_row, textvariable=self.edit_imagegrid_var, width=10)
        grid_entry.pack(side="left", padx=6)
        grid_entry.bind("<FocusOut>", self._on_imagegrid_changed)
        grid_entry.bind("<Return>", self._on_imagegrid_changed)
        ttk.Label(grid_row, text="(Format: NxM, z.B. 4x4)").pack(side="left", padx=4)

        # --- Bild-Container (Platzhalter für 3 Bildvorschauen) ---
        self.edit_img_container = ttk.Frame(self.edit_scroll_frame)
        self.edit_img_container.pack(fill="x", padx=8, pady=4)
        self._populate_edit_images()

        # --- Diavite A & B ---
        diav_frame = ttk.LabelFrame(self.edit_scroll_frame, text="Diavite Messungen")
        diav_frame.pack(fill="x", padx=8, pady=4)
        for pos in typ["Positions"]:
            if pos["PathPosAction"] in ("DiaviteMeasurementA", "DiaviteMeasurementB"):
                label = pos["PathPosAction"].replace("DiaviteMeasurement", "Diavite ")
                row_f = ttk.Frame(diav_frame)
                row_f.pack(fill="x", padx=4, pady=2)
                ttk.Label(row_f, text=label, width=14).pack(side="left")
                vars_ = self._xyz_entries(row_f, pos["Position"])
                self.edit_diavite_entries.append((pos["PathPosAction"], vars_))

        # --- PreStart ---
        for pos in typ["Positions"]:
            if pos["PathPosAction"] == "PreStart":
                ps_frame = ttk.LabelFrame(self.edit_scroll_frame, text="PreStart")
                ps_frame.pack(fill="x", padx=8, pady=4)
                row_f = ttk.Frame(ps_frame)
                row_f.pack(fill="x", padx=4, pady=2)
                self.edit_prestart_entry = self._xyz_entries(row_f, pos["Position"])

        # --- CaptureImage Matrix (Platzhalter) ---
        rows, cols = parse_grid(typ["ImageGrid"])
        self.edit_cap_outer = ttk.Frame(self.edit_scroll_frame)
        self.edit_cap_outer.pack(fill="x", padx=8, pady=4)
        self._populate_edit_cap_matrix(rows, cols, typ)

        # --- Process ---
        for pos in typ["Positions"]:
            if pos["PathPosAction"] == "Process":
                pr_frame = ttk.LabelFrame(self.edit_scroll_frame, text="Process")
                pr_frame.pack(fill="x", padx=8, pady=4)
                row_f = ttk.Frame(pr_frame)
                row_f.pack(fill="x", padx=4, pady=2)
                self.edit_process_entry = self._xyz_entries(row_f, pos["Position"])

    def _build_three_images(self, parent, img, grid_str, photo_list):
        """Baut 3 Bildvorschauen in parent: skaliertes Original, Einzelbilder, Grid-Mosaik (3px Abstand).

        photo_list: Liste, in die ImageTk.PhotoImage-Referenzen eingefügt werden (gegen GC).
        """
        MAX_W = 400
        img_w, img_h = img.size
        scale = min(1.0, MAX_W / img_w)
        disp_w = max(1, int(img_w * scale))
        disp_h = max(1, int(img_h * scale))
        disp_img = img.resize((disp_w, disp_h), Image.LANCZOS)

        try:
            rows, cols = parse_grid(grid_str)
            if rows <= 0 or cols <= 0:
                raise ValueError("rows/cols müssen positiv sein")
        except Exception:
            rows, cols = 1, 1

        # Grid-Spacing: gleichmäßiger Abstand zwischen Kachel-Starts (in Original-Pixeln)
        step_x = img_w / cols
        step_y = img_h / rows

        # Tile-Größe aus Limits (mit Fallback) in Original-Pixeln
        tile_w = self.limits.get("TileWidth", DEFAULT_TILE_WIDTH)
        tile_h = self.limits.get("TileHeight", DEFAULT_TILE_HEIGHT)

        # Display-Tile-Größe für Einzelbilder (skaliert, inkl. Überlapp-Faktor)
        tile_w_px = max(1, int(tile_w * scale))
        tile_h_px = max(1, int(tile_h * scale))

        # Bild 1: Skaliertes Originalbild
        lf1 = ttk.LabelFrame(parent, text="Originalbild")
        lf1.pack(side="left", padx=4, anchor="n")
        ph1 = ImageTk.PhotoImage(disp_img)
        lbl1 = ttk.Label(lf1, image=ph1)
        lbl1.image = ph1
        lbl1.pack()
        photo_list.append(ph1)

        # Bild 2: Einzelbilder nebeneinander im Grid-Layout (mit Überlapp falls tile_w > step_x)
        lf2 = ttk.LabelFrame(parent, text=f"Einzelbilder ({grid_str})")
        lf2.pack(side="left", padx=4, anchor="n")
        for r in range(rows):
            row_f = ttk.Frame(lf2)
            row_f.pack(anchor="w")
            for c in range(cols):
                x0 = int(c * step_x)
                y0 = int(r * step_y)
                x1 = min(x0 + tile_w, img_w)
                y1 = min(y0 + tile_h, img_h)
                tile = img.crop((x0, y0, x1, y1))
                tile_s = tile.resize((tile_w_px, tile_h_px), Image.LANCZOS)
                ph = ImageTk.PhotoImage(tile_s)
                lbl = ttk.Label(row_f, image=ph)
                lbl.image = ph
                lbl.pack(side="left")
                photo_list.append(ph)

        # Bild 3: Überlapp-Visualisierung – Gesamtbild mit eingezeichneten Kachel-Positionen
        overlay_base = disp_img.convert("RGBA")
        overlay = Image.new("RGBA", overlay_base.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)
        tile_w_s = max(1, int(tile_w * scale))
        tile_h_s = max(1, int(tile_h * scale))
        for r in range(rows):
            for c in range(cols):
                x0_s = int(c * step_x * scale)
                y0_s = int(r * step_y * scale)
                x1_s = min(x0_s + tile_w_s, disp_w)
                y1_s = min(y0_s + tile_h_s, disp_h)
                draw.rectangle([x0_s, y0_s, x1_s, y1_s], fill=(100, 150, 255, 60))
                draw.rectangle([x0_s, y0_s, x1_s, y1_s], outline=(0, 80, 200, 180), width=2)
        result3 = Image.alpha_composite(overlay_base, overlay).convert("RGB")
        lf3 = ttk.LabelFrame(parent, text=f"Kachel-Lage ({grid_str})")
        lf3.pack(side="left", padx=4, anchor="n")
        ph3 = ImageTk.PhotoImage(result3)
        lbl3 = ttk.Label(lf3, image=ph3)
        lbl3.image = ph3
        lbl3.pack()
        photo_list.append(ph3)

    def _populate_edit_images(self):
        """Befüllt self.edit_img_container mit 3 Bildvorschauen (Original, Einzelbilder, Mosaik)."""
        for w in self.edit_img_container.winfo_children():
            w.destroy()
        self._grid_photos = []

        if not PIL_AVAILABLE:
            return

        typ_key = self.current_edit_type
        grid_str = self.edit_imagegrid_var.get().strip()
        base_dir = os.path.dirname(os.path.abspath(__file__))
        img_path = None
        for ext in (".png", ".jpg", ".jpeg", ".bmp"):
            candidate = os.path.join(base_dir, typ_key + ext)
            if os.path.isfile(candidate):
                img_path = candidate
                break
        if img_path is None:
            return

        try:
            img = Image.open(img_path)
        except Exception:
            return

        img_lf = ttk.LabelFrame(self.edit_img_container, text="Typbild")
        img_lf.pack(fill="x")
        content = ttk.Frame(img_lf)
        content.pack(fill="x", padx=4, pady=4)
        self._build_three_images(content, img, grid_str, self._grid_photos)

    def _populate_edit_cap_matrix(self, rows, cols, typ):
        """Befüllt self.edit_cap_outer mit der CaptureImage-Matrix (rows × cols)."""
        for w in self.edit_cap_outer.winfo_children():
            w.destroy()
        self.edit_capture_entries.clear()

        grid_str = f"{rows}x{cols}"
        cap_frame = ttk.LabelFrame(self.edit_cap_outer, text=f"CaptureImage  ({grid_str})")
        cap_frame.pack(fill="x")

        header = ttk.Frame(cap_frame)
        header.pack(fill="x", padx=4, pady=2)
        ttk.Label(header, text="", width=10).pack(side="left")
        for c in range(cols):
            ttk.Label(header, text=f"Col {c}", width=28, anchor="center").pack(side="left", padx=2)

        for r in range(rows):
            row_frame = ttk.Frame(cap_frame)
            row_frame.pack(fill="x", padx=4, pady=1)
            ttk.Label(row_frame, text=f"Row {r}", width=10).pack(side="left")
            for c in range(cols):
                cell_frame = ttk.Frame(row_frame, relief="groove", borderwidth=1)
                cell_frame.pack(side="left", padx=2, pady=1)
                p = self._find_capture(typ["Positions"], r, c)
                pos_data = p["Position"] if p else {"X": 0, "Y": 0, "Z": 0}
                vars_ = self._xyz_entries_compact(cell_frame, pos_data)
                self.edit_capture_entries[(r, c)] = vars_

    def _on_imagegrid_changed(self, _event=None):
        """Wird aufgerufen wenn das ImageGrid-Feld geändert wird (FocusOut / Return)."""
        grid_str = self.edit_imagegrid_var.get().strip()
        try:
            rows, cols = parse_grid(grid_str)
            if rows <= 0 or cols <= 0:
                raise ValueError("rows/cols müssen positiv sein")
        except Exception:
            messagebox.showwarning("Eingabe", "Ungültiges Grid-Format. Bitte z.B. '4x4' eingeben.")
            return

        typ = self.data.get(self.current_edit_type)
        if typ is None:
            return

        self._populate_edit_images()
        self._populate_edit_cap_matrix(rows, cols, typ)

    def _delete_type(self):
        """Löscht den aktuell ausgewählten Typ nach Bestätigung."""
        typ_key = self.edit_type_var.get()
        if not typ_key:
            messagebox.showwarning("Kein Typ", "Bitte zuerst einen Typ auswählen.")
            return
        if not messagebox.askyesno("Typ löschen", f"Typ {typ_key} wirklich löschen?"):
            return

        del self.data[typ_key]
        save_json(self.data)

        # Dropdown aktualisieren
        self.edit_combo["values"] = sorted(self.data.keys())
        self.edit_type_var.set("")

        # Bearbeitungsbereich leeren
        for w in self.edit_scroll_frame.winfo_children():
            w.destroy()
        self.edit_diavite_entries.clear()
        self.edit_capture_entries.clear()
        self.current_edit_type = None

        messagebox.showinfo("Gelöscht", f"Typ {typ_key} wurde gelöscht.")

    def _save_edit(self):
        typ_key = self.edit_type_var.get()
        if not typ_key:
            messagebox.showwarning("Kein Typ", "Bitte zuerst einen Typ auswählen.")
            return

        typ = self.data[typ_key]
        errors = []

        # Diavite
        for action_name, (xv, yv, zv) in self.edit_diavite_entries:
            try:
                pos = {"X": float(xv.get()), "Y": float(yv.get()), "Z": float(zv.get())}
            except ValueError:
                errors.append(f"{action_name}: ungültige Zahlenwerte")
                continue
            errors.extend(validate_position(pos, "Diavite", self.limits))
            for p in typ["Positions"]:
                if p["PathPosAction"] == action_name:
                    p["Position"] = pos

        # PreStart
        if self.edit_prestart_entry:
            xv, yv, zv = self.edit_prestart_entry
            try:
                pos = {"X": float(xv.get()), "Y": float(yv.get()), "Z": float(zv.get())}
                for p in typ["Positions"]:
                    if p["PathPosAction"] == "PreStart":
                        p["Position"] = pos
            except ValueError:
                errors.append("PreStart: ungültige Zahlenwerte")

        # CaptureImage – rebuild from current edit entries (handles grid changes correctly)
        grid_str = self.edit_imagegrid_var.get().strip()
        try:
            ec_rows, ec_cols = parse_grid(grid_str)
            if ec_rows <= 0 or ec_cols <= 0:
                raise ValueError("rows/cols müssen positiv sein")
        except Exception:
            errors.append("ImageGrid: Ungültiges Format (erwartet z.B. '4x4')")
            ec_rows, ec_cols = 0, 0

        cap_positions_new = []
        for r in range(ec_rows):
            # Serpentinen-Muster: gerade Zeilen links→rechts, ungerade rechts→links
            col_range = range(ec_cols) if r % 2 == 0 else range(ec_cols - 1, -1, -1)
            for c in col_range:
                if (r, c) not in self.edit_capture_entries:
                    continue
                xv, yv, zv = self.edit_capture_entries[(r, c)]
                try:
                    pos = {"X": float(xv.get()), "Y": float(yv.get()), "Z": float(zv.get())}
                except ValueError:
                    errors.append(f"CaptureImage Row {r} Col {c}: ungültige Zahlenwerte")
                    continue
                errors.extend(validate_position(pos, "CaptureImage", self.limits))
                cap_positions_new.append({
                    "PathPosAction": "CaptureImage",
                    "Position": pos,
                    "CellRow": r,
                    "CellColumn": c,
                })

        # Process
        if self.edit_process_entry:
            xv, yv, zv = self.edit_process_entry
            try:
                pos = {"X": float(xv.get()), "Y": float(yv.get()), "Z": float(zv.get())}
                for p in typ["Positions"]:
                    if p["PathPosAction"] == "Process":
                        p["Position"] = pos
            except ValueError:
                errors.append("Process: ungültige Zahlenwerte")

        if errors:
            messagebox.showerror("Plausibilitätsprüfung fehlgeschlagen", "\n".join(errors))
            return

        # ImageGrid aktualisieren
        typ["ImageGrid"] = grid_str

        # CaptureImage-Positionen neu aufbauen (in Serpentinen-Reihenfolge, nach PreStart)
        non_cap = [p for p in typ["Positions"] if p["PathPosAction"] != "CaptureImage"]
        new_positions = []
        cap_inserted = False
        for p in non_cap:
            new_positions.append(p)
            if p["PathPosAction"] == "PreStart" and not cap_inserted:
                new_positions.extend(cap_positions_new)
                cap_inserted = True
        if not cap_inserted:
            new_positions.extend(cap_positions_new)
        typ["Positions"] = new_positions

        save_json(self.data)
        messagebox.showinfo("Gespeichert", f"Typ {typ_key} erfolgreich in Typdaten.json gespeichert.")

    # ───────────────────────────────────────────────────────────────────────
    #  TAB 2 – Neuen Typ anlegen
    # ───────────────────────────────────────────────────────────────────────
    def _build_new_tab(self):
        self.new_step = 0  # 0=Typnummer, 1=ImageGrid+Vorschau, 2=Maße, 3=Positionen

        # Container für Wizard-Schritte
        self.new_wizard = ttk.Frame(self.tab_new)
        self.new_wizard.pack(fill="both", expand=True, padx=8, pady=6)

        self._new_show_step0()

    def _clear_wizard(self):
        for w in self.new_wizard.winfo_children():
            w.destroy()

    def _new_show_step0(self):
        self._clear_wizard()
        self.new_step = 0
        # Reset step-specific vars so they are recreated fresh on next run
        for attr in ("new_grid_var", "new_width_var", "new_height_var",
                     "new_origin_x_var", "new_origin_y_var", "new_preview_frame"):
            if hasattr(self, attr):
                delattr(self, attr)
        ttk.Label(self.new_wizard, text="Schritt 1: Typnummer eingeben", font=("", 12, "bold")).pack(anchor="w",
                                                                                                      pady=(0, 8))
        f = ttk.Frame(self.new_wizard)
        f.pack(fill="x")
        ttk.Label(f, text="Typnummer:").pack(side="left")
        self.new_typnr_var = tk.StringVar()
        ttk.Entry(f, textvariable=self.new_typnr_var, width=20).pack(side="left", padx=6)
        ttk.Button(self.new_wizard, text="Weiter ➜", command=self._new_goto_step1).pack(anchor="e", pady=10)

    def _new_goto_step1(self):
        tnr = self.new_typnr_var.get().strip()
        if not tnr:
            messagebox.showwarning("Eingabe", "Bitte Typnummer eingeben.")
            return
        if tnr in self.data:
            messagebox.showwarning("Duplikat", f"Typ {tnr} existiert bereits.")
            return
        self._new_display_step1()

    def _new_display_step1(self):
        """Schritt 2: ImageGrid eingeben + Bildvorschau."""
        self._clear_wizard()
        self.new_step = 1
        tnr = self.new_typnr_var.get().strip()
        ttk.Label(self.new_wizard, text=f"Schritt 2: ImageGrid für Typ {tnr}",
                  font=("", 12, "bold")).pack(anchor="w", pady=(0, 8))
        f = ttk.Frame(self.new_wizard)
        f.pack(fill="x")
        ttk.Label(f, text="ImageGrid (z.B. 4x4):").pack(side="left")
        if not hasattr(self, "new_grid_var"):
            self.new_grid_var = tk.StringVar()
            self.new_grid_var.trace_add("write", lambda *_a: self._new_update_preview())
        ttk.Entry(f, textvariable=self.new_grid_var, width=10).pack(side="left", padx=6)

        # Preview-Frame (wird von _new_update_preview befüllt)
        self.new_preview_frame = ttk.Frame(self.new_wizard)
        self.new_preview_frame.pack(fill="x", pady=(8, 0))

        bf = ttk.Frame(self.new_wizard)
        bf.pack(fill="x", pady=10)
        ttk.Button(bf, text="⬅ Zurück", command=self._new_show_step0).pack(side="left")
        ttk.Button(bf, text="Weiter ➜", command=self._new_goto_step2).pack(side="right")

        # Vorschau anzeigen falls bereits ein Wert vorhanden (z.B. beim Zurücknavigieren)
        self._new_update_preview()

    def _new_update_preview(self):
        """Sucht ein passendes Bild und zeigt 3 Bildvorschauen an (Original, Einzelbilder, Grid-Mosaik)."""
        if not hasattr(self, "new_preview_frame") or not self.new_preview_frame.winfo_exists():
            return
        for w in self.new_preview_frame.winfo_children():
            w.destroy()
        self._new_img_photos = []

        grid_str = self.new_grid_var.get().strip()
        if not grid_str:
            return
        try:
            rows, cols = parse_grid(grid_str)
            if rows <= 0 or cols <= 0:
                return
        except Exception:
            return

        if not PIL_AVAILABLE:
            ttk.Label(self.new_preview_frame,
                      text="Pillow (PIL) nicht verfügbar – Bildvorschau nicht möglich.").pack(anchor="w")
            return

        tnr = self.new_typnr_var.get().strip()
        script_dir = os.path.dirname(os.path.abspath(__file__))
        img_path = None
        for fname in os.listdir(script_dir):
            if fname.lower().endswith(".png"):
                stem = os.path.splitext(fname)[0]
                if stem in tnr:
                    img_path = os.path.join(script_dir, fname)
                    break
        if img_path is None:
            return

        try:
            original_img = Image.open(img_path)
        except Exception:
            return

        self._build_three_images(self.new_preview_frame, original_img, grid_str, self._new_img_photos)

    def _new_goto_step2(self):
        """Schritt 2 → 3: ImageGrid validieren, dann Maße-Schritt anzeigen."""
        grid_str = self.new_grid_var.get().strip()
        try:
            rows, cols = parse_grid(grid_str)
            assert rows > 0 and cols > 0
        except Exception:
            messagebox.showwarning("Eingabe", "Ungültiges Grid-Format. Bitte z.B. '4x4' eingeben.")
            return
        self._new_display_step2()

    def _new_display_step2(self):
        """Schritt 3: Ursprung X/Y, Breite und Höhe eingeben."""
        self._clear_wizard()
        self.new_step = 2
        tnr = self.new_typnr_var.get().strip()
        ttk.Label(self.new_wizard, text=f"Schritt 3: Maße für Typ {tnr}",
                  font=("", 12, "bold")).pack(anchor="w", pady=(0, 8))

        f_origin = ttk.Frame(self.new_wizard)
        f_origin.pack(fill="x", pady=(4, 0))
        ttk.Label(f_origin, text="Ursprung X:").pack(side="left")
        if not hasattr(self, "new_origin_x_var"):
            self.new_origin_x_var = tk.StringVar(value=str(GRID_ORIGIN_X))
        ttk.Entry(f_origin, textvariable=self.new_origin_x_var, width=20).pack(side="left", padx=6)
        ttk.Label(f_origin, text="Ursprung Y:").pack(side="left", padx=(8, 0))
        if not hasattr(self, "new_origin_y_var"):
            self.new_origin_y_var = tk.StringVar(value=str(GRID_ORIGIN_Y))
        ttk.Entry(f_origin, textvariable=self.new_origin_y_var, width=20).pack(side="left", padx=6)

        f2 = ttk.Frame(self.new_wizard)
        f2.pack(fill="x", pady=(4, 0))
        ttk.Label(f2, text="Breite:").pack(side="left")
        if not hasattr(self, "new_width_var"):
            self.new_width_var = tk.StringVar(value=str(DEFAULT_GRID_WIDTH))
        ttk.Entry(f2, textvariable=self.new_width_var, width=20).pack(side="left", padx=6)
        ttk.Label(f2, text="Höhe:").pack(side="left", padx=(8, 0))
        if not hasattr(self, "new_height_var"):
            self.new_height_var = tk.StringVar(value=str(DEFAULT_GRID_HEIGHT))
        ttk.Entry(f2, textvariable=self.new_height_var, width=20).pack(side="left", padx=6)

        bf = ttk.Frame(self.new_wizard)
        bf.pack(fill="x", pady=10)
        ttk.Button(bf, text="⬅ Zurück", command=self._new_display_step1).pack(side="left")
        ttk.Button(bf, text="Weiter ➜", command=self._new_goto_step3).pack(side="right")

    def _new_goto_step3(self):
        """Schritt 3 → 4: Maße validieren, dann Positionen-Schritt anzeigen."""
        grid_str = self.new_grid_var.get().strip()
        rows, cols = parse_grid(grid_str)

        try:
            width = float(self.new_width_var.get())
            height = float(self.new_height_var.get())
            assert width > 0 and height > 0
        except Exception:
            messagebox.showwarning("Eingabe", "Breite und Höhe müssen positive Zahlen sein.")
            return

        try:
            origin_x = float(self.new_origin_x_var.get())
            origin_y = float(self.new_origin_y_var.get())
        except Exception:
            messagebox.showwarning("Eingabe", "Ursprung X und Y müssen gültige Zahlen sein.")
            return

        self._clear_wizard()
        self.new_step = 3
        tnr = self.new_typnr_var.get().strip()

        ttk.Label(self.new_wizard, text=f"Schritt 4: Positionen für Typ {tnr} ({grid_str})",
                  font=("", 12, "bold")).pack(anchor="w", pady=(0, 8))

        # Scrollbar
        container = ttk.Frame(self.new_wizard)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container)
        sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        # Diavite
        diav_frame = ttk.LabelFrame(scroll_frame, text="Diavite Messungen")
        diav_frame.pack(fill="x", padx=4, pady=4)
        self.new_diavite = []
        for label in ("Diavite A", "Diavite B"):
            rf = ttk.Frame(diav_frame)
            rf.pack(fill="x", padx=4, pady=2)
            ttk.Label(rf, text=label, width=14).pack(side="left")
            self.new_diavite.append(self._xyz_entries(rf, {"X": 0, "Y": 0, "Z": 0}))

        # PreStart
        ps_frame = ttk.LabelFrame(scroll_frame, text="PreStart")
        ps_frame.pack(fill="x", padx=4, pady=4)
        rf = ttk.Frame(ps_frame)
        rf.pack(fill="x", padx=4, pady=2)
        self.new_prestart = self._xyz_entries(rf, {"X": 0, "Y": 0, "Z": 0})

        # CaptureImage Matrix
        cap_frame = ttk.LabelFrame(scroll_frame, text=f"CaptureImage  ({grid_str})")
        cap_frame.pack(fill="x", padx=4, pady=4)

        header = ttk.Frame(cap_frame)
        header.pack(fill="x", padx=4, pady=2)
        ttk.Label(header, text="", width=10).pack(side="left")
        for c in range(cols):
            ttk.Label(header, text=f"Col {c}", width=28, anchor="center").pack(side="left", padx=2)

        self.new_capture_entries = {}
        for r in range(rows):
            row_frame = ttk.Frame(cap_frame)
            row_frame.pack(fill="x", padx=4, pady=1)
            ttk.Label(row_frame, text=f"Row {r}", width=10).pack(side="left")
            for c in range(cols):
                cell = ttk.Frame(row_frame, relief="groove", borderwidth=1)
                cell.pack(side="left", padx=2, pady=1)
                px, py, pz = self._get_initial_position(rows, cols, r, c, width, height, origin_x, origin_y)
                init_pos = {"X": round(px, 4), "Y": round(py, 4), "Z": round(pz, 4)}
                self.new_capture_entries[(r, c)] = self._xyz_entries_compact(cell, init_pos)

        # Process
        pr_frame = ttk.LabelFrame(scroll_frame, text="Process")
        pr_frame.pack(fill="x", padx=4, pady=4)
        rf2 = ttk.Frame(pr_frame)
        rf2.pack(fill="x", padx=4, pady=2)
        self.new_process = self._xyz_entries(rf2, {"X": 0, "Y": 0, "Z": 0})

        # Buttons
        bf = ttk.Frame(self.new_wizard)
        bf.pack(fill="x", pady=6)
        ttk.Button(bf, text="⬅ Zurück", command=self._new_display_step2).pack(side="left")
        ttk.Button(bf, text="💾  Neuen Typ speichern", command=lambda: self._save_new(rows, cols, grid_str)).pack(
            side="right")

    def _save_new(self, rows, cols, grid_str):
        tnr = self.new_typnr_var.get().strip()
        errors = []
        positions = []

        # Diavite
        actions = ["DiaviteMeasurementA", "DiaviteMeasurementB"]
        for i, (xv, yv, zv) in enumerate(self.new_diavite):
            try:
                pos = {"X": float(xv.get()), "Y": float(yv.get()), "Z": float(zv.get())}
            except ValueError:
                errors.append(f"{actions[i]}: ungültige Zahlenwerte")
                continue
            errors.extend(validate_position(pos, "Diavite", self.limits))
            positions.append({"PathPosAction": actions[i], "Position": pos, "CellRow": -1, "CellColumn": -1})

        # PreStart
        xv, yv, zv = self.new_prestart
        try:
            pos = {"X": float(xv.get()), "Y": float(yv.get()), "Z": float(zv.get())}
            positions.append({"PathPosAction": "PreStart", "Position": pos, "CellRow": -1, "CellColumn": -1})
        except ValueError:
            errors.append("PreStart: ungültige Zahlenwerte")

        # CaptureImage – Serpentinenmuster (wie im Original)
        for r in range(rows):
            if r % 2 == 0:
                col_range = range(cols)
            else:
                col_range = range(cols - 1, -1, -1)
            for c in col_range:
                xv, yv, zv = self.new_capture_entries[(r, c)]
                try:
                    pos = {"X": float(xv.get()), "Y": float(yv.get()), "Z": float(zv.get())}
                except ValueError:
                    errors.append(f"CaptureImage Row {r} Col {c}: ungültige Zahlenwerte")
                    continue
                errors.extend(validate_position(pos, "CaptureImage", self.limits))
                positions.append({"PathPosAction": "CaptureImage", "Position": pos, "CellRow": r, "CellColumn": c})

        # Process
        xv, yv, zv = self.new_process
        try:
            pos = {"X": float(xv.get()), "Y": float(yv.get()), "Z": float(zv.get())}
            positions.append({"PathPosAction": "Process", "Position": pos, "CellRow": -1, "CellColumn": -1})
        except ValueError:
            errors.append("Process: ungültige Zahlenwerte")

        if errors:
            messagebox.showerror("Plausibilitätsprüfung fehlgeschlagen", "\n".join(errors))
            return

        self.data[tnr] = {
            "Description": "Python Script Generated",
            "ImageGrid": grid_str,
            "Positions": positions,
        }
        save_json(self.data)

        # Dropdown in Tab 1 aktualisieren
        self.edit_combo["values"] = sorted(self.data.keys())

        messagebox.showinfo("Gespeichert", f"Neuer Typ {tnr} erfolgreich gespeichert.")
        self._new_show_step0()

    # ───────────────────────────────────────────────────────────────────────
    #  TAB 3 – Neuer Typ (Eckpunkte)
    # ───────────────────────────────────────────────────────────────────────
    def _build_corners_tab(self):
        self.corners_step = 0

        self.corners_wizard = ttk.Frame(self.tab_corners)
        self.corners_wizard.pack(fill="both", expand=True, padx=8, pady=6)

        self._corners_show_step0()

    def _clear_corners_wizard(self):
        for w in self.corners_wizard.winfo_children():
            w.destroy()

    def _corners_show_step0(self):
        self._clear_corners_wizard()
        self.corners_step = 0
        for attr in ("corners_grid_var", "corners_preview_frame"):
            if hasattr(self, attr):
                delattr(self, attr)
        ttk.Label(self.corners_wizard, text="Schritt 1: Typnummer eingeben", font=("", 12, "bold")).pack(
            anchor="w", pady=(0, 8))
        f = ttk.Frame(self.corners_wizard)
        f.pack(fill="x")
        ttk.Label(f, text="Typnummer:").pack(side="left")
        self.corners_typnr_var = tk.StringVar()
        ttk.Entry(f, textvariable=self.corners_typnr_var, width=20).pack(side="left", padx=6)
        ttk.Button(self.corners_wizard, text="Weiter ➜", command=self._corners_goto_step1).pack(anchor="e", pady=10)

    def _corners_goto_step1(self):
        tnr = self.corners_typnr_var.get().strip()
        if not tnr:
            messagebox.showwarning("Eingabe", "Bitte Typnummer eingeben.")
            return
        if tnr in self.data:
            messagebox.showwarning("Duplikat", f"Typ {tnr} existiert bereits.")
            return
        self._corners_display_step1()

    def _corners_display_step1(self):
        """Schritt 2: ImageGrid eingeben + Bildvorschau."""
        self._clear_corners_wizard()
        self.corners_step = 1
        tnr = self.corners_typnr_var.get().strip()
        ttk.Label(self.corners_wizard, text=f"Schritt 2: ImageGrid für Typ {tnr}",
                  font=("", 12, "bold")).pack(anchor="w", pady=(0, 8))
        f = ttk.Frame(self.corners_wizard)
        f.pack(fill="x")
        ttk.Label(f, text="ImageGrid (z.B. 4x4):").pack(side="left")
        if not hasattr(self, "corners_grid_var"):
            self.corners_grid_var = tk.StringVar()
            self.corners_grid_var.trace_add("write", lambda *_a: self._corners_update_preview())
        ttk.Entry(f, textvariable=self.corners_grid_var, width=10).pack(side="left", padx=6)

        # Preview-Frame (wird von _corners_update_preview befüllt)
        self.corners_preview_frame = ttk.Frame(self.corners_wizard)
        self.corners_preview_frame.pack(fill="x", pady=(8, 0))

        bf = ttk.Frame(self.corners_wizard)
        bf.pack(fill="x", pady=10)
        ttk.Button(bf, text="⬅ Zurück", command=self._corners_show_step0).pack(side="left")
        ttk.Button(bf, text="Weiter ➜", command=self._corners_goto_step2).pack(side="right")

        # Vorschau anzeigen falls bereits ein Wert vorhanden
        self._corners_update_preview()

    def _corners_update_preview(self):
        """Sucht ein passendes Bild und zeigt 3 Bildvorschauen an (Original, Einzelbilder, Kachel-Lage)."""
        if not hasattr(self, "corners_preview_frame") or not self.corners_preview_frame.winfo_exists():
            return
        for w in self.corners_preview_frame.winfo_children():
            w.destroy()
        self._corners_img_photos = []

        grid_str = self.corners_grid_var.get().strip()
        if not grid_str:
            return
        try:
            rows, cols = parse_grid(grid_str)
            if rows <= 0 or cols <= 0:
                return
        except Exception:
            return

        if not PIL_AVAILABLE:
            ttk.Label(self.corners_preview_frame,
                      text="Pillow (PIL) nicht verfügbar – Bildvorschau nicht möglich.").pack(anchor="w")
            return

        tnr = self.corners_typnr_var.get().strip()
        script_dir = os.path.dirname(os.path.abspath(__file__))
        img_path = None
        for fname in os.listdir(script_dir):
            if fname.lower().endswith(".png"):
                stem = os.path.splitext(fname)[0]
                if stem in tnr:
                    img_path = os.path.join(script_dir, fname)
                    break
        if img_path is None:
            return

        try:
            original_img = Image.open(img_path)
        except Exception:
            return

        self._build_three_images(self.corners_preview_frame, original_img, grid_str, self._corners_img_photos)

    def _corners_goto_step2(self):
        """Schritt 2 → 3: ImageGrid validieren, dann Eckpunkte-Schritt anzeigen."""
        grid_str = self.corners_grid_var.get().strip()
        try:
            rows, cols = parse_grid(grid_str)
            assert rows > 0 and cols > 0
        except Exception:
            messagebox.showwarning("Eingabe", "Ungültiges Grid-Format. Bitte z.B. '4x4' eingeben.")
            return
        self._corners_display_step2()

    def _corners_display_step2(self):
        """Schritt 3: Eckpunkte eingeben (Grid wurde bereits in Schritt 2 erfasst)."""
        self._clear_corners_wizard()
        self.corners_step = 2
        tnr = self.corners_typnr_var.get().strip()
        grid_str = self.corners_grid_var.get().strip()
        ttk.Label(self.corners_wizard, text=f"Schritt 3: Grid & Eckpunkte für Typ {tnr}",
                  font=("", 12, "bold")).pack(anchor="w", pady=(0, 8))

        # Position oben links (Row 0, Col 0)
        tl_frame = ttk.LabelFrame(self.corners_wizard, text="Position oben links (Row 0, Col 0)")
        tl_frame.pack(fill="x", pady=(8, 4))
        tl_inner = ttk.Frame(tl_frame)
        tl_inner.pack(fill="x", padx=4, pady=4)
        self.corners_tl_vars = self._xyz_entries(tl_inner, {"X": 0, "Y": 0, "Z": DEFAULT_CAPTURE_Z})

        # Position unten rechts (Row max, Col max)
        br_frame = ttk.LabelFrame(self.corners_wizard, text="Position unten rechts (Row max, Col max)")
        br_frame.pack(fill="x", pady=4)
        br_inner = ttk.Frame(br_frame)
        br_inner.pack(fill="x", padx=4, pady=4)
        self.corners_br_vars = self._xyz_entries(br_inner, {"X": 0, "Y": 0, "Z": DEFAULT_CAPTURE_Z})

        bf = ttk.Frame(self.corners_wizard)
        bf.pack(fill="x", pady=10)
        ttk.Button(bf, text="⬅ Zurück", command=self._corners_display_step1).pack(side="left")
        ttk.Button(bf, text="Weiter ➜", command=self._corners_goto_step3).pack(side="right")

    def _corners_goto_step3(self):
        grid_str = self.corners_grid_var.get().strip()
        try:
            rows, cols = parse_grid(grid_str)
            assert rows > 0 and cols > 0
        except Exception:
            messagebox.showwarning("Eingabe", "Ungültiges Grid-Format. Bitte z.B. '4x4' eingeben.")
            return

        try:
            tl_x = float(self.corners_tl_vars[0].get())
            tl_y = float(self.corners_tl_vars[1].get())
            tl_z = float(self.corners_tl_vars[2].get())
            br_x = float(self.corners_br_vars[0].get())
            br_y = float(self.corners_br_vars[1].get())
            br_z = float(self.corners_br_vars[2].get())
        except ValueError:
            messagebox.showwarning("Eingabe", "Bitte gültige Zahlenwerte für die Eckpunkte eingeben.")
            return

        self._clear_corners_wizard()
        self.corners_step = 3
        tnr = self.corners_typnr_var.get().strip()

        ttk.Label(self.corners_wizard, text=f"Schritt 4: Positionen für Typ {tnr} ({grid_str})",
                  font=("", 12, "bold")).pack(anchor="w", pady=(0, 8))

        # Scrollbar
        container = ttk.Frame(self.corners_wizard)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container)
        sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        # Diavite
        diav_frame = ttk.LabelFrame(scroll_frame, text="Diavite Messungen")
        diav_frame.pack(fill="x", padx=4, pady=4)
        self.corners_diavite = []
        for label in ("Diavite A", "Diavite B"):
            rf = ttk.Frame(diav_frame)
            rf.pack(fill="x", padx=4, pady=2)
            ttk.Label(rf, text=label, width=14).pack(side="left")
            self.corners_diavite.append(self._xyz_entries(rf, {"X": 0, "Y": 0, "Z": 0}))

        # PreStart
        ps_frame = ttk.LabelFrame(scroll_frame, text="PreStart")
        ps_frame.pack(fill="x", padx=4, pady=4)
        rf = ttk.Frame(ps_frame)
        rf.pack(fill="x", padx=4, pady=2)
        self.corners_prestart = self._xyz_entries(rf, {"X": 0, "Y": 0, "Z": 0})

        # CaptureImage Matrix – berechnet aus den Eckpunkten
        cap_frame = ttk.LabelFrame(scroll_frame, text=f"CaptureImage  ({grid_str})")
        cap_frame.pack(fill="x", padx=4, pady=4)

        header = ttk.Frame(cap_frame)
        header.pack(fill="x", padx=4, pady=2)
        ttk.Label(header, text="", width=10).pack(side="left")
        for c in range(cols):
            ttk.Label(header, text=f"Col {c}", width=28, anchor="center").pack(side="left", padx=2)

        self.corners_capture_entries = {}
        for r in range(rows):
            row_frame = ttk.Frame(cap_frame)
            row_frame.pack(fill="x", padx=4, pady=1)
            ttk.Label(row_frame, text=f"Row {r}", width=10).pack(side="left")
            for c in range(cols):
                cell = ttk.Frame(row_frame, relief="groove", borderwidth=1)
                cell.pack(side="left", padx=2, pady=1)
                px, py, pz = self._calc_corner_position(
                    rows, cols, r, c, tl_x, tl_y, tl_z, br_x, br_y, br_z)
                init_pos = {"X": round(px, 4), "Y": round(py, 4), "Z": round(pz, 4)}
                self.corners_capture_entries[(r, c)] = self._xyz_entries_compact(cell, init_pos)

        # Process
        pr_frame = ttk.LabelFrame(scroll_frame, text="Process")
        pr_frame.pack(fill="x", padx=4, pady=4)
        rf2 = ttk.Frame(pr_frame)
        rf2.pack(fill="x", padx=4, pady=2)
        self.corners_process = self._xyz_entries(rf2, {"X": 0, "Y": 0, "Z": 0})

        # Buttons
        bf = ttk.Frame(self.corners_wizard)
        bf.pack(fill="x", pady=6)
        ttk.Button(bf, text="⬅ Zurück", command=self._corners_display_step2).pack(side="left")
        ttk.Button(bf, text="💾  Neuen Typ speichern",
                   command=lambda: self._save_corners(rows, cols, grid_str)).pack(side="right")

    @staticmethod
    def _calc_corner_position(rows, cols, row, col, tl_x, tl_y, tl_z, br_x, br_y, br_z):
        """Berechnet die Position einer Grid-Zelle aus den beiden Eckpunkten (oben links / unten rechts).

        Args:
            rows, cols: Grid-Dimensionen.
            row, col: Aktuelle Zelle (0-basiert, logische Position).
            tl_x, tl_y, tl_z: Position oben links (Row 0, Col 0).
            br_x, br_y, br_z: Position unten rechts (Row max, Col max).

        Returns:
            (posX, posY, posZ) als Tuple von float.
        """
        pos_x = tl_x + (br_x - tl_x) / (rows - 1) * row if rows > 1 else tl_x
        pos_y = tl_y + (br_y - tl_y) / (cols - 1) * col if cols > 1 else tl_y
        pos_z = tl_z + (br_z - tl_z) / (rows - 1) * row if rows > 1 else tl_z

        return pos_x, pos_y, pos_z

    def _save_corners(self, rows, cols, grid_str):
        tnr = self.corners_typnr_var.get().strip()
        errors = []
        positions = []

        # Diavite
        actions = ["DiaviteMeasurementA", "DiaviteMeasurementB"]
        for i, (xv, yv, zv) in enumerate(self.corners_diavite):
            try:
                pos = {"X": float(xv.get()), "Y": float(yv.get()), "Z": float(zv.get())}
            except ValueError:
                errors.append(f"{actions[i]}: ungültige Zahlenwerte")
                continue
            errors.extend(validate_position(pos, "Diavite", self.limits))
            positions.append({"PathPosAction": actions[i], "Position": pos, "CellRow": -1, "CellColumn": -1})

        # PreStart
        xv, yv, zv = self.corners_prestart
        try:
            pos = {"X": float(xv.get()), "Y": float(yv.get()), "Z": float(zv.get())}
            positions.append({"PathPosAction": "PreStart", "Position": pos, "CellRow": -1, "CellColumn": -1})
        except ValueError:
            errors.append("PreStart: ungültige Zahlenwerte")

        # CaptureImage – Serpentinenmuster
        for r in range(rows):
            if r % 2 == 0:
                col_range = range(cols)
            else:
                col_range = range(cols - 1, -1, -1)
            for c in col_range:
                xv, yv, zv = self.corners_capture_entries[(r, c)]
                try:
                    pos = {"X": float(xv.get()), "Y": float(yv.get()), "Z": float(zv.get())}
                except ValueError:
                    errors.append(f"CaptureImage Row {r} Col {c}: ungültige Zahlenwerte")
                    continue
                errors.extend(validate_position(pos, "CaptureImage", self.limits))
                positions.append({"PathPosAction": "CaptureImage", "Position": pos, "CellRow": r, "CellColumn": c})

        # Process
        xv, yv, zv = self.corners_process
        try:
            pos = {"X": float(xv.get()), "Y": float(yv.get()), "Z": float(zv.get())}
            positions.append({"PathPosAction": "Process", "Position": pos, "CellRow": -1, "CellColumn": -1})
        except ValueError:
            errors.append("Process: ungültige Zahlenwerte")

        if errors:
            messagebox.showerror("Plausibilitätsprüfung fehlgeschlagen", "\n".join(errors))
            return

        self.data[tnr] = {
            "Description": "Python Script Generated",
            "ImageGrid": grid_str,
            "Positions": positions,
        }
        save_json(self.data)

        # Dropdown in Tab 1 aktualisieren
        self.edit_combo["values"] = sorted(self.data.keys())

        messagebox.showinfo("Gespeichert", f"Neuer Typ {tnr} erfolgreich gespeichert.")
        self._corners_show_step0()

    # ───────────────────────────────────────────────────────────────────────
    #  TAB 4 – Einrichter (Limits)
    # ───────────────────────────────────────────────────────────────────────
    def _build_limits_tab(self):
        ttk.Label(self.tab_limits, text="Min / Max Limits für Plausibilitätsprüfung",
                  font=("", 12, "bold")).pack(anchor="w", padx=8, pady=(8, 12))

        self.limit_entries = {}
        for category in ("Diavite", "CaptureImage"):
            lf = ttk.LabelFrame(self.tab_limits, text=category)
            lf.pack(fill="x", padx=8, pady=6)
            self.limit_entries[category] = {}
            for axis in ("X", "Y", "Z"):
                rf = ttk.Frame(lf)
                rf.pack(fill="x", padx=4, pady=2)
                ttk.Label(rf, text=f"{axis}:", width=4).pack(side="left")
                ttk.Label(rf, text="Min:").pack(side="left")
                min_var = tk.StringVar(value=str(self.limits.get(category, {}).get(axis, {}).get("min", 0)))
                ttk.Entry(rf, textvariable=min_var, width=10).pack(side="left", padx=(2, 12))
                ttk.Label(rf, text="Max:").pack(side="left")
                max_var = tk.StringVar(value=str(self.limits.get(category, {}).get(axis, {}).get("max", 2000)))
                ttk.Entry(rf, textvariable=max_var, width=10).pack(side="left", padx=2)
                self.limit_entries[category][axis] = (min_var, max_var)

        # Einzelbild-Größe für Grid-Vorschau
        tile_lf = ttk.LabelFrame(self.tab_limits, text="Einzelbild-Größe (für Grid-Vorschau)")
        tile_lf.pack(fill="x", padx=8, pady=6)
        rf_tile = ttk.Frame(tile_lf)
        rf_tile.pack(fill="x", padx=4, pady=4)
        ttk.Label(rf_tile, text="Einzelbild-Breite (px):").pack(side="left")
        self.tile_width_var = tk.StringVar(value=str(self.limits.get("TileWidth", DEFAULT_TILE_WIDTH)))
        ttk.Entry(rf_tile, textvariable=self.tile_width_var, width=10).pack(side="left", padx=6)
        ttk.Label(rf_tile, text="Einzelbild-Höhe (px):").pack(side="left", padx=(16, 0))
        self.tile_height_var = tk.StringVar(value=str(self.limits.get("TileHeight", DEFAULT_TILE_HEIGHT)))
        ttk.Entry(rf_tile, textvariable=self.tile_height_var, width=10).pack(side="left", padx=6)

        ttk.Button(self.tab_limits, text="💾  Limits speichern", command=self._save_limits).pack(anchor="e", padx=8,
                                                                                                  pady=10)

    def _save_limits(self):
        new_limits = {}
        for category, axes in self.limit_entries.items():
            new_limits[category] = {}
            for axis, (min_var, max_var) in axes.items():
                try:
                    lo = float(min_var.get())
                    hi = float(max_var.get())
                    if lo > hi:
                        messagebox.showerror("Fehler", f"{category} {axis}: Min ({lo}) > Max ({hi})")
                        return
                    new_limits[category][axis] = {"min": lo, "max": hi}
                except ValueError:
                    messagebox.showerror("Fehler", f"{category} {axis}: ungültige Zahlenwerte")
                    return
        try:
            tw = int(float(self.tile_width_var.get()))
            th = int(float(self.tile_height_var.get()))
            if tw <= 0 or th <= 0:
                raise ValueError
            new_limits["TileWidth"] = tw
            new_limits["TileHeight"] = th
        except ValueError:
            messagebox.showerror("Fehler", "Einzelbild-Breite/-Höhe: ungültige Zahlenwerte (müssen > 0 sein)")
            return
        self.limits = new_limits
        save_limits(new_limits)
        messagebox.showinfo("Gespeichert", "Limits erfolgreich gespeichert.")

    # ───────────────────────────────────────────────────────────────────────
    #  Widgets-Helfer
    # ───────────────────────────────────────────────────────────────────────
    @staticmethod
    def _get_initial_position(rows, cols, row, col, width, height, origin_x=None, origin_y=None):
        """Berechnet die initiale X/Y/Z Position für eine Grid-Zelle.

        Koordinatensystem:
          - X entspricht der Zeilen-Richtung (row), NICHT der Spalten-Richtung.
          - Y entspricht der Spalten-Richtung (col).
          - Das Serpentinenmuster wechselt die Spalten-Durchlaufrichtung pro Zeile:
            gerade Zeilen (0, 2, …) → col aufsteigend,
            ungerade Zeilen (1, 3, …) → col absteigend.

        Args:
            rows: Anzahl Zeilen im Grid.
            cols: Anzahl Spalten im Grid.
            row:  Aktuelle Zeile (0-basiert).
            col:  Aktuelle Spalte (0-basiert, logische Position in der Matrix).
            width: Breite = X-Span (Zeilen-Richtung).
            height: Höhe = Y-Span (Spalten-Richtung).
            origin_x: Ursprung X (Standard: GRID_ORIGIN_X).
            origin_y: Ursprung Y (Standard: GRID_ORIGIN_Y).

        Returns:
            (posX, posY, posZ) als Tuple von float.
        """
        if origin_x is None:
            origin_x = GRID_ORIGIN_X
        if origin_y is None:
            origin_y = GRID_ORIGIN_Y

        effective_col = col if row % 2 == 0 else cols - col - 1

        min_bounds = (origin_x, origin_y)
        max_bounds = (origin_x + width, origin_y + height)

        span_x = max_bounds[0] - min_bounds[0]
        span_y = max_bounds[1] - min_bounds[1]

        pos_x = min_bounds[0] + (span_x / (rows - 1)) * row if rows > 1 else min_bounds[0]
        pos_y = min_bounds[1] + (span_y / (cols - 1)) * effective_col if cols > 1 else min_bounds[1]
        pos_z = DEFAULT_CAPTURE_Z

        return pos_x, pos_y, pos_z

    @staticmethod
    def _xyz_entries(parent, pos_dict):
        """Erzeugt X/Y/Z Eingabefelder in einer Zeile und gibt (x_var, y_var, z_var) zurück."""
        vars_ = []
        for axis in ("X", "Y", "Z"):
            ttk.Label(parent, text=f"{axis}:").pack(side="left", padx=(8, 0))
            v = tk.StringVar(value=str(pos_dict.get(axis, 0)))
            ttk.Entry(parent, textvariable=v, width=12).pack(side="left", padx=2)
            vars_.append(v)
        return tuple(vars_)

    @staticmethod
    def _xyz_entries_compact(parent, pos_dict):
        """Kompakte X/Y/Z Felder für Matrixzellen."""
        vars_ = []
        for axis in ("X", "Y", "Z"):
            f = ttk.Frame(parent)
            f.pack(side="left", padx=1)
            ttk.Label(f, text=axis, font=("", 7)).pack()
            v = tk.StringVar(value=str(pos_dict.get(axis, 0)))
            ttk.Entry(f, textvariable=v, width=8, font=("", 8)).pack()
            vars_.append(v)
        return tuple(vars_)

    @staticmethod
    def _find_capture(positions, row, col):
        for p in positions:
            if p["PathPosAction"] == "CaptureImage" and p["CellRow"] == row and p["CellColumn"] == col:
                return p
        return None


# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.mainloop()