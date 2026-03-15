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
    from PIL import Image, ImageTk
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

def save_json(data):
    if os.path.exists(JSON_FILE):
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        backup_path = os.path.join(
            os.path.dirname(JSON_FILE),
            f"Typdaten_backup_{timestamp}.json",
        )
        shutil.copy2(JSON_FILE, backup_path)
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent="\t", ensure_ascii=False)

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

        # --- Tab 3: Einrichter (Limits) ---
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

        # Speichern-Button
        btn_frame = ttk.Frame(self.tab_edit)
        btn_frame.pack(fill="x", padx=8, pady=6)
        ttk.Button(btn_frame, text="💾  Änderungen speichern", command=self._save_edit).pack(side="right")

        # Datencontainer
        self.edit_diavite_entries = []  # [(action_name, (x_var, y_var, z_var)), ...]
        self.edit_capture_entries = {}  # {(row, col): (x_var, y_var, z_var)}
        self.edit_prestart_entry = None
        self.edit_process_entry = None

    def _on_type_selected(self, _event=None):
        typ_key = self.edit_type_var.get()
        if not typ_key:
            return
        typ = self.data[typ_key]

        # Alten Inhalt löschen
        for w in self.edit_scroll_frame.winfo_children():
            w.destroy()
        self.edit_diavite_entries.clear()
        self.edit_capture_entries.clear()

        # --- Description & Grid ---
        info = ttk.LabelFrame(self.edit_scroll_frame, text="Info")
        info.pack(fill="x", padx=8, pady=4)
        ttk.Label(info, text=f"Description: {typ['Description']}    |    ImageGrid: {typ['ImageGrid']}").pack(
            anchor="w", padx=4, pady=2)

        # --- Bild anzeigen (falls vorhanden) ---
        self._show_type_image(typ_key, typ["ImageGrid"])

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

        # --- CaptureImage Matrix ---
        rows, cols = parse_grid(typ["ImageGrid"])
        cap_frame = ttk.LabelFrame(self.edit_scroll_frame, text=f"CaptureImage  ({typ['ImageGrid']})")
        cap_frame.pack(fill="x", padx=8, pady=4)

        # Spaltenüberschriften
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
                # Finde passende Position
                p = self._find_capture(typ["Positions"], r, c)
                pos_data = p["Position"] if p else {"X": 0, "Y": 0, "Z": 0}
                vars_ = self._xyz_entries_compact(cell_frame, pos_data)
                self.edit_capture_entries[(r, c)] = vars_

        # --- Process ---
        for pos in typ["Positions"]:
            if pos["PathPosAction"] == "Process":
                pr_frame = ttk.LabelFrame(self.edit_scroll_frame, text="Process")
                pr_frame.pack(fill="x", padx=8, pady=4)
                row_f = ttk.Frame(pr_frame)
                row_f.pack(fill="x", padx=4, pady=2)
                self.edit_process_entry = self._xyz_entries(row_f, pos["Position"])

    def _show_type_image(self, typ_key, grid_str):
        """Zeigt das Typbild und die Grid-Teilbilder an, falls eine Bilddatei existiert."""
        if not PIL_AVAILABLE:
            return
        base_dir = os.path.dirname(os.path.abspath(__file__))
        img_path = None
        for ext in (".png", ".jpg", ".jpeg", ".bmp"):
            candidate = os.path.join(base_dir, typ_key + ext)
            if os.path.isfile(candidate):
                img_path = candidate
                break
        if img_path is None:
            return

        img = Image.open(img_path)

        img_frame = ttk.LabelFrame(self.edit_scroll_frame, text="Typbild")
        img_frame.pack(fill="x", padx=8, pady=4)

        # Originalbild und Grid-Teilbilder nebeneinander
        content = ttk.Frame(img_frame)
        content.pack(fill="x", padx=4, pady=4)

        # --- Originalbild (links) ---
        orig_frame = ttk.Frame(content)
        orig_frame.pack(side="left", anchor="n", padx=(0, 20))
        ttk.Label(orig_frame, text="Original", font=("", 9, "bold")).pack()
        photo = ImageTk.PhotoImage(img)
        lbl = ttk.Label(orig_frame, image=photo)
        lbl.image = photo  # Referenz halten
        lbl.pack()

        # --- Grid-Teilbilder (rechts) ---
        rows, cols = parse_grid(grid_str)
        w, h = img.size
        cell_w = w // cols
        cell_h = h // rows
        gap = 10

        grid_frame = ttk.Frame(content)
        grid_frame.pack(side="left", anchor="n")
        ttk.Label(grid_frame, text=f"Grid ({grid_str})", font=("", 9, "bold")).pack(anchor="w")

        # Referenzliste für PhotoImages (gegen Garbage Collection)
        self._grid_photos = []
        for r in range(rows):
            row_f = ttk.Frame(grid_frame)
            row_f.pack(anchor="w", pady=(gap if r > 0 else 0, 0))
            for c in range(cols):
                x0 = c * cell_w
                y0 = r * cell_h
                x1 = x0 + cell_w
                y1 = y0 + cell_h
                cell_img = img.crop((x0, y0, x1, y1))
                cell_photo = ImageTk.PhotoImage(cell_img)
                cell_lbl = ttk.Label(row_f, image=cell_photo)
                cell_lbl.image = cell_photo
                cell_lbl.pack(side="left", padx=(gap if c > 0 else 0, 0))
                self._grid_photos.append(cell_photo)

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

        # CaptureImage
        for (r, c), (xv, yv, zv) in self.edit_capture_entries.items():
            try:
                pos = {"X": float(xv.get()), "Y": float(yv.get()), "Z": float(zv.get())}
            except ValueError:
                errors.append(f"CaptureImage Row {r} Col {c}: ungültige Zahlenwerte")
                continue
            errors.extend(validate_position(pos, "CaptureImage", self.limits))
            p = self._find_capture(typ["Positions"], r, c)
            if p:
                p["Position"] = pos

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
        for attr in ("new_grid_var", "new_width_var", "new_height_var", "new_preview_frame"):
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
        ttk.Label(self.new_wizard, text=f"Schritt 2: ImageGrid für Typ {tnr}", font=("", 12, "bold")).pack(
            anchor="w", pady=(0, 8))
        f = ttk.Frame(self.new_wizard)
        f.pack(fill="x")
        ttk.Label(f, text="ImageGrid (z.B. 4x4):").pack(side="left")
        if not hasattr(self, "new_grid_var"):
            self.new_grid_var = tk.StringVar()
            self.new_grid_var.trace_add("write", lambda *_a: self._new_update_preview())
        ttk.Entry(f, textvariable=self.new_grid_var, width=10).pack(side="left", padx=6)

        f2 = ttk.Frame(self.new_wizard)
        f2.pack(fill="x", pady=(4, 0))
        ttk.Label(f2, text="Breite links:").pack(side="left")
        self.new_width_left_var = tk.StringVar(value=str(DEFAULT_GRID_WIDTH / 2))
        ttk.Entry(f2, textvariable=self.new_width_left_var, width=14).pack(side="left", padx=6)
        ttk.Label(f2, text="Breite rechts:").pack(side="left", padx=(8, 0))
        self.new_width_right_var = tk.StringVar(value=str(DEFAULT_GRID_WIDTH / 2))
        ttk.Entry(f2, textvariable=self.new_width_right_var, width=14).pack(side="left", padx=6)

        f3 = ttk.Frame(self.new_wizard)
        f3.pack(fill="x", pady=(4, 0))
        ttk.Label(f3, text="Höhe oben:").pack(side="left")
        self.new_height_above_var = tk.StringVar(value=str(DEFAULT_GRID_HEIGHT / 2))
        ttk.Entry(f3, textvariable=self.new_height_above_var, width=14).pack(side="left", padx=6)
        ttk.Label(f3, text="Höhe unten:").pack(side="left", padx=(8, 0))
        self.new_height_below_var = tk.StringVar(value=str(DEFAULT_GRID_HEIGHT / 2))
        ttk.Entry(f3, textvariable=self.new_height_below_var, width=14).pack(side="left", padx=6)

        bf = ttk.Frame(self.new_wizard)
        bf.pack(fill="x", pady=10)
        ttk.Button(bf, text="⬅ Zurück", command=self._new_show_step0).pack(side="left")
        ttk.Button(bf, text="Weiter ➜", command=self._new_goto_step2).pack(side="right")

        # Show preview if a value is already present (e.g. when navigating back)
        self._new_update_preview()

    def _new_update_preview(self):
        """Sucht ein passendes Bild und zeigt Original + Grid-Overlay an."""
        if not hasattr(self, "new_preview_frame") or not self.new_preview_frame.winfo_exists():
            return
        for w in self.new_preview_frame.winfo_children():
            w.destroy()

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

        # Passendes Bild suchen (Dateiname ohne .png muss im Typnamen enthalten sein)
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

        # Skalierung: max. 300 px Höhe
        MAX_H = 300
        img_w, img_h = original_img.size
        scale = min(1.0, MAX_H / img_h)
        disp_w = max(1, int(img_w * scale))
        disp_h = max(1, int(img_h * scale))
        disp_original = original_img.resize((disp_w, disp_h), Image.LANCZOS)

        # Einzelbild-Größe aus Limits lesen
        tile_w = int(self.limits.get("TileWidth", DEFAULT_TILE_WIDTH))
        tile_h = int(self.limits.get("TileHeight", DEFAULT_TILE_HEIGHT))
        t_w = max(1, int(tile_w * scale))
        t_h = max(1, int(tile_h * scale))

        # Grid-Overlay zeichnen
        # Startposition: start[i] = i * (total - tile) / (count - 1)  für count > 1, sonst 0.
        # Negative Koordinaten (wenn Kachel größer als Bild) sind gewollt (Überlapp) –
        # PIL schneidet Zeichnungen außerhalb des Bildes automatisch ab.
        grid_img = disp_original.copy().convert("RGB")
        draw = ImageDraw.Draw(grid_img)
        for r in range(rows):
            y0 = int(r * (disp_h - t_h) / (rows - 1)) if rows > 1 else 0
            for c in range(cols):
                x0 = int(c * (disp_w - t_w) / (cols - 1)) if cols > 1 else 0
                draw.rectangle([x0, y0, x0 + t_w, y0 + t_h], outline="red", width=2)

        left_lf = ttk.LabelFrame(self.new_preview_frame, text="Originalbild")
        left_lf.pack(side="left", padx=4, anchor="n")
        photo_orig = ImageTk.PhotoImage(disp_original)
        lbl_orig = ttk.Label(left_lf, image=photo_orig)
        lbl_orig.image = photo_orig
        lbl_orig.pack()

        right_lf = ttk.LabelFrame(self.new_preview_frame, text=f"Grid-Aufteilung ({grid_str})")
        right_lf.pack(side="left", padx=4, anchor="n")
        photo_grid = ImageTk.PhotoImage(grid_img)
        lbl_grid = ttk.Label(right_lf, image=photo_grid)
        lbl_grid.image = photo_grid
        lbl_grid.pack()

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
        """Schritt 3: Breite und Höhe eingeben."""
        self._clear_wizard()
        self.new_step = 2
        tnr = self.new_typnr_var.get().strip()
        ttk.Label(self.new_wizard, text=f"Schritt 3: Maße für Typ {tnr}", font=("", 12, "bold")).pack(
            anchor="w", pady=(0, 8))

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
        """Schritt 3 → 4: Breite/Höhe validieren, dann Positionen-Schritt anzeigen."""
        grid_str = self.new_grid_var.get().strip()
        rows, cols = parse_grid(grid_str)

        try:
            width_left = float(self.new_width_left_var.get())
            width_right = float(self.new_width_right_var.get())
            height_above = float(self.new_height_above_var.get())
            height_below = float(self.new_height_below_var.get())
            assert width_left >= 0 and width_right >= 0
            assert height_above >= 0 and height_below >= 0
            width = width_left + width_right
            height = height_above + height_below
            assert width > 0 and height > 0
        except Exception:
            messagebox.showwarning("Eingabe", "Breite und Höhe müssen positive Zahlen sein.")
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
                px, py, pz = self._get_initial_position(rows, cols, r, c, width, height)
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
    #  TAB 3 – Einrichter (Limits)
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
    def _get_initial_position(rows, cols, row, col, width, height):
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

        Returns:
            (posX, posY, posZ) als Tuple von float.
        """
        effective_col = col if row % 2 == 0 else cols - col - 1

        min_bounds = (GRID_ORIGIN_X, GRID_ORIGIN_Y)
        max_bounds = (GRID_ORIGIN_X + width, GRID_ORIGIN_Y + height)

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