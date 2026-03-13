#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI zum Bearbeiten und Anlegen von Typen in Typdaten.json
"""

import json
import os
import tkinter as tk
from tkinter import ttk, messagebox

JSON_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Typdaten.json")

# ─── Standardlimits (werden über Tab "Einrichter" angepasst) ────────────────
DEFAULT_LIMITS = {
    "Diavite": {"X": {"min": 0, "max": 2000}, "Y": {"min": 0, "max": 500}, "Z": {"min": 0, "max": 200}},
    "CaptureImage": {"X": {"min": 0, "max": 2000}, "Y": {"min": 0, "max": 500}, "Z": {"min": 0, "max": 200}},
}
LIMITS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Limits.json")


# ═══════════════════════════════════════════════════════════════════════════════
#  Hilfsfunktionen
# ═══════════════════════════════════════════════════════════════════════════════
def load_json():
    with open(JSON_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_json(data):
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
        self.new_step = 0  # 0=Typnummer, 1=Grid, 2=Positionen

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
        self._clear_wizard()
        self.new_step = 1
        ttk.Label(self.new_wizard, text=f"Schritt 2: ImageGrid für Typ {tnr}", font=("", 12, "bold")).pack(
            anchor="w", pady=(0, 8))
        f = ttk.Frame(self.new_wizard)
        f.pack(fill="x")
        ttk.Label(f, text="ImageGrid (z.B. 4x4):").pack(side="left")
        self.new_grid_var = tk.StringVar()
        ttk.Entry(f, textvariable=self.new_grid_var, width=10).pack(side="left", padx=6)

        bf = ttk.Frame(self.new_wizard)
        bf.pack(fill="x", pady=10)
        ttk.Button(bf, text="⬅ Zurück", command=self._new_show_step0).pack(side="left")
        ttk.Button(bf, text="Weiter ➜", command=self._new_goto_step2).pack(side="right")

    def _new_goto_step2(self):
        grid_str = self.new_grid_var.get().strip()
        try:
            rows, cols = parse_grid(grid_str)
            assert rows > 0 and cols > 0
        except Exception:
            messagebox.showwarning("Eingabe", "Ungültiges Grid-Format. Bitte z.B. '4x4' eingeben.")
            return

        self._clear_wizard()
        self.new_step = 2
        tnr = self.new_typnr_var.get().strip()

        ttk.Label(self.new_wizard, text=f"Schritt 3: Positionen für Typ {tnr} ({grid_str})",
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
                self.new_capture_entries[(r, c)] = self._xyz_entries_compact(cell, {"X": 0, "Y": 0, "Z": 0})

        # Process
        pr_frame = ttk.LabelFrame(scroll_frame, text="Process")
        pr_frame.pack(fill="x", padx=4, pady=4)
        rf2 = ttk.Frame(pr_frame)
        rf2.pack(fill="x", padx=4, pady=2)
        self.new_process = self._xyz_entries(rf2, {"X": 0, "Y": 0, "Z": 0})

        # Buttons
        bf = ttk.Frame(self.new_wizard)
        bf.pack(fill="x", pady=6)
        ttk.Button(bf, text="⬅ Zurück", command=self._new_goto_step1).pack(side="left")
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
        self.limits = new_limits
        save_limits(new_limits)
        messagebox.showinfo("Gespeichert", "Limits erfolgreich gespeichert.")

    # ───────────────────────────────────────────────────────────────────────
    #  Widgets-Helfer
    # ───────────────────────────────────────────────────────────────────────
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