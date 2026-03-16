import customtkinter as ctk
import json
import os
import copy
import zipfile
import re
from itertools import permutations as iter_permutations
from tkinter import colorchooser, messagebox, filedialog
import math
from datetime import datetime

# ── KONSTANTEN ────────────────────────────────────────────────────────────────
MAX_SEQ_LEN     = 10
DEFAULT_TD      = 5.0
DE_GOOD         = 3.0
DE_OK           = 6.0
GAMUT_WARN_DE   = 25.0
MAX_VIRTUAL     = 20   # V5 … V24

DEFAULT_LIBRARY = {
    "Bambu Lab Basic": [
        {"name": "Jade White",  "hex": "#FFFFFF", "td": 5.0},
        {"name": "Cyan",        "hex": "#0086D6", "td": 4.0},
        {"name": "Magenta",     "hex": "#EC008C", "td": 8.0},
        {"name": "Yellow",      "hex": "#FCE300", "td": 6.0},
        {"name": "Black",       "hex": "#101010", "td": 0.4},
        {"name": "Red",         "hex": "#D32F2F", "td": 3.0},
        {"name": "Green",       "hex": "#388E3C", "td": 4.5},
    ],
    "Bambu Lab Matte": [
        {"name": "Ivory White",  "hex": "#F2F2F2", "td": 3.5},
        {"name": "Charcoal",     "hex": "#2D2926", "td": 0.2},
        {"name": "Lilac Purple", "hex": "#A181C1", "td": 2.5},
    ],
    "Prusament": [
        {"name": "Vanilla White", "hex": "#D9D4C4", "td": 7.1},
        {"name": "Jet Black",     "hex": "#24292A", "td": 0.3},
        {"name": "Prusa Orange",  "hex": "#FE6E31", "td": 6.6},
        {"name": "Galaxy Silver", "hex": "#868F98", "td": 1.5},
    ],
    "Anycubic": [
        {"name": "White",   "hex": "#FFFFFF", "td": 8.5},
        {"name": "Cyan",    "hex": "#00FFFF", "td": 4.2},
        {"name": "Magenta", "hex": "#FF00FF", "td": 3.8},
        {"name": "Yellow",  "hex": "#FFFF00", "td": 5.5},
    ],
    "Eigene Favoriten": [],
}

# ── FARB-MATHEMATIK ───────────────────────────────────────────────────────────

def hex_to_rgb(hex_str):
    s = hex_str.strip().lstrip("#")
    if len(s) == 8: s = s[:6]
    if len(s) != 6: return (128, 128, 128)
    try:   return tuple(int(s[i:i+2], 16) for i in (0, 2, 4))
    except ValueError: return (128, 128, 128)

def rgb_to_lab(rgb):
    r, g, b = [x / 255.0 for x in rgb]
    r = (r / 12.92) if r <= 0.04045 else ((r + 0.055) / 1.055) ** 2.4
    g = (g / 12.92) if g <= 0.04045 else ((g + 0.055) / 1.055) ** 2.4
    b = (b / 12.92) if b <= 0.04045 else ((b + 0.055) / 1.055) ** 2.4
    x = (r * 0.4124 + g * 0.3576 + b * 0.1805) / 0.95047
    y =  r * 0.2126 + g * 0.7152 + b * 0.0722
    z = (r * 0.0193 + g * 0.1192 + b * 0.9505) / 1.08883
    x, y, z = [v**(1/3) if v > 0.008856 else 7.787*v + 16/116 for v in [x, y, z]]
    return (116*y - 16, 500*(x - y), 200*(y - z))

def lab_to_hex(lab):
    L, a, b = lab
    fy = (L + 16) / 116
    fx = a / 500 + fy
    fz = fy - b / 200
    x = 0.95047 * (fx**3 if fx**3 > 0.008856 else (fx - 16/116) / 7.787)
    y = 1.00000 * (fy**3 if fy**3 > 0.008856 else (fy - 16/116) / 7.787)
    z = 1.08883 * (fz**3 if fz**3 > 0.008856 else (fz - 16/116) / 7.787)
    r =  3.2406*x - 1.5372*y - 0.4986*z
    g = -0.9689*x + 1.8758*y + 0.0415*z
    b_ = 0.0557*x - 0.2040*y + 1.0570*z
    def gamma(c):
        c = max(0.0, min(1.0, c))
        return 1.055 * c**(1/2.4) - 0.055 if c > 0.0031308 else 12.92 * c
    return "#{:02X}{:02X}{:02X}".format(int(gamma(r)*255), int(gamma(g)*255), int(gamma(b_)*255))

def delta_e(lab1, lab2):
    return math.sqrt(sum((a - b)**2 for a, b in zip(lab1, lab2)))

def safe_td(value):
    try:
        v = float(str(value).strip())
        return v if v > 0 else DEFAULT_TD
    except (ValueError, TypeError):
        return DEFAULT_TD

def get_layer_weights(n):
    """Lineare progressive Gewichtung für beliebige Sequenzlänge: 1.0 (unten) → 1.5 (oben)."""
    if n == 1:
        return [1.0]
    return [1.0 + 0.5 * i / (n - 1) for i in range(n)]

def de_color(de):
    return "#4ade80" if de < DE_GOOD else "#fbbf24" if de < DE_OK else "#f87171"

def seq_to_runs(sequence):
    """'11122333' → [(1,3), (2,2), (3,3)] — konsekutive Runs pro Filament."""
    if not sequence: return []
    runs, cur, cnt = [], int(sequence[0]), 1
    for ch in sequence[1:]:
        fid = int(ch)
        if fid == cur: cnt += 1
        else: runs.append((cur, cnt)); cur, cnt = fid, 1
    runs.append((cur, cnt))
    return runs

def calc_cadence(sequence, layer_height):
    """Cadence Heights aus Sequenz + Schichthöhe.
    Liefert {filament_id: cadence_mm} basierend auf dem längsten Run je Filament.
    Physikalisch korrekt für 2-Filament-Sequenzen; Näherung bei mehr.
    """
    runs = seq_to_runs(sequence)
    max_run = {}
    for fid, n in runs:
        max_run[fid] = max(max_run.get(fid, 0), n)
    return {fid: round(n * layer_height, 4) for fid, n in max_run.items()}

def seq_filament_count(sequence):
    return len(set(sequence))

def de_label_text(de):
    q = "✓ gut" if de < DE_GOOD else "~ ok" if de < DE_OK else "✗ weit"
    return f"ΔE {de:.1f}  {q}"

# ── 3MF FARBEXTRAKTION ────────────────────────────────────────────────────────

HEX_RE = re.compile(r'#([0-9A-Fa-f]{6})\b')

def parse_3mf_colors(filepath):
    found = set()
    try:
        with zipfile.ZipFile(filepath, 'r') as zf:
            for name in zf.namelist():
                if not any(name.endswith(e) for e in ('.model', '.config', '.xml', '.json')):
                    continue
                try:
                    content = zf.read(name).decode('utf-8', errors='replace')
                except Exception:
                    continue
                if 'settings' in name or name.endswith('.json'):
                    try:
                        data = json.loads(content)
                        for key in ('filament_colour', 'filament_color'):
                            val = data.get(key, [])
                            if isinstance(val, list):
                                for c in val:
                                    m = HEX_RE.match(str(c).strip())
                                    if m: found.add('#' + m.group(1).upper())
                    except (json.JSONDecodeError, AttributeError):
                        pass
                for m in HEX_RE.finditer(content):
                    found.add('#' + m.group(1).upper())
    except Exception as e:
        return [], str(e)
    trivial = {"#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF",
               "#AAAAAA", "#333333", "#CCCCCC"}
    meaningful = [c for c in found if c not in trivial]
    return (meaningful if meaningful else list(found)), None


# ═══════════════════════════════════════════════════════════════════════════════
#  HAUPTANWENDUNG
# ═══════════════════════════════════════════════════════════════════════════════

class U1FullSpectrumApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("U1 FullSpectrum Ultimate — Pro Edition")
        self.geometry("1420x1000")
        ctk.set_appearance_mode("dark")
        self.db_file       = "filament_db.json"
        self.preset_file   = "presets.json"
        self.history       = []
        self.presets       = {}
        self.virtual_fils  = []   # list of virtual filament dicts
        self.last_result   = {}   # last calc() result for "hinzufügen"
        self.load_db()
        self.load_presets()
        self.setup_ui()

    # ── DATENBANK ──────────────────────────────────────────────────────────────

    def load_db(self):
        self.library = copy.deepcopy(DEFAULT_LIBRARY)
        if not os.path.exists(self.db_file): return
        try:
            with open(self.db_file, "r", encoding="utf-8") as f:
                saved = json.load(f)
            if not isinstance(saved, dict): raise ValueError("Kein JSON-Objekt.")
            for brand, fils in saved.items():
                if not isinstance(fils, list): continue
                self.library[brand] = [
                    {**x, "td": safe_td(x.get("td", DEFAULT_TD))}
                    for x in fils if isinstance(x, dict) and "name" in x and "hex" in x
                ]
        except (json.JSONDecodeError, ValueError, IOError) as e:
            messagebox.showwarning("DB-Fehler",
                f"filament_db.json Ladefehler:\n{e}\nStandardwerte aktiv.")

    def save_db(self):
        try:
            with open(self.db_file, "w", encoding="utf-8") as f:
                json.dump(self.library, f, indent=2, ensure_ascii=False)
        except IOError as e:
            messagebox.showerror("Speicherfehler", str(e))

    # ── PRESETS ────────────────────────────────────────────────────────────────

    def load_presets(self):
        if not os.path.exists(self.preset_file): return
        try:
            with open(self.preset_file, "r", encoding="utf-8") as f:
                self.presets = json.load(f)
        except Exception: self.presets = {}

    def save_presets(self):
        try:
            with open(self.preset_file, "w", encoding="utf-8") as f:
                json.dump(self.presets, f, indent=2, ensure_ascii=False)
        except IOError as e:
            messagebox.showerror("Fehler", str(e))

    def save_preset(self):
        name = ctk.CTkInputDialog(text="Preset-Name:", title="Preset speichern").get_input()
        if not name: return
        self.presets[name.strip()] = [
            {"brand": s["brand"].get(), "filament": s["color"].get(),
             "hex": s["hex"].get(), "td": safe_td(s["td"].get())}
            for s in self.slots
        ]
        self.save_presets(); self._refresh_preset_dropdown()
        self.preset_var.set(name.strip())
        messagebox.showinfo("Gespeichert", f'Preset "{name}" gespeichert.')

    def load_preset(self):
        name = self.preset_var.get()
        if not name or name not in self.presets: return
        for i, entry in enumerate(self.presets[name][:4]):
            s = self.slots[i]
            brand = entry.get("brand", "")
            if brand in self.library:
                s["brand"].set(brand); self.update_menu(i)
                fn = entry.get("filament", "")
                if fn in [f["name"] for f in self.library.get(brand, [])]:
                    s["color"].set(fn); self.apply_f(i); continue
            h = entry.get("hex", "#808080")
            s["hex"].delete(0, "end"); s["hex"].insert(0, h)
            s["td"].delete(0, "end");  s["td"].insert(0, str(entry.get("td", DEFAULT_TD)))
            s["preview"].configure(fg_color=h)

    def _refresh_preset_dropdown(self):
        names = list(self.presets.keys()) or ["(keine Presets)"]
        self.preset_dropdown.configure(values=names)
        if self.preset_var.get() not in names: self.preset_var.set(names[0])

    # ── UI-AUFBAU ──────────────────────────────────────────────────────────────

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── SIDEBAR: PHYSISCHE SLOTS ─────────────────────────────────────────
        self.sidebar = ctk.CTkScrollableFrame(self, width=430, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew", padx=(10, 5), pady=10)

        ctk.CTkLabel(self.sidebar, text="Physische Druckköpfe  T1–T4",
                     font=("Segoe UI", 18, "bold"), text_color="#38bdf8").pack(pady=(14, 4))
        ctk.CTkLabel(self.sidebar,
                     text="Diese 4 Filamente sind real im Drucker geladen.",
                     font=("Segoe UI", 10), text_color="#475569").pack(pady=(0, 10))

        self.slots = []
        for i in range(4):
            frame = ctk.CTkFrame(self.sidebar, border_width=1, border_color="#334155")
            frame.pack(fill="x", padx=12, pady=5)

            hdr = ctk.CTkFrame(frame, fg_color="transparent")
            hdr.pack(fill="x", padx=5, pady=(6, 0))
            ctk.CTkLabel(hdr, text=f"T{i+1} — WERKZEUG {i+1}",
                         font=("Segoe UI", 12, "bold"), text_color="#94a3b8").pack(side="left", padx=5)
            ctk.CTkButton(hdr, text="+", width=26, height=20, fg_color="#15803d",
                          command=lambda idx=i: self.add_filament(idx)).pack(side="right", padx=2)
            ctk.CTkButton(hdr, text="💾", width=26, height=20,
                          command=lambda idx=i: self.save_current(idx)).pack(side="right", padx=2)

            brand = ctk.CTkOptionMenu(frame, values=list(self.library.keys()),
                                       command=lambda x, idx=i: self.update_menu(idx))
            brand.pack(padx=10, pady=(3, 2), fill="x")
            color = ctk.CTkOptionMenu(frame, values=["Lade..."],
                                       command=lambda x, idx=i: self.apply_f(idx))
            color.pack(padx=10, pady=2, fill="x")

            row = ctk.CTkFrame(frame, fg_color="transparent")
            row.pack(fill="x", pady=(3, 8))
            hx = ctk.CTkEntry(row, width=88, placeholder_text="#RRGGBB")
            hx.pack(side="left", padx=(10, 2))
            preview = ctk.CTkLabel(row, text="", width=26, height=26,
                                   fg_color="#1e293b", corner_radius=5)
            preview.pack(side="left", padx=(0, 2))
            ctk.CTkButton(row, text="🎨", width=28, height=26, fg_color="#334155",
                          command=lambda idx=i: self.pick_slot_color(idx)).pack(side="left", padx=(0, 6))
            ctk.CTkLabel(row, text="TD:", font=("Segoe UI", 10)).pack(side="right", padx=(0, 4))
            td = ctk.CTkEntry(row, width=54, placeholder_text="5.0")
            td.pack(side="right", padx=(0, 4))

            self.slots.append({"brand": brand, "color": color,
                                "hex": hx, "td": td, "preview": preview})
            self.update_menu(i)

        # Presets
        ctk.CTkLabel(self.sidebar, text="SLOT-PRESETS",
                     font=("Segoe UI", 10, "bold"), text_color="#64748b").pack(pady=(12, 2))
        pf = ctk.CTkFrame(self.sidebar, fg_color="#0f172a", corner_radius=8)
        pf.pack(fill="x", padx=12, pady=(0, 5))
        self.preset_var = ctk.StringVar(value="(keine Presets)")
        self.preset_dropdown = ctk.CTkOptionMenu(pf, variable=self.preset_var,
                                                  values=["(keine Presets)"])
        self.preset_dropdown.pack(fill="x", padx=10, pady=(8, 4))
        pb = ctk.CTkFrame(pf, fg_color="transparent")
        pb.pack(fill="x", padx=10, pady=(0, 8))
        pb.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(pb, text="LADEN", fg_color="#1e3a5f",
                      command=self.load_preset).grid(row=0, column=0, sticky="ew", padx=(0, 3))
        ctk.CTkButton(pb, text="SPEICHERN", fg_color="#1e3a5f",
                      command=self.save_preset).grid(row=0, column=1, sticky="ew", padx=(3, 0))
        self._refresh_preset_dropdown()

        # Sidebar-Tools
        ctk.CTkButton(self.sidebar, text="＋  Neue Marke", fg_color="#1e3a5f",
                      command=self.add_brand).pack(fill="x", padx=12, pady=(8, 3))
        ctk.CTkButton(self.sidebar, text="📚  Bibliothek verwalten", fg_color="#374151",
                      command=self.open_library_manager).pack(fill="x", padx=12, pady=(3, 8))

        # Schichthöhe — global für Cadence-Berechnung
        lh_frame = ctk.CTkFrame(self.sidebar, fg_color="#0f172a", corner_radius=8)
        lh_frame.pack(fill="x", padx=12, pady=(0, 15))
        lh_inner = ctk.CTkFrame(lh_frame, fg_color="transparent")
        lh_inner.pack(fill="x", padx=10, pady=8)
        ctk.CTkLabel(lh_inner, text="Schichthöhe (mm):",
                     font=("Segoe UI", 11)).pack(side="left")
        self.layer_height_entry = ctk.CTkEntry(lh_inner, width=60, placeholder_text="0.2")
        self.layer_height_entry.insert(0, "0.2")
        self.layer_height_entry.pack(side="left", padx=(6, 0))
        ctk.CTkLabel(lh_inner, text="= Dithering Step Size",
                     font=("Segoe UI", 9), text_color="#475569").pack(side="left", padx=6)

        # ── HAUPTBEREICH ─────────────────────────────────────────────────────
        self.main = ctk.CTkScrollableFrame(self)
        self.main.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=10)
        self.main.grid_columnconfigure(0, weight=1)

        # ── ABSCHNITT 1: EINZELFARBEN-RECHNER ────────────────────────────────
        sec1 = ctk.CTkFrame(self.main, fg_color="#0f172a", corner_radius=10)
        sec1.grid(row=0, column=0, padx=0, pady=(0, 12), sticky="ew")
        sec1.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(sec1, text="EINZELFARBEN-RECHNER",
                     font=("Segoe UI", 13, "bold"), text_color="#38bdf8").grid(
            row=0, column=0, padx=20, pady=(14, 6), sticky="w")

        # Zielfarbe
        top = ctk.CTkFrame(sec1, fg_color="transparent")
        top.grid(row=1, column=0, padx=20, pady=(0, 8), sticky="ew")
        top.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(top, text="ZIELFARBE  🎨", command=self.pick,
                      height=46, font=("Segoe UI", 13, "bold"), width=165).grid(
            row=0, column=0, padx=(0, 8))
        self.hex_target_entry = ctk.CTkEntry(top, placeholder_text="#RRGGBB eingeben…",
                                              font=("Courier New", 13), height=46)
        self.hex_target_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.hex_target_entry.bind("<Return>", self._on_hex_target_enter)
        self.prev = ctk.CTkLabel(top, text="", width=46, height=46,
                                  fg_color="#1a1a1a", corner_radius=23)
        self.prev.grid(row=0, column=2)

        # Gamut-Warnung
        self.gamut_label = ctk.CTkLabel(
            sec1, text="⚠  Zielfarbe möglicherweise nicht erreichbar (ΔE > 25 zu allen Filamenten)",
            fg_color="#7c2d12", corner_radius=6, text_color="#fcd34d",
            font=("Segoe UI", 10))

        # Sequenz-Ergebnis
        self.res = ctk.CTkLabel(sec1, text="----------",
                                 font=("Courier New", 58, "bold"), text_color="#4ade80")
        self.res.grid(row=3, column=0, pady=(4, 2))

        # 10 Segmente
        sf = ctk.CTkFrame(sec1, fg_color="transparent")
        sf.grid(row=4, column=0, padx=40, pady=(0, 2), sticky="ew")
        self.segs = []
        for i in range(10):
            lbl = ctk.CTkLabel(sf, text=str(i+1), width=46, height=46,
                                fg_color="#1e293b", corner_radius=8,
                                font=("Segoe UI", 9), text_color="#475569")
            lbl.pack(side="left", expand=True, padx=2)
            self.segs.append(lbl)

        # Gewichtungshinweis
        wh = ctk.CTkFrame(sec1, fg_color="transparent")
        wh.grid(row=5, column=0, padx=40, pady=(0, 6), sticky="ew")
        ctk.CTkLabel(wh, text="L1 unten  1.0", font=("Segoe UI", 8),
                     text_color="#475569").pack(side="left")
        ctk.CTkLabel(wh, text="— progressive Gewichtung →", font=("Segoe UI", 8),
                     text_color="#334155").pack(side="left", expand=True)
        ctk.CTkLabel(wh, text="L10 oben  1.5", font=("Segoe UI", 8),
                     text_color="#475569").pack(side="right")

        # Mix-Vorschau + ΔE
        mf = ctk.CTkFrame(sec1, fg_color="#1e293b", corner_radius=8)
        mf.grid(row=6, column=0, padx=40, pady=6, sticky="ew")
        mf.grid_columnconfigure(2, weight=1)
        ctk.CTkLabel(mf, text="Ziel", font=("Segoe UI", 10),
                     text_color="#64748b").grid(row=0, column=0, padx=(14, 4), pady=10)
        self.target_circle = ctk.CTkLabel(mf, text="", width=42, height=42,
                                           fg_color="#1e293b", corner_radius=21)
        self.target_circle.grid(row=0, column=1, padx=4)
        self.de_label = ctk.CTkLabel(mf, text="ΔE  —",
                                      font=("Segoe UI", 16, "bold"), text_color="#475569")
        self.de_label.grid(row=0, column=2, padx=14)
        self.sim_circle = ctk.CTkLabel(mf, text="", width=42, height=42,
                                        fg_color="#1e293b", corner_radius=21)
        self.sim_circle.grid(row=0, column=3, padx=4)
        ctk.CTkLabel(mf, text="Simuliert", font=("Segoe UI", 10),
                     text_color="#64748b").grid(row=0, column=4, padx=(4, 14))

        # Sequenzlänge + Auto-Modus
        lr = ctk.CTkFrame(sec1, fg_color="#1a2535", corner_radius=8)
        lr.grid(row=7, column=0, padx=40, pady=(4, 4), sticky="ew")
        lr.grid_columnconfigure(1, weight=1)

        self.len_label = ctk.CTkLabel(lr, text="Länge: 10",
                                       font=("Segoe UI", 11, "bold"), width=80)
        self.len_label.grid(row=0, column=0, padx=(14, 6), pady=10)

        self.len_slider = ctk.CTkSlider(lr, from_=1, to=10, number_of_steps=9,
                                         command=self._on_len_slider)
        self.len_slider.set(10)
        self.len_slider.grid(row=0, column=1, sticky="ew", padx=6)

        self.auto_len_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(lr, text="Auto\n(kürzeste)", variable=self.auto_len_var,
                        font=("Segoe UI", 9),
                        command=self._on_auto_toggle).grid(row=0, column=2, padx=(8, 4))

        ctk.CTkLabel(lr, text="ΔE≤", font=("Segoe UI", 10),
                     text_color="#64748b").grid(row=0, column=3, padx=(4, 0))
        self.auto_thresh_entry = ctk.CTkEntry(lr, width=46, placeholder_text="2.0")
        self.auto_thresh_entry.insert(0, "2.0")
        self.auto_thresh_entry.grid(row=0, column=4, padx=(2, 14))

        # Steuerleiste
        bl = ctk.CTkFrame(sec1, fg_color="transparent")
        bl.grid(row=8, column=0, padx=40, pady=(4, 16), sticky="ew")
        bl.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(bl, text="SEQUENZ BERECHNEN", fg_color="#2563eb",
                      command=self.calc, height=46,
                      font=("Segoe UI", 14, "bold")).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.optimizer_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(bl, text="Optimizer\n(24 Combos)", variable=self.optimizer_var,
                        font=("Segoe UI", 9)).grid(row=0, column=1, padx=(0, 6))
        ctk.CTkButton(bl, text="➕  Als virtuellen\nDruckkopf hinzufügen",
                      fg_color="#15803d", hover_color="#166534",
                      command=self.add_to_virtual, height=46,
                      font=("Segoe UI", 11, "bold")).grid(row=0, column=2, padx=(0, 6))
        ctk.CTkButton(bl, text="EXPORT", fg_color="#7c3aed",
                      command=self.open_export_dialog, height=46, width=90,
                      font=("Segoe UI", 12, "bold")).grid(row=0, column=3)

        # ── ABSCHNITT 2: VIRTUELLE DRUCKKÖPFE ────────────────────────────────
        sec2_hdr = ctk.CTkFrame(self.main, fg_color="transparent")
        sec2_hdr.grid(row=1, column=0, padx=0, pady=(0, 6), sticky="ew")
        sec2_hdr.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(sec2_hdr,
                     text="VIRTUELLE DRUCKKÖPFE  V5 – V24",
                     font=("Segoe UI", 15, "bold"), text_color="#a78bfa").grid(
            row=0, column=0, sticky="w")
        ctk.CTkLabel(sec2_hdr,
                     text=f"Jeder virtuelle Kopf = FullSpectrum-Sequenz (1–10 Layer) aus T1–T4  ·  max. {MAX_VIRTUAL}",
                     font=("Segoe UI", 10), text_color="#64748b").grid(
            row=1, column=0, sticky="w")

        btn_row2 = ctk.CTkFrame(sec2_hdr, fg_color="transparent")
        btn_row2.grid(row=0, column=1, sticky="e", padx=(10, 0))
        ctk.CTkButton(btn_row2, text="🔬  3MF Assistent", fg_color="#0f4c81",
                      hover_color="#1e3a5f", height=40, width=160,
                      font=("Segoe UI", 12, "bold"),
                      command=self.open_3mf_assistant).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btn_row2, text="🗑  Alle löschen", fg_color="#7f1d1d",
                      hover_color="#991b1b", height=40, width=120,
                      command=self.clear_virtual).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btn_row2, text="📤  Alle exportieren", fg_color="#374151",
                      height=40, width=140,
                      command=self.open_export_dialog).pack(side="left")

        # Virtual Filament Grid Header
        gh = ctk.CTkFrame(self.main, fg_color="#1e293b", corner_radius=6)
        gh.grid(row=2, column=0, sticky="ew", pady=(0, 2))
        gh.grid_columnconfigure(4, weight=1)
        for col, (txt, w) in enumerate([("ID", 55), ("Ziel", 50),
                                         ("Sequenz", 160), ("Simuliert", 60),
                                         ("ΔE / Qualität", 130), ("Label", 0),
                                         ("", 40)]):
            ctk.CTkLabel(gh, text=txt, font=("Segoe UI", 10, "bold"),
                         text_color="#64748b", width=w).grid(
                row=0, column=col, padx=8, pady=6, sticky="w" if w == 0 else "")

        # Scrollbares Grid
        self.vgrid = ctk.CTkScrollableFrame(self.main, height=380, fg_color="#0f172a",
                                             corner_radius=8)
        self.vgrid.grid(row=3, column=0, sticky="ew", pady=(0, 12))
        self.vgrid.grid_columnconfigure(4, weight=1)
        self._refresh_virtual_grid()

    # ── SLOT-LOGIK ─────────────────────────────────────────────────────────────

    def update_menu(self, idx):
        b = self.slots[idx]["brand"].get()
        cols = [f["name"] for f in self.library.get(b, [])] or ["(leer)"]
        self.slots[idx]["color"].configure(values=cols)
        self.slots[idx]["color"].set(cols[0])
        self.apply_f(idx)

    def apply_f(self, idx):
        b = self.slots[idx]["brand"].get()
        n = self.slots[idx]["color"].get()
        if n in ("(leer)", "(manuell)"): return
        f = next((x for x in self.library.get(b, []) if x["name"] == n), None)
        if f is None: return
        self.slots[idx]["hex"].delete(0, "end"); self.slots[idx]["hex"].insert(0, f["hex"])
        self.slots[idx]["td"].delete(0, "end");  self.slots[idx]["td"].insert(0, str(f["td"]))
        self.slots[idx]["preview"].configure(fg_color=f["hex"])

    def pick_slot_color(self, idx):
        cur = self.slots[idx]["hex"].get().strip() or "#808080"
        r = colorchooser.askcolor(color=cur, title=f"Farbe für T{idx+1}")
        if r[1] is None: return
        h = r[1].upper()
        self.slots[idx]["hex"].delete(0, "end"); self.slots[idx]["hex"].insert(0, h)
        self.slots[idx]["preview"].configure(fg_color=h)
        cur_vals = [f["name"] for f in self.library.get(self.slots[idx]["brand"].get(), [])]
        self.slots[idx]["color"].configure(values=["(manuell)"] + cur_vals)
        self.slots[idx]["color"].set("(manuell)")

    def save_current(self, idx):
        n = ctk.CTkInputDialog(text="Filament-Name:", title="In Favoriten speichern").get_input()
        if not n: return
        self.library.setdefault("Eigene Favoriten", []).append(
            {"name": n.strip(), "hex": self.slots[idx]["hex"].get(),
             "td": safe_td(self.slots[idx]["td"].get())})
        self.save_db(); self._refresh_brand_menus()
        messagebox.showinfo("Gespeichert", f'"{n}" gespeichert.')

    def add_filament(self, idx):
        b = self.slots[idx]["brand"].get()
        n = ctk.CTkInputDialog(text="Name:", title=f"Filament zu '{b}'").get_input()
        if not n: return
        h = ctk.CTkInputDialog(text="Hex (#RRGGBB):", title="Farbe").get_input()
        if not h: return
        h = h.strip()
        if not h.startswith("#"): h = "#" + h
        td_raw = ctk.CTkInputDialog(text=f"TD-Wert (Standard {DEFAULT_TD}):", title="TD").get_input()
        td_val = safe_td(td_raw) if td_raw else DEFAULT_TD
        self.library.setdefault(b, []).append({"name": n.strip(), "hex": h, "td": td_val})
        self.save_db(); self.update_menu(idx)

    def add_brand(self):
        name = ctk.CTkInputDialog(text="Marken-Name:", title="Neue Marke").get_input()
        if not name: return
        name = name.strip()
        if name in self.library:
            messagebox.showwarning("Vorhanden", f'"{name}" existiert bereits.'); return
        self.library[name] = []
        self.save_db(); self._refresh_brand_menus()

    def _refresh_brand_menus(self):
        brands = list(self.library.keys())
        for i, s in enumerate(self.slots):
            cur = s["brand"].get()
            s["brand"].configure(values=brands)
            s["brand"].set(cur if cur in brands else brands[0])
            self.update_menu(i)

    # ── ZIELFARBE ──────────────────────────────────────────────────────────────

    def pick(self):
        c = colorchooser.askcolor()[1]
        if c: self._apply_target(c.upper())

    def _on_hex_target_enter(self, event=None):
        raw = self.hex_target_entry.get().strip()
        if not raw.startswith("#"): raw = "#" + raw
        if len(raw) == 7: self._apply_target(raw.upper())

    def _apply_target(self, hex_str):
        self.target = hex_str
        self.prev.configure(fg_color=hex_str)
        self.target_circle.configure(fg_color=hex_str)
        self.hex_target_entry.delete(0, "end")
        self.hex_target_entry.insert(0, hex_str)
        self.calc()

    def _on_len_slider(self, val):
        n = int(val)
        self.len_label.configure(text=f"Länge: {n}")
        # Nicht genutzte Segmente ausblenden (Preview ohne Neuberechnung)
        for i, seg in enumerate(self.segs):
            if i < n: seg.pack(side="left", expand=True, padx=2)
            else:      seg.pack_forget()

    def _on_auto_toggle(self):
        # Slider deaktivieren wenn Auto aktiv
        state = "disabled" if self.auto_len_var.get() else "normal"
        self.len_slider.configure(state=state)

    # ── BERECHNUNGS-KERN ────────────────────────────────────────────────────────

    def _get_fils(self):
        """Liest die 4 physischen Filamente aus den Slots."""
        fils = []
        for i, s in enumerate(self.slots):
            h = s["hex"].get().strip() or "#808080"
            td = safe_td(s["td"].get())
            s["td"].delete(0, "end"); s["td"].insert(0, str(td))
            fils.append({"id": i+1, "lab": rgb_to_lab(hex_to_rgb(h)), "td": td, "hex": h})
        return fils

    def _simulate_mix(self, sequence, fils):
        """Gewichteter Lab-Durchschnitt — funktioniert für jede Sequenzlänge."""
        n = len(sequence)
        weights = get_layer_weights(n)
        total_w = sum(weights)
        lab_sum = [0.0, 0.0, 0.0]
        by_id = {f["id"]: f for f in fils}
        for pos, fid in enumerate(sequence):
            w = weights[pos]
            for j, v in enumerate(by_id[fid]["lab"]):
                lab_sum[j] += w * v
        return tuple(v / total_w for v in lab_sum)

    def _build_sequence(self, ordered, tot, n):
        """Baut eine n-Layer-Sequenz (1–10). Höchster Score → oberste Positionen."""
        counts = [max(0, round((s["w"] / tot) * n)) for s in ordered]
        diff = n - sum(counts)
        i = 0
        while diff > 0: counts[i % len(counts)] += 1; diff -= 1; i += 1
        while diff < 0:
            k = i % len(counts)
            if counts[k] > 0: counts[k] -= 1; diff += 1
            i += 1
        seq = [None] * n; pos = n - 1
        for j, s in enumerate(ordered):
            for _ in range(counts[j]):
                if pos >= 0: seq[pos] = s["id"]; pos -= 1
        return [seq[k] or ordered[0]["id"] for k in range(n)]

    def _calc_for_color(self, target_hex, optimizer=False, seq_len=None, auto=False,
                        auto_threshold=2.0):
        """Berechnet Sequenz für eine Farbe — ohne UI-Seiteneffekte.

        seq_len : feste Länge 1–10 (None = aktueller Slider-Wert)
        auto    : findet kürzeste Länge mit ΔE < auto_threshold
        """
        t_lab = rgb_to_lab(hex_to_rgb(target_hex))
        fils  = self._get_fils()
        scores = [
            {"id": f["id"],
             "w": (1 / (delta_e(t_lab, f["lab"]) + 0.1)) * (10 / f["td"]),
             "h": f["hex"]}
            for f in fils
        ]
        tot = sum(s["w"] for s in scores)
        if tot == 0: return None

        def best_seq_for_n(n):
            if optimizer:
                best, best_dv = None, float("inf")
                for perm in iter_permutations(scores):
                    seq = self._build_sequence(list(perm), tot, n)
                    dv  = delta_e(self._simulate_mix(seq, fils), t_lab)
                    if dv < best_dv: best_dv = dv; best = seq
                return best
            return self._build_sequence(
                sorted(scores, key=lambda x: x["w"], reverse=True), tot, n)

        if auto:
            chosen_seq = None
            for n in range(1, MAX_SEQ_LEN + 1):
                seq = best_seq_for_n(n)
                dv  = delta_e(self._simulate_mix(seq, fils), t_lab)
                if dv <= auto_threshold:
                    chosen_seq = seq
                    break
            if chosen_seq is None:
                chosen_seq = best_seq_for_n(MAX_SEQ_LEN)
        else:
            n = seq_len if seq_len is not None else MAX_SEQ_LEN
            chosen_seq = best_seq_for_n(n)

        sim_lab = self._simulate_mix(chosen_seq, fils)
        dv      = delta_e(sim_lab, t_lab)
        return {
            "target_hex": target_hex,
            "sequence":   "".join(map(str, chosen_seq)),
            "sim_hex":    lab_to_hex(sim_lab),
            "de":         dv,
            "seq_len":    len(chosen_seq),
        }

    def calc(self):
        if not hasattr(self, "target"):
            messagebox.showinfo("Hinweis", "Bitte zuerst eine Zielfarbe auswählen."); return

        t_lab = rgb_to_lab(hex_to_rgb(self.target))
        fils  = self._get_fils()

        # Gamut
        if min(delta_e(t_lab, f["lab"]) for f in fils) > GAMUT_WARN_DE:
            self.gamut_label.grid(row=2, column=0, padx=20, pady=(0, 4), sticky="ew")
        else:
            self.gamut_label.grid_forget()

        auto      = self.auto_len_var.get()
        threshold = safe_td(self.auto_thresh_entry.get()) if auto else 2.0
        seq_len   = int(self.len_slider.get()) if not auto else None

        result = self._calc_for_color(
            self.target, self.optimizer_var.get(),
            seq_len=seq_len, auto=auto, auto_threshold=threshold)
        if result is None: return

        seq = result["sequence"]
        n   = result["seq_len"]

        # Slider-Wert auf tatsächliche Länge setzen (bei Auto)
        if auto:
            self.len_slider.set(n)
            self.len_label.configure(text=f"Länge: {n}")

        self.res.configure(text=seq)

        # Segmente dynamisch aktualisieren
        fils_hex = {f["id"]: f["hex"] for f in fils}
        for i, seg in enumerate(self.segs):
            if i < n:
                seg.configure(fg_color=fils_hex.get(int(seq[i]), "#1e293b"), text="")
                seg.pack(side="left", expand=True, padx=2)
            else:
                seg.pack_forget()

        self.sim_circle.configure(fg_color=result["sim_hex"])
        self.target_circle.configure(fg_color=self.target)
        dv = result["de"]
        self.de_label.configure(text=de_label_text(dv), text_color=de_color(dv))

        self.last_result = result

    # ── VIRTUELLE DRUCKKÖPFE ───────────────────────────────────────────────────

    def add_to_virtual(self, result=None):
        if len(self.virtual_fils) >= MAX_VIRTUAL:
            messagebox.showwarning("Maximum", f"Bereits {MAX_VIRTUAL} virtuelle Druckköpfe definiert.")
            return
        if result is None:
            result = self.last_result
        if not result:
            messagebox.showinfo("Hinweis", "Bitte zuerst eine Sequenz berechnen."); return
        vid = 5 + len(self.virtual_fils)
        self.virtual_fils.append({
            "vid":        vid,
            "target_hex": result["target_hex"],
            "sequence":   result["sequence"],
            "sim_hex":    result["sim_hex"],
            "de":         result["de"],
            "label":      f"Virtuell {vid}",
        })
        self._refresh_virtual_grid()

    def clear_virtual(self):
        if self.virtual_fils and messagebox.askyesno(
                "Löschen", "Alle virtuellen Druckköpfe löschen?"):
            self.virtual_fils.clear()
            self._refresh_virtual_grid()

    def _refresh_virtual_grid(self):
        for w in self.vgrid.winfo_children():
            w.destroy()
        if not self.virtual_fils:
            ctk.CTkLabel(self.vgrid,
                         text="Noch keine virtuellen Druckköpfe — "
                              "Einzelfarbe berechnen und '➕ Hinzufügen' klicken, "
                              "oder '🔬 3MF Assistent' öffnen.",
                         text_color="#334155", font=("Segoe UI", 11),
                         wraplength=600).pack(pady=30)
            return

        self.vgrid.grid_columnconfigure(4, weight=1)
        for row_data in self.virtual_fils:
            self._build_virtual_row(row_data)

    def _build_virtual_row(self, vf):
        outer = ctk.CTkFrame(self.vgrid, fg_color="#1e293b", corner_radius=7)
        outer.pack(fill="x", padx=6, pady=3)
        outer.grid_columnconfigure(4, weight=1)

        # Zeile 0: Haupt-Info
        ctk.CTkLabel(outer, text=f"V{vf['vid']}",
                     font=("Segoe UI", 13, "bold"), text_color="#a78bfa",
                     width=50).grid(row=0, column=0, padx=(10, 4), pady=(8, 2))
        ctk.CTkLabel(outer, text="", width=36, height=36,
                     fg_color=vf["target_hex"], corner_radius=18).grid(
            row=0, column=1, padx=4, pady=(8, 2))
        ctk.CTkLabel(outer, text=vf["sequence"],
                     font=("Courier New", 17, "bold"), text_color="#4ade80",
                     width=155).grid(row=0, column=2, padx=6, pady=(8, 2))
        ctk.CTkLabel(outer, text="", width=36, height=36,
                     fg_color=vf["sim_hex"], corner_radius=18).grid(
            row=0, column=3, padx=4, pady=(8, 2))
        ctk.CTkLabel(outer, text=de_label_text(vf["de"]),
                     font=("Segoe UI", 11, "bold"), text_color=de_color(vf["de"])).grid(
            row=0, column=4, padx=6, sticky="w", pady=(8, 2))
        lbl_entry = ctk.CTkEntry(outer, width=155, font=("Segoe UI", 11))
        lbl_entry.insert(0, vf["label"])
        lbl_entry.bind("<FocusOut>", lambda e, vfd=vf, en=lbl_entry:
                       vfd.update({"label": en.get()}))
        lbl_entry.grid(row=0, column=5, padx=6, pady=(8, 2))
        ctk.CTkButton(outer, text="✕", width=30, height=30,
                      fg_color="#7f1d1d", hover_color="#991b1b",
                      command=lambda vid=vf["vid"]: self._remove_virtual(vid)).grid(
            row=0, column=6, padx=(4, 10), pady=(8, 2))

        # Zeile 1: Runs-Visualisierung + Cadence-Hinweis
        info_row = ctk.CTkFrame(outer, fg_color="transparent")
        info_row.grid(row=1, column=0, columnspan=7, padx=10, pady=(0, 8), sticky="w")

        runs = seq_to_runs(vf["sequence"])
        n_fils = seq_filament_count(vf["sequence"])
        lh = safe_td(self.layer_height_entry.get()) if hasattr(self, "layer_height_entry") else 0.2

        # Farbige Run-Blöcke
        fils_hex = {s["id"]: s["hex"] for s in [
            {"id": i+1, "hex": sl["hex"].get()} for i, sl in enumerate(self.slots)]}
        for fid, cnt in runs:
            col = fils_hex.get(fid, "#64748b")
            ctk.CTkLabel(info_row,
                         text=f"T{fid}×{cnt}",
                         width=max(28, cnt*14), height=20,
                         fg_color=col, corner_radius=4,
                         font=("Segoe UI", 9, "bold"), text_color="#000000").pack(
                side="left", padx=2)

        # Cadence-Hinweis
        pattern_str = "/".join(vf["sequence"])   # "1121" → "1/1/2/1"

        if n_fils == 1:
            hint = "  →  Reine Farbe — kein Mix nötig"
            hint_color = "#94a3b8"
        elif n_fils == 2 and lh > 0:
            cad = calc_cadence(vf["sequence"], lh)
            ids = sorted(cad.keys())
            hint = (f"  →  Cadence A={cad[ids[0]]}mm / B={cad[ids[1]]}mm  "
                    f"oder Pattern: {pattern_str}")
            hint_color = "#4ade80"
        else:
            # 3-4 Filamente → Pattern Mode im Slicer (Others → Dithering → Pattern)
            hint = f"  →  Pattern Mode: {pattern_str}"
            hint_color = "#a78bfa"

        ctk.CTkLabel(info_row, text=hint,
                     font=("Segoe UI", 9), text_color=hint_color).pack(side="left", padx=4)

    def _remove_virtual(self, vid):
        self.virtual_fils = [v for v in self.virtual_fils if v["vid"] != vid]
        # Renumber
        for i, v in enumerate(self.virtual_fils):
            v["vid"] = 5 + i
        self._refresh_virtual_grid()

    # ── 3MF ASSISTENT ──────────────────────────────────────────────────────────

    def open_3mf_assistant(self):
        path = filedialog.askopenfilename(
            filetypes=[("3MF-Dateien", "*.3mf"), ("Alle Dateien", "*.*")],
            title="3MF-Datei für Farbanalyse öffnen")
        if not path: return

        colors, err = parse_3mf_colors(path)
        if not colors:
            messagebox.showinfo("3MF Assistent",
                f"Keine Farbdaten gefunden.\n{err or 'Keine verwertbaren Farben im Modell.'}"); return

        win = ctk.CTkToplevel(self)
        win.title(f"3MF Assistent — {os.path.basename(path)}")
        win.geometry("860x680")
        win.grab_set()

        # Header
        ctk.CTkLabel(win,
                     text=f"3MF Farbanalyse  ·  {len(colors)} Farbe(n) gefunden",
                     font=("Segoe UI", 15, "bold"), text_color="#38bdf8").pack(pady=(16, 2))
        ctk.CTkLabel(win,
                     text=f"Physische Basis: T1={self.slots[0]['hex'].get()}  "
                          f"T2={self.slots[1]['hex'].get()}  "
                          f"T3={self.slots[2]['hex'].get()}  "
                          f"T4={self.slots[3]['hex'].get()}",
                     font=("Segoe UI", 10), text_color="#64748b").pack(pady=(0, 6))

        opt_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(win, text="Optimizer aktivieren (langsamer, aber genauer)",
                        variable=opt_var, font=("Segoe UI", 11)).pack(pady=(0, 4))

        prog_label = ctk.CTkLabel(win, text="Bereit — 'Alle berechnen' drücken.",
                                   font=("Segoe UI", 11), text_color="#94a3b8")
        prog_label.pack(pady=(0, 6))

        # Spalten-Header
        hdr = ctk.CTkFrame(win, fg_color="#1e293b", corner_radius=6)
        hdr.pack(fill="x", padx=14, pady=(0, 4))
        hdr.grid_columnconfigure(4, weight=1)
        for col, txt in enumerate(["#", "Zielfarbe", "Sequenz", "Simuliert",
                                    "ΔE / Qualität", "VID"]):
            ctk.CTkLabel(hdr, text=txt, font=("Segoe UI", 10, "bold"),
                         text_color="#64748b").grid(row=0, column=col, padx=10, pady=6,
                                                     sticky="w")

        # Ergebnis-Scroll
        scroll = ctk.CTkScrollableFrame(win, height=380, fg_color="#0f172a")
        scroll.pack(fill="both", expand=True, padx=14, pady=4)
        scroll.grid_columnconfigure(4, weight=1)

        results = [None] * len(colors)   # Berechnete Ergebnisse
        vid_vars = []                     # BooleanVar: Zeile auswählen

        def render_rows():
            for w in scroll.winfo_children(): w.destroy()
            vid_vars.clear()
            for idx, hex_c in enumerate(colors[:MAX_VIRTUAL]):
                r = results[idx]
                row = ctk.CTkFrame(scroll, fg_color="#1e293b", corner_radius=6)
                row.pack(fill="x", padx=4, pady=2)
                row.grid_columnconfigure(4, weight=1)
                ctk.CTkLabel(row, text=str(idx+1), width=30, font=("Segoe UI", 10),
                             text_color="#64748b").grid(row=0, column=0, padx=8, pady=7)
                ctk.CTkLabel(row, text="", width=34, height=34,
                             fg_color=hex_c, corner_radius=17).grid(row=0, column=1, padx=4)
                if r:
                    ctk.CTkLabel(row, text=r["sequence"],
                                 font=("Courier New", 15, "bold"), text_color="#4ade80",
                                 width=145).grid(row=0, column=2, padx=6)
                    ctk.CTkLabel(row, text="", width=34, height=34,
                                 fg_color=r["sim_hex"], corner_radius=17).grid(row=0, column=3, padx=4)
                    ctk.CTkLabel(row, text=de_label_text(r["de"]),
                                 font=("Segoe UI", 11, "bold"),
                                 text_color=de_color(r["de"])).grid(
                        row=0, column=4, padx=6, sticky="w")
                else:
                    ctk.CTkLabel(row, text="—  noch nicht berechnet",
                                 text_color="#334155").grid(row=0, column=2,
                                                             columnspan=3, padx=6, sticky="w")
                sv = ctk.BooleanVar(value=True if r else False)
                vid_vars.append(sv)
                ctk.CTkCheckBox(row, text="übernehmen", variable=sv,
                                font=("Segoe UI", 9), width=100).grid(
                    row=0, column=5, padx=(4, 10))

        render_rows()

        def run_all():
            opt = opt_var.get()
            free = MAX_VIRTUAL - len(self.virtual_fils)
            to_calc = min(len(colors), free) if free < len(colors) else len(colors)
            for idx in range(min(len(colors), MAX_VIRTUAL)):
                prog_label.configure(
                    text=f"Berechne {idx+1}/{to_calc}  ({colors[idx]}) …")
                win.update_idletasks()
                r = self._calc_for_color(colors[idx], optimizer=opt,
                                        seq_len=int(self.len_slider.get()),
                                        auto=self.auto_len_var.get(),
                                        auto_threshold=safe_td(self.auto_thresh_entry.get()))
                results[idx] = r
            prog_label.configure(text=f"Fertig — {to_calc} Farben berechnet.")
            render_rows()

        def apply_selected():
            added = 0
            for idx, sv in enumerate(vid_vars):
                if not sv.get() or results[idx] is None: continue
                if len(self.virtual_fils) >= MAX_VIRTUAL: break
                self.add_to_virtual(results[idx])
                added += 1
            win.destroy()
            messagebox.showinfo("3MF Assistent",
                f"{added} virtuelle Druckköpfe wurden hinzugefügt.")

        btm = ctk.CTkFrame(win, fg_color="transparent")
        btm.pack(fill="x", padx=14, pady=(6, 14))
        ctk.CTkButton(btm, text="⚙  Alle berechnen", fg_color="#2563eb",
                      command=run_all, height=42, font=("Segoe UI", 12, "bold")).pack(
            side="left", padx=(0, 6))
        ctk.CTkButton(btm, text="✅  Ausgewählte übernehmen", fg_color="#15803d",
                      command=apply_selected, height=42,
                      font=("Segoe UI", 12, "bold")).pack(side="left")
        ctk.CTkButton(btm, text="Abbrechen", fg_color="#374151",
                      command=win.destroy, height=42).pack(side="right")

    # ── EXPORT ─────────────────────────────────────────────────────────────────

    def open_export_dialog(self):
        win = ctk.CTkToplevel(self)
        win.title("Export — Snapmaker U1 FullSpectrum")
        win.geometry("500x440")
        win.grab_set()

        ctk.CTkLabel(win, text="Export — U1 FullSpectrum",
                     font=("Segoe UI", 14, "bold")).pack(pady=(18, 10))

        fmt_var = ctk.StringVar(value="JSON")
        ff = ctk.CTkFrame(win, fg_color="transparent")
        ff.pack()
        ctk.CTkRadioButton(ff, text="JSON (.json)", variable=fmt_var,
                           value="JSON").pack(side="left", padx=12)
        ctk.CTkRadioButton(ff, text="Text (.txt)", variable=fmt_var,
                           value="TXT").pack(side="left", padx=12)

        # Schichthöhe + auto Cadence
        ctk.CTkLabel(win, text="Dithering-Einstellungen (OrcaSlicer FullSpectrum)",
                     font=("Segoe UI", 11, "bold"), text_color="#64748b").pack(pady=(16, 4))
        ctk.CTkLabel(win,
                     text="Schichthöhe = Dithering Step Size  ·  Cadence A/B aus Sequenz auto-berechnet",
                     font=("Segoe UI", 9), text_color="#475569").pack(pady=(0, 6))

        lhf = ctk.CTkFrame(win, fg_color="transparent"); lhf.pack()
        ctk.CTkLabel(lhf, text="Schichthöhe:").pack(side="left", padx=6)
        lh_val = self.layer_height_entry.get() if hasattr(self, "layer_height_entry") else "0.2"
        lh_exp = ctk.CTkEntry(lhf, width=60); lh_exp.insert(0, lh_val)
        lh_exp.pack(side="left", padx=4)
        ctk.CTkLabel(lhf, text="mm  (= Dithering Step Size)").pack(side="left", padx=4)

        cf = ctk.CTkFrame(win, fg_color="transparent"); cf.pack(pady=(6, 0))
        ctk.CTkLabel(cf, text="Cadence A:").pack(side="left", padx=6)
        ca = ctk.CTkEntry(cf, width=70, placeholder_text="auto"); ca.pack(side="left", padx=4)
        ctk.CTkLabel(cf, text="mm     B:").pack(side="left", padx=4)
        cb = ctk.CTkEntry(cf, width=70, placeholder_text="auto"); cb.pack(side="left", padx=4)
        ctk.CTkLabel(cf, text="mm  (leer = auto)", font=("Segoe UI", 9),
                     text_color="#475569").pack(side="left", padx=4)

        scope_var = ctk.StringVar(value="virtual" if self.virtual_fils else "single")
        sf = ctk.CTkFrame(win, fg_color="transparent")
        sf.pack(pady=(14, 4))
        ctk.CTkRadioButton(sf, text="Aktuelle Einzelsequenz",
                           variable=scope_var, value="single").pack(side="left", padx=12)
        ctk.CTkRadioButton(sf, text=f"Alle {len(self.virtual_fils)} virtuellen Druckköpfe",
                           variable=scope_var, value="virtual").pack(side="left", padx=12)

        def do_export():
            fmt   = fmt_var.get()
            ext   = ".json" if fmt == "JSON" else ".txt"
            scope = scope_var.get()
            path  = filedialog.asksaveasfilename(
                defaultextension=ext,
                filetypes=[(fmt, f"*{ext}")],
                title="Speichern")
            if not path: return

            lh     = safe_td(lh_exp.get()) if lh_exp.get().strip() else 0.2
            seq_now = self.res.cget("text") if hasattr(self, "target") else ""

            # Cadence auto aus aktueller Sequenz wenn Felder leer
            def resolve_cadence(seq):
                if not seq: return 0.0, 0.0
                cad = calc_cadence(seq, lh)
                ids = sorted(cad.keys())
                a = safe_td(ca.get()) if ca.get().strip() else (cad.get(ids[0], lh) if ids else lh)
                b = safe_td(cb.get()) if cb.get().strip() else (cad.get(ids[1], lh) if len(ids) > 1 else lh)
                return round(a, 4), round(b, 4)

            physical = [
                {"id": i+1, "brand": s["brand"].get(),
                 "filament": s["color"].get(), "hex": s["hex"].get(),
                 "td": safe_td(s["td"].get())}
                for i, s in enumerate(self.slots)
            ]

            def vf_export_entry(v):
                a_v, b_v = resolve_cadence(v["sequence"])
                runs  = seq_to_runs(v["sequence"])
                n_f   = seq_filament_count(v["sequence"])
                pat   = "/".join(v["sequence"])
                mode  = "cadence" if n_f <= 2 else "pattern"
                return {**v, "runs": runs, "filament_count": n_f,
                        "slicer_input_mode": mode,
                        "pattern_string": pat,
                        "cadence_a_mm": a_v if mode == "cadence" else None,
                        "cadence_b_mm": b_v if mode == "cadence" else None,
                        "dithering_step_size_mm": lh}

            try:
                if fmt == "JSON":
                    n_now = int(self.len_slider.get()) if not self.auto_len_var.get() else len(seq_now)
                    lw = get_layer_weights(n_now)
                    payload = {
                        "generator":            "U1 FullSpectrum Ultimate",
                        "timestamp":            datetime.now().isoformat(),
                        "dithering_step_size":  lh,
                        "layer_weights":        {f"L{i+1}": round(lw[i], 4) for i in range(len(lw))},
                        "physical_slots":       physical,
                    }
                    if scope == "virtual":
                        payload["virtual_filaments"] = [vf_export_entry(v) for v in self.virtual_fils]
                    else:
                        a_v, b_v = resolve_cadence(seq_now)
                        payload["sequence"]       = seq_now
                        payload["target_hex"]     = getattr(self, "target", "")
                        payload["cadence_a_mm"]   = a_v
                        payload["cadence_b_mm"]   = b_v
                        payload["runs"]           = seq_to_runs(seq_now)
                        payload["slicer_compatible"] = seq_filament_count(seq_now) <= 2
                    with open(path, "w", encoding="utf-8") as f:
                        json.dump(payload, f, indent=2, ensure_ascii=False)
                else:
                    lines = [
                        "=" * 56,
                        "  U1 FullSpectrum Ultimate — OrcaSlicer FullSpectrum Export",
                        "=" * 56,
                        f"Datum:              {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                        f"Schichthöhe:        {lh} mm  (= Dithering Step Size)",
                        "", "Physische Druckköpfe:", "-" * 36,
                    ]
                    for p in physical:
                        lines.append(f"  T{p['id']}: {p['brand']} — {p['filament']}  ({p['hex']}, TD={p['td']})")
                    if scope == "virtual" and self.virtual_fils:
                        lines += ["", "Virtuelle Druckköpfe (OrcaSlicer: Others → Dithering):",
                                  "-" * 56]
                        for v in self.virtual_fils:
                            a_v, b_v = resolve_cadence(v["sequence"])
                            n_f = seq_filament_count(v["sequence"])
                            runs_str = "  ".join(f"T{fid}×{cnt}" for fid, cnt in seq_to_runs(v["sequence"]))
                            pat_str  = "/".join(v["sequence"])
                            if n_f == 1:
                                slicer_hint = "Reine Farbe — kein Mix"
                            elif n_f == 2:
                                slicer_hint = f"Cadence A={a_v}mm / B={b_v}mm  oder Pattern: {pat_str}"
                            else:
                                slicer_hint = f"Pattern Mode: {pat_str}"
                            lines += [
                                f"  V{v['vid']}  [{v['label']}]  Ziel:{v['target_hex']}  ΔE={v['de']:.1f}",
                                f"    Sequenz: {v['sequence']}",
                                f"    Runs:    {runs_str}",
                                f"    Slicer:  {slicer_hint}",
                            ]
                    else:
                        a_v, b_v = resolve_cadence(seq_now)
                        lines += ["", f"Sequenz:  {seq_now}",
                                  f"Ziel:     {getattr(self, 'target', '')}",
                                  f"Cadence:  A={a_v}mm / B={b_v}mm",
                                  f"Runs:     {'  '.join(f'T{fid}×{cnt}' for fid,cnt in seq_to_runs(seq_now))}"]
                    lines += ["", "github.com/ratdoux/OrcaSlicer-FullSpectrum", "=" * 56]
                    with open(path, "w", encoding="utf-8") as f:
                        f.write("\n".join(lines))
                messagebox.showinfo("Export", f"Gespeichert:\n{path}")
                win.destroy()
            except IOError as e:
                messagebox.showerror("Fehler", str(e))

        ctk.CTkButton(win, text="EXPORTIEREN", fg_color="#7c3aed",
                      command=do_export, height=44,
                      font=("Segoe UI", 14, "bold")).pack(pady=(18, 6), padx=40, fill="x")
        ctk.CTkButton(win, text="Abbrechen", fg_color="#334155",
                      command=win.destroy, height=36).pack(padx=40, fill="x")

    # ── BIBLIOTHEK-MANAGER ─────────────────────────────────────────────────────

    def open_library_manager(self):
        win = ctk.CTkToplevel(self)
        win.title("Filament-Bibliothek"); win.geometry("720x560"); win.grab_set()

        top = ctk.CTkFrame(win, fg_color="transparent")
        top.pack(fill="x", padx=14, pady=10)
        ctk.CTkLabel(top, text="Marke:").pack(side="left", padx=5)
        bv = ctk.StringVar(value=list(self.library.keys())[0])
        bm = ctk.CTkOptionMenu(top, variable=bv,
                                values=list(self.library.keys()),
                                command=lambda x: refresh())
        bm.pack(side="left", padx=5, fill="x", expand=True)
        ctk.CTkButton(top, text="Marke löschen", fg_color="#991b1b",
                      command=lambda: del_brand()).pack(side="right", padx=5)

        scroll = ctk.CTkScrollableFrame(win, height=380)
        scroll.pack(fill="both", expand=True, padx=14, pady=5)

        def refresh():
            for w in scroll.winfo_children(): w.destroy()
            brand = bv.get()
            fils = self.library.get(brand, [])
            if not fils:
                ctk.CTkLabel(scroll, text="Keine Filamente.",
                              text_color="#475569").pack(pady=20); return
            for idx, fil in enumerate(fils):
                row = ctk.CTkFrame(scroll, fg_color="#1e293b", corner_radius=6)
                row.pack(fill="x", padx=5, pady=2)
                row.grid_columnconfigure(1, weight=1)
                ctk.CTkLabel(row, text="", width=28, height=28,
                              fg_color=fil["hex"], corner_radius=4).grid(
                    row=0, column=0, padx=(8, 6), pady=6)
                ctk.CTkLabel(row, text=fil["name"],
                              font=("Segoe UI", 12, "bold")).grid(row=0, column=1, sticky="w", padx=5)
                ctk.CTkLabel(row, text=f'{fil["hex"]}  ·  TD: {fil["td"]}',
                              text_color="#64748b", font=("Segoe UI", 10)).grid(
                    row=0, column=2, padx=10)
                ctk.CTkButton(row, text="✕", width=26, height=26, fg_color="#7f1d1d",
                               command=lambda b=brand, i=idx: del_fil(b, i)).grid(
                    row=0, column=3, padx=(0, 8))

        def del_fil(brand, idx):
            n = self.library[brand][idx]["name"]
            if messagebox.askyesno("Löschen", f'"{n}" löschen?'):
                self.library[brand].pop(idx)
                self.save_db(); self._refresh_brand_menus(); refresh()

        def del_brand():
            b = bv.get()
            if b in DEFAULT_LIBRARY and DEFAULT_LIBRARY.get(b):
                messagebox.showerror("Geschützt", "Standard-Marken können nicht gelöscht werden."); return
            if messagebox.askyesno("Löschen", f'Marke "{b}" löschen?'):
                del self.library[b]; self.save_db(); self._refresh_brand_menus()
                nb = list(self.library.keys())
                bm.configure(values=nb); bv.set(nb[0] if nb else ""); refresh()

        refresh()
        btm = ctk.CTkFrame(win, fg_color="transparent")
        btm.pack(fill="x", padx=14, pady=(5, 14))
        ctk.CTkButton(btm, text="+ Filament hinzufügen", fg_color="#15803d",
                      command=lambda: self._lm_add(bv.get(), refresh)).pack(side="left")
        ctk.CTkButton(btm, text="Schließen", fg_color="#334155",
                      command=win.destroy).pack(side="right")

    def _lm_add(self, brand, cb):
        n = ctk.CTkInputDialog(text="Name:", title="Hinzufügen").get_input()
        if not n: return
        h = ctk.CTkInputDialog(text="Hex (#RRGGBB):", title="Farbe").get_input()
        if not h: return
        h = h.strip()
        if not h.startswith("#"): h = "#" + h
        td_r = ctk.CTkInputDialog(text=f"TD (Standard {DEFAULT_TD}):", title="TD").get_input()
        self.library.setdefault(brand, []).append(
            {"name": n.strip(), "hex": h, "td": safe_td(td_r) if td_r else DEFAULT_TD})
        self.save_db(); self._refresh_brand_menus(); cb()


if __name__ == "__main__":
    U1FullSpectrumApp().mainloop()
