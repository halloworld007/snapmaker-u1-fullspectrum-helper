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
try:
    from CTkColorPicker import AskColor as _AskColor
    _HAS_CTKPICKER = True
except ImportError:
    _HAS_CTKPICKER = False

try:
    from CTkToolTip import CTkToolTip as _Tip
    _HAS_TOOLTIP = True
except ImportError:
    _HAS_TOOLTIP = False

try:
    from PIL import Image as _PILImage, ImageDraw as _ImageDraw, ImageTk as _ImageTk
    _HAS_PIL = True
except ImportError:
    _HAS_PIL = False

try:
    import matplotlib
    matplotlib.use("TkAgg")
    import matplotlib.pyplot as _plt
    from mpl_toolkits.mplot3d import Axes3D as _Axes3D   # noqa: F401
    _HAS_MPL = True
except ImportError:
    _HAS_MPL = False

# ── ÜBERSETZUNGEN ─────────────────────────────────────────────────────────────

STRINGS = {
"de": {
    "app_title": "U1 FullSpectrum Ultimate — Pro Edition",
    "lang_btn": "🌐 EN",
    "phys_heads_title": "Physische Druckköpfe  T1–T4",
    "phys_heads_desc": "Diese 4 Filamente sind real im Drucker geladen.",
    "tool_header": "T{i} — WERKZEUG {i}",
    "slot_presets": "SLOT-PRESETS",
    "no_presets": "(keine Presets)",
    "btn_load": "LADEN",
    "btn_save": "SPEICHERN",
    "btn_new_brand": "＋  Neue Marke",
    "btn_library": "📚  Bibliothek verwalten",
    "layer_height_label": "Schichthöhe (mm):",
    "dithering_step": "= Dithering Step Size",
    "sec1_title": "EINZELFARBEN-RECHNER",
    "btn_target_color": "ZIELFARBE  🎨",
    "hex_placeholder": "#RRGGBB eingeben…",
    "gamut_warning": "⚠  Zielfarbe möglicherweise nicht erreichbar (ΔE > 25 zu allen Filamenten)",
    "label_target": "Ziel",
    "label_simulated": "Simuliert",
    "length_label": "Länge: {n}",
    "auto_check": "Auto\n(kürzeste)",
    "btn_calculate": "SEQUENZ BERECHNEN",
    "optimizer_check": "Optimizer\n(24 Combos)",
    "btn_add_virtual": "➕  Als virtuellen\nDruckkopf hinzufügen",
    "weight_bottom": "L1 unten  1.0",
    "weight_arrow": "— progressive Gewichtung →",
    "weight_top": "L10 oben  1.5",
    "sec2_title": "VIRTUELLE DRUCKKÖPFE  V5 – V24",
    "sec2_desc": "Jeder virtuelle Kopf = FullSpectrum-Sequenz (1–10 Layer) aus T1–T4  ·  max. {max_v}",
    "btn_3mf": "🔬  3MF Assistent",
    "btn_del_all": "🗑  Alle löschen",
    "btn_export_all": "📤  Alle exportieren",
    "grid_target": "Ziel",
    "grid_sequence": "Sequenz",
    "grid_simulated": "Simuliert",
    "grid_quality": "ΔE / Qualität",
    "grid_label": "Label",
    "empty_virtual": "Noch keine virtuellen Druckköpfe — Einzelfarbe berechnen und '➕ Hinzufügen' klicken, oder '🔬 3MF Assistent' öffnen.",
    "hint_pure": "  →  Reine Farbe — kein Mix nötig",
    "hint_cadence": "  →  Cadence A={a}mm / B={b}mm  oder Pattern: {p}",
    "hint_pattern": "  →  Pattern Mode: {p}",
    "de_good": "✓ gut",
    "de_ok": "~ ok",
    "de_far": "✗ weit",
    "dlg_db_err_title": "DB-Fehler",
    "dlg_db_err_msg": "filament_db.json Ladefehler:\n{e}\nStandardwerte aktiv.",
    "dlg_save_err": "Speicherfehler",
    "dlg_error": "Fehler",
    "dlg_saved": "Gespeichert",
    "dlg_preset_saved": 'Preset "{name}" gespeichert.',
    "dlg_fil_saved": '"{n}" gespeichert.',
    "dlg_exists": "Vorhanden",
    "dlg_exists_msg": '"{name}" existiert bereits.',
    "dlg_note": "Hinweis",
    "dlg_select_color": "Bitte zuerst eine Zielfarbe auswählen.",
    "dlg_max_virtual": "Maximum",
    "dlg_max_virtual_msg": "Bereits {max_v} virtuelle Druckköpfe definiert.",
    "dlg_no_seq": "Bitte zuerst eine Sequenz berechnen.",
    "dlg_del_title": "Löschen",
    "dlg_del_virtual": "Alle virtuellen Druckköpfe löschen?",
    "dlg_3mf_title": "3MF Assistent",
    "dlg_3mf_no_colors_fallback": "Keine verwertbaren Farben im Modell.",
    "dlg_3mf_added": "{n} virtuelle Druckköpfe wurden hinzugefügt.",
    "dlg_export_saved": "Gespeichert:\n{path}",
    "inp_preset_name": "Preset-Name:",
    "inp_preset_title": "Preset speichern",
    "inp_fil_name": "Filament-Name:",
    "inp_save_fav": "In Favoriten speichern",
    "inp_name": "Name:",
    "inp_add_fil_title": "Filament zu '{b}'",
    "inp_hex": "Hex (#RRGGBB):",
    "inp_color_title": "Farbe",
    "inp_td": "TD-Wert (Standard {td}):",
    "inp_td_title": "TD",
    "inp_brand_name": "Marken-Name:",
    "inp_brand_title": "Neue Marke",
    "inp_add_title": "Hinzufügen",
    "inp_td2": "TD (Standard {td}):",
    "exp_title": "Export — Snapmaker U1 FullSpectrum",
    "exp_header": "Export — U1 FullSpectrum",
    "exp_dither_title": "Dithering-Einstellungen (OrcaSlicer FullSpectrum)",
    "exp_dither_desc": "Schichthöhe = Dithering Step Size  ·  Cadence A/B aus Sequenz auto-berechnet",
    "exp_lh_label": "Schichthöhe:",
    "exp_lh_unit": "mm  (= Dithering Step Size)",
    "exp_cadence_a": "Cadence A:",
    "exp_cadence_sep": "mm     B:",
    "exp_cadence_auto": "mm  (leer = auto)",
    "exp_scope_single": "Aktuelle Einzelsequenz",
    "exp_scope_virtual": "Alle {n} virtuellen Druckköpfe",
    "exp_btn": "EXPORTIEREN",
    "exp_cancel": "Abbrechen",
    "3mf_analysis_title": "3MF Farbanalyse  ·  {n} Farbe(n) gefunden",
    "3mf_basis": "Physische Basis: T1={t1}  T2={t2}  T3={t3}  T4={t4}",
    "3mf_optimizer": "Optimizer aktivieren (langsamer, aber genauer)",
    "3mf_ready": "Bereit — 'Alle berechnen' drücken.",
    "3mf_col_target": "Zielfarbe",
    "3mf_col_seq": "Sequenz",
    "3mf_col_sim": "Simuliert",
    "3mf_col_quality": "ΔE / Qualität",
    "3mf_not_calc": "—  noch nicht berechnet",
    "3mf_include": "übernehmen",
    "3mf_progress": "Berechne {i}/{total}  ({c}) …",
    "3mf_done": "Fertig — {n} Farben berechnet.",
    "3mf_btn_calc": "⚙  Alle berechnen",
    "3mf_btn_apply": "✅  Ausgewählte übernehmen",
    "3mf_btn_cancel": "Abbrechen",
    "lib_title": "Filament-Bibliothek",
    "lib_brand": "Marke:",
    "lib_del_brand": "Marke löschen",
    "lib_no_fils": "Keine Filamente.",
    "lib_del_fil": '"{n}" löschen?',
    "lib_protected": "Geschützt",
    "lib_protected_msg": "Standard-Marken können nicht gelöscht werden.",
    "lib_del_brand_msg": 'Marke "{b}" löschen?',
    "lib_add_fil": "+ Filament hinzufügen",
    "lib_close": "Schließen",
    "txt_date": "Datum:",
    "txt_layer_height": "Schichthöhe:",
    "txt_physical_heads": "Physische Druckköpfe:",
    "txt_virtual_heads": "Virtuelle Druckköpfe (OrcaSlicer: Others → Dithering):",
    "txt_pure": "Reine Farbe — kein Mix",
    "txt_cadence": "Cadence A={a}mm / B={b}mm  oder Pattern: {p}",
    "txt_pattern": "Pattern Mode: {p}",
    "txt_sequence": "Sequenz:",
    "txt_target": "Ziel:",
    "txt_cadence2": "Cadence:",
    "txt_runs": "Runs:",
    "empty_slot": "(leer)",
    "manual_color": "(manuell)",
    "virtual_label_default": "Virtuell {vid}",
    "save_dialog_title": "Speichern",
    "open_3mf_title": "3MF-Datei für Farbanalyse öffnen",
    "3mf_filetypes": "3MF-Dateien",
    "btn_copy": "📋 Kopieren",
    "btn_random": "🎲 Zufall",
    "copied_msg": "Sequenz in Zwischenablage kopiert!",
    "btn_batch": "🎨 Batch-Farben",
    "batch_title": "Batch-Farben berechnen",
    "batch_desc": "Hex-Codes eingeben (eine Farbe pro Zeile, z.B. #FF5500):",
    "batch_btn_calc": "⚙  Alle berechnen & übernehmen",
    "batch_btn_cancel": "Abbrechen",
    "batch_done": "{n} virtuelle Druckköpfe aus Batch hinzugefügt.",
    "batch_warn_max": "Maximale Anzahl virtueller Köpfe erreicht ({max_v}).",
    "settings_title": "Einstellungen",
    "settings_saved": "Einstellungen gespeichert.",
    "tab_calculator": "Rechner",
    "tab_virtual": "Virtuelle Köpfe",
    "color_picker_title": "Farbe wählen — T{i}",
    "target_picker_title": "Zielfarbe wählen",
    "btn_img_pick": "🖼 Aus Bild",
    "img_pick_title": "Bild öffnen — Farbe auswählen",
    "img_filetypes": "Bilddateien",
    "img_instruction": "Klicke auf eine Farbe im Bild",
    "btn_undo": "↩ Rückgängig",
    "undo_empty": "Nichts zum Rückgängigmachen.",
    "tip_calculate": "Sequenz berechnen (Enter)",
    "tip_optimizer": "Testet alle 24 Filament-Reihenfolgen",
    "tip_add_virtual": "Als virtuellen Druckkopf V5-V24 hinzufügen",
    "tip_copy": "Sequenz in Zwischenablage kopieren",
    "tip_export": "Als JSON oder TXT exportieren",
    "tip_random": "Zufällige Zielfarbe",
    "tip_img_pick": "Farbe aus Bild wählen",
    "tip_3mf": "3MF-Datei öffnen und Farben analysieren",
    "tip_batch": "Mehrere Hex-Codes auf einmal berechnen",
    "tip_lang": "Sprache umschalten",
    "tip_td": "Transmission Distance: 0=deckend, 10=transparent",
    "colorinfo_label": "RGB: {r} {g} {b}   HSV: {h:.0f}° {s:.0f}% {v:.0f}%   Lab: {L:.0f} {a:.0f} {b_:.0f}",
    # Einstellungen
    "settings_btn": "⚙  Einstellungen",
    "settings_title": "Einstellungen",
    "settings_max_virtual": "Max. virtuelle Druckköpfe:",
    "settings_theme": "Erscheinungsbild:",
    "settings_theme_dark": "Dunkel",
    "settings_theme_light": "Hell",
    "settings_theme_system": "System",
    "settings_save_btn": "Speichern & Schließen",
    "settings_saved": "Einstellungen gespeichert.",
    # Dithering-Profile
    "dither_profiles": "Dithering-Profile:",
    "dither_fine": "🔬 Fein (kurz)",
    "dither_balanced": "⚖ Ausgewogen",
    "dither_smooth": "🌊 Sanft (lang)",
    "dither_auto": "🤖 Auto",
    "tip_dither_fine": "Kurze Sequenz (2-3 Layer) — präzise Farben, mehr Werkzeugwechsel",
    "tip_dither_balanced": "Mittlere Sequenz (5 Layer) — ausgewogener Kompromiss",
    "tip_dither_smooth": "Lange Sequenz (8-10 Layer) — sanfte Übergänge, weniger Wechsel",
    "tip_dither_auto": "Automatisch kürzeste Sequenz mit ΔE ≤ Schwellwert",
    # Farbsehschwäche
    "colorblind_label": "Simulation:",
    "colorblind_normal": "Normal",
    "colorblind_prot": "Protanopie",
    "colorblind_deut": "Deuteranopie",
    "colorblind_trit": "Tritanopie",
    # 3D Lab Plot
    "btn_lab_plot": "🔭 Lab-Farbraum",
    "tip_lab_plot": "3D-Visualisierung der Filamente und Zielfarbe im CIE-Lab-Farbraum",
    "lab_plot_title": "CIE Lab Farbraum — Filamente & Zielfarbe",
    # Swatch
    "btn_swatch": "🖼 Swatch speichern",
    "tip_swatch": "Farbvergleich (Ziel vs. Simuliert) als PNG speichern",
    "swatch_saved": "Swatch gespeichert:\n{path}",
    # Slicer-Guide
    "btn_slicer_guide": "📖 Slicer-Guide",
    "slicer_guide_title": "Export → OrcaSlicer FullSpectrum",
    # Gradient-Generator
    "btn_gradient": "🌈 Gradient",
    "tip_gradient": "Farbverlauf zwischen zwei Farben berechnen",
    "gradient_title": "Gradient-Generator",
    "gradient_from": "Von:",
    "gradient_to": "Bis:",
    "gradient_steps": "Schritte:",
    "gradient_btn_calc": "⚙  Berechnen & Hinzufügen",
    "gradient_done": "{n} Gradient-Schritte als virtuelle Köpfe hinzugefügt.",
    "gradient_warn_max": "Maximal {n} weitere Schritte passen hinzu.",
    # Web-Update
    "btn_web_update": "🌐 Online-Update",
    "tip_web_update": "Neue Filament-Farben aus der Community-Datenbank laden",
    "web_update_title": "Filament-Datenbank aktualisieren",
    "web_update_fetching": "Lade Daten von GitHub …",
    "web_update_ok": "✅  {n_brands} Marken / {n_fils} Filamente geladen.\n{new} neue Einträge hinzugefügt.",
    "web_update_no_new": "Datenbank ist bereits aktuell.",
    "web_update_err": "Fehler beim Laden:\n{e}",
    # OrcaSlicer Direkt-Export
    "btn_orca_export": "🚀 → OrcaSlicer",
    "tip_orca_export": "Filament-Profile direkt in OrcaSlicer importieren",
    "orca_title": "Direkt-Export → OrcaSlicer",
    "orca_header": "Filament-Profile in OrcaSlicer schreiben",
    "orca_path_label": "OrcaSlicer Profil-Ordner:",
    "orca_path_auto": "(wird automatisch erkannt)",
    "orca_path_browse": "📂 Ändern",
    "orca_scope_phys": "Physische Slots T1–T4 (Farbe + TD-Info)",
    "orca_scope_virt": "Virtuelle Köpfe V5+ (simulierte Farben + Dithering-Notiz)",
    "orca_scope_both": "Beides",
    "orca_prefix_label": "Profil-Prefix:",
    "orca_prefix_hint": "z.B. 'U1' → U1-T1, U1-V5 …",
    "orca_btn_export": "PROFILE SCHREIBEN",
    "orca_btn_cancel": "Abbrechen",
    "orca_success": "✅  {n} Profile erfolgreich nach OrcaSlicer geschrieben!\n\nOrcaSlicer neu starten, um die Profile zu laden.",
    "orca_no_path": "OrcaSlicer-Profilordner nicht gefunden.\nBitte Pfad manuell auswählen.",
    "orca_no_virtual": "Keine virtuellen Druckköpfe vorhanden.",
    "orca_overwrite_confirm": "{n} Profile werden ggf. überschrieben. Fortfahren?",
    "orca_filament_notes_t": "U1 FullSpectrum — T{i} | {brand} {name} | TD={td}",
    "orca_filament_notes_v": "U1 FullSpectrum Sequenz: {seq} | ΔE={de:.1f} | {hint}",
},
"en": {
    "app_title": "U1 FullSpectrum Ultimate — Pro Edition",
    "lang_btn": "🌐 DE",
    "phys_heads_title": "Physical Print Heads  T1–T4",
    "phys_heads_desc": "These 4 filaments are physically loaded in the printer.",
    "tool_header": "T{i} — TOOL {i}",
    "slot_presets": "SLOT PRESETS",
    "no_presets": "(no presets)",
    "btn_load": "LOAD",
    "btn_save": "SAVE",
    "btn_new_brand": "＋  New Brand",
    "btn_library": "📚  Manage Library",
    "layer_height_label": "Layer Height (mm):",
    "dithering_step": "= Dithering Step Size",
    "sec1_title": "SINGLE COLOR CALCULATOR",
    "btn_target_color": "TARGET COLOR  🎨",
    "hex_placeholder": "Enter #RRGGBB…",
    "gamut_warning": "⚠  Target color may be out of gamut (ΔE > 25 to all filaments)",
    "label_target": "Target",
    "label_simulated": "Simulated",
    "length_label": "Length: {n}",
    "auto_check": "Auto\n(shortest)",
    "btn_calculate": "CALCULATE SEQUENCE",
    "optimizer_check": "Optimizer\n(24 Combos)",
    "btn_add_virtual": "➕  Add as Virtual\nPrint Head",
    "weight_bottom": "L1 bottom  1.0",
    "weight_arrow": "— progressive weighting →",
    "weight_top": "L10 top  1.5",
    "sec2_title": "VIRTUAL PRINT HEADS  V5 – V24",
    "sec2_desc": "Each virtual head = FullSpectrum sequence (1–10 layers) from T1–T4  ·  max. {max_v}",
    "btn_3mf": "🔬  3MF Assistant",
    "btn_del_all": "🗑  Delete All",
    "btn_export_all": "📤  Export All",
    "grid_target": "Target",
    "grid_sequence": "Sequence",
    "grid_simulated": "Simulated",
    "grid_quality": "ΔE / Quality",
    "grid_label": "Label",
    "empty_virtual": "No virtual print heads yet — calculate a single color and click '➕ Add', or open the '🔬 3MF Assistant'.",
    "hint_pure": "  →  Pure color — no mix needed",
    "hint_cadence": "  →  Cadence A={a}mm / B={b}mm  or Pattern: {p}",
    "hint_pattern": "  →  Pattern Mode: {p}",
    "de_good": "✓ good",
    "de_ok": "~ ok",
    "de_far": "✗ far",
    "dlg_db_err_title": "DB Error",
    "dlg_db_err_msg": "filament_db.json load error:\n{e}\nUsing defaults.",
    "dlg_save_err": "Save Error",
    "dlg_error": "Error",
    "dlg_saved": "Saved",
    "dlg_preset_saved": 'Preset "{name}" saved.',
    "dlg_fil_saved": '"{n}" saved.',
    "dlg_exists": "Exists",
    "dlg_exists_msg": '"{name}" already exists.',
    "dlg_note": "Note",
    "dlg_select_color": "Please select a target color first.",
    "dlg_max_virtual": "Maximum",
    "dlg_max_virtual_msg": "Already {max_v} virtual print heads defined.",
    "dlg_no_seq": "Please calculate a sequence first.",
    "dlg_del_title": "Delete",
    "dlg_del_virtual": "Delete all virtual print heads?",
    "dlg_3mf_title": "3MF Assistant",
    "dlg_3mf_no_colors_fallback": "No usable colors in the model.",
    "dlg_3mf_added": "{n} virtual print heads added.",
    "dlg_export_saved": "Saved:\n{path}",
    "inp_preset_name": "Preset name:",
    "inp_preset_title": "Save Preset",
    "inp_fil_name": "Filament name:",
    "inp_save_fav": "Save to Favorites",
    "inp_name": "Name:",
    "inp_add_fil_title": "Add filament to '{b}'",
    "inp_hex": "Hex (#RRGGBB):",
    "inp_color_title": "Color",
    "inp_td": "TD value (default {td}):",
    "inp_td_title": "TD",
    "inp_brand_name": "Brand name:",
    "inp_brand_title": "New Brand",
    "inp_add_title": "Add",
    "inp_td2": "TD (default {td}):",
    "exp_title": "Export — Snapmaker U1 FullSpectrum",
    "exp_header": "Export — U1 FullSpectrum",
    "exp_dither_title": "Dithering Settings (OrcaSlicer FullSpectrum)",
    "exp_dither_desc": "Layer Height = Dithering Step Size  ·  Cadence A/B auto-calculated from sequence",
    "exp_lh_label": "Layer Height:",
    "exp_lh_unit": "mm  (= Dithering Step Size)",
    "exp_cadence_a": "Cadence A:",
    "exp_cadence_sep": "mm     B:",
    "exp_cadence_auto": "mm  (empty = auto)",
    "exp_scope_single": "Current Single Sequence",
    "exp_scope_virtual": "All {n} virtual print heads",
    "exp_btn": "EXPORT",
    "exp_cancel": "Cancel",
    "3mf_analysis_title": "3MF Color Analysis  ·  {n} color(s) found",
    "3mf_basis": "Physical basis: T1={t1}  T2={t2}  T3={t3}  T4={t4}",
    "3mf_optimizer": "Enable optimizer (slower but more accurate)",
    "3mf_ready": "Ready — click 'Calculate All'.",
    "3mf_col_target": "Target Color",
    "3mf_col_seq": "Sequence",
    "3mf_col_sim": "Simulated",
    "3mf_col_quality": "ΔE / Quality",
    "3mf_not_calc": "—  not yet calculated",
    "3mf_include": "include",
    "3mf_progress": "Calculating {i}/{total}  ({c}) …",
    "3mf_done": "Done — {n} colors calculated.",
    "3mf_btn_calc": "⚙  Calculate All",
    "3mf_btn_apply": "✅  Add Selected",
    "3mf_btn_cancel": "Cancel",
    "lib_title": "Filament Library",
    "lib_brand": "Brand:",
    "lib_del_brand": "Delete Brand",
    "lib_no_fils": "No filaments.",
    "lib_del_fil": 'Delete "{n}"?',
    "lib_protected": "Protected",
    "lib_protected_msg": "Default brands cannot be deleted.",
    "lib_del_brand_msg": 'Delete brand "{b}"?',
    "lib_add_fil": "+ Add Filament",
    "lib_close": "Close",
    "txt_date": "Date:",
    "txt_layer_height": "Layer height:",
    "txt_physical_heads": "Physical Print Heads:",
    "txt_virtual_heads": "Virtual Print Heads (OrcaSlicer: Others → Dithering):",
    "txt_pure": "Pure color — no mix",
    "txt_cadence": "Cadence A={a}mm / B={b}mm  or Pattern: {p}",
    "txt_pattern": "Pattern Mode: {p}",
    "txt_sequence": "Sequence:",
    "txt_target": "Target:",
    "txt_cadence2": "Cadence:",
    "txt_runs": "Runs:",
    "empty_slot": "(empty)",
    "manual_color": "(manual)",
    "virtual_label_default": "Virtual {vid}",
    "save_dialog_title": "Save",
    "open_3mf_title": "Open 3MF file for color analysis",
    "3mf_filetypes": "3MF Files",
    "btn_copy": "📋 Copy",
    "btn_random": "🎲 Random",
    "copied_msg": "Sequence copied to clipboard!",
    "btn_batch": "🎨 Batch Colors",
    "batch_title": "Batch Color Calculator",
    "batch_desc": "Enter hex codes (one color per line, e.g. #FF5500):",
    "batch_btn_calc": "⚙  Calculate All & Add",
    "batch_btn_cancel": "Cancel",
    "batch_done": "{n} virtual print heads added from batch.",
    "batch_warn_max": "Maximum virtual heads reached ({max_v}).",
    "settings_title": "Settings",
    "settings_saved": "Settings saved.",
    "tab_calculator": "Calculator",
    "tab_virtual": "Virtual Heads",
    "color_picker_title": "Pick Color — T{i}",
    "target_picker_title": "Pick Target Color",
    "btn_img_pick": "🖼 From Image",
    "img_pick_title": "Open Image — Pick Color",
    "img_filetypes": "Image Files",
    "img_instruction": "Click on a color in the image",
    "btn_undo": "↩ Undo",
    "undo_empty": "Nothing to undo.",
    "tip_calculate": "Calculate sequence (Enter)",
    "tip_optimizer": "Tests all 24 filament permutations",
    "tip_add_virtual": "Add as virtual print head V5-V24",
    "tip_copy": "Copy sequence to clipboard",
    "tip_export": "Export as JSON or TXT",
    "tip_random": "Random target color",
    "tip_img_pick": "Pick color from an image file",
    "tip_3mf": "Open 3MF file and analyze colors",
    "tip_batch": "Calculate multiple hex codes at once",
    "tip_lang": "Switch language",
    "tip_td": "Transmission Distance: 0=opaque, 10=transparent",
    "colorinfo_label": "RGB: {r} {g} {b}   HSV: {h:.0f}° {s:.0f}% {v:.0f}%   Lab: {L:.0f} {a:.0f} {b_:.0f}",
    "settings_btn": "⚙  Settings",
    "settings_title": "Settings",
    "settings_max_virtual": "Max. virtual print heads:",
    "settings_theme": "Appearance:",
    "settings_theme_dark": "Dark",
    "settings_theme_light": "Light",
    "settings_theme_system": "System",
    "settings_save_btn": "Save & Close",
    "settings_saved": "Settings saved.",
    "dither_profiles": "Dithering Profiles:",
    "dither_fine": "🔬 Fine (short)",
    "dither_balanced": "⚖ Balanced",
    "dither_smooth": "🌊 Smooth (long)",
    "dither_auto": "🤖 Auto",
    "tip_dither_fine": "Short sequence (2-3 layers) — precise colors, more tool changes",
    "tip_dither_balanced": "Medium sequence (5 layers) — balanced compromise",
    "tip_dither_smooth": "Long sequence (8-10 layers) — smooth transitions, fewer changes",
    "tip_dither_auto": "Automatically find shortest sequence with ΔE ≤ threshold",
    "colorblind_label": "Simulation:",
    "colorblind_normal": "Normal",
    "colorblind_prot": "Protanopia",
    "colorblind_deut": "Deuteranopia",
    "colorblind_trit": "Tritanopia",
    "btn_lab_plot": "🔭 Lab Space",
    "tip_lab_plot": "3D visualization of filaments and target in CIE Lab color space",
    "lab_plot_title": "CIE Lab Color Space — Filaments & Target",
    "btn_swatch": "🖼 Save Swatch",
    "tip_swatch": "Save color comparison (target vs. simulated) as PNG",
    "swatch_saved": "Swatch saved:\n{path}",
    "btn_slicer_guide": "📖 Slicer Guide",
    "slicer_guide_title": "Export → OrcaSlicer FullSpectrum",
    # Gradient generator
    "btn_gradient": "🌈 Gradient",
    "tip_gradient": "Calculate color gradient between two colors",
    "gradient_title": "Gradient Generator",
    "gradient_from": "From:",
    "gradient_to": "To:",
    "gradient_steps": "Steps:",
    "gradient_btn_calc": "⚙  Calculate & Add",
    "gradient_done": "{n} gradient steps added as virtual heads.",
    "gradient_warn_max": "Max {n} more steps fit.",
    # Web update
    "btn_web_update": "🌐 Online Update",
    "tip_web_update": "Load new filament colors from the community database",
    "web_update_title": "Update Filament Database",
    "web_update_fetching": "Fetching data from GitHub …",
    "web_update_ok": "✅  {n_brands} brands / {n_fils} filaments loaded.\n{new} new entries added.",
    "web_update_no_new": "Database is already up to date.",
    "web_update_err": "Error loading data:\n{e}",
    # OrcaSlicer direct export
    "btn_orca_export": "🚀 → OrcaSlicer",
    "tip_orca_export": "Import filament profiles directly into OrcaSlicer",
    "orca_title": "Direct Export → OrcaSlicer",
    "orca_header": "Write Filament Profiles to OrcaSlicer",
    "orca_path_label": "OrcaSlicer profile folder:",
    "orca_path_auto": "(auto-detected)",
    "orca_path_browse": "📂 Browse",
    "orca_scope_phys": "Physical slots T1–T4 (color + TD info)",
    "orca_scope_virt": "Virtual heads V5+ (simulated colors + dithering note)",
    "orca_scope_both": "Both",
    "orca_prefix_label": "Profile prefix:",
    "orca_prefix_hint": "e.g. 'U1' → U1-T1, U1-V5 …",
    "orca_btn_export": "WRITE PROFILES",
    "orca_btn_cancel": "Cancel",
    "orca_success": "✅  {n} profiles written to OrcaSlicer!\n\nRestart OrcaSlicer to load the profiles.",
    "orca_no_path": "OrcaSlicer profile folder not found.\nPlease select the path manually.",
    "orca_no_virtual": "No virtual print heads defined.",
    "orca_overwrite_confirm": "{n} profiles may be overwritten. Continue?",
    "orca_filament_notes_t": "U1 FullSpectrum — T{i} | {brand} {name} | TD={td}",
    "orca_filament_notes_v": "U1 FullSpectrum sequence: {seq} | ΔE={de:.1f} | {hint}",
},
}

_SLOT_SKIP = {"(leer)", "(empty)", "(manuell)", "(manual)"}

# ── KONSTANTEN ────────────────────────────────────────────────────────────────
MAX_SEQ_LEN     = 10
DEFAULT_TD      = 5.0
DE_GOOD         = 3.0
DE_OK           = 6.0
GAMUT_WARN_DE   = 25.0
MAX_VIRTUAL     = 20   # Default; wird aus settings.json überschrieben
MAX_VIRTUAL_HARD = 24  # Absolutes Maximum

DEFAULT_LIBRARY = {
    # ── Bambu Lab Basic PLA ─────────────────────────────────────────────────────
    # Hex-Codes basierend auf Bambu Lab Produktkatalog (bambulab.com)
    # TD-Werte: gemessen/geschätzt für FullSpectrum (höher = transparenter)
    "Bambu Lab Basic": [
        {"name": "Jade White",       "hex": "#F5F5F3", "td": 8.5},
        {"name": "Cream White",      "hex": "#FFFAEF", "td": 8.0},
        {"name": "Silver",           "hex": "#B8BCBE", "td": 2.5},
        {"name": "Light Gray",       "hex": "#A8AAAC", "td": 2.0},
        {"name": "Gray",             "hex": "#6B6D6F", "td": 1.5},
        {"name": "Dark Gray",        "hex": "#414345", "td": 0.8},
        {"name": "Charcoal",         "hex": "#2D2926", "td": 0.4},
        {"name": "Black",            "hex": "#101012", "td": 0.3},
        {"name": "Lemon Yellow",     "hex": "#FFF176", "td": 7.5},
        {"name": "Yellow",           "hex": "#FCE300", "td": 6.5},
        {"name": "Banana Yellow",    "hex": "#FFD54F", "td": 6.0},
        {"name": "Gold",             "hex": "#FFAB00", "td": 5.5},
        {"name": "Tangerine",        "hex": "#FF7043", "td": 5.5},
        {"name": "Orange",           "hex": "#E65100", "td": 4.5},
        {"name": "Flame Red",        "hex": "#D32F2F", "td": 3.5},
        {"name": "Vivid Red",        "hex": "#E53935", "td": 3.5},
        {"name": "Cherry Red",       "hex": "#B71C1C", "td": 2.5},
        {"name": "Coral",            "hex": "#FF6B6B", "td": 5.0},
        {"name": "Sakura Pink",      "hex": "#F8BBD0", "td": 7.5},
        {"name": "Hot Pink",         "hex": "#F06292", "td": 6.5},
        {"name": "Magenta",          "hex": "#EC008C", "td": 8.0},
        {"name": "Fuchsia",          "hex": "#D500F9", "td": 7.0},
        {"name": "Lilac",            "hex": "#CE93D8", "td": 6.0},
        {"name": "Lavender",         "hex": "#B39DDB", "td": 5.5},
        {"name": "Purple",           "hex": "#7B1FA2", "td": 3.5},
        {"name": "Violet",           "hex": "#4527A0", "td": 3.0},
        {"name": "Cobalt Blue",      "hex": "#1565C0", "td": 3.5},
        {"name": "Azure Blue",       "hex": "#1976D2", "td": 4.0},
        {"name": "Blue",             "hex": "#0047AB", "td": 3.5},
        {"name": "Sky Blue",         "hex": "#64B5F6", "td": 6.0},
        {"name": "Baby Blue",        "hex": "#BBDEFB", "td": 7.5},
        {"name": "Cyan",             "hex": "#0086D6", "td": 5.0},
        {"name": "Teal",             "hex": "#00796B", "td": 3.0},
        {"name": "Mint",             "hex": "#A5D6A7", "td": 7.0},
        {"name": "Grass Green",      "hex": "#43A047", "td": 4.5},
        {"name": "Forest Green",     "hex": "#2E7D32", "td": 3.0},
        {"name": "Lime Green",       "hex": "#C6E24E", "td": 7.0},
        {"name": "Olive",            "hex": "#827717", "td": 2.5},
        {"name": "Caramel",          "hex": "#BF7E45", "td": 3.0},
        {"name": "Brown",            "hex": "#795548", "td": 2.0},
        {"name": "Skin",             "hex": "#FFCCB3", "td": 7.5},
    ],
    # ── Bambu Lab Matte PLA ─────────────────────────────────────────────────────
    "Bambu Lab Matte": [
        {"name": "Ivory White",      "hex": "#F2EFDF", "td": 7.0},
        {"name": "Beige",            "hex": "#E8D8B8", "td": 6.0},
        {"name": "Cream",            "hex": "#FFF8E1", "td": 6.5},
        {"name": "Matte Black",      "hex": "#1A1A1A", "td": 0.3},
        {"name": "Charcoal",         "hex": "#2D2926", "td": 0.4},
        {"name": "Stone Gray",       "hex": "#7B7B7B", "td": 1.5},
        {"name": "Sunflower",        "hex": "#FFC107", "td": 6.0},
        {"name": "Terracotta",       "hex": "#C1440E", "td": 3.0},
        {"name": "Brick Red",        "hex": "#A93226", "td": 2.5},
        {"name": "Dusty Rose",       "hex": "#D4848A", "td": 5.0},
        {"name": "Mauve",            "hex": "#BD7EA6", "td": 4.5},
        {"name": "Lilac Purple",     "hex": "#A181C1", "td": 4.0},
        {"name": "Deep Purple",      "hex": "#512DA8", "td": 2.5},
        {"name": "Ocean Blue",       "hex": "#1565C0", "td": 3.0},
        {"name": "Powder Blue",      "hex": "#B0BEC5", "td": 6.0},
        {"name": "Sage Green",       "hex": "#8FAF8B", "td": 4.0},
        {"name": "Moss Green",       "hex": "#6B7C45", "td": 3.0},
        {"name": "Army Green",       "hex": "#4A5240", "td": 2.0},
        {"name": "Desert Sand",      "hex": "#C2B280", "td": 4.5},
        {"name": "Caramel Brown",    "hex": "#9C6330", "td": 2.5},
    ],
    # ── Bambu Lab Silk PLA ──────────────────────────────────────────────────────
    "Bambu Lab Silk": [
        {"name": "Silk Gold",        "hex": "#D4A843", "td": 2.0},
        {"name": "Silk Rose Gold",   "hex": "#B76E79", "td": 2.0},
        {"name": "Silk Copper",      "hex": "#B87333", "td": 1.8},
        {"name": "Silk Bronze",      "hex": "#8C6239", "td": 1.8},
        {"name": "Silk Silver",      "hex": "#C0C4C8", "td": 2.5},
        {"name": "Silk Black",       "hex": "#1A1A1A", "td": 0.5},
        {"name": "Silk Red",         "hex": "#B71C1C", "td": 2.5},
        {"name": "Silk Ruby",        "hex": "#9B111E", "td": 2.2},
        {"name": "Silk Sapphire",    "hex": "#1A237E", "td": 2.0},
        {"name": "Silk Jade",        "hex": "#1A6644", "td": 2.0},
        {"name": "Silk Galaxy Black","hex": "#1C1C2E", "td": 0.6},
        {"name": "Silk Rainbow",     "hex": "#FF6B35", "td": 2.5},
        {"name": "Silk Ice",         "hex": "#C9E8F0", "td": 3.5},
    ],
    # ── Prusament PLA ───────────────────────────────────────────────────────────
    # Hex aus Prusa Color Picker, TD geschätzt
    "Prusament PLA": [
        {"name": "Vanilla White",    "hex": "#D9D4C4", "td": 7.0},
        {"name": "Chalk White",      "hex": "#F0EDDE", "td": 7.5},
        {"name": "Jet Black",        "hex": "#24292A", "td": 0.3},
        {"name": "Galaxy Black",     "hex": "#17191A", "td": 0.4},
        {"name": "Grey Matter",      "hex": "#908E8E", "td": 1.8},
        {"name": "Galaxy Silver",    "hex": "#868F98", "td": 2.0},
        {"name": "Prusa Orange",     "hex": "#FE6E31", "td": 6.5},
        {"name": "Carrot Orange",    "hex": "#E9601A", "td": 5.5},
        {"name": "Mango Yellow",     "hex": "#FFB21F", "td": 6.5},
        {"name": "Sunflower Yellow", "hex": "#F9D011", "td": 6.5},
        {"name": "Lipstick Red",     "hex": "#B11A29", "td": 3.0},
        {"name": "Fire Engine Red",  "hex": "#CE1F1A", "td": 3.5},
        {"name": "Terracotta",       "hex": "#C04821", "td": 3.5},
        {"name": "Coral",            "hex": "#E95641", "td": 5.0},
        {"name": "Raspberry Red",    "hex": "#9B1B30", "td": 2.5},
        {"name": "Azure Blue",       "hex": "#0762AD", "td": 3.5},
        {"name": "Ultramarine",      "hex": "#2E3AB1", "td": 3.5},
        {"name": "Cobalt Blue",      "hex": "#004EA1", "td": 3.5},
        {"name": "Mystic Petrol",    "hex": "#1E6E7E", "td": 3.0},
        {"name": "Pirate Blue",      "hex": "#193B81", "td": 2.5},
        {"name": "Gentian Blue",     "hex": "#3558C1", "td": 4.0},
        {"name": "Urban Grey",       "hex": "#595959", "td": 1.2},
        {"name": "Anthracite Grey",  "hex": "#3A3A3A", "td": 0.6},
        {"name": "Lemon Yellow",     "hex": "#F8E045", "td": 7.0},
        {"name": "Mystic Green",     "hex": "#256B3A", "td": 3.0},
    ],
    # ── eSUN PLA+ ───────────────────────────────────────────────────────────────
    "eSUN PLA+": [
        {"name": "Cold White",       "hex": "#F8F8F8", "td": 8.5},
        {"name": "Warm White",       "hex": "#FFF8E7", "td": 8.0},
        {"name": "Black",            "hex": "#0A0A0A", "td": 0.3},
        {"name": "Silver",           "hex": "#C0C0C0", "td": 2.5},
        {"name": "Dark Gray",        "hex": "#404040", "td": 0.8},
        {"name": "Fire Red",         "hex": "#CC1010", "td": 3.5},
        {"name": "Red",              "hex": "#C0392B", "td": 3.0},
        {"name": "Pink",             "hex": "#FF69B4", "td": 6.5},
        {"name": "Orange",           "hex": "#FF5500", "td": 5.5},
        {"name": "Yellow",           "hex": "#FFD700", "td": 6.5},
        {"name": "Lemon Yellow",     "hex": "#FFF44F", "td": 7.0},
        {"name": "Grass Green",      "hex": "#00A86B", "td": 4.5},
        {"name": "Pine Green",       "hex": "#01796F", "td": 3.5},
        {"name": "Teal",             "hex": "#008B8B", "td": 3.5},
        {"name": "Blue",             "hex": "#1560BD", "td": 3.5},
        {"name": "Dark Blue",        "hex": "#00008B", "td": 2.5},
        {"name": "Light Blue",       "hex": "#5B9BD5", "td": 5.5},
        {"name": "Purple",           "hex": "#7B2FBE", "td": 3.5},
        {"name": "Magenta",          "hex": "#E040FB", "td": 7.0},
        {"name": "Skin",             "hex": "#FFDAB9", "td": 7.5},
        {"name": "Brown",            "hex": "#8B4513", "td": 2.0},
        {"name": "Wood",             "hex": "#966432", "td": 2.5},
    ],
    # ── Polymaker PolyTerra PLA ─────────────────────────────────────────────────
    "Polymaker PolyTerra": [
        {"name": "Cotton White",     "hex": "#F5F0EB", "td": 7.5},
        {"name": "Stone White",      "hex": "#E0DDD5", "td": 6.5},
        {"name": "Charcoal",         "hex": "#2D2D2D", "td": 0.4},
        {"name": "Stone Grey",       "hex": "#7C7C70", "td": 1.5},
        {"name": "Sakura Pink",      "hex": "#F4C2C2", "td": 7.0},
        {"name": "Coral",            "hex": "#F28B66", "td": 5.5},
        {"name": "Muted Rose",       "hex": "#C48B8B", "td": 4.5},
        {"name": "Army Red",         "hex": "#9B2335", "td": 2.5},
        {"name": "Savanna Yellow",   "hex": "#C8A020", "td": 5.0},
        {"name": "Muted Orange",     "hex": "#C0672A", "td": 4.0},
        {"name": "Jungle Green",     "hex": "#29AB87", "td": 4.0},
        {"name": "Forest Green",     "hex": "#228B22", "td": 3.5},
        {"name": "Army Green",       "hex": "#4B5320", "td": 2.0},
        {"name": "Muted Teal",       "hex": "#4E8577", "td": 3.5},
        {"name": "Midnight Blue",    "hex": "#1A1A6E", "td": 2.0},
        {"name": "Cloud Blue",       "hex": "#ADC8E6", "td": 6.5},
        {"name": "Indigo",           "hex": "#4B0082", "td": 2.5},
        {"name": "Desert Sand",      "hex": "#C2B280", "td": 4.5},
        {"name": "Coffee Brown",     "hex": "#7B4A25", "td": 2.0},
        {"name": "Cotton Candy",     "hex": "#FFBCD9", "td": 7.5},
    ],
    # ── Hatchbox PLA ────────────────────────────────────────────────────────────
    "Hatchbox PLA": [
        {"name": "White",            "hex": "#FDFDFD", "td": 8.5},
        {"name": "Black",            "hex": "#0D0D0D", "td": 0.3},
        {"name": "Silver",           "hex": "#C0C0C0", "td": 2.5},
        {"name": "Gold",             "hex": "#CFB53B", "td": 2.5},
        {"name": "True Red",         "hex": "#CC1010", "td": 3.5},
        {"name": "Galaxy Red",       "hex": "#8B0000", "td": 2.0},
        {"name": "Pink",             "hex": "#FF69B4", "td": 6.5},
        {"name": "Orange",           "hex": "#FF6600", "td": 5.5},
        {"name": "Yellow",           "hex": "#FFE900", "td": 6.5},
        {"name": "Teal",             "hex": "#008080", "td": 3.5},
        {"name": "Blue",             "hex": "#0047AB", "td": 3.5},
        {"name": "Green",            "hex": "#00A550", "td": 4.5},
        {"name": "Purple",           "hex": "#7B2FBE", "td": 3.5},
        {"name": "Brown",            "hex": "#8B4513", "td": 2.0},
        {"name": "Wood",             "hex": "#8B6914", "td": 2.5},
        {"name": "Glow Green",       "hex": "#39FF14", "td": 7.5},
    ],
    # ── Overture PLA ────────────────────────────────────────────────────────────
    "Overture PLA": [
        {"name": "White",            "hex": "#F8F8F8", "td": 8.5},
        {"name": "Black",            "hex": "#141414", "td": 0.3},
        {"name": "Gray",             "hex": "#787878", "td": 1.5},
        {"name": "Marble White",     "hex": "#E8E8E8", "td": 5.0},
        {"name": "True Red",         "hex": "#C01010", "td": 3.5},
        {"name": "Scarlet Red",      "hex": "#E81010", "td": 3.5},
        {"name": "Pink",             "hex": "#FF69B4", "td": 6.5},
        {"name": "Orange",           "hex": "#FF6600", "td": 5.5},
        {"name": "Yellow",           "hex": "#FFD700", "td": 6.5},
        {"name": "Blue",             "hex": "#2060C0", "td": 3.5},
        {"name": "Sky Blue",         "hex": "#6EB5FF", "td": 6.0},
        {"name": "Green",            "hex": "#009050", "td": 4.5},
        {"name": "Purple",           "hex": "#6B2FBE", "td": 3.5},
        {"name": "Glow Blue",        "hex": "#00BFFF", "td": 7.0},
    ],
    # ── Snapmaker PLA ───────────────────────────────────────────────────────────
    "Snapmaker PLA": [
        {"name": "White",            "hex": "#F8F8F8", "td": 8.5},
        {"name": "Black",            "hex": "#141414", "td": 0.3},
        {"name": "Gray",             "hex": "#808080", "td": 1.5},
        {"name": "Red",              "hex": "#CC2020", "td": 3.5},
        {"name": "Orange",           "hex": "#FF8C00", "td": 5.5},
        {"name": "Yellow",           "hex": "#FFD700", "td": 6.5},
        {"name": "Green",            "hex": "#228B22", "td": 3.5},
        {"name": "Blue",             "hex": "#0047AB", "td": 3.5},
        {"name": "Purple",           "hex": "#7B2FBE", "td": 3.5},
        {"name": "Pink",             "hex": "#FF69B4", "td": 6.5},
    ],
    # ── Anycubic PLA ────────────────────────────────────────────────────────────
    "Anycubic": [
        {"name": "White",            "hex": "#F8F8F6", "td": 8.5},
        {"name": "Silk White",       "hex": "#F0F0EC", "td": 6.5},
        {"name": "Black",            "hex": "#121212", "td": 0.3},
        {"name": "Gray",             "hex": "#808080", "td": 1.5},
        {"name": "Cyan",             "hex": "#00C8FF", "td": 5.5},
        {"name": "Magenta",          "hex": "#FF0090", "td": 7.5},
        {"name": "Yellow",           "hex": "#FFE500", "td": 6.5},
        {"name": "Red",              "hex": "#CC1010", "td": 3.5},
        {"name": "Orange",           "hex": "#FF6000", "td": 5.5},
        {"name": "Blue",             "hex": "#0050C8", "td": 3.5},
        {"name": "Green",            "hex": "#00A040", "td": 4.5},
        {"name": "Purple",           "hex": "#8020C0", "td": 3.5},
        {"name": "Pink",             "hex": "#FF80A0", "td": 7.0},
        {"name": "Gold Silk",        "hex": "#D4A843", "td": 2.0},
        {"name": "Silver Silk",      "hex": "#C0C4C8", "td": 2.5},
    ],
    # ── SUNLU PLA+ ──────────────────────────────────────────────────────────────
    "SUNLU PLA+": [
        {"name": "White",            "hex": "#F5F5F5", "td": 8.5},
        {"name": "Black",            "hex": "#121212", "td": 0.3},
        {"name": "Gray",             "hex": "#7C7C7C", "td": 1.5},
        {"name": "Red",              "hex": "#D01010", "td": 3.5},
        {"name": "Orange",           "hex": "#FF5500", "td": 5.5},
        {"name": "Yellow",           "hex": "#FFDD00", "td": 6.5},
        {"name": "Green",            "hex": "#009050", "td": 4.5},
        {"name": "Blue",             "hex": "#0055CC", "td": 3.5},
        {"name": "Purple",           "hex": "#7020B0", "td": 3.5},
        {"name": "Pink",             "hex": "#FF70A0", "td": 7.0},
        {"name": "Skin",             "hex": "#FFDAB9", "td": 7.5},
        {"name": "Glow Green",       "hex": "#39FF14", "td": 7.5},
        {"name": "Glow Blue",        "hex": "#00BFFF", "td": 7.5},
    ],
    # ── BambuLab PETG HF ────────────────────────────────────────────────────────
    "Bambu Lab PETG": [
        {"name": "Jade White",       "hex": "#F4F2ED", "td": 7.5},
        {"name": "Black",            "hex": "#101010", "td": 0.3},
        {"name": "Translucent",      "hex": "#E8EFF5", "td": 9.5},
        {"name": "Red",              "hex": "#D01515", "td": 3.5},
        {"name": "Yellow",           "hex": "#FFD600", "td": 6.5},
        {"name": "Blue",             "hex": "#0052CC", "td": 3.5},
        {"name": "Green",            "hex": "#00913F", "td": 4.5},
        {"name": "Orange",           "hex": "#FF6A00", "td": 5.5},
        {"name": "Gray",             "hex": "#8A8A8A", "td": 1.5},
    ],
    # ── FullSpectrum CMY-System ─────────────────────────────────────────────────
    # Optimierte Primärfarben für maximale Mischgamut (ΔE-minimiert)
    # Orientiert an Druckfarben-CMY für beste optische Mischung
    "FS CMY System": [
        {"name": "FS White",         "hex": "#F8F8F8", "td": 9.0},
        {"name": "FS Cyan",          "hex": "#00ADEF", "td": 6.0},
        {"name": "FS Magenta",       "hex": "#EC008C", "td": 8.0},
        {"name": "FS Yellow",        "hex": "#FFD700", "td": 7.0},
        {"name": "FS Black",         "hex": "#0A0A0A", "td": 0.3},
        {"name": "FS Orange",        "hex": "#FF6600", "td": 5.5},
        {"name": "FS Blue",          "hex": "#0047AB", "td": 4.0},
        {"name": "FS Green",         "hex": "#00A651", "td": 5.0},
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

def de_label_text(de, lang="de"):
    s = STRINGS[lang]
    q = s["de_good"] if de < DE_GOOD else s["de_ok"] if de < DE_OK else s["de_far"]
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
        self.settings_file = "settings.json"
        self.settings      = self._load_settings()
        self.lang          = self.settings.get("lang", "de")
        self.title(STRINGS[self.lang]["app_title"])
        self.geometry("1420x1000")
        ctk.set_appearance_mode(self.settings.get("theme", "dark"))
        self.db_file       = "filament_db.json"
        self.preset_file   = "presets.json"
        self.history       = []
        self.presets       = {}
        self.virtual_fils  = []   # list of virtual filament dicts
        self.virtual_undo  = []   # undo stack for virtual heads
        self.last_result   = {}   # last calc() result for "hinzufügen"
        self._max_virtual  = int(self.settings.get("max_virtual", MAX_VIRTUAL))
        self.load_db()
        self.load_presets()
        self.setup_ui()
        # Einstellungen wiederherstellen
        lh = self.settings.get("layer_height", "0.2")
        self.layer_height_entry.delete(0, "end")
        self.layer_height_entry.insert(0, lh)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        self._save_settings()
        self.destroy()

    def t(self, key, **kwargs):
        s = STRINGS[self.lang].get(key, STRINGS["de"].get(key, key))
        return s.format(**kwargs) if kwargs else s

    def tip(self, widget, key):
        """Tooltip hinzufügen wenn CTkToolTip verfügbar."""
        if _HAS_TOOLTIP:
            _Tip(widget, message=self.t(key), delay=600)

    def de_label(self, de):
        return de_label_text(de, self.lang)

    def _load_settings(self):
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _save_settings(self):
        self.settings["lang"]         = self.lang
        self.settings["max_virtual"]  = self._max_virtual
        self.settings["layer_height"] = self.layer_height_entry.get() if hasattr(self, "layer_height_entry") else "0.2"
        try:
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=2)
        except IOError:
            pass

    def _save_slot_values(self):
        return [{"brand": s["brand"].get(), "hex": s["hex"].get(),
                 "td": s["td"].get()} for s in self.slots]

    def _restore_slot_values(self, vals):
        for i, v in enumerate(vals):
            s = self.slots[i]
            if v["brand"] in self.library:
                s["brand"].set(v["brand"])
                self.update_menu(i)
            if v["hex"]:
                s["hex"].delete(0, "end"); s["hex"].insert(0, v["hex"])
                s["preview"].configure(fg_color=v["hex"])
            if v["td"]:
                s["td"].delete(0, "end"); s["td"].insert(0, v["td"])

    def toggle_lang(self):
        self._save_settings()
        slot_vals = self._save_slot_values()
        target    = getattr(self, "target", None)
        self.lang = "en" if self.lang == "de" else "de"
        self.title(self.t("app_title"))
        for w in self.winfo_children():
            w.destroy()
        self.setup_ui()
        self._restore_slot_values(slot_vals)
        if target:
            self._apply_target(target)
        self._save_settings()

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
            messagebox.showwarning(self.t("dlg_db_err_title"),
                self.t("dlg_db_err_msg", e=e))

    def save_db(self):
        try:
            with open(self.db_file, "w", encoding="utf-8") as f:
                json.dump(self.library, f, indent=2, ensure_ascii=False)
        except IOError as e:
            messagebox.showerror(self.t("dlg_save_err"), str(e))

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
            messagebox.showerror(self.t("dlg_error"), str(e))

    def save_preset(self):
        name = ctk.CTkInputDialog(text=self.t("inp_preset_name"), title=self.t("inp_preset_title")).get_input()
        if not name: return
        self.presets[name.strip()] = [
            {"brand": s["brand"].get(), "filament": s["color"].get(),
             "hex": s["hex"].get(), "td": safe_td(s["td"].get())}
            for s in self.slots
        ]
        self.save_presets(); self._refresh_preset_dropdown()
        self.preset_var.set(name.strip())
        messagebox.showinfo(self.t("dlg_saved"), self.t("dlg_preset_saved", name=name))

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
        names = list(self.presets.keys()) or [self.t("no_presets")]
        self.preset_dropdown.configure(values=names)
        if self.preset_var.get() not in names: self.preset_var.set(names[0])

    # ── UI-AUFBAU ──────────────────────────────────────────────────────────────

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── SIDEBAR: PHYSISCHE SLOTS ─────────────────────────────────────────
        self.sidebar = ctk.CTkScrollableFrame(self, width=430, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew", padx=(10, 5), pady=10)

        # Sprach-Toggle
        lang_row = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        lang_row.pack(fill="x", padx=12, pady=(6, 0))
        lang_btn = ctk.CTkButton(lang_row, text=self.t("lang_btn"), width=70, height=26,
                      fg_color="#1e3a5f", font=("Segoe UI", 10),
                      command=self.toggle_lang)
        lang_btn.pack(side="right")
        self.tip(lang_btn, "tip_lang")

        sett_btn = ctk.CTkButton(lang_row, text=self.t("settings_btn"), width=120, height=26,
                      fg_color="#374151", font=("Segoe UI", 10),
                      command=self.open_settings_dialog)
        sett_btn.pack(side="right", padx=(0, 6))

        ctk.CTkLabel(self.sidebar, text=self.t("phys_heads_title"),
                     font=("Segoe UI", 18, "bold"), text_color="#38bdf8").pack(pady=(8, 4))
        ctk.CTkLabel(self.sidebar,
                     text=self.t("phys_heads_desc"),
                     font=("Segoe UI", 10), text_color="#475569").pack(pady=(0, 10))

        self.slots = []
        for i in range(4):
            frame = ctk.CTkFrame(self.sidebar, border_width=1, border_color="#334155")
            frame.pack(fill="x", padx=12, pady=5)

            hdr = ctk.CTkFrame(frame, fg_color="transparent")
            hdr.pack(fill="x", padx=5, pady=(6, 0))
            ctk.CTkLabel(hdr, text=self.t("tool_header", i=i+1),
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
        ctk.CTkLabel(self.sidebar, text=self.t("slot_presets"),
                     font=("Segoe UI", 10, "bold"), text_color="#64748b").pack(pady=(12, 2))
        pf = ctk.CTkFrame(self.sidebar, fg_color="#0f172a", corner_radius=8)
        pf.pack(fill="x", padx=12, pady=(0, 5))
        self.preset_var = ctk.StringVar(value=self.t("no_presets"))
        self.preset_dropdown = ctk.CTkOptionMenu(pf, variable=self.preset_var,
                                                  values=[self.t("no_presets")])
        self.preset_dropdown.pack(fill="x", padx=10, pady=(8, 4))
        pb = ctk.CTkFrame(pf, fg_color="transparent")
        pb.pack(fill="x", padx=10, pady=(0, 8))
        pb.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(pb, text=self.t("btn_load"), fg_color="#1e3a5f",
                      command=self.load_preset).grid(row=0, column=0, sticky="ew", padx=(0, 3))
        ctk.CTkButton(pb, text=self.t("btn_save"), fg_color="#1e3a5f",
                      command=self.save_preset).grid(row=0, column=1, sticky="ew", padx=(3, 0))
        self._refresh_preset_dropdown()

        # Sidebar-Tools
        ctk.CTkButton(self.sidebar, text=self.t("btn_new_brand"), fg_color="#1e3a5f",
                      command=self.add_brand).pack(fill="x", padx=12, pady=(8, 3))
        ctk.CTkButton(self.sidebar, text=self.t("btn_library"), fg_color="#374151",
                      command=self.open_library_manager).pack(fill="x", padx=12, pady=(3, 2))
        web_btn = ctk.CTkButton(self.sidebar, text=self.t("btn_web_update"),
                      fg_color="#164e63", hover_color="#155e75",
                      command=self.web_update_library)
        web_btn.pack(fill="x", padx=12, pady=(0, 8))
        self.tip(web_btn, "tip_web_update")

        # Schichthöhe — global für Cadence-Berechnung
        lh_frame = ctk.CTkFrame(self.sidebar, fg_color="#0f172a", corner_radius=8)
        lh_frame.pack(fill="x", padx=12, pady=(0, 15))
        lh_inner = ctk.CTkFrame(lh_frame, fg_color="transparent")
        lh_inner.pack(fill="x", padx=10, pady=8)
        ctk.CTkLabel(lh_inner, text=self.t("layer_height_label"),
                     font=("Segoe UI", 11)).pack(side="left")
        self.layer_height_entry = ctk.CTkEntry(lh_inner, width=60, placeholder_text="0.2")
        self.layer_height_entry.insert(0, "0.2")
        self.layer_height_entry.pack(side="left", padx=(6, 0))
        ctk.CTkLabel(lh_inner, text=self.t("dithering_step"),
                     font=("Segoe UI", 9), text_color="#475569").pack(side="left", padx=6)

        # ── HAUPTBEREICH ─────────────────────────────────────────────────────
        self.main = ctk.CTkScrollableFrame(self)
        self.main.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=10)
        self.main.grid_columnconfigure(0, weight=1)

        # ── ABSCHNITT 1: EINZELFARBEN-RECHNER ────────────────────────────────
        sec1 = ctk.CTkFrame(self.main, fg_color="#0f172a", corner_radius=10)
        sec1.grid(row=0, column=0, padx=0, pady=(0, 12), sticky="ew")
        sec1.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(sec1, text=self.t("sec1_title"),
                     font=("Segoe UI", 13, "bold"), text_color="#38bdf8").grid(
            row=0, column=0, padx=20, pady=(14, 6), sticky="w")

        # Zielfarbe
        top = ctk.CTkFrame(sec1, fg_color="transparent")
        top.grid(row=1, column=0, padx=20, pady=(0, 8), sticky="ew")
        top.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(top, text=self.t("btn_target_color"), command=self.pick,
                      height=46, font=("Segoe UI", 13, "bold"), width=165).grid(
            row=0, column=0, padx=(0, 8))
        self.hex_target_entry = ctk.CTkEntry(top, placeholder_text=self.t("hex_placeholder"),
                                              font=("Courier New", 13), height=46)
        self.hex_target_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.hex_target_entry.bind("<Return>", self._on_hex_target_enter)
        self.prev = ctk.CTkLabel(top, text="", width=46, height=46,
                                  fg_color="#1a1a1a", corner_radius=23)
        self.prev.grid(row=0, column=2)

        # Gamut-Warnung
        self.gamut_label = ctk.CTkLabel(
            sec1, text=self.t("gamut_warning"),
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
        wh.grid(row=5, column=0, padx=40, pady=(0, 2), sticky="ew")
        ctk.CTkLabel(wh, text=self.t("weight_bottom"), font=("Segoe UI", 8),
                     text_color="#475569").pack(side="left")
        ctk.CTkLabel(wh, text=self.t("weight_arrow"), font=("Segoe UI", 8),
                     text_color="#334155").pack(side="left", expand=True)
        ctk.CTkLabel(wh, text=self.t("weight_top"), font=("Segoe UI", 8),
                     text_color="#475569").pack(side="right")

        # Gamut-Vorschau
        gf = ctk.CTkFrame(sec1, fg_color="#0f172a", corner_radius=6)
        gf.grid(row=5, column=0, padx=40, pady=(28, 4), sticky="ew")
        ctk.CTkLabel(gf, text="Gamut:", font=("Segoe UI", 8),
                     text_color="#475569").pack(side="left", padx=(8, 4), pady=3)
        self.gamut_strip = ctk.CTkFrame(gf, fg_color="transparent", height=16)
        self.gamut_strip.pack(side="left", fill="x", expand=True, padx=4, pady=3)
        self._gamut_cells = []
        for _ in range(40):
            c = ctk.CTkLabel(self.gamut_strip, text="", width=0, height=16,
                              fg_color="#1e293b", corner_radius=0)
            c.pack(side="left", fill="x", expand=True)
            self._gamut_cells.append(c)
        self.after(200, self._update_gamut_strip)

        # Mix-Vorschau + ΔE
        mf = ctk.CTkFrame(sec1, fg_color="#1e293b", corner_radius=8)
        mf.grid(row=6, column=0, padx=40, pady=6, sticky="ew")
        mf.grid_columnconfigure(1, weight=1)

        # Ziel-Seite
        tgt_col = ctk.CTkFrame(mf, fg_color="transparent")
        tgt_col.grid(row=0, column=0, padx=(20, 8), pady=12)
        ctk.CTkLabel(tgt_col, text=self.t("label_target"), font=("Segoe UI", 10),
                     text_color="#64748b").pack()
        self.target_circle = ctk.CTkLabel(tgt_col, text="", width=64, height=64,
                                           fg_color="#1e293b", corner_radius=32)
        self.target_circle.pack(pady=(4, 2))
        self.target_hex_lbl = ctk.CTkLabel(tgt_col, text="—", font=("Courier New", 9),
                                            text_color="#475569")
        self.target_hex_lbl.pack()

        # ΔE Mitte
        de_col = ctk.CTkFrame(mf, fg_color="transparent")
        de_col.grid(row=0, column=1, padx=20)
        self.de_disp = ctk.CTkLabel(de_col, text="ΔE  —",
                                     font=("Segoe UI", 18, "bold"), text_color="#475569")
        self.de_disp.pack()
        self.de_quality_lbl = ctk.CTkLabel(de_col, text="", font=("Segoe UI", 9),
                                            text_color="#475569")
        self.de_quality_lbl.pack()

        # Simuliert-Seite
        sim_col = ctk.CTkFrame(mf, fg_color="transparent")
        sim_col.grid(row=0, column=2, padx=(8, 20), pady=12)
        ctk.CTkLabel(sim_col, text=self.t("label_simulated"), font=("Segoe UI", 10),
                     text_color="#64748b").pack()
        self.sim_circle = ctk.CTkLabel(sim_col, text="", width=64, height=64,
                                        fg_color="#1e293b", corner_radius=32)
        self.sim_circle.pack(pady=(4, 2))
        self.sim_hex_lbl = ctk.CTkLabel(sim_col, text="—", font=("Courier New", 9),
                                         text_color="#475569")
        self.sim_hex_lbl.pack()

        # Sequenzlänge + Auto-Modus
        lr = ctk.CTkFrame(sec1, fg_color="#1a2535", corner_radius=8)
        lr.grid(row=7, column=0, padx=40, pady=(4, 4), sticky="ew")
        lr.grid_columnconfigure(1, weight=1)

        self.len_label = ctk.CTkLabel(lr, text=self.t("length_label", n=10),
                                       font=("Segoe UI", 11, "bold"), width=80)
        self.len_label.grid(row=0, column=0, padx=(14, 6), pady=10)

        self.len_slider = ctk.CTkSlider(lr, from_=1, to=10, number_of_steps=9,
                                         command=self._on_len_slider)
        self.len_slider.set(10)
        self.len_slider.grid(row=0, column=1, sticky="ew", padx=6)

        self.auto_len_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(lr, text=self.t("auto_check"), variable=self.auto_len_var,
                        font=("Segoe UI", 9),
                        command=self._on_auto_toggle).grid(row=0, column=2, padx=(8, 4))

        ctk.CTkLabel(lr, text="ΔE≤", font=("Segoe UI", 10),
                     text_color="#64748b").grid(row=0, column=3, padx=(4, 0))
        self.auto_thresh_entry = ctk.CTkEntry(lr, width=46, placeholder_text="2.0")
        self.auto_thresh_entry.insert(0, "2.0")
        self.auto_thresh_entry.grid(row=0, column=4, padx=(2, 14))

        # Zielfarbe: Random + Bild-Picker Buttons
        rnd_btn = ctk.CTkButton(top, text=self.t("btn_random"), width=85, height=46,
                      fg_color="#374151", font=("Segoe UI", 11),
                      command=self.pick_random_color)
        rnd_btn.grid(row=0, column=3, padx=(4, 0))
        self.tip(rnd_btn, "tip_random")

        img_btn = ctk.CTkButton(top, text=self.t("btn_img_pick"), width=100, height=46,
                      fg_color="#374151", font=("Segoe UI", 11),
                      command=self.pick_color_from_image)
        img_btn.grid(row=0, column=4, padx=(4, 0))
        self.tip(img_btn, "tip_img_pick")

        # Farbinfo-Label (RGB + Lab)
        self.colorinfo_label = ctk.CTkLabel(sec1, text="",
                                             font=("Segoe UI", 9), text_color="#475569")
        self.colorinfo_label.grid(row=2, column=0, padx=20, pady=(0, 2), sticky="w")

        # Steuerleiste
        bl = ctk.CTkFrame(sec1, fg_color="transparent")
        bl.grid(row=8, column=0, padx=40, pady=(4, 16), sticky="ew")
        bl.grid_columnconfigure(0, weight=1)
        calc_btn = ctk.CTkButton(bl, text=self.t("btn_calculate"), fg_color="#2563eb",
                      command=self.calc, height=46,
                      font=("Segoe UI", 14, "bold"))
        calc_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.tip(calc_btn, "tip_calculate")

        self.optimizer_var = ctk.BooleanVar(value=False)
        opt_cb = ctk.CTkCheckBox(bl, text=self.t("optimizer_check"), variable=self.optimizer_var,
                        font=("Segoe UI", 9))
        opt_cb.grid(row=0, column=1, padx=(0, 6))
        self.tip(opt_cb, "tip_optimizer")

        add_btn = ctk.CTkButton(bl, text=self.t("btn_add_virtual"),
                      fg_color="#15803d", hover_color="#166534",
                      command=self.add_to_virtual, height=46,
                      font=("Segoe UI", 11, "bold"))
        add_btn.grid(row=0, column=2, padx=(0, 6))
        self.tip(add_btn, "tip_add_virtual")

        copy_btn = ctk.CTkButton(bl, text=self.t("btn_copy"), fg_color="#374151",
                      command=self.copy_sequence, height=46, width=110,
                      font=("Segoe UI", 11))
        copy_btn.grid(row=0, column=3, padx=(0, 6))
        self.tip(copy_btn, "tip_copy")

        exp_btn = ctk.CTkButton(bl, text="EXPORT", fg_color="#7c3aed",
                      command=self.open_export_dialog, height=46, width=90,
                      font=("Segoe UI", 12, "bold"))
        exp_btn.grid(row=0, column=4)
        self.tip(exp_btn, "tip_export")

        # Keyboard shortcut: Enter = Berechnen
        self.bind("<Return>", lambda e: self.calc())

        # ── Dithering-Profile ─────────────────────────────────────────────────
        dp_frame = ctk.CTkFrame(sec1, fg_color="#0f172a", corner_radius=8)
        dp_frame.grid(row=9, column=0, padx=40, pady=(0, 6), sticky="ew")
        ctk.CTkLabel(dp_frame, text=self.t("dither_profiles"),
                     font=("Segoe UI", 10), text_color="#64748b").pack(side="left", padx=(12, 8))
        for key, fn in [("dither_fine", lambda: self._apply_dither_profile(2, False)),
                        ("dither_balanced", lambda: self._apply_dither_profile(5, False)),
                        ("dither_smooth", lambda: self._apply_dither_profile(9, False)),
                        ("dither_auto", lambda: self._apply_dither_profile(None, True))]:
            b = ctk.CTkButton(dp_frame, text=self.t(key), height=28, width=100,
                              fg_color="#1e3a5f", font=("Segoe UI", 10),
                              command=fn)
            b.pack(side="left", padx=3, pady=6)
            self.tip(b, "tip_" + key)

        # ── Farbsehschwäche-Simulation + Extras ──────────────────────────────
        sim_frame = ctk.CTkFrame(sec1, fg_color="#0f172a", corner_radius=8)
        sim_frame.grid(row=10, column=0, padx=40, pady=(0, 12), sticky="ew")
        ctk.CTkLabel(sim_frame, text=self.t("colorblind_label"),
                     font=("Segoe UI", 10), text_color="#64748b").pack(side="left", padx=(12, 6))
        self.colorblind_var = ctk.StringVar(value="normal")
        for val, key in [("normal", "colorblind_normal"), ("prot", "colorblind_prot"),
                         ("deut", "colorblind_deut"), ("trit", "colorblind_trit")]:
            ctk.CTkRadioButton(sim_frame, text=self.t(key), variable=self.colorblind_var,
                               value=val, font=("Segoe UI", 10),
                               command=self._update_colorblind_preview).pack(side="left", padx=4, pady=6)

        # Extra-Buttons (Lab-Plot, Swatch, Slicer-Guide)
        extra_row = ctk.CTkFrame(sec1, fg_color="transparent")
        extra_row.grid(row=11, column=0, padx=40, pady=(0, 14), sticky="ew")
        lab_btn = ctk.CTkButton(extra_row, text=self.t("btn_lab_plot"), fg_color="#0f4c81",
                                height=32, font=("Segoe UI", 11), command=self.show_lab_plot)
        lab_btn.pack(side="left", padx=(0, 6))
        self.tip(lab_btn, "tip_lab_plot")
        sw_btn = ctk.CTkButton(extra_row, text=self.t("btn_swatch"), fg_color="#374151",
                               height=32, font=("Segoe UI", 11), command=self.save_swatch)
        sw_btn.pack(side="left", padx=(0, 6))
        self.tip(sw_btn, "tip_swatch")
        guide_btn = ctk.CTkButton(extra_row, text=self.t("btn_slicer_guide"), fg_color="#7c3aed",
                                  height=32, font=("Segoe UI", 11), command=self.open_slicer_guide)
        guide_btn.pack(side="left", padx=(0, 6))
        grad_btn = ctk.CTkButton(extra_row, text=self.t("btn_gradient"), fg_color="#0e7490",
                                 height=32, font=("Segoe UI", 11),
                                 command=self.open_gradient_dialog)
        grad_btn.pack(side="left")
        self.tip(grad_btn, "tip_gradient")

        # ── ABSCHNITT 2: VIRTUELLE DRUCKKÖPFE ────────────────────────────────
        sec2_hdr = ctk.CTkFrame(self.main, fg_color="transparent")
        sec2_hdr.grid(row=1, column=0, padx=0, pady=(0, 6), sticky="ew")
        sec2_hdr.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(sec2_hdr,
                     text=self.t("sec2_title"),
                     font=("Segoe UI", 15, "bold"), text_color="#a78bfa").grid(
            row=0, column=0, sticky="w")
        ctk.CTkLabel(sec2_hdr,
                     text=self.t("sec2_desc", max_v=self._max_virtual),
                     font=("Segoe UI", 10), text_color="#64748b").grid(
            row=1, column=0, sticky="w")

        btn_row2 = ctk.CTkFrame(sec2_hdr, fg_color="transparent")
        btn_row2.grid(row=0, column=1, sticky="e", padx=(10, 0))
        tmf_btn = ctk.CTkButton(btn_row2, text=self.t("btn_3mf"), fg_color="#0f4c81",
                      hover_color="#1e3a5f", height=40, width=160,
                      font=("Segoe UI", 12, "bold"),
                      command=self.open_3mf_assistant)
        tmf_btn.pack(side="left", padx=(0, 6))
        self.tip(tmf_btn, "tip_3mf")

        bat_btn = ctk.CTkButton(btn_row2, text=self.t("btn_batch"), fg_color="#4338ca",
                      hover_color="#3730a3", height=40, width=140,
                      font=("Segoe UI", 11, "bold"),
                      command=self.open_batch_dialog)
        bat_btn.pack(side="left", padx=(0, 6))
        self.tip(bat_btn, "tip_batch")
        undo_btn = ctk.CTkButton(btn_row2, text=self.t("btn_undo"), fg_color="#374151",
                      height=40, width=110, command=self.undo_virtual)
        undo_btn.pack(side="left", padx=(0, 6))
        self.tip(undo_btn, "tip_calculate")  # reuse

        ctk.CTkButton(btn_row2, text=self.t("btn_del_all"), fg_color="#7f1d1d",
                      hover_color="#991b1b", height=40, width=120,
                      command=self.clear_virtual).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btn_row2, text=self.t("btn_export_all"), fg_color="#374151",
                      height=40, width=140,
                      command=self.open_export_dialog).pack(side="left", padx=(0, 6))
        orca_btn = ctk.CTkButton(btn_row2, text=self.t("btn_orca_export"), fg_color="#0f766e",
                      hover_color="#0d6660", height=40, width=140,
                      font=("Segoe UI", 11, "bold"),
                      command=self.open_orca_export_dialog)
        orca_btn.pack(side="left")
        self.tip(orca_btn, "tip_orca_export")

        # Virtual Filament Grid Header
        gh = ctk.CTkFrame(self.main, fg_color="#1e293b", corner_radius=6)
        gh.grid(row=2, column=0, sticky="ew", pady=(0, 2))
        gh.grid_columnconfigure(4, weight=1)
        for col, (txt, w) in enumerate([("ID", 55), (self.t("grid_target"), 50),
                                         (self.t("grid_sequence"), 160), (self.t("grid_simulated"), 60),
                                         (self.t("grid_quality"), 130), (self.t("grid_label"), 0),
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
        cols = [f["name"] for f in self.library.get(b, [])] or [self.t("empty_slot")]
        self.slots[idx]["color"].configure(values=cols)
        self.slots[idx]["color"].set(cols[0])
        self.apply_f(idx)

    def apply_f(self, idx):
        b = self.slots[idx]["brand"].get()
        n = self.slots[idx]["color"].get()
        if n in _SLOT_SKIP: return
        f = next((x for x in self.library.get(b, []) if x["name"] == n), None)
        if f is None: return
        self.slots[idx]["hex"].delete(0, "end"); self.slots[idx]["hex"].insert(0, f["hex"])
        self.slots[idx]["td"].delete(0, "end");  self.slots[idx]["td"].insert(0, str(f["td"]))
        self.slots[idx]["preview"].configure(fg_color=f["hex"])
        self.after(100, self._update_gamut_strip)

    def pick_slot_color(self, idx):
        cur = self.slots[idx]["hex"].get().strip() or "#808080"
        h = self._ask_color(initial=cur, title=self.t("color_picker_title", i=idx+1))
        if h is None: return
        self.slots[idx]["hex"].delete(0, "end"); self.slots[idx]["hex"].insert(0, h)
        self.slots[idx]["preview"].configure(fg_color=h)
        cur_vals = [f["name"] for f in self.library.get(self.slots[idx]["brand"].get(), [])]
        manual = self.t("manual_color")
        self.slots[idx]["color"].configure(values=[manual] + cur_vals)
        self.slots[idx]["color"].set(manual)
        self.after(100, self._update_gamut_strip)

    def save_current(self, idx):
        n = ctk.CTkInputDialog(text=self.t("inp_fil_name"), title=self.t("inp_save_fav")).get_input()
        if not n: return
        self.library.setdefault("Eigene Favoriten", []).append(
            {"name": n.strip(), "hex": self.slots[idx]["hex"].get(),
             "td": safe_td(self.slots[idx]["td"].get())})
        self.save_db(); self._refresh_brand_menus()
        messagebox.showinfo(self.t("dlg_saved"), self.t("dlg_fil_saved", n=n))

    def add_filament(self, idx):
        b = self.slots[idx]["brand"].get()
        n = ctk.CTkInputDialog(text=self.t("inp_name"), title=self.t("inp_add_fil_title", b=b)).get_input()
        if not n: return
        h = ctk.CTkInputDialog(text=self.t("inp_hex"), title=self.t("inp_color_title")).get_input()
        if not h: return
        h = h.strip()
        if not h.startswith("#"): h = "#" + h
        td_raw = ctk.CTkInputDialog(text=self.t("inp_td", td=DEFAULT_TD), title=self.t("inp_td_title")).get_input()
        td_val = safe_td(td_raw) if td_raw else DEFAULT_TD
        self.library.setdefault(b, []).append({"name": n.strip(), "hex": h, "td": td_val})
        self.save_db(); self.update_menu(idx)

    def add_brand(self):
        name = ctk.CTkInputDialog(text=self.t("inp_brand_name"), title=self.t("inp_brand_title")).get_input()
        if not name: return
        name = name.strip()
        if name in self.library:
            messagebox.showwarning(self.t("dlg_exists"), self.t("dlg_exists_msg", name=name)); return
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

    def _ask_color(self, initial="#808080", title="Pick Color"):
        """Unified color picker: CTkColorPicker wenn verfügbar, sonst tkinter Fallback."""
        if _HAS_CTKPICKER:
            picker = _AskColor(width=350, initial_color=initial,
                               title=title, text="OK")
            result = picker.get()
            return result.upper() if result else None
        else:
            result = colorchooser.askcolor(color=initial, title=title)
            return result[1].upper() if result[1] else None

    def pick(self):
        c = self._ask_color(title=self.t("target_picker_title"))
        if c: self._apply_target(c)

    def _on_hex_target_enter(self, event=None):
        raw = self.hex_target_entry.get().strip()
        if not raw.startswith("#"): raw = "#" + raw
        if len(raw) == 7: self._apply_target(raw.upper())

    @staticmethod
    def _rgb_to_hsv(r, g, b):
        r_, g_, b_ = r/255, g/255, b/255
        mx, mn = max(r_, g_, b_), min(r_, g_, b_)
        df = mx - mn
        if df == 0: h = 0
        elif mx == r_: h = (60 * ((g_ - b_) / df) + 360) % 360
        elif mx == g_: h = (60 * ((b_ - r_) / df) + 120) % 360
        else:          h = (60 * ((r_ - g_) / df) + 240) % 360
        s = 0 if mx == 0 else (df / mx * 100)
        v = mx * 100
        return h, s, v

    def _apply_target(self, hex_str):
        self.target = hex_str
        self.prev.configure(fg_color=hex_str)
        self.target_circle.configure(fg_color=hex_str)
        self.hex_target_entry.delete(0, "end")
        self.hex_target_entry.insert(0, hex_str)
        rgb = hex_to_rgb(hex_str)
        lab = rgb_to_lab(rgb)
        h, s, v = self._rgb_to_hsv(*rgb)
        info = self.t("colorinfo_label", r=rgb[0], g=rgb[1], b=rgb[2],
                      h=h, s=s, v=v, L=lab[0], a=lab[1], b_=lab[2])
        if hasattr(self, "colorinfo_label"):
            self.colorinfo_label.configure(text=info)
        self.calc()

    def _on_len_slider(self, val):
        n = int(val)
        self.len_label.configure(text=self.t("length_label", n=n))
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

    def _update_gamut_strip(self):
        """Füllt die Gamut-Strip-Zellen mit erreichbaren Mischfarben."""
        if not hasattr(self, "_gamut_cells"):
            return
        fils = self._get_fils()
        if not fils or len(fils) < 2:
            return
        import itertools
        samples = []
        for f in fils:
            samples.append(f["hex"])
        for f1, f2 in itertools.combinations(fils, 2):
            for w in [0.15, 0.3, 0.45, 0.55, 0.7, 0.85]:
                lab1, lab2 = f1["lab"], f2["lab"]
                mix = tuple(lab1[k] * (1-w) + lab2[k] * w for k in range(3))
                samples.append(lab_to_hex(mix))
        if len(fils) >= 3:
            for combo in itertools.combinations(fils, 3):
                for w in [(0.33,0.33,0.34),(0.5,0.25,0.25),(0.25,0.5,0.25),(0.25,0.25,0.5)]:
                    mix = tuple(sum(combo[j]["lab"][k]*w[j] for j in range(3))
                                 for k in range(3))
                    samples.append(lab_to_hex(mix))
        if len(fils) >= 4:
            mix4 = tuple(sum(f["lab"][k] for f in fils)/4 for k in range(3))
            samples.append(lab_to_hex(mix4))

        def hue_key(h):
            r, g, b = hex_to_rgb(h)
            hue, s, v = self._rgb_to_hsv(r, g, b)
            return (0 if s < 10 else 1, hue)
        samples.sort(key=hue_key)

        n = len(self._gamut_cells)
        for i, cell in enumerate(self._gamut_cells):
            idx = int(i * len(samples) / n)
            cell.configure(fg_color=samples[min(idx, len(samples)-1)])

    def calc(self):
        if not hasattr(self, "target"):
            messagebox.showinfo(self.t("dlg_note"), self.t("dlg_select_color")); return

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
            self.len_label.configure(text=self.t("length_label", n=n))

        self.res.configure(text=seq)

        # Segmente dynamisch aktualisieren
        fils_hex = {f["id"]: f["hex"] for f in fils}
        for i, seg in enumerate(self.segs):
            if i < n:
                fid = int(seq[i])
                bg  = fils_hex.get(fid, "#1e293b")
                r_, g_, b_ = hex_to_rgb(bg)
                lum = 0.299*r_ + 0.587*g_ + 0.114*b_
                txt_col = "#111111" if lum > 140 else "#eeeeee"
                seg.configure(fg_color=bg, text=f"T{fid}", text_color=txt_col,
                               font=("Segoe UI", 10, "bold"))
                seg.pack(side="left", expand=True, padx=2)
            else:
                seg.pack_forget()

        self._last_sim_hex = result["sim_hex"]
        self._last_de      = result["de"]
        dv = result["de"]

        # Farbsehschwäche-Simulation anwenden
        mode = self.colorblind_var.get() if hasattr(self, "colorblind_var") else "normal"
        sim_display = self._simulate_colorblind(result["sim_hex"], mode)
        tgt_display = self._simulate_colorblind(self.target, mode)
        self.sim_circle.configure(fg_color=sim_display)
        self.target_circle.configure(fg_color=tgt_display)
        self.de_disp.configure(text=f"ΔE  {dv:.1f}", text_color=de_color(dv))
        if hasattr(self, "target_hex_lbl"):
            self.target_hex_lbl.configure(text=self.target.upper())
        if hasattr(self, "sim_hex_lbl"):
            self.sim_hex_lbl.configure(text=result["sim_hex"].upper())
        if hasattr(self, "de_quality_lbl"):
            quality = "ausgezeichnet ✓" if dv < 3.0 else "gut" if dv < 6.0 else "sichtbar"
            if self.lang == "en":
                quality = "excellent ✓" if dv < 3.0 else "good" if dv < 6.0 else "visible"
            self.de_quality_lbl.configure(text=quality, text_color=de_color(dv))

        self.last_result = result

    # ── VIRTUELLE DRUCKKÖPFE ───────────────────────────────────────────────────

    def add_to_virtual(self, result=None):
        if len(self.virtual_fils) >= self._max_virtual:
            messagebox.showwarning(self.t("dlg_max_virtual"), self.t("dlg_max_virtual_msg", max_v=self._max_virtual))
            return
        if result is None:
            result = self.last_result
        if not result:
            messagebox.showinfo(self.t("dlg_note"), self.t("dlg_no_seq")); return
        self.virtual_undo.append(copy.deepcopy(self.virtual_fils))
        vid = 5 + len(self.virtual_fils)
        self.virtual_fils.append({
            "vid":        vid,
            "target_hex": result["target_hex"],
            "sequence":   result["sequence"],
            "sim_hex":    result["sim_hex"],
            "de":         result["de"],
            "label":      self.t("virtual_label_default", vid=vid),
        })
        self._refresh_virtual_grid()

    def clear_virtual(self):
        if self.virtual_fils and messagebox.askyesno(
                self.t("dlg_del_title"), self.t("dlg_del_virtual")):
            self.virtual_undo.append(copy.deepcopy(self.virtual_fils))
            self.virtual_fils.clear()
            self._refresh_virtual_grid()

    def undo_virtual(self):
        if not self.virtual_undo:
            messagebox.showinfo(self.t("dlg_note"), self.t("undo_empty")); return
        self.virtual_fils = self.virtual_undo.pop()
        self._refresh_virtual_grid()

    def pick_color_from_image(self):
        """Farbe aus Bilddatei wählen (Pillow)."""
        if not _HAS_PIL:
            messagebox.showinfo(self.t("dlg_note"), "Pillow not installed."); return
        path = filedialog.askopenfilename(
            filetypes=[(self.t("img_filetypes"), "*.png *.jpg *.jpeg *.bmp *.webp *.gif"),
                       ("*", "*.*")],
            title=self.t("img_pick_title"))
        if not path: return

        img = _PILImage.open(path).convert("RGB")
        # Bild auf max 600px skalieren für schnelles Anzeigen
        max_size = 600
        w, h = img.size
        scale = min(max_size / w, max_size / h, 1.0)
        disp_w, disp_h = int(w * scale), int(h * scale)
        img_disp = img.resize((disp_w, disp_h), _PILImage.LANCZOS)

        win = ctk.CTkToplevel(self)
        win.title(self.t("img_pick_title"))
        win.geometry(f"{disp_w + 40}x{disp_h + 100}")
        win.grab_set()
        win.resizable(False, False)

        ctk.CTkLabel(win, text=self.t("img_instruction"),
                     font=("Segoe UI", 11), text_color="#94a3b8").pack(pady=(10, 4))

        import tkinter as tk
        from PIL import ImageTk
        tk_img = ImageTk.PhotoImage(img_disp)

        canvas = tk.Canvas(win, width=disp_w, height=disp_h,
                           highlightthickness=0, bg="#0f172a", cursor="crosshair")
        canvas.pack(padx=20)
        canvas.create_image(0, 0, anchor="nw", image=tk_img)
        canvas._img = tk_img  # reference to avoid GC

        preview_frame = ctk.CTkFrame(win, fg_color="transparent")
        preview_frame.pack(pady=6)
        preview_dot = ctk.CTkLabel(preview_frame, text="", width=32, height=32,
                                   fg_color="#808080", corner_radius=16)
        preview_dot.pack(side="left", padx=6)
        preview_hex = ctk.CTkLabel(preview_frame, text="—",
                                   font=("Courier New", 13), text_color="#94a3b8")
        preview_hex.pack(side="left")

        chosen = [None]

        def on_move(event):
            px = int(event.x / scale)
            py = int(event.y / scale)
            px = max(0, min(px, w - 1))
            py = max(0, min(py, h - 1))
            r, g, b = img.getpixel((px, py))
            h_str = f"#{r:02X}{g:02X}{b:02X}"
            preview_dot.configure(fg_color=h_str)
            preview_hex.configure(text=h_str)
            chosen[0] = h_str

        def on_click(event):
            if chosen[0]:
                win.destroy()
                self._apply_target(chosen[0])

        canvas.bind("<Motion>", on_move)
        canvas.bind("<Button-1>", on_click)
        win.mainloop()

    # ── DITHERING-PROFILE ──────────────────────────────────────────────────────

    def _apply_dither_profile(self, length, auto):
        """Schnellprofil: Slider + Auto-Checkbox setzen und neu berechnen."""
        self.auto_len_var.set(auto)
        self.len_slider.configure(state="normal")
        if auto:
            self.len_slider.configure(state="disabled")
        elif length is not None:
            self.len_slider.set(length)
            self.len_label.configure(text=self.t("length_label", n=length))
            for i, seg in enumerate(self.segs):
                if i < length: seg.pack(side="left", expand=True, padx=2)
                else: seg.pack_forget()
        if hasattr(self, "target"):
            self.calc()

    # ── FARBSEHSCHWÄCHE-SIMULATION ─────────────────────────────────────────────

    @staticmethod
    def _simulate_colorblind(hex_str, mode):
        """Approximation einer Farbsehschwäche via Transformationsmatrix."""
        r, g, b = [x/255 for x in hex_to_rgb(hex_str)]
        if mode == "prot":   # Protanopie (Rot-Blind)
            r2 = 0.567*r + 0.433*g
            g2 = 0.558*r + 0.442*g
            b2 = 0.242*g + 0.758*b
        elif mode == "deut": # Deuteranopie (Grün-Blind)
            r2 = 0.625*r + 0.375*g
            g2 = 0.700*r + 0.300*g
            b2 = 0.300*g + 0.700*b
        elif mode == "trit": # Tritanopie (Blau-Blind)
            r2 = 0.950*r + 0.050*b
            g2 = 0.433*g + 0.567*b
            b2 = 0.475*g + 0.525*b
        else:
            return hex_str
        ri, gi, bi = (int(max(0,min(1,v))*255) for v in (r2, g2, b2))
        return f"#{ri:02X}{gi:02X}{bi:02X}"

    def _update_colorblind_preview(self):
        """Sim-Kreis und Ziel-Kreis mit Farbsehschwäche-Transformation aktualisieren."""
        mode = self.colorblind_var.get() if hasattr(self, "colorblind_var") else "normal"
        if not hasattr(self, "target"): return
        tgt_c = self._simulate_colorblind(self.target, mode)
        self.target_circle.configure(fg_color=tgt_c)
        if hasattr(self, "_last_sim_hex") and self._last_sim_hex:
            sim_c = self._simulate_colorblind(self._last_sim_hex, mode)
            self.sim_circle.configure(fg_color=sim_c)

    # ── 3D LAB-VISUALISIERUNG ─────────────────────────────────────────────────

    def show_lab_plot(self):
        """Matplotlib 3D-Plot: T1-T4 und Zielfarbe im CIE-Lab-Farbraum."""
        if not _HAS_MPL:
            messagebox.showinfo(self.t("dlg_note"), "matplotlib not installed."); return
        fils = self._get_fils()
        fig = _plt.figure(figsize=(8, 6))
        ax = fig.add_subplot(111, projection='3d')
        ax.set_facecolor("#0f172a")
        fig.patch.set_facecolor("#1e293b")
        for spine in ax.spines.values():
            spine.set_color("#475569")
        ax.tick_params(colors="#94a3b8")
        ax.xaxis.label.set_color("#94a3b8")
        ax.yaxis.label.set_color("#94a3b8")
        ax.zaxis.label.set_color("#94a3b8")
        for f in fils:
            L, a, b = f["lab"]
            ax.scatter(a, b, L, c=f["hex"], s=180, edgecolors="#ffffff",
                       linewidths=1.5, zorder=5)
            ax.text(a, b, L + 2, f"T{f['id']}", color="#ffffff", fontsize=9)
        if hasattr(self, "target"):
            tlab = rgb_to_lab(hex_to_rgb(self.target))
            L, a, b = tlab
            ax.scatter(a, b, L, c=self.target, s=260, marker='*',
                       edgecolors="#ffffff", linewidths=1.5, zorder=6)
            ax.text(a, b, L + 2, "⊙ Ziel", color="#fbbf24", fontsize=9)
            if hasattr(self, "_last_sim_hex") and self._last_sim_hex:
                slab = rgb_to_lab(hex_to_rgb(self._last_sim_hex))
                ax.scatter(slab[1], slab[2], slab[0], c=self._last_sim_hex,
                           s=180, marker='^', edgecolors="#ffffff", linewidths=1.5, zorder=6)
                ax.text(slab[1], slab[2], slab[0] + 2, "≈ Sim", color="#4ade80", fontsize=9)
        ax.set_xlabel("a* (grün–rot)")
        ax.set_ylabel("b* (blau–gelb)")
        ax.set_zlabel("L* (Helligkeit)")
        ax.set_title(self.t("lab_plot_title"), color="#e2e8f0", pad=10)
        _plt.tight_layout()
        _plt.show()

    # ── SWATCH SPEICHERN ──────────────────────────────────────────────────────

    def save_swatch(self):
        """Farbvergleich (Ziel vs. Simuliert) als PNG speichern."""
        if not _HAS_PIL:
            messagebox.showinfo(self.t("dlg_note"), "Pillow not installed."); return
        if not hasattr(self, "target"):
            messagebox.showinfo(self.t("dlg_note"), self.t("dlg_select_color")); return
        path = filedialog.asksaveasfilename(
            defaultextension=".png", filetypes=[("PNG", "*.png")],
            title=self.t("btn_swatch"))
        if not path: return
        W, H, PAD = 500, 120, 16
        img = _PILImage.new("RGB", (W, H), (15, 23, 42))
        draw = _ImageDraw.Draw(img)
        # Ziel-Quadrat
        tr, tg, tb = hex_to_rgb(self.target)
        draw.rectangle([PAD, PAD, W//2 - PAD//2, H - PAD], fill=(tr, tg, tb))
        # Simuliert-Quadrat
        sim_h = getattr(self, "_last_sim_hex", "#808080")
        sr, sg, sb = hex_to_rgb(sim_h)
        draw.rectangle([W//2 + PAD//2, PAD, W - PAD, H - PAD], fill=(sr, sg, sb))
        # Labels
        draw.text((PAD + 4, PAD + 4), f"Target\n{self.target}", fill=(255,255,255))
        draw.text((W//2 + PAD//2 + 4, PAD + 4), f"Simulated\n{sim_h}", fill=(255,255,255))
        if hasattr(self, "de_label") and hasattr(self, "_last_de"):
            draw.text((W//2 - 30, H//2 - 8), f"ΔE {self._last_de:.1f}", fill=(255,200,100))
        img.save(path)
        messagebox.showinfo(self.t("dlg_saved"), self.t("swatch_saved", path=path))

    # ── SLICER-GUIDE ──────────────────────────────────────────────────────────

    def open_slicer_guide(self):
        """Schritt-für-Schritt Anleitung: Ergebnis in OrcaSlicer FullSpectrum einbinden."""
        win = ctk.CTkToplevel(self)
        win.title(self.t("slicer_guide_title"))
        win.geometry("720x640")
        win.grab_set()

        scroll = ctk.CTkScrollableFrame(win)
        scroll.pack(fill="both", expand=True, padx=16, pady=16)

        def section(title, color="#38bdf8"):
            ctk.CTkLabel(scroll, text=title, font=("Segoe UI", 13, "bold"),
                         text_color=color).pack(anchor="w", pady=(14, 2))

        def para(text, color="#cbd5e1"):
            ctk.CTkLabel(scroll, text=text, font=("Segoe UI", 11),
                         text_color=color, wraplength=660, justify="left").pack(anchor="w", pady=2)

        def code(text):
            ctk.CTkLabel(scroll, text=text, font=("Courier New", 11),
                         fg_color="#0f172a", corner_radius=6,
                         text_color="#4ade80").pack(anchor="w", fill="x", padx=4, pady=3)

        if self.lang == "de":
            section("① Voraussetzung: OrcaSlicer FullSpectrum")
            para("Installiere den OrcaSlicer-FullSpectrum Fork (nicht den Standard-OrcaSlicer):")
            code("  github.com/ratdoux/OrcaSlicer-FullSpectrum")
            para("Die FullSpectrum-Funktion befindet sich unter:  Filament → Others → Dithering")

            section("② Filamente im Slicer konfigurieren")
            para("Lade die 4 physischen Filamente T1–T4 in OrcaSlicer als separate Filament-Profile.\n"
                 "Stelle sicher, dass Farbe und TD-Wert mit deinen Slot-Einstellungen hier übereinstimmen.")

            section("③ Virtuellen Druckkopf im Slicer anlegen", "#a78bfa")
            para("Für jeden virtuellen Druckkopf (V5, V6, ...) musst du ein neues Filament-Profil anlegen\n"
                 "und die Dithering-Einstellungen setzen:")

            section("  2-Filament-Sequenz → Cadence Height", "#4ade80")
            para("Beispiel: Sequenz  1121  →  T1×2, T2×1, T1×1")
            code("  Others → Dithering → Enable Dithering: ✓")
            code("  Dithering Step Size: 0.2 mm  (= deine Schichthöhe)")
            code("  Cadence Height A: aus dem Export  (z.B. 0.4 mm)")
            code("  Cadence Height B: aus dem Export  (z.B. 0.2 mm)")

            section("  3–4-Filament-Sequenz → Pattern Mode", "#a78bfa")
            para("Beispiel: Sequenz  11213  →  Pattern String  1/1/2/1/3")
            code("  Others → Dithering → Enable Dithering: ✓")
            code("  Dithering Step Size: 0.2 mm")
            code("  Pattern: 1/1/2/1/3  ← aus dem Export kopieren")

            section("④ Dithering Step Size", "#fbbf24")
            para("Der Dithering Step Size ist GLEICH der Schichthöhe (z.B. 0.2 mm).\n"
                 "Dies ist NICHT der Düsendurchmesser (0.4 oder 0.8 mm)!")

            section("⑤ Objekt einfärben")
            para("Wähle im Slicer das Objekt aus und weise Flächen/Regionen\n"
                 "dem virtuellen Filament V5, V6, ... zu (Paint-Brush oder per Filament-Zuweisung).")

            section("⑥ Tipps für beste Ergebnisse", "#fbbf24")
            para("• ΔE < 3: Ausgezeichnete Übereinstimmung — direkt drucken\n"
                 "• ΔE 3–6: Gute Annäherung — leichte Farbabweichung sichtbar\n"
                 "• ΔE > 6: Starke Abweichung — Sequenz optimieren oder andere Filamente wählen\n"
                 "• Für glatte Farbverläufe: Sanft-Profil (8-10 Layer)\n"
                 "• Für präzise Farben: Fein-Profil (2-3 Layer)\n"
                 "• Immer zuerst einen kleinen Testdruck machen!")
        else:
            section("① Prerequisite: OrcaSlicer FullSpectrum")
            para("Install the OrcaSlicer-FullSpectrum fork (not standard OrcaSlicer):")
            code("  github.com/ratdoux/OrcaSlicer-FullSpectrum")
            para("The FullSpectrum feature is at:  Filament → Others → Dithering")

            section("② Configure Filaments in Slicer")
            para("Load the 4 physical filaments T1–T4 in OrcaSlicer as separate filament profiles.\n"
                 "Make sure color and TD values match your slot settings here.")

            section("③ Create Virtual Print Head in Slicer", "#a78bfa")
            para("For each virtual head (V5, V6, ...) create a new filament profile\n"
                 "and set the dithering parameters:")

            section("  2-Filament Sequence → Cadence Height", "#4ade80")
            para("Example: Sequence  1121  →  T1×2, T2×1, T1×1")
            code("  Others → Dithering → Enable Dithering: ✓")
            code("  Dithering Step Size: 0.2 mm  (= your layer height)")
            code("  Cadence Height A: from export  (e.g. 0.4 mm)")
            code("  Cadence Height B: from export  (e.g. 0.2 mm)")

            section("  3–4-Filament Sequence → Pattern Mode", "#a78bfa")
            para("Example: Sequence  11213  →  Pattern String  1/1/2/1/3")
            code("  Others → Dithering → Enable Dithering: ✓")
            code("  Dithering Step Size: 0.2 mm")
            code("  Pattern: 1/1/2/1/3  ← copy from export")

            section("④ Dithering Step Size", "#fbbf24")
            para("The Dithering Step Size EQUALS the layer height (e.g. 0.2 mm).\n"
                 "This is NOT the nozzle diameter (0.4 or 0.8 mm)!")

            section("⑤ Paint the Object")
            para("Select the object in the slicer and assign surfaces/regions\n"
                 "to virtual filament V5, V6, ... (paint brush or filament assignment).")

            section("⑥ Tips for Best Results", "#fbbf24")
            para("• ΔE < 3: Excellent match — print directly\n"
                 "• ΔE 3–6: Good approximation — slight color deviation visible\n"
                 "• ΔE > 6: Large deviation — optimize sequence or choose other filaments\n"
                 "• For smooth gradients: Smooth profile (8-10 layers)\n"
                 "• For precise colors: Fine profile (2-3 layers)\n"
                 "• Always do a small test print first!")

        ctk.CTkButton(win, text="OK", fg_color="#2563eb",
                      command=win.destroy, height=36).pack(pady=(0, 14))

    # ── EINSTELLUNGEN-DIALOG ──────────────────────────────────────────────────

    def open_settings_dialog(self):
        win = ctk.CTkToplevel(self)
        win.title(self.t("settings_title"))
        win.geometry("400x300")
        win.grab_set()

        ctk.CTkLabel(win, text=self.t("settings_title"),
                     font=("Segoe UI", 14, "bold")).pack(pady=(18, 12))

        # Max virtuelle Druckköpfe
        mv_frame = ctk.CTkFrame(win, fg_color="transparent")
        mv_frame.pack(fill="x", padx=24, pady=6)
        ctk.CTkLabel(mv_frame, text=self.t("settings_max_virtual"),
                     font=("Segoe UI", 11)).pack(side="left")
        mv_var = ctk.IntVar(value=self._max_virtual)
        mv_spin = ctk.CTkEntry(mv_frame, width=60, textvariable=mv_var)
        mv_spin.pack(side="left", padx=(8, 4))
        ctk.CTkLabel(mv_frame, text=f"(1–{MAX_VIRTUAL_HARD})",
                     font=("Segoe UI", 9), text_color="#475569").pack(side="left")

        # Erscheinungsbild
        th_frame = ctk.CTkFrame(win, fg_color="transparent")
        th_frame.pack(fill="x", padx=24, pady=6)
        ctk.CTkLabel(th_frame, text=self.t("settings_theme"),
                     font=("Segoe UI", 11)).pack(side="left")
        theme_var = ctk.StringVar(value=self.settings.get("theme", "dark"))
        for val, key in [("dark", "settings_theme_dark"), ("light", "settings_theme_light"),
                         ("system", "settings_theme_system")]:
            ctk.CTkRadioButton(th_frame, text=self.t(key), variable=theme_var,
                               value=val, font=("Segoe UI", 10)).pack(side="left", padx=6)

        def do_save():
            try:
                new_max = max(1, min(MAX_VIRTUAL_HARD, int(mv_var.get())))
            except (ValueError, TypeError):
                new_max = MAX_VIRTUAL
            self._max_virtual = new_max
            self.settings["theme"] = theme_var.get()
            ctk.set_appearance_mode(theme_var.get())
            self._save_settings()
            win.destroy()
            messagebox.showinfo(self.t("dlg_saved"), self.t("settings_saved"))

        ctk.CTkButton(win, text=self.t("settings_save_btn"), fg_color="#2563eb",
                      command=do_save, height=40,
                      font=("Segoe UI", 12, "bold")).pack(pady=(20, 8), padx=24, fill="x")

    def copy_sequence(self):
        """Aktuelle Sequenz in die Zwischenablage kopieren."""
        seq = self.res.cget("text") if hasattr(self, "res") else ""
        if not seq or seq == "----------": return
        self.clipboard_clear()
        self.clipboard_append(seq)
        messagebox.showinfo(self.t("dlg_note"), self.t("copied_msg"))

    def pick_random_color(self):
        """Zufällige Farbe wählen und berechnen."""
        import random
        r = random.randint(0, 255)
        g = random.randint(0, 255)
        b = random.randint(0, 255)
        h = f"#{r:02X}{g:02X}{b:02X}"
        self._apply_target(h)

    def open_batch_dialog(self):
        """Batch-Eingabe: Mehrere Hex-Codes auf einmal berechnen und als virtuelle Köpfe hinzufügen."""
        win = ctk.CTkToplevel(self)
        win.title(self.t("batch_title"))
        win.geometry("480x460")
        win.grab_set()

        ctk.CTkLabel(win, text=self.t("batch_title"),
                     font=("Segoe UI", 14, "bold")).pack(pady=(16, 4))
        ctk.CTkLabel(win, text=self.t("batch_desc"),
                     font=("Segoe UI", 10), text_color="#64748b",
                     wraplength=420).pack(pady=(0, 8))

        # Farbvorschau-Reihe
        preview_row = ctk.CTkFrame(win, fg_color="transparent")
        preview_row.pack(fill="x", padx=20, pady=(0, 4))

        txt = ctk.CTkTextbox(win, height=200, font=("Courier New", 12))
        txt.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        # Vorschau aktualisieren beim Tippen
        preview_labels = []
        def update_preview(event=None):
            for w in preview_row.winfo_children(): w.destroy()
            preview_labels.clear()
            lines = txt.get("1.0", "end").strip().splitlines()
            for line in lines[:20]:
                h = line.strip()
                if not h.startswith("#"): h = "#" + h
                if re.match(r'^#[0-9A-Fa-f]{6}$', h):
                    lbl = ctk.CTkLabel(preview_row, text="", width=24, height=24,
                                       fg_color=h, corner_radius=4)
                    lbl.pack(side="left", padx=2)
                    preview_labels.append(lbl)
        txt.bind("<KeyRelease>", update_preview)

        prog = ctk.CTkLabel(win, text="", font=("Segoe UI", 10), text_color="#94a3b8")
        prog.pack()

        def do_batch():
            lines = txt.get("1.0", "end").strip().splitlines()
            hexes = []
            for line in lines:
                h = line.strip()
                if not h: continue
                if not h.startswith("#"): h = "#" + h
                if re.match(r'^#[0-9A-Fa-f]{6}$', h):
                    hexes.append(h.upper())
            added = 0
            for i, h in enumerate(hexes):
                if len(self.virtual_fils) >= self._max_virtual:
                    messagebox.showwarning(self.t("dlg_max_virtual"),
                                           self.t("batch_warn_max", max_v=self._max_virtual))
                    break
                prog.configure(text=f"{i+1}/{len(hexes)}  {h}")
                win.update_idletasks()
                r = self._calc_for_color(h,
                    optimizer=self.optimizer_var.get(),
                    seq_len=int(self.len_slider.get()) if not self.auto_len_var.get() else None,
                    auto=self.auto_len_var.get(),
                    auto_threshold=safe_td(self.auto_thresh_entry.get()))
                if r:
                    self.add_to_virtual(r)
                    added += 1
            win.destroy()
            messagebox.showinfo(self.t("dlg_3mf_title"), self.t("batch_done", n=added))

        btm = ctk.CTkFrame(win, fg_color="transparent")
        btm.pack(fill="x", padx=20, pady=(0, 14))
        ctk.CTkButton(btm, text=self.t("batch_btn_calc"), fg_color="#2563eb",
                      command=do_batch, height=40,
                      font=("Segoe UI", 12, "bold")).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btm, text=self.t("batch_btn_cancel"), fg_color="#374151",
                      command=win.destroy, height=40).pack(side="right")

    def _refresh_virtual_grid(self):
        for w in self.vgrid.winfo_children():
            w.destroy()
        if not self.virtual_fils:
            ctk.CTkLabel(self.vgrid,
                         text=self.t("empty_virtual"),
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
        ctk.CTkLabel(outer, text=self.de_label(vf["de"]),
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
            hint = self.t("hint_pure")
            hint_color = "#94a3b8"
        elif n_fils == 2 and lh > 0:
            cad = calc_cadence(vf["sequence"], lh)
            ids = sorted(cad.keys())
            hint = self.t("hint_cadence", a=cad[ids[0]], b=cad[ids[1]], p=pattern_str)
            hint_color = "#4ade80"
        else:
            hint = self.t("hint_pattern", p=pattern_str)
            hint_color = "#a78bfa"

        ctk.CTkLabel(info_row, text=hint,
                     font=("Segoe UI", 9), text_color=hint_color).pack(side="left", padx=4)

    def _remove_virtual(self, vid):
        self.virtual_undo.append(copy.deepcopy(self.virtual_fils))
        self.virtual_fils = [v for v in self.virtual_fils if v["vid"] != vid]
        for i, v in enumerate(self.virtual_fils):
            v["vid"] = 5 + i
        self._refresh_virtual_grid()

    # ── 3MF ASSISTENT ──────────────────────────────────────────────────────────

    def open_3mf_assistant(self):
        path = filedialog.askopenfilename(
            filetypes=[(self.t("3mf_filetypes"), "*.3mf"), ("*", "*.*")],
            title=self.t("open_3mf_title"))
        if not path: return

        colors, err = parse_3mf_colors(path)
        if not colors:
            messagebox.showinfo(self.t("dlg_3mf_title"),
                f"{self.t('dlg_3mf_no_colors_fallback') if not err else err}"); return

        win = ctk.CTkToplevel(self)
        win.title(f"{self.t('dlg_3mf_title')} — {os.path.basename(path)}")
        win.geometry("860x680")
        win.grab_set()

        # Header
        ctk.CTkLabel(win,
                     text=self.t("3mf_analysis_title", n=len(colors)),
                     font=("Segoe UI", 15, "bold"), text_color="#38bdf8").pack(pady=(16, 2))
        ctk.CTkLabel(win,
                     text=self.t("3mf_basis",
                          t1=self.slots[0]['hex'].get(), t2=self.slots[1]['hex'].get(),
                          t3=self.slots[2]['hex'].get(), t4=self.slots[3]['hex'].get()),
                     font=("Segoe UI", 10), text_color="#64748b").pack(pady=(0, 6))

        opt_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(win, text=self.t("3mf_optimizer"),
                        variable=opt_var, font=("Segoe UI", 11)).pack(pady=(0, 4))

        prog_label = ctk.CTkLabel(win, text=self.t("3mf_ready"),
                                   font=("Segoe UI", 11), text_color="#94a3b8")
        prog_label.pack(pady=(0, 6))

        # Spalten-Header
        hdr = ctk.CTkFrame(win, fg_color="#1e293b", corner_radius=6)
        hdr.pack(fill="x", padx=14, pady=(0, 4))
        hdr.grid_columnconfigure(4, weight=1)
        for col, txt in enumerate(["#", self.t("3mf_col_target"), self.t("3mf_col_seq"),
                                    self.t("3mf_col_sim"), self.t("3mf_col_quality"), "VID"]):
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
            for idx, hex_c in enumerate(colors[:self._max_virtual]):
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
                    ctk.CTkLabel(row, text=self.de_label(r["de"]),
                                 font=("Segoe UI", 11, "bold"),
                                 text_color=de_color(r["de"])).grid(
                        row=0, column=4, padx=6, sticky="w")
                else:
                    ctk.CTkLabel(row, text=self.t("3mf_not_calc"),
                                 text_color="#334155").grid(row=0, column=2,
                                                             columnspan=3, padx=6, sticky="w")
                sv = ctk.BooleanVar(value=True if r else False)
                vid_vars.append(sv)
                ctk.CTkCheckBox(row, text=self.t("3mf_include"), variable=sv,
                                font=("Segoe UI", 9), width=100).grid(
                    row=0, column=5, padx=(4, 10))

        render_rows()

        def run_all():
            opt = opt_var.get()
            free = self._max_virtual - len(self.virtual_fils)
            to_calc = min(len(colors), free) if free < len(colors) else len(colors)
            for idx in range(min(len(colors), self._max_virtual)):
                prog_label.configure(
                    text=self.t("3mf_progress", i=idx+1, total=to_calc, c=colors[idx]))
                win.update_idletasks()
                r = self._calc_for_color(colors[idx], optimizer=opt,
                                        seq_len=int(self.len_slider.get()),
                                        auto=self.auto_len_var.get(),
                                        auto_threshold=safe_td(self.auto_thresh_entry.get()))
                results[idx] = r
            prog_label.configure(text=self.t("3mf_done", n=to_calc))
            render_rows()

        def apply_selected():
            added = 0
            for idx, sv in enumerate(vid_vars):
                if not sv.get() or results[idx] is None: continue
                if len(self.virtual_fils) >= self._max_virtual: break
                self.add_to_virtual(results[idx])
                added += 1
            win.destroy()
            messagebox.showinfo(self.t("dlg_3mf_title"), self.t("dlg_3mf_added", n=added))

        btm = ctk.CTkFrame(win, fg_color="transparent")
        btm.pack(fill="x", padx=14, pady=(6, 14))
        ctk.CTkButton(btm, text=self.t("3mf_btn_calc"), fg_color="#2563eb",
                      command=run_all, height=42, font=("Segoe UI", 12, "bold")).pack(
            side="left", padx=(0, 6))
        ctk.CTkButton(btm, text=self.t("3mf_btn_apply"), fg_color="#15803d",
                      command=apply_selected, height=42,
                      font=("Segoe UI", 12, "bold")).pack(side="left")
        ctk.CTkButton(btm, text=self.t("3mf_btn_cancel"), fg_color="#374151",
                      command=win.destroy, height=42).pack(side="right")

    # ── EXPORT ─────────────────────────────────────────────────────────────────

    def open_export_dialog(self):
        win = ctk.CTkToplevel(self)
        win.title(self.t("exp_title"))
        win.geometry("500x440")
        win.grab_set()

        ctk.CTkLabel(win, text=self.t("exp_header"),
                     font=("Segoe UI", 14, "bold")).pack(pady=(18, 10))

        fmt_var = ctk.StringVar(value="JSON")
        ff = ctk.CTkFrame(win, fg_color="transparent")
        ff.pack()
        ctk.CTkRadioButton(ff, text="JSON (.json)", variable=fmt_var,
                           value="JSON").pack(side="left", padx=12)
        ctk.CTkRadioButton(ff, text="Text (.txt)", variable=fmt_var,
                           value="TXT").pack(side="left", padx=12)

        ctk.CTkLabel(win, text=self.t("exp_dither_title"),
                     font=("Segoe UI", 11, "bold"), text_color="#64748b").pack(pady=(16, 4))
        ctk.CTkLabel(win, text=self.t("exp_dither_desc"),
                     font=("Segoe UI", 9), text_color="#475569").pack(pady=(0, 6))

        lhf = ctk.CTkFrame(win, fg_color="transparent"); lhf.pack()
        ctk.CTkLabel(lhf, text=self.t("exp_lh_label")).pack(side="left", padx=6)
        lh_val = self.layer_height_entry.get() if hasattr(self, "layer_height_entry") else "0.2"
        lh_exp = ctk.CTkEntry(lhf, width=60); lh_exp.insert(0, lh_val)
        lh_exp.pack(side="left", padx=4)
        ctk.CTkLabel(lhf, text=self.t("exp_lh_unit")).pack(side="left", padx=4)

        cf = ctk.CTkFrame(win, fg_color="transparent"); cf.pack(pady=(6, 0))
        ctk.CTkLabel(cf, text=self.t("exp_cadence_a")).pack(side="left", padx=6)
        ca = ctk.CTkEntry(cf, width=70, placeholder_text="auto"); ca.pack(side="left", padx=4)
        ctk.CTkLabel(cf, text=self.t("exp_cadence_sep")).pack(side="left", padx=4)
        cb = ctk.CTkEntry(cf, width=70, placeholder_text="auto"); cb.pack(side="left", padx=4)
        ctk.CTkLabel(cf, text=self.t("exp_cadence_auto"), font=("Segoe UI", 9),
                     text_color="#475569").pack(side="left", padx=4)

        scope_var = ctk.StringVar(value="virtual" if self.virtual_fils else "single")
        sf = ctk.CTkFrame(win, fg_color="transparent")
        sf.pack(pady=(14, 4))
        ctk.CTkRadioButton(sf, text=self.t("exp_scope_single"),
                           variable=scope_var, value="single").pack(side="left", padx=12)
        ctk.CTkRadioButton(sf, text=self.t("exp_scope_virtual", n=len(self.virtual_fils)),
                           variable=scope_var, value="virtual").pack(side="left", padx=12)

        def do_export():
            fmt   = fmt_var.get()
            ext   = ".json" if fmt == "JSON" else ".txt"
            scope = scope_var.get()
            path  = filedialog.asksaveasfilename(
                defaultextension=ext,
                filetypes=[(fmt, f"*{ext}")],
                title=self.t("save_dialog_title"))
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
                        f"{self.t('txt_date')}              {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                        f"{self.t('txt_layer_height')}        {lh} mm  (= Dithering Step Size)",
                        "", self.t("txt_physical_heads"), "-" * 36,
                    ]
                    for p in physical:
                        lines.append(f"  T{p['id']}: {p['brand']} — {p['filament']}  ({p['hex']}, TD={p['td']})")
                    if scope == "virtual" and self.virtual_fils:
                        lines += ["", self.t("txt_virtual_heads"), "-" * 56]
                        for v in self.virtual_fils:
                            a_v, b_v = resolve_cadence(v["sequence"])
                            n_f = seq_filament_count(v["sequence"])
                            runs_str = "  ".join(f"T{fid}×{cnt}" for fid, cnt in seq_to_runs(v["sequence"]))
                            pat_str  = "/".join(v["sequence"])
                            if n_f == 1:
                                slicer_hint = self.t("txt_pure")
                            elif n_f == 2:
                                slicer_hint = self.t("txt_cadence", a=a_v, b=b_v, p=pat_str)
                            else:
                                slicer_hint = self.t("txt_pattern", p=pat_str)
                            lines += [
                                f"  V{v['vid']}  [{v['label']}]  {self.t('grid_target')}:{v['target_hex']}  ΔE={v['de']:.1f}",
                                f"    {self.t('txt_sequence')} {v['sequence']}",
                                f"    Runs:    {runs_str}",
                                f"    Slicer:  {slicer_hint}",
                            ]
                    else:
                        a_v, b_v = resolve_cadence(seq_now)
                        lines += ["", f"{self.t('txt_sequence')}  {seq_now}",
                                  f"{self.t('txt_target')}     {getattr(self, 'target', '')}",
                                  f"{self.t('txt_cadence2')}  A={a_v}mm / B={b_v}mm",
                                  f"{self.t('txt_runs')}     {'  '.join(f'T{fid}×{cnt}' for fid,cnt in seq_to_runs(seq_now))}"]
                    lines += ["", "github.com/ratdoux/OrcaSlicer-FullSpectrum", "=" * 56]
                    with open(path, "w", encoding="utf-8") as f:
                        f.write("\n".join(lines))
                messagebox.showinfo(self.t("dlg_saved"), self.t("dlg_export_saved", path=path))
                win.destroy()
            except IOError as e:
                messagebox.showerror(self.t("dlg_error"), str(e))

        ctk.CTkButton(win, text=self.t("exp_btn"), fg_color="#7c3aed",
                      command=do_export, height=44,
                      font=("Segoe UI", 14, "bold")).pack(pady=(18, 4), padx=40, fill="x")
        ctk.CTkButton(win, text=self.t("btn_orca_export"), fg_color="#0f766e",
                      command=lambda: [win.destroy(), self.open_orca_export_dialog()],
                      height=40, font=("Segoe UI", 13, "bold")).pack(pady=(0, 4), padx=40, fill="x")
        ctk.CTkButton(win, text=self.t("exp_cancel"), fg_color="#334155",
                      command=win.destroy, height=36).pack(padx=40, fill="x")

    # ── ORCASLICER DIREKT-EXPORT ───────────────────────────────────────────────

    @staticmethod
    def _find_orca_installations() -> list:
        """
        Sucht alle OrcaSlicer-kompatiblen Installationen auf dem System.
        Gibt eine Liste von Dicts zurück: {"label": str, "path": str, "exists": bool}
        Durchsucht AppData/Roaming nach Ordnern mit user/default/filament Struktur,
        sowie bekannte portable Pfade (Exe-Verzeichnis + data/).
        """
        import platform
        found = []
        seen = set()

        def add(label, path):
            norm = os.path.normcase(os.path.normpath(path))
            if norm not in seen:
                seen.add(norm)
                found.append({"label": label, "path": path,
                               "exists": os.path.isdir(path)})

        if platform.system() == "Windows":
            appdata = os.environ.get("APPDATA", "")
            if appdata and os.path.isdir(appdata):
                # Alle Unterordner von AppData\Roaming durchsuchen
                for entry in os.scandir(appdata):
                    if not entry.is_dir():
                        continue
                    name = entry.name
                    # OrcaSlicer-ähnliche Ordnernamen
                    if any(x in name.lower() for x in
                           ["orca", "snapmaker_orca", "bambu", "orca slicer"]):
                        for sub in ["user/default/filament", "user/filament"]:
                            p = os.path.join(entry.path, sub)
                            add(f"{name}  ({p})", p)

            # Bekannte portable Pfade: wenn neben einer Exe ein "data"-Ordner liegt
            portable_hints = [
                os.path.expanduser("~/Desktop"),
                os.path.expanduser("~/Downloads"),
                "C:/OrcaSlicer",
                "D:/OrcaSlicer",
                "C:/Snapmaker_Orca",
            ]
            for base in portable_hints:
                if not os.path.isdir(base):
                    continue
                for entry in os.scandir(base):
                    if not entry.is_dir():
                        continue
                    # data/user/default/filament neben Exe
                    p = os.path.join(entry.path, "data", "user", "default", "filament")
                    if os.path.isdir(p):
                        add(f"{entry.name} [portable]  ({p})", p)
                    # AppData-style innerhalb portablem Ordner
                    p2 = os.path.join(entry.path, "user", "default", "filament")
                    if os.path.isdir(p2):
                        add(f"{entry.name} [portable]  ({p2})", p2)

        elif platform.system() == "Darwin":
            base = os.path.expanduser("~/Library/Application Support")
            if os.path.isdir(base):
                for entry in os.scandir(base):
                    if entry.is_dir() and "orca" in entry.name.lower():
                        for sub in ["user/default/filament", "user/filament"]:
                            p = os.path.join(entry.path, sub)
                            add(f"{entry.name}  ({p})", p)
        else:
            base = os.path.expanduser("~/.config")
            if os.path.isdir(base):
                for entry in os.scandir(base):
                    if entry.is_dir() and "orca" in entry.name.lower():
                        p = os.path.join(entry.path, "user", "default", "filament")
                        add(f"{entry.name}  ({p})", p)

        # Existierende zuerst, dann potenzielle
        found.sort(key=lambda x: (0 if x["exists"] else 1, x["label"]))
        return found

    @staticmethod
    def _detect_orca_filament_path() -> str:
        """Gibt den ersten gefundenen Filament-Pfad zurück (Rückwärtskompatibilität)."""
        installs = App._find_orca_installations()
        for i in installs:
            if i["exists"]:
                return i["path"]
        return installs[0]["path"] if installs else ""

    def _build_orca_filament_json(self, name: str, hex_color: str, notes: str,
                                   filament_type: str = "PLA",
                                   target_path: str = "") -> dict:
        """
        Erstellt ein OrcaSlicer-kompatibles Filament-Profil-JSON.
        Wählt automatisch das richtige 'inherits'-Format je nach Slicer
        (OrcaSlicer: fdm_filament_pla  vs.  Snapmaker_Orca: Generic PLA).
        """
        hex_clean = hex_color if hex_color.startswith("#") else f"#{hex_color}"
        ft = filament_type.upper()
        # Snapmaker_Orca und ähnliche nutzen "Generic XYZ" als Basisklasse
        if "snapmaker" in target_path.lower() or "snapmaker_orca" in target_path.lower():
            inherits = f"Generic {ft}"
        else:
            inherits = f"fdm_filament_{ft.lower()}"
        return {
            "type": "filament",
            "name": name,
            "inherits": inherits,
            "from": "User",
            "is_custom_defined": "0",
            "instantiation": "true",
            "filament_vendor": ["U1 FullSpectrum"],
            "filament_notes": [notes],
            "default_filament_colour": [hex_clean],
            "compatible_printers": [],
            "filament_settings_id": [name],
        }

    def open_orca_export_dialog(self):
        installs = self._find_orca_installations()
        path_var = ctk.StringVar(value=installs[0]["path"] if installs else "")

        win = ctk.CTkToplevel(self)
        win.title(self.t("orca_title"))
        win.geometry("600x560")
        win.grab_set()

        ctk.CTkLabel(win, text=self.t("orca_header"),
                     font=("Segoe UI", 14, "bold")).pack(pady=(18, 6))

        # Slicer-Auswahl (Dropdown wenn mehrere gefunden)
        if installs:
            ctk.CTkLabel(win, text="Erkannte Slicer-Installation:",
                         font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=20)
            # Kurze Labels für Dropdown (nur Ordnername)
            def short_label(i):
                parts = i["path"].replace("\\", "/").split("/")
                # z.B. "Snapmaker_Orca ✓" oder "OrcaSlicer (leer)"
                appname = parts[-4] if len(parts) >= 4 else i["label"]
                status = " ✓" if i["exists"] else " (noch nicht vorhanden)"
                return f"{appname}{status}"
            labels = [short_label(i) for i in installs]
            sel_var = ctk.StringVar(value=labels[0])
            def on_slicer_select(choice):
                idx = labels.index(choice)
                path_var.set(installs[idx]["path"])
            ctk.CTkOptionMenu(win, variable=sel_var, values=labels,
                              command=on_slicer_select,
                              width=560).pack(padx=20, pady=(2, 8))
        else:
            ctk.CTkLabel(win, text=self.t("orca_no_path"),
                         text_color="#f87171", font=("Segoe UI", 10)).pack(pady=(0, 4))

        # Pfad-Zeile (manuell überschreibbar)
        pf = ctk.CTkFrame(win, fg_color="transparent"); pf.pack(fill="x", padx=20, pady=(0, 4))
        ctk.CTkLabel(pf, text=self.t("orca_path_label"),
                     font=("Segoe UI", 9), text_color="#64748b").pack(anchor="w")
        pe = ctk.CTkEntry(pf, textvariable=path_var, width=460,
                          font=("Segoe UI", 9))
        pe.pack(side="left", fill="x", expand=True, pady=2)
        def browse_path():
            p = filedialog.askdirectory(title=self.t("orca_path_label"))
            if p: path_var.set(p)
        ctk.CTkButton(pf, text=self.t("orca_path_browse"), width=80,
                      command=browse_path).pack(side="left", padx=(6, 0))

        # Scope
        ctk.CTkLabel(win, text="Exportieren:", font=("Segoe UI", 11, "bold")).pack(
            pady=(10, 2))
        scope_var = ctk.StringVar(value="both")
        sf = ctk.CTkFrame(win, fg_color="transparent"); sf.pack()
        ctk.CTkRadioButton(sf, text=self.t("orca_scope_phys"),
                           variable=scope_var, value="phys").pack(anchor="w", padx=20, pady=1)
        ctk.CTkRadioButton(sf, text=self.t("orca_scope_virt"),
                           variable=scope_var, value="virt").pack(anchor="w", padx=20, pady=1)
        ctk.CTkRadioButton(sf, text=self.t("orca_scope_both"),
                           variable=scope_var, value="both").pack(anchor="w", padx=20, pady=1)

        # Filament-Typ
        ctk.CTkLabel(win, text="Basis-Filamenttyp:", font=("Segoe UI", 11, "bold")).pack(
            pady=(10, 2))
        ftype_var = ctk.StringVar(value="PLA")
        ft = ctk.CTkFrame(win, fg_color="transparent"); ft.pack()
        for ft_name in ["PLA", "PETG", "ABS", "TPU"]:
            ctk.CTkRadioButton(ft, text=ft_name, variable=ftype_var,
                               value=ft_name).pack(side="left", padx=10)

        # Prefix
        pxf = ctk.CTkFrame(win, fg_color="transparent"); pxf.pack(pady=(10, 0))
        ctk.CTkLabel(pxf, text=self.t("orca_prefix_label")).pack(side="left", padx=6)
        px_entry = ctk.CTkEntry(pxf, width=80, placeholder_text="U1"); px_entry.insert(0, "U1")
        px_entry.pack(side="left", padx=4)
        ctk.CTkLabel(pxf, text=self.t("orca_prefix_hint"),
                     text_color="#64748b", font=("Segoe UI", 9)).pack(side="left", padx=4)

        # Layer height for dithering step
        lhf = ctk.CTkFrame(win, fg_color="transparent"); lhf.pack(pady=(6, 0))
        ctk.CTkLabel(lhf, text=self.t("exp_lh_label")).pack(side="left", padx=6)
        lh_val = self.layer_height_entry.get() if hasattr(self, "layer_height_entry") else "0.2"
        lh_entry = ctk.CTkEntry(lhf, width=60); lh_entry.insert(0, lh_val)
        lh_entry.pack(side="left", padx=4)
        ctk.CTkLabel(lhf, text="mm").pack(side="left")

        def do_orca_export():
            folder = path_var.get().strip()
            if not folder:
                messagebox.showerror(self.t("dlg_error"), self.t("orca_no_path"))
                return
            scope   = scope_var.get()
            prefix  = px_entry.get().strip() or "U1"
            ftype   = ftype_var.get()
            lh      = safe_td(lh_entry.get()) if lh_entry.get().strip() else 0.2

            profiles_to_write = []  # list of (filename, dict)

            # Physische Slots
            if scope in ("phys", "both"):
                for i, s in enumerate(self.slots):
                    hex_c  = s["hex"].get().strip() or "#888888"
                    brand  = s["brand"].get().strip()
                    name   = s["color"].get().strip()
                    td     = s["td"].get().strip()
                    pname  = f"{prefix}-T{i+1} {name}" if name and name not in _SLOT_SKIP else f"{prefix}-T{i+1}"
                    notes  = self.t("orca_filament_notes_t", i=i+1, brand=brand, name=name, td=td)
                    data   = self._build_orca_filament_json(pname, hex_c, notes, ftype, folder)
                    safe_fn = re.sub(r'[\\/:*?"<>|]', "_", pname) + ".json"
                    profiles_to_write.append((safe_fn, data))

            # Virtuelle Köpfe
            if scope in ("virt", "both"):
                if not self.virtual_fils:
                    messagebox.showwarning(self.t("dlg_saved"), self.t("orca_no_virtual"))
                else:
                    for v in self.virtual_fils:
                        sim_hex = v.get("sim_hex", v.get("target_hex", "#888888"))
                        seq     = v["sequence"]
                        n_f     = seq_filament_count(seq)
                        runs_str = "  ".join(f"T{fid}×{cnt}" for fid, cnt in seq_to_runs(seq))
                        cad     = calc_cadence(seq, lh)
                        ids     = sorted(cad.keys())
                        if n_f == 1:
                            hint = "Pure color"
                        elif n_f == 2:
                            a = round(cad.get(ids[0], lh), 3)
                            b = round(cad.get(ids[1], lh) if len(ids) > 1 else lh, 3)
                            hint = f"Cadence A={a}mm B={b}mm | Step={lh}mm"
                        else:
                            pat = "/".join(seq)
                            hint = f"Pattern={pat} | Step={lh}mm"
                        label  = v.get("label", f"V{v['vid']}")
                        pname  = f"{prefix}-V{v['vid']} {label}"
                        notes  = self.t("orca_filament_notes_v",
                                        seq=runs_str, de=v.get("de", 0.0), hint=hint)
                        data   = self._build_orca_filament_json(pname, sim_hex, notes, ftype, folder)
                        safe_fn = re.sub(r'[\\/:*?"<>|]', "_", pname) + ".json"
                        profiles_to_write.append((safe_fn, data))

            if not profiles_to_write:
                return

            # Überprüfen ob Dateien überschrieben werden
            existing = [fn for fn, _ in profiles_to_write
                        if os.path.exists(os.path.join(folder, fn))]
            if existing:
                if not messagebox.askyesno(self.t("orca_title"),
                    self.t("orca_overwrite_confirm", n=len(existing))):
                    return

            # Ordner anlegen falls nötig
            os.makedirs(folder, exist_ok=True)

            written = 0
            errors  = []
            for fn, data in profiles_to_write:
                try:
                    fpath = os.path.join(folder, fn)
                    with open(fpath, "w", encoding="utf-8") as f:
                        json.dump(data, f, indent=4, ensure_ascii=False)
                    written += 1
                except IOError as e:
                    errors.append(str(e))

            if errors:
                messagebox.showerror(self.t("dlg_error"), "\n".join(errors[:3]))
            else:
                messagebox.showinfo(self.t("dlg_saved"),
                    self.t("orca_success", n=written))
                win.destroy()

        ctk.CTkButton(win, text=self.t("orca_btn_export"), fg_color="#0f766e",
                      command=do_orca_export, height=44,
                      font=("Segoe UI", 14, "bold")).pack(pady=(16, 4), padx=40, fill="x")
        ctk.CTkButton(win, text=self.t("orca_btn_cancel"), fg_color="#334155",
                      command=win.destroy, height=36).pack(padx=40, fill="x")

    # ── GRADIENT-GENERATOR ─────────────────────────────────────────────────────

    def open_gradient_dialog(self):
        """Farbverlauf zwischen zwei Farben berechnen und als virtuelle Köpfe hinzufügen."""
        win = ctk.CTkToplevel(self)
        win.title(self.t("gradient_title"))
        win.geometry("480x420")
        win.grab_set()

        ctk.CTkLabel(win, text=self.t("gradient_title"),
                     font=("Segoe UI", 14, "bold")).pack(pady=(16, 10))

        from_var = ctk.StringVar(value=getattr(self, "target", "#FF2020"))
        to_var   = ctk.StringVar(value="#2020FF")

        def make_color_row(label_key, var):
            f = ctk.CTkFrame(win, fg_color="transparent")
            f.pack(fill="x", padx=24, pady=3)
            ctk.CTkLabel(f, text=self.t(label_key), width=50).pack(side="left")
            sw = ctk.CTkLabel(f, text="", width=36, height=36,
                               fg_color=var.get() if len(var.get())==7 else "#888888",
                               corner_radius=8)
            sw.pack(side="left", padx=6)
            entry = ctk.CTkEntry(f, textvariable=var, width=100)
            entry.pack(side="left", padx=4)
            def pick(v=var, s=sw):
                from tkinter import colorchooser as _cc
                col = _cc.askcolor(color=v.get(), title="Farbe")[1]
                if col: v.set(col.upper()); s.configure(fg_color=col)
            def on_change(*a, v=var, s=sw):
                h = v.get().strip()
                if len(h) == 7 and h.startswith("#"):
                    try: s.configure(fg_color=h)
                    except: pass
            var.trace_add("write", on_change)
            ctk.CTkButton(f, text="🎨", width=34, height=34,
                          fg_color="#374151", command=pick).pack(side="left", padx=2)

        make_color_row("gradient_from", from_var)
        make_color_row("gradient_to",   to_var)

        sf = ctk.CTkFrame(win, fg_color="transparent")
        sf.pack(pady=8)
        ctk.CTkLabel(sf, text=self.t("gradient_steps")).pack(side="left", padx=6)
        steps_var = ctk.IntVar(value=6)
        steps_lbl = ctk.CTkLabel(sf, text="6", width=30, font=("Segoe UI", 12, "bold"))
        def on_steps(v):
            steps_lbl.configure(text=str(int(float(v))))
            update_preview()
        ctk.CTkSlider(sf, from_=2, to=min(16, self._max_virtual),
                       number_of_steps=14, variable=steps_var,
                       command=on_steps).pack(side="left", padx=6)
        steps_lbl.pack(side="left")

        # Preview strip
        pf = ctk.CTkFrame(win, fg_color="#0f172a", corner_radius=6, height=28)
        pf.pack(fill="x", padx=24, pady=8)
        pcells = []
        for _ in range(16):
            c = ctk.CTkLabel(pf, text="", width=0, height=28,
                              fg_color="#1e293b", corner_radius=0)
            c.pack(side="left", fill="x", expand=True)
            pcells.append(c)

        def update_preview(*a):
            try:
                h1, h2 = from_var.get().strip(), to_var.get().strip()
                if len(h1) != 7 or len(h2) != 7: return
                l1 = rgb_to_lab(hex_to_rgb(h1))
                l2 = rgb_to_lab(hex_to_rgb(h2))
                for i, cell in enumerate(pcells):
                    t = i / (len(pcells) - 1)
                    mix = tuple(l1[k] + t*(l2[k]-l1[k]) for k in range(3))
                    cell.configure(fg_color=lab_to_hex(mix))
            except: pass

        from_var.trace_add("write", update_preview)
        to_var.trace_add("write", update_preview)
        update_preview()

        def do_gradient():
            h1, h2 = from_var.get().strip(), to_var.get().strip()
            n = int(steps_var.get())
            if len(h1) != 7 or len(h2) != 7: return
            slots_left = self._max_virtual - len(self.virtual_fils)
            if slots_left <= 0:
                messagebox.showwarning(self.t("dlg_max_virtual"),
                    self.t("batch_warn_max", max_v=self._max_virtual)); return
            n = min(n, slots_left)
            l1 = rgb_to_lab(hex_to_rgb(h1))
            l2 = rgb_to_lab(hex_to_rgb(h2))
            self.virtual_undo.append(copy.deepcopy(self.virtual_fils))
            added = 0
            for step in range(n):
                t = step / max(n - 1, 1)
                target_hex = lab_to_hex(tuple(l1[k] + t*(l2[k]-l1[k]) for k in range(3)))
                res = self._calc_for_color(target_hex, False,
                                            seq_len=None, auto=True, auto_threshold=3.0)
                if res is None or len(self.virtual_fils) >= self._max_virtual:
                    break
                vid = 5 + len(self.virtual_fils)
                self.virtual_fils.append({
                    "vid": vid, "target_hex": target_hex,
                    "sequence": res["sequence"], "sim_hex": res["sim_hex"],
                    "de": res["de"], "label": f"Grad {step+1}/{n}",
                })
                added += 1
            self._refresh_virtual_grid()
            messagebox.showinfo(self.t("dlg_saved"), self.t("gradient_done", n=added))
            win.destroy()

        ctk.CTkButton(win, text=self.t("gradient_btn_calc"), fg_color="#0e7490",
                      command=do_gradient, height=42,
                      font=("Segoe UI", 13, "bold")).pack(pady=(4, 4), padx=24, fill="x")
        ctk.CTkButton(win, text=self.t("exp_cancel"), fg_color="#334155",
                      command=win.destroy, height=36).pack(padx=24, fill="x")

    # ── WEB-UPDATE ─────────────────────────────────────────────────────────────

    # Community-Datenbank: als Fallback auch lokal gespeichert
    _COMMUNITY_URL = (
        "https://raw.githubusercontent.com/halloworld007/"
        "snapmaker-u1-fullspectrum-helper/main/filament_community.json"
    )

    def web_update_library(self):
        """Lädt Community-Farbdaten von GitHub und merged sie in die Bibliothek."""
        import threading

        win = ctk.CTkToplevel(self)
        win.title(self.t("web_update_title"))
        win.geometry("420x160")
        win.grab_set()
        lbl = ctk.CTkLabel(win, text=self.t("web_update_fetching"),
                           font=("Segoe UI", 12))
        lbl.pack(expand=True, pady=30)
        pb = ctk.CTkProgressBar(win, mode="indeterminate")
        pb.pack(padx=40, fill="x"); pb.start()

        def fetch():
            try:
                import urllib.request
                with urllib.request.urlopen(self._COMMUNITY_URL, timeout=10) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
            except Exception as e:
                win.after(0, lambda: done(None, str(e)))
                return
            win.after(0, lambda: done(data, None))

        def done(data, err):
            pb.stop(); pb.pack_forget()
            if err:
                lbl.configure(text=self.t("web_update_err", e=err),
                               text_color="#f87171", wraplength=380)
                ctk.CTkButton(win, text="OK", command=win.destroy,
                               width=100).pack(pady=8)
                return

            if not isinstance(data, dict):
                lbl.configure(text=self.t("web_update_err", e="Invalid format"),
                               text_color="#f87171")
                ctk.CTkButton(win, text="OK", command=win.destroy,
                               width=100).pack(pady=8)
                return

            new_count = 0
            n_brands  = 0
            n_fils    = 0
            for brand, entries in data.items():
                if brand == "Eigene Favoriten":
                    continue
                n_brands += 1
                existing_names = {f["name"] for f in self.library.get(brand, [])}
                for entry in entries:
                    n_fils += 1
                    if entry.get("name") not in existing_names:
                        self.library.setdefault(brand, []).append(entry)
                        new_count += 1

            if new_count:
                self.save_db()
                msg = self.t("web_update_ok", n_brands=n_brands,
                              n_fils=n_fils, new=new_count)
                lbl.configure(text=msg, text_color="#4ade80", wraplength=380)
            else:
                lbl.configure(text=self.t("web_update_no_new"),
                               text_color="#94a3b8")
            ctk.CTkButton(win, text="OK", command=win.destroy,
                           width=100).pack(pady=8)

        threading.Thread(target=fetch, daemon=True).start()

    # ── BIBLIOTHEK-MANAGER ─────────────────────────────────────────────────────

    def open_library_manager(self):
        win = ctk.CTkToplevel(self)
        win.title(self.t("lib_title")); win.geometry("720x560"); win.grab_set()

        top = ctk.CTkFrame(win, fg_color="transparent")
        top.pack(fill="x", padx=14, pady=10)
        ctk.CTkLabel(top, text=self.t("lib_brand")).pack(side="left", padx=5)
        bv = ctk.StringVar(value=list(self.library.keys())[0])
        bm = ctk.CTkOptionMenu(top, variable=bv,
                                values=list(self.library.keys()),
                                command=lambda x: refresh())
        bm.pack(side="left", padx=5, fill="x", expand=True)
        ctk.CTkButton(top, text=self.t("lib_del_brand"), fg_color="#991b1b",
                      command=lambda: del_brand()).pack(side="right", padx=5)

        scroll = ctk.CTkScrollableFrame(win, height=380)
        scroll.pack(fill="both", expand=True, padx=14, pady=5)

        def refresh():
            for w in scroll.winfo_children(): w.destroy()
            brand = bv.get()
            fils = self.library.get(brand, [])
            if not fils:
                ctk.CTkLabel(scroll, text=self.t("lib_no_fils"),
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
            if messagebox.askyesno(self.t("dlg_del_title"), self.t("lib_del_fil", n=n)):
                self.library[brand].pop(idx)
                self.save_db(); self._refresh_brand_menus(); refresh()

        def del_brand():
            b = bv.get()
            if b in DEFAULT_LIBRARY and DEFAULT_LIBRARY.get(b):
                messagebox.showerror(self.t("lib_protected"), self.t("lib_protected_msg")); return
            if messagebox.askyesno(self.t("dlg_del_title"), self.t("lib_del_brand_msg", b=b)):
                del self.library[b]; self.save_db(); self._refresh_brand_menus()
                nb = list(self.library.keys())
                bm.configure(values=nb); bv.set(nb[0] if nb else ""); refresh()

        refresh()
        btm = ctk.CTkFrame(win, fg_color="transparent")
        btm.pack(fill="x", padx=14, pady=(5, 14))
        ctk.CTkButton(btm, text=self.t("lib_add_fil"), fg_color="#15803d",
                      command=lambda: self._lm_add(bv.get(), refresh)).pack(side="left")
        ctk.CTkButton(btm, text=self.t("lib_close"), fg_color="#334155",
                      command=win.destroy).pack(side="right")

    def _lm_add(self, brand, cb):
        n = ctk.CTkInputDialog(text=self.t("inp_name"), title=self.t("inp_add_title")).get_input()
        if not n: return
        h = ctk.CTkInputDialog(text=self.t("inp_hex"), title=self.t("inp_color_title")).get_input()
        if not h: return
        h = h.strip()
        if not h.startswith("#"): h = "#" + h
        td_r = ctk.CTkInputDialog(text=self.t("inp_td2", td=DEFAULT_TD), title=self.t("inp_td_title")).get_input()
        self.library.setdefault(brand, []).append(
            {"name": n.strip(), "hex": h, "td": safe_td(td_r) if td_r else DEFAULT_TD})
        self.save_db(); self._refresh_brand_menus(); cb()


if __name__ == "__main__":
    U1FullSpectrumApp().mainloop()
