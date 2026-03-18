import customtkinter as ctk
import json
import os
import copy
import zipfile
import re
from itertools import permutations as iter_permutations
import tkinter as tk
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
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAS_DND = True
except ImportError:
    _HAS_DND = False

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
    "lh_hint": "⚠ 0.08 mm empfohlen — Farbschichten werden unsichtbar",
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
    "sec2_desc": "Jeder virtuelle Kopf = FullSpectrum-Sequenz (1–48 Layer) aus T1–T4  ·  max. {max_v}",
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
    "tip_add_virtual": "Als virtuellen Druckkopf V5-V24 hinzufügen (Ctrl+Enter)",
    "tip_copy": "Sequenz in Zwischenablage kopieren (Ctrl+C)",
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
    "orca_fs_hint": "⚠ FullSpectrum-Slicer erkannt: Nur T1–T4 exportieren! V5+ werden von FullSpectrum automatisch als 'Mixed Filaments' generiert — wenn du sie zusätzlich importierst, entstehen doppelte physische Slots.",
    # Projekt
    "btn_proj_save": "💾 Projekt speichern",
    "btn_proj_load": "📂 Projekt laden",
    "proj_saved": "Projekt gespeichert:\n{path}",
    "proj_loaded": "Projekt geladen:\n{path}",
    "proj_filetypes": "U1-Projekt",
    "proj_err": "Fehler beim Laden:\n{e}",
    # TD-Kalibrierung
    "btn_td_cal": "🔬 TD kalibrieren",
    "td_cal_title": "TD-Kalibrierungsassistent",
    "td_cal_desc": "Drucke T{a} und T{b} abwechselnd (z.B. 5+5 Layer) und miss die Mischfarbe.",
    "td_cal_measured": "Gemessene Hex-Farbe des Testdrucks:",
    "td_cal_result": "Empfohlene TD-Werte:  T{a} = {ta:.1f}    T{b} = {tb:.1f}",
    "td_cal_apply": "Übernehmen",
    "td_cal_slot": "Slot:",
    # Top-3
    "top3_title": "Top-3 Sequenzen",
    "top3_rank": "#{r}",
    "top3_add": "+ Hinzufügen",
    # Slot-Optimizer
    "btn_slot_opt": "🎯 Slot-Optimizer",
    "slot_opt_title": "Slot-Optimizer — Besten 4 Filamente finden",
    "slot_opt_desc": "Berechnet welche 4 Filamente aus der Bibliothek den größten Farbraum abdecken.",
    "slot_opt_targets": "Zielfarben (aus virtuellen Köpfen oder manuell):",
    "slot_opt_use_virtual": "Virtuelle Köpfe als Ziele verwenden",
    "slot_opt_running": "Berechne … {i}/{total}",
    "slot_opt_done": "Fertig! Bestes Set gefunden.",
    "slot_opt_apply": "Als T1–T4 übernehmen",
    "slot_opt_result": "Bestes Set (ΔE Ø {de:.1f}):",
    # Paletten-Import
    "btn_palette": "🖼 Palette aus Bild",
    "tip_palette": "Dominante Farben aus Bild extrahieren und als Batch hinzufügen",
    "palette_title": "Palette aus Bild importieren",
    "palette_colors": "Farben:",
    "palette_btn_add": "Alle berechnen & hinzufügen",
    # 3MF zurückschreiben
    "btn_3mf_write": "✏️ 3MF schreiben",
    "tip_3mf_write": "Filament-Farben der virtuellen Köpfe in eine 3MF-Datei schreiben",
    "3mf_write_title": "Filamentfarben in 3MF schreiben",
    "3mf_write_desc": "Wähle eine 3MF-Datei. Die Filament-Farben werden mit den simulierten Farben der virtuellen Köpfe aktualisiert.",
    "3mf_write_ok": "3MF erfolgreich aktualisiert:\n{path}",
    "3mf_write_err": "Fehler beim Schreiben:\n{e}",
    "remap_title": "3MF Extruder-Remap",
    "remap_col_hdr": "Originale Extruder  →  Ziel (T1–T4 oder V-Kopf)",
    "remap_keep": "(unverändert lassen)",
    "remap_btn_write": "3MF schreiben",
    # Werkzeugwechsel
    "btn_tc_est": "🔄 Werkzeugwechsel",
    "tc_title": "Werkzeugwechsel-Schätzung",
    "tc_layers": "Druckschichten gesamt:",
    "tc_result": "Geschätzte Werkzeugwechsel: {n}",
    "tc_time": "Geschätzte Zusatzzeit: ~{min:.0f} min  (bei {sec}s/Wechsel)",
    "tc_purge": "Geschätzter Purge-Verbrauch: ~{g:.0f}g",
    # Farbharmonien
    "btn_harmonies": "🎨 Harmonien",
    "tip_harmonies": "Komplementär-, Triad- und Analogfarben zur Zielfarbe",
    "harm_title": "Farbharmonien",
    "harm_complement": "Komplementär",
    "harm_triadic": "Triade",
    "harm_analogous": "Analog",
    "harm_split": "Split-Komplementär",
    "harm_add_all": "Alle berechnen & hinzufügen",
    # Farbmodell
    "model_label": "Farbmodell:",
    "model_linear": "Additiv — FullSpectrum-kompatibel",
    "model_td": "TD-gewichtet (physikalisch)",
    "model_subtractive": "Subtraktiv (Pigment-Modell)",
    "model_filamentmixer": "FilamentMixer (Pigment)",
    # Streifenrisiko
    "stripe_risk": "\u26a0 Streifenrisiko \u2014 Sequenzl\u00e4nge {n} ung\u00fcnstig",
    "stripe_ok": "\u2713 Kein Streifenrisiko",
    # Multi-Gradient
    "btn_multi_gradient": "\U0001f308 Multi-Gradient",
    "multi_gradient_title": "Multi-Gradient virtueller Kopf",
    "multi_gradient_desc": "Gewichteten Verlauf aus allen 4 Slots erstellen",
    "multi_gradient_auto": "Auto-Balance",
    "multi_gradient_add": "Als virtuellen Kopf hinzuf\u00fcgen",
    # Alle Cadence-Werte kopieren
    "btn_copy_all_cad": "📋 Alle Cadence-Werte",
    "copy_all_title": "Alle Dithering-Werte",
    "copy_all_btn": "Alles in Zwischenablage",
    # Tabs
    "tab_tools": "Werkzeuge",
    # Gamut-Plot
    "btn_gamut_plot": "🎯 Gamut-Plot",
    # Batch-Neuberechnung
    "btn_recalc_all": "🔄 Alle neu berechnen",
    "recalc_all_done": "{n} virtuelle Köpfe neu berechnet.",
    # ΔE-Übersicht
    "btn_de_overview": "📊 ΔE-Übersicht",
    "de_overview_title": "ΔE-Übersicht — Alle virtuellen Köpfe",
    "de_overview_col_id": "ID",
    "de_overview_col_label": "Label",
    "de_overview_col_seq": "Sequenz",
    "de_overview_col_de": "ΔE",
    "de_overview_col_quality": "Qualität",
    "de_overview_col_tc": "WW/Zyklus",
    # Farbrezept
    "btn_recipe": "📜 Farbrezept",
    "recipe_title": "Farbrezept-Export",
    "recipe_copy_btn": "In Zwischenablage",
    # Multi-Ziel-Optimizer
    "btn_multitarget": "🎯 Multi-Ziel",
    "mt_title": "Mehrfach-Zielfarben-Optimierung",
    "mt_desc": "Findet die beste Sequenz, die mehrere Zielfarben gleichzeitig approximiert.",
    "mt_target": "Ziel {n}:",
    "mt_add_target": "+ Zielfarbe hinzufügen",
    "mt_calc": "⚙  Optimieren",
    "mt_result": "Beste Sequenz: {seq}   ΔE Ø {de:.1f}",
    "mt_add_btn": "➕ Als virtuellen Kopf hinzufügen",
    "mt_no_targets": "Bitte mindestens 2 Zielfarben angeben.",
    # Slot geladen-Status
    "slot_loaded": "Geladen",
    # Werkzeugwechsel-Warnung pro V-Kopf
    "tc_warn_badge": "⚠ {n}×WW",
    # Slot-Vergleich (Change 10)
    "btn_slot_compare": "🔀 Slot-Vergleich",
    "slot_compare_title": "Slot-Vergleich — Live ΔE-Auswirkung",
    "slot_compare_slot": "Zu vergleichender Slot:",
    "slot_compare_alt": "Alternatives Filament:",
    "slot_compare_col_vid": "V-Kopf",
    "slot_compare_col_cur": "Aktuell ΔE",
    "slot_compare_col_new": "Neu ΔE",
    "slot_compare_col_delta": "Δ",
    # Werkzeugwechsel-Kosten (Change 11)
    "tc_cost_per_kg": "Filamentpreis (€/kg):",
    "tc_density": "Dichte (g/cm³):",
    "tc_purge_cost": "Purge-Kosten gesamt: ~{cost:.2f}€",
    "tc_purge_layers": "Purge entspricht: ~{n} Druckschichten",
    # TD-Schätzung (Change 8)
    "btn_estimate_td": "TD schätzen",
    # DnD-Hint (Change 5)
    "dnd_hint": "← .3mf oder .u1proj hierher ziehen",
    # Auto-found label (Change 5 UI)
    "auto_found": "Auto: Länge {n} gefunden",
    # Status bar strings
    "status_ready": "Bereit",
    "status_calculated": "Berechnet — ΔE {de:.1f} — Sequenz: {seq}",
    "status_added": "V{vid} hinzugefügt",
    "status_exported": "Exportiert: {f}",
    "status_3mf": "3MF geladen — {n} Farben gefunden",
    # Slot undo button
    "btn_undo_slot": "↩ Slot",
    # Virt sort/filter labels
    "virt_sort_added": "Hinzugefügt",
    "virt_sort_de_asc": "ΔE ↑",
    "virt_sort_de_desc": "ΔE ↓",
    "virt_sort_label": "Label A-Z",
    "virt_filter_placeholder": "Filtern…",
    # Click-to-copy hint
    "click_to_copy": "(klick zum Kopieren)",
    # 3MF Farb-Wizard
    "wizard_btn": "🧙 3MF Wizard",
    "wizard_title": "3MF Farb-Wizard",
    "wizard_step1": "Schritt 1 / 3 — 3MF Datei laden",
    "wizard_step2": "Schritt 2 / 3 — Beste 4 Filamente suchen",
    "wizard_step3": "Schritt 3 / 3 — Ergebnis",
    "wizard_load_btn": "📂 3MF Datei öffnen",
    "wizard_no_file": "Keine Datei geladen",
    "wizard_colors_found": "{n} Farbe(n) im Modell gefunden",
    "wizard_next": "Weiter →",
    "wizard_info": "Durchsuche {n_lib} Filamente nach bester Kombination für {n_col} Zielfarben.",
    "wizard_start": "Optimierung starten",
    "wizard_checking": "Prüfe Kombination {i}/{total}…",
    "wizard_avg_de": "Durchschnittliche ΔE: {de:.1f}",
    "wizard_apply": "✅ Als T1–T4 übernehmen",
    "wizard_add_virtual": "Virtuelle Köpfe für alle Modellfarben berechnen",
    "wizard_close": "Schließen",
    "wizard_applied": "Beste 4 Filamente als T1–T4 gesetzt.",
    "wizard_coverage": "Farb-Abdeckung",
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
    "lh_hint": "⚠ 0.08 mm recommended — color layers become invisible",
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
    "sec2_desc": "Each virtual head = FullSpectrum sequence (1–48 layers) from T1–T4  ·  max. {max_v}",
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
    "tip_add_virtual": "Add as virtual print head V5-V24 (Ctrl+Enter)",
    "tip_copy": "Copy sequence to clipboard (Ctrl+C)",
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
    "orca_fs_hint": "⚠ FullSpectrum slicer detected: Export T1–T4 only! V5+ are auto-generated by FullSpectrum as 'Mixed Filaments' — importing them additionally creates duplicate physical slots.",
    # Project
    "btn_proj_save": "💾 Save Project",
    "btn_proj_load": "📂 Load Project",
    "proj_saved": "Project saved:\n{path}",
    "proj_loaded": "Project loaded:\n{path}",
    "proj_filetypes": "U1 Project",
    "proj_err": "Load error:\n{e}",
    # TD calibration
    "btn_td_cal": "🔬 Calibrate TD",
    "td_cal_title": "TD Calibration Assistant",
    "td_cal_desc": "Print T{a} and T{b} alternating (e.g. 5+5 layers) and measure the mix color.",
    "td_cal_measured": "Measured hex color of the test print:",
    "td_cal_result": "Recommended TD values:  T{a} = {ta:.1f}    T{b} = {tb:.1f}",
    "td_cal_apply": "Apply",
    "td_cal_slot": "Slot:",
    # Top-3
    "top3_title": "Top-3 Sequences",
    "top3_rank": "#{r}",
    "top3_add": "+ Add",
    # Slot optimizer
    "btn_slot_opt": "🎯 Slot Optimizer",
    "slot_opt_title": "Slot Optimizer — Find Best 4 Filaments",
    "slot_opt_desc": "Finds which 4 filaments from the library cover the largest color gamut.",
    "slot_opt_targets": "Target colors (from virtual heads or manual):",
    "slot_opt_use_virtual": "Use virtual heads as targets",
    "slot_opt_running": "Calculating … {i}/{total}",
    "slot_opt_done": "Done! Best set found.",
    "slot_opt_apply": "Apply as T1–T4",
    "slot_opt_result": "Best set (avg ΔE {de:.1f}):",
    # Palette import
    "btn_palette": "🖼 Palette from Image",
    "tip_palette": "Extract dominant colors from image and add as batch",
    "palette_title": "Import Palette from Image",
    "palette_colors": "Colors:",
    "palette_btn_add": "Calculate all & add",
    # 3MF write-back
    "btn_3mf_write": "✏️ Write 3MF",
    "tip_3mf_write": "Write virtual head colors into a 3MF file",
    "3mf_write_title": "Write Filament Colors to 3MF",
    "3mf_write_desc": "Select a 3MF file. Filament colors will be updated with the simulated colors of the virtual heads.",
    "3mf_write_ok": "3MF updated successfully:\n{path}",
    "3mf_write_err": "Write error:\n{e}",
    "remap_title": "3MF Extruder Remap",
    "remap_col_hdr": "Original extruders  →  Target (T1–T4 or virtual head)",
    "remap_keep": "(keep unchanged)",
    "remap_btn_write": "Write 3MF",
    # Tool changes
    "btn_tc_est": "🔄 Tool Changes",
    "tc_title": "Tool Change Estimator",
    "tc_layers": "Total print layers:",
    "tc_result": "Estimated tool changes: {n}",
    "tc_time": "Estimated extra time: ~{min:.0f} min  (at {sec}s/change)",
    "tc_purge": "Estimated purge waste: ~{g:.0f}g",
    # Color harmonies
    "btn_harmonies": "🎨 Harmonies",
    "tip_harmonies": "Complementary, triadic and analogous colors for the target",
    "harm_title": "Color Harmonies",
    "harm_complement": "Complement",
    "harm_triadic": "Triadic",
    "harm_analogous": "Analogous",
    "harm_split": "Split-Complement",
    "harm_add_all": "Calculate all & add",
    # Color model
    "model_label": "Color model:",
    "model_linear": "Additive — FullSpectrum compatible",
    "model_td": "TD-weighted (physical)",
    "model_subtractive": "Subtractive (pigment model)",
    "model_filamentmixer": "FilamentMixer (Pigment)",
    # Stripe risk
    "stripe_risk": "\u26a0 Stripe risk \u2014 sequence length {n} unfavorable",
    "stripe_ok": "\u2713 No stripe risk",
    # Multi-Gradient
    "btn_multi_gradient": "\U0001f308 Multi-Gradient",
    "multi_gradient_title": "Multi-Gradient Virtual Head",
    "multi_gradient_desc": "Create weighted gradient from all 4 slots",
    "multi_gradient_auto": "Auto-Balance",
    "multi_gradient_add": "Add as Virtual Head",
    # Copy all cadence values
    "btn_copy_all_cad": "📋 All Cadence Values",
    "copy_all_title": "All Dithering Values",
    "copy_all_btn": "Copy All to Clipboard",
    # Tabs
    "tab_tools": "Werkzeuge",
    # Gamut plot
    "btn_gamut_plot": "🎯 Gamut Plot",
    # Batch recalc
    "btn_recalc_all": "🔄 Recalc All",
    "recalc_all_done": "{n} virtual heads recalculated.",
    # ΔE overview
    "btn_de_overview": "📊 ΔE Overview",
    "de_overview_title": "ΔE Overview — All Virtual Heads",
    "de_overview_col_id": "ID",
    "de_overview_col_label": "Label",
    "de_overview_col_seq": "Sequence",
    "de_overview_col_de": "ΔE",
    "de_overview_col_quality": "Quality",
    "de_overview_col_tc": "TC/Cycle",
    # Color recipe
    "btn_recipe": "📜 Recipe",
    "recipe_title": "Color Recipe Export",
    "recipe_copy_btn": "Copy to Clipboard",
    # Multi-target optimizer
    "btn_multitarget": "🎯 Multi-Target",
    "mt_title": "Multi-Target Color Optimizer",
    "mt_desc": "Finds the best sequence that approximates multiple target colors simultaneously.",
    "mt_target": "Target {n}:",
    "mt_add_target": "+ Add Target",
    "mt_calc": "⚙  Optimize",
    "mt_result": "Best sequence: {seq}   avg ΔE {de:.1f}",
    "mt_add_btn": "➕ Add as Virtual Head",
    "mt_no_targets": "Please specify at least 2 target colors.",
    # Slot loaded status
    "slot_loaded": "Loaded",
    # Tool change warning per V-head
    "tc_warn_badge": "⚠ {n}×TC",
    # Slot comparison (Change 10)
    "btn_slot_compare": "🔀 Slot Compare",
    "slot_compare_title": "Slot Compare — Live ΔE Impact",
    "slot_compare_slot": "Slot to compare:",
    "slot_compare_alt": "Alternative filament:",
    "slot_compare_col_vid": "V-Head",
    "slot_compare_col_cur": "Current ΔE",
    "slot_compare_col_new": "New ΔE",
    "slot_compare_col_delta": "Δ",
    # Tool change costs (Change 11)
    "tc_cost_per_kg": "Filament price (€/kg):",
    "tc_density": "Density (g/cm³):",
    "tc_purge_cost": "Total purge cost: ~{cost:.2f}€",
    "tc_purge_layers": "Purge equals: ~{n} print layers",
    # TD estimation (Change 8)
    "btn_estimate_td": "Estimate TD",
    # DnD hint (Change 5)
    "dnd_hint": "← drop .3mf or .u1proj here",
    # Auto-found label (Change 5 UI)
    "auto_found": "Auto: length {n} found",
    # Status bar strings
    "status_ready": "Ready",
    "status_calculated": "Calculated — ΔE {de:.1f} — Sequence: {seq}",
    "status_added": "V{vid} added",
    "status_exported": "Exported: {f}",
    "status_3mf": "3MF loaded — {n} colors found",
    # Slot undo button
    "btn_undo_slot": "↩ Slot",
    # Virt sort/filter labels
    "virt_sort_added": "Added",
    "virt_sort_de_asc": "ΔE ↑",
    "virt_sort_de_desc": "ΔE ↓",
    "virt_sort_label": "Label A-Z",
    "virt_filter_placeholder": "Filter…",
    # Click-to-copy hint
    "click_to_copy": "(click to copy)",
    # 3MF Color Wizard
    "wizard_btn": "🧙 3MF Wizard",
    "wizard_title": "3MF Color Wizard",
    "wizard_step1": "Step 1 / 3 — Load 3MF File",
    "wizard_step2": "Step 2 / 3 — Find Best 4 Filaments",
    "wizard_step3": "Step 3 / 3 — Result",
    "wizard_load_btn": "📂 Open 3MF File",
    "wizard_no_file": "No file loaded",
    "wizard_colors_found": "{n} color(s) found in model",
    "wizard_next": "Next →",
    "wizard_info": "Searching {n_lib} filaments for best combination for {n_col} target colors.",
    "wizard_start": "Start Optimization",
    "wizard_checking": "Checking combination {i}/{total}…",
    "wizard_avg_de": "Average ΔE: {de:.1f}",
    "wizard_apply": "✅ Apply as T1–T4",
    "wizard_add_virtual": "Calculate virtual heads for all model colors",
    "wizard_close": "Close",
    "wizard_applied": "Best 4 filaments set as T1–T4.",
    "wizard_coverage": "Color Coverage",
},
}

_SLOT_SKIP = {"(leer)", "(empty)", "(manuell)", "(manual)"}

# ── KONSTANTEN ────────────────────────────────────────────────────────────────
MAX_SEQ_LEN     = 48
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

# ── EINGEBAUTE SLOT-PRESETS ────────────────────────────────────────────────────
# Wissenschaftlich optimierte Kombis für maximale Farbgamut (FullSpectrum additiv)
# Werden immer angezeigt (★-Präfix), können nicht überschrieben werden.
# Interne Keys sind sprachunabhängig — Anzeigenamen kommen aus BUILTIN_PRESET_LABELS.
BUILTIN_PRESETS = {
    "★ CMYK": [
        {"brand": "Bambu Lab Basic", "filament": "Cyan",         "hex": "#0086D6", "td": 5.0},
        {"brand": "Bambu Lab Basic", "filament": "Magenta",      "hex": "#EC008C", "td": 8.0},
        {"brand": "Bambu Lab Basic", "filament": "Lemon Yellow", "hex": "#FFF176", "td": 7.5},
        {"brand": "Bambu Lab Basic", "filament": "Black",        "hex": "#101012", "td": 0.3},
    ],
    "★ RGB+W": [
        {"brand": "Bambu Lab Basic", "filament": "Vivid Red",    "hex": "#E53935", "td": 3.5},
        {"brand": "Bambu Lab Basic", "filament": "Grass Green",  "hex": "#43A047", "td": 4.5},
        {"brand": "Bambu Lab Basic", "filament": "Cobalt Blue",  "hex": "#1565C0", "td": 3.5},
        {"brand": "Bambu Lab Basic", "filament": "Jade White",   "hex": "#F5F5F3", "td": 8.5},
    ],
    "★ Primary": [
        {"brand": "Bambu Lab Basic", "filament": "Jade White",   "hex": "#F5F5F3", "td": 8.5},
        {"brand": "Bambu Lab Basic", "filament": "Yellow",       "hex": "#FCE300", "td": 6.5},
        {"brand": "Bambu Lab Basic", "filament": "Vivid Red",    "hex": "#E53935", "td": 3.5},
        {"brand": "Bambu Lab Basic", "filament": "Cobalt Blue",  "hex": "#1565C0", "td": 3.5},
    ],
    "★ Warm": [
        {"brand": "Bambu Lab Basic", "filament": "Jade White",   "hex": "#F5F5F3", "td": 8.5},
        {"brand": "Bambu Lab Basic", "filament": "Lemon Yellow", "hex": "#FFF176", "td": 7.5},
        {"brand": "Bambu Lab Basic", "filament": "Tangerine",    "hex": "#FF7043", "td": 5.5},
        {"brand": "Bambu Lab Basic", "filament": "Flame Red",    "hex": "#D32F2F", "td": 3.5},
    ],
    "★ Cool": [
        {"brand": "Bambu Lab Basic", "filament": "Jade White",   "hex": "#F5F5F3", "td": 8.5},
        {"brand": "Bambu Lab Basic", "filament": "Mint",         "hex": "#A5D6A7", "td": 7.0},
        {"brand": "Bambu Lab Basic", "filament": "Cyan",         "hex": "#0086D6", "td": 5.0},
        {"brand": "Bambu Lab Basic", "filament": "Violet",       "hex": "#4527A0", "td": 3.0},
    ],
    "★ Snapmaker": [
        {"brand": "Snapmaker PLA",   "filament": "White",        "hex": "#F8F8F8", "td": 8.5},
        {"brand": "Snapmaker PLA",   "filament": "Yellow",       "hex": "#FFD700", "td": 6.5},
        {"brand": "Snapmaker PLA",   "filament": "Red",          "hex": "#CC2020", "td": 3.5},
        {"brand": "Snapmaker PLA",   "filament": "Blue",         "hex": "#0047AB", "td": 3.5},
    ],
}

# Übersetzung der Anzeigenamen für BUILTIN_PRESETS (interner Key → DE/EN Label)
BUILTIN_PRESET_LABELS = {
    "★ CMYK":      {"de": "★ CMYK — Max. Gamut",       "en": "★ CMYK — Max. Gamut"},
    "★ RGB+W":     {"de": "★ RGB + Weiß — Additiv",    "en": "★ RGB + White — Additive"},
    "★ Primary":   {"de": "★ Primärfarben Classic",     "en": "★ Primary Colors Classic"},
    "★ Warm":      {"de": "★ Warm-Palette",             "en": "★ Warm Tones"},
    "★ Cool":      {"de": "★ Kalt-Palette",             "en": "★ Cool Tones"},
    "★ Snapmaker": {"de": "★ Snapmaker Standard",       "en": "★ Snapmaker Standard"},
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

def find_best_4_filaments(target_labs, library_fils, progress_cb=None):
    """Find best 4 filaments from library to cover target_labs (list of Lab tuples).

    Algorithm: k-medoids inspired greedy search
    1. For each target color, find top-20 closest library filaments
    2. Build candidate pool from union of all top-20 lists
    3. Try all C(candidate_pool, 4) combinations (usually <5000)
    4. For each combo: score = avg of min_delta_e per target across 4 slots
    5. Return best combo

    library_fils: list of {"name": str, "hex": str, "td": float, "brand": str, "lab": tuple}
    Returns: (best_4_list, avg_de, scores_per_target)
    """
    if not target_labs or not library_fils:
        return [], 99.0, []

    # Step 1: for each target, find top-20 closest library filaments
    candidate_ids = set()
    for t_lab in target_labs:
        distances = [(i, delta_e(t_lab, f["lab"])) for i, f in enumerate(library_fils)]
        distances.sort(key=lambda x: x[1])
        for idx, _ in distances[:20]:
            candidate_ids.add(idx)

    candidates = [library_fils[i] for i in sorted(candidate_ids)]

    # Step 2: try all C(candidates, 4) combos
    from itertools import combinations as _combs
    best_combo = None
    best_score = float("inf")
    combos = list(_combs(range(len(candidates)), 4))

    for ci, combo_idxs in enumerate(combos):
        if progress_cb and ci % 100 == 0:
            progress_cb(ci, len(combos))
        combo = [candidates[i] for i in combo_idxs]
        combo_labs = [f["lab"] for f in combo]
        total_de = 0.0
        for t_lab in target_labs:
            min_de = min(delta_e(t_lab, c_lab) for c_lab in combo_labs)
            total_de += min_de
        avg_de = total_de / len(target_labs)
        if avg_de < best_score:
            best_score = avg_de
            best_combo = combo

    if best_combo is None:
        best_combo = candidates[:4] if len(candidates) >= 4 else candidates

    # Compute per-target scores for display
    scores = []
    if best_combo:
        combo_labs = [f["lab"] for f in best_combo]
        for t_lab in target_labs:
            min_de = min(delta_e(t_lab, c_lab) for c_lab in combo_labs)
            scores.append(min_de)

    return best_combo, best_score, scores


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
    """Cadence Heights matching FullSpectrum slicer v0.92+ logic.

    Slicer formula: minority_side=1, majority=max(1, round(major_pct/minor_pct))
    Reduces GCD so cycle is minimal. Returns {filament_id: cadence_mm}.
    """
    if not sequence:
        return {}
    # Count occurrences per filament
    counts = {}
    for fid in sequence:
        fid_int = int(fid)
        counts[fid_int] = counts.get(fid_int, 0) + 1
    if len(counts) == 1:
        fid = list(counts.keys())[0]
        return {fid: round(len(sequence) * layer_height, 3)}

    total = sum(counts.values())
    # Sort: minority first, majority last
    sorted_ids = sorted(counts.keys(), key=lambda k: counts[k])

    # Slicer formula: minority anchors to 1, majority scales
    minority_count = counts[sorted_ids[0]]
    result = {}
    for fid in sorted_ids:
        pct = counts[fid] / total
        minority_pct = minority_count / total
        if fid == sorted_ids[0]:  # minority
            layers = 1
        else:
            layers = max(1, round(pct / minority_pct))
        result[fid] = round(layers * layer_height, 3)
    return result

def seq_filament_count(sequence):
    return len(set(sequence))

def check_stripe_risk(sequence):
    """Check if a sequence has stripe risk based on FullSpectrum phase-shift formula.

    Slicer uses phase_step = (cycle / 2 + 1) to avoid striping.
    Risk exists when cycle length and phase_step share common factors > 1.
    Returns (risk: bool, message: str)
    """
    if not sequence:
        return False, ""
    n = len(sequence)
    if n <= 1:
        return False, ""
    # Unique filament count
    unique = len(set(str(x) for x in sequence))
    if unique == 1:
        return False, ""
    phase_step = (n // 2) + 1
    gcd = math.gcd(n, phase_step)
    if gcd > 1:
        return True, f"\u26a0 Streifenrisiko (Zyklus {n}, Phase {phase_step}, GCD={gcd})"
    return False, f"\u2713 Kein Streifenrisiko (Zyklus {n}, Phase {phase_step})"

def filament_mixer_lerp(r1, g1, b1, r2, g2, b2, t):
    """Polynomial pigment blending approximating Mixbox behavior.
    Degree-4 polynomial regression. t=0 -> color1, t=1 -> color2.
    Returns (r, g, b) as 0-255 integers.
    """
    # Normalize to 0-1
    r1f, g1f, b1f = r1/255, g1/255, b1/255
    r2f, g2f, b2f = r2/255, g2/255, b2/255

    # Simple but physically-motivated pigment model:
    # Convert to "pigment concentration" space (roughly sqrt for scattering)
    # then blend, then convert back
    def to_pigment(c):
        return math.sqrt(max(0.0, c))
    def from_pigment(c):
        return min(1.0, c * c)

    # Blend in pigment space
    rp = to_pigment(r1f) * (1-t) + to_pigment(r2f) * t
    gp = to_pigment(g1f) * (1-t) + to_pigment(g2f) * t
    bp = to_pigment(b1f) * (1-t) + to_pigment(b2f) * t

    r = int(round(from_pigment(rp) * 255))
    g = int(round(from_pigment(gp) * 255))
    b = int(round(from_pigment(bp) * 255))
    return (max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b)))

def build_weighted_gradient_sequence(weights, max_len=48):
    """Build a dithering sequence from weighted filament list.
    weights: list of (filament_id, weight_pct) tuples, weights sum to 100.
    Uses Bresenham-style distribution like the FullSpectrum slicer.
    Returns sequence as list of filament IDs.
    """
    if not weights:
        return []
    total = sum(w for _, w in weights)
    if total == 0:
        return [weights[0][0]]

    # Normalize to max_len slots
    slots = []
    for fid, w in weights:
        exact = (w / total) * max_len
        n = int(exact)
        remainder_new = exact - n
        slots.append((fid, n, remainder_new))

    # Distribute remainder slots by largest fractional part
    n_total = sum(n for _, n, _ in slots)
    deficit = max_len - n_total
    sorted_by_rem = sorted(slots, key=lambda x: x[2], reverse=True)
    final = {fid: n for fid, n, _ in slots}
    for i in range(deficit):
        fid = sorted_by_rem[i % len(sorted_by_rem)][0]
        final[fid] = final.get(fid, 0) + 1

    # Interleave using Bresenham
    sequence = []
    accumulators = {fid: 0.0 for fid, _ in weights}
    for step in range(max_len):
        best_fid = None
        best_acc = -1
        for fid, w in weights:
            accumulators[fid] += (w / total)
            if accumulators[fid] > best_acc:
                best_acc = accumulators[fid]
                best_fid = fid
        sequence.append(best_fid)
        accumulators[best_fid] -= 1.0

    return sequence

def compute_layer_schedule(sequence, n_layers=12):
    """Simulate which filament is active at each layer index using slicer formula.
    sequence: list/string of filament IDs forming one cycle.
    Returns list of filament_ids for layers 0..n_layers-1.
    """
    cycle = len(sequence)
    if cycle == 0:
        return []
    seq = [int(s) for s in sequence]
    result = []
    for layer_idx in range(n_layers):
        pos = layer_idx % cycle
        result.append(seq[pos])
    return result

def estimate_td(hex_color):
    """Estimate TD from hex brightness. Bright/saturated = higher TD (more transparent)."""
    try:
        r, g, b = hex_to_rgb(hex_color)
        # Perceived brightness 0-1
        brightness = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        # Very dark colors (black): TD ~0.3; white: TD ~8.5
        return round(0.3 + brightness * 8.2, 1)
    except Exception:
        return 5.0


def de_label_text(de, lang="de"):
    s = STRINGS[lang]
    q = s["de_good"] if de < DE_GOOD else s["de_ok"] if de < DE_OK else s["de_far"]
    return f"ΔE {de:.1f}  {q}"

# ── 3MF FARBEXTRAKTION ────────────────────────────────────────────────────────

HEX_RE = re.compile(r'#([0-9A-Fa-f]{6})\b')

def parse_3mf_colors(filepath):
    found = set()
    try:
        import xml.etree.ElementTree as _ET3MF
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
                # Generic hex extraction from all file types
                for m in HEX_RE.finditer(content):
                    found.add('#' + m.group(1).upper())
                # XML-based extraction for .model files (PrusaSlicer p:color, BambuStudio m:colorgroup)
                if name.endswith('.model') or name.endswith('.xml'):
                    try:
                        root = _ET3MF.fromstring(content)
                        # Walk all elements looking for color-related attributes
                        for elem in root.iter():
                            # p:color attribute (PrusaSlicer)
                            for attr_name, attr_val in elem.attrib.items():
                                lname = attr_name.lower()
                                if 'color' in lname and isinstance(attr_val, str):
                                    m = HEX_RE.search(attr_val.strip())
                                    if m: found.add('#' + m.group(1).upper())
                            # p:colorgroup / m:colorgroup children
                            tag_local = elem.tag.split('}')[-1].lower() if '}' in elem.tag else elem.tag.lower()
                            if tag_local in ('colorgroup', 'color', 'basematerials', 'basematerial'):
                                for child in elem:
                                    for attr_val in child.attrib.values():
                                        m = HEX_RE.search(str(attr_val).strip())
                                        if m: found.add('#' + m.group(1).upper())
                                # Also check element's own text/attribs
                                for attr_val in elem.attrib.values():
                                    m = HEX_RE.search(str(attr_val).strip())
                                    if m: found.add('#' + m.group(1).upper())
                    except Exception:
                        pass
    except Exception as e:
        return [], str(e)
    trivial = {"#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF",
               "#AAAAAA", "#333333", "#CCCCCC"}
    meaningful = [c for c in found if c not in trivial]
    raw = meaningful if meaningful else list(found)
    # Deduplicate: remove near-duplicates with ΔE < 3
    deduped = []
    deduped_labs = []
    for hex_c in raw:
        try:
            lab = rgb_to_lab(hex_to_rgb(hex_c))
        except Exception:
            deduped.append(hex_c)
            continue
        if not any(delta_e(lab, el) < 3.0 for el in deduped_labs):
            deduped.append(hex_c)
            deduped_labs.append(lab)
    return deduped, None


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
        self._building_ui  = True   # suppress expensive redraws during startup
        self._gamut_job    = None   # debounce handle for gamut strip
        self.history       = []
        self.presets       = {}
        self.virtual_fils  = []   # list of virtual filament dicts
        self.virtual_undo  = []   # undo stack for virtual heads
        self._slot_undo_stack = []  # undo stack for slot changes
        self.last_result   = {}   # last calc() result for "hinzufügen"
        self._max_virtual  = int(self.settings.get("max_virtual", MAX_VIRTUAL))
        self.favorites     = []
        self._3mf_win      = None   # singleton for 3MF assistant window
        self._3mf_wizard_win = None  # singleton for 3MF Wizard
        self._recent_colors = []    # recent target colors (max 10)
        self.load_db()
        self.load_presets()
        self._load_favorites()
        self.setup_ui()
        self._building_ui = False   # UI fully built — allow redraws now
        # Einstellungen wiederherstellen
        lh = self.settings.get("layer_height", "0.08")
        self.layer_height_entry.delete(0, "end")
        self.layer_height_entry.insert(0, lh)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        # Keyboard shortcuts
        self.bind("<Control-z>", lambda e: self.undo_last())
        self.bind("<Control-Return>", lambda e: self.add_to_virtual())
        self.bind("<Control-1>", lambda e: self._switch_tab(0))
        self.bind("<Control-2>", lambda e: self._switch_tab(1))
        self.bind("<Control-3>", lambda e: self._switch_tab(2))
        # Restore extended settings
        self._restore_extended_settings()
        # Single gamut update after everything is ready
        self.after(400, self._update_gamut_strip)
        # Drag & Drop support
        if _HAS_DND:
            try:
                self.drop_target_register(DND_FILES)
                self.dnd_bind('<<Drop>>', self._on_dnd_drop)
            except Exception:
                pass

    def _on_close(self):
        self._save_settings()
        self.destroy()

    def _set_status(self, msg, duration=0):
        """Update status bar text; auto-reset after duration ms if > 0."""
        if hasattr(self, "_status_bar"):
            self._status_bar.configure(text=msg)
            if duration > 0:
                self.after(duration, lambda: self._status_bar.configure(
                    text=self.t("status_ready")) if hasattr(self, "_status_bar") else None)

    def _toggle_slot(self, i):
        """Accordion toggle for slot i."""
        if not hasattr(self, "_slot_expanded"):
            return
        self._slot_expanded[i] = not self._slot_expanded[i]
        expanded = self._slot_expanded[i]
        body = self.slots[i].get("_body")
        btn = self._slot_toggle_btns[i] if hasattr(self, "_slot_toggle_btns") else None
        if body:
            if expanded:
                body.pack(fill="x")
            else:
                body.pack_forget()
        if btn:
            sym = "▼" if expanded else "▶"
            # Try to show the current hex color in the header when collapsed
            hex_val = ""
            try:
                hex_val = self.slots[i]["hex"].get().strip()
            except Exception:
                pass
            if not expanded and hex_val:
                btn.configure(text=f"{sym} T{i+1}  {hex_val}")
            else:
                btn.configure(text=f"{sym} T{i+1} — WERKZEUG {i+1}")

    def _on_dnd_drop(self, event):
        """Handle drag-and-drop of .3mf or .u1proj files onto the main window."""
        path = event.data.strip().strip('{}')
        if path.lower().endswith('.3mf'):
            self._open_3mf_with_path(path)
        elif path.lower().endswith('.u1proj'):
            self._load_project_from_path(path)

    def _open_3mf_with_path(self, path):
        """Open 3MF assistant with a given file path (used by DnD)."""
        colors, err = parse_3mf_colors(path)
        if not colors:
            messagebox.showinfo(self.t("dlg_3mf_title"),
                f"{self.t('dlg_3mf_no_colors_fallback') if not err else err}"); return
        # Reuse open_3mf_assistant logic but skip file dialog
        # Store path temporarily and call open with pre-loaded data
        self._dnd_3mf_path = path
        self._dnd_3mf_colors = colors
        self._open_3mf_with_colors(path, colors)

    def _open_3mf_with_colors(self, path, colors):
        """Open 3MF assistant window with pre-loaded colors."""
        if self._3mf_win is not None and self._3mf_win.winfo_exists():
            self._3mf_win.focus_force()
            self._3mf_win.lift()
            return
        win = ctk.CTkToplevel(self)
        self._3mf_win = win
        win.protocol("WM_DELETE_WINDOW", lambda: (setattr(self, "_3mf_win", None), win.destroy()))
        win.title(f"{self.t('dlg_3mf_title')} — {os.path.basename(path)}")
        win.geometry("860x680")
        win.grab_set()
        ctk.CTkLabel(win,
                     text=self.t("3mf_analysis_title", n=len(colors)),
                     font=("Segoe UI", 15, "bold"), text_color="#38bdf8").pack(pady=(16, 2))
        ctk.CTkLabel(win,
                     text=self.t("3mf_basis",
                          t1=self.slots[0]['hex'].get(), t2=self.slots[1]['hex'].get(),
                          t3=self.slots[2]['hex'].get(), t4=self.slots[3]['hex'].get()),
                     font=("Segoe UI", 10), text_color="#64748b").pack(pady=(0, 6))
        opt_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(win, text=self.t("3mf_optimizer"), variable=opt_var,
                        font=("Segoe UI", 11)).pack(pady=(0, 4))
        prog_label = ctk.CTkLabel(win, text=self.t("3mf_ready"),
                                   font=("Segoe UI", 11), text_color="#94a3b8")
        prog_label.pack(pady=(0, 6))
        hdr = ctk.CTkFrame(win, fg_color="#1e293b", corner_radius=6)
        hdr.pack(fill="x", padx=14, pady=(0, 4))
        hdr.grid_columnconfigure(4, weight=1)
        for col, txt in enumerate(["#", self.t("3mf_col_target"), self.t("3mf_col_seq"),
                                    self.t("3mf_col_sim"), self.t("3mf_col_quality"), "VID"]):
            ctk.CTkLabel(hdr, text=txt, font=("Segoe UI", 10, "bold"),
                         text_color="#64748b").grid(row=0, column=col, padx=10, pady=6, sticky="w")
        scroll = ctk.CTkScrollableFrame(win, height=380, fg_color="#0f172a")
        scroll.pack(fill="both", expand=True, padx=14, pady=4)
        scroll.grid_columnconfigure(4, weight=1)
        results = [None] * len(colors)
        vid_vars = []
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
                                 text_color=de_color(r["de"])).grid(row=0, column=4, padx=6, sticky="w")
                else:
                    ctk.CTkLabel(row, text=self.t("3mf_not_calc"),
                                 text_color="#334155").grid(row=0, column=2, columnspan=3, padx=6, sticky="w")
                sv = ctk.BooleanVar(value=True if r else False)
                vid_vars.append(sv)
                ctk.CTkCheckBox(row, text=self.t("3mf_include"), variable=sv,
                                font=("Segoe UI", 9), width=100).grid(row=0, column=5, padx=(4, 10))
        render_rows()
        def run_all():
            opt = opt_var.get()
            free = self._max_virtual - len(self.virtual_fils)
            to_calc = min(len(colors), free) if free < len(colors) else len(colors)
            for idx in range(min(len(colors), self._max_virtual)):
                prog_label.configure(text=self.t("3mf_progress", i=idx+1, total=to_calc, c=colors[idx]))
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
            setattr(self, "_3mf_win", None)
            messagebox.showinfo(self.t("dlg_3mf_title"), self.t("dlg_3mf_added", n=added))
        btm = ctk.CTkFrame(win, fg_color="transparent")
        btm.pack(fill="x", padx=14, pady=(6, 14))
        ctk.CTkButton(btm, text=self.t("3mf_btn_calc"), fg_color="#2563eb",
                      command=run_all, height=42, font=("Segoe UI", 12, "bold")).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btm, text=self.t("3mf_btn_apply"), fg_color="#15803d",
                      command=apply_selected, height=42, font=("Segoe UI", 12, "bold")).pack(side="left")
        ctk.CTkButton(btm, text=self.t("3mf_btn_cancel"), fg_color="#374151",
                      command=win.destroy, height=42).pack(side="right")

    def _load_project_from_path(self, path):
        """Load a .u1proj file from a given path (used by DnD)."""
        try:
            with open(path, encoding="utf-8") as f:
                project = json.load(f)
        except Exception as e:
            messagebox.showerror(self.t("dlg_error"), self.t("proj_err", e=e)); return
        self._save_slot_snapshot()
        if hasattr(self, "layer_height_entry") and "layer_height" in project:
            self.layer_height_entry.delete(0, "end")
            self.layer_height_entry.insert(0, project["layer_height"])
        for i, s_data in enumerate(project.get("slots", [])[:4]):
            s = self.slots[i]
            if s_data.get("brand") and s_data["brand"] in self.library:
                s["brand"].set(s_data["brand"])
                self.update_menu(i)
            if s_data.get("color"):
                try: s["color"].set(s_data["color"])
                except: pass
            if s_data.get("hex"):
                s["hex"].delete(0, "end")
                s["hex"].insert(0, s_data["hex"])
                try: s["preview"].configure(fg_color=s_data["hex"])
                except: pass
            if s_data.get("td"):
                s["td"].delete(0, "end")
                s["td"].insert(0, s_data["td"])
        self.virtual_fils = project.get("virtual_fils", [])
        if "recent_colors" in project:
            self._recent_colors = project["recent_colors"]
            self._update_recent_swatches()
        self._refresh_virtual_grid()
        self._schedule_gamut_update()
        messagebox.showinfo(self.t("dlg_saved"), self.t("proj_loaded", path=path))

    # ── FAVORITEN ──────────────────────────────────────────────────────────────

    def _load_favorites(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "favorites.json")
        try:
            with open(path) as f:
                self.favorites = json.load(f)
        except Exception:
            self.favorites = []

    def _save_favorites(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "favorites.json")
        try:
            with open(path, "w") as f:
                json.dump(self.favorites, f)
        except Exception:
            pass

    def _add_to_favorites(self):
        raw = self.hex_target_entry.get().strip() if hasattr(self, "hex_target_entry") else ""
        hex_val = raw.lstrip("#")
        if len(hex_val) == 6:
            entry = f"#{hex_val.upper()}"
            if entry not in self.favorites:
                self.favorites.append(entry)
                self._save_favorites()

    def open_favorites(self):
        win = ctk.CTkToplevel(self)
        win.title("Favoriten")
        win.geometry("280x400")
        sf = ctk.CTkScrollableFrame(win)
        sf.pack(fill="both", expand=True, padx=8, pady=8)
        for fav in list(self.favorites):
            row = ctk.CTkFrame(sf, fg_color="transparent")
            row.pack(fill="x", pady=2)
            try:
                swatch = ctk.CTkLabel(row, text="  ", width=30, fg_color=fav, corner_radius=4)
                swatch.pack(side="left", padx=4)
            except Exception:
                pass
            lbl = ctk.CTkLabel(row, text=fav, cursor="hand2")
            lbl.pack(side="left", padx=4)
            lbl.bind("<Button-1>", lambda e, h=fav: (
                self._apply_target(h),
                win.lift()))
            del_btn = ctk.CTkButton(row, text="✕", width=28, height=24,
                command=lambda h=fav, r=row: self._del_favorite(h, r))
            del_btn.pack(side="right", padx=4)

    def _del_favorite(self, hex_val, row_widget):
        if hex_val in self.favorites:
            self.favorites.remove(hex_val)
            self._save_favorites()
            row_widget.destroy()

    def t(self, key, **kwargs):
        s = STRINGS[self.lang].get(key, STRINGS["de"].get(key, key))
        return s.format(**kwargs) if kwargs else s

    def tip(self, widget, key):
        """Tooltip hinzufügen wenn CTkToolTip verfügbar."""
        if _HAS_TOOLTIP:
            _Tip(widget, message=self.t(key), delay=200)

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
        self.settings["layer_height"] = self.layer_height_entry.get() if hasattr(self, "layer_height_entry") else "0.08"
        self.settings["geometry"]     = self.geometry()
        self.settings["last_color"]   = self.hex_target_entry.get() if hasattr(self, "hex_target_entry") else ""
        self.settings["color_model"]  = self.color_model_var.get() if hasattr(self, "color_model_var") else "linear"
        try:
            tab_list = list(self.tabs._tab_dict.keys())
            current = self.tabs.get()
            self.settings["last_tab"] = tab_list.index(current)
        except Exception:
            self.settings["last_tab"] = 0
        try:
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=2)
        except IOError:
            pass

    def _restore_extended_settings(self):
        """Restore geometry, last color, color model, last tab after UI is built."""
        try:
            if self.settings.get("geometry"):
                self.geometry(self.settings["geometry"])
        except Exception:
            pass
        try:
            lc = self.settings.get("last_color", "")
            if lc and len(lc.lstrip("#")) == 6:
                raw = lc.lstrip("#")
                self.hex_target_entry.delete(0, "end")
                self.hex_target_entry.insert(0, lc)
                self._on_hex_live()
        except Exception:
            pass
        try:
            cm = self.settings.get("color_model", "")
            if cm and hasattr(self, "color_model_var"):
                self.color_model_var.set(cm)
        except Exception:
            pass
        try:
            idx = self.settings.get("last_tab")
            if idx is not None:
                tab_list = list(self.tabs._tab_dict.keys())
                if 0 <= idx < len(tab_list):
                    self.tabs.set(tab_list[idx])
        except Exception:
            pass

    def _switch_tab(self, idx):
        try:
            tab_list = list(self.tabs._tab_dict.keys())
            if 0 <= idx < len(tab_list):
                self.tabs.set(tab_list[idx])
        except Exception:
            pass

    def undo_last(self):
        """Ctrl+Z: undo last virtual head addition."""
        self.undo_virtual()

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

    def _save_slot_snapshot(self):
        """Push current slot state onto the slot undo stack (max 10 entries)."""
        if not hasattr(self, "slots"):
            return
        snap = []
        for s in self.slots:
            snap.append({
                "brand": s["brand"].get(),
                "color": s["color"].get(),
                "hex":   s["hex"].get(),
                "td":    s["td"].get(),
            })
        self._slot_undo_stack.append(snap)
        if len(self._slot_undo_stack) > 10:
            self._slot_undo_stack.pop(0)

    def _undo_slot(self):
        """Restore previous slot state from the slot undo stack."""
        if not self._slot_undo_stack:
            messagebox.showinfo(self.t("dlg_note"), self.t("undo_empty"))
            return
        snap = self._slot_undo_stack.pop()
        for i, v in enumerate(snap):
            s = self.slots[i]
            if v["brand"] in self.library:
                s["brand"].set(v["brand"])
                self.update_menu(i)
            col_val = v.get("color", "")
            if col_val:
                try:
                    s["color"].set(col_val)
                except Exception:
                    pass
            if v["hex"]:
                s["hex"].delete(0, "end"); s["hex"].insert(0, v["hex"])
                try:
                    s["preview"].configure(fg_color=v["hex"])
                except Exception:
                    pass
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

    # ── PROJEKT SPEICHERN/LADEN ────────────────────────────────────────────────

    def save_project(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".u1proj",
            filetypes=[(self.t("proj_filetypes"), "*.u1proj"), ("JSON", "*.json")],
            title=self.t("btn_proj_save"))
        if not path: return
        lh = self.layer_height_entry.get() if hasattr(self, "layer_height_entry") else "0.08"
        project = {
            "version": 1,
            "layer_height": lh,
            "slots": [
                {"brand": s["brand"].get(), "color": s["color"].get(),
                 "hex": s["hex"].get(), "td": s["td"].get()}
                for s in self.slots
            ],
            "virtual_fils": self.virtual_fils,
            "recent_colors": self._recent_colors,
        }
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(project, f, indent=2, ensure_ascii=False)
            messagebox.showinfo(self.t("dlg_saved"), self.t("proj_saved", path=path))
        except IOError as e:
            messagebox.showerror(self.t("dlg_error"), str(e))

    def load_project(self):
        path = filedialog.askopenfilename(
            filetypes=[(self.t("proj_filetypes"), "*.u1proj"), ("JSON", "*.json")],
            title=self.t("btn_proj_load"))
        if not path: return
        try:
            with open(path, encoding="utf-8") as f:
                project = json.load(f)
        except Exception as e:
            messagebox.showerror(self.t("dlg_error"), self.t("proj_err", e=e)); return
        self._save_slot_snapshot()
        # Schichthöhe
        if hasattr(self, "layer_height_entry") and "layer_height" in project:
            self.layer_height_entry.delete(0, "end")
            self.layer_height_entry.insert(0, project["layer_height"])
        # Slots
        for i, s_data in enumerate(project.get("slots", [])[:4]):
            s = self.slots[i]
            if s_data.get("brand") and s_data["brand"] in self.library:
                s["brand"].set(s_data["brand"])
                self.update_menu(i)
            if s_data.get("color"):
                try: s["color"].set(s_data["color"])
                except: pass
            if s_data.get("hex"):
                s["hex"].delete(0, "end")
                s["hex"].insert(0, s_data["hex"])
                try: s["preview"].configure(fg_color=s_data["hex"])
                except: pass
            if s_data.get("td"):
                s["td"].delete(0, "end")
                s["td"].insert(0, s_data["td"])
        # Virtuelle Köpfe
        self.virtual_fils = project.get("virtual_fils", [])
        # Ensure stable_id on loaded virtual heads
        for i, vf in enumerate(self.virtual_fils):
            if "stable_id" not in vf:
                vf["stable_id"] = vf.get("vid", 5 + i)
        # Recent colors
        if "recent_colors" in project:
            self._recent_colors = project["recent_colors"]
            self._update_recent_swatches()
        self._refresh_virtual_grid()
        self._schedule_gamut_update()
        messagebox.showinfo(self.t("dlg_saved"), self.t("proj_loaded", path=path))

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
        if name.strip().startswith("★"):
            messagebox.showwarning(self.t("dlg_saved"), "Namen mit ★ sind für eingebaute Presets reserviert.")
            return
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
        if not name: return
        # Displayed name → internal key (for built-in presets)
        entries = self.presets.get(name)
        if entries is None:
            for key, labels in BUILTIN_PRESET_LABELS.items():
                if name in labels.values():
                    entries = BUILTIN_PRESETS.get(key)
                    break
        if entries is None: return
        self._save_slot_snapshot()
        for i, entry in enumerate(entries[:4]):
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
        # Built-in presets first (★, übersetzt), dann User-Presets
        builtin = [BUILTIN_PRESET_LABELS[k][self.lang] for k in BUILTIN_PRESETS]
        names = builtin + list(self.presets.keys())
        if not names: names = [self.t("no_presets")]
        self.preset_dropdown.configure(values=names)
        if self.preset_var.get() not in names: self.preset_var.set(names[0])

    # ── UI-AUFBAU ──────────────────────────────────────────────────────────────

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)

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

        # Feature 13: Light/Dark mode toggle
        _cur_mode = ctk.get_appearance_mode()
        _appearance_icon = "🌙" if _cur_mode == "Dark" else "☀️"
        self.appearance_btn = ctk.CTkButton(lang_row, text=_appearance_icon, width=36, height=26,
                      fg_color="#374151", font=("Segoe UI", 12),
                      command=self._toggle_appearance)
        self.appearance_btn.pack(side="right", padx=(0, 4))

        ctk.CTkLabel(self.sidebar, text=self.t("phys_heads_title"),
                     font=("Segoe UI", 18, "bold"), text_color="#38bdf8").pack(pady=(8, 4))
        ctk.CTkLabel(self.sidebar,
                     text=self.t("phys_heads_desc"),
                     font=("Segoe UI", 10), text_color="#475569").pack(pady=(0, 10))

        self.slots = []
        self._slot_expanded = [True, False, False, False]
        self._slot_toggle_btns = []

        for i in range(4):
            # Outer wrapper: colored left-border strip + main frame
            wrapper = ctk.CTkFrame(self.sidebar, fg_color="transparent")
            wrapper.pack(fill="x", padx=12, pady=5)
            wrapper.grid_columnconfigure(1, weight=1)

            # Left color strip (Change 2)
            color_strip = ctk.CTkFrame(wrapper, width=4, fg_color="#334155", corner_radius=2)
            color_strip.grid(row=0, column=0, sticky="ns", padx=(0, 2))

            frame = ctk.CTkFrame(wrapper, border_width=1, border_color="#334155")
            frame.grid(row=0, column=1, sticky="ew")

            # Toggle header button (Change 1)
            toggle_symbol = "▼" if self._slot_expanded[i] else "▶"
            toggle_btn = ctk.CTkButton(
                frame, text=f"{toggle_symbol} T{i+1} — WERKZEUG {i+1}",
                fg_color="transparent", hover_color="#1e293b",
                anchor="w", font=("Segoe UI", 11, "bold"), text_color="#94a3b8",
                height=32, command=lambda idx=i: self._toggle_slot(idx))
            toggle_btn.pack(fill="x", padx=5, pady=(4, 0))
            self._slot_toggle_btns.append(toggle_btn)

            # Body frame (collapsible)
            body = ctk.CTkFrame(frame, fg_color="transparent")
            body.pack(fill="x")

            # Header row with action buttons inside body
            hdr = ctk.CTkFrame(body, fg_color="transparent")
            hdr.pack(fill="x", padx=5, pady=(2, 0))
            ctk.CTkButton(hdr, text="+", width=26, height=20, fg_color="#15803d",
                          command=lambda idx=i: self.add_filament(idx)).pack(side="right", padx=2)
            ctk.CTkButton(hdr, text="💾", width=26, height=20,
                          command=lambda idx=i: self.save_current(idx)).pack(side="right", padx=2)
            ctk.CTkButton(hdr, text="🔍", width=26, height=20, fg_color="#0e7490",
                          command=lambda idx=i: self.open_filament_search(idx)).pack(side="right", padx=2)

            brand = ctk.CTkOptionMenu(body, values=list(self.library.keys()),
                                       command=lambda x, idx=i: self.update_menu(idx))
            brand.pack(padx=10, pady=(3, 2), fill="x")
            color = ctk.CTkOptionMenu(body, values=["Lade..."],
                                       command=lambda x, idx=i: self.apply_f(idx))
            color.pack(padx=10, pady=2, fill="x")

            row = ctk.CTkFrame(body, fg_color="transparent")
            row.pack(fill="x", pady=(3, 8))
            hx = ctk.CTkEntry(row, width=88, placeholder_text="#RRGGBB")
            hx.pack(side="left", padx=(10, 2))
            # Change 2: larger preview (36×36)
            preview = ctk.CTkLabel(row, text="", width=36, height=36,
                                   fg_color="#1e293b", corner_radius=5)
            preview.pack(side="left", padx=(0, 2))
            ctk.CTkButton(row, text="🎨", width=28, height=26, fg_color="#334155",
                          command=lambda idx=i: self.pick_slot_color(idx)).pack(side="left", padx=(0, 6))
            ctk.CTkLabel(row, text="TD:", font=("Segoe UI", 10)).pack(side="right", padx=(0, 4))
            td = ctk.CTkEntry(row, width=54, placeholder_text="5.0")
            td.pack(side="right", padx=(0, 4))
            loaded_var = ctk.BooleanVar(value=True)
            loaded_sw = ctk.CTkSwitch(row, text=self.t("slot_loaded"), variable=loaded_var,
                                      width=80, font=("Segoe UI", 9),
                                      onvalue=True, offvalue=False)
            loaded_sw.pack(side="right", padx=(0, 6))

            self.slots.append({"brand": brand, "color": color,
                                "hex": hx, "td": td, "preview": preview,
                                "loaded": loaded_var,
                                "_body": body, "_strip": color_strip})
            self.update_menu(i)

            # Apply initial collapsed state (T2-T4 collapsed)
            if not self._slot_expanded[i]:
                body.pack_forget()


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
        pb.grid_columnconfigure((0, 1, 2), weight=1)
        ctk.CTkButton(pb, text=self.t("btn_load"), fg_color="#1e3a5f",
                      command=self.load_preset).grid(row=0, column=0, sticky="ew", padx=(0, 3))
        ctk.CTkButton(pb, text=self.t("btn_save"), fg_color="#1e3a5f",
                      command=self.save_preset).grid(row=0, column=1, sticky="ew", padx=(3, 3))
        ctk.CTkButton(pb, text=self.t("btn_undo_slot"), fg_color="#374151",
                      command=self._undo_slot, height=28,
                      font=("Segoe UI", 9)).grid(row=0, column=2, sticky="ew", padx=(3, 0))
        self._refresh_preset_dropdown()

        # Projekt speichern/laden
        proj_frame = ctk.CTkFrame(self.sidebar, fg_color="#0f172a", corner_radius=8)
        proj_frame.pack(fill="x", padx=12, pady=(4, 8))
        proj_frame.grid_columnconfigure((0,1), weight=1)
        ctk.CTkButton(proj_frame, text=self.t("btn_proj_save"), fg_color="#1e3a5f",
                      height=32, command=self.save_project).grid(
            row=0, column=0, sticky="ew", padx=(6,3), pady=6)
        ctk.CTkButton(proj_frame, text=self.t("btn_proj_load"), fg_color="#1e3a5f",
                      height=32, command=self.load_project).grid(
            row=0, column=1, sticky="ew", padx=(3,6), pady=6)


        # Schichthöhe — global für Cadence-Berechnung
        lh_frame = ctk.CTkFrame(self.sidebar, fg_color="#0f172a", corner_radius=8)
        lh_frame.pack(fill="x", padx=12, pady=(0, 15))
        lh_inner = ctk.CTkFrame(lh_frame, fg_color="transparent")
        lh_inner.pack(fill="x", padx=10, pady=8)
        ctk.CTkLabel(lh_inner, text=self.t("layer_height_label"),
                     font=("Segoe UI", 11)).pack(side="left")
        self.layer_height_entry = ctk.CTkEntry(lh_inner, width=60, placeholder_text="0.08")
        self.layer_height_entry.insert(0, "0.08")
        self.layer_height_entry.pack(side="left", padx=(6, 0))
        self.layer_height_entry.bind("<KeyRelease>",
            lambda e: self.after(400, self._refresh_virtual_grid))
        self.layer_height_entry.bind("<FocusOut>",
            lambda e: self._refresh_virtual_grid())
        ctk.CTkLabel(lh_inner, text=self.t("dithering_step"),
                     font=("Segoe UI", 9), text_color="#475569").pack(side="left", padx=6)
        ctk.CTkLabel(lh_inner, text=self.t("lh_hint"),
                     font=("Segoe UI", 8), text_color="#f59e0b").pack(side="left", padx=(8, 0))

        model_frame = ctk.CTkFrame(self.sidebar, fg_color="#0f172a", corner_radius=8)
        model_frame.pack(fill="x", padx=10, pady=(6,0))
        ctk.CTkLabel(model_frame, text=self.t("model_label"),
                     font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=8, pady=(6,2))
        self.color_model_var = ctk.StringVar(value="linear")
        for val, key in [("linear","model_linear"),("td","model_td"),("subtractive","model_subtractive"),("filamentmixer","model_filamentmixer")]:
            ctk.CTkRadioButton(model_frame, text=self.t(key),
                               variable=self.color_model_var, value=val,
                               command=self._on_model_change).pack(anchor="w", padx=16, pady=1)
        ctk.CTkLabel(model_frame, text="", height=4).pack()

        # ── HAUPTBEREICH: TABS ───────────────────────────────────────────────
        self.tabs = ctk.CTkTabview(
            self, corner_radius=8,
            fg_color="#0a0f1e",
            segmented_button_fg_color="#1e293b",
            segmented_button_selected_color="#2563eb",
            segmented_button_selected_hover_color="#1d4ed8",
            segmented_button_unselected_color="#1e293b",
            segmented_button_unselected_hover_color="#334155",
            text_color="#e2e8f0",
            text_color_disabled="#475569",
        )
        self.tabs.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=10)

        # ── STATUS BAR ────────────────────────────────────────────────────────
        self._status_bar = ctk.CTkLabel(self, text=self.t("status_ready"),
            font=("Segoe UI", 10), text_color="#64748b",
            anchor="w", height=24, fg_color="#0a1628")
        self._status_bar.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 2))

        _t1_name = "🎨  " + self.t("tab_calculator")
        _t2_name = "🔲  " + self.t("tab_virtual")
        _t3_name = "🛠  " + self.t("tab_tools")
        self.tabs.add(_t1_name)
        self.tabs.add(_t2_name)
        self.tabs.add(_t3_name)

        # scrollbare Container je Tab
        tab1 = ctk.CTkScrollableFrame(self.tabs.tab(_t1_name),
                                       fg_color="transparent", corner_radius=0)
        tab1.pack(fill="both", expand=True)
        tab1.grid_columnconfigure(0, weight=1)

        # Tab2: normaler Frame, damit das innere vgrid allein scrollt
        tab2 = ctk.CTkFrame(self.tabs.tab(_t2_name), fg_color="transparent", corner_radius=0)
        tab2.pack(fill="both", expand=True)
        tab2.grid_columnconfigure(0, weight=1)
        tab2.grid_rowconfigure(5, weight=1)

        tab3 = ctk.CTkScrollableFrame(self.tabs.tab(_t3_name),
                                       fg_color="transparent", corner_radius=0)
        tab3.pack(fill="both", expand=True)
        tab3.grid_columnconfigure(0, weight=1)

        # ────────────────────────────────────────────────────────────────────
        # TAB 1 — EINZELFARBEN-RECHNER
        # ────────────────────────────────────────────────────────────────────
        sec1 = ctk.CTkFrame(tab1, fg_color="#0f172a", corner_radius=10)
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
        self.hex_target_entry.bind("<KeyRelease>", self._on_hex_live)
        self.prev = ctk.CTkLabel(top, text="", width=46, height=46,
                                  fg_color="#1a1a1a", corner_radius=23)
        self.prev.grid(row=0, column=2)
        self.live_de_label = ctk.CTkLabel(top, text="", font=("Segoe UI", 11, "bold"),
                                           text_color="#94a3b8", width=80)
        self.live_de_label.grid(row=0, column=3, padx=(6, 0))

        rnd_btn = ctk.CTkButton(top, text=self.t("btn_random"), width=85, height=46,
                      fg_color="#374151", font=("Segoe UI", 11),
                      command=self.pick_random_color)
        rnd_btn.grid(row=0, column=4, padx=(4, 0))
        self.tip(rnd_btn, "tip_random")

        img_btn = ctk.CTkButton(top, text=self.t("btn_img_pick"), width=100, height=46,
                      fg_color="#374151", font=("Segoe UI", 11),
                      command=self.pick_color_from_image)
        img_btn.grid(row=0, column=5, padx=(4, 0))
        self.tip(img_btn, "tip_img_pick")

        fav_star_btn = ctk.CTkButton(top, text="⭐", width=40, height=46,
                      fg_color="#374151", font=("Segoe UI", 14),
                      command=self._add_to_favorites)
        fav_star_btn.grid(row=0, column=6, padx=(4, 0))

        fav_list_btn = ctk.CTkButton(top, text="📋", width=40, height=46,
                      fg_color="#374151", font=("Segoe UI", 14),
                      command=self.open_favorites)
        fav_list_btn.grid(row=0, column=7, padx=(2, 0))

        # Suggestion label (Feature 9)
        self.suggestion_label = ctk.CTkLabel(sec1, text="", text_color="gray",
                                              font=("Segoe UI", 11), wraplength=500)
        self.suggestion_label.grid(row=1, column=0, padx=20, pady=(2, 0), sticky="w")

        # Gamut-Warnung
        self.gamut_label = ctk.CTkLabel(
            sec1, text=self.t("gamut_warning"),
            fg_color="#7c2d12", corner_radius=6, text_color="#fcd34d",
            font=("Segoe UI", 10))

        # Farbinfo-Label (RGB + Lab)
        self.colorinfo_label = ctk.CTkLabel(sec1, text="",
                                             font=("Segoe UI", 9), text_color="#475569")
        self.colorinfo_label.grid(row=2, column=0, padx=20, pady=(0, 2), sticky="w")

        # Sequenz-Ergebnis + Kopier-Button
        res_outer = ctk.CTkFrame(sec1, fg_color="transparent")
        res_outer.grid(row=3, column=0, pady=(4, 0))
        res_frame = ctk.CTkFrame(res_outer, fg_color="transparent")
        res_frame.pack()
        self.res = ctk.CTkLabel(res_frame, text="----------",
                                 font=("Courier New", 58, "bold"), text_color="#4ade80",
                                 cursor="hand2")
        self.res.pack(side="left")
        self.res.bind("<Button-1>", self._copy_seq_flash)
        self._res_copy_btn = ctk.CTkButton(
            res_frame, text="📋", width=32, height=32,
            fg_color="#1e293b", hover_color="#334155",
            command=self._copy_sequence)
        self._res_copy_btn.pack(side="left", padx=(6, 0), anchor="s", pady=(0, 8))
        self._res_copy_hint = ctk.CTkLabel(res_outer, text=self.t("click_to_copy"),
            font=("Segoe UI", 9), text_color="#64748b")
        self._res_copy_hint.pack(pady=(0, 2))

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

        # Layer stack preview canvas (visual sequence bars)
        _seq_pf = ctk.CTkFrame(sec1, fg_color="#0f172a", corner_radius=6)
        _seq_pf.grid(row=5, column=0, padx=40, pady=(28, 2), sticky="ew")
        _seq_pf.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(_seq_pf, text="Schichten:", font=("Segoe UI", 8),
                     text_color="#475569").grid(row=0, column=0, padx=(8, 4), pady=3, sticky="w")
        self._seq_preview_canvas = tk.Canvas(_seq_pf, height=80, bg="#0f172a",
                                             highlightthickness=0, bd=0)
        self._seq_preview_canvas.grid(row=0, column=1, sticky="ew", padx=(0, 4), pady=3)

        # Gamut-Vorschau
        gf = ctk.CTkFrame(sec1, fg_color="#0f172a", corner_radius=6)
        gf.grid(row=6, column=0, padx=40, pady=(0, 4), sticky="ew")
        ctk.CTkLabel(gf, text="Gamut:", font=("Segoe UI", 8),
                     text_color="#475569").pack(side="left", padx=(8, 4), pady=3)
        self._gamut_canvas = tk.Canvas(gf, height=16, bg="#1e293b",
                                       highlightthickness=0, bd=0)
        self._gamut_canvas.pack(side="left", fill="x", expand=True, padx=4, pady=3)
        self._gamut_rects = []   # canvas rect IDs, filled lazily

        # Mix-Vorschau + ΔE
        mf = ctk.CTkFrame(sec1, fg_color="#1e293b", corner_radius=8,
                          border_width=2, border_color="#334155")
        mf.grid(row=7, column=0, padx=40, pady=6, sticky="ew")
        mf.grid_columnconfigure(1, weight=1)
        self._result_frame = mf

        tgt_col = ctk.CTkFrame(mf, fg_color="transparent")
        tgt_col.grid(row=0, column=0, padx=(20, 8), pady=12)
        ctk.CTkLabel(tgt_col, text=self.t("label_target"), font=("Segoe UI", 10),
                     text_color="#64748b").pack()
        self.target_circle = ctk.CTkLabel(tgt_col, text="", width=64, height=64,
                                           fg_color="#1e293b", corner_radius=32)
        self.target_circle.pack(pady=(4, 2))
        self.target_hex_lbl = ctk.CTkLabel(tgt_col, text="—", font=("Courier New", 9),
                                            text_color="#475569", cursor="hand2")
        self.target_hex_lbl.pack()
        self.target_hex_lbl.bind("<Button-1>", lambda e: (
            self.clipboard_clear(),
            self.clipboard_append(self.target_hex_lbl.cget("text"))))

        de_col = ctk.CTkFrame(mf, fg_color="transparent")
        de_col.grid(row=0, column=1, padx=20)
        self.de_disp = ctk.CTkLabel(de_col, text="ΔE  —",
                                     font=("Segoe UI", 18, "bold"), text_color="#475569")
        self.de_disp.pack()
        self.de_quality_lbl = ctk.CTkLabel(de_col, text="", font=("Segoe UI", 9),
                                            text_color="#475569")
        self.de_quality_lbl.pack()

        sim_col = ctk.CTkFrame(mf, fg_color="transparent")
        sim_col.grid(row=0, column=2, padx=(8, 20), pady=12)
        ctk.CTkLabel(sim_col, text=self.t("label_simulated"), font=("Segoe UI", 10),
                     text_color="#64748b").pack()
        self.sim_circle = ctk.CTkLabel(sim_col, text="", width=64, height=64,
                                        fg_color="#1e293b", corner_radius=32)
        self.sim_circle.pack(pady=(4, 2))
        self.sim_hex_lbl = ctk.CTkLabel(sim_col, text="—", font=("Courier New", 9),
                                         text_color="#475569", cursor="hand2")
        self.sim_hex_lbl.pack()
        self.sim_hex_lbl.bind("<Button-1>", lambda e: (
            self.clipboard_clear(),
            self.clipboard_append(self.sim_hex_lbl.cget("text"))))

        # Sequenzlänge + Auto-Modus
        lr = ctk.CTkFrame(sec1, fg_color="#1a2535", corner_radius=8)
        lr.grid(row=8, column=0, padx=40, pady=(4, 4), sticky="ew")
        lr.grid_columnconfigure(1, weight=1)

        self.len_label = ctk.CTkLabel(lr, text=self.t("length_label", n=10),
                                       font=("Segoe UI", 11, "bold"), width=80)
        self.len_label.grid(row=0, column=0, padx=(14, 6), pady=10)

        self.len_slider = ctk.CTkSlider(lr, from_=1, to=48, number_of_steps=47,
                                         command=self._on_len_slider)
        self.len_slider.set(10)
        self.len_slider.grid(row=0, column=1, sticky="ew", padx=6)

        self._auto_found_label = ctk.CTkLabel(lr, text="Auto — Länge wird berechnet",
                                               font=("Segoe UI", 9), text_color="#64748b")
        self._auto_found_label.grid(row=0, column=1, sticky="ew", padx=6)

        self.auto_len_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(lr, text=self.t("auto_check"), variable=self.auto_len_var,
                        font=("Segoe UI", 9),
                        command=self._on_auto_toggle).grid(row=0, column=2, padx=(8, 4))
        # Auto is default — hide slider, show label
        self.len_slider.grid_remove()

        ctk.CTkLabel(lr, text="ΔE≤", font=("Segoe UI", 10),
                     text_color="#64748b").grid(row=0, column=3, padx=(4, 0))
        self.auto_thresh_entry = ctk.CTkEntry(lr, width=46, placeholder_text="2.0")
        self.auto_thresh_entry.insert(0, "2.0")
        self.auto_thresh_entry.grid(row=0, column=4, padx=(2, 14))

        # Steuerleiste
        bl = ctk.CTkFrame(sec1, fg_color="transparent")
        bl.grid(row=9, column=0, padx=40, pady=(4, 16), sticky="ew")
        bl.grid_columnconfigure(0, weight=1)
        calc_btn = ctk.CTkButton(bl, text=self.t("btn_calculate"), fg_color="#2563eb",
                      command=self.calc, height=46,
                      font=("Segoe UI", 13, "bold"))
        calc_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._calc_btn = calc_btn
        self.tip(calc_btn, "tip_calculate")

        self.optimizer_var = ctk.BooleanVar(value=False)
        opt_cb = ctk.CTkCheckBox(bl, text=self.t("optimizer_check"), variable=self.optimizer_var,
                        font=("Segoe UI", 9))
        opt_cb.grid(row=0, column=1, padx=(0, 6))
        self.tip(opt_cb, "tip_optimizer")

        add_btn = ctk.CTkButton(bl, text=self.t("btn_add_virtual"),
                      fg_color="#16a34a", hover_color="#15803d",
                      command=self.add_to_virtual, height=44,
                      font=("Segoe UI", 13, "bold"))
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

        # Top-3 Sequenzen (nach Optimizer)
        self.top3_frame = ctk.CTkFrame(sec1, fg_color="#0f172a", corner_radius=8)
        self.top3_frame.grid(row=9, column=0, padx=40, pady=(0, 4), sticky="ew")
        self.top3_frame.grid_remove()

        # Keyboard shortcut: Enter = Berechnen
        self.bind("<Return>", lambda e: self.calc())

        # Stripe risk label (Change 4)
        self._stripe_label = ctk.CTkLabel(sec1, text="", font=("Segoe UI", 9),
                                          text_color="#64748b")
        self._stripe_label.grid(row=10, column=0, padx=40, pady=(0, 2), sticky="ew")

        # Layer schedule canvas (Change 8)
        _lsf = ctk.CTkFrame(sec1, fg_color="#0f172a", corner_radius=6)
        _lsf.grid(row=10, column=0, padx=40, pady=(18, 2), sticky="ew")
        _lsf.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(_lsf, text="L1-12:", font=("Segoe UI", 8),
                     text_color="#475569").grid(row=0, column=0, padx=(8, 4), pady=3, sticky="w")
        self._layer_sched_canvas = tk.Canvas(_lsf, height=32, bg="#0f172a",
                                             highlightthickness=0, bd=0)
        self._layer_sched_canvas.grid(row=0, column=1, sticky="ew", padx=(0, 4), pady=3)

        # Dithering-Profile
        dp_frame = ctk.CTkFrame(sec1, fg_color="#0f172a", corner_radius=8)
        dp_frame.grid(row=11, column=0, padx=40, pady=(0, 6), sticky="ew")
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

        # Farbsehschwäche-Simulation
        sim_frame = ctk.CTkFrame(sec1, fg_color="#0f172a", corner_radius=8)
        sim_frame.grid(row=12, column=0, padx=40, pady=(0, 12), sticky="ew")
        ctk.CTkLabel(sim_frame, text=self.t("colorblind_label"),
                     font=("Segoe UI", 10), text_color="#64748b").pack(side="left", padx=(12, 6))
        self.colorblind_var = ctk.StringVar(value="normal")
        for val, key in [("normal", "colorblind_normal"), ("prot", "colorblind_prot"),
                         ("deut", "colorblind_deut"), ("trit", "colorblind_trit")]:
            ctk.CTkRadioButton(sim_frame, text=self.t(key), variable=self.colorblind_var,
                               value=val, font=("Segoe UI", 10),
                               command=self._update_colorblind_preview).pack(side="left", padx=4, pady=6)

        # Recent colors palette (Change 7)
        recent_outer = ctk.CTkFrame(sec1, fg_color="#0f172a", corner_radius=8)
        recent_outer.grid(row=13, column=0, padx=40, pady=(0, 4), sticky="ew")
        ctk.CTkLabel(recent_outer, text="Zuletzt:", font=("Segoe UI", 9),
                     text_color="#64748b").pack(side="left", padx=(8, 4), pady=4)
        self._recent_frame = ctk.CTkFrame(recent_outer, fg_color="transparent")
        self._recent_frame.pack(side="left", fill="x", expand=True, pady=4)
        self._update_recent_swatches()

        # Feature 5: History panel (last 10 calculations)
        hist_outer = ctk.CTkFrame(sec1, fg_color="#0f172a", corner_radius=8)
        hist_outer.grid(row=14, column=0, padx=40, pady=(0, 12), sticky="ew")
        ctk.CTkLabel(hist_outer, text="🕘 Verlauf",
                     font=("Segoe UI", 10), text_color="#64748b").pack(anchor="w", padx=12, pady=(6, 2))
        self._history_frame = ctk.CTkScrollableFrame(hist_outer, height=100, fg_color="transparent")
        self._history_frame.pack(fill="x", padx=8, pady=(0, 6))

        # ────────────────────────────────────────────────────────────────────
        # TAB 2 — VIRTUELLE DRUCKKÖPFE
        # ────────────────────────────────────────────────────────────────────

        # Titel
        v_title = ctk.CTkFrame(tab2, fg_color="transparent")
        v_title.grid(row=0, column=0, padx=8, pady=(8, 4), sticky="ew")
        ctk.CTkLabel(v_title, text=self.t("sec2_title"),
                     font=("Segoe UI", 15, "bold"), text_color="#a78bfa").pack(side="left")
        ctk.CTkLabel(v_title, text="  " + self.t("sec2_desc", max_v=self._max_virtual),
                     font=("Segoe UI", 10), text_color="#64748b").pack(side="left")

        # Toolbar-Zeile 1: Kern-Operationen
        tb1 = ctk.CTkFrame(tab2, fg_color="#1e293b", corner_radius=8)
        tb1.grid(row=1, column=0, padx=8, pady=(0, 3), sticky="ew")

        tmf_btn = ctk.CTkButton(tb1, text=self.t("btn_3mf"), fg_color="#0f4c81",
                      hover_color="#1e3a5f", height=36, font=("Segoe UI", 11, "bold"),
                      command=self.open_3mf_assistant)
        tmf_btn.pack(side="left", padx=(8, 4), pady=6)
        self.tip(tmf_btn, "tip_3mf")

        wizard_btn = ctk.CTkButton(tb1, text=self.t("wizard_btn"), fg_color="#7c3aed",
                      hover_color="#6d28d9", height=36, font=("Segoe UI", 11, "bold"),
                      command=self.open_3mf_wizard)
        wizard_btn.pack(side="left", padx=(0, 4), pady=6)

        bat_btn = ctk.CTkButton(tb1, text=self.t("btn_batch"), fg_color="#4338ca",
                      hover_color="#3730a3", height=36, font=("Segoe UI", 11, "bold"),
                      command=self.open_batch_dialog)
        bat_btn.pack(side="left", padx=(0, 4), pady=6)
        self.tip(bat_btn, "tip_batch")

        undo_btn = ctk.CTkButton(tb1, text=self.t("btn_undo"), fg_color="#374151",
                      height=36, command=self.undo_virtual)
        undo_btn.pack(side="left", padx=(0, 4), pady=6)

        ctk.CTkButton(tb1, text=self.t("btn_del_all"), fg_color="#7f1d1d",
                      hover_color="#991b1b", height=36,
                      command=self.clear_virtual).pack(side="left", padx=(0, 4), pady=6)

        ctk.CTkButton(tb1, text=self.t("btn_recalc_all"), fg_color="#065f46",
                      hover_color="#047857", height=36,
                      command=self.recalc_all_virtual).pack(side="left", padx=(0, 8), pady=6)

        # Toolbar-Zeile 2: Export & Analyse
        tb2 = ctk.CTkFrame(tab2, fg_color="#1a2535", corner_radius=8)
        tb2.grid(row=2, column=0, padx=8, pady=(0, 6), sticky="ew")

        ctk.CTkButton(tb2, text=self.t("btn_export_all"), fg_color="#374151",
                      height=32, command=self.open_export_dialog).pack(side="left", padx=(8, 4), pady=5)

        orca_btn = ctk.CTkButton(tb2, text=self.t("btn_orca_export"), fg_color="#0f766e",
                      hover_color="#0d6660", height=32, font=("Segoe UI", 10, "bold"),
                      command=self.open_orca_export_dialog)
        orca_btn.pack(side="left", padx=(0, 4), pady=5)
        self.tip(orca_btn, "tip_orca_export")

        ctk.CTkButton(tb2, text=self.t("btn_3mf_write"), fg_color="#374151",
                      height=32, command=self.write_3mf_colors).pack(side="left", padx=(0, 4), pady=5)

        ctk.CTkButton(tb2, text=self.t("btn_de_overview"), fg_color="#1e3a5f",
                      height=32, command=self.open_de_overview).pack(side="left", padx=(0, 4), pady=5)

        ctk.CTkButton(tb2, text=self.t("btn_recipe"), fg_color="#374151",
                      height=32, command=self.open_recipe_export).pack(side="left", padx=(0, 4), pady=5)

        ctk.CTkButton(tb2, text=self.t("btn_copy_all_cad"), fg_color="#0e7490",
                      height=32, command=self.open_copy_all_cadence).pack(side="left", padx=(0, 8), pady=5)

        # Sort / Filter bar for virtual heads
        sf_bar = ctk.CTkFrame(tab2, fg_color="#1a2535", corner_radius=6)
        sf_bar.grid(row=3, column=0, padx=8, pady=(0, 2), sticky="ew")
        ctk.CTkLabel(sf_bar, text="🔍", font=("Segoe UI", 12),
                     text_color="#64748b").pack(side="left", padx=(8, 2), pady=5)
        self._virt_filter_var = ctk.StringVar()
        _virt_filter_entry = ctk.CTkEntry(sf_bar, textvariable=self._virt_filter_var,
                                           placeholder_text=self.t("virt_filter_placeholder"),
                                           width=160, height=28)
        _virt_filter_entry.pack(side="left", padx=(0, 8), pady=5)
        self._virt_sort_var = ctk.StringVar(value=self.t("virt_sort_added"))
        _sort_opts = [self.t("virt_sort_added"), self.t("virt_sort_de_asc"),
                      self.t("virt_sort_de_desc"), self.t("virt_sort_label")]
        _sort_menu = ctk.CTkOptionMenu(sf_bar, variable=self._virt_sort_var,
                                        values=_sort_opts, width=130, height=28,
                                        command=lambda _: self._refresh_virtual_grid())
        _sort_menu.pack(side="left", padx=(0, 8), pady=5)
        self._virt_filter_var.trace_add("write", lambda *a: self.after(200, self._refresh_virtual_grid))

        # Grid-Header
        gh = ctk.CTkFrame(tab2, fg_color="#1e293b", corner_radius=6)
        gh.grid(row=4, column=0, padx=8, sticky="ew", pady=(0, 2))
        gh.grid_columnconfigure(4, weight=1)
        for col, (txt, w) in enumerate([("ID", 55), (self.t("grid_target"), 50),
                                         (self.t("grid_sequence"), 160), (self.t("grid_simulated"), 60),
                                         (self.t("grid_quality"), 130), (self.t("grid_label"), 0),
                                         ("", 40)]):
            ctk.CTkLabel(gh, text=txt, font=("Segoe UI", 10, "bold"),
                         text_color="#64748b", width=w).grid(
                row=0, column=col, padx=8, pady=6, sticky="w" if w == 0 else "")

        # Scrollbares Grid
        self.vgrid = ctk.CTkScrollableFrame(tab2, fg_color="#0f172a", corner_radius=8)
        self.vgrid.grid(row=5, column=0, padx=8, sticky="nsew", pady=(0, 0))
        tab2.grid_rowconfigure(5, weight=1)
        tab2.grid_rowconfigure(2, weight=0)
        self.vgrid.grid_columnconfigure(4, weight=1)
        self._refresh_virtual_grid()

        # DnD hint label + export hint label
        _hint_row = ctk.CTkFrame(tab2, fg_color="transparent")
        _hint_row.grid(row=6, column=0, padx=8, pady=(2, 6), sticky="ew")
        _hint_row.grid_columnconfigure(0, weight=1)
        if _HAS_DND:
            ctk.CTkLabel(_hint_row, text="← .3mf oder .u1proj hierher ziehen",
                         font=("Segoe UI", 9), text_color="#334155").grid(
                row=0, column=0, sticky="w", padx=4)
        self._export_hint_label = ctk.CTkLabel(_hint_row, text="",
                                               font=("Segoe UI", 9), text_color="#64748b")
        self._export_hint_label.grid(row=0, column=1, sticky="e", padx=4)

        # ────────────────────────────────────────────────────────────────────
        # TAB 3 — WERKZEUGE
        # ────────────────────────────────────────────────────────────────────

        def _tools_section(parent, title):
            ctk.CTkLabel(parent, text=title, font=("Segoe UI", 11, "bold"),
                         text_color="#64748b").pack(anchor="w", padx=16, pady=(14, 4))
            f = ctk.CTkFrame(parent, fg_color="#1e293b", corner_radius=8)
            f.pack(fill="x", padx=12, pady=(0, 6))
            return f

        def _tool_btn(parent, text, color, cmd, tip_key=None):
            b = ctk.CTkButton(parent, text=text, fg_color=color, height=36,
                              font=("Segoe UI", 11), command=cmd)
            b.pack(side="left", padx=6, pady=8)
            if tip_key:
                self.tip(b, tip_key)
            return b

        # Farb-Visualisierung
        f = _tools_section(tab3, "🔭  Analyse & Visualisierung")
        _tool_btn(f, self.t("btn_lab_plot"), "#0f4c81", self.show_lab_plot, "tip_lab_plot")
        _tool_btn(f, self.t("btn_gamut_plot"), "#0f4c81", self.open_gamut_plot)
        _tool_btn(f, self.t("btn_swatch"), "#374151", self.save_swatch, "tip_swatch")
        _tool_btn(f, self.t("btn_slicer_guide"), "#7c3aed", self.open_slicer_guide)
        _tool_btn(f, self.t("btn_tc_est"), "#374151", self.open_tc_estimator)
        _tool_btn(f, "📊 ΔE-Matrix", "#1e3a5f", self.open_filament_matrix)
        _tool_btn(f, "🖼 PNG Export", "#374151", self.export_png_summary)

        # Farb-Generierung
        f = _tools_section(tab3, "🌈  Farb-Generierung")
        _tool_btn(f, self.t("btn_gradient"), "#0e7490", self.open_gradient_dialog, "tip_gradient")
        _tool_btn(f, self.t("btn_harmonies"), "#7c3aed", self.open_harmonies_dialog, "tip_harmonies")
        _tool_btn(f, self.t("btn_palette"), "#374151", self.import_palette_from_image, "tip_palette")
        _tool_btn(f, self.t("btn_multitarget"), "#7c3aed", self.open_multitarget_optimizer)
        _tool_btn(f, self.t("btn_multi_gradient"), "#0e7490", self.open_multi_gradient_dialog)

        # Optimierung
        f = _tools_section(tab3, "🎯  Optimierung & Kalibrierung")
        _tool_btn(f, self.t("btn_slot_opt"), "#7c3aed", self.open_slot_optimizer)
        _tool_btn(f, self.t("btn_td_cal"), "#0f4c81", self.open_td_calibration)
        _tool_btn(f, self.t("btn_slot_compare"), "#7c3aed", self.open_slot_compare)

        # Bibliothek
        f = _tools_section(tab3, "📚  Bibliothek & Datenbank")
        _tool_btn(f, self.t("btn_new_brand"), "#1e3a5f", self.add_brand)
        _tool_btn(f, self.t("btn_library"), "#374151", self.open_library_manager)
        web_b = _tool_btn(f, self.t("btn_web_update"), "#164e63", self.web_update_library,
                          "tip_web_update")

    # ── SLOT-LOGIK ─────────────────────────────────────────────────────────────

    def update_menu(self, idx):
        b = self.slots[idx]["brand"].get()
        cols = [f["name"] for f in self.library.get(b, [])] or [self.t("empty_slot")]
        self.slots[idx]["color"].configure(values=cols)
        self.slots[idx]["color"].set(cols[0])
        self.apply_f(idx)

    def _update_slot_strip(self, idx):
        """Update the colored left-border strip to match slot's current hex color."""
        try:
            strip = self.slots[idx].get("_strip")
            hex_val = self.slots[idx]["hex"].get().strip() or "#334155"
            if strip:
                strip.configure(fg_color=hex_val)
        except Exception:
            pass

    def apply_f(self, idx):
        b = self.slots[idx]["brand"].get()
        n = self.slots[idx]["color"].get()
        if n in _SLOT_SKIP: return
        f = next((x for x in self.library.get(b, []) if x["name"] == n), None)
        if f is None: return
        if not getattr(self, "_building_ui", False):
            self._save_slot_snapshot()
        self.slots[idx]["hex"].delete(0, "end"); self.slots[idx]["hex"].insert(0, f["hex"])
        self.slots[idx]["td"].delete(0, "end");  self.slots[idx]["td"].insert(0, str(f["td"]))
        self.slots[idx]["preview"].configure(fg_color=f["hex"])
        self._update_slot_strip(idx)
        if not getattr(self, "_building_ui", False):
            self._schedule_gamut_update()

    def pick_slot_color(self, idx):
        cur = self.slots[idx]["hex"].get().strip() or "#808080"
        h = self._ask_color(initial=cur, title=self.t("color_picker_title", i=idx+1))
        if h is None: return
        self.slots[idx]["hex"].delete(0, "end"); self.slots[idx]["hex"].insert(0, h)
        self.slots[idx]["preview"].configure(fg_color=h)
        self._update_slot_strip(idx)
        cur_vals = [f["name"] for f in self.library.get(self.slots[idx]["brand"].get(), [])]
        manual = self.t("manual_color")
        self.slots[idx]["color"].configure(values=[manual] + cur_vals)
        self.slots[idx]["color"].set(manual)
        self._schedule_gamut_update()

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
        """Show/hide slider vs auto-found label depending on auto mode."""
        if self.auto_len_var.get():
            self.len_slider.grid_remove()
            if hasattr(self, "_auto_found_label"):
                self._auto_found_label.grid()
        else:
            if hasattr(self, "_auto_found_label"):
                self._auto_found_label.grid_remove()
            self.len_slider.grid()

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
        """Farbmodell umschaltbar: linear (Standard), TD-gewichtet, subtraktiv, FilamentMixer."""
        by_id = {f["id"]: f for f in fils}
        counts = {}
        for fid in sequence:
            counts[fid] = counts.get(fid, 0) + 1
        total = len(sequence)
        model = getattr(self, "color_model_var", None)
        model = model.get() if model else "linear"

        if model == "filamentmixer":
            # Blend pairs sequentially using pigment model
            if not sequence or not fils:
                return rgb_to_lab((128, 128, 128))
            fids = [int(s) for s in sequence]
            fil_map = {f["id"]: hex_to_rgb(f["hex"]) for f in fils}
            r, g, b = fil_map.get(fids[0], (128, 128, 128))
            for i in range(1, len(fids)):
                r2, g2, b2 = fil_map.get(fids[i], (r, g, b))
                t = 1.0 / (i + 1)  # progressive blend weight
                r, g, b = filament_mixer_lerp(r, g, b, r2, g2, b2, t)
            return rgb_to_lab((r, g, b))

        r_acc = g_acc = b_acc = 0.0
        total_w = 0.0
        for fid, cnt in counts.items():
            r, g, b = hex_to_rgb(by_id[fid]["hex"])
            td = max(0.1, float(by_id[fid].get("td", 5.0)))
            base_w = cnt / total
            w = (base_w / td) if model == "td" else base_w
            total_w += w
            rl = (r/255)**2.2; gl = (g/255)**2.2; bl = (b/255)**2.2
            if model == "subtractive":
                r_acc += (1-rl)*w; g_acc += (1-gl)*w; b_acc += (1-bl)*w
            else:
                r_acc += rl*w; g_acc += gl*w; b_acc += bl*w
        if total_w > 0:
            r_acc /= total_w; g_acc /= total_w; b_acc /= total_w
        if model == "subtractive":
            r_acc = 1-r_acc; g_acc = 1-g_acc; b_acc = 1-b_acc
        def tog(v): return round(min(255, max(0, v**(1/2.2)*255)))
        return rgb_to_lab((tog(r_acc), tog(g_acc), tog(b_acc)))

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

        seq_len : feste Länge 1–48 (None = aktueller Slider-Wert)
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

        all_opt_results = []  # collect all permutation results for top-3

        def best_seq_for_n(n):
            if optimizer:
                best, best_dv = None, float("inf")
                for perm in iter_permutations(scores):
                    seq = self._build_sequence(list(perm), tot, n)
                    dv  = delta_e(self._simulate_mix(seq, fils), t_lab)
                    seq_str = "".join(map(str, seq))
                    all_opt_results.append({
                        "target_hex": target_hex,
                        "sequence": seq_str,
                        "sim_hex": lab_to_hex(self._simulate_mix(seq, fils)),
                        "de": dv,
                        "seq_len": n,
                    })
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

        # Store top-3 for display (only when optimizer was run)
        if optimizer and all_opt_results:
            all_opt_results.sort(key=lambda x: x["de"])
            # Deduplicate by sequence string
            seen_seqs = set()
            unique_results = []
            for r in all_opt_results:
                if r["sequence"] not in seen_seqs:
                    seen_seqs.add(r["sequence"])
                    unique_results.append(r)
            self._top3_results = unique_results[:3]
        else:
            self._top3_results = []

        sim_lab = self._simulate_mix(chosen_seq, fils)
        dv      = delta_e(sim_lab, t_lab)
        return {
            "target_hex": target_hex,
            "sequence":   "".join(map(str, chosen_seq)),
            "sim_hex":    lab_to_hex(sim_lab),
            "de":         dv,
            "seq_len":    len(chosen_seq),
        }

    def _draw_seq_preview(self, sequence):
        """Draw colored horizontal bars representing each layer in the sequence."""
        if not hasattr(self, '_seq_preview_canvas'):
            return
        c = self._seq_preview_canvas
        c.update_idletasks()
        W = c.winfo_width() or 300
        n = len(sequence)
        if n == 0:
            c.delete("all")
            return
        bar_h = max(1, 80 // n)
        fils = {f["id"]: f["hex"] for f in self._get_fils()}
        c.delete("all")
        for i, fid_str in enumerate(reversed(sequence)):
            try:
                fid = int(fid_str)
            except (ValueError, TypeError):
                continue
            color = fils.get(fid, "#888888")
            y0 = i * bar_h
            y1 = y0 + bar_h
            c.create_rectangle(0, y0, W, y1, fill=color, outline="")

    def _draw_layer_schedule(self, sequence):
        """Draw colored squares for first 12 layers showing active filament (Change 8)."""
        if not hasattr(self, "_layer_sched_canvas"):
            return
        c = self._layer_sched_canvas
        c.update_idletasks()
        W = c.winfo_width() or 400
        H = 32
        c.delete("all")
        if not sequence:
            return
        schedule = compute_layer_schedule(sequence, n_layers=12)
        fils_hex = {f["id"]: f["hex"] for f in self._get_fils()}
        n = len(schedule)
        sq = min(28, (W - 4) // max(n, 1))
        for i, fid in enumerate(schedule):
            color = fils_hex.get(fid, "#888888")
            x0 = 2 + i * (sq + 2)
            c.create_rectangle(x0, 2, x0 + sq, H - 2, fill=color, outline="#334155")
            # Label Ln
            r_, g_, b_ = hex_to_rgb(color)
            lum = 0.299 * r_ + 0.587 * g_ + 0.114 * b_
            tc = "#111111" if lum > 140 else "#eeeeee"
            c.create_text(x0 + sq // 2, H // 2, text=f"L{i+1}", fill=tc,
                          font=("Segoe UI", 7))

    def _schedule_gamut_update(self, delay=150):
        """Debounced gamut strip update — cancels any pending call and reschedules."""
        if self._gamut_job:
            self.after_cancel(self._gamut_job)
        self._gamut_job = self.after(delay, self._run_gamut_update)

    def _update_gamut_strip(self):
        """Public entry point — delegates through the debouncer."""
        self._schedule_gamut_update()

    def _run_gamut_update(self):
        """Redraws the gamut strip Canvas with reachable mixed colors."""
        self._gamut_job = None
        if not hasattr(self, "_gamut_canvas"):
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

        # Draw onto Canvas — much faster than 40 individual CTkLabel widgets
        canvas = self._gamut_canvas
        canvas.update_idletasks()
        W = canvas.winfo_width()
        H = canvas.winfo_height() or 16
        if W < 2:
            W = 400
        n = min(len(samples), 60)
        cell_w = W / n
        canvas.delete("all")
        for i in range(n):
            idx = int(i * len(samples) / n)
            color = samples[min(idx, len(samples) - 1)]
            x0 = int(i * cell_w)
            x1 = int((i + 1) * cell_w) + 1  # +1 avoids gaps between rects
            canvas.create_rectangle(x0, 0, x1, H, fill=color, outline="")

    def _show_top3(self):
        """Zeigt Top-3 Sequenzen nach Optimizer in einem kompakten Frame."""
        for w in self.top3_frame.winfo_children():
            w.destroy()
        results = getattr(self, "_top3_results", [])
        if len(results) < 2:
            self.top3_frame.grid_remove()
            return
        self.top3_frame.grid()
        ctk.CTkLabel(self.top3_frame, text=self.t("top3_title"),
                     font=("Segoe UI", 10, "bold"), text_color="#64748b").pack(
            side="left", padx=(10, 8), pady=6)
        for idx, r in enumerate(results[:3]):
            card = ctk.CTkFrame(self.top3_frame, fg_color="#1e293b", corner_radius=6)
            card.pack(side="left", padx=3, pady=6, fill="y")
            ctk.CTkLabel(card, text=self.t("top3_rank", r=idx+1),
                         font=("Segoe UI", 9, "bold"),
                         text_color="#64748b").pack(side="left", padx=(6, 2))
            ctk.CTkLabel(card, text=" ".join(r["sequence"]),
                         font=("Courier New", 10, "bold"),
                         text_color="#4ade80").pack(side="left", padx=2)
            ctk.CTkLabel(card, text=f"  ΔE {r['de']:.1f}",
                         font=("Segoe UI", 9),
                         text_color=de_color(r["de"])).pack(side="left", padx=2)
            sim_c = ctk.CTkLabel(card, text="", width=20, height=20,
                                  fg_color=r["sim_hex"], corner_radius=4)
            sim_c.pack(side="left", padx=4)
            def add_this(res=r):
                self.last_result = res
                self.add_to_virtual(res)
            ctk.CTkButton(card, text=self.t("top3_add"), width=70, height=24,
                          fg_color="#15803d", command=add_this).pack(
                side="left", padx=(2, 6))

    def calc(self):
        if not hasattr(self, "target"):
            messagebox.showinfo(self.t("dlg_note"), self.t("dlg_select_color")); return

        # Loading state
        if hasattr(self, "_calc_btn"):
            self._calc_btn.configure(text="⏳ Berechne…", state="disabled")
            self.update_idletasks()

        try:
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
                if hasattr(self, "_auto_found_label"):
                    self._auto_found_label.configure(text=self.t("auto_found", n=n))

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
            self._show_top3()
            self._draw_seq_preview(seq)

            # Stripe risk check (Change 4)
            _risk, _risk_msg = check_stripe_risk(seq)
            if hasattr(self, "_stripe_label"):
                if _risk:
                    self._stripe_label.configure(text=_risk_msg, text_color="#f97316")
                else:
                    self._stripe_label.configure(text=_risk_msg, text_color="#4ade80")

            # Layer schedule (Change 8)
            self._draw_layer_schedule(seq)

            # High ΔE visual warning on result border
            self._set_result_border(dv)

            # Update recent colors
            self._add_recent_color(self.target)

            # Feature 9: Auto-suggestion
            if dv > 6:
                self._check_auto_suggestion(self.target)
            else:
                if hasattr(self, "suggestion_label"):
                    self.suggestion_label.configure(text="")

            # Feature 10: Material compatibility warning
            seq_ids = list(set(int(c) - 1 for c in seq))
            mat_warn = self._check_material_compatibility(seq_ids)
            if mat_warn and hasattr(self, "suggestion_label"):
                cur = self.suggestion_label.cget("text")
                combined = mat_warn if not cur else f"{cur}  |  {mat_warn}"
                self.suggestion_label.configure(text=combined, text_color="#f59e0b")

            # Update history
            self._update_history(self.target, result["sim_hex"], result["de"], seq)

            # Status bar update
            self._set_status(self.t("status_calculated", de=dv, seq=seq), 5000)

        finally:
            if hasattr(self, "_calc_btn"):
                self._calc_btn.configure(text=self.t("btn_calculate"), state="normal")

    # ── VIRTUELLE DRUCKKÖPFE ───────────────────────────────────────────────────

    def _copy_sequence(self):
        seq = self.res.cget("text").strip()
        if seq and seq != "----------":
            self.clipboard_clear()
            self.clipboard_append(seq)

    def _copy_seq_flash(self, event=None):
        """Click handler on result label — copy to clipboard and flash green."""
        seq = self.res.cget("text")
        if not seq or seq == "----------" or seq == "—":
            return
        self.clipboard_clear()
        self.clipboard_append(seq)
        try:
            orig = self.res.cget("fg_color")
            self.res.configure(fg_color="#166534")
            self.after(600, lambda: self.res.configure(fg_color=orig))
        except Exception:
            pass

    def _set_result_border(self, de):
        """Set the result frame border color based on ΔE value."""
        if not hasattr(self, "_result_frame"):
            return
        if de < DE_GOOD:
            color = "#16a34a"
        elif de < DE_OK:
            color = "#d97706"
        else:
            color = "#dc2626"
            self._pulse_result_border(3)
        self._result_frame.configure(border_color=color, border_width=2)

    def _pulse_result_border(self, n):
        """Pulse the result border width n times for high ΔE warning."""
        if n <= 0 or not hasattr(self, "_result_frame"):
            return
        try:
            cur_w = self._result_frame.cget("border_width")
            self._result_frame.configure(border_width=4 if cur_w == 2 else 2)
            self.after(500, lambda: self._pulse_result_border(n - 1))
        except Exception:
            pass

    def _on_model_change(self):
        if hasattr(self, "target") and self.target:
            self.calculate()
        self._refresh_virtual_grid()

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
            "stable_id":  5 + len(self.virtual_fils),
            "target_hex": result["target_hex"],
            "sequence":   result["sequence"],
            "sim_hex":    result["sim_hex"],
            "de":         result["de"],
            "label":      self.t("virtual_label_default", vid=vid),
        })
        self._refresh_virtual_grid()
        self._set_status(self.t("status_added", vid=vid), 3000)

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

    def _get_sorted_filtered_virtual(self):
        """Return virtual_fils list filtered and sorted per UI controls."""
        fils = list(self.virtual_fils)
        # Filter
        ftext = ""
        if hasattr(self, "_virt_filter_var"):
            ftext = self._virt_filter_var.get().lower().strip()
        if ftext:
            fils = [v for v in fils if ftext in v.get("label", "").lower()
                    or ftext in v.get("sequence", "").lower()
                    or ftext in v.get("target_hex", "").lower()]
        # Sort
        sk = ""
        if hasattr(self, "_virt_sort_var"):
            sk = self._virt_sort_var.get()
        if "ΔE ↑" in sk:
            fils.sort(key=lambda v: v.get("de", 99))
        elif "ΔE ↓" in sk:
            fils.sort(key=lambda v: v.get("de", 0), reverse=True)
        elif "Label" in sk or "A-Z" in sk:
            fils.sort(key=lambda v: v.get("label", ""))
        return fils

    def _refresh_virtual_grid(self):
        if not hasattr(self, "vgrid"):
            return
        for w in self.vgrid.winfo_children():
            w.destroy()
        if not self.virtual_fils:
            ctk.CTkLabel(self.vgrid,
                         text=self.t("empty_virtual"),
                         text_color="#334155", font=("Segoe UI", 11),
                         wraplength=600).pack(pady=30)
            return

        self.vgrid.grid_columnconfigure(4, weight=1)
        for row_data in self._get_sorted_filtered_virtual():
            self._build_virtual_row(row_data)

    def _build_virtual_row(self, vf):
        outer = ctk.CTkFrame(self.vgrid, fg_color="#1e293b", corner_radius=7)
        outer.pack(fill="x", padx=6, pady=3)
        outer.grid_columnconfigure(5, weight=1)

        # Get index for reorder buttons (Feature 4)
        try:
            row_idx = self.virtual_fils.index(vf)
        except ValueError:
            row_idx = 0

        # Feature 4: ↑↓ reorder buttons
        nav_frame = ctk.CTkFrame(outer, fg_color="transparent")
        nav_frame.grid(row=0, column=0, padx=(4, 0), pady=(8, 2))
        ctk.CTkButton(nav_frame, text="↑", width=22, height=20, fg_color="#334155",
                      command=lambda i=row_idx: self._vhead_move(i, -1)).pack(pady=(0, 1))
        ctk.CTkButton(nav_frame, text="↓", width=22, height=20, fg_color="#334155",
                      command=lambda i=row_idx: self._vhead_move(i, +1)).pack()

        # Zeile 0: Haupt-Info
        ctk.CTkLabel(outer, text=f"V{vf['vid']}",
                     font=("Segoe UI", 13, "bold"), text_color="#a78bfa",
                     width=44).grid(row=0, column=1, padx=(4, 2), pady=(8, 2))

        # Feature 3: clickable target circle
        def _copy_hex_to_clipboard(hex_val, widget):
            self.clipboard_clear()
            self.clipboard_append(hex_val)
            try:
                orig = widget.cget("text")
                widget.configure(text="✓")
                widget.after(800, lambda: widget.configure(text=orig))
            except Exception:
                pass

        tgt_lbl = ctk.CTkLabel(outer, text="", width=36, height=36,
                               fg_color=vf["target_hex"], corner_radius=18, cursor="hand2")
        tgt_lbl.grid(row=0, column=2, padx=4, pady=(8, 2))
        tgt_lbl.bind("<Button-1>", lambda e, h=vf["target_hex"], w=tgt_lbl:
                     _copy_hex_to_clipboard(h, w))

        ctk.CTkLabel(outer, text=vf["sequence"],
                     font=("Courier New", 17, "bold"), text_color="#4ade80",
                     width=155).grid(row=0, column=3, padx=6, pady=(8, 2))

        # Feature 3: clickable sim circle
        sim_lbl = ctk.CTkLabel(outer, text="", width=36, height=36,
                               fg_color=vf["sim_hex"], corner_radius=18, cursor="hand2")
        sim_lbl.grid(row=0, column=4, padx=4, pady=(8, 2))
        sim_lbl.bind("<Button-1>", lambda e, h=vf["sim_hex"], w=sim_lbl:
                     _copy_hex_to_clipboard(h, w))

        ctk.CTkLabel(outer, text=self.de_label(vf["de"]),
                     font=("Segoe UI", 11, "bold"), text_color=de_color(vf["de"])).grid(
            row=0, column=5, padx=6, sticky="w", pady=(8, 2))
        lbl_entry = ctk.CTkEntry(outer, width=130, font=("Segoe UI", 11))
        lbl_entry.insert(0, vf.get("label", ""))
        lbl_entry.bind("<FocusOut>", lambda e, vfd=vf, en=lbl_entry:
                       vfd.update({"label": en.get()}))
        lbl_entry.grid(row=0, column=6, padx=4, pady=(8, 2))

        # Feature 12: sequence editor button
        ctk.CTkButton(outer, text="✏", width=28, height=30,
                      fg_color="#1e3a5f", hover_color="#2563eb",
                      command=lambda i=row_idx: self.open_sequence_editor(i)).grid(
            row=0, column=7, padx=(2, 2), pady=(8, 2))

        ctk.CTkButton(outer, text="✕", width=30, height=30,
                      fg_color="#7f1d1d", hover_color="#991b1b",
                      command=lambda vid=vf["vid"]: self._remove_virtual(vid)).grid(
            row=0, column=8, padx=(2, 10), pady=(8, 2))

        # Zeile 1: Runs-Visualisierung + Cadence-Hinweis
        info_row = ctk.CTkFrame(outer, fg_color="transparent")
        info_row.grid(row=1, column=0, columnspan=9, padx=10, pady=(0, 8), sticky="w")

        runs = seq_to_runs(vf["sequence"])
        n_fils = seq_filament_count(vf["sequence"])
        lh = safe_td(self.layer_height_entry.get()) if hasattr(self, "layer_height_entry") else 0.08

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

        # Werkzeugwechsel-Warnung: Anzahl Transitionen pro Zyklus
        transitions = sum(1 for k in range(len(vf["sequence"]) - 1)
                          if vf["sequence"][k] != vf["sequence"][k + 1])
        if transitions > 2:
            ctk.CTkLabel(info_row, text=self.t("tc_warn_badge", n=transitions),
                         font=("Segoe UI", 8, "bold"), text_color="#f59e0b",
                         fg_color="#292524", corner_radius=4).pack(side="left", padx=(0, 4))

        # Cadence-Hinweis
        pattern_str = ",".join(vf["sequence"])   # "1121" → "1,1,2,1"

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

        import tkinter as _tk
        _hint_entry = _tk.Entry(
            info_row, font=("Segoe UI", 9), fg=hint_color,
            bg="#0f172a", readonlybackground="#0f172a",
            relief="flat", bd=0, highlightthickness=0,
            state="readonly", width=max(len(hint), 10))
        _hint_entry.configure(state="normal")
        _hint_entry.delete(0, "end")
        _hint_entry.insert(0, hint)
        _hint_entry.configure(state="readonly")
        _hint_entry.pack(side="left", padx=4)
        # 📋 Kopier-Button — kopiert nur die reine Ziffernkette (z.B. "1121")
        raw_seq = "".join(vf["sequence"])
        def _copy_hint(s=raw_seq):
            self.clipboard_clear(); self.clipboard_append(s)
        ctk.CTkButton(info_row, text="📋", width=24, height=20,
                      fg_color="#1e293b", hover_color="#334155",
                      command=_copy_hint).pack(side="left", padx=(0, 4))

    def _remove_virtual(self, vid):
        self.virtual_undo.append(copy.deepcopy(self.virtual_fils))
        self.virtual_fils = [v for v in self.virtual_fils if v["vid"] != vid]
        for i, v in enumerate(self.virtual_fils):
            v["vid"] = 5 + i
        self._refresh_virtual_grid()

    # ── LIVE ΔE ────────────────────────────────────────────────────────────────

    def _on_hex_live(self, event=None):
        """Quick ΔE preview while the user types a hex code (debounced 300 ms)."""
        if not hasattr(self, "_hex_live_job"):
            self._hex_live_job = None
        if self._hex_live_job:
            self.after_cancel(self._hex_live_job)
        self._hex_live_job = self.after(300, self._run_hex_live)

    def _run_hex_live(self):
        self._hex_live_job = None
        raw = self.hex_target_entry.get().strip()
        if not raw.startswith("#"):
            raw = "#" + raw
        if len(raw) != 7:
            if hasattr(self, "live_de_label"):
                self.live_de_label.configure(text="")
            return
        try:
            hex_to_rgb(raw)
        except Exception:
            return
        result = self._calc_for_color(raw, optimizer=False, seq_len=4, auto=False)
        if result and hasattr(self, "live_de_label"):
            de = result["de"]
            self.live_de_label.configure(
                text=f"≈ΔE {de:.1f}", text_color=de_color(de))

    # ── BATCH-NEUBERECHNUNG ─────────────────────────────────────────────────────

    def recalc_all_virtual(self):
        """Berechnet alle virtuellen Köpfe mit den aktuellen Slot-Einstellungen neu."""
        if not self.virtual_fils:
            messagebox.showinfo(self.t("dlg_note"), self.t("orca_no_virtual")); return
        self.virtual_undo.append(copy.deepcopy(self.virtual_fils))
        use_opt  = self.optimizer_var.get() if hasattr(self, "optimizer_var") else False
        use_auto = self.auto_len_var.get() if hasattr(self, "auto_len_var") else True
        thresh   = safe_td(self.auto_thresh_entry.get()) if hasattr(self, "auto_thresh_entry") else 2.0
        for vf in self.virtual_fils:
            result = self._calc_for_color(vf["target_hex"], optimizer=use_opt,
                                          seq_len=None if use_auto else len(vf["sequence"]),
                                          auto=use_auto, auto_threshold=thresh)
            if result:
                vf["sequence"] = result["sequence"]
                vf["sim_hex"]  = result["sim_hex"]
                vf["de"]       = result["de"]
        self._refresh_virtual_grid()
        messagebox.showinfo(self.t("dlg_note"),
                            self.t("recalc_all_done", n=len(self.virtual_fils)))

    # ── ΔE-ÜBERSICHT ───────────────────────────────────────────────────────────

    def open_de_overview(self):
        """Zeigt eine Tabelle aller virtuellen Köpfe mit ΔE, Qualität und WW-Info."""
        if not self.virtual_fils:
            messagebox.showinfo(self.t("dlg_note"), self.t("orca_no_virtual")); return
        import tkinter as _tk

        win = ctk.CTkToplevel(self)
        win.title(self.t("de_overview_title"))
        win.geometry("820x480")
        win.grab_set()

        ctk.CTkLabel(win, text=self.t("de_overview_title"),
                     font=("Segoe UI", 13, "bold"), text_color="#38bdf8").pack(pady=(14, 6))

        cols = [
            (self.t("de_overview_col_id"),      50),
            (self.t("de_overview_col_label"),   160),
            (self.t("de_overview_col_seq"),     110),
            ("Ziel",                             46),
            ("Sim",                              46),
            (self.t("de_overview_col_de"),       60),
            (self.t("de_overview_col_quality"), 110),
            (self.t("de_overview_col_tc"),       80),
        ]
        hdr = ctk.CTkFrame(win, fg_color="#1e293b", corner_radius=6)
        hdr.pack(fill="x", padx=14, pady=(0, 4))
        for c, (txt, w) in enumerate(cols):
            ctk.CTkLabel(hdr, text=txt, font=("Segoe UI", 10, "bold"),
                         text_color="#64748b", width=w).grid(row=0, column=c, padx=6, pady=5)

        sf = ctk.CTkScrollableFrame(win, fg_color="#0f172a", corner_radius=8)
        sf.pack(fill="both", expand=True, padx=14, pady=4)

        lh = safe_td(self.layer_height_entry.get()) if hasattr(self, "layer_height_entry") else 0.08
        for vf in self.virtual_fils:
            seq = vf["sequence"]
            de  = vf.get("de", 0.0)
            transitions = sum(1 for k in range(len(seq) - 1) if seq[k] != seq[k + 1])
            if de < DE_GOOD:
                q_text = "✓ " + ("ausgezeichnet" if self.lang == "de" else "excellent")
                q_col  = "#4ade80"
            elif de < DE_OK:
                q_text = "~ " + ("gut" if self.lang == "de" else "good")
                q_col  = "#fbbf24"
            else:
                q_text = "✗ " + ("sichtbar" if self.lang == "de" else "visible")
                q_col  = "#f87171"

            row_fr = ctk.CTkFrame(sf, fg_color="#1e293b", corner_radius=5)
            row_fr.pack(fill="x", padx=4, pady=2)
            vals = [
                (f"V{vf['vid']}",     50,  "#a78bfa", None),
                (vf.get("label",""),  160, "#e2e8f0", None),
                (seq,                 110, "#4ade80", None),
                ("",                   46, vf["target_hex"], vf["target_hex"]),
                ("",                   46, vf["sim_hex"],    vf["sim_hex"]),
                (f"{de:.1f}",          60, de_color(de),    None),
                (q_text,              110, q_col,            None),
                (str(transitions),     80,
                 "#f59e0b" if transitions > 2 else "#94a3b8", None),
            ]
            for c, (txt, w, col, bg) in enumerate(vals):
                kw = {"fg_color": bg, "corner_radius": 20} if bg else {}
                ctk.CTkLabel(row_fr, text=txt, width=w,
                             font=("Segoe UI", 10), text_color=col, **kw).grid(
                    row=0, column=c, padx=6, pady=4)

        ctk.CTkButton(win, text=self.t("exp_cancel"), fg_color="#374151",
                      command=win.destroy, height=34).pack(pady=8, padx=20, fill="x")

    # ── FARBREZEPT-EXPORT ──────────────────────────────────────────────────────

    def open_recipe_export(self):
        """Generiert ein menschenlesbares Farbrezept für alle virtuellen Köpfe."""
        if not self.virtual_fils:
            messagebox.showinfo(self.t("dlg_note"), self.t("orca_no_virtual")); return
        import tkinter as _tk

        lh  = safe_td(self.layer_height_entry.get()) if hasattr(self, "layer_height_entry") else 0.08
        now = datetime.now().strftime("%Y-%m-%d %H:%M")

        lines = [
            "=" * 60,
            f"  U1 FullSpectrum — Farbrezept",
            f"  Erstellt: {now}   Schichthöhe: {lh} mm",
            "=" * 60, "",
        ]

        slot_names = []
        for i, s in enumerate(self.slots):
            brand = s["brand"].get()
            color = s["color"].get()
            hx    = s["hex"].get()
            name  = f"T{i+1}: {brand} {color} ({hx})"
            slot_names.append(name)
            lines.append(f"  {name}")
        lines += ["", "-" * 60, ""]

        fils_hex  = {i+1: self.slots[i]["hex"].get() for i in range(4)}
        fils_name = {i+1: f"{self.slots[i]['brand'].get()} {self.slots[i]['color'].get()}"
                     for i in range(4)}

        for vf in self.virtual_fils:
            seq  = vf["sequence"]
            de   = vf.get("de", 0.0)
            n_f  = seq_filament_count(seq)
            runs = seq_to_runs(seq)
            cad  = calc_cadence(seq, lh)
            ids  = sorted(cad.keys())
            label = vf.get("label", f"V{vf['vid']}")
            q_str = ("ausgezeichnet" if de < DE_GOOD else "gut" if de < DE_OK else "sichtbar") \
                    if self.lang == "de" else \
                    ("excellent" if de < DE_GOOD else "good" if de < DE_OK else "visible")

            lines.append(f"V{vf['vid']}  \"{label}\"")
            lines.append(f"  Ziel: {vf['target_hex']}   Simuliert: {vf['sim_hex']}   ΔE = {de:.1f} ({q_str})")
            lines.append(f"  Sequenz: {''.join(seq)}   ({len(seq)} Schichten je Wiederholung)")
            for fid, cnt in runs:
                lines.append(f"    T{fid} {fils_name.get(fid, '')}  × {cnt} Schicht{'en' if cnt > 1 else ''}")
            if n_f == 1:
                lines.append(f"  → Reine Farbe — kein Dithering nötig")
            elif n_f == 2:
                a = round(cad.get(ids[0], lh), 3)
                b = round(cad.get(ids[1], lh) if len(ids) > 1 else lh, 3)
                lines.append(f"  → Cadence A = {a} mm   B = {b} mm   Step = {lh} mm")
            else:
                lines.append(f"  → Pattern Mode: {''.join(seq)}   Step = {lh} mm")
            lines.append("")

        lines.append("=" * 60)
        full_text = "\n".join(lines)

        win = ctk.CTkToplevel(self)
        win.title(self.t("recipe_title"))
        win.geometry("700x500")
        win.grab_set()

        ctk.CTkLabel(win, text=self.t("recipe_title"),
                     font=("Segoe UI", 13, "bold")).pack(pady=(14, 6))
        txt_box = _tk.Text(win, bg="#0f172a", fg="#e2e8f0",
                           font=("Courier New", 10), relief="flat", bd=0,
                           wrap="none", padx=10, pady=6)
        txt_box.pack(fill="both", expand=True, padx=16, pady=4)
        txt_box.insert("1.0", full_text)
        txt_box.configure(state="disabled")

        def _copy():
            self.clipboard_clear(); self.clipboard_append(full_text)
        ctk.CTkButton(win, text=self.t("recipe_copy_btn"), fg_color="#0f766e",
                      command=_copy, height=36).pack(pady=6, padx=20, fill="x")
        ctk.CTkButton(win, text=self.t("exp_cancel"), fg_color="#374151",
                      command=win.destroy, height=34).pack(pady=(0, 10), padx=20, fill="x")

    # ── MEHRFACH-ZIELFARBEN-OPTIMIZER ──────────────────────────────────────────

    def open_multitarget_optimizer(self):
        """Optimiert eine Sequenz für mehrere Zielfarben gleichzeitig (Min. Ø-ΔE)."""
        import tkinter as _tk
        from itertools import product as _product

        win = ctk.CTkToplevel(self)
        win.title(self.t("mt_title"))
        win.geometry("600x560")
        win.grab_set()

        ctk.CTkLabel(win, text=self.t("mt_title"),
                     font=("Segoe UI", 13, "bold"), text_color="#38bdf8").pack(pady=(14, 4))
        ctk.CTkLabel(win, text=self.t("mt_desc"),
                     font=("Segoe UI", 10), text_color="#94a3b8",
                     wraplength=560).pack(pady=(0, 8), padx=20)

        target_frame = ctk.CTkScrollableFrame(win, height=180, fg_color="#0f172a", corner_radius=8)
        target_frame.pack(fill="x", padx=16, pady=(0, 8))

        targets = []  # list of {"hex": str, "preview": CTkLabel, "entry": CTkEntry}

        def add_target_row(initial="#808080"):
            row = ctk.CTkFrame(target_frame, fg_color="transparent")
            row.pack(fill="x", pady=3)
            n = len(targets) + 1
            ctk.CTkLabel(row, text=self.t("mt_target", n=n),
                         width=60, font=("Segoe UI", 10)).pack(side="left", padx=(4, 4))
            entry = ctk.CTkEntry(row, width=100, font=("Courier New", 11),
                                 placeholder_text="#RRGGBB")
            entry.insert(0, initial)
            entry.pack(side="left", padx=(0, 4))
            prev = ctk.CTkLabel(row, text="", width=28, height=28,
                                 fg_color=initial, corner_radius=14)
            prev.pack(side="left", padx=(0, 4))

            def _pick(e=entry, p=prev):
                c = self._ask_color(initial=e.get(), title=self.t("target_picker_title"))
                if c:
                    e.delete(0, "end"); e.insert(0, c); p.configure(fg_color=c)
            ctk.CTkButton(row, text="🎨", width=28, height=28, fg_color="#334155",
                          command=_pick).pack(side="left")
            targets.append({"entry": entry, "preview": prev})

        add_target_row("#FF4444")
        add_target_row("#44AAFF")

        ctk.CTkButton(win, text=self.t("mt_add_target"), fg_color="#1e3a5f",
                      command=add_target_row, height=30).pack(pady=4, padx=20, fill="x")

        # Sequenzlänge
        len_row = ctk.CTkFrame(win, fg_color="transparent")
        len_row.pack(fill="x", padx=20, pady=4)
        ctk.CTkLabel(len_row, text="Max. Länge:", font=("Segoe UI", 10)).pack(side="left")
        len_var = ctk.StringVar(value="5")
        ctk.CTkOptionMenu(len_row, variable=len_var,
                          values=[str(n) for n in range(1, MAX_SEQ_LEN + 1)],
                          width=80).pack(side="left", padx=8)

        result_lbl = ctk.CTkLabel(win, text="", font=("Segoe UI", 11),
                                   text_color="#4ade80", wraplength=560)
        result_lbl.pack(pady=6, padx=20)

        _mt_result = {"seq": None, "de": None, "sim_hex": None}

        def _optimize():
            hexes = []
            for t in targets:
                raw = t["entry"].get().strip()
                if not raw.startswith("#"): raw = "#" + raw
                if len(raw) == 7:
                    t["preview"].configure(fg_color=raw)
                    hexes.append(raw)
            if len(hexes) < 2:
                messagebox.showinfo(self.t("dlg_note"), self.t("mt_no_targets")); return

            fils    = self._get_fils()
            fids    = [f["id"] for f in fils]
            t_labs  = [rgb_to_lab(hex_to_rgb(h)) for h in hexes]
            max_len = int(len_var.get())

            best_seq, best_avg_de, best_sim = None, float("inf"), None
            for length in range(1, max_len + 1):
                for combo in _product(fids, repeat=length):
                    seq_list = list(combo)
                    sim_lab  = self._simulate_mix(seq_list, fils)
                    avg_de   = sum(delta_e(sim_lab, tl) for tl in t_labs) / len(t_labs)
                    if avg_de < best_avg_de:
                        best_avg_de = avg_de
                        best_seq    = "".join(map(str, seq_list))
                        best_sim    = lab_to_hex(sim_lab)

            _mt_result["seq"]     = best_seq
            _mt_result["de"]      = best_avg_de
            _mt_result["sim_hex"] = best_sim
            _mt_result["hexes"]   = hexes
            result_lbl.configure(
                text=self.t("mt_result", seq=best_seq, de=best_avg_de))
            add_btn.configure(state="normal")

        add_btn_frame = ctk.CTkFrame(win, fg_color="transparent")
        add_btn_frame.pack(fill="x", padx=20, pady=4)
        ctk.CTkButton(add_btn_frame, text=self.t("mt_calc"), fg_color="#7c3aed",
                      hover_color="#6d28d9", height=40,
                      command=_optimize).pack(side="left", expand=True, fill="x", padx=(0, 4))

        add_btn = ctk.CTkButton(add_btn_frame, text=self.t("mt_add_btn"),
                                fg_color="#065f46", hover_color="#047857",
                                height=40, state="disabled",
                                command=lambda: self._mt_add_result(_mt_result, win))
        add_btn.pack(side="left", expand=True, fill="x", padx=(4, 0))

        ctk.CTkButton(win, text=self.t("exp_cancel"), fg_color="#374151",
                      command=win.destroy, height=34).pack(pady=(4, 12), padx=20, fill="x")

    def _mt_add_result(self, result, win):
        """Fügt das Multi-Target-Ergebnis als virtuellen Kopf hinzu."""
        if not result.get("seq"): return
        n_virts = len(self.virtual_fils)
        max_v   = self.settings.get("max_virtual", MAX_VIRTUAL)
        if n_virts >= max_v:
            messagebox.showinfo(self.t("dlg_max_virtual"),
                                self.t("dlg_max_virtual_msg", max_v=max_v)); return
        vid     = 5 + n_virts
        # Durchschnitt der Zielfarben als target_hex
        hexes   = result.get("hexes", [result.get("sim_hex", "#808080")])
        r_avg   = round(sum(hex_to_rgb(h)[0] for h in hexes) / len(hexes))
        g_avg   = round(sum(hex_to_rgb(h)[1] for h in hexes) / len(hexes))
        b_avg   = round(sum(hex_to_rgb(h)[2] for h in hexes) / len(hexes))
        avg_hex = f"#{r_avg:02X}{g_avg:02X}{b_avg:02X}"
        self.virtual_undo.append(copy.deepcopy(self.virtual_fils))
        self.virtual_fils.append({
            "vid":        vid,
            "target_hex": avg_hex,
            "sequence":   result["seq"],
            "sim_hex":    result["sim_hex"],
            "de":         result["de"],
            "label":      f"MultiTarget V{vid}",
        })
        self._refresh_virtual_grid()
        win.destroy()

    # ── 3MF ASSISTENT ──────────────────────────────────────────────────────────

    def open_3mf_assistant(self):
        # Singleton: focus existing window if open
        if self._3mf_win is not None and self._3mf_win.winfo_exists():
            self._3mf_win.focus_force()
            self._3mf_win.lift()
            return

        path = filedialog.askopenfilename(
            filetypes=[(self.t("3mf_filetypes"), "*.3mf"), ("*", "*.*")],
            title=self.t("open_3mf_title"))
        if not path: return

        colors, err = parse_3mf_colors(path)
        if not colors:
            messagebox.showinfo(self.t("dlg_3mf_title"),
                f"{self.t('dlg_3mf_no_colors_fallback') if not err else err}"); return

        win = ctk.CTkToplevel(self)
        self._3mf_win = win
        win.protocol("WM_DELETE_WINDOW", lambda: (setattr(self, "_3mf_win", None), win.destroy()))
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
            self._set_status(self.t("status_3mf", n=added), 4000)

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

    # ── 3MF FARB-WIZARD ────────────────────────────────────────────────────────

    def _get_all_library_fils(self):
        """Returns all filaments from DEFAULT_LIBRARY + user DB as list of dicts with lab."""
        result = []
        all_data = {}
        for brand, fils in DEFAULT_LIBRARY.items():
            all_data[brand] = fils[:]
        for brand, fils in self.library.items():
            if brand not in all_data:
                all_data[brand] = []
            all_data[brand] = all_data[brand] + [f for f in fils if f not in all_data[brand]]
        for brand, fils in all_data.items():
            for f in fils:
                try:
                    lab = rgb_to_lab(hex_to_rgb(f["hex"]))
                    result.append({
                        "brand": brand,
                        "name": f["name"],
                        "hex": f["hex"],
                        "td": f.get("td", DEFAULT_TD),
                        "lab": lab,
                    })
                except Exception:
                    pass
        return result

    def open_3mf_wizard(self):
        if self._3mf_wizard_win and self._3mf_wizard_win.winfo_exists():
            self._3mf_wizard_win.focus_force()
            return
        win = ThreeMFWizard(self)
        self._3mf_wizard_win = win
        win.protocol("WM_DELETE_WINDOW",
                     lambda: (setattr(self, "_3mf_wizard_win", None), win.destroy()))

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
        lh_val = self.layer_height_entry.get() if hasattr(self, "layer_height_entry") else "0.08"
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
                pat   = ",".join(v["sequence"])
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
                            pat_str  = ",".join(v["sequence"])
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
                self._set_status(self.t("status_exported", f=os.path.basename(path)), 4000)
                # Show slicer hint (Change 9)
                if scope == "virtual":
                    self._show_export_hint_for_virtual()
                elif scope == "single" and hasattr(self, "last_result") and self.last_result:
                    self._show_export_hint_for_seq(self.last_result.get("sequence", ""))
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

    def _show_export_hint_for_seq(self, seq):
        """Show post-export slicer hint for a single sequence (Change 9)."""
        if not hasattr(self, '_export_hint_label') or not seq:
            return
        lh = safe_td(self.layer_height_entry.get()) if hasattr(self, "layer_height_entry") else 0.08
        n_f = seq_filament_count(seq)
        cad = calc_cadence(seq, lh)
        ids = sorted(cad.keys())
        if n_f <= 2 and len(ids) >= 2:
            a = round(cad.get(ids[0], lh), 3)
            b = round(cad.get(ids[1], lh), 3)
            hint = f"OrcaSlicer: Others → Dithering → Cadence Height  A={a}mm / B={b}mm"
        elif n_f >= 3:
            hint = f"OrcaSlicer: Others → Dithering → Pattern Mode  {''.join(seq)}"
        else:
            hint = ""
        if hint:
            self._export_hint_label.configure(text=hint)
            self.after(8000, lambda: self._export_hint_label.configure(text="") if self._export_hint_label.winfo_exists() else None)

    def _show_export_hint_for_virtual(self):
        """Show post-export slicer hint for all virtual heads (Change 9)."""
        if not hasattr(self, '_export_hint_label') or not self.virtual_fils:
            return
        # Show hint for first virtual head as example
        vf = self.virtual_fils[0]
        self._show_export_hint_for_seq(vf.get("sequence", ""))

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

        # FullSpectrum-Hinweis-Label (wird je nach Slicer ein-/ausgeblendet)
        def _is_fullspectrum(path):
            p = path.lower()
            return "snapmaker" in p or "fullspectrum" in p or "full_spectrum" in p

        fs_hint = ctk.CTkLabel(
            win,
            text=self.t("orca_fs_hint"),
            font=("Segoe UI", 9), text_color="#f59e0b",
            wraplength=520, justify="left")

        # Scope
        ctk.CTkLabel(win, text="Exportieren:", font=("Segoe UI", 11, "bold")).pack(
            pady=(10, 2))
        initial_scope = "phys" if _is_fullspectrum(path_var.get()) else "both"
        scope_var = ctk.StringVar(value=initial_scope)
        sf = ctk.CTkFrame(win, fg_color="transparent"); sf.pack()
        ctk.CTkRadioButton(sf, text=self.t("orca_scope_phys"),
                           variable=scope_var, value="phys").pack(anchor="w", padx=20, pady=1)
        ctk.CTkRadioButton(sf, text=self.t("orca_scope_virt"),
                           variable=scope_var, value="virt").pack(anchor="w", padx=20, pady=1)
        ctk.CTkRadioButton(sf, text=self.t("orca_scope_both"),
                           variable=scope_var, value="both").pack(anchor="w", padx=20, pady=1)

        def _update_fs_hint(*_):
            if _is_fullspectrum(path_var.get()):
                scope_var.set("phys")
                fs_hint.pack(padx=20, pady=(2, 0), anchor="w")
            else:
                fs_hint.pack_forget()
        path_var.trace_add("write", _update_fs_hint)
        _update_fs_hint()  # initial state

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
        lh_val = self.layer_height_entry.get() if hasattr(self, "layer_height_entry") else "0.08"
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
            lh      = safe_td(lh_entry.get()) if lh_entry.get().strip() else 0.08

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

    # ── MULTI-GRADIENT DIALOG (Change 7) ──────────────────────────────────────

    def open_multi_gradient_dialog(self):
        """Multi-Gradient: weighted blend of all 4 slots as virtual head."""
        fils = self._get_fils()
        win = ctk.CTkToplevel(self)
        win.title(self.t("multi_gradient_title"))
        win.geometry("480x440")
        win.grab_set()

        ctk.CTkLabel(win, text=self.t("multi_gradient_title"),
                     font=("Segoe UI", 14, "bold")).pack(pady=(14, 6))
        ctk.CTkLabel(win, text=self.t("multi_gradient_desc"),
                     font=("Segoe UI", 10), text_color="#64748b").pack(pady=(0, 10))

        weight_vars = []
        swatch_labels = []
        rows_frame = ctk.CTkFrame(win, fg_color="transparent")
        rows_frame.pack(fill="x", padx=24)

        for i, f in enumerate(fils):
            row = ctk.CTkFrame(rows_frame, fg_color="transparent")
            row.pack(fill="x", pady=3)
            sw = ctk.CTkLabel(row, text="", width=32, height=32,
                              fg_color=f["hex"], corner_radius=6)
            sw.pack(side="left", padx=(0, 8))
            ctk.CTkLabel(row, text=f"T{f['id']}", font=("Segoe UI", 10, "bold"),
                         width=30).pack(side="left")
            wv = ctk.StringVar(value="25")
            weight_vars.append(wv)
            sp = ctk.CTkEntry(row, textvariable=wv, width=60)
            sp.pack(side="left", padx=8)
            ctk.CTkLabel(row, text="%", font=("Segoe UI", 9)).pack(side="left")
            swatch_labels.append(sw)

        preview_label = ctk.CTkLabel(win, text="", width=80, height=32,
                                     fg_color="#808080", corner_radius=8)
        preview_label.pack(pady=8)

        def _update_preview(*a):
            try:
                ws = [(f["id"], float(wv.get() or "0")) for f, wv in zip(fils, weight_vars)]
                seq = build_weighted_gradient_sequence(ws, max_len=MAX_SEQ_LEN)
                if not seq:
                    return
                sim_lab = self._simulate_mix([str(x) for x in seq], fils)
                preview_label.configure(fg_color=lab_to_hex(sim_lab))
            except Exception:
                pass

        for wv in weight_vars:
            wv.trace_add("write", _update_preview)
        _update_preview()

        def _auto_balance():
            n = len(weight_vars)
            base = 100 // n
            rem = 100 - base * n
            for i, wv in enumerate(weight_vars):
                wv.set(str(base + (1 if i < rem else 0)))

        ctk.CTkButton(win, text=self.t("multi_gradient_auto"), fg_color="#1e3a5f",
                      command=_auto_balance, height=32).pack(pady=(0, 4), padx=24, fill="x")

        def _add_as_virtual():
            try:
                ws = [(f["id"], float(wv.get() or "0")) for f, wv in zip(fils, weight_vars)]
                seq = build_weighted_gradient_sequence(ws, max_len=MAX_SEQ_LEN)
                if not seq:
                    return
                seq_str = "".join(str(x) for x in seq)
                sim_lab = self._simulate_mix(list(seq_str), fils)
                sim_hex = lab_to_hex(sim_lab)
                t_lab = sim_lab
                dv = delta_e(sim_lab, t_lab)
                result = {
                    "target_hex": sim_hex,
                    "sequence": seq_str,
                    "sim_hex": sim_hex,
                    "de": 0.0,
                    "seq_len": len(seq_str),
                }
                self.add_to_virtual(result)
                win.destroy()
            except Exception as e:
                messagebox.showerror(self.t("dlg_error"), str(e))

        ctk.CTkButton(win, text=self.t("multi_gradient_add"), fg_color="#16a34a",
                      command=_add_as_virtual, height=42,
                      font=("Segoe UI", 13, "bold")).pack(pady=(4, 4), padx=24, fill="x")
        ctk.CTkButton(win, text=self.t("exp_cancel"), fg_color="#334155",
                      command=win.destroy, height=36).pack(padx=24, fill="x")

    # ── SLOT-VERGLEICH (Change 10) ─────────────────────────────────────────────

    def open_slot_compare(self):
        """Opens a window showing ΔE impact of replacing a slot with an alternative filament."""
        if not self.virtual_fils:
            messagebox.showinfo(self.t("dlg_note"), self.t("orca_no_virtual")); return
        win = ctk.CTkToplevel(self)
        win.title(self.t("slot_compare_title"))
        win.geometry("700x520")
        win.grab_set()
        ctk.CTkLabel(win, text=self.t("slot_compare_title"),
                     font=("Segoe UI", 13, "bold"), text_color="#38bdf8").pack(pady=(14, 8))

        ctrl = ctk.CTkFrame(win, fg_color="#1e293b", corner_radius=8)
        ctrl.pack(fill="x", padx=16, pady=(0, 8))

        # Slot selector
        slot_f = ctk.CTkFrame(ctrl, fg_color="transparent")
        slot_f.pack(side="left", padx=12, pady=8)
        ctk.CTkLabel(slot_f, text=self.t("slot_compare_slot"),
                     font=("Segoe UI", 10)).pack(anchor="w")
        slot_var = ctk.StringVar(value="T2")
        slot_opts = [f"T{i+1}" for i in range(4)]
        ctk.CTkOptionMenu(slot_f, variable=slot_var, values=slot_opts, width=80).pack()

        # Alternative filament selector
        alt_f = ctk.CTkFrame(ctrl, fg_color="transparent")
        alt_f.pack(side="left", padx=12, pady=8, expand=True, fill="x")
        ctk.CTkLabel(alt_f, text=self.t("slot_compare_alt"),
                     font=("Segoe UI", 10)).pack(anchor="w")
        all_fils = [(brand, fil) for brand, fils in self.library.items() for fil in fils]
        alt_opts = [f"{b} — {f['name']}" for b, f in all_fils] if all_fils else ["(none)"]
        alt_var = ctk.StringVar(value=alt_opts[0] if alt_opts else "")
        alt_menu = ctk.CTkOptionMenu(alt_f, variable=alt_var, values=alt_opts, width=260)
        alt_menu.pack()

        # Results table
        hdr = ctk.CTkFrame(win, fg_color="#1e293b", corner_radius=6)
        hdr.pack(fill="x", padx=16, pady=(0, 2))
        for col, (txt, w) in enumerate([(self.t("slot_compare_col_vid"), 60),
                                         (self.t("slot_compare_col_cur"), 100),
                                         (self.t("slot_compare_col_new"), 100),
                                         (self.t("slot_compare_col_delta"), 80)]):
            ctk.CTkLabel(hdr, text=txt, font=("Segoe UI", 10, "bold"),
                         text_color="#64748b", width=w).grid(row=0, column=col, padx=8, pady=6, sticky="w")

        scroll = ctk.CTkScrollableFrame(win, height=300, fg_color="#0f172a")
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        def run_compare():
            for w in scroll.winfo_children(): w.destroy()
            slot_idx = int(slot_var.get()[1]) - 1  # T1→0, T2→1, ...
            # Find alternative filament
            alt_label = alt_var.get()
            alt_fil = None
            for brand, fil in all_fils:
                if f"{brand} — {fil['name']}" == alt_label:
                    alt_fil = fil
                    break
            if alt_fil is None: return

            fils_current = self._get_fils()
            # Build alternative filament set
            alt_hex = alt_fil.get("hex", "#888888")
            alt_td = alt_fil.get("td", DEFAULT_TD)
            alt_lab = rgb_to_lab(hex_to_rgb(alt_hex))

            fils_alt = []
            for f in fils_current:
                if f["id"] == slot_idx + 1:
                    fils_alt.append({"id": f["id"], "hex": alt_hex, "td": alt_td, "lab": alt_lab})
                else:
                    fils_alt.append(f)

            for vf in self.virtual_fils:
                seq = vf["sequence"]
                target_hex = vf.get("target_hex", "")
                if not target_hex: continue
                t_lab = rgb_to_lab(hex_to_rgb(target_hex))
                cur_de = vf.get("de", 0.0)
                # Simulate with alternative filament (convert string seq to int list)
                seq_ids = [int(c) for c in seq]
                new_sim = self._simulate_mix(seq_ids, fils_alt)
                new_de = delta_e(new_sim, t_lab)
                delta = new_de - cur_de

                row = ctk.CTkFrame(scroll, fg_color="#1e293b", corner_radius=6)
                row.pack(fill="x", padx=4, pady=2)
                ctk.CTkLabel(row, text=f"V{vf['vid']}", width=60,
                             font=("Segoe UI", 10, "bold")).grid(row=0, column=0, padx=8, pady=6)
                ctk.CTkLabel(row, text=f"{cur_de:.1f}", width=100,
                             text_color=de_color(cur_de)).grid(row=0, column=1, padx=8)
                ctk.CTkLabel(row, text=f"{new_de:.1f}", width=100,
                             text_color=de_color(new_de)).grid(row=0, column=2, padx=8)
                delta_col = "#4ade80" if delta < 0 else "#f87171" if delta > 0 else "#94a3b8"
                delta_txt = f"+{delta:.1f}" if delta > 0 else f"{delta:.1f}"
                ctk.CTkLabel(row, text=delta_txt, width=80, font=("Segoe UI", 10, "bold"),
                             text_color=delta_col).grid(row=0, column=3, padx=8)

        ctk.CTkButton(ctrl, text="▶  Vergleichen", fg_color="#2563eb",
                      command=run_compare, height=36).pack(side="right", padx=12, pady=8)
        ctk.CTkButton(win, text=self.t("exp_cancel"), fg_color="#334155",
                      command=win.destroy, height=34).pack(padx=16, pady=(0, 10), fill="x")

    # ── SLOT-OPTIMIZER ────────────────────────────────────────────────────────

    def open_slot_optimizer(self):
        """Findet die 4 Filamente aus der Bibliothek mit dem größten Farbraum."""
        win = ctk.CTkToplevel(self)
        win.title(self.t("slot_opt_title"))
        win.geometry("520x480")
        win.grab_set()
        ctk.CTkLabel(win, text=self.t("slot_opt_title"),
                     font=("Segoe UI", 13, "bold")).pack(pady=(14, 4))
        ctk.CTkLabel(win, text=self.t("slot_opt_desc"),
                     font=("Segoe UI", 10), text_color="#64748b",
                     wraplength=460).pack(pady=(0, 8))

        # Zielfarben-Option
        use_virtual_var = ctk.BooleanVar(value=bool(self.virtual_fils))
        ctk.CTkCheckBox(win, text=self.t("slot_opt_use_virtual"),
                        variable=use_virtual_var).pack(anchor="w", padx=24)

        progress_lbl = ctk.CTkLabel(win, text="", font=("Segoe UI", 10),
                                     text_color="#64748b")
        progress_lbl.pack(pady=4)
        pb = ctk.CTkProgressBar(win); pb.pack(padx=30, fill="x"); pb.set(0)

        result_frame = ctk.CTkScrollableFrame(win, height=180, fg_color="#0f172a")
        result_frame.pack(fill="x", padx=20, pady=8)

        self._opt_best = None

        def run_optimizer():
            import threading

            # Alle Filamente flach aus der Library
            all_fils = []
            for brand, entries in self.library.items():
                if brand == "Eigene Favoriten": continue
                for e in entries:
                    if not e.get("hex") or len(e["hex"]) != 7: continue
                    all_fils.append({
                        "brand": brand, "name": e["name"],
                        "hex": e["hex"], "td": e.get("td", 5.0),
                        "lab": rgb_to_lab(hex_to_rgb(e["hex"]))
                    })

            # Zielfarben
            if use_virtual_var.get() and self.virtual_fils:
                targets = [rgb_to_lab(hex_to_rgb(vf["target_hex"]))
                           for vf in self.virtual_fils]
            else:
                # Standardmäßig gut verteilte Testfarben im Farbraum
                targets = [rgb_to_lab(hex_to_rgb(h)) for h in [
                    "#FF0000","#FF8000","#FFFF00","#00FF00",
                    "#00FFFF","#0000FF","#FF00FF","#FFFFFF",
                    "#808080","#000000"]]

            n = len(all_fils)
            k = len(targets)

            def worker():
                # --- Phase 1: Pre-compute distance matrix (n × k) ---
                dist = [[delta_e(targets[t], all_fils[i]["lab"])
                         for t in range(k)]
                        for i in range(n)]

                def _upd_ui(frac, msg):
                    def _cb():
                        try:
                            if win.winfo_exists():
                                pb.set(frac)
                                progress_lbl.configure(text=msg)
                        except Exception:
                            pass
                    win.after(0, _cb)

                _upd_ui(0.15, "Phase 1: Distanzmatrix …")

                def combo_score(chosen_idx):
                    total = 0.0
                    for t in range(k):
                        total += min(dist[i][t] for i in chosen_idx)
                    return total / k

                # --- Phase 2: Greedy selection ---
                chosen = []
                remaining = list(range(n))
                for step in range(4):
                    _upd_ui(0.15 + step * 0.15,
                            f"Phase 2: Greedy Schritt {step+1}/4 …")
                    # current best coverage for each target
                    if chosen:
                        cur_best = [min(dist[i][t] for i in chosen) for t in range(k)]
                    else:
                        cur_best = [float("inf")] * k

                    best_i, best_sc = None, float("inf")
                    for i in remaining:
                        sc = sum(min(cur_best[t], dist[i][t]) for t in range(k)) / k
                        if sc < best_sc:
                            best_sc, best_i = sc, i
                    chosen.append(best_i)
                    remaining.remove(best_i)

                _upd_ui(0.75, "Phase 3: Local Search …")

                # --- Phase 3: Local search (swap improvements) ---
                improved = True
                cur_score = combo_score(chosen)
                while improved:
                    improved = False
                    for ci in range(4):
                        old_c = chosen[ci]
                        for r in remaining:
                            chosen[ci] = r
                            sc = combo_score(chosen)
                            if sc < cur_score - 1e-9:
                                cur_score = sc
                                remaining.remove(r)
                                remaining.append(old_c)
                                old_c = r
                                improved = True
                                break
                            else:
                                chosen[ci] = old_c
                        if improved:
                            break

                self._opt_best = ([all_fils[i] for i in chosen], cur_score)
                def _done():
                    try:
                        if win.winfo_exists():
                            show_result()
                    except Exception:
                        pass
                win.after(0, _done)

            def show_result():
                pb.set(1.0)
                progress_lbl.configure(text=self.t("slot_opt_done"))
                for w in result_frame.winfo_children(): w.destroy()
                if not self._opt_best: return
                combo, score = self._opt_best
                ctk.CTkLabel(result_frame,
                              text=self.t("slot_opt_result", de=score),
                              font=("Segoe UI", 10, "bold"),
                              text_color="#4ade80").pack(anchor="w", padx=8, pady=4)
                for f in combo:
                    row = ctk.CTkFrame(result_frame, fg_color="#1e293b", corner_radius=6)
                    row.pack(fill="x", padx=4, pady=2)
                    ctk.CTkLabel(row, text="", width=28, height=28,
                                  fg_color=f["hex"], corner_radius=5).pack(
                        side="left", padx=6, pady=4)
                    ctk.CTkLabel(row, text=f"{f['brand']} — {f['name']}  {f['hex']}  TD={f['td']}",
                                  font=("Segoe UI", 10)).pack(side="left", padx=4)
                apply_btn.pack(pady=(4, 2), padx=20, fill="x")

            threading.Thread(target=worker, daemon=True).start()

        def apply_best():
            if not self._opt_best: return
            combo, _ = self._opt_best
            for i, f in enumerate(combo[:4]):
                s = self.slots[i]
                if f["brand"] in self.library:
                    s["brand"].set(f["brand"])
                    self.update_menu(i)
                s["hex"].delete(0, "end"); s["hex"].insert(0, f["hex"])
                s["td"].delete(0, "end");  s["td"].insert(0, str(f["td"]))
                try: s["preview"].configure(fg_color=f["hex"])
                except: pass
                # Filament-Name setzen falls vorhanden
                try: s["color"].set(f["name"])
                except: pass
            self._schedule_gamut_update()
            win.destroy()

        apply_btn = ctk.CTkButton(win, text=self.t("slot_opt_apply"),
                                   fg_color="#15803d", height=38)
        apply_btn.configure(command=apply_best)

        ctk.CTkButton(win, text="▶  " + self.t("slot_opt_title").split("—")[0].strip(),
                      fg_color="#2563eb", height=40,
                      font=("Segoe UI", 12, "bold"),
                      command=run_optimizer).pack(pady=(0, 4), padx=20, fill="x")
        ctk.CTkButton(win, text=self.t("exp_cancel"), fg_color="#334155",
                      command=win.destroy, height=34).pack(padx=20, fill="x")

    # ── 3MF ZURÜCKSCHREIBEN ────────────────────────────────────────────────────

    def write_3mf_colors(self):
        """Öffnet 3MF, liest alle Objekte + ihre Extruder-Zuweisungen,
        zeigt Remap-Dialog und schreibt gezielt filament_settings + model_settings."""
        import xml.etree.ElementTree as ET

        path = filedialog.askopenfilename(
            filetypes=[("3MF","*.3mf"),("*","*.*")], title=self.t("3mf_write_title"))
        if not path: return

        # ── 3MF analysieren ──────────────────────────────────────────────────
        try:
            with zipfile.ZipFile(path,"r") as zin:
                names = zin.namelist()
                # model_settings lesen → Objekte + Extruder
                ms_raw = zin.read("Metadata/model_settings.config").decode("utf-8")
                ms_root = ET.fromstring(ms_raw)
                # {extruder_str: [obj_name, ...]}
                ext_objects = {}
                for obj in ms_root.findall("object"):
                    obj_name = next((m.get("value") for m in obj.findall("metadata")
                                     if m.get("key")=="name"), "?")
                    ext_val  = next((m.get("value") for m in obj.findall("metadata")
                                     if m.get("key")=="extruder"), "1")
                    ext_objects.setdefault(ext_val, []).append(obj_name)

                # filament_settings_X vorhanden?
                fil_cfgs = sorted([n for n in names
                                   if re.match(r"Metadata/filament_settings_\d+\.config", n)])
        except Exception as e:
            messagebox.showerror(self.t("dlg_error"), str(e)); return

        # ── Remap-Dialog ─────────────────────────────────────────────────────
        win = ctk.CTkToplevel(self)
        win.title(self.t("remap_title"))
        win.geometry("680x520")
        win.grab_set()

        ctk.CTkLabel(win, text=self.t("remap_title"),
                     font=("Segoe UI",14,"bold"), text_color="#38bdf8").pack(pady=(16,4))
        ctk.CTkLabel(win, text=os.path.basename(path),
                     font=("Segoe UI",9), text_color="#64748b").pack()

        sf = ctk.CTkScrollableFrame(win, fg_color="#0f172a", height=300)
        sf.pack(fill="x", padx=16, pady=(10,4))

        ctk.CTkLabel(sf, text=self.t("remap_col_hdr"),
                     font=("Segoe UI",10,"bold"), text_color="#94a3b8").pack(anchor="w", padx=8, pady=(6,2))

        # Optionen für Dropdown: T1–T4 + V5+
        slot_opts = [f"T{i+1} — {self.slots[i]['color'].get() or self.slots[i]['hex'].get()}"
                     for i in range(4)]
        virt_opts = [f"V{vf['vid']} — {vf.get('label','?')} (ΔE {vf.get('de',0):.1f})"
                     for vf in self.virtual_fils]
        all_opts  = slot_opts + virt_opts + [self.t("remap_keep")]

        remap_vars = {}  # ext_str → StringVar
        sorted_exts = sorted(ext_objects.keys(), key=lambda x: int(x) if x.isdigit() else 99)

        for ext in sorted_exts:
            row = ctk.CTkFrame(sf, fg_color="#1e293b", corner_radius=6)
            row.pack(fill="x", padx=4, pady=3)
            obj_list = ", ".join(ext_objects[ext])[:55]
            ctk.CTkLabel(row, text=f"Extruder {ext}",
                         font=("Segoe UI",10,"bold"), width=90).pack(side="left", padx=8)
            ctk.CTkLabel(row, text=obj_list, font=("Segoe UI",9),
                         text_color="#94a3b8").pack(side="left", padx=4, expand=True, anchor="w")
            var = ctk.StringVar(value=all_opts[min(int(ext)-1, len(all_opts)-1)]
                                if ext.isdigit() and int(ext)-1 < len(all_opts) else all_opts[-1])
            remap_vars[ext] = var
            ctk.CTkOptionMenu(row, variable=var, values=all_opts, width=240).pack(
                side="right", padx=8, pady=4)

        lh = safe_td(self.layer_height_entry.get()) if hasattr(self,"layer_height_entry") else 0.08

        def do_write():
            save_path = filedialog.asksaveasfilename(
                defaultextension=".3mf", filetypes=[("3MF","*.3mf")],
                initialfile=os.path.splitext(os.path.basename(path))[0]+"_u1.3mf",
                title="Speichern als")
            if not save_path: return

            # Create backup if file already exists
            backup_info = ""
            if os.path.exists(save_path):
                import shutil as _shutil_bak
                backup_path = save_path + ".bak"
                try:
                    _shutil_bak.copy2(save_path, backup_path)
                    backup_info = f"\n(Backup: {backup_path})"
                except Exception:
                    pass

            # ext → neue Slot/V-ID ermitteln
            def parse_choice(choice, ext_str):
                if choice == self.t("remap_keep"):
                    return None, None  # unverändert
                if choice.startswith("T"):
                    slot_i = int(choice[1]) - 1
                    return int(ext_str), slot_i + 1   # neue Ext = T1→1, T2→2, …
                if choice.startswith("V"):
                    vid = int(choice.split()[0][1:])
                    return int(ext_str), vid
                return None, None

            ext_map = {}  # {old_ext_int: new_ext_int}
            vf_for_ext = {}  # {new_ext_int: vf-dict or None}
            slot_for_ext = {}  # {new_ext_int: slot_index}

            for ext, var in remap_vars.items():
                choice = var.get()
                old_e = int(ext) if ext.isdigit() else None
                if old_e is None: continue
                if choice == self.t("remap_keep"):
                    continue
                if choice.startswith("T"):
                    slot_i = int(choice[1]) - 1
                    new_e = slot_i + 1
                    ext_map[old_e] = new_e
                    slot_for_ext[new_e] = slot_i
                elif choice.startswith("V"):
                    vid = int(choice.split()[0][1:])
                    new_e = vid
                    ext_map[old_e] = new_e
                    vf_match = next((v for v in self.virtual_fils if v["vid"]==vid), None)
                    vf_for_ext[new_e] = vf_match

            try:
                import shutil, tempfile
                with zipfile.ZipFile(path,"r") as zin:
                    with tempfile.NamedTemporaryFile(suffix=".3mf", delete=False) as tmp:
                        tmp_path = tmp.name
                    new_filament_cfgs = {}  # fname → bytes to add

                    with zipfile.ZipFile(tmp_path,"w", zipfile.ZIP_DEFLATED) as zout:
                        for item in zin.infolist():
                            data = zin.read(item.filename)
                            fname = item.filename

                            # model_settings.config — extruder remappen
                            if fname == "Metadata/model_settings.config":
                                try:
                                    root2 = ET.fromstring(data.decode("utf-8"))
                                    for obj in root2.findall("object"):
                                        for meta in obj.findall("metadata"):
                                            if meta.get("key") == "extruder":
                                                old_e = int(meta.get("value","1"))
                                                if old_e in ext_map:
                                                    meta.set("value", str(ext_map[old_e]))
                                    data = ET.tostring(root2, encoding="unicode").encode("utf-8")
                                except Exception: pass

                            # filament_settings_X.config für neue V-Köpfe/Slots generieren
                            elif re.match(r"Metadata/filament_settings_\d+\.config", fname):
                                # Original beibehalten — neue kommen später separat
                                pass

                            zout.writestr(item, data)

                        # Neue filament_settings für remappte Extruder hinzufügen
                        def _dithering_note(vf):
                            seq = vf["sequence"]
                            n_f = seq_filament_count(seq)
                            cad = calc_cadence(seq, lh)
                            ids_s = sorted(cad.keys())
                            if n_f == 1:
                                return f"U1 FullSpectrum | Reine Farbe T{seq[0]} | dE={vf.get('de',0):.1f}"
                            elif n_f == 2:
                                a = round(cad.get(ids_s[0],lh),3)
                                b = round(cad.get(ids_s[1],lh) if len(ids_s)>1 else lh,3)
                                return (f"U1 FullSpectrum | Seq:{''.join(seq)} | "
                                        f"Cadence A={a}mm B={b}mm | Step={lh}mm | dE={vf.get('de',0):.1f}")
                            else:
                                return (f"U1 FullSpectrum | Seq:{''.join(seq)} | "
                                        f"Pattern={''.join(seq)} | Step={lh}mm | dE={vf.get('de',0):.1f}")

                        base_cfg = {"type":"filament","from":"User","instantiation":"true",
                                    "filament_type":["PLA"],"filament_diameter":["1.75"],
                                    "filament_flow_ratio":["1"],"compatible_printers":[]}

                        for new_e, vf in vf_for_ext.items():
                            if vf is None: continue
                            sim_hex = vf.get("sim_hex", vf.get("target_hex","#888888"))
                            if not sim_hex.startswith("#"): sim_hex = "#"+sim_hex
                            cfg = dict(base_cfg)
                            cfg["filament_settings_id"] = [f"U1-V{vf['vid']}"]
                            cfg["filament_vendor"] = ["U1 FullSpectrum"]
                            cfg["default_filament_colour"] = [sim_hex]
                            cfg["filament_notes"] = [_dithering_note(vf)]
                            fname_new = f"Metadata/filament_settings_{new_e}.config"
                            zout.writestr(fname_new,
                                          json.dumps(cfg, indent=4, ensure_ascii=False).encode("utf-8"))

                        for new_e, slot_i in slot_for_ext.items():
                            s = self.slots[slot_i]
                            sim_hex = s["hex"].get().strip() or "#888888"
                            if not sim_hex.startswith("#"): sim_hex = "#"+sim_hex
                            cfg = dict(base_cfg)
                            cfg["filament_settings_id"] = [f"U1-T{slot_i+1}"]
                            cfg["filament_vendor"] = [s["brand"].get() or "U1"]
                            cfg["default_filament_colour"] = [sim_hex]
                            cfg["filament_notes"] = [f"U1 FullSpectrum | T{slot_i+1} | TD={s['td'].get()}"]
                            fname_new = f"Metadata/filament_settings_{new_e}.config"
                            zout.writestr(fname_new,
                                          json.dumps(cfg, indent=4, ensure_ascii=False).encode("utf-8"))

                shutil.move(tmp_path, save_path)
                messagebox.showinfo(self.t("dlg_saved"), self.t("3mf_write_ok", path=save_path) + backup_info)
                win.destroy()
            except Exception as e:
                messagebox.showerror(self.t("dlg_error"), str(e))

        ctk.CTkButton(win, text=self.t("remap_btn_write"), fg_color="#0f766e",
                      height=40, font=("Segoe UI",12,"bold"),
                      command=do_write).pack(pady=(8,4), padx=20, fill="x")
        ctk.CTkButton(win, text=self.t("exp_cancel"), fg_color="#334155",
                      command=win.destroy, height=34).pack(padx=20, fill="x", pady=(0,10))

    # ── WERKZEUGWECHSEL-SCHÄTZUNG ─────────────────────────────────────────────

    def open_tc_estimator(self):
        """Schätzt Werkzeugwechsel, Zusatzzeit und Purge-Material inkl. Kosten."""
        if not self.virtual_fils:
            messagebox.showinfo(self.t("dlg_note"), self.t("orca_no_virtual")); return
        win = ctk.CTkToplevel(self)
        win.title(self.t("tc_title"))
        win.geometry("460x440")
        win.grab_set()
        ctk.CTkLabel(win, text=self.t("tc_title"),
                     font=("Segoe UI", 14, "bold")).pack(pady=(16, 10))

        # Eingaben
        in_frame = ctk.CTkFrame(win, fg_color="transparent"); in_frame.pack(padx=30, fill="x")
        ctk.CTkLabel(in_frame, text=self.t("tc_layers"), width=220).grid(row=0, column=0, sticky="w", pady=4)
        layers_entry = ctk.CTkEntry(in_frame, width=100, placeholder_text="200")
        layers_entry.insert(0, "200"); layers_entry.grid(row=0, column=1, padx=6)

        ctk.CTkLabel(in_frame, text="Sekunden/Werkzeugwechsel:", width=220).grid(row=1, column=0, sticky="w", pady=4)
        secs_entry = ctk.CTkEntry(in_frame, width=100, placeholder_text="30")
        secs_entry.insert(0, "30"); secs_entry.grid(row=1, column=1, padx=6)

        ctk.CTkLabel(in_frame, text="Purge-Menge (mm³/Wechsel):", width=220).grid(row=2, column=0, sticky="w", pady=4)
        purge_entry = ctk.CTkEntry(in_frame, width=100, placeholder_text="120")
        purge_entry.insert(0, "120"); purge_entry.grid(row=2, column=1, padx=6)

        # Cost inputs (Change 11)
        ctk.CTkLabel(in_frame, text=self.t("tc_cost_per_kg"), width=220).grid(row=3, column=0, sticky="w", pady=4)
        cost_entry = ctk.CTkEntry(in_frame, width=100, placeholder_text="25")
        cost_entry.insert(0, "25"); cost_entry.grid(row=3, column=1, padx=6)

        ctk.CTkLabel(in_frame, text=self.t("tc_density"), width=220).grid(row=4, column=0, sticky="w", pady=4)
        density_entry = ctk.CTkEntry(in_frame, width=100, placeholder_text="1.24")
        density_entry.insert(0, "1.24"); density_entry.grid(row=4, column=1, padx=6)

        result_frame = ctk.CTkFrame(win, fg_color="#1e293b", corner_radius=8)
        result_frame.pack(fill="x", padx=30, pady=12)
        res_lbl = ctk.CTkLabel(result_frame, text="", font=("Segoe UI", 11),
                                text_color="#4ade80", wraplength=380, justify="left")
        res_lbl.pack(pady=12, padx=14)

        def calculate():
            try:
                n_layers  = int(layers_entry.get())
                secs      = float(secs_entry.get())
                purge_vol = float(purge_entry.get())
                cost_per_kg = float(cost_entry.get())
                density = float(density_entry.get())
            except ValueError: return

            lh = safe_td(self.layer_height_entry.get()) if hasattr(self, "layer_height_entry") else 0.08
            total_changes = 0
            for vf in self.virtual_fils:
                seq = vf["sequence"]
                changes_per_cycle = sum(1 for i in range(len(seq)-1) if seq[i] != seq[i+1])
                cycle_len = len(seq)
                cycles = n_layers / max(cycle_len, 1)
                total_changes += int(cycles * changes_per_cycle)

            extra_min = total_changes * secs / 60
            # Purge in Gramm: mm³ × density(g/cm³) / 1000 (cm³)
            purge_cm3 = total_changes * purge_vol / 1000
            purge_g = purge_cm3 * density
            # Cost = volume_cm3 × density(g/cm3) × cost_per_kg/1000
            purge_cost = purge_cm3 * density * (cost_per_kg / 1000)
            # Equivalent print layers
            layer_vol_mm3 = purge_vol  # rough approximation: each layer ~ same as purge vol
            purge_layer_equiv = int(total_changes * purge_vol / max(layer_vol_mm3, 1)) if layer_vol_mm3 > 0 else 0

            res_lbl.configure(text=(
                f"{self.t('tc_result', n=total_changes)}\n"
                f"{self.t('tc_time', min=extra_min, sec=int(secs))}\n"
                f"{self.t('tc_purge', g=purge_g)}\n"
                f"{self.t('tc_purge_cost', cost=purge_cost)}\n"
                f"{self.t('tc_purge_layers', n=total_changes)}"
            ))

        ctk.CTkButton(win, text="Berechnen", fg_color="#2563eb",
                      command=calculate, height=40).pack(pady=(0,6), padx=30, fill="x")
        ctk.CTkButton(win, text=self.t("exp_cancel"), fg_color="#334155",
                      command=win.destroy, height=34).pack(padx=30, fill="x")

    # ── ALLE CADENCE-WERTE KOPIEREN ────────────────────────────────────────────

    def open_copy_all_cadence(self):
        if not self.virtual_fils:
            messagebox.showinfo(self.t("dlg_note"), self.t("orca_no_virtual")); return
        lh = safe_td(self.layer_height_entry.get()) if hasattr(self,"layer_height_entry") else 0.08
        lines = []
        for vf in self.virtual_fils:
            seq = vf["sequence"]
            n_f = seq_filament_count(seq)
            raw = "".join(seq)
            cad = calc_cadence(seq, lh)
            ids = sorted(cad.keys())
            label = vf.get("label", f"V{vf['vid']}")
            if n_f == 1:
                hint = f"Reine Farbe T{seq[0]}"
            elif n_f == 2:
                a = round(cad.get(ids[0], lh), 3)
                b = round(cad.get(ids[1], lh) if len(ids) > 1 else lh, 3)
                hint = f"Cadence A={a}mm  B={b}mm  |  Step={lh}mm"
            else:
                hint = f"Pattern: {raw}  |  Step={lh}mm"
            de = vf.get("de", 0.0)
            lines.append(f"V{vf['vid']} [{label}]  Seq:{raw}  ΔE:{de:.1f}  →  {hint}")

        win = ctk.CTkToplevel(self)
        win.title(self.t("copy_all_title"))
        win.geometry("720x380")
        win.grab_set()
        ctk.CTkLabel(win, text=self.t("copy_all_title"),
                     font=("Segoe UI",13,"bold")).pack(pady=(14,6))
        import tkinter as _tk2
        txt = _tk2.Text(win, bg="#0f172a", fg="#e2e8f0", font=("Courier New",10),
                        relief="flat", bd=0, wrap="none", padx=10, pady=6)
        txt.pack(fill="both", expand=True, padx=16, pady=4)
        full = "\n".join(lines)
        txt.insert("1.0", full)
        txt.configure(state="disabled")
        def _copy_all():
            self.clipboard_clear(); self.clipboard_append(full)
        ctk.CTkButton(win, text=self.t("copy_all_btn"), fg_color="#0f766e",
                      command=_copy_all, height=36).pack(pady=8, padx=20, fill="x")

    # ── PALETTEN-IMPORT ────────────────────────────────────────────────────────

    def import_palette_from_image(self):
        """Dominante Farben aus Bild extrahieren und als virtuelle Köpfe hinzufügen."""
        if not _HAS_PIL:
            messagebox.showinfo(self.t("dlg_note"), "Pillow not installed."); return
        path = filedialog.askopenfilename(
            filetypes=[(self.t("img_filetypes"), "*.png *.jpg *.jpeg *.bmp *.webp"),
                       ("*", "*.*")],
            title=self.t("palette_title"))
        if not path: return

        try:
            img = _PILImage.open(path).convert("RGB")
            img.thumbnail((200, 200))
        except Exception as e:
            messagebox.showerror(self.t("dlg_error"), str(e)); return

        win = ctk.CTkToplevel(self)
        win.title(self.t("palette_title"))
        win.geometry("480x380")
        win.grab_set()
        ctk.CTkLabel(win, text=self.t("palette_title"),
                     font=("Segoe UI", 14, "bold")).pack(pady=(14, 6))

        # Simple Quantisierung: Farbraum in 6-Bucket-Raster aufteilen
        from collections import Counter
        pixels = list(img.getdata())
        def quantize(px, levels=6):
            step = 256 // levels
            return tuple((c // step) * step + step//2 for c in px[:3])
        counts = Counter(quantize(p) for p in pixels)
        # Top-Farben nach Häufigkeit
        top = [rgb for rgb, _ in counts.most_common(50)]
        # Nur Farben mit Mindestabstand auswählen
        selected = []
        for rgb in top:
            hx = "#{:02X}{:02X}{:02X}".format(*rgb)
            lab = rgb_to_lab(rgb)
            if all(delta_e(lab, rgb_to_lab(hex_to_rgb(s))) > 12 for s in selected):
                selected.append(hx)
            if len(selected) >= 8:
                break

        ctk.CTkLabel(win, text=self.t("palette_colors"),
                     font=("Segoe UI", 10), text_color="#64748b").pack(pady=(0, 4))

        check_vars = []
        swatch_frame = ctk.CTkScrollableFrame(win, height=180)
        swatch_frame.pack(fill="x", padx=20, pady=4)
        for hx in selected:
            row = ctk.CTkFrame(swatch_frame, fg_color="#1e293b", corner_radius=6)
            row.pack(fill="x", pady=2)
            var = ctk.BooleanVar(value=True)
            check_vars.append((var, hx))
            ctk.CTkCheckBox(row, text="", variable=var, width=30).pack(side="left", padx=4)
            ctk.CTkLabel(row, text="", width=36, height=28,
                          fg_color=hx, corner_radius=6).pack(side="left", padx=4, pady=4)
            ctk.CTkLabel(row, text=hx, font=("Courier New", 10)).pack(side="left", padx=4)

        def add_selected():
            selected_colors = [hx for var, hx in check_vars if var.get()]
            if not selected_colors: return
            self.virtual_undo.append(copy.deepcopy(self.virtual_fils))
            added = 0
            for hx in selected_colors:
                if len(self.virtual_fils) >= self._max_virtual: break
                res = self._calc_for_color(hx, False, seq_len=None,
                                            auto=True, auto_threshold=3.0)
                if res is None: continue
                vid = 5 + len(self.virtual_fils)
                self.virtual_fils.append({
                    "vid": vid, "target_hex": hx,
                    "sequence": res["sequence"], "sim_hex": res["sim_hex"],
                    "de": res["de"], "label": f"Pal {added+1}",
                })
                added += 1
            self._refresh_virtual_grid()
            messagebox.showinfo(self.t("dlg_saved"),
                self.t("batch_done", n=added))
            win.destroy()

        ctk.CTkButton(win, text=self.t("palette_btn_add"), fg_color="#0e7490",
                      command=add_selected, height=40).pack(pady=(8,4), padx=20, fill="x")
        ctk.CTkButton(win, text=self.t("exp_cancel"), fg_color="#334155",
                      command=win.destroy, height=34).pack(padx=20, fill="x")

    # ── FARBHARMONIEN ─────────────────────────────────────────────────────────

    def open_harmonies_dialog(self):
        """Zeigt Komplementär, Triad, Analog und Split-Komplementär zur Zielfarbe."""
        if not hasattr(self, "target"):
            messagebox.showinfo(self.t("dlg_note"), self.t("dlg_select_color")); return
        win = ctk.CTkToplevel(self)
        win.title(self.t("harm_title"))
        win.geometry("520x360")
        win.grab_set()
        ctk.CTkLabel(win, text=self.t("harm_title"),
                     font=("Segoe UI", 14, "bold")).pack(pady=(14, 8))

        r, g, b = hex_to_rgb(self.target)
        h, s, v = self._rgb_to_hsv(r, g, b)

        def hsv_to_hex(hue, sat, val):
            hue = hue % 360
            h_i = int(hue / 60) % 6
            f   = hue / 60 - int(hue / 60)
            p   = val * (1 - sat / 100)
            q   = val * (1 - f * sat / 100)
            t   = val * (1 - (1 - f) * sat / 100)
            v2  = val / 100
            p2, q2, t2 = p/100, q/100, t/100
            rgb_map = [(v2,t2,p2),(q2,v2,p2),(p2,v2,t2),(p2,q2,v2),(t2,p2,v2),(v2,p2,q2)]
            rr, gg, bb = rgb_map[h_i]
            return "#{:02X}{:02X}{:02X}".format(int(rr*255), int(gg*255), int(bb*255))

        harmonies = {
            self.t("harm_complement"):  [hsv_to_hex(h+180, s, v)],
            self.t("harm_triadic"):     [hsv_to_hex(h+120, s, v), hsv_to_hex(h+240, s, v)],
            self.t("harm_analogous"):   [hsv_to_hex(h-30, s, v), hsv_to_hex(h+30, s, v)],
            self.t("harm_split"):       [hsv_to_hex(h+150, s, v), hsv_to_hex(h+210, s, v)],
        }

        all_colors = []
        for name, colors in harmonies.items():
            row = ctk.CTkFrame(win, fg_color="#1e293b", corner_radius=6)
            row.pack(fill="x", padx=20, pady=3)
            ctk.CTkLabel(row, text=name, width=160, font=("Segoe UI", 10),
                         text_color="#94a3b8").pack(side="left", padx=10)
            for hex_c in colors:
                all_colors.append(hex_c)
                swatch = ctk.CTkLabel(row, text="", width=36, height=36,
                                       fg_color=hex_c, corner_radius=6)
                swatch.pack(side="left", padx=4, pady=6)
                ctk.CTkLabel(row, text=hex_c, font=("Courier New", 9),
                              text_color="#64748b").pack(side="left", padx=2)
                def use_this(hc=hex_c):
                    self._apply_target(hc)
                    win.destroy()
                ctk.CTkButton(row, text="→ Ziel", width=60, height=26,
                              fg_color="#374151", command=use_this).pack(side="left", padx=4)

        def add_all():
            self.virtual_undo.append(copy.deepcopy(self.virtual_fils))
            added = 0
            for hex_c in all_colors:
                if len(self.virtual_fils) >= self._max_virtual: break
                res = self._calc_for_color(hex_c, False, seq_len=None,
                                            auto=True, auto_threshold=3.0)
                if res is None: continue
                vid = 5 + len(self.virtual_fils)
                self.virtual_fils.append({
                    "vid": vid, "target_hex": hex_c,
                    "sequence": res["sequence"], "sim_hex": res["sim_hex"],
                    "de": res["de"], "label": f"Harm {added+1}",
                })
                added += 1
            self._refresh_virtual_grid()
            win.destroy()

        ctk.CTkButton(win, text=self.t("harm_add_all"), fg_color="#0e7490",
                      command=add_all, height=38).pack(pady=(8, 4), padx=20, fill="x")
        ctk.CTkButton(win, text=self.t("exp_cancel"), fg_color="#334155",
                      command=win.destroy, height=34).pack(padx=20, fill="x")

    # ── GAMUT-PLOT ─────────────────────────────────────────────────────────────

    def open_gamut_plot(self):
        if not _HAS_MPL:
            messagebox.showinfo(self.t("dlg_note"),
                "matplotlib nicht installiert.\npip install matplotlib"); return
        fils = self._get_fils()
        if not fils:
            messagebox.showinfo(self.t("dlg_note"), "Keine Filamente geladen"); return

        fig, ax = _plt.subplots(figsize=(7, 7))
        ax.set_facecolor("#0f172a"); fig.patch.set_facecolor("#1e293b")
        ax.set_xlabel("a* (Grün ← → Rot)", color="#94a3b8")
        ax.set_ylabel("b* (Blau ← → Gelb)", color="#94a3b8")
        ax.set_title("Erreichbarer Gamut — CIE a*b*", color="#e2e8f0", fontsize=13, fontweight="bold")
        ax.tick_params(colors="#64748b")
        for spine in ax.spines.values(): spine.set_color("#334155")
        ax.axhline(0, color="#334155", lw=0.5)
        ax.axvline(0, color="#334155", lw=0.5)

        # Gamut-Wolke: alle Kombinationen 2–5 Schichten
        import itertools
        cloud_a, cloud_b = [], []
        fil_ids = [str(f["id"]) for f in fils]
        for length in range(2, 6):
            for combo in itertools.product(fil_ids, repeat=length):
                if len(set(combo)) < 2: continue
                try:
                    lab = self._simulate_mix(list(combo), fils)
                    cloud_a.append(lab[1]); cloud_b.append(lab[2])
                except Exception: pass
        if cloud_a:
            ax.scatter(cloud_a, cloud_b, s=6, alpha=0.12, color="#38bdf8", zorder=1)
            try:
                from scipy.spatial import ConvexHull
                import numpy as np
                pts = np.array(list(zip(cloud_a, cloud_b)))
                hull = ConvexHull(pts)
                verts = pts[hull.vertices]
                verts = list(verts) + [verts[0]]
                ax.fill([p[0] for p in verts], [p[1] for p in verts],
                        alpha=0.06, color="#38bdf8", zorder=2)
                ax.plot([p[0] for p in verts], [p[1] for p in verts],
                        color="#38bdf8", lw=1, alpha=0.35, zorder=2)
            except Exception: pass

        # Physische Filamente
        for f in fils:
            lab = f["lab"]
            ax.scatter(lab[1], lab[2], s=200, color=f["hex"],
                       edgecolors="#ffffff", lw=1.5, zorder=5)
            ax.annotate(f"T{f['id']}", (lab[1], lab[2]),
                        textcoords="offset points", xytext=(7, 4),
                        color="#e2e8f0", fontsize=10, fontweight="bold")

        # Virtuelle Köpfe
        for vf in self.virtual_fils:
            sh = vf.get("sim_hex", "#888888")
            try:
                lab = rgb_to_lab(hex_to_rgb(sh))
                ax.scatter(lab[1], lab[2], s=90, marker="D", color=sh,
                           edgecolors="#fbbf24", lw=1.2, zorder=5)
                ax.annotate(f"V{vf['vid']}", (lab[1], lab[2]),
                            textcoords="offset points", xytext=(5, 3),
                            color="#fbbf24", fontsize=8)
            except Exception: pass

        # Zielfarbe
        if hasattr(self, "target") and self.target:
            try:
                t_lab = rgb_to_lab(hex_to_rgb(self.target))
                ax.scatter(t_lab[1], t_lab[2], s=250, marker="*",
                           color=self.target, edgecolors="#f87171", lw=2, zorder=6)
                ax.annotate("Ziel", (t_lab[1], t_lab[2]),
                            textcoords="offset points", xytext=(7, 4),
                            color="#f87171", fontsize=10)
            except Exception: pass

        _plt.tight_layout()
        _plt.show()

    # ── TD-KALIBRIERUNG ────────────────────────────────────────────────────────

    def open_td_calibration(self):
        """TD-Werte anhand eines Testdrucks kalibrieren."""
        win = ctk.CTkToplevel(self)
        win.title(self.t("td_cal_title"))
        win.geometry("500x420")
        win.grab_set()

        ctk.CTkLabel(win, text=self.t("td_cal_title"),
                     font=("Segoe UI", 14, "bold")).pack(pady=(16, 6))

        # Slot-Auswahl A und B
        sel_frame = ctk.CTkFrame(win, fg_color="transparent"); sel_frame.pack(pady=6)
        ctk.CTkLabel(sel_frame, text=self.t("td_cal_slot")).pack(side="left", padx=6)
        slot_a = ctk.IntVar(value=1)
        slot_b = ctk.IntVar(value=2)
        ctk.CTkLabel(sel_frame, text="A:").pack(side="left", padx=(8, 2))
        ctk.CTkOptionMenu(sel_frame, variable=slot_a, values=["1","2","3","4"],
                          width=60, command=lambda x: update_desc()).pack(side="left")
        ctk.CTkLabel(sel_frame, text="  B:").pack(side="left", padx=(8, 2))
        ctk.CTkOptionMenu(sel_frame, variable=slot_b, values=["1","2","3","4"],
                          width=60, command=lambda x: update_desc()).pack(side="left")

        desc_lbl = ctk.CTkLabel(win, text="", font=("Segoe UI", 10),
                                 text_color="#64748b", wraplength=440)
        desc_lbl.pack(pady=4)

        def update_desc(*a):
            desc_lbl.configure(text=self.t("td_cal_desc", a=slot_a.get(), b=slot_b.get()))
        update_desc()

        # Vorschau: aktuelle Slot-Farben
        pf = ctk.CTkFrame(win, fg_color="#1e293b", corner_radius=8); pf.pack(pady=8, padx=30, fill="x")
        pf.grid_columnconfigure((0,1,2,3,4), weight=1)
        for col, (txt, key) in enumerate([("T{a} Farbe","a"),("",""),("50/50 Mix",""),("",""),("T{b} Farbe","b")]):
            if col in (1,3): continue
            label = txt.format(a=slot_a.get(), b=slot_b.get()) if "{" in txt else txt
            ctk.CTkLabel(pf, text=label, font=("Segoe UI", 9), text_color="#64748b").grid(
                row=0, column=col, pady=(6,2))
        hex_a_lbl = ctk.CTkLabel(pf, text="", width=50, height=50,
                                  fg_color=self.slots[0]["hex"].get() or "#888888",
                                  corner_radius=8)
        hex_a_lbl.grid(row=1, column=0, padx=10, pady=(0,8))
        ctk.CTkLabel(pf, text="⟷", font=("Segoe UI", 14), text_color="#64748b").grid(row=1, column=1)
        self._td_mix_preview = ctk.CTkLabel(pf, text="", width=50, height=50,
                                             fg_color="#888888", corner_radius=8)
        self._td_mix_preview.grid(row=1, column=2, padx=10, pady=(0,8))
        ctk.CTkLabel(pf, text="⟷", font=("Segoe UI", 14), text_color="#64748b").grid(row=1, column=3)
        hex_b_lbl = ctk.CTkLabel(pf, text="", width=50, height=50,
                                  fg_color=self.slots[1]["hex"].get() or "#888888",
                                  corner_radius=8)
        hex_b_lbl.grid(row=1, column=4, padx=10, pady=(0,8))

        # Gemessene Farbe
        ctk.CTkLabel(win, text=self.t("td_cal_measured"),
                     font=("Segoe UI", 10)).pack(pady=(8, 2))
        mf = ctk.CTkFrame(win, fg_color="transparent"); mf.pack()
        measured_var = ctk.StringVar(value="#888888")
        meas_swatch = ctk.CTkLabel(mf, text="", width=32, height=32,
                                    fg_color="#888888", corner_radius=6)
        meas_swatch.pack(side="left", padx=6)
        meas_entry = ctk.CTkEntry(mf, textvariable=measured_var, width=100)
        meas_entry.pack(side="left", padx=4)
        def on_measured(*a):
            h = measured_var.get().strip()
            if len(h) == 7 and h.startswith("#"):
                try: meas_swatch.configure(fg_color=h)
                except: pass
        measured_var.trace_add("write", on_measured)
        def pick_measured():
            from tkinter import colorchooser as _cc
            col = _cc.askcolor(title="Gemessene Farbe")[1]
            if col: measured_var.set(col.upper())
        ctk.CTkButton(mf, text="🎨", width=32, height=32, fg_color="#374151",
                      command=pick_measured).pack(side="left", padx=2)

        result_lbl = ctk.CTkLabel(win, text="", font=("Segoe UI", 11, "bold"),
                                   text_color="#4ade80")
        result_lbl.pack(pady=8)

        def calc_td():
            a_idx = int(slot_a.get()) - 1
            b_idx = int(slot_b.get()) - 1
            hex_a = self.slots[a_idx]["hex"].get().strip()
            hex_b = self.slots[b_idx]["hex"].get().strip()
            meas  = measured_var.get().strip()
            if not all(len(h) == 7 for h in [hex_a, hex_b, meas]):
                return
            lab_a    = rgb_to_lab(hex_to_rgb(hex_a))
            lab_b    = rgb_to_lab(hex_to_rgb(hex_b))
            lab_meas = rgb_to_lab(hex_to_rgb(meas))
            # Bestimme Mischgewicht aus gemessener Farbe (Projektion auf A-B-Linie)
            de_total = delta_e(lab_a, lab_b)
            if de_total < 1e-3: return
            de_a = delta_e(lab_meas, lab_a)
            de_b = delta_e(lab_meas, lab_b)
            w_b = de_a / (de_a + de_b + 1e-9)  # Gewicht B
            # Empfohlenes TD-Verhältnis: w_b/w_a = td_a/td_b (FullSpectrum-Formel)
            td_old_a = safe_td(self.slots[a_idx]["td"].get())
            td_old_b = safe_td(self.slots[b_idx]["td"].get())
            # Skaliere so dass Summe erhalten bleibt
            ratio = w_b / max(1 - w_b, 0.01)
            td_new_a = round(td_old_a * 2 / (1 + ratio), 1)
            td_new_b = round(td_old_b * 2 * ratio / (1 + ratio), 1)
            td_new_a = max(0.1, min(10.0, td_new_a))
            td_new_b = max(0.1, min(10.0, td_new_b))
            # Vorschau der vorhergesagten Mischfarbe
            mix_lab = tuple((lab_a[k]*(1-w_b) + lab_b[k]*w_b) for k in range(3))
            self._td_mix_preview.configure(fg_color=lab_to_hex(mix_lab))
            result_lbl.configure(
                text=self.t("td_cal_result", a=slot_a.get(), b=slot_b.get(),
                             ta=td_new_a, tb=td_new_b))
            # Apply-Button aktivieren
            def apply():
                self.slots[a_idx]["td"].delete(0, "end")
                self.slots[a_idx]["td"].insert(0, str(td_new_a))
                self.slots[b_idx]["td"].delete(0, "end")
                self.slots[b_idx]["td"].insert(0, str(td_new_b))
                win.destroy()
            apply_btn.configure(command=apply)

        ctk.CTkButton(win, text="Berechnen", fg_color="#2563eb",
                      command=calc_td, height=36).pack(pady=(4, 2), padx=30, fill="x")
        apply_btn = ctk.CTkButton(win, text=self.t("td_cal_apply"), fg_color="#15803d",
                                   height=36)
        apply_btn.pack(pady=(2, 8), padx=30, fill="x")

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
        # Pre-suggest TD from hex brightness (Change 8 / 13)
        est_td = estimate_td(h)
        td_r = ctk.CTkInputDialog(
            text=self.t("inp_td2", td=DEFAULT_TD) + f"\n(Schätzung aus Helligkeit: ~{est_td})",
            title=self.t("inp_td_title")).get_input()
        self.library.setdefault(brand, []).append(
            {"name": n.strip(), "hex": h, "td": safe_td(td_r) if td_r else est_td})
        self.save_db(); self._refresh_brand_menus(); cb()


    # ── FEATURE 3: HEX-CLICK IN V-GRID (added in _build_virtual_row) ──────────

    # ── FEATURE 4: V-HEAD REORDER ─────────────────────────────────────────────

    def _vhead_move(self, idx, direction):
        vf = self.virtual_fils
        new_idx = idx + direction
        if 0 <= new_idx < len(vf):
            vf[idx], vf[new_idx] = vf[new_idx], vf[idx]
            for i, v in enumerate(vf):
                v["vid"] = 5 + i
            self._refresh_virtual_grid()

    # ── FEATURE 5: IMPROVED HISTORY ───────────────────────────────────────────

    def _update_history(self, target_hex, sim_hex, de, seq):
        """Add entry to history (max 10) and refresh history UI if available."""
        entry = {"target_hex": target_hex, "sim_hex": sim_hex,
                 "de": de, "sequence": seq}
        # Remove duplicate target
        self.history = [h for h in self.history if h["target_hex"] != target_hex]
        self.history.insert(0, entry)
        if len(self.history) > 10:
            self.history = self.history[:10]
        self._refresh_history_ui()

    def _refresh_history_ui(self):
        """Refresh history cards panel if it exists."""
        if not hasattr(self, "_history_frame"):
            return
        frame = self._history_frame
        for w in frame.winfo_children():
            w.destroy()
        for entry in self.history:
            card = ctk.CTkFrame(frame, fg_color="#1e293b", corner_radius=6)
            card.pack(fill="x", pady=2, padx=2)
            try:
                swatch = ctk.CTkLabel(card, text="  ", width=20, height=20,
                                      fg_color=entry["target_hex"], corner_radius=3)
                swatch.pack(side="left", padx=(4, 2), pady=3)
            except Exception:
                pass
            seq_str = entry.get("sequence", "")
            de_val = entry.get("de", 0)
            lbl_text = f"{entry['target_hex']}  ΔE={de_val:.1f}  [{seq_str}]"
            lbl = ctk.CTkLabel(card, text=lbl_text, font=("Segoe UI", 9),
                               cursor="hand2", anchor="w")
            lbl.pack(side="left", padx=4, fill="x", expand=True)
            def _load_hist(e=entry):
                self._apply_target(e["target_hex"])
            lbl.bind("<Button-1>", lambda ev, e=entry: self._apply_target(e["target_hex"]))

    # ── RECENT COLORS (Change 7) ──────────────────────────────────────────────

    def _add_recent_color(self, hex_color):
        """Add color to recent colors list, deduplicate, keep max 10."""
        if not hex_color:
            return
        hex_color = hex_color.upper()
        self._recent_colors = [c for c in self._recent_colors if c != hex_color]
        self._recent_colors.insert(0, hex_color)
        self._recent_colors = self._recent_colors[:10]
        self._update_recent_swatches()

    def _update_recent_swatches(self):
        """Clear and redraw recent color swatches."""
        if not hasattr(self, '_recent_frame'):
            return
        for w in self._recent_frame.winfo_children():
            w.destroy()
        for hex_c in self._recent_colors[:10]:
            try:
                swatch = ctk.CTkLabel(self._recent_frame, text="", width=18, height=18,
                                      fg_color=hex_c, corner_radius=3, cursor="hand2")
                swatch.pack(side="left", padx=2, pady=2)
                def _on_click(e, h=hex_c):
                    self._apply_target(h)
                    self.calc()
                swatch.bind("<Button-1>", _on_click)
            except Exception:
                pass

    # ── FEATURE 7: FILAMENT SEARCH DIALOG ────────────────────────────────────

    def open_filament_search(self, slot_idx):
        # Singleton per slot — focus existing window instead of opening a second one
        if not hasattr(self, "_search_wins"):
            self._search_wins = {}
        existing = self._search_wins.get(slot_idx)
        if existing and existing.winfo_exists():
            existing.focus_force()
            existing.lift()
            return

        win = ctk.CTkToplevel(self)
        win.title("Filament suchen")
        win.geometry("420x560")
        self._search_wins[slot_idx] = win
        win.protocol("WM_DELETE_WINDOW", lambda: (self._search_wins.pop(slot_idx, None), win.destroy()))

        # Brand filter row
        brand_row = ctk.CTkFrame(win, fg_color="transparent")
        brand_row.pack(fill="x", padx=8, pady=(8, 2))
        brand_var = ctk.StringVar(value="Alle")
        all_brands = ["Alle"] + list(self.library.keys())
        brand_menu = ctk.CTkOptionMenu(brand_row, variable=brand_var, values=all_brands,
                                        width=180)
        brand_menu.pack(side="left")

        search_var = ctk.StringVar()
        entry = ctk.CTkEntry(win, textvariable=search_var, placeholder_text="Suchen...")
        entry.pack(fill="x", padx=8, pady=(2, 4))
        entry.focus()

        # Use tk.Listbox for keyboard navigation and double-click support
        list_frame = ctk.CTkFrame(win, fg_color="#0f172a", corner_radius=6)
        list_frame.pack(fill="both", expand=True, padx=8, pady=4)
        listbox = tk.Listbox(list_frame, bg="#0f172a", fg="#e2e8f0",
                             selectbackground="#2563eb", selectforeground="white",
                             font=("Segoe UI", 11), borderwidth=0, highlightthickness=0,
                             activestyle="none")
        listbox_scroll = tk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=listbox_scroll.set)
        listbox_scroll.pack(side="right", fill="y")
        listbox.pack(side="left", fill="both", expand=True, padx=4, pady=4)

        # Store search results for apply
        _results = []

        def _apply_selected():
            sel = listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            if idx >= len(_results):
                return
            brand, fil = _results[idx]
            self._save_slot_snapshot()
            s = self.slots[slot_idx]
            s["brand"].set(brand)
            self.update_menu(slot_idx)
            try:
                s["color"].set(fil.get("name", ""))
            except Exception:
                pass
            s["hex"].delete(0, "end")
            s["hex"].insert(0, fil.get("hex", "#FFFFFF"))
            s["td"].delete(0, "end")
            s["td"].insert(0, str(fil.get("td", 5.0)))
            try:
                s["preview"].configure(fg_color=fil.get("hex", "#FFFFFF"))
            except Exception:
                pass
            self._search_wins.pop(slot_idx, None)
            win.destroy()

        _debounce_id = [None]

        def _schedule_update(*args):
            if _debounce_id[0]:
                win.after_cancel(_debounce_id[0])
            _debounce_id[0] = win.after(200, _update_results)

        def _update_results():
            listbox.delete(0, "end")
            _results.clear()
            q = search_var.get().lower()
            brand_filter = brand_var.get()
            count = 0
            for brand, filaments in self.library.items():
                if brand_filter != "Alle" and brand != brand_filter:
                    continue
                for fil in filaments:
                    if q in brand.lower() or q in fil.get("name", "").lower() or q in fil.get("hex", "").lower():
                        if count >= 200:
                            break
                        count += 1
                        _results.append((brand, fil))
                        name_str = f"{brand} — {fil.get('name', '?')}  {fil.get('hex', '')}"
                        listbox.insert("end", name_str)

        listbox.bind("<Double-Button-1>", lambda e: _apply_selected())
        listbox.bind("<Return>", lambda e: _apply_selected())
        brand_var.trace_add("write", _schedule_update)
        search_var.trace_add("write", _schedule_update)

        # OK button at bottom
        ok_btn = ctk.CTkButton(win, text="✓ Übernehmen", fg_color="#15803d",
                                command=_apply_selected)
        ok_btn.pack(fill="x", padx=8, pady=(0, 8))

        _update_results()

    # ── FEATURE 8: FILAMENT DISTANCE MATRIX ──────────────────────────────────

    def open_filament_matrix(self):
        win = ctk.CTkToplevel(self)
        win.title("Filament ΔE-Matrix")
        win.geometry("440x300")

        colors = []
        names = []
        for s in self.slots:
            hex_c = s["hex"].get().strip().lstrip("#")
            try:
                lab = rgb_to_lab(hex_to_rgb(hex_c))
                colors.append(lab)
                names.append(s["color"].get() or f"T{self.slots.index(s)+1}")
            except Exception:
                colors.append((50, 0, 0))
                names.append(f"T{self.slots.index(s)+1}")

        frame = ctk.CTkFrame(win)
        frame.pack(fill="both", expand=True, padx=8, pady=8)

        ctk.CTkLabel(frame, text="", width=90).grid(row=0, column=0, padx=4, pady=4)
        for j, name in enumerate(names):
            ctk.CTkLabel(frame, text=name[:10],
                         font=ctk.CTkFont(weight="bold")).grid(row=0, column=j+1, padx=4, pady=4)

        for i in range(4):
            ctk.CTkLabel(frame, text=names[i][:10],
                         font=ctk.CTkFont(weight="bold")).grid(row=i+1, column=0, padx=4, pady=4)
            for j in range(4):
                if i == j:
                    val = "—"
                    fg = None
                else:
                    de = delta_e(colors[i], colors[j])
                    val = f"{de:.1f}"
                    fg = "#2d8a2d" if de < 10 else ("#d4a000" if de < 30 else "#c0392b")
                lbl = ctk.CTkLabel(frame, text=val, text_color=fg)
                lbl.grid(row=i+1, column=j+1, padx=4, pady=4)

    # ── FEATURE 9: AUTO-SLOT SUGGESTION ──────────────────────────────────────

    def _check_auto_suggestion(self, target_hex):
        target_lab = rgb_to_lab(hex_to_rgb(target_hex.lstrip("#")))
        best_de = float('inf')
        best_fil = None
        best_brand = None
        loaded_hexes = {s["hex"].get().strip().lstrip("#").upper() for s in self.slots}

        for brand, filaments in self.library.items():
            for fil in filaments:
                fhex = fil.get("hex", "").lstrip("#")
                if fhex.upper() in loaded_hexes:
                    continue
                try:
                    flab = rgb_to_lab(hex_to_rgb(fhex))
                    de = delta_e(target_lab, flab)
                    if de < best_de:
                        best_de = de
                        best_fil = fil
                        best_brand = brand
                except Exception:
                    pass

        if best_fil and best_de < 15 and hasattr(self, "suggestion_label"):
            msg = f"💡 Tipp: «{best_brand} {best_fil.get('name','')}» könnte ΔE auf ~{best_de:.1f} senken"
            self.suggestion_label.configure(text=msg, text_color="gray")

    # ── FEATURE 10: MATERIAL COMPATIBILITY WARNING ────────────────────────────

    def _check_material_compatibility(self, sequence_slot_indices):
        used_slots = set(sequence_slot_indices)
        types_used = set()
        for idx in used_slots:
            if idx < 0 or idx >= len(self.slots):
                continue
            slot_name = self.slots[idx]["color"].get()
            slot_brand = self.slots[idx]["brand"].get()
            fil_type = None
            for fil in self.library.get(slot_brand, []):
                if fil.get("name") == slot_name:
                    fil_type = fil.get("type", None)
                    break
            if fil_type is None:
                n = slot_name.upper()
                if "ABS" in n:
                    fil_type = "ABS"
                elif "PETG" in n:
                    fil_type = "PETG"
                elif "TPU" in n:
                    fil_type = "TPU"
                elif "ASA" in n:
                    fil_type = "ASA"
                else:
                    fil_type = None  # unknown, skip
            if fil_type:
                types_used.add(fil_type)

        if len(types_used) < 2:
            return None
        incompatible_pairs = {("PLA", "ABS"), ("ABS", "PLA"),
                               ("PLA", "ASA"), ("ASA", "PLA"),
                               ("ABS", "PETG"), ("PETG", "ABS")}
        for a, b in incompatible_pairs:
            if a in types_used and b in types_used:
                return f"⚠ Materialwarnung: {a} + {b} — unterschiedliche Drucktemperaturen!"
        return None

    # ── FEATURE 11: PNG SUMMARY EXPORT ───────────────────────────────────────

    def export_png_summary(self):
        try:
            from PIL import Image, ImageDraw, ImageFont
        except ImportError:
            messagebox.showerror("Fehler", "Pillow nicht installiert. pip install Pillow")
            return

        if not self.virtual_fils:
            messagebox.showinfo("Info", "Keine virtuellen Köpfe vorhanden.")
            return

        path = filedialog.asksaveasfilename(defaultextension=".png",
            filetypes=[("PNG", "*.png")], title="PNG exportieren")
        if not path:
            return

        cols = 4
        rows = math.ceil(len(self.virtual_fils) / cols)
        cell_w, cell_h = 160, 80
        img_w = cols * cell_w + 20
        img_h = rows * cell_h + 40

        img = Image.new("RGB", (img_w, img_h), (30, 30, 30))
        draw = ImageDraw.Draw(img)

        try:
            font = ImageFont.truetype("arial.ttf", 12)
            small_font = ImageFont.truetype("arial.ttf", 10)
        except Exception:
            font = ImageFont.load_default()
            small_font = font

        for i, vf in enumerate(self.virtual_fils):
            col = i % cols
            row_idx = i // cols
            x = col * cell_w + 10
            y = row_idx * cell_h + 30

            sim_hex = vf.get("sim_hex", "#888888").lstrip("#")
            try:
                r, g, b = int(sim_hex[0:2], 16), int(sim_hex[2:4], 16), int(sim_hex[4:6], 16)
                draw.rectangle([x, y, x + cell_w - 10, y + 50], fill=(r, g, b))
            except Exception:
                draw.rectangle([x, y, x + cell_w - 10, y + 50], fill=(128, 128, 128))

            name = vf.get("name", vf.get("label", f"V{vf.get('vid', i+5)}"))
            de = vf.get("de", 0)
            seq = vf.get("sequence", "")
            draw.text((x + 2, y + 52),
                      f"{name}  ΔE={de:.1f}  [{seq}]",
                      fill=(200, 200, 200), font=small_font)

        draw.text((10, 8), "U1 FullSpectrum — Virtual Heads", fill=(255, 255, 255), font=font)

        img.save(path)
        messagebox.showinfo("Export", f"PNG gespeichert: {path}")

    # ── FEATURE 12: SEQUENCE EDITOR ──────────────────────────────────────────

    def open_sequence_editor(self, vf_idx):
        vf = self.virtual_fils[vf_idx]
        seq = list(vf["sequence"])

        win = ctk.CTkToplevel(self)
        win.title(f"Sequenz bearbeiten — {vf.get('label', vf.get('name', ''))}")
        win.geometry("440x240")

        frame = ctk.CTkFrame(win)
        frame.pack(fill="x", padx=10, pady=10)

        tiles = []
        tile_vars = [ctk.IntVar(value=int(c) - 1) for c in seq]

        def _draw_tile(tile_frame, idx):
            for w in tile_frame.winfo_children():
                w.destroy()
            sv = tile_vars[idx].get()
            sh = self.slots[sv]["hex"].get().lstrip("#") if sv < len(self.slots) else "888888"
            try:
                color = f"#{sh}" if sh else "#888888"
                swatch = ctk.CTkLabel(tile_frame, text=f"T{sv+1}",
                    fg_color=color, width=50, height=50, corner_radius=6,
                    cursor="hand2", font=ctk.CTkFont(weight="bold"))
                swatch.pack(padx=2, pady=2)
                swatch.bind("<Button-1>", lambda e, i=idx: _cycle_tile(i))
            except Exception:
                pass

        def _cycle_tile(idx):
            tile_vars[idx].set((tile_vars[idx].get() + 1) % len(self.slots))
            _draw_tile(tiles[idx], idx)

        def _make_tile(parent, i):
            tf = ctk.CTkFrame(parent, width=54, height=54, corner_radius=8)
            tf.pack(side="left", padx=3)
            tiles.append(tf)
            _draw_tile(tf, i)

        for i in range(len(seq)):
            _make_tile(frame, i)

        def _add_layer():
            if len(tile_vars) < 10:
                tile_vars.append(ctk.IntVar(value=0))
                _make_tile(frame, len(tile_vars) - 1)

        def _remove_layer():
            if len(tile_vars) > 1:
                tile_vars.pop()
                tiles.pop().destroy()

        ctrl = ctk.CTkFrame(win, fg_color="transparent")
        ctrl.pack(fill="x", padx=10, pady=4)
        ctk.CTkButton(ctrl, text="+ Layer", width=80, command=_add_layer).pack(side="left", padx=4)
        ctk.CTkButton(ctrl, text="− Layer", width=80, command=_remove_layer).pack(side="left", padx=4)

        def _save():
            new_seq = "".join(str(tile_vars[i].get() + 1) for i in range(len(tile_vars)))
            vf["sequence"] = new_seq
            colors_lin = []
            for ch in new_seq:
                si = int(ch) - 1
                hx = self.slots[si]["hex"].get().lstrip("#") if si < len(self.slots) else "808080"
                try:
                    r, g, b = (int(hx[0:2], 16) / 255,
                                int(hx[2:4], 16) / 255,
                                int(hx[4:6], 16) / 255)
                    colors_lin.append((r ** 2.2, g ** 2.2, b ** 2.2))
                except Exception:
                    colors_lin.append((0.5, 0.5, 0.5))
            if colors_lin:
                avg = tuple(sum(c[i] for c in colors_lin) / len(colors_lin) for i in range(3))
                avg_srgb = tuple(min(255, int((c ** (1 / 2.2)) * 255)) for c in avg)
                vf["sim_hex"] = "#{:02X}{:02X}{:02X}".format(*avg_srgb)
            self._refresh_virtual_grid()
            win.destroy()

        ctk.CTkButton(win, text="💾 Speichern", command=_save).pack(pady=8)

    # ── FEATURE 13: LIGHT MODE TOGGLE ────────────────────────────────────────

    def _toggle_appearance(self):
        current = ctk.get_appearance_mode()
        new_mode = "Light" if current == "Dark" else "Dark"
        ctk.set_appearance_mode(new_mode)
        icon = "🌙" if new_mode == "Dark" else "☀️"
        if hasattr(self, "appearance_btn"):
            self.appearance_btn.configure(text=icon)


# ── 3MF FARB-WIZARD (CTkToplevel) ─────────────────────────────────────────────

class ThreeMFWizard(ctk.CTkToplevel):
    """3-step wizard: load 3MF → optimize 4 filaments → show result & apply."""

    def __init__(self, app):
        super().__init__(app)
        self._app = app
        self.title(app.t("wizard_title"))
        self.geometry("700x560")
        self.resizable(True, True)
        self.grab_set()

        # State
        self._colors = []         # hex strings from 3MF
        self._best4 = []
        self._avg_de = 0.0
        self._scores = []
        self._thread = None
        self._thread_result = None  # (best4, avg_de, scores) when done
        self._add_virtual_var = ctk.BooleanVar(value=False)

        # Container that holds pages
        self._container = ctk.CTkFrame(self, fg_color="transparent")
        self._container.pack(fill="both", expand=True, padx=0, pady=0)

        self._page1 = None
        self._page2 = None
        self._page3 = None
        self._show_page1()

    # ── PAGE 1 ─────────────────────────────────────────────────────────────────

    def _show_page1(self):
        self._clear_container()
        app = self._app
        t = app.t

        ctk.CTkLabel(self._container, text=t("wizard_step1"),
                     font=("Segoe UI", 16, "bold"), text_color="#38bdf8").pack(pady=(18, 8))

        # Drop zone / load button
        load_frame = ctk.CTkFrame(self._container, fg_color="#1e293b",
                                   corner_radius=12, height=100)
        load_frame.pack(fill="x", padx=24, pady=6)
        load_frame.pack_propagate(False)
        load_btn = ctk.CTkButton(load_frame, text=t("wizard_load_btn"),
                                  fg_color="#0f4c81", hover_color="#1e3a5f",
                                  height=52, font=("Segoe UI", 13, "bold"),
                                  command=self._load_file)
        load_btn.pack(expand=True)

        self._p1_path_lbl = ctk.CTkLabel(self._container, text=t("wizard_no_file"),
                                          font=("Segoe UI", 10), text_color="#64748b")
        self._p1_path_lbl.pack(pady=(4, 2))

        self._p1_count_lbl = ctk.CTkLabel(self._container, text="",
                                           font=("Segoe UI", 12, "bold"), text_color="#4ade80")
        self._p1_count_lbl.pack(pady=(0, 8))

        # Swatch frame
        self._p1_swatch_frame = ctk.CTkScrollableFrame(self._container, height=120,
                                                         fg_color="#0f172a")
        self._p1_swatch_frame.pack(fill="x", padx=24, pady=(0, 8))

        self._p1_next_btn = ctk.CTkButton(self._container, text=t("wizard_next"),
                                           fg_color="#15803d", hover_color="#166534",
                                           height=42, font=("Segoe UI", 13, "bold"),
                                           state="disabled",
                                           command=self._show_page2)
        self._p1_next_btn.pack(pady=(8, 16), padx=24, fill="x")

    def _load_file(self):
        app = self._app
        path = filedialog.askopenfilename(
            filetypes=[("3MF-Dateien" if app.lang == "de" else "3MF Files", "*.3mf"),
                       ("*", "*.*")],
            title=app.t("wizard_load_btn"))
        if not path:
            return
        colors, err = parse_3mf_colors(path)
        if not colors:
            messagebox.showinfo(app.t("wizard_title"),
                                err if err else app.t("wizard_no_file"))
            return
        self._colors = colors
        self._p1_path_lbl.configure(text=os.path.basename(path))
        self._p1_count_lbl.configure(
            text=app.t("wizard_colors_found", n=len(colors)))
        # Draw swatches
        for w in self._p1_swatch_frame.winfo_children():
            w.destroy()
        row_f = ctk.CTkFrame(self._p1_swatch_frame, fg_color="transparent")
        row_f.pack(fill="x")
        for i, hex_c in enumerate(colors[:24]):
            ctk.CTkLabel(row_f, text="", width=26, height=26,
                         fg_color=hex_c, corner_radius=4,
                         tooltip_text=hex_c if False else "").grid(
                row=i // 12, column=i % 12, padx=3, pady=3)
        self._p1_next_btn.configure(state="normal")

    # ── PAGE 2 ─────────────────────────────────────────────────────────────────

    def _show_page2(self):
        self._clear_container()
        app = self._app
        t = app.t

        lib = app._get_all_library_fils()
        self._lib_fils = lib

        ctk.CTkLabel(self._container, text=t("wizard_step2"),
                     font=("Segoe UI", 16, "bold"), text_color="#38bdf8").pack(pady=(18, 4))

        info_txt = t("wizard_info", n_lib=len(lib), n_col=len(self._colors))
        ctk.CTkLabel(self._container, text=info_txt,
                     font=("Segoe UI", 10), text_color="#94a3b8",
                     wraplength=600).pack(pady=(0, 12))

        self._p2_progress = ctk.CTkProgressBar(self._container, width=500)
        self._p2_progress.set(0)
        self._p2_progress.pack(pady=(0, 6))

        self._p2_status_lbl = ctk.CTkLabel(self._container, text="",
                                             font=("Segoe UI", 10), text_color="#64748b")
        self._p2_status_lbl.pack(pady=(0, 16))

        btn_row = ctk.CTkFrame(self._container, fg_color="transparent")
        btn_row.pack(fill="x", padx=24, pady=4)

        self._p2_start_btn = ctk.CTkButton(btn_row, text=t("wizard_start"),
                                            fg_color="#2563eb", hover_color="#1d4ed8",
                                            height=42, font=("Segoe UI", 13, "bold"),
                                            command=self._start_optimization)
        self._p2_start_btn.pack(side="left", expand=True, fill="x", padx=(0, 6))

        self._p2_next_btn = ctk.CTkButton(btn_row, text=t("wizard_next"),
                                           fg_color="#15803d", hover_color="#166534",
                                           height=42, font=("Segoe UI", 13, "bold"),
                                           state="disabled",
                                           command=self._show_page3)
        self._p2_next_btn.pack(side="left", expand=True, fill="x")

    def _start_optimization(self):
        import threading
        self._p2_start_btn.configure(state="disabled")
        self._p2_next_btn.configure(state="disabled")
        self._p2_progress.set(0)
        self._thread_result = None
        target_labs = [rgb_to_lab(hex_to_rgb(h)) for h in self._colors]
        lib = self._lib_fils

        def worker():
            def progress_cb(i, total):
                self._thread_progress = (i, total)
            best4, avg_de, scores = find_best_4_filaments(target_labs, lib, progress_cb)
            self._thread_result = (best4, avg_de, scores)

        self._thread_progress = (0, 1)
        t = threading.Thread(target=worker, daemon=True)
        t.start()
        self._poll_thread(t)

    def _poll_thread(self, t):
        app = self._app
        if t.is_alive():
            i, total = getattr(self, "_thread_progress", (0, 1))
            if total > 0:
                self._p2_progress.set(i / total)
                self._p2_status_lbl.configure(
                    text=app.t("wizard_checking", i=i, total=total))
            self.after(50, lambda: self._poll_thread(t))
        else:
            self._p2_progress.set(1.0)
            if self._thread_result:
                best4, avg_de, scores = self._thread_result
                self._best4 = best4
                self._avg_de = avg_de
                self._scores = scores
                self._p2_status_lbl.configure(
                    text=app.t("wizard_avg_de", de=avg_de))
                self._p2_next_btn.configure(state="normal")
            else:
                self._p2_start_btn.configure(state="normal")

    # ── PAGE 3 ─────────────────────────────────────────────────────────────────

    def _show_page3(self):
        self._clear_container()
        app = self._app
        t = app.t

        ctk.CTkLabel(self._container, text=t("wizard_step3"),
                     font=("Segoe UI", 16, "bold"), text_color="#38bdf8").pack(pady=(14, 6))

        # 4 filament cards
        cards_frame = ctk.CTkFrame(self._container, fg_color="#1e293b", corner_radius=8)
        cards_frame.pack(fill="x", padx=20, pady=(0, 8))
        slot_labels = ["T1", "T2", "T3", "T4"]
        best4 = self._best4 if len(self._best4) >= 4 else (self._best4 + [None] * 4)[:4]
        for i, fil in enumerate(best4):
            card = ctk.CTkFrame(cards_frame, fg_color="#0f172a", corner_radius=6)
            card.grid(row=0, column=i, padx=6, pady=8, sticky="nsew")
            cards_frame.grid_columnconfigure(i, weight=1)
            if fil:
                ctk.CTkLabel(card, text="", width=40, height=40,
                             fg_color=fil.get("hex", "#808080"),
                             corner_radius=6).pack(pady=(8, 4))
                ctk.CTkLabel(card, text=fil.get("brand", ""),
                             font=("Segoe UI", 8), text_color="#94a3b8").pack()
                ctk.CTkLabel(card, text=fil.get("name", ""),
                             font=("Segoe UI", 10, "bold"),
                             text_color="#e2e8f0", wraplength=110).pack(padx=4)
                ctk.CTkLabel(card, text=f"TD={fil.get('td', DEFAULT_TD):.1f}",
                             font=("Segoe UI", 9), text_color="#64748b").pack()
            ctk.CTkLabel(card, text=f"→ {slot_labels[i]}",
                         font=("Segoe UI", 11, "bold"), text_color="#a78bfa").pack(pady=(2, 8))

        # Summary
        de_col = de_color(self._avg_de)
        ctk.CTkLabel(self._container,
                     text=t("wizard_avg_de", de=self._avg_de),
                     font=("Segoe UI", 13, "bold"),
                     text_color=de_col).pack(pady=(2, 6))

        # Coverage table
        ctk.CTkLabel(self._container, text=t("wizard_coverage"),
                     font=("Segoe UI", 11, "bold"), text_color="#94a3b8").pack()
        cov_scroll = ctk.CTkScrollableFrame(self._container, height=140,
                                             fg_color="#0f172a")
        cov_scroll.pack(fill="x", padx=20, pady=(2, 8))
        for i, (hex_c, sc) in enumerate(zip(self._colors, self._scores)):
            row = ctk.CTkFrame(cov_scroll, fg_color="#1e293b", corner_radius=4)
            row.pack(fill="x", padx=2, pady=1)
            ctk.CTkLabel(row, text="", width=20, height=20,
                         fg_color=hex_c, corner_radius=3).pack(side="left", padx=6, pady=4)
            ctk.CTkLabel(row, text=hex_c,
                         font=("Courier New", 10), text_color="#94a3b8").pack(side="left", padx=4)
            sc_lbl = f"ΔE {sc:.1f}"
            ctk.CTkLabel(row, text=sc_lbl,
                         font=("Segoe UI", 10, "bold"),
                         text_color=de_color(sc)).pack(side="right", padx=10)

        # Checkbox: also add virtual heads
        ctk.CTkCheckBox(self._container,
                        text=t("wizard_add_virtual"),
                        variable=self._add_virtual_var,
                        font=("Segoe UI", 10)).pack(pady=(0, 6))

        # Buttons
        btn_row = ctk.CTkFrame(self._container, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(0, 14))
        ctk.CTkButton(btn_row, text=t("wizard_apply"),
                      fg_color="#15803d", hover_color="#166534",
                      height=42, font=("Segoe UI", 12, "bold"),
                      command=self._apply_and_close).pack(side="left", expand=True,
                                                          fill="x", padx=(0, 6))
        ctk.CTkButton(btn_row, text=t("wizard_close"),
                      fg_color="#374151",
                      height=42,
                      command=self.destroy).pack(side="left", expand=True, fill="x")

    def _apply_and_close(self):
        app = self._app
        best4 = self._best4
        # Apply up to 4 filaments to slots T1-T4
        for i, fil in enumerate(best4[:4]):
            if fil is None:
                continue
            brand = fil.get("brand", "")
            name = fil.get("name", "")
            hex_c = fil.get("hex", "#808080")
            td = fil.get("td", DEFAULT_TD)
            s = app.slots[i]
            # Try to set via library lookup
            if brand in app.library:
                match = next((f for f in app.library[brand] if f["name"] == name), None)
                if match:
                    s["brand"].set(brand)
                    app.update_menu(i)
                    s["color"].set(name)
                    app.apply_f(i)
                    continue
            # Fallback: set hex + td directly
            s["hex"].delete(0, "end")
            s["hex"].insert(0, hex_c)
            s["td"].delete(0, "end")
            s["td"].insert(0, str(td))
            s["preview"].configure(fg_color=hex_c)

        # Optionally add virtual heads for all model colors
        if self._add_virtual_var.get():
            for hex_c in self._colors:
                if len(app.virtual_fils) >= app._max_virtual:
                    break
                result = app._calc_for_color(hex_c, auto=True)
                if result:
                    app.add_to_virtual(result)

        messagebox.showinfo(app.t("wizard_title"), app.t("wizard_applied"))
        self.destroy()
        setattr(app, "_3mf_wizard_win", None)

    # ── HELPERS ────────────────────────────────────────────────────────────────

    def _clear_container(self):
        for w in self._container.winfo_children():
            w.destroy()


if __name__ == "__main__":
    import sys
    import argparse
    parser = argparse.ArgumentParser(description="U1 FullSpectrum Helper", add_help=False)
    parser.add_argument("--target", help="Target hex color e.g. #FF5500")
    parser.add_argument("--t1", help="T1 hex color")
    parser.add_argument("--t2", help="T2 hex color")
    parser.add_argument("--t3", help="T3 hex color")
    parser.add_argument("--t4", help="T4 hex color")
    parser.add_argument("--auto", action="store_true", help="Use auto shortest sequence")
    parser.add_argument("--optimizer", action="store_true", help="Use optimizer")
    parser.add_argument("--lh", type=float, default=0.08, help="Layer height mm")
    parser.add_argument("--json", action="store_true", dest="json_out", help="Output as JSON")
    parser.add_argument("--help", "-h", action="store_true", help="Show help")

    # Only parse known args so CTk doesn't choke on unknown args
    args, unknown = parser.parse_known_args()

    if args.help:
        parser.print_help()
        sys.exit(0)

    if args.target:
        # CLI mode - no GUI
        slots = []
        for i, hex_c in enumerate([args.t1, args.t2, args.t3, args.t4]):
            if hex_c:
                _h = hex_c if hex_c.startswith("#") else "#" + hex_c
                slots.append({"id": i+1, "hex": _h, "td": DEFAULT_TD,
                              "lab": rgb_to_lab(hex_to_rgb(_h))})
        if not slots:
            print("Error: specify at least --t1 and --t2"); sys.exit(1)

        target_lab = rgb_to_lab(hex_to_rgb(args.target))
        scores = [{"id": f["id"], "w": (1/(delta_e(target_lab, f["lab"])+0.1))*(10/DEFAULT_TD), "h": f["hex"]} for f in slots]
        tot = sum(s["w"] for s in scores)
        if tot == 0:
            print("Error: could not compute scores"); sys.exit(1)

        def _build_seq_cli(sorted_scores, tot, n):
            counts = [max(0, round((s["w"]/tot)*n)) for s in sorted_scores]
            diff = n - sum(counts)
            i = 0
            while diff > 0: counts[i % len(counts)] += 1; diff -= 1; i += 1
            while diff < 0:
                for j in range(len(counts)):
                    if counts[j] > 0: counts[j] -= 1; diff += 1; break
            arr = [None] * n
            pos = n - 1
            for j, s in enumerate(sorted_scores):
                for _ in range(counts[j]):
                    if pos >= 0: arr[pos] = s["id"]; pos -= 1
            return [x or sorted_scores[0]["id"] for x in arr]

        def _sim_mix_cli(seq, slots_list):
            lin = [0.0, 0.0, 0.0]
            for fid in seq:
                slot = next((s for s in slots_list if s["id"] == fid), None)
                if slot is None: continue
                r, g, b = hex_to_rgb(slot["hex"])
                lin[0] += (r/255)**2.2; lin[1] += (g/255)**2.2; lin[2] += (b/255)**2.2
            n = len(seq)
            lin = [x/n for x in lin]
            return rgb_to_lab(tuple(int((x**(1/2.2))*255) for x in lin))

        result_seq = None
        result_de = float("inf")
        srt = sorted(scores, key=lambda x: x["w"], reverse=True)
        for n in range(1, 11):
            if args.optimizer:
                best_s, best_dv = None, float("inf")
                for perm in iter_permutations(srt):
                    seq = _build_seq_cli(list(perm), tot, n)
                    sim_lab = _sim_mix_cli(seq, slots)
                    dv = delta_e(sim_lab, target_lab)
                    if dv < best_dv: best_dv = dv; best_s = seq
                seq = best_s
                dv = best_dv
            else:
                seq = _build_seq_cli(srt, tot, n)
                sim_lab = _sim_mix_cli(seq, slots)
                dv = delta_e(sim_lab, target_lab)
            if dv < result_de:
                result_de = dv; result_seq = seq
            if args.auto and dv <= 2.0:
                break

        seq_str = "".join(map(str, result_seq)) if result_seq else ""
        cad = calc_cadence(seq_str, args.lh)
        ids = sorted(cad.keys())

        if args.json_out:
            import json as _json
            print(_json.dumps({"sequence": seq_str, "de": round(result_de, 2),
                               "cadence": cad, "layer_height": args.lh}))
        else:
            print(f"Sequence: {seq_str}")
            print(f"ΔE:       {result_de:.2f}")
            if len(ids) == 1:
                print("Mode:     Pure color")
            elif len(ids) == 2:
                print(f"Mode:     Cadence  A={cad[ids[0]]}mm  B={cad[ids[1]]}mm")
            else:
                print(f"Mode:     Pattern  {seq_str}")
    else:
        # GUI mode
        app = U1FullSpectrumApp()
        app.mainloop()
