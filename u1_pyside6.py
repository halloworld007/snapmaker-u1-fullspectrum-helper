"""
U1 FullSpectrum Ultimate — PySide6 Edition
Snapmaker U1 Toolchanger · FullSpectrum layer-dithering color calculator
"""
import sys, os, json, copy, zipfile, re, math, itertools
from itertools import permutations as iter_permutations
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QScrollArea, QTabWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout, QFormLayout,
    QLabel, QPushButton, QLineEdit, QComboBox, QCheckBox, QSlider,
    QFrame, QSplitter, QGroupBox, QTextEdit, QStackedWidget, QDialog, QDialogButtonBox,
    QMessageBox, QFileDialog, QColorDialog, QProgressDialog,
    QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem,
    QSizePolicy, QSpinBox, QDoubleSpinBox, QInputDialog, QAbstractItemView,
    QTextBrowser, QProgressBar, QHeaderView,
)
from PySide6.QtCore import (
    Qt, QTimer, QThread, Signal, QSize, QSettings, QRect,
)
from PySide6.QtGui import (
    QPalette, QColor, QFont, QPainter, QBrush, QKeySequence, QShortcut,
    QPixmap, QIcon, QCursor, QImage,
)

try:
    from PIL import Image as _PILImage, ImageDraw as _ImageDraw
    _HAS_PIL = True
except ImportError:
    _HAS_PIL = False

try:
    from filament_mixer import GPMixer as _GPMixer
    _HAS_FILAMENT_MIXER = True
except ImportError:
    _HAS_FILAMENT_MIXER = False

try:
    import matplotlib
    matplotlib.use("QtAgg")
    import matplotlib.pyplot as _plt
    _HAS_MPL = True
except ImportError:
    _HAS_MPL = False

# ── TRANSLATIONS ──────────────────────────────────────────────────────────────

STRINGS = {
"de": {
    "app_title": "U1 FullSpectrum Ultimate — PySide6 Edition",
    "lang_btn": "🌐 EN",
    "phys_heads_title": "Physische Druckköpfe  T1–T4",
    "phys_heads_desc": "Diese 4 Filamente sind real im Drucker geladen.",
    "tool_header": "T{i} — WERKZEUG {i}",
    "slot_presets": "SLOT-PRESETS",
    "no_presets": "(keine Presets)",
    "btn_load": "LADEN",
    "btn_save": "SPEICHERN",
    "layer_height_label": "Schichthöhe (mm):",
    "lh_hint": "⚠ 0.08 mm empfohlen",
    "sec1_title": "EINZELFARBEN-RECHNER",
    "hex_placeholder": "#RRGGBB eingeben…",
    "gamut_warning": "⚠  Zielfarbe außerhalb des Gamuts (ΔE > 25)",
    "label_target": "Ziel",
    "label_simulated": "Simuliert",
    "length_label": "Länge: {n}",
    "auto_check": "Auto (kürzeste)",
    "btn_calculate": "SEQUENZ BERECHNEN",
    "optimizer_check": "Optimizer (24 Combos)",
    "btn_add_virtual": "➕ Als virtuellen Druckkopf",
    "sec2_title": "VIRTUELLE DRUCKKÖPFE  V5–V24",
    "sec2_desc": "bis zu {max_v} virtuelle Köpfe",
    "dlg_save_err": "Speicherfehler",
    "tdcal_calc": "TD berechnen",
    "tdcal_white_warn": "⚠ Weißes Filament nicht geeignet für TD-Kalibrierung.",
    "tdcal_result": "Geschätztes TD: {:.2f}",
    "btn_3mf": "🔬 3MF Assistent",
    "btn_del_all": "🗑 Alle löschen",
    "btn_undo": "↩ Rückgängig",
    "btn_recalc_all": "🔄 Alle neu berechnen",
    "btn_export_all": "📤 Export",
    "btn_orca_export": "🚀 → OrcaSlicer",
    "btn_3mf_write": "✏️ 3MF schreiben",
    "btn_de_overview": "📊 ΔE-Übersicht",
    "btn_recipe": "📜 Farbrezept",
    "btn_copy_all_cad": "📋 Cadence-Werte",
    "btn_batch": "🎨 Batch",
    "empty_virtual": "Noch keine virtuellen Druckköpfe.",
    "hint_pure": "  →  Reine Farbe",
    "hint_cadence": "  →  Cadence A={a}mm / B={b}mm  Pattern: {p}",
    "hint_pattern": "  →  Pattern Mode: {p}",
    "de_good": "✓ gut",
    "de_ok": "~ ok",
    "de_far": "✗ weit",
    "dlg_db_err_title": "DB-Fehler",
    "dlg_db_err_msg": "filament_db.json Fehler:\n{e}",
    "dlg_error": "Fehler",
    "dlg_saved": "Gespeichert",
    "dlg_preset_saved": 'Preset "{name}" gespeichert.',
    "dlg_fil_saved": '"{n}" gespeichert.',
    "dlg_exists": "Vorhanden",
    "dlg_exists_msg": '"{name}" existiert bereits.',
    "dlg_note": "Hinweis",
    "dlg_select_color": "Bitte zuerst eine Zielfarbe auswählen.",
    "dlg_max_virtual": "Maximum erreicht",
    "dlg_max_virtual_msg": "Bereits {max_v} virtuelle Druckköpfe.",
    "dlg_no_seq": "Bitte zuerst eine Sequenz berechnen.",
    "dlg_del_title": "Löschen",
    "dlg_del_virtual": "Alle virtuellen Druckköpfe löschen?",
    "dlg_3mf_added": "{n} virtuelle Druckköpfe hinzugefügt.",
    "dlg_export_saved": "Gespeichert:\n{path}",
    "inp_preset_name": "Preset-Name:",
    "inp_preset_title": "Preset speichern",
    "inp_fil_name": "Filament-Name:",
    "inp_name": "Name:",
    "inp_add_fil_title": "Filament zu '{b}'",
    "inp_hex": "Hex (#RRGGBB):",
    "inp_td": "TD-Wert (Standard {td}):",
    "inp_brand_name": "Marken-Name:",
    "inp_brand_title": "Neue Marke",
    "exp_title": "Export — Snapmaker U1 FullSpectrum",
    "exp_btn": "EXPORTIEREN",
    "exp_cancel": "Abbrechen",
    "exp_scope_single": "Aktuelle Sequenz",
    "exp_scope_virtual": "Alle {n} virtuellen Köpfe",
    "3mf_analysis_title": "3MF Analyse  ·  {n} Farbe(n)",
    "3mf_optimizer": "Optimizer aktivieren",
    "3mf_ready": "Bereit — 'Alle berechnen' drücken.",
    "3mf_col_target": "Zielfarbe",
    "3mf_col_seq": "Sequenz",
    "3mf_col_sim": "Simuliert",
    "3mf_col_quality": "ΔE",
    "3mf_not_calc": "— noch nicht",
    "3mf_include": "übernehmen",
    "3mf_progress": "Berechne {i}/{total} …",
    "3mf_done": "Fertig — {n} Farben.",
    "3mf_btn_calc": "⚙ Alle berechnen",
    "3mf_btn_apply": "✅ Ausgewählte übernehmen",
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
    "txt_virtual_heads": "Virtuelle Druckköpfe:",
    "txt_pure": "Reine Farbe — kein Mix",
    "txt_cadence": "Cadence A={a}mm / B={b}mm  Pattern: {p}",
    "txt_sequence": "Sequenz:",
    "txt_target": "Ziel:",
    "empty_slot": "(leer)",
    "manual_color": "(manuell)",
    "virtual_label_default": "V{vid}",
    "btn_copy": "📋 Kopieren",
    "btn_random": "🎲 Zufall",
    "copied_msg": "Kopiert!",
    "batch_title": "Batch-Farben",
    "batch_desc": "Hex-Codes (eine Farbe pro Zeile):",
    "batch_btn_calc": "⚙ Berechnen & hinzufügen",
    "batch_btn_cancel": "Abbrechen",
    "batch_done": "{n} virtuelle Köpfe hinzugefügt.",
    "batch_warn_max": "Maximum {max_v} erreicht.",
    "settings_title": "Einstellungen",
    "settings_saved": "Einstellungen gespeichert.",
    "settings_btn": "⚙ Einstellungen",
    "settings_max_virtual": "Max. virtuelle Druckköpfe:",
    "tab_calculator": "Rechner",
    "tab_virtual": "V-Köpfe",
    "tab_tools": "Werkzeuge",
    "color_picker_title": "Farbe wählen — T{i}",
    "target_picker_title": "Zielfarbe wählen",
    "btn_img_pick": "🖼 Aus Bild",
    "undo_empty": "Nichts rückgängig zu machen.",
    "colorinfo_label": "RGB: {r} {g} {b}   HSV: {h:.0f}° {s:.0f}% {v:.0f}%   Lab: {L:.0f} {a:.0f} {b_:.0f}",
    "btn_new_brand": "＋ Neue Marke",
    "btn_library": "📚 Bibliothek",
    "btn_web_update": "🌐 Online-Update",
    "web_update_title": "Datenbank aktualisieren",
    "web_update_fetching": "Lade Daten von GitHub …",
    "web_update_ok": "✅ {n_brands} Marken / {n_fils} Filamente.\n{new} neue Einträge.",
    "web_update_no_new": "Datenbank ist aktuell.",
    "web_update_err": "Fehler:\n{e}",
    "orca_title": "Direkt-Export → OrcaSlicer",
    "orca_header": "Filament-Profile in OrcaSlicer schreiben",
    "orca_path_label": "OrcaSlicer Profil-Ordner:",
    "orca_path_browse": "📂 Ändern",
    "orca_scope_phys": "T1–T4 (physische Slots)",
    "orca_scope_virt": "V5+ (virtuelle Köpfe)",
    "orca_scope_both": "Beides",
    "orca_prefix_label": "Profil-Prefix:",
    "orca_btn_export": "PROFILE SCHREIBEN",
    "orca_btn_cancel": "Abbrechen",
    "orca_success": "✅ {n} Profile nach OrcaSlicer geschrieben!\nOrcaSlicer neu starten.",
    "orca_no_path": "OrcaSlicer-Ordner nicht gefunden.",
    "orca_no_virtual": "Keine virtuellen Druckköpfe.",
    "btn_slot_compare": "🔀 Slot-Vergleich",
    "slot_compare_title": "Slot-Vergleich — Live ΔE-Auswirkung",
    "slot_compare_slot": "Zu vergleichender Slot:",
    "slot_compare_alt": "Alternatives Filament:",
    "slot_compare_col_vid": "V-Kopf",
    "slot_compare_col_cur": "Aktuell ΔE",
    "slot_compare_col_new": "Neu ΔE",
    "slot_compare_col_delta": "Δ",
    "orca_overwrite_confirm": "{n} Profile werden überschrieben. Fortfahren?",
    "orca_filament_notes_t": "U1 FullSpectrum — T{i} | {brand} {name} | TD={td}",
    "orca_filament_notes_v": "U1 FullSpectrum Sequenz: {seq} | ΔE={de:.1f} | {hint}",
    "orca_fs_hint": "⚠ FullSpectrum-Slicer: Nur T1–T4 exportieren!",
    "btn_proj_save": "💾 Projekt speichern",
    "btn_proj_load": "📂 Projekt laden",
    "proj_saved": "Projekt gespeichert:\n{path}",
    "proj_loaded": "Projekt geladen:\n{path}",
    "proj_filetypes": "U1-Projekt",
    "proj_err": "Ladefehler:\n{e}",
    "btn_td_cal": "🔬 TD kalibrieren",
    "td_cal_title": "TD-Kalibrierung",
    "td_cal_measured": "Gemessene Hex-Farbe:",
    "td_cal_result": "T{a} = {ta:.1f}    T{b} = {tb:.1f}",
    "td_cal_apply": "Übernehmen",
    "td_cal_slot": "Slot:",
    "top3_title": "Top-3:",
    "top3_rank": "#{r}",
    "top3_add": "+ Add",
    "btn_slot_opt": "🎯 Slot-Optimizer",
    "slot_opt_title": "Slot-Optimizer",
    "slot_opt_desc": "Beste 4 Filamente für maximalen Farbraum.",
    "slot_opt_use_virtual": "V-Köpfe als Ziele",
    "slot_opt_done": "Fertig!",
    "slot_opt_apply": "Als T1–T4 übernehmen",
    "slot_opt_result": "Bestes Set (ΔE Ø {de:.1f}):",
    "btn_palette": "🖼 Palette aus Bild",
    "palette_title": "Palette aus Bild",
    "palette_btn_add": "Alle berechnen & hinzufügen",
    "palette_indexed_hint": "Indiziertes PNG — alle {n} Palettenfarben geladen",
    "obj_btn": "📦 OBJ/MTL",
    "obj_title": "OBJ/MTL Farbanalyse",
    "obj_no_colors": "Keine Farben gefunden. MTL-Datei vorhanden?",
    "obj_no_pil": "Pillow nicht installiert — Textur-Extraktion nicht möglich.\npip install Pillow",
    "obj_filetypes": "OBJ Dateien",
    "3mf_texture_hint": "Farben aus eingebetteten Texturen extrahiert",
    "remap_title": "3MF Extruder-Remap",
    "remap_keep": "(unverändert)",
    "remap_btn_write": "3MF schreiben",
    "btn_tc_est": "🔄 Werkzeugwechsel",
    "tc_title": "Werkzeugwechsel-Schätzung",
    "tc_layers": "Druckschichten gesamt:",
    "tc_result": "TC/Schicht: {tc}   Layers bis opak: {n}",
    "tc_time": "Zusatzzeit: ~{min:.0f} min  (bei {sec}s/Wechsel)",
    "tc_purge": "Purge: ~{g:.0f}g",
    "stats_summary": "📊 {layers} Schichten · {changes} Werkzeugwechsel",
    "stats_filament_row": "  T{fid}: {cnt} Lagen ({pct:.0f}%)",
    "stats_change_time": "  ~{min:.0f} min Wechselzeit (30s/Wechsel)",
    "btn_harmonies": "🎨 Harmonien",
    "harm_title": "Farbharmonien",
    "harm_complement": "Komplementär",
    "harm_triadic": "Triade",
    "harm_analogous": "Analog",
    "harm_split": "Split-Komplement",
    "harm_add_all": "Alle hinzufügen",
    "model_label": "Farbmodell:",
    "model_linear": "Additiv (FullSpectrum)",
    "model_td": "TD-gewichtet",
    "model_subtractive": "Subtraktiv",
    "model_filamentmixer": "FilamentMixer (GPMixer/Pigment)",
    "stripe_risk": "\u26a0 Streifenrisiko \u2014 Sequenzl\u00e4nge {n} ung\u00fcnstig",
    "stripe_ok": "\u2713 Kein Streifenrisiko",
    "btn_multi_gradient": "\U0001f308 Multi-Gradient",
    "multi_gradient_title": "Multi-Gradient virtueller Kopf",
    "multi_gradient_desc": "Gewichteten Verlauf aus allen 4 Slots erstellen",
    "multi_gradient_auto": "Auto-Balance",
    "multi_gradient_add": "Als virtuellen Kopf hinzuf\u00fcgen",
    "btn_gradient": "🌈 Gradient",
    "gradient_title": "Gradient-Generator",
    "gradient_from": "Von:",
    "gradient_to": "Bis:",
    "gradient_steps": "Schritte:",
    "gradient_btn_calc": "⚙ Berechnen & hinzufügen",
    "gradient_done": "{n} Schritte hinzugefügt.",
    "de_overview_title": "ΔE-Übersicht",
    "de_overview_col_id": "ID",
    "de_overview_col_seq": "Sequenz",
    "de_overview_col_de": "ΔE",
    "de_overview_col_quality": "Qualität",
    "auto_found": "Auto: Länge {n} gefunden",
    "auto_finding": "Auto — Länge wird berechnet",
    "status_ready": "Bereit",
    "status_calculated": "Berechnet — ΔE {de:.1f} — Sequenz: {seq}",
    "status_added": "V{vid} hinzugefügt",
    "status_exported": "Exportiert: {f}",
    "status_3mf": "3MF: {n} Farben gefunden",
    "recipe_title": "Farbrezept-Export",
    "recipe_copy_btn": "In Zwischenablage",
    "btn_multitarget": "🎯 Multi-Ziel",
    "mt_title": "Mehrfach-Ziel-Optimizer",
    "mt_add_target": "+ Zielfarbe",
    "mt_calc": "⚙ Optimieren",
    "mt_result": "Sequenz: {seq}   ΔE Ø {de:.1f}",
    "mt_add_btn": "➕ Als V-Kopf",
    "mt_no_targets": "Bitte mind. 2 Zielfarben.",
    "slot_loaded": "Geladen",
    "tc_warn_badge": "⚠ {n}×WW",
    "recalc_all_done": "{n} V-Köpfe neu berechnet.",
    "btn_lab_plot": "🔭 Lab-Farbraum",
    "btn_gamut_plot": "🎯 Gamut-Plot",
    "btn_swatch": "🖼 Swatch",
    "swatch_saved": "Swatch gespeichert:\n{path}",
    "btn_slicer_guide": "📖 Slicer-Guide",
    "open_3mf_title": "3MF-Datei öffnen",
    "save_dialog_title": "Speichern",
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
    # FullSpectrum Direct 3MF Export
    "fs_export_btn": "💉 FS 3MF Export",
    "fs_export_title": "FullSpectrum Direct 3MF Export",
    "fs_export_desc": "Schreibt alle virtuellen Druckköpfe als Mixed Filaments direkt in eine .3mf Datei. Im FullSpectrum Slicer öffnen — Cadence & Pattern bereits konfiguriert.",
    "fs_src_label": "Quelldatei (.3mf):",
    "fs_dst_label": "Ausgabedatei:",
    "fs_overwrite": "Quelldatei überschreiben (Backup wird erstellt)",
    "fs_save_as": "Neue Datei speichern",
    "fs_lh_label": "Schichthöhe (mm):",
    "fs_count": "{n} virtuelle Druckköpfe werden eingeschrieben",
    "fs_preview_btn": "🔍 Vorschau",
    "fs_write_btn": "💉 Einschreiben",
    "fs_success": "Erfolgreich eingeschrieben:\n{path}",
    "fs_no_virtual": "Keine virtuellen Druckköpfe definiert.",
    "fs_no_src": "Bitte Quelldatei auswählen.",
    "fs_warn_fullspectrum": "⚠ Nur für FullSpectrum Slicer (Snapmaker_Orca).\nStandard OrcaSlicer unterstützt mixed_filament_definitions nicht.",
    "fs_guide_title": "Anleitung",
    "skin_tone_check": "Hauttöne-Modus",
    "skin_tone_mode": "Hauttöne-Modus: ΔE-Schwelle {de:.1f} aktiv",
    "batch_import_svg": "SVG importieren",
    # ── newly added keys ──
    "btn_target_color": "🎨 Zielfarbe wählen",
    "copy_all_btn": "📋 Alles kopieren",
    "copy_all_title": "Cadence-Werte aller V-Köpfe",
    "de_overview_col_label": "Bezeichnung",
    "dlg_3mf_no_colors_fallback": "Keine Farben in der 3MF-Datei gefunden.",
    "dlg_3mf_title": "3MF Assistent",
    "exp_header": "Export — Snapmaker U1 FullSpectrum",
    "exp_lh_label": "Schichthöhe (mm):",
    "exp_lh_unit": "mm",
    "gamut_plot_title": "Gamut-Plot",
    "grad_add_virtual": "➕ Als virtuelle Köpfe hinzufügen",
    "grad_added": "{n} Schritte als V-Köpfe hinzugefügt.",
    "grad_generate": "⚙ Verlauf berechnen",
    "grad_generate_first": "Bitte zuerst den Verlauf berechnen.",
    "grad_pick_stops": "Farbstopps wählen:",
    "grad_stops": "Farbstopps:",
    "guide_title": "Slicer-Einrichtungsanleitung",
    "harm_add_virtual": "➕ Als virtuelle Köpfe hinzufügen",
    "harm_added": "{n} Harmoniefarben als V-Köpfe hinzugefügt.",
    "harm_base": "Basisfarbe:",
    "harm_tetradic": "Tetraden",
    "harm_type": "Harmonieart:",
    "img_pick_title": "Farbe aus Bild wählen",
    "inp_add_title": "Filament hinzufügen",
    "inp_color_title": "Farbe eingeben",
    "inp_td2": "TD-Wert (Standard {td}):",
    "inp_td_title": "TD-Wert",
    "lab_plot_title": "Lab-Farbraum",
    "layer_height": "Schichthöhe (mm):",
    "matrix_desc": "ΔE-Abstände zwischen allen geladenen Filamenten.",
    "matrix_title": "Filament-Mischmatrix",
    "multi_add_color": "➕ Zielfarbe hinzufügen",
    "multi_best_seq": "Beste Sequenz: {}   Ø ΔE: {}",
    "multi_desc": "Optimiere eine Sequenz für mehrere Zielfarben gleichzeitig.",
    "multi_no_result": "Kein Ergebnis gefunden.",
    "multi_no_targets": "Bitte mind. 1 Zielfarbe hinzufügen.",
    "multi_optimize": "⚙ Optimieren",
    "multi_result": "Optimierungs-Ergebnis",
    "multi_title": "Mehrfach-Ziel-Optimizer",
    "no_filaments": "Keine Filamente geladen.",
    "orca_prefix_hint": "Prefix für OrcaSlicer-Profilnamen",
    "pal_add_virtual": "➕ Als virtuelle Köpfe hinzufügen",
    "pal_added": "{n} Palettenfarben als V-Köpfe hinzugefügt.",
    "pal_browse": "📂 Bild öffnen",
    "pal_colors": "Anzahl Farben:",
    "pal_no_file": "(kein Bild geladen)",
    "pal_title": "Palette aus Bild",
    "png_export_title": "PNG-Übersicht speichern",
    "png_saved": "PNG gespeichert:\n{}",
    "remap_col_hdr": "Extruder-Zuordnung neu belegen",
    "seqed_add": "➕ Hinzufügen",
    "seqed_added": "Sequenz '{}' als V-Kopf hinzugefügt.",
    "seqed_apply": "✅ Als V-Kopf übernehmen",
    "seqed_clear": "🗑 Leeren",
    "seqed_desc": "Sequenz manuell zusammenstellen und als virtuellen Kopf speichern.",
    "seqed_max": "Maximale Sequenzlänge erreicht.",
    "seqed_remove": "↩ Letzten entfernen",
    "seqed_title": "Sequenz-Editor",
    "slotopt_desc": "Beste 4 Filamente für maximalen Gamut-Abdeckung.",
    "slotopt_running": "Optimizer läuft …",
    "slotopt_start": "▶ Starten",
    "slotopt_title": "Slot-Optimizer",
    "swatch_save_title": "Swatch speichern",
    "tc_desc": "Schätze Transmissionskoeffizient und Layers-bis-Opak.",
    "td_label": "TD-Wert:",
    "tdcal_desc": "Filament-Hex und gemessene Druckfarbe eingeben, um TD zu schätzen.",
    "tdcal_fil_hex": "Filament-Hex-Farbe:",
    "tdcal_layers": "Anzahl Schichten:",
    "tdcal_measured": "Gemessene Druckfarbe (Hex):",
    "tdcal_title": "TD-Kalibrierung",
    "txt_pattern": "Pattern Mode: {p}",
    # ── UI labels & dialog strings ──
    "search_window_title": "Filament-Suche — Slot T{i}",
    "search_color_filter": "Farbfilter:",
    "search_color_tip": "Farbe wählen (Ergebnisse nach ΔE sortiert)",
    "search_color_placeholder": "#RRGGBB  (leer = kein Filter)",
    "search_slot_color": "🎯 Slot-Farbe",
    "search_slot_color_tip": "Farbe aus aktuellem Slot übernehmen",
    "search_clear_tip": "Farbfilter löschen",
    "search_placeholder": "Name / Marke suchen…",
    "search_all_brands": "Alle Marken",
    "search_orca_import": "📂 OrcaSlicer importieren",
    "search_orca_tip": "Lokale OrcaSlicer-Profile in Suche laden",
    "search_select": "Auswählen",
    "seqed_seq_empty": "Sequenz: (leer)",
    "seqed_seq_prefix": "Sequenz: ",
    "lbl_gamut": "Gamut:",
    "lbl_length": "Länge:",
    "lbl_mix_pct": "Mix %:",
    "lbl_sort": "Sortierung:",
    "lbl_print_height": "Druckhöhe:",
    "lbl_td": "TD:",
    "compare_run_btn": "▶  Vergleichen",
    "de_matrix_btn": "📊 ΔE-Matrix",
    "png_export_btn": "🖼 PNG Export",
    "txt_json_export_btn": "TXT/JSON Export",
    "tools_analysis": "Analyse & Visualisierung",
    "tools_color_gen": "Farb-Generierung",
    "tools_optimization": "Optimierung & Kalibrierung",
    "tools_library": "Bibliothek & Datenbank",
    "tools_export": "Export",
    "btn_print_stats": "📊 Print-Statistik",
    "print_stats_title": "Filament-Statistik",
    "print_stats_desc": "Filamentanteil und -verbrauch pro virtuellem Kopf.",
    "print_stats_height": "Druckhöhe (mm):",
    "print_stats_col_head": "Kopf",
    "print_stats_col_seq": "Sequenz",
    "print_stats_col_de": "ΔE",
    "btn_layer_preview": "🔬 Schicht-Vorschau",
    "layer_preview_title": "Schichtfolgen-Vorschau",
    "layer_preview_layers": "Schichten anzeigen:",
    "lbl_local_z": "Local-Z Dithering",
    "lbl_local_z_tip": "dithering_local_z_mode=1: Jede bemalte Zone erhält eigene Z-Höhenkontrolle.\nVerbessert Qualität bei Multi-Zonen-Drucken (empfohlen ab FS v0.7).",
    "lbl_adv_dither": "Advanced Dithering",
    "lbl_adv_dither_tip": "mixed_filament_advanced_dithering=1: Erweiterte Dithering-Kontrolle.\n(Experimentell — nur für erfahrene Nutzer)",
    "de_quality_excellent": "ausgezeichnet",
    "de_quality_good": "gut",
    "de_quality_visible": "sichtbar",
    # ── newly added translation keys ──
    "slot_undo_btn": "↩ Slot",
    "slot_undo_tip": "Slot-Änderung rückgängig",
    "project_groupbox": "Projekt",
    "transluc_check": "🔆 Transluz.",
    "transluc_tip": "Per-Slot: Beer-Lambert TD-Modell für dieses Filament.\nÜberschreibt das globale Farbmodell für diesen Slot.",
    "de_thresh_label": "ΔE≤",
    "skin_tone_tip": "Hautton-Modus: Engere ΔE-Schwelle (1.5) global,\n1.0 für LAB im Bereich L*40–80 a*5–25 b*10–30.",
    "mix_pct_tip": "Direkte Prozent-Eingabe für 2-Filament-Mischung.\nBeispiel: 30% → Sequenz 1222 (T2 dominiert)",
    "mix_seq_btn": "→ Seq",
    "mix_seq_tip": "Sequenz aus %-Verhältnis erzeugen (nur 2 Filamente)",
    "seq_click_copy_hint": "(Klick zum Kopieren)",
    "seq_label_click_tip": "Klick zum Kopieren",
    "seq_preview_tip": "Sequenz-Vorschau: jeder Block = 1 Schicht im Zyklus",
    "hist_groupbox": "🕘 Verlauf",
    "sort_added": "Hinzugefügt",
    "sort_de_asc": "ΔE ↑ (besser zuerst)",
    "sort_de_desc": "ΔE ↓",
    "sort_label_az": "Label A-Z",
    "3mf_filetypes": "3MF-Dateien",
    "wizard_cov_color": "Farbe",
    "wizard_cov_hex": "Hex",
    "wizard_cov_de": "ΔE",
    "wizard_cov_quality": "Qualität",
    "wizard_load_spinner": "⏳ Lade & analysiere 3MF …",
    "fs_preview_groupbox": "mixed_filament_definitions Vorschau",
    "fs_lh_warn": "⚠ >0.15 mm → Streifen!",
    "head_label": "Kopf:",
    "lh_warn_striping": "⚠ >0.15mm",
    "auto_suggest_tip": "💡 Tipp: «{brand} {name}» könnte ΔE auf ~{de:.1f} senken",
    "material_warn": "⚠ Materialwarnung: {a} + {b} — unterschiedliche Drucktemperaturen!",
    "web_update_downloading": "⏳ Community-DB wird heruntergeladen…",
    "web_update_added": "✅ {added} neue Filamente hinzugefügt.",
},
"en": {
    "app_title": "U1 FullSpectrum Ultimate — PySide6 Edition",
    "lang_btn": "🌐 DE",
    "phys_heads_title": "Physical Print Heads  T1–T4",
    "phys_heads_desc": "These 4 filaments are physically loaded.",
    "tool_header": "T{i} — TOOL {i}",
    "slot_presets": "SLOT PRESETS",
    "no_presets": "(no presets)",
    "btn_load": "LOAD",
    "btn_save": "SAVE",
    "layer_height_label": "Layer Height (mm):",
    "lh_hint": "⚠ 0.08 mm recommended",
    "sec1_title": "SINGLE COLOR CALCULATOR",
    "hex_placeholder": "Enter #RRGGBB…",
    "gamut_warning": "⚠  Target out of gamut (ΔE > 25)",
    "label_target": "Target",
    "label_simulated": "Simulated",
    "length_label": "Length: {n}",
    "auto_check": "Auto (shortest)",
    "btn_calculate": "CALCULATE SEQUENCE",
    "optimizer_check": "Optimizer (24 Combos)",
    "btn_add_virtual": "➕ Add as Virtual Head",
    "sec2_title": "VIRTUAL PRINT HEADS  V5–V24",
    "sec2_desc": "up to {max_v} virtual heads",
    "dlg_save_err": "Save Error",
    "tdcal_calc": "Calculate TD",
    "tdcal_white_warn": "⚠ White filament not suitable for TD calibration.",
    "tdcal_result": "Estimated TD: {:.2f}",
    "btn_3mf": "🔬 3MF Assistant",
    "btn_del_all": "🗑 Delete All",
    "btn_undo": "↩ Undo",
    "btn_recalc_all": "🔄 Recalc All",
    "btn_export_all": "📤 Export",
    "btn_orca_export": "🚀 → OrcaSlicer",
    "btn_3mf_write": "✏️ Write 3MF",
    "btn_de_overview": "📊 ΔE Overview",
    "btn_recipe": "📜 Recipe",
    "btn_copy_all_cad": "📋 Cadence Values",
    "btn_batch": "🎨 Batch",
    "empty_virtual": "No virtual print heads yet.",
    "hint_pure": "  →  Pure color",
    "hint_cadence": "  →  Cadence A={a}mm / B={b}mm  Pattern: {p}",
    "hint_pattern": "  →  Pattern Mode: {p}",
    "de_good": "✓ good",
    "de_ok": "~ ok",
    "de_far": "✗ far",
    "dlg_db_err_title": "DB Error",
    "dlg_db_err_msg": "filament_db.json error:\n{e}",
    "dlg_error": "Error",
    "dlg_saved": "Saved",
    "dlg_preset_saved": 'Preset "{name}" saved.',
    "dlg_fil_saved": '"{n}" saved.',
    "dlg_exists": "Exists",
    "dlg_exists_msg": '"{name}" already exists.',
    "dlg_note": "Note",
    "dlg_select_color": "Please select a target color first.",
    "dlg_max_virtual": "Maximum reached",
    "dlg_max_virtual_msg": "Already {max_v} virtual heads.",
    "dlg_no_seq": "Please calculate a sequence first.",
    "dlg_del_title": "Delete",
    "dlg_del_virtual": "Delete all virtual print heads?",
    "dlg_3mf_added": "{n} virtual heads added.",
    "dlg_export_saved": "Saved:\n{path}",
    "inp_preset_name": "Preset name:",
    "inp_preset_title": "Save Preset",
    "inp_fil_name": "Filament name:",
    "inp_name": "Name:",
    "inp_add_fil_title": "Add filament to '{b}'",
    "inp_hex": "Hex (#RRGGBB):",
    "inp_td": "TD value (default {td}):",
    "inp_brand_name": "Brand name:",
    "inp_brand_title": "New Brand",
    "exp_title": "Export — Snapmaker U1 FullSpectrum",
    "exp_btn": "EXPORT",
    "exp_cancel": "Cancel",
    "exp_scope_single": "Current sequence",
    "exp_scope_virtual": "All {n} virtual heads",
    "3mf_analysis_title": "3MF Analysis  ·  {n} color(s)",
    "3mf_optimizer": "Enable optimizer",
    "3mf_ready": "Ready — click 'Calculate All'.",
    "3mf_col_target": "Target",
    "3mf_col_seq": "Sequence",
    "3mf_col_sim": "Simulated",
    "3mf_col_quality": "ΔE",
    "3mf_not_calc": "— not yet",
    "3mf_include": "include",
    "3mf_progress": "Calculating {i}/{total} …",
    "3mf_done": "Done — {n} colors.",
    "3mf_btn_calc": "⚙ Calculate All",
    "3mf_btn_apply": "✅ Add Selected",
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
    "txt_virtual_heads": "Virtual Print Heads:",
    "txt_pure": "Pure color — no mix",
    "txt_cadence": "Cadence A={a}mm / B={b}mm  Pattern: {p}",
    "txt_sequence": "Sequence:",
    "txt_target": "Target:",
    "empty_slot": "(empty)",
    "manual_color": "(manual)",
    "virtual_label_default": "V{vid}",
    "btn_copy": "📋 Copy",
    "btn_random": "🎲 Random",
    "copied_msg": "Copied!",
    "batch_title": "Batch Colors",
    "batch_desc": "Enter hex codes (one per line):",
    "batch_btn_calc": "⚙ Calculate & Add",
    "batch_btn_cancel": "Cancel",
    "batch_done": "{n} virtual heads added.",
    "batch_warn_max": "Maximum {max_v} reached.",
    "settings_title": "Settings",
    "settings_saved": "Settings saved.",
    "settings_btn": "⚙ Settings",
    "settings_max_virtual": "Max. virtual heads:",
    "tab_calculator": "Calculator",
    "tab_virtual": "Virtual Heads",
    "tab_tools": "Tools",
    "color_picker_title": "Pick Color — T{i}",
    "target_picker_title": "Pick Target Color",
    "btn_img_pick": "🖼 From Image",
    "undo_empty": "Nothing to undo.",
    "colorinfo_label": "RGB: {r} {g} {b}   HSV: {h:.0f}° {s:.0f}% {v:.0f}%   Lab: {L:.0f} {a:.0f} {b_:.0f}",
    "btn_new_brand": "＋ New Brand",
    "btn_library": "📚 Library",
    "btn_web_update": "🌐 Online Update",
    "web_update_title": "Update Database",
    "web_update_fetching": "Fetching from GitHub …",
    "web_update_ok": "✅ {n_brands} brands / {n_fils} filaments.\n{new} new entries.",
    "web_update_no_new": "Database is up to date.",
    "web_update_err": "Error:\n{e}",
    "orca_title": "Direct Export → OrcaSlicer",
    "orca_header": "Write Filament Profiles to OrcaSlicer",
    "orca_path_label": "OrcaSlicer profile folder:",
    "orca_path_browse": "📂 Browse",
    "orca_scope_phys": "T1–T4 (physical slots)",
    "orca_scope_virt": "V5+ (virtual heads)",
    "orca_scope_both": "Both",
    "orca_prefix_label": "Profile prefix:",
    "orca_btn_export": "WRITE PROFILES",
    "orca_btn_cancel": "Cancel",
    "orca_success": "✅ {n} profiles written!\nRestart OrcaSlicer.",
    "orca_no_path": "OrcaSlicer folder not found.",
    "orca_no_virtual": "No virtual heads.",
    "btn_slot_compare": "🔀 Slot Compare",
    "slot_compare_title": "Slot Compare — Live ΔE Impact",
    "slot_compare_slot": "Slot to compare:",
    "slot_compare_alt": "Alternative filament:",
    "slot_compare_col_vid": "V-Head",
    "slot_compare_col_cur": "Current ΔE",
    "slot_compare_col_new": "New ΔE",
    "slot_compare_col_delta": "Δ",
    "orca_overwrite_confirm": "{n} profiles may be overwritten. Continue?",
    "orca_filament_notes_t": "U1 FullSpectrum — T{i} | {brand} {name} | TD={td}",
    "orca_filament_notes_v": "U1 FullSpectrum: {seq} | ΔE={de:.1f} | {hint}",
    "orca_fs_hint": "⚠ FullSpectrum slicer: Export T1–T4 only!",
    "btn_proj_save": "💾 Save Project",
    "btn_proj_load": "📂 Load Project",
    "proj_saved": "Project saved:\n{path}",
    "proj_loaded": "Project loaded:\n{path}",
    "proj_filetypes": "U1 Project",
    "proj_err": "Load error:\n{e}",
    "btn_td_cal": "🔬 Calibrate TD",
    "td_cal_title": "TD Calibration",
    "td_cal_measured": "Measured hex color:",
    "td_cal_result": "T{a} = {ta:.1f}    T{b} = {tb:.1f}",
    "td_cal_apply": "Apply",
    "td_cal_slot": "Slot:",
    "top3_title": "Top-3:",
    "top3_rank": "#{r}",
    "top3_add": "+ Add",
    "btn_slot_opt": "🎯 Slot Optimizer",
    "slot_opt_title": "Slot Optimizer",
    "slot_opt_desc": "Best 4 filaments for maximum gamut.",
    "slot_opt_use_virtual": "Use virtual heads as targets",
    "slot_opt_done": "Done!",
    "slot_opt_apply": "Apply as T1–T4",
    "slot_opt_result": "Best set (avg ΔE {de:.1f}):",
    "btn_palette": "🖼 Palette from Image",
    "palette_title": "Palette from Image",
    "palette_btn_add": "Calculate all & add",
    "palette_indexed_hint": "Indexed PNG — all {n} palette colors loaded",
    "obj_btn": "📦 OBJ/MTL",
    "obj_title": "OBJ/MTL Color Analysis",
    "obj_no_colors": "No colors found. Is the MTL file present?",
    "obj_no_pil": "Pillow not installed — texture extraction unavailable.\npip install Pillow",
    "obj_filetypes": "OBJ Files",
    "3mf_texture_hint": "Colors extracted from embedded textures",
    "remap_title": "3MF Extruder Remap",
    "remap_keep": "(keep unchanged)",
    "remap_btn_write": "Write 3MF",
    "btn_tc_est": "🔄 Tool Changes",
    "tc_title": "Tool Change Estimator",
    "tc_layers": "Total print layers:",
    "tc_result": "TC/layer: {tc}   Layers to opaque: {n}",
    "tc_time": "Extra time: ~{min:.0f} min  (at {sec}s/change)",
    "tc_purge": "Purge: ~{g:.0f}g",
    "stats_summary": "📊 {layers} layers · {changes} tool changes",
    "stats_filament_row": "  T{fid}: {cnt} layers ({pct:.0f}%)",
    "stats_change_time": "  ~{min:.0f} min change time (30s/change)",
    "btn_harmonies": "🎨 Harmonies",
    "harm_title": "Color Harmonies",
    "harm_complement": "Complement",
    "harm_triadic": "Triadic",
    "harm_analogous": "Analogous",
    "harm_split": "Split-Complement",
    "harm_add_all": "Calculate all & add",
    "model_label": "Color model:",
    "model_linear": "Additive (FullSpectrum)",
    "model_td": "TD-weighted",
    "model_subtractive": "Subtractive",
    "model_filamentmixer": "FilamentMixer (GPMixer/Pigment)",
    "stripe_risk": "\u26a0 Stripe risk \u2014 sequence length {n} unfavorable",
    "stripe_ok": "\u2713 No stripe risk",
    "btn_multi_gradient": "\U0001f308 Multi-Gradient",
    "multi_gradient_title": "Multi-Gradient Virtual Head",
    "multi_gradient_desc": "Create weighted gradient from all 4 slots",
    "multi_gradient_auto": "Auto-Balance",
    "multi_gradient_add": "Add as Virtual Head",
    "btn_gradient": "🌈 Gradient",
    "gradient_title": "Gradient Generator",
    "gradient_from": "From:",
    "gradient_to": "To:",
    "gradient_steps": "Steps:",
    "gradient_btn_calc": "⚙ Calculate & Add",
    "gradient_done": "{n} steps added.",
    "de_overview_title": "ΔE Overview",
    "de_overview_col_id": "ID",
    "de_overview_col_seq": "Sequence",
    "de_overview_col_de": "ΔE",
    "de_overview_col_quality": "Quality",
    "auto_found": "Auto: length {n} found",
    "auto_finding": "Auto — calculating length",
    "status_ready": "Ready",
    "status_calculated": "Calculated — ΔE {de:.1f} — Sequence: {seq}",
    "status_added": "V{vid} added",
    "status_exported": "Exported: {f}",
    "status_3mf": "3MF: {n} colors found",
    "recipe_title": "Color Recipe Export",
    "recipe_copy_btn": "Copy to Clipboard",
    "btn_multitarget": "🎯 Multi-Target",
    "mt_title": "Multi-Target Optimizer",
    "mt_add_target": "+ Add Target",
    "mt_calc": "⚙ Optimize",
    "mt_result": "Sequence: {seq}   avg ΔE {de:.1f}",
    "mt_add_btn": "➕ Add as Virtual Head",
    "mt_no_targets": "Please add at least 2 targets.",
    "slot_loaded": "Loaded",
    "tc_warn_badge": "⚠ {n}×TC",
    "recalc_all_done": "{n} virtual heads recalculated.",
    "btn_lab_plot": "🔭 Lab Space",
    "btn_gamut_plot": "🎯 Gamut Plot",
    "btn_swatch": "🖼 Swatch",
    "swatch_saved": "Swatch saved:\n{path}",
    "btn_slicer_guide": "📖 Slicer Guide",
    "open_3mf_title": "Open 3MF File",
    "save_dialog_title": "Save",
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
    # FullSpectrum Direct 3MF Export
    "fs_export_btn": "💉 FS 3MF Export",
    "fs_export_title": "FullSpectrum Direct 3MF Export",
    "fs_export_desc": "Writes all virtual print heads as Mixed Filaments directly into a .3mf file. Open in FullSpectrum slicer — Cadence & Pattern already configured.",
    "fs_src_label": "Source file (.3mf):",
    "fs_dst_label": "Output file:",
    "fs_overwrite": "Overwrite source file (backup will be created)",
    "fs_save_as": "Save as new file",
    "fs_lh_label": "Layer height (mm):",
    "fs_count": "{n} virtual print heads will be written",
    "fs_preview_btn": "🔍 Preview",
    "fs_write_btn": "💉 Write",
    "fs_success": "Successfully written:\n{path}",
    "fs_no_virtual": "No virtual print heads defined.",
    "fs_no_src": "Please select a source file.",
    "fs_warn_fullspectrum": "⚠ For FullSpectrum Slicer (Snapmaker_Orca) only.\nStandard OrcaSlicer does not support mixed_filament_definitions.",
    "fs_guide_title": "Instructions",
    "skin_tone_check": "Skin-Tone Mode",
    "skin_tone_mode": "Skin-Tone Mode: ΔE threshold {de:.1f} active",
    "batch_import_svg": "Import SVG",
    # ── newly added keys ──
    "btn_target_color": "🎨 Pick Target Color",
    "copy_all_btn": "📋 Copy All",
    "copy_all_title": "Cadence Values — All Virtual Heads",
    "de_overview_col_label": "Label",
    "dlg_3mf_no_colors_fallback": "No colors found in 3MF file.",
    "dlg_3mf_title": "3MF Assistant",
    "exp_header": "Export — Snapmaker U1 FullSpectrum",
    "exp_lh_label": "Layer Height (mm):",
    "exp_lh_unit": "mm",
    "gamut_plot_title": "Gamut Plot",
    "grad_add_virtual": "➕ Add as Virtual Heads",
    "grad_added": "{n} steps added as virtual heads.",
    "grad_generate": "⚙ Generate Gradient",
    "grad_generate_first": "Please generate the gradient first.",
    "grad_pick_stops": "Pick color stops:",
    "grad_stops": "Color stops:",
    "guide_title": "Slicer Setup Guide",
    "harm_add_virtual": "➕ Add as Virtual Heads",
    "harm_added": "{n} harmony colors added as virtual heads.",
    "harm_base": "Base color:",
    "harm_tetradic": "Tetradic",
    "harm_type": "Harmony type:",
    "img_pick_title": "Pick Color from Image",
    "inp_add_title": "Add Filament",
    "inp_color_title": "Enter Color",
    "inp_td2": "TD value (default {td}):",
    "inp_td_title": "TD Value",
    "lab_plot_title": "Lab Color Space",
    "layer_height": "Layer Height (mm):",
    "matrix_desc": "ΔE distances between all loaded filaments.",
    "matrix_title": "Filament Mix Matrix",
    "multi_add_color": "➕ Add Target Color",
    "multi_best_seq": "Best sequence: {}   avg ΔE: {}",
    "multi_desc": "Optimize a sequence for multiple target colors simultaneously.",
    "multi_no_result": "No result found.",
    "multi_no_targets": "Please add at least 1 target color.",
    "multi_optimize": "⚙ Optimize",
    "multi_result": "Optimization Result",
    "multi_title": "Multi-Target Optimizer",
    "no_filaments": "No filaments loaded.",
    "orca_prefix_hint": "Prefix for OrcaSlicer profile names",
    "pal_add_virtual": "➕ Add as Virtual Heads",
    "pal_added": "{n} palette colors added as virtual heads.",
    "pal_browse": "📂 Open Image",
    "pal_colors": "Number of colors:",
    "pal_no_file": "(no image loaded)",
    "pal_title": "Palette from Image",
    "png_export_title": "Save PNG Summary",
    "png_saved": "PNG saved:\n{}",
    "remap_col_hdr": "Remap Extruder Assignments",
    "seqed_add": "➕ Add",
    "seqed_added": "Sequence '{}' added as virtual head.",
    "seqed_apply": "✅ Add as Virtual Head",
    "seqed_clear": "🗑 Clear",
    "seqed_desc": "Manually compose a sequence and save as a virtual head.",
    "seqed_max": "Maximum sequence length reached.",
    "seqed_remove": "↩ Remove Last",
    "seqed_title": "Sequence Editor",
    "slotopt_desc": "Find the best 4 filaments for maximum gamut coverage.",
    "slotopt_running": "Optimizer running …",
    "slotopt_start": "▶ Start",
    "slotopt_title": "Slot Optimizer",
    "swatch_save_title": "Save Swatch",
    "tc_desc": "Estimate transmission coefficient and layers-to-opaque.",
    "td_label": "TD value:",
    "tdcal_desc": "Enter filament hex and measured print color to estimate TD.",
    "tdcal_fil_hex": "Filament hex color:",
    "tdcal_layers": "Number of layers:",
    "tdcal_measured": "Measured print color (hex):",
    "tdcal_title": "TD Calibration",
    "txt_pattern": "Pattern Mode: {p}",
    # ── UI labels & dialog strings ──
    "search_window_title": "Filament Search — Slot T{i}",
    "search_color_filter": "Color filter:",
    "search_color_tip": "Pick color (results sorted by ΔE)",
    "search_color_placeholder": "#RRGGBB  (empty = no filter)",
    "search_slot_color": "🎯 Slot Color",
    "search_slot_color_tip": "Use color from current slot",
    "search_clear_tip": "Clear color filter",
    "search_placeholder": "Search name / brand…",
    "search_all_brands": "All Brands",
    "search_orca_import": "📂 Import OrcaSlicer",
    "search_orca_tip": "Load local OrcaSlicer profiles into search",
    "search_select": "Select",
    "seqed_seq_empty": "Sequence: (empty)",
    "seqed_seq_prefix": "Sequence: ",
    "lbl_gamut": "Gamut:",
    "lbl_length": "Length:",
    "lbl_mix_pct": "Mix %:",
    "lbl_sort": "Sort:",
    "lbl_print_height": "Print height:",
    "lbl_td": "TD:",
    "compare_run_btn": "▶  Compare",
    "de_matrix_btn": "📊 ΔE Matrix",
    "png_export_btn": "🖼 PNG Export",
    "txt_json_export_btn": "TXT/JSON Export",
    "tools_analysis": "Analysis & Visualization",
    "tools_color_gen": "Color Generation",
    "tools_optimization": "Optimization & Calibration",
    "tools_library": "Library & Database",
    "tools_export": "Export",
    "btn_print_stats": "📊 Print Stats",
    "print_stats_title": "Filament Statistics",
    "print_stats_desc": "Filament fraction and usage per virtual head.",
    "print_stats_height": "Print height (mm):",
    "print_stats_col_head": "Head",
    "print_stats_col_seq": "Sequence",
    "print_stats_col_de": "ΔE",
    "btn_layer_preview": "🔬 Layer Preview",
    "layer_preview_title": "Layer Sequence Preview",
    "layer_preview_layers": "Layers to show:",
    "lbl_local_z": "Local-Z Dithering",
    "lbl_local_z_tip": "dithering_local_z_mode=1: Each painted zone gets its own Z-height control.\nImproves quality for multi-zone prints (recommended from FS v0.7).",
    "lbl_adv_dither": "Advanced Dithering",
    "lbl_adv_dither_tip": "mixed_filament_advanced_dithering=1: Advanced dithering control.\n(Experimental — for experienced users only)",
    "de_quality_excellent": "excellent",
    "de_quality_good": "good",
    "de_quality_visible": "visible",
    # ── newly added translation keys ──
    "slot_undo_btn": "↩ Slot",
    "slot_undo_tip": "Undo slot change",
    "project_groupbox": "Project",
    "transluc_check": "🔆 Transluc.",
    "transluc_tip": "Per-slot: use Beer-Lambert TD model.\nOverrides global model for this slot.",
    "de_thresh_label": "ΔE≤",
    "skin_tone_tip": "Skin-Tone Mode: tighter ΔE threshold (1.5) globally,\n1.0 for LAB in range L*40–80 a*5–25 b*10–30.",
    "mix_pct_tip": "Direct percentage input for 2-filament mix.\nExample: 30% → sequence 1222 (T2 dominant)",
    "mix_seq_btn": "→ Seq",
    "mix_seq_tip": "Generate sequence from % ratio (2 filaments only)",
    "seq_click_copy_hint": "(Click to copy)",
    "seq_label_click_tip": "Click to copy",
    "seq_preview_tip": "Sequence preview: each block = 1 layer in cycle",
    "hist_groupbox": "🕘 History",
    "sort_added": "Added",
    "sort_de_asc": "ΔE ↑ (best first)",
    "sort_de_desc": "ΔE ↓",
    "sort_label_az": "Label A-Z",
    "3mf_filetypes": "3MF Files",
    "wizard_cov_color": "Color",
    "wizard_cov_hex": "Hex",
    "wizard_cov_de": "ΔE",
    "wizard_cov_quality": "Quality",
    "wizard_load_spinner": "⏳ Loading & analyzing 3MF …",
    "fs_preview_groupbox": "mixed_filament_definitions preview",
    "fs_lh_warn": "⚠ >0.15 mm → striping!",
    "head_label": "Head:",
    "lh_warn_striping": "⚠ >0.15mm",
    "auto_suggest_tip": "💡 Tip: «{brand} {name}» could reduce ΔE to ~{de:.1f}",
    "material_warn": "⚠ Material warning: {a} + {b} — different print temperatures!",
    "web_update_downloading": "⏳ Downloading community DB…",
    "web_update_added": "✅ {added} new filaments added.",
},
}

_SLOT_SKIP = {"(leer)", "(empty)", "(manuell)", "(manual)"}
MAX_SEQ_LEN      = 48
DEFAULT_TD       = 5.0
DE_GOOD          = 3.0
DE_OK            = 6.0
GAMUT_WARN_DE    = 25.0
MAX_VIRTUAL      = 20
MAX_VIRTUAL_HARD = 24

DEFAULT_LIBRARY = {
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
    "Bambu Lab Matte": [
        {"name": "Ivory White",      "hex": "#F2EFDF", "td": 7.0},
        {"name": "Beige",            "hex": "#E8D8B8", "td": 6.0},
        {"name": "Matte Black",      "hex": "#1A1A1A", "td": 0.3},
        {"name": "Stone Gray",       "hex": "#7B7B7B", "td": 1.5},
        {"name": "Sunflower",        "hex": "#FFC107", "td": 6.0},
        {"name": "Terracotta",       "hex": "#C1440E", "td": 3.0},
        {"name": "Brick Red",        "hex": "#A93226", "td": 2.5},
        {"name": "Dusty Rose",       "hex": "#D4848A", "td": 5.0},
        {"name": "Mauve",            "hex": "#BD7EA6", "td": 4.5},
        {"name": "Deep Purple",      "hex": "#512DA8", "td": 2.5},
        {"name": "Ocean Blue",       "hex": "#1565C0", "td": 3.0},
        {"name": "Sage Green",       "hex": "#8FAF8B", "td": 4.0},
        {"name": "Army Green",       "hex": "#4A5240", "td": 2.0},
        {"name": "Desert Sand",      "hex": "#C2B280", "td": 4.5},
        {"name": "Caramel Brown",    "hex": "#9C6330", "td": 2.5},
    ],
    "Bambu Lab Silk": [
        {"name": "Silk Gold",        "hex": "#D4A843", "td": 2.0},
        {"name": "Silk Rose Gold",   "hex": "#B76E79", "td": 2.0},
        {"name": "Silk Copper",      "hex": "#B87333", "td": 1.8},
        {"name": "Silk Silver",      "hex": "#C0C4C8", "td": 2.5},
        {"name": "Silk Black",       "hex": "#1A1A1A", "td": 0.5},
        {"name": "Silk Red",         "hex": "#B71C1C", "td": 2.5},
        {"name": "Silk Sapphire",    "hex": "#1A237E", "td": 2.0},
        {"name": "Silk Jade",        "hex": "#1A6644", "td": 2.0},
    ],
    "Prusament PLA": [
        {"name": "Vanilla White",    "hex": "#D9D4C4", "td": 7.0},
        {"name": "Chalk White",      "hex": "#F0EDDE", "td": 7.5},
        {"name": "Jet Black",        "hex": "#24292A", "td": 0.3},
        {"name": "Galaxy Black",     "hex": "#17191A", "td": 0.4},
        {"name": "Grey Matter",      "hex": "#908E8E", "td": 1.8},
        {"name": "Prusa Orange",     "hex": "#FE6E31", "td": 6.5},
        {"name": "Mango Yellow",     "hex": "#FFB21F", "td": 6.5},
        {"name": "Sunflower Yellow", "hex": "#F9D011", "td": 6.5},
        {"name": "Lipstick Red",     "hex": "#B11A29", "td": 3.0},
        {"name": "Fire Engine Red",  "hex": "#CE1F1A", "td": 3.5},
        {"name": "Azure Blue",       "hex": "#0762AD", "td": 3.5},
        {"name": "Cobalt Blue",      "hex": "#004EA1", "td": 3.5},
        {"name": "Mystic Petrol",    "hex": "#1E6E7E", "td": 3.0},
        {"name": "Lemon Yellow",     "hex": "#F8E045", "td": 7.0},
        {"name": "Mystic Green",     "hex": "#256B3A", "td": 3.0},
    ],
    "eSUN PLA+": [
        {"name": "Cold White",       "hex": "#F8F8F8", "td": 8.5},
        {"name": "Warm White",       "hex": "#FFF8E7", "td": 8.0},
        {"name": "Black",            "hex": "#0A0A0A", "td": 0.3},
        {"name": "Silver",           "hex": "#C0C0C0", "td": 2.5},
        {"name": "Fire Red",         "hex": "#CC1010", "td": 3.5},
        {"name": "Pink",             "hex": "#FF69B4", "td": 6.5},
        {"name": "Orange",           "hex": "#FF5500", "td": 5.5},
        {"name": "Yellow",           "hex": "#FFD700", "td": 6.5},
        {"name": "Grass Green",      "hex": "#00A86B", "td": 4.5},
        {"name": "Teal",             "hex": "#008B8B", "td": 3.5},
        {"name": "Blue",             "hex": "#1560BD", "td": 3.5},
        {"name": "Purple",           "hex": "#7B2FBE", "td": 3.5},
        {"name": "Magenta",          "hex": "#E040FB", "td": 7.0},
        {"name": "Skin",             "hex": "#FFDAB9", "td": 7.5},
        {"name": "Brown",            "hex": "#8B4513", "td": 2.0},
    ],
    "Polymaker PolyTerra": [
        {"name": "Cotton White",     "hex": "#F5F0EB", "td": 7.5},
        {"name": "Charcoal",         "hex": "#2D2D2D", "td": 0.4},
        {"name": "Sakura Pink",      "hex": "#F4C2C2", "td": 7.0},
        {"name": "Coral",            "hex": "#F28B66", "td": 5.5},
        {"name": "Army Red",         "hex": "#9B2335", "td": 2.5},
        {"name": "Savanna Yellow",   "hex": "#C8A020", "td": 5.0},
        {"name": "Jungle Green",     "hex": "#29AB87", "td": 4.0},
        {"name": "Forest Green",     "hex": "#228B22", "td": 3.5},
        {"name": "Midnight Blue",    "hex": "#1A1A6E", "td": 2.0},
        {"name": "Cloud Blue",       "hex": "#ADC8E6", "td": 6.5},
        {"name": "Desert Sand",      "hex": "#C2B280", "td": 4.5},
    ],
    "Hatchbox PLA": [
        {"name": "White",            "hex": "#FDFDFD", "td": 8.5},
        {"name": "Black",            "hex": "#0D0D0D", "td": 0.3},
        {"name": "Silver",           "hex": "#C0C0C0", "td": 2.5},
        {"name": "Gold",             "hex": "#CFB53B", "td": 2.5},
        {"name": "True Red",         "hex": "#CC1010", "td": 3.5},
        {"name": "Pink",             "hex": "#FF69B4", "td": 6.5},
        {"name": "Orange",           "hex": "#FF6600", "td": 5.5},
        {"name": "Yellow",           "hex": "#FFE900", "td": 6.5},
        {"name": "Teal",             "hex": "#008080", "td": 3.5},
        {"name": "Blue",             "hex": "#0047AB", "td": 3.5},
        {"name": "Green",            "hex": "#00A550", "td": 4.5},
        {"name": "Purple",           "hex": "#7B2FBE", "td": 3.5},
    ],
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
    "Bambu Lab PETG": [
        {"name": "Jade White",       "hex": "#F4F2ED", "td": 7.5},
        {"name": "Black",            "hex": "#101010", "td": 0.3},
        {"name": "Red",              "hex": "#D01515", "td": 3.5},
        {"name": "Yellow",           "hex": "#FFD600", "td": 6.5},
        {"name": "Blue",             "hex": "#0052CC", "td": 3.5},
        {"name": "Green",            "hex": "#00913F", "td": 4.5},
        {"name": "Orange",           "hex": "#FF6A00", "td": 5.5},
        {"name": "Gray",             "hex": "#8A8A8A", "td": 1.5},
    ],
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

BUILTIN_PRESET_LABELS = {
    "★ CMYK":      {"de": "★ CMYK — Max. Gamut",       "en": "★ CMYK — Max. Gamut"},
    "★ RGB+W":     {"de": "★ RGB + Weiß — Additiv",    "en": "★ RGB + White — Additive"},
    "★ Primary":   {"de": "★ Primärfarben Classic",     "en": "★ Primary Colors Classic"},
    "★ Warm":      {"de": "★ Warm-Palette",             "en": "★ Warm Tones"},
    "★ Cool":      {"de": "★ Kalt-Palette",             "en": "★ Cool Tones"},
    "★ Snapmaker": {"de": "★ Snapmaker Standard",       "en": "★ Snapmaker Standard"},
}

# ── COLOR MATH ────────────────────────────────────────────────────────────────

def hex_to_rgb(hex_str):
    s = str(hex_str).strip().lstrip("#")
    if len(s) == 8: s = s[:6]
    if len(s) != 6:
        return (128, 128, 128)
    try:
        return tuple(int(s[i:i+2], 16) for i in (0, 2, 4))
    except ValueError:
        return (128, 128, 128)

def rgb_to_hex(r, g, b):
    return "#{:02X}{:02X}{:02X}".format(int(max(0, min(255, r))),
                                         int(max(0, min(255, g))),
                                         int(max(0, min(255, b))))

def estimate_td(hex_color):
    """Estimate TD from hex brightness. Bright/saturated = higher TD (more transparent)."""
    try:
        r, g, b = hex_to_rgb(hex_color)
        brightness = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        return round(0.3 + brightness * 8.2, 1)
    except Exception:
        return 5.0

_LAB_CACHE: dict = {}

def rgb_to_lab(rgb):
    key = (int(rgb[0]), int(rgb[1]), int(rgb[2]))
    if key in _LAB_CACHE:
        return _LAB_CACHE[key]
    r, g, b = [x / 255.0 for x in key]
    r = (r / 12.92) if r <= 0.04045 else ((r + 0.055) / 1.055) ** 2.4
    g = (g / 12.92) if g <= 0.04045 else ((g + 0.055) / 1.055) ** 2.4
    b = (b / 12.92) if b <= 0.04045 else ((b + 0.055) / 1.055) ** 2.4
    x = (r * 0.4124 + g * 0.3576 + b * 0.1805) / 0.95047
    y =  r * 0.2126 + g * 0.7152 + b * 0.0722
    z = (r * 0.0193 + g * 0.1192 + b * 0.9505) / 1.08883
    x, y, z = [v**(1/3) if v > 0.008856 else 7.787*v + 16/116 for v in [x, y, z]]
    result = (116*y - 16, 500*(x - y), 200*(y - z))
    _LAB_CACHE[key] = result
    return result

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
    """CIEDE2000 perceptual color difference — more accurate than simple ΔE76."""
    L1, a1, b1 = lab1
    L2, a2, b2 = lab2
    # C*ab
    C1 = math.sqrt(a1**2 + b1**2)
    C2 = math.sqrt(a2**2 + b2**2)
    C_avg = (C1 + C2) / 2.0
    C_avg7 = C_avg**7
    G = 0.5 * (1.0 - math.sqrt(C_avg7 / (C_avg7 + 6103515625.0)))  # 25^7
    a1p = a1 * (1.0 + G);  a2p = a2 * (1.0 + G)
    C1p = math.sqrt(a1p**2 + b1**2)
    C2p = math.sqrt(a2p**2 + b2**2)
    def _h(ap, bp):
        if ap == 0.0 and bp == 0.0: return 0.0
        v = math.degrees(math.atan2(bp, ap))
        return v + 360.0 if v < 0.0 else v
    h1p = _h(a1p, b1);  h2p = _h(a2p, b2)
    dLp = L2 - L1
    dCp = C2p - C1p
    if C1p * C2p == 0.0:
        dhp = 0.0
    elif abs(h2p - h1p) <= 180.0:
        dhp = h2p - h1p
    elif h2p - h1p > 180.0:
        dhp = h2p - h1p - 360.0
    else:
        dhp = h2p - h1p + 360.0
    dHp = 2.0 * math.sqrt(C1p * C2p) * math.sin(math.radians(dhp / 2.0))
    Lp_avg = (L1 + L2) / 2.0
    Cp_avg = (C1p + C2p) / 2.0
    if C1p * C2p == 0.0:
        hp_avg = h1p + h2p
    elif abs(h1p - h2p) <= 180.0:
        hp_avg = (h1p + h2p) / 2.0
    elif h1p + h2p < 360.0:
        hp_avg = (h1p + h2p + 360.0) / 2.0
    else:
        hp_avg = (h1p + h2p - 360.0) / 2.0
    T = (1.0
         - 0.17 * math.cos(math.radians(hp_avg - 30.0))
         + 0.24 * math.cos(math.radians(2.0 * hp_avg))
         + 0.32 * math.cos(math.radians(3.0 * hp_avg + 6.0))
         - 0.20 * math.cos(math.radians(4.0 * hp_avg - 63.0)))
    SL = 1.0 + 0.015 * (Lp_avg - 50.0)**2 / math.sqrt(20.0 + (Lp_avg - 50.0)**2)
    SC = 1.0 + 0.045 * Cp_avg
    SH = 1.0 + 0.015 * Cp_avg * T
    Cp_avg7 = Cp_avg**7
    RC = 2.0 * math.sqrt(Cp_avg7 / (Cp_avg7 + 6103515625.0))
    d_theta = 30.0 * math.exp(-((hp_avg - 275.0) / 25.0)**2)
    RT = -math.sin(math.radians(2.0 * d_theta)) * RC
    return math.sqrt(
        (dLp / SL)**2 + (dCp / SC)**2 + (dHp / SH)**2
        + RT * (dCp / SC) * (dHp / SH)
    )

def find_best_4_filaments(target_labs, library_fils, progress_cb=None):
    """Find best 4 filaments from library to cover target_labs.

    Algorithm: Greedy initialisation + local search (swap).
    ~300x faster than brute-force C(n,4) — typically <1 second.

    Phase 1 — Greedy (O(n_fils * 4 * n_targets)):
      Pick each of 4 slots by choosing the filament that most reduces
      the total uncovered ΔE given the slots already chosen.

    Phase 2 — Local search (O(n_fils * 4 * iterations)):
      For each slot try replacing with every other filament.
      Keep the swap if it improves the score.  Repeat until stable.

    Returns: (best_4_list, avg_de, scores_per_target)
    """
    if not target_labs or not library_fils:
        return [], 99.0, []

    n_t = len(target_labs)
    n_f = len(library_fils)

    # Pre-compute full distance matrix (n_targets × n_fils)
    try:
        import numpy as np
        t_arr = np.array(target_labs, dtype=np.float32)
        f_arr = np.array([f["lab"] for f in library_fils], dtype=np.float32)
        diff = t_arr[:, None, :] - f_arr[None, :, :]
        dist = np.sqrt((diff ** 2).sum(axis=2))

        def score(indices):
            return float(dist[:, list(indices)].min(axis=1).mean())

    except ImportError:
        dist_py = [[delta_e(target_labs[i], library_fils[j]["lab"])
                    for j in range(n_f)]
                   for i in range(n_t)]

        def score(indices):
            idx_list = list(indices)
            total = sum(min(dist_py[i][j] for j in idx_list)
                        for i in range(n_t))
            return total / n_t

    if progress_cb:
        progress_cb(0, 100)

    # ── Phase 1: Greedy ───────────────────────────────────────────────────────
    chosen = []
    for step in range(min(4, n_f)):
        best_j, best_s = -1, float("inf")
        for j in range(n_f):
            if j in chosen:
                continue
            s = score(chosen + [j])
            if s < best_s:
                best_s, best_j = s, j
        chosen.append(best_j)
        if progress_cb:
            progress_cb(10 + step * 10, 100)

    # ── Phase 2: Local search ─────────────────────────────────────────────────
    improved = True
    iteration = 0
    while improved and iteration < 30:
        improved = False
        iteration += 1
        for slot in range(len(chosen)):
            cur_s = score(chosen)
            for j in range(n_f):
                if j in chosen:
                    continue
                trial = chosen[:]
                trial[slot] = j
                s = score(trial)
                if s < cur_s - 1e-6:
                    cur_s = s
                    chosen[slot] = j
                    improved = True
        if progress_cb:
            progress_cb(min(95, 40 + iteration * 3), 100)

    if progress_cb:
        progress_cb(100, 100)

    best_combo = [library_fils[i] for i in chosen]
    best_score = score(chosen)

    combo_labs = [f["lab"] for f in best_combo]
    scores = [min(delta_e(t, c) for c in combo_labs) for t in target_labs]

    return best_combo, best_score, scores


def safe_td(value):
    try:
        v = float(str(value).strip())
        return v if v > 0 else DEFAULT_TD
    except (ValueError, TypeError):
        return DEFAULT_TD

def de_color(de):
    return "#4ade80" if de < DE_GOOD else "#fbbf24" if de < DE_OK else "#f87171"

def de_label_text(de, lang="de"):
    s = STRINGS[lang]
    q = s["de_good"] if de < DE_GOOD else s["de_ok"] if de < DE_OK else s["de_far"]
    return f"ΔE {de:.1f}  {q}"

def seq_to_runs(sequence):
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


def build_mixed_filament_definitions(virtual_heads, layer_height=0.08,
                                      local_z=False, advanced_dithering=False):
    """Build the mixed_filament_definitions string for FullSpectrum slicer.

    Each virtual head becomes one row in the semicolon-delimited string.
    Format per row: component_a,component_b,enabled,custom,mix_b_percent,
                    pointillism,g<ids>,w<weights>,m<mode>,d<deleted>,
                    o<origin>,u<stable_id>[,manual_pattern]

    virtual_heads: list of dicts with keys:
        vid (int), sequence (str of digit chars), target_hex (str),
        sim_hex (str), de (float), label (str), stable_id (int, optional)

    Returns: (definitions_string, extra_config_dict)
    """
    rows = []
    extra_cfg = {}

    # Collect cadence heights from all 2-filament heads
    all_cadence_a = []
    all_cadence_b = []

    for i, vf in enumerate(virtual_heads):
        seq = vf.get("sequence", "")
        if not seq:
            continue

        stable_id = vf.get("stable_id", 1000 + i)
        unique_ids = list(dict.fromkeys(int(c) for c in seq))  # preserve order, dedupe
        n_unique = len(unique_ids)

        if n_unique == 0:
            continue
        elif n_unique == 1:
            # Pure color — single filament, no mix needed
            fid = unique_ids[0]
            row = f"{fid},{fid},1,1,0,0,g,w,m2,d0,o0,u{stable_id}"
            rows.append(row)
        elif n_unique == 2:
            # 2-filament cadence mode
            fid_a, fid_b = unique_ids[0], unique_ids[1]
            count_a = seq.count(str(fid_a))
            count_b = seq.count(str(fid_b))
            total = count_a + count_b
            mix_b_pct = round(count_b / total * 100) if total > 0 else 50

            # Cadence heights using slicer formula
            cad = calc_cadence(seq, layer_height)
            ids_sorted = sorted(cad.keys())
            if len(ids_sorted) >= 2:
                ca = cad.get(fid_a, layer_height)
                cb = cad.get(fid_b, layer_height)
                all_cadence_a.append(ca)
                all_cadence_b.append(cb)

            row = f"{fid_a},{fid_b},1,1,{mix_b_pct},0,g,w,m2,d0,o0,u{stable_id}"
            rows.append(row)
        else:
            # 3-4 filament: use manual pattern mode
            grad_ids = "".join(str(x) for x in unique_ids)
            total = len(seq)
            weights = [round(seq.count(str(fid)) / total * 100) for fid in unique_ids]
            # Normalize to sum 100
            diff = 100 - sum(weights)
            if diff != 0:
                weights[0] += diff
            grad_w = "/".join(str(w) for w in weights)

            # Manual pattern: compact form (e.g. "11112222") — native slicer format
            pattern = "".join(seq)

            # Use first and last unique filament as component_a/b
            fid_a, fid_b = unique_ids[0], unique_ids[-1]
            mix_b_pct = weights[-1]

            row = f"{fid_a},{fid_b},1,1,{mix_b_pct},0,g{grad_ids},w{grad_w},m2,d0,o0,u{stable_id},{pattern}"
            rows.append(row)

    # Global cadence heights: use median of all 2-filament heads
    if all_cadence_a:
        all_cadence_a.sort()
        all_cadence_b.sort()
        mid = len(all_cadence_a) // 2
        extra_cfg["mixed_color_layer_height_a"] = str(all_cadence_a[mid])
        extra_cfg["mixed_color_layer_height_b"] = str(all_cadence_b[mid])
    else:
        extra_cfg["mixed_color_layer_height_a"] = str(layer_height)
        extra_cfg["mixed_color_layer_height_b"] = str(layer_height)

    extra_cfg["dithering_z_step_size"] = str(layer_height)   # must match printer Z resolution
    extra_cfg["dithering_step_size"] = str(layer_height)     # v0.9.4+ alias
    # Process-level cadence heights (mirrors per-head values for compatibility)
    cad_a = extra_cfg.get("mixed_color_layer_height_a", str(layer_height))
    cad_b = extra_cfg.get("mixed_color_layer_height_b", str(layer_height))
    extra_cfg["dithering_cadence_height_a"] = cad_a
    extra_cfg["dithering_cadence_height_b"] = cad_b
    extra_cfg["dithering_local_z_mode"] = "1" if local_z else "0"
    extra_cfg["dithering_step_painted_zones_only"] = "1"
    extra_cfg["mixed_filament_advanced_dithering"] = "1" if advanced_dithering else "0"
    extra_cfg["mixed_filament_gradient_mode"] = "0"
    extra_cfg["mixed_filament_height_lower_bound"] = "0.04"
    extra_cfg["mixed_filament_height_upper_bound"] = str(round(layer_height * 4, 3))

    return ";".join(rows), extra_cfg


def inject_fullspectrum_into_3mf(input_path, output_path, virtual_heads,
                                   physical_filaments, layer_height=0.08,
                                   local_z=False, advanced_dithering=False):
    """Write FullSpectrum mixed_filament_definitions into a .3mf project file.

    input_path:        path to existing .3mf file (or None to create minimal skeleton)
    output_path:       path to write modified .3mf
    virtual_heads:     list of virtual head dicts (from app._virtual / virtual_fils)
    physical_filaments: list of 4 dicts with {hex, name, brand, td}
    layer_height:      layer height in mm (default 0.08)

    Returns: (success: bool, message: str)
    """
    import shutil

    if not virtual_heads:
        return False, "Keine virtuellen Druckköpfe definiert."

    # Build the definitions string
    defs_str, extra_cfg = build_mixed_filament_definitions(
        virtual_heads, layer_height, local_z=local_z, advanced_dithering=advanced_dithering)

    # Physical filament colors for filament_colour array
    phys_colors = [f.get("hex", "#808080") for f in physical_filaments if f.get("hex")]
    while len(phys_colors) < 4:
        phys_colors.append("#808080")
    phys_colors = phys_colors[:4]

    if input_path is None:
        return False, "Keine .3mf Quelldatei — bitte zuerst eine .3mf Datei auswählen."

    if not os.path.exists(input_path):
        return False, f"Datei nicht gefunden: {input_path}"

    # Create backup
    backup_path = input_path + ".bak"
    if input_path == output_path:
        shutil.copy2(input_path, backup_path)

    try:
        with zipfile.ZipFile(input_path, 'r') as zin:
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                config_written = False
                for item in zin.infolist():
                    data = zin.read(item.filename)

                    if item.filename == 'Metadata/project_settings.config':
                        # Parse existing config
                        try:
                            cfg = json.loads(data.decode('utf-8'))
                        except Exception:
                            cfg = {}

                        # Inject settings
                        cfg['mixed_filament_definitions'] = defs_str
                        cfg['filament_colour'] = phys_colors
                        cfg.update(extra_cfg)

                        data = json.dumps(cfg, indent=4, ensure_ascii=False).encode('utf-8')
                        config_written = True

                    zout.writestr(item, data)

                # If project_settings.config didn't exist, create it
                if not config_written:
                    cfg = {
                        "version": "01.09.04.52",
                        "name": "U1 FullSpectrum Export",
                        "from": "project",
                        "filament_colour": phys_colors,
                        "mixed_filament_definitions": defs_str,
                    }
                    cfg.update(extra_cfg)
                    zout.writestr(
                        'Metadata/project_settings.config',
                        json.dumps(cfg, indent=4, ensure_ascii=False).encode('utf-8')
                    )

        n_heads = len([v for v in virtual_heads if v.get("sequence")])
        msg = f"✅ {n_heads} virtuelle Druckköpfe in .3mf eingeschrieben."
        if input_path == output_path:
            msg += f"\n(Backup: {backup_path})"
        return True, msg

    except Exception as e:
        return False, f"Fehler beim Schreiben: {e}"


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
    r1f, g1f, b1f = r1/255, g1/255, b1/255
    r2f, g2f, b2f = r2/255, g2/255, b2/255

    def to_pigment(c):
        return math.sqrt(max(0.0, c))
    def from_pigment(c):
        return min(1.0, c * c)

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

    slots = []
    for fid, w in weights:
        exact = (w / total) * max_len
        n = int(exact)
        remainder_new = exact - n
        slots.append((fid, n, remainder_new))

    n_total = sum(n for _, n, _ in slots)
    deficit = max_len - n_total
    sorted_by_rem = sorted(slots, key=lambda x: x[2], reverse=True)
    final = {fid: n for fid, n, _ in slots}
    for i in range(deficit):
        fid = sorted_by_rem[i % len(sorted_by_rem)][0]
        final[fid] = final.get(fid, 0) + 1

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

HEX_RE = re.compile(r'#([0-9A-Fa-f]{6})\b')

def _extract_colors_from_3mf_textures(zip_ref):
    """Extract dominant colors from PNG/JPG textures embedded in a .3mf zip."""
    colors = []
    if not _HAS_PIL:
        return colors
    for name in zip_ref.namelist():
        lname = name.lower()
        if lname.endswith('.png') or lname.endswith('.jpg') or lname.endswith('.jpeg'):
            try:
                with zip_ref.open(name) as img_file:
                    img = _PILImage.open(img_file)
                    if img.mode == 'P':
                        palette = img.getpalette()
                        used_indices = set(img.getdata())
                        for idx in sorted(used_indices):
                            r = palette[idx * 3]
                            g = palette[idx * 3 + 1]
                            b = palette[idx * 3 + 2]
                            brightness = (r + g + b) / 3
                            if brightness < 10 or brightness > 245:
                                continue
                            colors.append(f"#{r:02X}{g:02X}{b:02X}")
                    else:
                        img = img.convert('RGB')
                        quantized = img.quantize(colors=16,
                                                 method=_PILImage.Quantize.MEDIANCUT)
                        palette_img = quantized.convert('RGB')
                        palette = palette_img.getpalette()
                        if palette:
                            for i in range(0, min(len(palette), 16 * 3), 3):
                                r, g, b = palette[i], palette[i + 1], palette[i + 2]
                                brightness = (r + g + b) / 3
                                if brightness < 10 or brightness > 245:
                                    continue
                                colors.append(f"#{r:02X}{g:02X}{b:02X}")
            except Exception:
                pass
    return colors


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
                    except Exception:
                        pass
                for m in HEX_RE.finditer(content):
                    found.add('#' + m.group(1).upper())
            # FIX 1: Also extract colors from embedded PNG/JPG textures
            texture_colors = _extract_colors_from_3mf_textures(zf)
            extracted_so_far = list(found)
            for hx in texture_colors:
                try:
                    t_lab = rgb_to_lab(hex_to_rgb(hx))
                    if all(delta_e(t_lab, rgb_to_lab(hex_to_rgb(ex))) >= 3.0
                           for ex in extracted_so_far):
                        found.add(hx)
                        extracted_so_far.append(hx)
                except Exception:
                    found.add(hx)
                    extracted_so_far.append(hx)
    except Exception as e:
        return [], str(e)
    trivial = {"#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF",
               "#AAAAAA", "#333333", "#CCCCCC"}
    meaningful = [c for c in found if c not in trivial]
    return (meaningful if meaningful else list(found)), None


def parse_obj_mtl_colors(obj_path):
    """Extract colors from OBJ+MTL file combination.
    Returns list of hex color strings.
    """
    colors = []
    obj_dir = os.path.dirname(obj_path)

    # Find MTL file(s) referenced in OBJ
    mtl_files = []
    try:
        with open(obj_path, 'r', encoding='utf-8', errors='replace') as f:
            for line in f:
                line = line.strip()
                if line.startswith('mtllib '):
                    mtl_name = line[7:].strip()
                    mtl_path = os.path.join(obj_dir, mtl_name)
                    if os.path.exists(mtl_path):
                        mtl_files.append(mtl_path)
    except Exception:
        pass

    if not mtl_files:
        base = os.path.splitext(obj_path)[0]
        candidate = base + '.mtl'
        if os.path.exists(candidate):
            mtl_files.append(candidate)

    for mtl_path in mtl_files:
        try:
            with open(mtl_path, 'r', encoding='utf-8', errors='replace') as f:
                for line in f:
                    line = line.strip()
                    if line.lower().startswith('kd '):
                        parts = line.split()
                        if len(parts) >= 4:
                            try:
                                r = max(0, min(255, int(float(parts[1]) * 255)))
                                g = max(0, min(255, int(float(parts[2]) * 255)))
                                b = max(0, min(255, int(float(parts[3]) * 255)))
                                colors.append(f"#{r:02X}{g:02X}{b:02X}")
                            except ValueError:
                                pass
                    elif line.lower().startswith('map_kd '):
                        tex_name = line.split(None, 1)[1].strip()
                        tex_path = os.path.join(obj_dir, tex_name)
                        if os.path.exists(tex_path) and _HAS_PIL:
                            try:
                                img = _PILImage.open(tex_path)
                                if img.mode == 'P':
                                    palette = img.getpalette()
                                    used = set(img.getdata())
                                    for idx in sorted(used):
                                        r = palette[idx * 3]
                                        g = palette[idx * 3 + 1]
                                        b = palette[idx * 3 + 2]
                                        if (r + g + b) / 3 < 10 or (r + g + b) / 3 > 245:
                                            continue
                                        colors.append(f"#{r:02X}{g:02X}{b:02X}")
                                else:
                                    img = img.convert('RGB')
                                    q = img.quantize(colors=16,
                                                     method=_PILImage.Quantize.MEDIANCUT)
                                    pal = q.convert('RGB').getpalette()
                                    if pal:
                                        for i in range(0, min(len(pal), 48), 3):
                                            r, g, b = pal[i], pal[i + 1], pal[i + 2]
                                            if (r + g + b) / 3 < 10 or (r + g + b) / 3 > 245:
                                                continue
                                            colors.append(f"#{r:02X}{g:02X}{b:02X}")
                            except Exception:
                                pass
        except Exception:
            pass

    # Deduplicate by ΔE < 3
    unique = []
    for hx in colors:
        try:
            lab = rgb_to_lab(hex_to_rgb(hx))
            if all(delta_e(lab, rgb_to_lab(hex_to_rgb(u))) >= 3.0 for u in unique):
                unique.append(hx)
        except Exception:
            pass
    return unique

# ── DARK STYLESHEET ───────────────────────────────────────────────────────────

DARK_QSS = """
QWidget { background-color: #0f172a; color: #e2e8f0; font-family: 'Segoe UI'; font-size: 10pt; }
QMainWindow { background-color: #0f172a; }
QScrollArea { border: none; background-color: #0f172a; }
QScrollArea > QWidget > QWidget { background-color: #0f172a; }

QPushButton {
    background-color: #1e3a5f; color: #e2e8f0;
    border: 1px solid #334155; border-radius: 5px;
    padding: 4px 10px; min-height: 26px;
}
QPushButton:hover { background-color: #2563eb; }
QPushButton:pressed { background-color: #1d4ed8; }
QPushButton:disabled { background-color: #1e293b; color: #475569; }
QPushButton#btn_primary {
    background-color: #1d4ed8; font-weight: bold; font-size: 11pt;
    min-height: 36px;
}
QPushButton#btn_primary:hover { background-color: #2563eb; }
QPushButton#btn_green { background-color: #15803d; }
QPushButton#btn_green:hover { background-color: #16a34a; }
QPushButton#btn_red { background-color: #991b1b; }
QPushButton#btn_red:hover { background-color: #dc2626; }
QPushButton#btn_small { padding: 2px 6px; min-height: 22px; font-size: 9pt; }
QPushButton#btn_swatch { border-radius: 13px; min-width: 26px; max-width: 26px; min-height: 26px; max-height: 26px; border: 2px solid #334155; }

QLineEdit {
    background-color: #1e293b; border: 1px solid #334155;
    border-radius: 4px; padding: 3px 6px; color: #e2e8f0;
}
QLineEdit:focus { border-color: #3b82f6; }

QComboBox {
    background-color: #1e293b; border: 1px solid #334155;
    border-radius: 4px; padding: 3px 6px; color: #e2e8f0;
    min-height: 26px;
}
QComboBox:hover { border-color: #3b82f6; }
QComboBox QAbstractItemView {
    background-color: #1e293b; border: 1px solid #334155;
    color: #e2e8f0; selection-background-color: #1d4ed8;
}
QComboBox::drop-down { border: none; }

QCheckBox { spacing: 6px; }
QCheckBox::indicator { width: 16px; height: 16px; border: 1px solid #334155; border-radius: 3px; background-color: #1e293b; }
QCheckBox::indicator:checked { background-color: #3b82f6; border-color: #3b82f6; }

QSlider::groove:horizontal { height: 6px; background-color: #334155; border-radius: 3px; }
QSlider::handle:horizontal { background-color: #3b82f6; border-radius: 8px; width: 16px; height: 16px; margin: -5px 0; }
QSlider::sub-page:horizontal { background-color: #3b82f6; border-radius: 3px; }

QTabWidget::pane { border: 1px solid #334155; background-color: #0f172a; }
QTabBar::tab { background-color: #1e293b; color: #94a3b8; padding: 8px 20px; border: 1px solid #334155; border-bottom: none; border-radius: 4px 4px 0 0; }
QTabBar::tab:selected { background-color: #0f172a; color: #e2e8f0; border-bottom: 2px solid #3b82f6; }
QTabBar::tab:hover:!selected { background-color: #263048; color: #e2e8f0; }

QFrame#slot_frame { background-color: #1e293b; border: 1px solid #334155; border-radius: 6px; }
QFrame#card { background-color: #1e293b; border: 1px solid #334155; border-radius: 6px; }
QFrame#dark_card { background-color: #0f172a; border: 1px solid #1e293b; border-radius: 6px; }

QLabel#section_title { color: #38bdf8; font-size: 13pt; font-weight: bold; }
QLabel#slot_title { color: #94a3b8; font-size: 10pt; font-weight: bold; }
QLabel#hint { color: #64748b; font-size: 9pt; }
QLabel#de_label { font-size: 18pt; font-weight: bold; }
QLabel#gamut_warn { color: #fbbf24; font-size: 9pt; }

QGroupBox {
    border: 1px solid #334155; border-radius: 6px;
    margin-top: 10px; padding-top: 6px;
    font-weight: bold; color: #64748b;
}
QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 4px; }

QTextEdit { background-color: #1e293b; border: 1px solid #334155; border-radius: 4px; color: #e2e8f0; }
QListWidget { background-color: #1e293b; border: 1px solid #334155; border-radius: 4px; color: #e2e8f0; alternate-background-color: #263048; }
QListWidget::item:selected { background-color: #1d4ed8; }
QTableWidget { background-color: #1e293b; border: 1px solid #334155; gridline-color: #334155; color: #e2e8f0; }
QTableWidget::item:selected { background-color: #1d4ed8; }
QHeaderView::section { background-color: #0f172a; color: #94a3b8; border: 1px solid #334155; padding: 4px; }

QScrollBar:vertical { background-color: #1e293b; width: 10px; border-radius: 5px; }
QScrollBar::handle:vertical { background-color: #334155; border-radius: 5px; min-height: 20px; }
QScrollBar::handle:vertical:hover { background-color: #475569; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal { background-color: #1e293b; height: 10px; }
QScrollBar::handle:horizontal { background-color: #334155; border-radius: 5px; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }

QDialog { background-color: #0f172a; }
QInputDialog { background-color: #0f172a; }
QMessageBox { background-color: #0f172a; }
QProgressDialog { background-color: #0f172a; }
QSpinBox { background-color: #1e293b; border: 1px solid #334155; border-radius: 4px; padding: 3px 6px; color: #e2e8f0; }
"""

LIGHT_QSS = """
QWidget { background-color: #f8fafc; color: #1e293b; font-family: 'Segoe UI'; font-size: 10pt; }
QPushButton { background-color: #e2e8f0; color: #1e293b; border: 1px solid #cbd5e1; border-radius: 5px; padding: 4px 10px; min-height: 26px; }
QPushButton:hover { background-color: #bfdbfe; }
QPushButton#btn_primary { background-color: #2563eb; color: white; font-weight: bold; font-size: 11pt; min-height: 36px; }
QPushButton#btn_green { background-color: #16a34a; color: white; }
QLineEdit { background-color: white; border: 1px solid #cbd5e1; border-radius: 4px; padding: 3px 6px; }
QComboBox { background-color: white; border: 1px solid #cbd5e1; border-radius: 4px; padding: 3px 6px; min-height: 26px; }
QComboBox QAbstractItemView { background-color: white; color: #1e293b; selection-background-color: #bfdbfe; }
QCheckBox::indicator { width: 16px; height: 16px; border: 1px solid #cbd5e1; border-radius: 3px; background-color: white; }
QCheckBox::indicator:checked { background-color: #3b82f6; border-color: #3b82f6; }
QTabWidget::pane { border: 1px solid #cbd5e1; }
QTabBar::tab { background-color: #e2e8f0; color: #64748b; padding: 8px 20px; border: 1px solid #cbd5e1; border-bottom: none; border-radius: 4px 4px 0 0; }
QTabBar::tab:selected { background-color: #f8fafc; color: #1e293b; border-bottom: 2px solid #3b82f6; }
QFrame#slot_frame { background-color: white; border: 1px solid #cbd5e1; border-radius: 6px; }
QFrame#card { background-color: white; border: 1px solid #cbd5e1; border-radius: 6px; }
QLabel#section_title { color: #1d4ed8; font-size: 13pt; font-weight: bold; }
QGroupBox { border: 1px solid #cbd5e1; border-radius: 6px; margin-top: 10px; }
QTextEdit { background-color: white; border: 1px solid #cbd5e1; border-radius: 4px; }
QListWidget { background-color: white; border: 1px solid #cbd5e1; }
QTableWidget { background-color: white; border: 1px solid #cbd5e1; }
QScrollBar:vertical { background-color: #f1f5f9; width: 10px; }
QScrollBar::handle:vertical { background-color: #cbd5e1; border-radius: 5px; min-height: 20px; }
"""

# ── GAMUT STRIP WIDGET ────────────────────────────────────────────────────────

class GamutStrip(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(16)
        self._samples = []
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self._do_paint)

    def schedule_update(self, samples):
        self._pending = samples
        self._timer.start(100)

    def _do_paint(self):
        self._samples = getattr(self, "_pending", [])
        self.update()

    def paintEvent(self, event):
        if not self._samples:
            return
        painter = QPainter(self)
        w, h = self.width(), self.height()
        n = len(self._samples)
        if n == 0:
            return
        cell_w = w / n
        for i, hex_c in enumerate(self._samples):
            try:
                r, g, b = hex_to_rgb(hex_c)
                painter.fillRect(QRect(int(i * cell_w), 0,
                                       int(cell_w) + 1, h),
                                  QColor(r, g, b))
            except Exception:
                pass
        painter.end()

# ── COLOR SWATCH LABEL ────────────────────────────────────────────────────────

class SwatchLabel(QLabel):
    """Colored square label that copies hex to clipboard on click."""
    clicked = Signal(str)

    def __init__(self, hex_color="#808080", size=36, radius=6, parent=None):
        super().__init__(parent)
        self.setFixedSize(size, size)
        self._hex = hex_color
        self._radius = radius
        self.setCursor(QCursor(Qt.PointingHandCursor))
        self._apply(hex_color)

    def set_color(self, hex_color):
        self._hex = hex_color
        self._apply(hex_color)

    def _apply(self, hex_color):
        self._hex = hex_color
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        try:
            r, g, b = hex_to_rgb(self._hex)
            color = QColor(r, g, b)
        except Exception:
            color = QColor(128, 128, 128)
        # Border
        painter.setPen(QColor(0x33, 0x41, 0x55))
        painter.setBrush(color)
        rect = self.rect().adjusted(1, 1, -1, -1)
        painter.drawRoundedRect(rect, self._radius, self._radius)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            QApplication.clipboard().setText(self._hex.upper())
            self.clicked.emit(self._hex)
        super().mousePressEvent(event)


# ── CLICKABLE LABEL ───────────────────────────────────────────────────────────

class ClickableLabel(QLabel):
    """A QLabel that copies its text to clipboard on click and flashes green."""

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setCursor(QCursor(Qt.PointingHandCursor))
        self._original_style = ""

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            txt = self.text().strip()
            if txt and txt != "----------":
                QApplication.clipboard().setText(txt)
                self._flash()
        super().mousePressEvent(event)

    def _flash(self):
        self._original_style = self.styleSheet()
        self.setStyleSheet(self._original_style + " background-color: #16a34a; border-radius: 4px;")
        QTimer.singleShot(600, self._restore_style)

    def _restore_style(self):
        self.setStyleSheet(self._original_style)


_SLOT_SKIP = {"(leer)", "(empty)", "(manuell)", "(manual)"}

_COMMUNITY_URL = (
    "https://raw.githubusercontent.com/halloworld007/"
    "snapmaker-u1-fullspectrum-helper/main/filament_community.json"
)

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

BUILTIN_PRESET_LABELS = {
    "★ CMYK":      {"de": "★ CMYK — Max. Gamut",       "en": "★ CMYK — Max. Gamut"},
    "★ RGB+W":     {"de": "★ RGB + Weiß — Additiv",    "en": "★ RGB + White — Additive"},
    "★ Primary":   {"de": "★ Primärfarben Classic",     "en": "★ Primary Colors Classic"},
    "★ Warm":      {"de": "★ Warm-Palette",             "en": "★ Warm Tones"},
    "★ Cool":      {"de": "★ Kalt-Palette",             "en": "★ Cool Tones"},
    "★ Snapmaker": {"de": "★ Snapmaker Standard",       "en": "★ Snapmaker Standard"},
}

# ── HELPER FUNCTIONS (not yet defined) ────────────────────────────────────────

def _get_layer_weights(n):
    if n == 1:
        return [1.0]
    return [1.0 + 0.5 * i / (n - 1) for i in range(n)]

def _de_color(de):
    return "#4ade80" if de < DE_GOOD else "#fbbf24" if de < DE_OK else "#f87171"

def _seq_filament_count(sequence):
    return len(set(sequence))

def _de_label_text(de, lang="de"):
    s = STRINGS[lang]
    q = s["de_good"] if de < DE_GOOD else s["de_ok"] if de < DE_OK else s["de_far"]
    return f"ΔE {de:.1f}  {q}"

HEX_RE = re.compile(r'#([0-9A-Fa-f]{6})\b')

def _parse_3mf_colors(filepath):
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
                                    if m:
                                        found.add('#' + m.group(1).upper())
                    except (json.JSONDecodeError, AttributeError):
                        pass
                for m in HEX_RE.finditer(content):
                    found.add('#' + m.group(1).upper())
            # FIX 1: Also extract colors from embedded PNG/JPG textures
            texture_colors = _extract_colors_from_3mf_textures(zf)
            extracted_so_far = list(found)
            for hx in texture_colors:
                try:
                    t_lab = rgb_to_lab(hex_to_rgb(hx))
                    if all(delta_e(t_lab, rgb_to_lab(hex_to_rgb(ex))) >= 3.0
                           for ex in extracted_so_far):
                        found.add(hx)
                        extracted_so_far.append(hx)
                except Exception:
                    found.add(hx)
                    extracted_so_far.append(hx)
    except Exception as e:
        return [], str(e)
    trivial = {"#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF",
               "#AAAAAA", "#333333", "#CCCCCC"}
    meaningful = [c for c in found if c not in trivial]
    return (meaningful if meaningful else list(found)), None


# ═══════════════════════════════════════════════════════════════════════════════
#  FILAMENT SEARCH DIALOG
# ═══════════════════════════════════════════════════════════════════════════════

class FilamentSearchDialog(QDialog):
    filament_selected = Signal(int, dict)  # slot_idx, filament_dict

    def __init__(self, slot_idx, library, parent=None, slot_hex=None, lang="de"):
        super().__init__(parent)
        self.slot_idx = slot_idx
        self.slot_hex = slot_hex  # pre-fill color from slot
        self.lang = lang
        # deep-copy so orca imports don't mutate parent library
        self.library = {k: list(v) for k, v in library.items()}
        self._filter_hex = None   # active color filter
        self.setWindowTitle(STRINGS[self.lang].get("search_window_title", "Filament Search — Slot T{i}").format(i=slot_idx + 1))
        self.resize(500, 560)
        self._build_ui()
        if slot_hex:
            self._set_color_filter(slot_hex)

    def _t(self, key, **kwargs):
        s = STRINGS[self.lang].get(key, STRINGS["de"].get(key, key))
        return s.format(**kwargs) if kwargs else s

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(5)

        # ── Color filter row ──────────────────────────────────────────────
        color_row = QHBoxLayout()
        color_lbl = QLabel(self._t("search_color_filter"))
        color_row.addWidget(color_lbl)

        self._color_swatch = QPushButton()
        self._color_swatch.setFixedSize(28, 28)
        self._color_swatch.setToolTip(self._t("search_color_tip"))
        self._color_swatch.clicked.connect(self._pick_color)
        color_row.addWidget(self._color_swatch)

        self._color_hex_edit = QLineEdit()
        self._color_hex_edit.setPlaceholderText(self._t("search_color_placeholder"))
        self._color_hex_edit.setMaximumWidth(120)
        self._color_hex_edit.textChanged.connect(self._on_color_hex_changed)
        color_row.addWidget(self._color_hex_edit)

        slot_btn = QPushButton(self._t("search_slot_color"))
        slot_btn.setToolTip(self._t("search_slot_color_tip"))
        slot_btn.clicked.connect(self._use_slot_color)
        color_row.addWidget(slot_btn)

        clear_btn = QPushButton("✕")
        clear_btn.setFixedWidth(28)
        clear_btn.setToolTip(self._t("search_clear_tip"))
        clear_btn.clicked.connect(self._clear_color_filter)
        color_row.addWidget(clear_btn)
        color_row.addStretch()
        layout.addLayout(color_row)

        # ── Text search row ───────────────────────────────────────────────
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText(self._t("search_placeholder"))
        layout.addWidget(self.search_edit)

        # ── Brand filter ──────────────────────────────────────────────────
        brand_row = QHBoxLayout()
        brand_row.addWidget(QLabel(self._t("lib_brand")))
        self.brand_combo = QComboBox()
        self.brand_combo.addItem(self._t("search_all_brands"))
        for brand in sorted(self.library.keys()):
            self.brand_combo.addItem(brand)
        brand_row.addWidget(self.brand_combo, 1)

        orca_btn = QPushButton(self._t("search_orca_import"))
        orca_btn.setToolTip(self._t("search_orca_tip"))
        orca_btn.clicked.connect(self._import_orca_profiles)
        brand_row.addWidget(orca_btn)
        layout.addLayout(brand_row)

        # ── Results list ─────────────────────────────────────────────────
        self.results_list = QListWidget()
        self.results_list.setAlternatingRowColors(True)
        layout.addWidget(self.results_list, 1)

        # ── Status label ─────────────────────────────────────────────────
        self._status_lbl = QLabel("")
        self._status_lbl.setObjectName("hint")
        layout.addWidget(self._status_lbl)

        # ── Buttons ───────────────────────────────────────────────────────
        btn_row = QHBoxLayout()
        self.select_btn = QPushButton(self._t("search_select"))
        self.select_btn.setObjectName("btn_green")
        cancel_btn = QPushButton(self._t("exp_cancel"))
        btn_row.addWidget(self.select_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

        # ── Connections ───────────────────────────────────────────────────
        self._debounce = QTimer()
        self._debounce.setSingleShot(True)
        self._debounce.timeout.connect(self._update_results)
        self.search_edit.textChanged.connect(lambda: self._debounce.start(200))
        self.brand_combo.currentIndexChanged.connect(self._update_results)
        self.results_list.itemDoubleClicked.connect(lambda _: self._on_select())
        self.select_btn.clicked.connect(self._on_select)
        cancel_btn.clicked.connect(self.reject)

        self._refresh_swatch()
        self._update_results()

    # ── Color filter helpers ─────────────────────────────────────────────

    def _refresh_swatch(self):
        if self._filter_hex:
            self._color_swatch.setStyleSheet(
                f"background:{self._filter_hex}; border:2px solid #888; border-radius:4px;")
        else:
            self._color_swatch.setStyleSheet(
                "background:#444; border:2px solid #888; border-radius:4px;")

    def _pick_color(self):
        initial = QColor(self._filter_hex) if self._filter_hex else QColor("#808080")
        col = QColorDialog.getColor(initial, self, "Zielfarbe wählen")
        if col.isValid():
            self._set_color_filter(col.name().upper())

    def _set_color_filter(self, hex_color):
        self._filter_hex = hex_color
        self._color_hex_edit.blockSignals(True)
        self._color_hex_edit.setText(hex_color)
        self._color_hex_edit.blockSignals(False)
        self._refresh_swatch()
        self._update_results()

    def _on_color_hex_changed(self, text):
        text = text.strip()
        if not text.startswith("#"):
            text = "#" + text
        if len(text) == 7:
            try:
                QColor(text)
                self._filter_hex = text.upper()
                self._refresh_swatch()
                self._debounce.start(300)
                return
            except Exception:
                pass
        if not text or text == "#":
            self._filter_hex = None
            self._refresh_swatch()
            self._debounce.start(300)

    def _clear_color_filter(self):
        self._filter_hex = None
        self._color_hex_edit.blockSignals(True)
        self._color_hex_edit.clear()
        self._color_hex_edit.blockSignals(False)
        self._refresh_swatch()
        self._update_results()

    def _use_slot_color(self):
        if self.slot_hex:
            self._set_color_filter(self.slot_hex)
        else:
            QMessageBox.information(self, "Info", "Kein Slot-Farbwert verfügbar.")

    # ── OrcaSlicer local import ──────────────────────────────────────────

    def _find_orca_dirs(self):
        import platform
        found = []
        seen = set()
        def add(path):
            norm = os.path.normcase(os.path.normpath(path))
            if norm not in seen and os.path.isdir(path):
                seen.add(norm)
                found.append(path)
        if platform.system() == "Windows":
            appdata = os.environ.get("APPDATA", "")
            if appdata and os.path.isdir(appdata):
                for entry in os.scandir(appdata):
                    if not entry.is_dir(): continue
                    if any(x in entry.name.lower() for x in ["orca", "snapmaker_orca", "bambu"]):
                        for sub in ["user/default/filament", "user/filament"]:
                            add(os.path.join(entry.path, sub))
        elif platform.system() == "Darwin":
            base = os.path.expanduser("~/Library/Application Support")
            if os.path.isdir(base):
                for entry in os.scandir(base):
                    if entry.is_dir() and "orca" in entry.name.lower():
                        for sub in ["user/default/filament", "user/filament"]:
                            add(os.path.join(entry.path, sub))
        else:
            base = os.path.expanduser("~/.config")
            if os.path.isdir(base):
                for entry in os.scandir(base):
                    if entry.is_dir() and "orca" in entry.name.lower():
                        add(os.path.join(entry.path, "user", "default", "filament"))
        return found

    def _import_orca_profiles(self):
        dirs = self._find_orca_dirs()
        if not dirs:
            QMessageBox.information(self, "OrcaSlicer", "Keine OrcaSlicer-Installation gefunden.")
            return
        brand = "OrcaSlicer (lokal)"
        imported = []
        for d in dirs:
            for fname in os.listdir(d):
                if not fname.endswith(".json"):
                    continue
                try:
                    with open(os.path.join(d, fname), "r", encoding="utf-8") as f:
                        data = json.load(f)
                    name = data.get("filament_settings_id") or data.get("name") or fname[:-5]
                    if isinstance(name, list): name = name[0]
                    colors = data.get("filament_colour", ["#808080"])
                    if isinstance(colors, list) and colors:
                        hex_c = colors[0].strip()
                    else:
                        hex_c = "#808080"
                    if not hex_c.startswith("#"):
                        hex_c = "#" + hex_c
                    td_raw = data.get("filament_density", [None])
                    if isinstance(td_raw, list): td_raw = td_raw[0]
                    try:
                        td = float(td_raw) if td_raw and td_raw != "nil" else 1.24
                    except Exception:
                        td = 1.24
                    imported.append({"name": name, "hex": hex_c, "td": td})
                except Exception:
                    continue
        if not imported:
            QMessageBox.information(self, "OrcaSlicer", "Keine Profile gefunden.")
            return
        self.library[brand] = imported
        # add to brand combo if not present
        existing = [self.brand_combo.itemText(i) for i in range(self.brand_combo.count())]
        if brand not in existing:
            self.brand_combo.addItem(brand)
        self._status_lbl.setText(f"✅ {len(imported)} OrcaSlicer-Profile geladen")
        self._update_results()

    # ── Results ──────────────────────────────────────────────────────────

    def _update_results(self):
        self.results_list.clear()
        q = self.search_edit.text().lower().strip()
        brand_filter = self.brand_combo.currentText()

        # Compute target lab if color filter active
        target_lab = None
        if self._filter_hex:
            try:
                target_lab = rgb_to_lab(hex_to_rgb(self._filter_hex))
            except Exception:
                target_lab = None

        # Collect matching candidates
        candidates = []
        for brand, filaments in self.library.items():
            if brand_filter != "Alle Brands" and brand != brand_filter:
                continue
            for fil in filaments:
                name = fil.get("name", "")
                if q and q not in brand.lower() and q not in name.lower():
                    continue
                de = None
                if target_lab:
                    try:
                        fil_lab = rgb_to_lab(hex_to_rgb(fil.get("hex", "#808080")))
                        de = delta_e(target_lab, fil_lab)
                    except Exception:
                        de = 9999.0
                candidates.append((de, brand, fil))

        # Sort: by ΔE if color filter active, else by brand+name
        if target_lab:
            candidates.sort(key=lambda x: (x[0] if x[0] is not None else 9999))
        else:
            candidates.sort(key=lambda x: (x[1].lower(), x[2].get("name", "").lower()))

        for i, (de, brand, fil) in enumerate(candidates[:300]):
            name = fil.get("name", "")
            hex_c = fil.get("hex", "#808080")
            if de is not None:
                display = f"ΔE {de:5.1f}  {brand} — {name}  ({hex_c})"
            else:
                display = f"{brand} — {name}  ({hex_c})"
            item = QListWidgetItem(display)
            item.setData(Qt.UserRole, {"brand": brand, **fil})
            try:
                r, g, b = hex_to_rgb(hex_c)
                item.setBackground(QColor(r, g, b))
                lum = 0.299 * r + 0.587 * g + 0.114 * b
                item.setForeground(QColor("#111111" if lum > 128 else "#eeeeee"))
            except Exception:
                pass
            self.results_list.addItem(item)

        total = len(candidates)
        shown = min(300, total)
        if target_lab:
            self._status_lbl.setText(f"{shown}/{total} Filamente — sortiert nach ΔE (CIEDE2000)")
        else:
            self._status_lbl.setText(f"{shown}/{total} Filamente")

    def _on_select(self):
        item = self.results_list.currentItem()
        if not item:
            return
        data = item.data(Qt.UserRole)
        self.filament_selected.emit(self.slot_idx, data)
        self.accept()


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION CLASS
# ═══════════════════════════════════════════════════════════════════════════════

class U1App(QMainWindow):

    def __init__(self):
        super().__init__()
        self._settings = QSettings("U1FullSpectrum", "U1AppPySide6")
        self.lang = self._settings.value("lang", "de")
        self._theme = self._settings.value("theme", "dark")
        self._max_virtual = int(self._settings.value("max_virtual", MAX_VIRTUAL))
        self._virtual = []
        self._history = []
        self._undo_stack = []
        self._slot_undo_stack = []
        self._slot_expanded = [True, False, False, False]
        self._search_wins = {}
        self._last_result = {}
        self._last_sim_hex = None
        self._last_de = None
        self._target_hex = None
        self._3mf_wizard = None  # singleton for 3MF Wizard
        self._fs_export_dlg = None  # singleton for FS 3MF Export dialog
        self._gamut_timer = QTimer()
        self._gamut_timer.setSingleShot(True)
        self._gamut_timer.timeout.connect(self._run_gamut_update)

        # DB files
        self._base_dir = os.path.dirname(os.path.abspath(__file__))
        self.db_file = os.path.join(self._base_dir, "filament_db.json")
        self.preset_file = os.path.join(self._base_dir, "presets.json")

        self._load_db()
        self._load_presets()

        self._build_ui()
        self._apply_theme(self._theme)
        self.setAcceptDrops(True)

        # Restore geometry
        geom = self._settings.value("geometry")
        if geom:
            try:
                self.restoreGeometry(geom)
            except Exception:
                self.resize(1420, 900)
        else:
            self.resize(1420, 900)

        # Restore last color
        last_color = self._settings.value("last_color", "")
        if last_color and len(last_color.lstrip("#")) == 6:
            QTimer.singleShot(200, lambda: self._apply_target(last_color))

        # Restore color model
        cm = self._settings.value("color_model", "linear")
        idx_cm = {"linear": 0, "td": 1, "subtractive": 2}.get(cm, 0)
        if hasattr(self, "_model_combo"):
            self._model_combo.setCurrentIndex(idx_cm)

        # Restore tab
        last_tab = self._settings.value("last_tab", 0, int)
        QTimer.singleShot(300, lambda: self._tabs.setCurrentIndex(last_tab))

        # Gamut strip initial update
        QTimer.singleShot(500, self._update_gamut_strip)

        # Keyboard shortcuts
        QShortcut(QKeySequence("Ctrl+Z"), self, self._undo_last)
        QShortcut(QKeySequence("Return"), self, self._calc)
        QShortcut(QKeySequence("Ctrl+S"), self, self._save_project)
        QShortcut(QKeySequence("Ctrl+O"), self, self._load_project)
        QShortcut(QKeySequence("Delete"), self, self._delete_selected_virtual)

    # ── DRAG & DROP ──────────────────────────────────────────────────────────

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(u.toLocalFile().lower().endswith((".3mf", ".obj")) for u in urls):
                event.acceptProposedAction()
                return
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".3mf"):
                self._open_3mf_with_path(path)
                break
            elif path.lower().endswith(".obj"):
                self._open_obj_assistant(path)
                break
        event.acceptProposedAction()

    def closeEvent(self, event):
        self._save_settings()
        self._settings.setValue("geometry", self.saveGeometry())
        event.accept()

    def _save_settings(self):
        self._settings.setValue("lang", self.lang)
        self._settings.setValue("theme", self._theme)
        self._settings.setValue("max_virtual", self._max_virtual)
        self._settings.setValue("geometry", self.saveGeometry())
        if hasattr(self, "_target_hex") and self._target_hex:
            self._settings.setValue("last_color", self._target_hex)
        if hasattr(self, "_model_combo"):
            models = ["linear", "td", "subtractive"]
            idx = self._model_combo.currentIndex()
            self._settings.setValue("color_model", models[idx] if 0 <= idx < 3 else "linear")
        if hasattr(self, "_tabs"):
            self._settings.setValue("last_tab", self._tabs.currentIndex())
        lh = self._lh_spin.value() if hasattr(self, "_lh_spin") else 0.08
        self._settings.setValue("layer_height", lh)

    def t(self, key, **kwargs):
        s = STRINGS[self.lang].get(key, STRINGS["de"].get(key, key))
        return s.format(**kwargs) if kwargs else s

    # ── STATUS BAR ─────────────────────────────────────────────────────────────

    def _set_status(self, msg, duration=0):
        self.statusBar().showMessage(msg, duration)

    # ── SLOT UNDO ──────────────────────────────────────────────────────────────

    def _save_slot_snapshot(self):
        snapshot = []
        for s in getattr(self, "_slots", []):
            snapshot.append({
                "brand": s["brand_combo"].currentText(),
                "fil": s["fil_combo"].currentText(),
                "hex": s["hex_edit"].text(),
                "td": s["td_spin"].value(),
            })
        self._slot_undo_stack.append(snapshot)
        if len(self._slot_undo_stack) > 10:
            self._slot_undo_stack.pop(0)

    def _undo_slot(self):
        if not self._slot_undo_stack:
            return
        snapshot = self._slot_undo_stack.pop()
        for i, data in enumerate(snapshot):
            if i >= len(self._slots):
                break
            s = self._slots[i]
            brands = [s["brand_combo"].itemText(j) for j in range(s["brand_combo"].count())]
            if data["brand"] in brands:
                s["brand_combo"].setCurrentText(data["brand"])
                self._update_filament_combo(i)
            for j in range(s["fil_combo"].count()):
                if s["fil_combo"].itemText(j) == data["fil"]:
                    s["fil_combo"].setCurrentIndex(j)
                    break
            s["hex_edit"].setText(data["hex"])
            s["td_spin"].setValue(data["td"])
            self._update_slot_preview(i)
        self._update_gamut_strip()

    # ── SLOT COLOR STRIP ───────────────────────────────────────────────────────

    def _update_slot_strip(self, idx):
        s = self._slots[idx]
        hex_val = s["hex_edit"].text().strip()
        if not hex_val.startswith("#"):
            hex_val = "#" + hex_val
        strip = s.get("color_strip")
        if strip:
            try:
                r, g, b = hex_to_rgb(hex_val)
                strip.setStyleSheet(
                    f"background-color: #{r:02X}{g:02X}{b:02X}; "
                    f"border-radius: 2px; border: none;")
            except Exception:
                pass

    # ── AUTO TOGGLE ────────────────────────────────────────────────────────────

    def _on_auto_toggle(self, checked):
        if hasattr(self, "_len_spin"):
            self._len_spin.setVisible(not checked)
        if hasattr(self, "_auto_found_label"):
            self._auto_found_label.setVisible(checked)

    # ── RESULT BORDER ──────────────────────────────────────────────────────────

    def _set_result_border(self, de):
        if de < 3.0:
            color = "#16a34a"
        elif de < 6.0:
            color = "#d97706"
        else:
            color = "#dc2626"
            QTimer.singleShot(0, lambda: self._pulse_result(3))
        if hasattr(self, "_result_frame"):
            self._result_frame.setStyleSheet(
                f"QFrame {{ border: 2px solid {color}; border-radius: 6px; }}")

    def _pulse_result(self, n):
        if n <= 0 or not hasattr(self, "_result_frame"):
            return
        cur = self._result_frame.styleSheet()
        if "4px" in cur:
            new_w = "2px"
        else:
            new_w = "4px"
        import re as _re
        new_style = _re.sub(r'\d+px solid', f'{new_w} solid', cur)
        self._result_frame.setStyleSheet(new_style)
        QTimer.singleShot(500, lambda: self._pulse_result(n - 1))

    # ── VIRTUAL SORT / FILTER ──────────────────────────────────────────────────

    def _get_sorted_filtered_virtual(self):
        fils = list(self._virtual)
        ftext = self._virt_filter.text().lower().strip() if hasattr(self, "_virt_filter") else ""
        if ftext:
            fils = [v for v in fils if ftext in v.get("label", "").lower()
                    or ftext in v.get("sequence", "").lower()
                    or ftext in v.get("target_hex", "").lower()]
        si = self._virt_sort.currentIndex() if hasattr(self, "_virt_sort") else 0
        if si == 1:
            fils.sort(key=lambda v: v.get("de", 99))
        elif si == 2:
            fils.sort(key=lambda v: v.get("de", 0), reverse=True)
        elif si == 3:
            fils.sort(key=lambda v: v.get("label", ""))
        return fils

    # ── DATABASE ───────────────────────────────────────────────────────────────

    def _load_db(self):
        self.library = copy.deepcopy(DEFAULT_LIBRARY)
        if not os.path.exists(self.db_file):
            return
        try:
            with open(self.db_file, "r", encoding="utf-8") as f:
                saved = json.load(f)
            if not isinstance(saved, dict):
                raise ValueError("Not a JSON object.")
            for brand, fils in saved.items():
                if not isinstance(fils, list):
                    continue
                self.library[brand] = [
                    {**x, "td": safe_td(x.get("td", DEFAULT_TD))}
                    for x in fils
                    if isinstance(x, dict) and "name" in x and "hex" in x
                ]
        except Exception as e:
            QMessageBox.warning(self, self.t("dlg_db_err_title"),
                                self.t("dlg_db_err_msg", e=e))

    def _save_db(self):
        try:
            with open(self.db_file, "w", encoding="utf-8") as f:
                json.dump(self.library, f, indent=2, ensure_ascii=False)
        except IOError as e:
            QMessageBox.critical(self, self.t("dlg_save_err"), str(e))

    def _load_presets(self):
        self.presets = {}
        if not os.path.exists(self.preset_file):
            return
        try:
            with open(self.preset_file, "r", encoding="utf-8") as f:
                self.presets = json.load(f)
        except Exception:
            self.presets = {}

    def _save_presets(self):
        try:
            with open(self.preset_file, "w", encoding="utf-8") as f:
                json.dump(self.presets, f, indent=2, ensure_ascii=False)
        except IOError as e:
            QMessageBox.critical(self, self.t("dlg_error"), str(e))

    # ── THEME ──────────────────────────────────────────────────────────────────

    def _apply_theme(self, theme):
        self._theme = theme
        if theme == "light":
            QApplication.instance().setStyleSheet(LIGHT_QSS)
        else:
            QApplication.instance().setStyleSheet(DARK_QSS)

    def _toggle_theme(self):
        new_theme = "light" if self._theme == "dark" else "dark"
        self._apply_theme(new_theme)
        if hasattr(self, "_theme_btn"):
            self._theme_btn.setText("☀️" if new_theme == "light" else "🌙")

    # ── LANGUAGE ───────────────────────────────────────────────────────────────

    def _toggle_lang(self):
        self._save_settings()
        # Save slot values
        slot_vals = self._get_slot_vals()
        target = self._target_hex
        self.lang = "en" if self.lang == "de" else "de"
        # Rebuild UI
        old_widget = self.centralWidget()
        self._build_ui()
        if old_widget:
            old_widget.deleteLater()
        self._apply_theme(self._theme)
        self._restore_slot_vals(slot_vals)
        if target:
            QTimer.singleShot(100, lambda: self._apply_target(target))
        QTimer.singleShot(400, self._update_gamut_strip)

    def _get_slot_vals(self):
        result = []
        for s in getattr(self, "_slots", []):
            result.append({
                "brand": s["brand_combo"].currentText(),
                "hex": s["hex_edit"].text(),
                "td": s["td_spin"].value(),
            })
        return result

    def _restore_slot_vals(self, vals):
        for i, v in enumerate(vals):
            if i >= len(self._slots):
                break
            s = self._slots[i]
            brands = [s["brand_combo"].itemText(j)
                      for j in range(s["brand_combo"].count())]
            if v["brand"] in brands:
                s["brand_combo"].setCurrentText(v["brand"])
                self._update_filament_combo(i)
            if v["hex"]:
                s["hex_edit"].setText(v["hex"])
                self._update_slot_preview(i)
            s["td_spin"].setValue(v["td"])

    # ── BUILD UI ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.setWindowTitle(self.t("app_title"))
        self.statusBar().setStyleSheet(
            "QStatusBar { background: #0a1628; color: #64748b; font-size: 10px; }")
        self.statusBar().showMessage(self.t("status_ready"))

        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(8, 8, 8, 8)
        root_layout.setSpacing(6)

        # Sidebar
        sidebar_scroll = QScrollArea()
        sidebar_scroll.setWidgetResizable(True)
        sidebar_scroll.setMinimumWidth(360)
        sidebar_scroll.setMaximumWidth(420)
        sidebar_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        sidebar_inner = QWidget()
        sidebar_layout = QVBoxLayout(sidebar_inner)
        sidebar_layout.setContentsMargins(6, 6, 6, 6)
        sidebar_layout.setSpacing(6)
        sidebar_scroll.setWidget(sidebar_inner)
        root_layout.addWidget(sidebar_scroll)

        self._build_sidebar(sidebar_layout)
        sidebar_layout.addStretch()

        # Main tabs
        self._tabs = QTabWidget()
        root_layout.addWidget(self._tabs, 1)

        self._build_tab_calculator()
        self._build_tab_virtual()
        self._build_tab_tools()

    def _build_sidebar(self, layout):
        # Top button row: lang + theme
        top_row = QHBoxLayout()
        self._lang_btn = QPushButton(self.t("lang_btn"))
        self._lang_btn.setFixedHeight(28)
        self._lang_btn.clicked.connect(self._toggle_lang)
        top_row.addWidget(self._lang_btn)

        self._theme_btn = QPushButton("🌙" if self._theme == "dark" else "☀️")
        self._theme_btn.setFixedHeight(28)
        self._theme_btn.setFixedWidth(36)
        self._theme_btn.clicked.connect(self._toggle_theme)
        top_row.addWidget(self._theme_btn)
        layout.addLayout(top_row)

        # Title
        title_lbl = QLabel(self.t("phys_heads_title"))
        title_lbl.setObjectName("section_title")
        title_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_lbl)

        desc_lbl = QLabel(self.t("phys_heads_desc"))
        desc_lbl.setObjectName("hint")
        desc_lbl.setWordWrap(True)
        desc_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(desc_lbl)

        # Slot frames
        self._slots = []
        for i in range(4):
            slot_frame = self._build_slot_frame(i)
            layout.addWidget(slot_frame)

        # Preset section
        preset_gb = QGroupBox(self.t("slot_presets"))
        preset_layout = QVBoxLayout(preset_gb)
        self._preset_combo = QComboBox()
        preset_layout.addWidget(self._preset_combo)
        self._refresh_preset_combo()
        preset_btn_row = QHBoxLayout()
        load_preset_btn = QPushButton(self.t("btn_load"))
        save_preset_btn = QPushButton(self.t("btn_save"))
        slot_undo_btn = QPushButton(self.t("slot_undo_btn"))
        slot_undo_btn.setFixedHeight(26)
        slot_undo_btn.setToolTip(self.t("slot_undo_tip"))
        slot_undo_btn.clicked.connect(self._undo_slot)
        load_preset_btn.clicked.connect(self._load_preset)
        save_preset_btn.clicked.connect(self._save_preset)
        preset_btn_row.addWidget(load_preset_btn)
        preset_btn_row.addWidget(save_preset_btn)
        preset_btn_row.addWidget(slot_undo_btn)
        preset_layout.addLayout(preset_btn_row)
        layout.addWidget(preset_gb)

        # Layer height
        lh_gb = QGroupBox(self.t("layer_height_label"))
        lh_layout = QHBoxLayout(lh_gb)
        self._lh_spin = QDoubleSpinBox()
        self._lh_spin.setRange(0.05, 0.50)
        self._lh_spin.setSingleStep(0.01)
        self._lh_spin.setDecimals(2)
        lh_val = float(self._settings.value("layer_height", 0.08))
        self._lh_spin.setValue(lh_val)
        self._lh_spin.setSuffix(" mm")
        self._lh_spin.setToolTip(
            "Optimal 0.08–0.12 mm für FullSpectrum Farbblending.\n"
            "Optimal 0.08–0.12 mm for FullSpectrum color blending.\n"
            "> 0.15 mm → sichtbare Streifen / visible striping.")
        self._lh_spin.valueChanged.connect(self._on_lh_changed)
        lh_layout.addWidget(self._lh_spin)
        self._lh_warn_main = QLabel("")
        self._lh_warn_main.setStyleSheet("color: #fb923c; font-size: 11px;")
        lh_layout.addWidget(self._lh_warn_main)
        # Print height input (for statistics)
        lh_layout.addWidget(QLabel("/" + self.t("lbl_print_height")))
        self._print_h_spin = QDoubleSpinBox()
        self._print_h_spin.setRange(0.0, 500.0)
        self._print_h_spin.setSingleStep(1.0)
        self._print_h_spin.setDecimals(1)
        self._print_h_spin.setValue(0.0)
        self._print_h_spin.setSuffix(" mm")
        self._print_h_spin.setToolTip("Druckhöhe für Statistik (0 = deaktiviert)" if self.lang == "de"
                                       else "Print height for statistics (0 = disabled)")
        self._print_h_spin.valueChanged.connect(self._update_print_stats)
        lh_layout.addWidget(self._print_h_spin)
        lh_hint = QLabel(self.t("lh_hint"))
        lh_hint.setObjectName("hint")
        lh_hint.setWordWrap(True)
        lh_layout.addWidget(lh_hint)
        layout.addWidget(lh_gb)

        # Color model
        model_gb = QGroupBox(self.t("model_label"))
        model_layout = QHBoxLayout(model_gb)
        self._model_combo = QComboBox()
        self._model_combo.addItems([
            self.t("model_linear"),
            self.t("model_td"),
            self.t("model_subtractive"),
            self.t("model_filamentmixer"),
        ])
        self._model_combo.currentIndexChanged.connect(self._on_model_change)
        model_layout.addWidget(self._model_combo)
        layout.addWidget(model_gb)

        # Project save/load
        proj_gb = QGroupBox(self.t("project_groupbox"))
        proj_layout = QHBoxLayout(proj_gb)
        proj_save_btn = QPushButton(self.t("btn_proj_save"))
        proj_load_btn = QPushButton(self.t("btn_proj_load"))
        proj_save_btn.clicked.connect(self._save_project)
        proj_load_btn.clicked.connect(self._load_project)
        proj_layout.addWidget(proj_save_btn)
        proj_layout.addWidget(proj_load_btn)
        layout.addWidget(proj_gb)

    def _build_slot_frame(self, idx):
        # Outer frame with left color strip
        outer_frame = QFrame()
        outer_frame.setObjectName("slot_frame")
        outer_layout = QHBoxLayout(outer_frame)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        # Left color strip (4px wide)
        color_strip = QFrame()
        color_strip.setFixedWidth(4)
        color_strip.setStyleSheet("background-color: #808080; border-radius: 2px; border: none;")
        outer_layout.addWidget(color_strip)

        # Main slot content
        frame = QFrame()
        frame_layout = QVBoxLayout(frame)
        frame_layout.setContentsMargins(6, 4, 6, 6)
        frame_layout.setSpacing(4)
        outer_layout.addWidget(frame, 1)

        # Accordion header button
        is_expanded = self._slot_expanded[idx]
        toggle_btn = QPushButton(
            f"▼ T{idx+1}" if is_expanded else f"▶ T{idx+1}")
        toggle_btn.setFixedHeight(26)
        toggle_btn.setStyleSheet(
            "QPushButton { text-align: left; padding-left: 4px; "
            "font-weight: bold; background-color: transparent; border: none; "
            "color: #94a3b8; } "
            "QPushButton:hover { color: #e2e8f0; }")
        frame_layout.addWidget(toggle_btn)

        # Body widget (collapsible)
        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(4)
        body.setVisible(is_expanded)

        # Header row inside body
        hdr = QHBoxLayout()
        search_btn = QPushButton("🔍")
        search_btn.setFixedSize(26, 26)
        search_btn.clicked.connect(lambda checked, i=idx: self._open_filament_search(i))
        hdr.addStretch()
        hdr.addWidget(search_btn)
        body_layout.addLayout(hdr)

        # Brand combo
        brand_combo = QComboBox()
        brand_combo.addItems(list(self.library.keys()))
        brand_combo.currentTextChanged.connect(lambda t, i=idx: self._update_filament_combo(i))
        body_layout.addWidget(brand_combo)

        # Filament combo
        fil_combo = QComboBox()
        fil_combo.currentTextChanged.connect(lambda t, i=idx: self._apply_filament(i))
        body_layout.addWidget(fil_combo)

        # Hex + preview + TD row
        hex_row = QHBoxLayout()
        hex_edit = QLineEdit()
        hex_edit.setPlaceholderText("#RRGGBB")
        hex_edit.setMaximumWidth(90)
        hex_edit.textChanged.connect(lambda t, i=idx: self._update_slot_preview(i))
        hex_row.addWidget(hex_edit)

        preview_lbl = QLabel()
        preview_lbl.setFixedSize(36, 36)
        preview_lbl.setStyleSheet("background-color: #808080; border-radius: 4px; border: 1px solid #334155;")
        hex_row.addWidget(preview_lbl)

        pick_btn = QPushButton("🎨")
        pick_btn.setFixedSize(26, 26)
        pick_btn.clicked.connect(lambda checked, i=idx: self._pick_slot_color(i))
        hex_row.addWidget(pick_btn)

        hex_row.addStretch()
        td_lbl = QLabel(self.t("lbl_td"))
        hex_row.addWidget(td_lbl)
        td_spin = QDoubleSpinBox()
        td_spin.setRange(0.1, 10.0)
        td_spin.setSingleStep(0.1)
        td_spin.setDecimals(1)
        td_spin.setValue(DEFAULT_TD)
        td_spin.setMaximumWidth(70)
        td_spin.valueChanged.connect(lambda v, i=idx: self._update_gamut_strip())
        hex_row.addWidget(td_spin)

        body_layout.addLayout(hex_row)

        translucent_check = QCheckBox(self.t("transluc_check"))
        translucent_check.setToolTip(self.t("transluc_tip"))
        body_layout.addWidget(translucent_check)

        frame_layout.addWidget(body)

        slot_data = {
            "brand_combo": brand_combo,
            "fil_combo": fil_combo,
            "hex_edit": hex_edit,
            "preview_lbl": preview_lbl,
            "td_spin": td_spin,
            "translucent_check": translucent_check,
            "toggle_btn": toggle_btn,
            "body": body,
            "color_strip": color_strip,
        }
        self._slots.append(slot_data)

        # Wire accordion toggle
        toggle_btn.clicked.connect(lambda checked, i=idx: self._toggle_slot(i))

        # Initialize filament combo
        self._update_filament_combo(idx)
        return outer_frame

    def _toggle_slot(self, i):
        self._slot_expanded[i] = not self._slot_expanded[i]
        s = self._slots[i]
        expanded = self._slot_expanded[i]
        s["body"].setVisible(expanded)
        if expanded:
            s["toggle_btn"].setText(f"▼ T{i+1}")
        else:
            hex_val = s["hex_edit"].text().strip() or "#808080"
            if not hex_val.startswith("#"):
                hex_val = "#" + hex_val
            s["toggle_btn"].setText(f"▶ T{i+1}  {hex_val.upper()}")

    def _update_filament_combo(self, idx):
        s = self._slots[idx]
        brand = s["brand_combo"].currentText()
        fils = self.library.get(brand, [])
        s["fil_combo"].blockSignals(True)
        s["fil_combo"].clear()
        names = [f["name"] for f in fils] or [self.t("empty_slot")]
        s["fil_combo"].addItems(names)
        s["fil_combo"].blockSignals(False)
        if fils:
            self._apply_filament(idx)

    def _apply_filament(self, idx):
        s = self._slots[idx]
        brand = s["brand_combo"].currentText()
        name = s["fil_combo"].currentText()
        if name in _SLOT_SKIP:
            return
        fil = next((f for f in self.library.get(brand, []) if f["name"] == name), None)
        if fil is None:
            return
        s["hex_edit"].blockSignals(True)
        s["hex_edit"].setText(fil["hex"])
        s["hex_edit"].blockSignals(False)
        s["td_spin"].setValue(fil.get("td", DEFAULT_TD))
        self._update_slot_preview(idx)
        self._update_gamut_strip()

    def _update_slot_preview(self, idx):
        s = self._slots[idx]
        hex_val = s["hex_edit"].text().strip()
        if not hex_val.startswith("#"):
            hex_val = "#" + hex_val
        try:
            r, g, b = hex_to_rgb(hex_val)
            s["preview_lbl"].setStyleSheet(
                f"background-color: #{r:02X}{g:02X}{b:02X}; border-radius: 4px; border: 1px solid #334155;")
        except Exception:
            pass
        self._update_slot_strip(idx)
        # Update toggle button text when collapsed
        if not self._slot_expanded[idx]:
            s["toggle_btn"].setText(f"▶ T{idx+1}  {hex_val.upper()}")
        self._update_gamut_strip()

    def _pick_slot_color(self, idx):
        s = self._slots[idx]
        cur = s["hex_edit"].text().strip() or "#808080"
        try:
            r, g, b = hex_to_rgb(cur)
            initial = QColor(r, g, b)
        except Exception:
            initial = QColor(128, 128, 128)
        color = QColorDialog.getColor(initial, self,
                                      self.t("color_picker_title", i=idx + 1))
        if color.isValid():
            hex_str = color.name().upper()
            s["hex_edit"].setText(hex_str)
            self._update_slot_preview(idx)

    def _slot_filaments(self):
        result = []
        for i, s in enumerate(self._slots):
            hex_val = s["hex_edit"].text().strip() or "#808080"
            if not hex_val.startswith("#"):
                hex_val = "#" + hex_val
            td = s["td_spin"].value()
            brand = s["brand_combo"].currentText()
            name = s["fil_combo"].currentText()
            result.append({
                "id": i + 1,
                "hex": hex_val,
                "td": td,
                "lab": rgb_to_lab(hex_to_rgb(hex_val)),
                "brand": brand,
                "name": name,
            })
        return result

    def _apply_slot(self, slot_idx, entry):
        """Apply a filament entry dict to a slot."""
        s = self._slots[slot_idx]
        brand = entry.get("brand", "")
        brands = [s["brand_combo"].itemText(j) for j in range(s["brand_combo"].count())]
        if brand in brands:
            s["brand_combo"].setCurrentText(brand)
            self._update_filament_combo(slot_idx)
            name = entry.get("filament", entry.get("name", ""))
            for j in range(s["fil_combo"].count()):
                if s["fil_combo"].itemText(j) == name:
                    s["fil_combo"].setCurrentIndex(j)
                    return
        hex_val = entry.get("hex", "#808080")
        s["hex_edit"].setText(hex_val)
        s["td_spin"].setValue(entry.get("td", DEFAULT_TD))
        self._update_slot_preview(slot_idx)

    # ── PRESET MANAGEMENT ─────────────────────────────────────────────────────

    def _refresh_preset_combo(self):
        self._preset_combo.blockSignals(True)
        self._preset_combo.clear()
        builtin = [BUILTIN_PRESET_LABELS[k][self.lang] for k in BUILTIN_PRESETS]
        self._preset_combo.addItems(builtin + list(self.presets.keys()))
        self._preset_combo.blockSignals(False)

    def _load_preset(self):
        name = self._preset_combo.currentText()
        if not name:
            return
        entries = self.presets.get(name)
        if entries is None:
            for key, labels in BUILTIN_PRESET_LABELS.items():
                if name in labels.values():
                    entries = BUILTIN_PRESETS.get(key)
                    break
        if entries is None:
            return
        self._save_slot_snapshot()
        for i, entry in enumerate(entries[:4]):
            self._apply_slot(i, entry)
        self._update_gamut_strip()

    def _save_preset(self):
        name, ok = QInputDialog.getText(self, self.t("inp_preset_title"),
                                        self.t("inp_preset_name"))
        if not ok or not name:
            return
        name = name.strip()
        if name.startswith("★"):
            QMessageBox.warning(self, self.t("dlg_saved"),
                                "Names starting with ★ are reserved for built-in presets.")
            return
        self.presets[name] = [
            {
                "brand": s["brand_combo"].currentText(),
                "filament": s["fil_combo"].currentText(),
                "hex": s["hex_edit"].text(),
                "td": s["td_spin"].value(),
            }
            for s in self._slots
        ]
        self._save_presets()
        self._refresh_preset_combo()
        self._preset_combo.setCurrentText(name)
        QMessageBox.information(self, self.t("dlg_saved"),
                                self.t("dlg_preset_saved", name=name))

    # ── PROJECT SAVE/LOAD ─────────────────────────────────────────────────────

    def _save_project(self):
        path, _ = QFileDialog.getSaveFileName(
            self, self.t("btn_proj_save"), "",
            f"{self.t('proj_filetypes')} (*.u1proj);;JSON (*.json)")
        if not path:
            return
        project = {
            "version": 1,
            "layer_height": self._lh_spin.value(),
            "slots": [
                {
                    "brand": s["brand_combo"].currentText(),
                    "color": s["fil_combo"].currentText(),
                    "hex": s["hex_edit"].text(),
                    "td": s["td_spin"].value(),
                }
                for s in self._slots
            ],
            "virtual_fils": self._virtual,
        }
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(project, f, indent=2, ensure_ascii=False)
            QMessageBox.information(self, self.t("dlg_saved"),
                                    self.t("proj_saved", path=path))
        except IOError as e:
            QMessageBox.critical(self, self.t("dlg_error"), str(e))

    def _load_project(self):
        path, _ = QFileDialog.getOpenFileName(
            self, self.t("btn_proj_load"), "",
            f"{self.t('proj_filetypes')} (*.u1proj);;JSON (*.json)")
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as f:
                project = json.load(f)
        except Exception as e:
            QMessageBox.critical(self, self.t("dlg_error"),
                                 self.t("proj_err", e=e))
            return
        if "layer_height" in project:
            self._lh_spin.setValue(float(project["layer_height"]))
        self._save_slot_snapshot()
        for i, s_data in enumerate(project.get("slots", [])[:4]):
            self._apply_slot(i, s_data)
        self._virtual = project.get("virtual_fils", [])
        # Ensure stable_id on loaded virtual heads
        for i, vf in enumerate(self._virtual):
            if "stable_id" not in vf:
                vf["stable_id"] = vf.get("vid", 5 + i)
        self._refresh_virtual_grid()
        self._update_gamut_strip()
        QMessageBox.information(self, self.t("dlg_saved"),
                                self.t("proj_loaded", path=path))

    # ── MISC HANDLERS ─────────────────────────────────────────────────────────

    def _on_lh_changed(self, val):
        if hasattr(self, "_lh_warn_main"):
            if val > 0.15:
                self._lh_warn_main.setText(self.t("lh_warn_striping"))
            else:
                self._lh_warn_main.setText("")
        self._refresh_virtual_grid()

    def _on_model_change(self, idx):
        if self._target_hex:
            self._calc()
        self._refresh_virtual_grid()

    def _apply_mix_pct(self):
        """Build a 2-filament sequence from a direct percentage value."""
        if not self._target_hex:
            QMessageBox.information(self, self.t("dlg_note"), self.t("dlg_select_color"))
            return
        pct_b = self._mix_pct_spin.value() / 100.0  # fraction for T2
        pct_a = 1.0 - pct_b

        # Find best 2 filaments for this target
        fils = self._slot_filaments()
        t_lab = rgb_to_lab(hex_to_rgb(self._target_hex))
        best_pair = None
        best_de = float("inf")
        for i in range(len(fils)):
            for j in range(i + 1, len(fils)):
                sim = self._simulate_mix(
                    [fils[i]["id"]] * round(pct_a * 10) + [fils[j]["id"]] * round(pct_b * 10),
                    fils)
                de = delta_e(sim, t_lab)
                if de < best_de:
                    best_de = de
                    best_pair = (fils[i], fils[j])

        if best_pair is None:
            return

        # Build balanced integer sequence (minority=1)
        fa, fb = best_pair
        if pct_a <= pct_b:
            minority_id, majority_id = fa["id"], fb["id"]
            ratio = round(pct_b / pct_a) if pct_a > 0 else 8
        else:
            minority_id, majority_id = fb["id"], fa["id"]
            ratio = round(pct_a / pct_b) if pct_b > 0 else 8
        ratio = max(1, min(ratio, 24))
        seq = [str(minority_id)] + [str(majority_id)] * ratio

        # Trigger full calc with this sequence as target
        self._auto_check.setChecked(False)
        self._len_spin.setValue(len(seq))
        result = self._calc_for_color(
            self._target_hex, optimizer=False, seq_len=len(seq), auto=False)
        if result:
            self._seq_label.setText(result["sequence"])
            self._last_sim_hex = result["sim_hex"]
            self._last_de = result["de"]
            self._last_result = result
            self._set_status(
                f"Mix {int(self._mix_pct_spin.value())}% T{fb['id']} → "
                f"Seq {''.join(seq)} — ΔE {result['de']:.1f}", 5000)

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 1 — CALCULATOR
    # ═══════════════════════════════════════════════════════════════════════════

    def _build_tab_calculator(self):
        tab = QScrollArea()
        tab.setWidgetResizable(True)
        tab.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        inner = QWidget()
        tab.setWidget(inner)
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # Section title
        sec_title = QLabel(self.t("sec1_title"))
        sec_title.setObjectName("section_title")
        layout.addWidget(sec_title)

        # Target color row
        target_frame = QFrame()
        target_frame.setObjectName("card")
        target_layout = QVBoxLayout(target_frame)
        target_layout.setContentsMargins(10, 10, 10, 10)
        target_layout.setSpacing(8)

        row1 = QHBoxLayout()
        pick_btn = QPushButton(self.t("btn_target_color"))
        pick_btn.setObjectName("btn_primary")
        pick_btn.setFixedHeight(42)
        pick_btn.clicked.connect(self._pick_target_color)
        row1.addWidget(pick_btn)

        self._hex_target_edit = QLineEdit()
        self._hex_target_edit.setPlaceholderText(self.t("hex_placeholder"))
        self._hex_target_edit.setFixedHeight(42)
        self._hex_target_edit.returnPressed.connect(self._on_hex_target_enter)
        self._hex_target_edit.textChanged.connect(self._on_hex_live)
        row1.addWidget(self._hex_target_edit, 1)

        self._target_preview = SwatchLabel("#808080", size=42, radius=8)
        row1.addWidget(self._target_preview)

        self._live_de_label = QLabel("")
        self._live_de_label.setObjectName("hint")
        self._live_de_label.setFixedWidth(80)
        row1.addWidget(self._live_de_label)

        rnd_btn = QPushButton(self.t("btn_random"))
        rnd_btn.setFixedHeight(42)
        rnd_btn.clicked.connect(self._pick_random_color)
        row1.addWidget(rnd_btn)

        if _HAS_PIL:
            img_btn = QPushButton(self.t("btn_img_pick"))
            img_btn.setFixedHeight(42)
            img_btn.clicked.connect(self._pick_color_from_image)
            row1.addWidget(img_btn)

        target_layout.addLayout(row1)

        # Color info label
        self._colorinfo_label = QLabel("")
        self._colorinfo_label.setObjectName("hint")
        target_layout.addWidget(self._colorinfo_label)

        # Gamut warning
        self._gamut_warn = QLabel(self.t("gamut_warning"))
        self._gamut_warn.setObjectName("gamut_warn")
        self._gamut_warn.setWordWrap(True)
        self._gamut_warn.hide()
        target_layout.addWidget(self._gamut_warn)

        layout.addWidget(target_frame)

        # Gamut strip
        gamut_frame = QFrame()
        gamut_frame.setObjectName("card")
        gamut_frame_layout = QHBoxLayout(gamut_frame)
        gamut_frame_layout.setContentsMargins(6, 4, 6, 4)
        gamut_lbl = QLabel(self.t("lbl_gamut"))
        gamut_lbl.setObjectName("hint")
        gamut_frame_layout.addWidget(gamut_lbl)
        self._gamut_strip = GamutStrip()
        self._gamut_strip.setFixedHeight(20)
        gamut_frame_layout.addWidget(self._gamut_strip, 1)
        layout.addWidget(gamut_frame)

        # Options row: length + auto + optimizer
        opts_frame = QFrame()
        opts_frame.setObjectName("card")
        opts_layout = QHBoxLayout(opts_frame)
        opts_layout.setContentsMargins(8, 6, 8, 6)

        opts_layout.addWidget(QLabel(self.t("lbl_length")))
        self._len_spin = QSpinBox()
        self._len_spin.setRange(1, 48)
        self._len_spin.setValue(10)
        opts_layout.addWidget(self._len_spin)

        self._auto_found_label = QLabel(self.t("auto_finding"))
        self._auto_found_label.setObjectName("hint")
        opts_layout.addWidget(self._auto_found_label)

        self._auto_check = QCheckBox(self.t("auto_check").replace("\n", " "))
        self._auto_check.setChecked(True)
        self._auto_check.toggled.connect(self._on_auto_toggle)
        opts_layout.addWidget(self._auto_check)
        # Apply initial state: auto is default True
        self._len_spin.setVisible(False)
        self._auto_found_label.setVisible(True)

        opts_layout.addWidget(QLabel(self.t("de_thresh_label")))
        self._auto_thresh_spin = QDoubleSpinBox()
        self._auto_thresh_spin.setRange(0.5, 10.0)
        self._auto_thresh_spin.setSingleStep(0.5)
        self._auto_thresh_spin.setValue(2.0)
        self._auto_thresh_spin.setMaximumWidth(70)
        opts_layout.addWidget(self._auto_thresh_spin)

        self._optimizer_check = QCheckBox(self.t("optimizer_check").replace("\n", " "))
        opts_layout.addWidget(self._optimizer_check)

        self._skin_tone_check = QCheckBox(self.t("skin_tone_check"))
        self._skin_tone_check.setToolTip(self.t("skin_tone_tip"))
        opts_layout.addWidget(self._skin_tone_check)

        opts_layout.addSpacing(8)
        opts_layout.addWidget(QLabel(self.t("lbl_mix_pct")))
        self._mix_pct_spin = QDoubleSpinBox()
        self._mix_pct_spin.setRange(1.0, 99.0)
        self._mix_pct_spin.setSingleStep(5.0)
        self._mix_pct_spin.setValue(50.0)
        self._mix_pct_spin.setDecimals(0)
        self._mix_pct_spin.setMaximumWidth(70)
        self._mix_pct_spin.setToolTip(self.t("mix_pct_tip"))
        opts_layout.addWidget(self._mix_pct_spin)
        mix_btn = QPushButton(self.t("mix_seq_btn"))
        mix_btn.setFixedWidth(55)
        mix_btn.setToolTip(self.t("mix_seq_tip"))
        mix_btn.clicked.connect(self._apply_mix_pct)
        opts_layout.addWidget(mix_btn)

        opts_layout.addStretch()

        # Color model in options
        opts_layout.addWidget(QLabel(self.t("model_label")))
        if not hasattr(self, "_model_combo"):
            self._model_combo = QComboBox()
            self._model_combo.addItems([
                self.t("model_linear"),
                self.t("model_td"),
                self.t("model_subtractive"),
                self.t("model_filamentmixer"),
            ])
            self._model_combo.currentIndexChanged.connect(self._on_model_change)
        opts_layout.addWidget(self._model_combo)
        layout.addWidget(opts_frame)

        # Calculate button
        self._calc_btn = QPushButton(self.t("btn_calculate"))
        self._calc_btn.setObjectName("btn_primary")
        self._calc_btn.setFixedHeight(44)
        self._calc_btn.setStyleSheet(
            "QPushButton { background-color: #2563eb; color: white; font-size: 13px; "
            "font-weight: bold; border-radius: 6px; } "
            "QPushButton:hover { background-color: #1d4ed8; } "
            "QPushButton:disabled { background-color: #334155; }")
        self._calc_btn.setToolTip(self.t("btn_calculate") + " (Enter)")
        self._calc_btn.clicked.connect(self._calc)
        layout.addWidget(self._calc_btn)

        # Result area
        self._result_frame = QFrame()
        self._result_frame.setObjectName("card")
        result_frame = self._result_frame
        result_layout = QHBoxLayout(result_frame)
        result_layout.setContentsMargins(16, 12, 16, 12)
        result_layout.setSpacing(16)

        # Target swatch
        tgt_col = QVBoxLayout()
        tgt_col.addWidget(QLabel(self.t("label_target")))
        self._result_target_swatch = SwatchLabel("#808080", size=64, radius=10)
        tgt_col.addWidget(self._result_target_swatch)
        self._result_target_hex_lbl = QLabel("—")
        self._result_target_hex_lbl.setObjectName("hint")
        tgt_col.addWidget(self._result_target_hex_lbl)
        result_layout.addLayout(tgt_col)

        # ΔE display
        de_col = QVBoxLayout()
        de_col.setAlignment(Qt.AlignCenter)
        self._de_label = QLabel("ΔE —")
        self._de_label.setObjectName("de_label")
        self._de_label.setAlignment(Qt.AlignCenter)
        de_col.addWidget(self._de_label)
        self._de_quality_lbl = QLabel("")
        self._de_quality_lbl.setObjectName("hint")
        self._de_quality_lbl.setAlignment(Qt.AlignCenter)
        de_col.addWidget(self._de_quality_lbl)
        result_layout.addLayout(de_col)

        # Sim swatch
        sim_col = QVBoxLayout()
        sim_col.addWidget(QLabel(self.t("label_simulated")))
        self._result_sim_swatch = SwatchLabel("#808080", size=64, radius=10)
        sim_col.addWidget(self._result_sim_swatch)
        self._result_sim_hex_lbl = QLabel("—")
        self._result_sim_hex_lbl.setObjectName("hint")
        sim_col.addWidget(self._result_sim_hex_lbl)
        result_layout.addLayout(sim_col)

        # Sequence text + copy
        seq_col = QVBoxLayout()
        seq_col.setAlignment(Qt.AlignCenter)
        self._seq_label = ClickableLabel("----------")
        self._seq_label.setFont(QFont("Courier New", 13, QFont.Bold))
        self._seq_label.setStyleSheet("color: #4ade80; letter-spacing: 1px;")
        self._seq_label.setAlignment(Qt.AlignCenter)
        self._seq_label.setWordWrap(True)
        self._seq_label.setToolTip(self.t("seq_label_click_tip"))
        seq_col.addWidget(self._seq_label)
        _copy_hint_lbl = QLabel(self.t("seq_click_copy_hint"))
        _copy_hint_lbl.setObjectName("hint")
        _copy_hint_lbl.setAlignment(Qt.AlignCenter)
        seq_col.addWidget(_copy_hint_lbl)
        self._hint_label = QLabel("")
        self._hint_label.setObjectName("hint")
        self._hint_label.setWordWrap(True)
        seq_col.addWidget(self._hint_label)
        self._stats_label = QLabel("")
        self._stats_label.setObjectName("hint")
        self._stats_label.setWordWrap(True)
        seq_col.addWidget(self._stats_label)
        # Stripe risk label (Change 4)
        self._stripe_label = QLabel("")
        self._stripe_label.setObjectName("hint")
        self._stripe_label.setWordWrap(True)
        seq_col.addWidget(self._stripe_label)
        # Layer schedule row (Change 8)
        self._layer_sched_labels = []
        layer_sched_row = QHBoxLayout()
        layer_sched_row.setSpacing(2)
        for i in range(12):
            lbl = QLabel(f"L{i+1}")
            lbl.setFixedSize(28, 28)
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet("background-color: #1e293b; border-radius: 4px; font-size: 8px;")
            self._layer_sched_labels.append(lbl)
            layer_sched_row.addWidget(lbl)
        layer_sched_row.addStretch()
        seq_col.addLayout(layer_sched_row)
        # Sequence preview bar — full-width colored strip
        self._seq_preview = QLabel()
        self._seq_preview.setFixedHeight(20)
        self._seq_preview.setMinimumWidth(200)
        self._seq_preview.setToolTip(self.t("seq_preview_tip"))
        seq_col.addWidget(self._seq_preview)
        copy_btn = QPushButton(self.t("btn_copy"))
        copy_btn.setFixedHeight(30)
        copy_btn.clicked.connect(self._copy_sequence)
        seq_col.addWidget(copy_btn)
        result_layout.addLayout(seq_col, 1)
        layout.addWidget(result_frame)

        # Add virtual button
        add_virt_btn = QPushButton(self.t("btn_add_virtual").replace("\n", " "))
        add_virt_btn.setObjectName("btn_green")
        add_virt_btn.setFixedHeight(44)
        add_virt_btn.setStyleSheet(
            "QPushButton { background-color: #15803d; color: white; font-size: 13px; "
            "font-weight: bold; border-radius: 6px; } "
            "QPushButton:hover { background-color: #16a34a; } "
            "QPushButton:disabled { background-color: #334155; }")
        add_virt_btn.setToolTip(self.t("btn_add_virtual").replace("\n", " ") + " (Ctrl+Enter)")
        add_virt_btn.clicked.connect(self.add_virtual)
        layout.addWidget(add_virt_btn)
        QShortcut(QKeySequence("Ctrl+Return"), self, self.add_virtual)

        # Top-3 frame (hidden by default)
        self._top3_frame = QFrame()
        self._top3_frame.setObjectName("card")
        self._top3_layout = QHBoxLayout(self._top3_frame)
        self._top3_layout.setContentsMargins(8, 6, 8, 6)
        self._top3_frame.hide()
        layout.addWidget(self._top3_frame)

        # History section
        hist_gb = QGroupBox(self.t("hist_groupbox"))
        hist_layout = QVBoxLayout(hist_gb)
        self._history_scroll = QScrollArea()
        self._history_scroll.setWidgetResizable(True)
        self._history_scroll.setFixedHeight(110)
        self._history_inner = QWidget()
        self._history_inner_layout = QVBoxLayout(self._history_inner)
        self._history_inner_layout.setContentsMargins(2, 2, 2, 2)
        self._history_inner_layout.setSpacing(2)
        self._history_inner_layout.addStretch()
        self._history_scroll.setWidget(self._history_inner)
        hist_layout.addWidget(self._history_scroll)
        layout.addWidget(hist_gb)

        layout.addStretch()
        self._tabs.addTab(tab, "🎨  " + self.t("tab_calculator"))

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 2 — VIRTUAL HEADS
    # ═══════════════════════════════════════════════════════════════════════════

    def _build_tab_virtual(self):
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(8, 8, 8, 8)
        tab_layout.setSpacing(6)

        # Title row
        title_row = QHBoxLayout()
        sec2_title = QLabel(self.t("sec2_title"))
        sec2_title.setObjectName("section_title")
        title_row.addWidget(sec2_title)
        desc = QLabel(self.t("sec2_desc", max_v=self._max_virtual))
        desc.setObjectName("hint")
        title_row.addWidget(desc)
        title_row.addStretch()
        tab_layout.addLayout(title_row)

        # Toolbar 1
        tb1 = QFrame()
        tb1.setObjectName("card")
        tb1_layout = QHBoxLayout(tb1)
        tb1_layout.setContentsMargins(6, 4, 6, 4)
        tb1_layout.setSpacing(4)

        def _mkbtn(text, color, slot, layout=tb1_layout):
            btn = QPushButton(text)
            btn.setFixedHeight(34)
            btn.setStyleSheet(f"background-color: {color};")
            btn.clicked.connect(slot)
            layout.addWidget(btn)
            return btn

        _mkbtn(self.t("btn_3mf"), "#0f4c81", self._open_3mf_assistant)
        _mkbtn(self.t("obj_btn"), "#065f46", self._open_obj_assistant)
        _mkbtn(self.t("wizard_btn"), "#7c3aed", self._open_3mf_wizard)
        _mkbtn(self.t("btn_batch"), "#4338ca", self._open_batch_dialog)
        _mkbtn(self.t("btn_undo"), "#374151", self._undo_last)
        _mkbtn(self.t("btn_del_all"), "#7f1d1d", self._clear_virtual)
        _mkbtn(self.t("btn_recalc_all"), "#065f46", self._recalc_all_virtual)
        tb1_layout.addStretch()
        tab_layout.addWidget(tb1)

        # Toolbar 2
        tb2 = QFrame()
        tb2.setObjectName("card")
        tb2_layout = QHBoxLayout(tb2)
        tb2_layout.setContentsMargins(6, 4, 6, 4)
        tb2_layout.setSpacing(4)

        def _mkbtn2(text, color, slot):
            btn = QPushButton(text)
            btn.setFixedHeight(30)
            btn.setStyleSheet(f"background-color: {color};")
            btn.clicked.connect(slot)
            tb2_layout.addWidget(btn)
            return btn

        _mkbtn2(self.t("btn_export_all"), "#374151", self._open_export_dialog)
        _mkbtn2(self.t("btn_orca_export"), "#0f766e", self._open_orca_export_dialog)
        _mkbtn2(self.t("btn_3mf_write"), "#374151", self._write_3mf_colors)
        _mkbtn2(self.t("btn_copy_all_cad"), "#0e7490", self._open_copy_all_cadence)
        _mkbtn2(self.t("btn_de_overview"), "#1e3a5f", self._open_de_overview)
        _mkbtn2(self.t("btn_recalc_all"), "#065f46", self._recalc_all_virtual)
        _mkbtn2(self.t("fs_export_btn"), "#7c2d12", self._open_fs_export_dialog)
        tb2_layout.addStretch()
        tab_layout.addWidget(tb2)

        # Sort / Filter bar
        sort_filter_bar = QFrame()
        sort_filter_bar.setObjectName("card")
        sf_layout = QHBoxLayout(sort_filter_bar)
        sf_layout.setContentsMargins(6, 4, 6, 4)
        sf_layout.setSpacing(6)
        sf_layout.addWidget(QLabel(self.t("lbl_sort")))
        self._virt_sort = QComboBox()
        self._virt_sort.addItems([
            self.t("sort_added"), self.t("sort_de_asc"),
            self.t("sort_de_desc"), self.t("sort_label_az"),
        ])
        self._virt_sort.setFixedWidth(130)
        self._virt_sort.currentIndexChanged.connect(self._refresh_virtual_grid)
        sf_layout.addWidget(self._virt_sort)
        self._virt_filter = QLineEdit()
        self._virt_filter.setPlaceholderText("Filter…")
        self._virt_filter.textChanged.connect(self._refresh_virtual_grid)
        sf_layout.addWidget(self._virt_filter, 1)
        tab_layout.addWidget(sort_filter_bar)

        # Virtual grid scroll area
        self._vgrid_scroll = QScrollArea()
        self._vgrid_scroll.setWidgetResizable(True)
        self._vgrid_inner = QWidget()
        self._vgrid_layout = QVBoxLayout(self._vgrid_inner)
        self._vgrid_layout.setContentsMargins(4, 4, 4, 4)
        self._vgrid_layout.setSpacing(4)
        self._vgrid_scroll.setWidget(self._vgrid_inner)
        tab_layout.addWidget(self._vgrid_scroll, 1)

        self._refresh_virtual_grid()
        self._tabs.addTab(tab, "🔲  " + self.t("tab_virtual"))

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 3 — TOOLS
    # ═══════════════════════════════════════════════════════════════════════════

    def _build_tab_tools(self):
        tab = QScrollArea()
        tab.setWidgetResizable(True)
        tab.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        inner = QWidget()
        tab.setWidget(inner)
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        def _section(title):
            gb = QGroupBox(title)
            gb_layout = QHBoxLayout(gb)
            gb_layout.setContentsMargins(8, 8, 8, 8)
            gb_layout.setSpacing(6)
            layout.addWidget(gb)
            return gb_layout

        def _btn(parent_layout, text, color, slot):
            b = QPushButton(text)
            b.setFixedHeight(36)
            b.setStyleSheet(f"background-color: {color};")
            b.clicked.connect(slot)
            parent_layout.addWidget(b)
            return b

        # Gamut & Analysis
        f = _section("🔭  " + self.t("tools_analysis"))
        if _HAS_MPL:
            _btn(f, self.t("btn_lab_plot"), "#0f4c81", self._show_lab_plot)
            _btn(f, self.t("btn_gamut_plot"), "#0f4c81", self._open_gamut_plot)
        _btn(f, self.t("btn_swatch"), "#374151", self._save_swatch)
        _btn(f, self.t("btn_slicer_guide"), "#7c3aed", self._open_slicer_guide)
        _btn(f, self.t("btn_tc_est"), "#374151", self._open_tc_estimator)
        _btn(f, self.t("btn_print_stats"), "#374151", self._open_print_stats)
        _btn(f, self.t("btn_layer_preview"), "#0f4c81", self._open_layer_preview)
        _btn(f, self.t("de_matrix_btn"), "#1e3a5f", self._open_filament_matrix)
        if _HAS_PIL:
            _btn(f, self.t("png_export_btn"), "#374151", self._export_png_summary)
        f.addStretch()

        # Color generation
        f2 = _section("🌈  " + self.t("tools_color_gen"))
        _btn(f2, self.t("btn_gradient"), "#0e7490", self._open_gradient_dialog)
        _btn(f2, self.t("btn_harmonies"), "#7c3aed", self._open_harmonies_dialog)
        _btn(f2, self.t("btn_multi_gradient"), "#0e7490", self._open_multi_gradient_dialog)
        if _HAS_PIL:
            _btn(f2, self.t("btn_palette"), "#374151", self._import_palette_from_image)
        _btn(f2, self.t("btn_multitarget"), "#7c3aed", self._open_multitarget_optimizer)
        f2.addStretch()

        # Optimization
        f3 = _section("🎯  " + self.t("tools_optimization"))
        _btn(f3, self.t("btn_slot_opt"), "#7c3aed", self._open_slot_optimizer)
        _btn(f3, self.t("btn_td_cal"), "#0f4c81", self._open_td_calibration)
        f3.addStretch()

        # Library
        f4 = _section("📚  " + self.t("tools_library"))
        _btn(f4, self.t("btn_new_brand"), "#1e3a5f", self._add_brand)
        _btn(f4, self.t("btn_library"), "#374151", self._open_library_manager)
        _btn(f4, self.t("btn_web_update"), "#164e63", self._web_update_library)
        f4.addStretch()

        # Export
        f5 = _section("📤  " + self.t("tools_export"))
        _btn(f5, self.t("txt_json_export_btn"), "#374151", self._open_export_dialog)
        _btn(f5, self.t("btn_orca_export"), "#0f766e", self._open_orca_export_dialog)
        _btn(f5, self.t("btn_recipe"), "#374151", self._open_recipe_export)
        f5.addStretch()

        layout.addStretch()
        self._tabs.addTab(tab, "🛠  " + self.t("tab_tools"))

    # ═══════════════════════════════════════════════════════════════════════════
    # CALCULATION CORE
    # ═══════════════════════════════════════════════════════════════════════════

    def _simulate_mix(self, sequence, fils):
        """Color simulation: linear (additive), TD-weighted, subtractive, or FilamentMixer."""
        by_id = {f["id"]: f for f in fils}
        counts = {}
        for fid in sequence:
            counts[fid] = counts.get(fid, 0) + 1
        total = len(sequence)
        models = ["linear", "td", "subtractive", "filamentmixer"]
        idx = self._model_combo.currentIndex() if hasattr(self, "_model_combo") else 0
        model = models[idx] if 0 <= idx < len(models) else "linear"

        if model == "filamentmixer":
            if not sequence or not fils:
                return rgb_to_lab((128, 128, 128))
            fil_map = {f["id"]: hex_to_rgb(f["hex"]) for f in fils}
            if _HAS_FILAMENT_MIXER:
                # Use GPMixer (dE≈1.79) — matches FullSpectrum slicer preview
                try:
                    colors = [fil_map.get(int(s), (128, 128, 128)) for s in sequence]
                    weights = [1.0 / len(colors)] * len(colors)
                    mixer = _GPMixer()
                    result_rgb = mixer.mix_n_colors(
                        [(r / 255, g / 255, b / 255) for r, g, b in colors], weights)
                    r = int(result_rgb[0] * 255)
                    g = int(result_rgb[1] * 255)
                    b = int(result_rgb[2] * 255)
                    return rgb_to_lab((r, g, b))
                except Exception:
                    pass  # fall through to polynomial approximation
            fids = [int(s) for s in sequence]
            r, g, b = fil_map.get(fids[0], (128, 128, 128))
            for i in range(1, len(fids)):
                r2, g2, b2 = fil_map.get(fids[i], (r, g, b))
                t = 1.0 / (i + 1)
                r, g, b = filament_mixer_lerp(r, g, b, r2, g2, b2, t)
            return rgb_to_lab((r, g, b))

        # Build per-slot translucency map (slot index → bool)
        _slot_translucent = {}
        for _si, _sd in enumerate(getattr(self, "_slots", [])):
            _tc = _sd.get("translucent_check")
            _slot_translucent[_si + 1] = (_tc.isChecked() if _tc is not None else False)

        r_acc = g_acc = b_acc = 0.0
        total_w = 0.0
        for fid, cnt in counts.items():
            r, g, b = hex_to_rgb(by_id[fid]["hex"])
            td = max(0.1, float(by_id[fid].get("td", 5.0)))
            base_w = cnt / total
            # Per-slot: if translucent_check is checked for this fid, use TD weighting
            # regardless of global model; otherwise respect global model
            slot_is_translucent = _slot_translucent.get(int(fid), False)
            if slot_is_translucent:
                w = base_w / td
            elif model == "td":
                w = base_w / td
            else:
                w = base_w
            total_w += w
            rl = (r / 255) ** 2.2
            gl = (g / 255) ** 2.2
            bl = (b / 255) ** 2.2
            if model == "subtractive":
                r_acc += (1 - rl) * w
                g_acc += (1 - gl) * w
                b_acc += (1 - bl) * w
            else:
                r_acc += rl * w
                g_acc += gl * w
                b_acc += bl * w
        if total_w > 0:
            r_acc /= total_w
            g_acc /= total_w
            b_acc /= total_w
        if model == "subtractive":
            r_acc = 1 - r_acc
            g_acc = 1 - g_acc
            b_acc = 1 - b_acc

        def tog(v):
            return round(min(255, max(0, v ** (1 / 2.2) * 255)))

        return rgb_to_lab((tog(r_acc), tog(g_acc), tog(b_acc)))

    def _build_sequence(self, ordered, tot, n):
        """Build n-layer sequence from sorted filaments by weight."""
        counts = [max(0, round((s["w"] / tot) * n)) for s in ordered]
        diff = n - sum(counts)
        i = 0
        while diff > 0:
            counts[i % len(counts)] += 1
            diff -= 1
            i += 1
        while diff < 0:
            k = i % len(counts)
            if counts[k] > 0:
                counts[k] -= 1
                diff += 1
            i += 1
        seq = [None] * n
        pos = n - 1
        for j, s in enumerate(ordered):
            for _ in range(counts[j]):
                if pos >= 0:
                    seq[pos] = s["id"]
                    pos -= 1
        return [seq[k] or ordered[0]["id"] for k in range(n)]

    def _calc_for_color(self, target_hex, optimizer=False, seq_len=None,
                        auto=False, auto_threshold=2.0):
        """Calculate sequence for a color — no UI side effects."""
        t_lab = rgb_to_lab(hex_to_rgb(target_hex))
        fils = self._slot_filaments()
        scores = [
            {
                "id": f["id"],
                "w": (1 / (delta_e(t_lab, f["lab"]) + 0.1)) * (10 / f["td"]),
                "h": f["hex"],
            }
            for f in fils
        ]
        tot = sum(s["w"] for s in scores)
        if tot == 0:
            return None

        all_opt_results = []

        def best_seq_for_n(n):
            if optimizer:
                best, best_dv = None, float("inf")
                for perm in iter_permutations(scores):
                    seq = self._build_sequence(list(perm), tot, n)
                    dv = delta_e(self._simulate_mix(seq, fils), t_lab)
                    seq_str = "".join(map(str, seq))
                    all_opt_results.append({
                        "target_hex": target_hex,
                        "sequence": seq_str,
                        "sim_hex": lab_to_hex(self._simulate_mix(seq, fils)),
                        "de": dv,
                        "seq_len": n,
                    })
                    if dv < best_dv:
                        best_dv = dv
                        best = seq
                return best
            return self._build_sequence(
                sorted(scores, key=lambda x: x["w"], reverse=True), tot, n)

        if auto:
            chosen_seq = None
            for n in range(1, MAX_SEQ_LEN + 1):
                seq = best_seq_for_n(n)
                dv = delta_e(self._simulate_mix(seq, fils), t_lab)
                if dv <= auto_threshold:
                    chosen_seq = seq
                    break
            if chosen_seq is None:
                chosen_seq = best_seq_for_n(MAX_SEQ_LEN)
        else:
            n = seq_len if seq_len is not None else MAX_SEQ_LEN
            chosen_seq = best_seq_for_n(n)

        if optimizer and all_opt_results:
            all_opt_results.sort(key=lambda x: x["de"])
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
        dv = delta_e(sim_lab, t_lab)
        return {
            "target_hex": target_hex,
            "sequence": "".join(map(str, chosen_seq)),
            "sim_hex": lab_to_hex(sim_lab),
            "de": dv,
            "seq_len": len(chosen_seq),
        }

    # ── GAMUT STRIP ───────────────────────────────────────────────────────────

    def _update_gamut_strip(self):
        self._gamut_timer.start(150)

    def _run_gamut_update(self):
        if not hasattr(self, "_gamut_strip"):
            return
        fils = self._slot_filaments()
        if not fils or len(fils) < 2:
            return
        samples = [f["hex"] for f in fils]
        for f1, f2 in itertools.combinations(fils, 2):
            for w in [0.15, 0.3, 0.45, 0.55, 0.7, 0.85]:
                lab1, lab2 = f1["lab"], f2["lab"]
                mix = tuple(lab1[k] * (1 - w) + lab2[k] * w for k in range(3))
                samples.append(lab_to_hex(mix))
        if len(fils) >= 3:
            for combo in itertools.combinations(fils, 3):
                for w in [(0.33, 0.33, 0.34), (0.5, 0.25, 0.25),
                          (0.25, 0.5, 0.25), (0.25, 0.25, 0.5)]:
                    mix = tuple(sum(combo[j]["lab"][k] * w[j] for j in range(3))
                                for k in range(3))
                    samples.append(lab_to_hex(mix))
        if len(fils) >= 4:
            mix4 = tuple(sum(f["lab"][k] for f in fils) / 4 for k in range(3))
            samples.append(lab_to_hex(mix4))

        def hue_key(h):
            try:
                r, g, b = hex_to_rgb(h)
                r_, g_, b_ = r / 255, g / 255, b / 255
                mx, mn = max(r_, g_, b_), min(r_, g_, b_)
                df = mx - mn
                if df == 0:
                    hue = 0
                elif mx == r_:
                    hue = (60 * ((g_ - b_) / df) + 360) % 360
                elif mx == g_:
                    hue = (60 * ((b_ - r_) / df) + 120) % 360
                else:
                    hue = (60 * ((r_ - g_) / df) + 240) % 360
                s = 0 if mx == 0 else (df / mx * 100)
                return (0 if s < 10 else 1, hue)
            except Exception:
                return (0, 0)

        samples.sort(key=hue_key)
        self._gamut_strip.schedule_update(samples)

    # ── CALCULATOR ACTIONS ────────────────────────────────────────────────────

    def _pick_target_color(self):
        initial = QColor(128, 128, 128)
        if self._target_hex:
            try:
                r, g, b = hex_to_rgb(self._target_hex)
                initial = QColor(r, g, b)
            except Exception:
                pass
        color = QColorDialog.getColor(initial, self, self.t("target_picker_title"))
        if color.isValid():
            self._apply_target(color.name().upper())

    def _on_hex_target_enter(self):
        raw = self._hex_target_edit.text().strip()
        if not raw.startswith("#"):
            raw = "#" + raw
        if len(raw) == 7:
            self._apply_target(raw.upper())

    def _on_hex_live(self):
        if not hasattr(self, "_hex_live_timer"):
            self._hex_live_timer = QTimer()
            self._hex_live_timer.setSingleShot(True)
            self._hex_live_timer.timeout.connect(self._run_hex_live)
        self._hex_live_timer.start(300)

    def _run_hex_live(self):
        raw = self._hex_target_edit.text().strip()
        if not raw.startswith("#"):
            raw = "#" + raw
        if len(raw) != 7:
            self._live_de_label.setText("")
            return
        try:
            hex_to_rgb(raw)
        except Exception:
            return
        result = self._calc_for_color(raw, optimizer=False, seq_len=4, auto=False)
        if result and hasattr(self, "_live_de_label"):
            de = result["de"]
            self._live_de_label.setText(f"≈ΔE {de:.1f}")
            self._live_de_label.setStyleSheet(f"color: {_de_color(de)};")
        self._check_auto_suggestion(raw)

    def _apply_target(self, hex_str):
        if not hex_str.startswith("#"):
            hex_str = "#" + hex_str
        self._target_hex = hex_str.upper()
        try:
            r, g, b = hex_to_rgb(hex_str)
            self._hex_target_edit.blockSignals(True)
            self._hex_target_edit.setText(hex_str)
            self._hex_target_edit.blockSignals(False)
            self._target_preview.set_color(hex_str)
            # Color info
            r_, g_, b_ = r / 255, g / 255, b / 255
            mx, mn = max(r_, g_, b_), min(r_, g_, b_)
            df = mx - mn
            if df == 0:
                h = 0
            elif mx == r_:
                h = (60 * ((g_ - b_) / df) + 360) % 360
            elif mx == g_:
                h = (60 * ((b_ - r_) / df) + 120) % 360
            else:
                h = (60 * ((r_ - g_) / df) + 240) % 360
            s_val = 0 if mx == 0 else (df / mx * 100)
            v_val = mx * 100
            lab = rgb_to_lab((r, g, b))
            info = self.t("colorinfo_label", r=r, g=g, b=b,
                          h=h, s=s_val, v=v_val, L=lab[0], a=lab[1], b_=lab[2])
            self._colorinfo_label.setText(info)
        except Exception:
            pass
        self._calc()

    def _calc(self):
        if not self._target_hex:
            QMessageBox.information(self, self.t("dlg_note"), self.t("dlg_select_color"))
            return

        if hasattr(self, "_calc_btn"):
            self._calc_btn.setText("⏳ Berechne…")
            self._calc_btn.setEnabled(False)
            QApplication.processEvents()

        try:
            t_lab = rgb_to_lab(hex_to_rgb(self._target_hex))
            fils = self._slot_filaments()

            # Gamut warning
            if min(delta_e(t_lab, f["lab"]) for f in fils) > GAMUT_WARN_DE:
                self._gamut_warn.show()
            else:
                self._gamut_warn.hide()

            auto = self._auto_check.isChecked()
            threshold = self._auto_thresh_spin.value() if auto else 2.0
            seq_len = self._len_spin.value() if not auto else None

            # Skin-Tone Mode: override threshold if checkbox is active
            skin_tone = hasattr(self, "_skin_tone_check") and self._skin_tone_check.isChecked()
            if skin_tone and auto:
                threshold = min(threshold, 1.5)
                # Check if target is in skin-tone LAB range
                L_, a_, b_ = t_lab
                if 40 <= L_ <= 80 and 5 <= a_ <= 25 and 10 <= b_ <= 30:
                    threshold = 1.0
                    self._set_status(self.t("skin_tone_mode", de=1.0), 8000)
                else:
                    self._set_status(self.t("skin_tone_mode", de=1.5), 5000)

            result = self._calc_for_color(
                self._target_hex,
                self._optimizer_check.isChecked(),
                seq_len=seq_len,
                auto=auto,
                auto_threshold=threshold)
            if result is None:
                return

            seq = result["sequence"]
            if auto:
                if hasattr(self, "_auto_found_label"):
                    self._auto_found_label.setText(
                        self.t("auto_found", n=result["seq_len"]))

            self._seq_label.setText(seq)
            self._last_sim_hex = result["sim_hex"]
            self._last_de = result["de"]
            self._last_sequence = list(result.get("sequence", []))
            dv = result["de"]

            # Update result swatches
            self._result_target_swatch.set_color(self._target_hex)
            self._result_target_hex_lbl.setText(self._target_hex.upper())
            self._result_sim_swatch.set_color(result["sim_hex"])
            self._result_sim_hex_lbl.setText(result["sim_hex"].upper())

            # ΔE display
            self._de_label.setText(f"ΔE {dv:.1f}")
            self._de_label.setStyleSheet(f"color: {_de_color(dv)}; font-size: 22pt; font-weight: bold;")
            quality = ("excellent ✓" if dv < 3.0 else "good" if dv < 6.0 else "visible") \
                if self.lang == "en" else \
                ("ausgezeichnet ✓" if dv < 3.0 else "gut" if dv < 6.0 else "sichtbar")
            self._de_quality_lbl.setText(quality)
            self._de_quality_lbl.setStyleSheet(f"color: {_de_color(dv)};")

            # Result border color based on ΔE
            self._set_result_border(dv)

            # Hint — use run-length encoding for compact display
            lh = self._lh_spin.value()
            n_fils = _seq_filament_count(seq)
            runs = seq_to_runs(seq)
            pat_str = " ".join(f"T{fid}×{cnt}" for fid, cnt in runs)
            if n_fils == 1:
                hint = self.t("hint_pure")
            elif n_fils == 2 and lh > 0:
                cad = calc_cadence(seq, lh)
                ids = sorted(cad.keys())
                hint = self.t("hint_cadence", a=cad[ids[0]],
                              b=cad[ids[1]] if len(ids) > 1 else lh, p=pat_str)
            else:
                hint = self.t("hint_pattern", p=pat_str)
            self._hint_label.setText(hint)

            self._last_result = result
            self._show_top3()
            self._update_history(self._target_hex, result["sim_hex"], dv, seq)
            self._set_status(self.t("status_calculated", de=dv, seq=seq), 5000)

            # Stripe risk check (Change 4)
            _risk, _risk_msg = check_stripe_risk(seq)
            if hasattr(self, "_stripe_label"):
                if _risk:
                    self._stripe_label.setText(_risk_msg)
                    self._stripe_label.setStyleSheet("color: #f97316;")
                else:
                    self._stripe_label.setText(_risk_msg)
                    self._stripe_label.setStyleSheet("color: #4ade80;")

            # Material compatibility check
            seq_ids = [int(c) - 1 for c in seq if c.isdigit()]
            compat_warn = self._check_material_compatibility(seq_ids)
            if compat_warn:
                self._set_status(compat_warn, 10000)

            # Layer schedule (Change 8)
            self._draw_layer_schedule(seq)
            self._draw_seq_preview(seq)
            self._update_print_stats()
        finally:
            if hasattr(self, "_calc_btn"):
                self._calc_btn.setText(self.t("btn_calculate"))
                self._calc_btn.setEnabled(True)

    def _draw_layer_schedule(self, sequence):
        """Update colored squares for first 12 layers showing active filament (Change 8)."""
        if not hasattr(self, "_layer_sched_labels"):
            return
        if not sequence:
            for lbl in self._layer_sched_labels:
                lbl.setStyleSheet("background-color: #1e293b; border-radius: 4px; font-size: 8px;")
                lbl.setText("")
            return
        schedule = compute_layer_schedule(sequence, n_layers=12)
        fils_hex = {f["id"]: f["hex"] for f in self._slot_filaments()}
        for i, lbl in enumerate(self._layer_sched_labels):
            if i < len(schedule):
                fid = schedule[i]
                color = fils_hex.get(fid, "#888888")
                try:
                    r_, g_, b_ = hex_to_rgb(color)
                    lum = 0.299 * r_ + 0.587 * g_ + 0.114 * b_
                    tc = "#111111" if lum > 140 else "#eeeeee"
                except Exception:
                    tc = "#eeeeee"
                lbl.setStyleSheet(
                    f"background-color: {color}; border-radius: 4px; "
                    f"font-size: 8px; color: {tc};")
                lbl.setText(f"L{i+1}")
            else:
                lbl.setStyleSheet("background-color: #1e293b; border-radius: 4px; font-size: 8px;")
                lbl.setText("")

    def _draw_seq_preview(self, sequence):
        """Paint a full-width bar where each cell = one slot in the sequence."""
        if not hasattr(self, "_seq_preview"):
            return
        if not sequence:
            self._seq_preview.setPixmap(QPixmap())
            return
        fils_hex = {f["id"]: f["hex"] for f in self._slot_filaments()}
        n = len(sequence)
        w = max(self._seq_preview.width(), 200)
        h = 20
        cell_w = max(1, w // n)
        total_w = cell_w * n
        img = QImage(total_w, h, QImage.Format_RGB32)
        img.fill(QColor("#1e293b"))
        painter = QPainter(img)
        painter.setPen(Qt.NoPen)
        for i, fid in enumerate(sequence):
            color = fils_hex.get(int(fid), "#888888")
            painter.setBrush(QColor(color))
            painter.drawRect(i * cell_w, 0, cell_w, h)
        painter.end()
        self._seq_preview.setPixmap(QPixmap.fromImage(img))

    def _update_print_stats(self):
        """Recompute and display print statistics."""
        if not hasattr(self, "_stats_label"):
            return
        seq = getattr(self, "_seq_label", None)
        seq_text = seq.text().strip() if seq else ""
        if not seq_text or seq_text == "----------":
            self._stats_label.setText("")
            return
        lh = self._lh_spin.value() if hasattr(self, "_lh_spin") else 0.08
        ph = self._print_h_spin.value() if hasattr(self, "_print_h_spin") else 0.0
        if ph <= 0 or lh <= 0:
            self._stats_label.setText("")
            return
        sequence = list(seq_text)
        total_layers = max(1, round(ph / lh))
        cycle_len = len(sequence)
        # Count tool changes: whenever consecutive layers differ
        full_cycles = total_layers // cycle_len
        remainder = total_layers % cycle_len
        extended = sequence * full_cycles + sequence[:remainder]
        tool_changes = sum(1 for i in range(1, len(extended)) if extended[i] != extended[i-1])
        # layers per filament
        from collections import Counter
        counts = Counter(extended)
        # estimate change time: ~30s per change
        change_time_min = tool_changes * 30 / 60
        parts = [self.t("stats_summary").format(layers=total_layers, changes=tool_changes)]
        for fid, cnt in sorted(counts.items()):
            pct = 100.0 * cnt / total_layers
            parts.append(self.t("stats_filament_row").format(fid=fid, cnt=cnt, pct=pct))
        parts.append(self.t("stats_change_time").format(min=change_time_min))
        self._stats_label.setText("\n".join(parts))

    def _show_top3(self):
        # Clear existing
        while self._top3_layout.count():
            item = self._top3_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        results = getattr(self, "_top3_results", [])
        if len(results) < 2:
            self._top3_frame.hide()
            return

        self._top3_frame.show()
        top3_title = QLabel(self.t("top3_title"))
        top3_title.setObjectName("hint")
        self._top3_layout.addWidget(top3_title)

        for idx, r in enumerate(results[:3]):
            card = QFrame()
            card.setObjectName("card")
            card_layout = QHBoxLayout(card)
            card_layout.setContentsMargins(6, 4, 6, 4)
            card_layout.setSpacing(4)

            rank_lbl = QLabel(self.t("top3_rank", r=idx + 1))
            rank_lbl.setObjectName("hint")
            card_layout.addWidget(rank_lbl)

            seq_lbl = QLabel(" ".join(r["sequence"]))
            seq_lbl.setFont(QFont("Courier New", 10, QFont.Bold))
            seq_lbl.setStyleSheet("color: #4ade80;")
            card_layout.addWidget(seq_lbl)

            de_lbl = QLabel(f"ΔE {r['de']:.1f}")
            de_lbl.setStyleSheet(f"color: {_de_color(r['de'])}; font-size: 9pt;")
            card_layout.addWidget(de_lbl)

            sim_sw = SwatchLabel(r["sim_hex"], size=20, radius=4)
            card_layout.addWidget(sim_sw)

            add_btn = QPushButton(self.t("top3_add"))
            add_btn.setFixedHeight(24)
            add_btn.setFixedWidth(80)
            add_btn.setStyleSheet("background-color: #15803d;")
            add_btn.clicked.connect(lambda checked, res=r: self._add_top3_result(res))
            card_layout.addWidget(add_btn)
            self._top3_layout.addWidget(card)

        self._top3_layout.addStretch()

    def _add_top3_result(self, result):
        self._last_result = result
        self.add_virtual(result)

    def add_virtual(self, result=None):
        if len(self._virtual) >= self._max_virtual:
            QMessageBox.warning(self, self.t("dlg_max_virtual"),
                                self.t("dlg_max_virtual_msg", max_v=self._max_virtual))
            return
        if result is None:
            result = self._last_result
        if not result:
            QMessageBox.information(self, self.t("dlg_note"), self.t("dlg_no_seq"))
            return
        self._undo_stack.append(copy.deepcopy(self._virtual))
        if len(self._undo_stack) > 20:
            self._undo_stack.pop(0)
        vid = 5 + len(self._virtual)
        stable_id = vid  # assign before append so ID == vid
        self._virtual.append({
            "vid": vid,
            "stable_id": stable_id,
            "target_hex": result["target_hex"],
            "sequence": result["sequence"],
            "sim_hex": result["sim_hex"],
            "de": result["de"],
            "label": self.t("virtual_label_default", vid=vid),
        })
        self._refresh_virtual_grid()
        self._set_status(self.t("status_added", vid=vid), 3000)

    def _undo_last(self):
        if not self._undo_stack:
            QMessageBox.information(self, self.t("dlg_note"), self.t("undo_empty"))
            return
        self._virtual = self._undo_stack.pop()
        self._refresh_virtual_grid()

    def _clear_virtual(self):
        if not self._virtual:
            return
        if QMessageBox.question(self, self.t("dlg_del_title"),
                                self.t("dlg_del_virtual"),
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self._undo_stack.append(copy.deepcopy(self._virtual))
            self._virtual.clear()
            self._refresh_virtual_grid()

    def _copy_sequence(self):
        seq = self._seq_label.text().strip()
        if seq and seq != "----------":
            QApplication.clipboard().setText(seq)
            QMessageBox.information(self, self.t("dlg_note"), self.t("copied_msg"))

    def _pick_random_color(self):
        import random
        r = random.randint(0, 255)
        g = random.randint(0, 255)
        b = random.randint(0, 255)
        self._apply_target(f"#{r:02X}{g:02X}{b:02X}")

    # ── HISTORY ────────────────────────────────────────────────────────────────

    def _update_history(self, target_hex, sim_hex, de, seq):
        entry = {"target_hex": target_hex, "sim_hex": sim_hex, "de": de, "sequence": seq}
        self._history = [h for h in self._history if h["target_hex"] != target_hex]
        self._history.insert(0, entry)
        if len(self._history) > 10:
            self._history = self._history[:10]
        self._refresh_history_ui()

    def _refresh_history_ui(self):
        if not hasattr(self, "_history_inner_layout"):
            return
        # Clear
        while self._history_inner_layout.count():
            item = self._history_inner_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        for entry in self._history:
            card = QFrame()
            card.setObjectName("card")
            card_layout = QHBoxLayout(card)
            card_layout.setContentsMargins(4, 2, 4, 2)
            card_layout.setSpacing(4)

            sw = SwatchLabel(entry["target_hex"], size=18, radius=4)
            card_layout.addWidget(sw)

            de_val = entry.get("de", 0)
            seq_str = entry.get("sequence", "")
            lbl = QLabel(f"{entry['target_hex']}  ΔE={de_val:.1f}  [{seq_str}]")
            lbl.setObjectName("hint")
            lbl.setCursor(QCursor(Qt.PointingHandCursor))
            lbl.mousePressEvent = lambda e, h=entry["target_hex"]: self._apply_target(h)
            card_layout.addWidget(lbl, 1)
            self._history_inner_layout.addWidget(card)

        self._history_inner_layout.addStretch()

    # ── VIRTUAL HEADS GRID ────────────────────────────────────────────────────

    def _refresh_virtual_grid(self):
        # Clear existing widgets
        while self._vgrid_layout.count():
            item = self._vgrid_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not self._virtual:
            empty_lbl = QLabel(self.t("empty_virtual"))
            empty_lbl.setObjectName("hint")
            empty_lbl.setWordWrap(True)
            empty_lbl.setAlignment(Qt.AlignCenter)
            self._vgrid_layout.addWidget(empty_lbl)
            self._vgrid_layout.addStretch()
            return

        for vf in self._get_sorted_filtered_virtual():
            row_widget = self._build_virtual_row(vf)
            self._vgrid_layout.addWidget(row_widget)
        self._vgrid_layout.addStretch()

    def _build_virtual_row(self, vf):
        frame = QFrame()
        frame.setObjectName("card")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        # Top row
        top_row = QHBoxLayout()
        top_row.setSpacing(6)

        # Index
        try:
            row_idx = self._virtual.index(vf)
        except ValueError:
            row_idx = 0

        # Up/down buttons
        nav_col = QVBoxLayout()
        up_btn = QPushButton("↑")
        up_btn.setFixedSize(22, 20)
        dn_btn = QPushButton("↓")
        dn_btn.setFixedSize(22, 20)
        up_btn.clicked.connect(lambda checked, i=row_idx: self._vhead_move(i, -1))
        dn_btn.clicked.connect(lambda checked, i=row_idx: self._vhead_move(i, +1))
        nav_col.addWidget(up_btn)
        nav_col.addWidget(dn_btn)
        top_row.addLayout(nav_col)

        # V-ID
        id_lbl = QLabel(f"V{vf['vid']}")
        id_lbl.setStyleSheet("color: #a78bfa; font-weight: bold; font-size: 12pt;")
        id_lbl.setFixedWidth(44)
        top_row.addWidget(id_lbl)

        # Target swatch
        tgt_sw = SwatchLabel(vf["target_hex"], size=36, radius=6)
        top_row.addWidget(tgt_sw)

        # Sequence
        seq_lbl = QLabel(vf["sequence"])
        seq_lbl.setFont(QFont("Courier New", 16, QFont.Bold))
        seq_lbl.setStyleSheet("color: #4ade80;")
        seq_lbl.setMinimumWidth(160)
        top_row.addWidget(seq_lbl)

        # Sim swatch
        sim_sw = SwatchLabel(vf["sim_hex"], size=36, radius=6)
        top_row.addWidget(sim_sw)

        # ΔE
        de_lbl = QLabel(_de_label_text(vf["de"], self.lang))
        de_lbl.setStyleSheet(f"color: {_de_color(vf['de'])}; font-weight: bold;")
        top_row.addWidget(de_lbl)

        # Label edit
        lbl_edit = QLineEdit(vf.get("label", ""))
        lbl_edit.setMaximumWidth(150)
        lbl_edit.editingFinished.connect(
            lambda e=lbl_edit, vfd=vf: vfd.update({"label": e.text()}))
        top_row.addWidget(lbl_edit)

        # Edit sequence button
        edit_btn = QPushButton("✏")
        edit_btn.setFixedSize(28, 30)
        edit_btn.setStyleSheet("background-color: #1e3a5f;")
        edit_btn.clicked.connect(lambda checked: self._open_sequence_editor())
        top_row.addWidget(edit_btn)

        # Delete button
        del_btn = QPushButton("✕")
        del_btn.setFixedSize(30, 30)
        del_btn.setStyleSheet("background-color: #7f1d1d;")
        del_btn.clicked.connect(lambda checked, vid=vf["vid"]: self._remove_virtual(vid))
        top_row.addWidget(del_btn)
        top_row.addStretch()
        layout.addLayout(top_row)

        # Bottom info row: runs + cadence hint
        bottom_row = QHBoxLayout()
        bottom_row.setSpacing(4)

        fils_hex = {i + 1: self._slots[i]["hex_edit"].text() for i in range(4)}
        runs = seq_to_runs(vf["sequence"])
        lh = self._lh_spin.value()
        n_fils = _seq_filament_count(vf["sequence"])

        for fid, cnt in runs:
            col = fils_hex.get(fid, "#64748b")
            try:
                r, g, b = hex_to_rgb(col)
                lum = 0.299 * r + 0.587 * g + 0.114 * b
                txt_col = "#111111" if lum > 128 else "#eeeeee"
            except Exception:
                txt_col = "#eeeeee"
            run_lbl = QLabel(f"T{fid}×{cnt}")
            run_lbl.setFixedHeight(20)
            run_lbl.setStyleSheet(
                f"background-color: {col}; color: {txt_col}; border-radius: 4px; "
                f"padding: 0 4px; font-size: 9pt; font-weight: bold;")
            bottom_row.addWidget(run_lbl)

        # Tool change warning
        transitions = sum(1 for k in range(len(vf["sequence"]) - 1)
                          if vf["sequence"][k] != vf["sequence"][k + 1])
        if transitions > 2:
            warn_lbl = QLabel(self.t("tc_warn_badge", n=transitions))
            warn_lbl.setStyleSheet("color: #f59e0b; font-size: 8pt;")
            bottom_row.addWidget(warn_lbl)

        # Cadence hint
        pat_str = ",".join(vf["sequence"])
        if n_fils == 1:
            hint = self.t("hint_pure")
            hint_col = "#94a3b8"
        elif n_fils == 2 and lh > 0:
            cad = calc_cadence(vf["sequence"], lh)
            ids = sorted(cad.keys())
            hint = self.t("hint_cadence", a=cad[ids[0]],
                          b=cad[ids[1]] if len(ids) > 1 else lh, p=pat_str)
            hint_col = "#4ade80"
        else:
            hint = self.t("hint_pattern", p=pat_str)
            hint_col = "#a78bfa"

        hint_lbl = QLabel(hint)
        hint_lbl.setStyleSheet(f"color: {hint_col}; font-size: 9pt;")
        hint_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        bottom_row.addWidget(hint_lbl)

        # Copy sequence button
        raw_seq = "".join(vf["sequence"])
        copy_btn = QPushButton("📋")
        copy_btn.setFixedSize(24, 20)
        copy_btn.clicked.connect(
            lambda checked, s=raw_seq: QApplication.clipboard().setText(s))
        bottom_row.addWidget(copy_btn)

        bottom_row.addStretch()
        layout.addLayout(bottom_row)
        return frame

    def _remove_virtual(self, vid):
        self._undo_stack.append(copy.deepcopy(self._virtual))
        if len(self._undo_stack) > 20:
            self._undo_stack.pop(0)
        self._virtual = [v for v in self._virtual if v["vid"] != vid]
        for i, v in enumerate(self._virtual):
            v["vid"] = 5 + i
        self._refresh_virtual_grid()

    def _delete_selected_virtual(self):
        """Delete shortcut: removes the last virtual head (Delete key)."""
        if self._virtual:
            self._remove_virtual(self._virtual[-1]["vid"])

    def _vhead_move(self, idx, direction):
        vf = self._virtual
        new_idx = idx + direction
        if 0 <= new_idx < len(vf):
            vf[idx], vf[new_idx] = vf[new_idx], vf[idx]
            for i, v in enumerate(vf):
                v["vid"] = 5 + i
            self._refresh_virtual_grid()

    def _recalc_all_virtual(self):
        if not self._virtual:
            QMessageBox.information(self, self.t("dlg_note"), self.t("orca_no_virtual"))
            return
        self._undo_stack.append(copy.deepcopy(self._virtual))
        use_opt = self._optimizer_check.isChecked()
        use_auto = self._auto_check.isChecked() if hasattr(self, "_auto_check") else True
        thresh   = self._auto_thresh_spin.value() if hasattr(self, "_auto_thresh_spin") else 2.0
        for vf in self._virtual:
            result = self._calc_for_color(
                vf["target_hex"], optimizer=use_opt,
                seq_len=None if use_auto else len(vf["sequence"]),
                auto=use_auto, auto_threshold=thresh)
            if result:
                vf["sequence"] = result["sequence"]
                vf["sim_hex"] = result["sim_hex"]
                vf["de"] = result["de"]
        self._refresh_virtual_grid()
        QMessageBox.information(self, self.t("dlg_note"),
                                self.t("recalc_all_done", n=len(self._virtual)))

    # ── FILAMENT SEARCH ───────────────────────────────────────────────────────

    def _open_filament_search(self, slot_idx):
        existing = self._search_wins.get(slot_idx)
        if existing and existing.isVisible():
            existing.raise_()
            existing.activateWindow()
            return
        # get current slot hex to pre-fill color filter
        slot_hex = None
        if hasattr(self, "_slots") and slot_idx < len(self._slots):
            h = self._slots[slot_idx]["hex_edit"].text().strip()
            if len(h.lstrip("#")) == 6:
                slot_hex = h if h.startswith("#") else "#" + h
        dlg = FilamentSearchDialog(slot_idx, self.library, self, slot_hex=slot_hex, lang=self.lang)
        dlg.filament_selected.connect(self._on_filament_search_select)
        self._search_wins[slot_idx] = dlg
        dlg.show()

    def _on_filament_search_select(self, slot_idx, fil_data):
        self._save_slot_snapshot()
        s = self._slots[slot_idx]
        brand = fil_data.get("brand", "")
        brands = [s["brand_combo"].itemText(j) for j in range(s["brand_combo"].count())]
        if brand in brands:
            s["brand_combo"].setCurrentText(brand)
            self._update_filament_combo(slot_idx)
        name = fil_data.get("name", "")
        for j in range(s["fil_combo"].count()):
            if s["fil_combo"].itemText(j) == name:
                s["fil_combo"].setCurrentIndex(j)
                break
        hex_c = fil_data.get("hex", "")
        if hex_c:
            s["hex_edit"].setText(hex_c)
        td = fil_data.get("td", DEFAULT_TD)
        s["td_spin"].setValue(td)
        self._update_slot_preview(slot_idx)
        self._update_gamut_strip()

    # ── LIBRARY MANAGEMENT ────────────────────────────────────────────────────

    def _add_brand(self):
        name, ok = QInputDialog.getText(self, self.t("inp_brand_title"),
                                        self.t("inp_brand_name"))
        if not ok or not name:
            return
        name = name.strip()
        if name in self.library:
            QMessageBox.warning(self, self.t("dlg_exists"),
                                self.t("dlg_exists_msg", name=name))
            return
        self.library[name] = []
        self._save_db()
        self._refresh_slot_brand_combos()

    def _refresh_slot_brand_combos(self):
        brands = list(self.library.keys())
        for i, s in enumerate(self._slots):
            cur = s["brand_combo"].currentText()
            s["brand_combo"].blockSignals(True)
            s["brand_combo"].clear()
            s["brand_combo"].addItems(brands)
            if cur in brands:
                s["brand_combo"].setCurrentText(cur)
            else:
                s["brand_combo"].setCurrentIndex(0)
            s["brand_combo"].blockSignals(False)
            self._update_filament_combo(i)

    def _open_library_manager(self):
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("lib_title"))
        dlg.resize(720, 560)
        layout = QVBoxLayout(dlg)

        top_row = QHBoxLayout()
        top_row.addWidget(QLabel(self.t("lib_brand")))
        brand_combo = QComboBox()
        brand_combo.addItems(list(self.library.keys()))
        top_row.addWidget(brand_combo, 1)

        del_brand_btn = QPushButton(self.t("lib_del_brand"))
        del_brand_btn.setStyleSheet("background-color: #7f1d1d;")
        top_row.addWidget(del_brand_btn)
        layout.addLayout(top_row)

        fil_list = QListWidget()
        fil_list.setAlternatingRowColors(True)
        layout.addWidget(fil_list, 1)

        def refresh():
            fil_list.clear()
            brand = brand_combo.currentText()
            for fil in self.library.get(brand, []):
                item = QListWidgetItem(f"{fil['name']}  {fil.get('hex', '')}  TD={fil.get('td', '')}")
                item.setData(Qt.UserRole, fil)
                hex_c = fil.get("hex", "#808080")
                try:
                    r, g, b = hex_to_rgb(hex_c)
                    item.setBackground(QColor(r, g, b))
                    lum = 0.299 * r + 0.587 * g + 0.114 * b
                    item.setForeground(QColor("#111111" if lum > 128 else "#eeeeee"))
                except Exception:
                    pass
                fil_list.addItem(item)

        brand_combo.currentTextChanged.connect(lambda t: refresh())

        def del_fil():
            item = fil_list.currentItem()
            if not item:
                return
            fil = item.data(Qt.UserRole)
            brand = brand_combo.currentText()
            if QMessageBox.question(dlg, self.t("dlg_del_title"),
                                    self.t("lib_del_fil", n=fil["name"]),
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                self.library[brand] = [f for f in self.library[brand]
                                       if f["name"] != fil["name"]]
                self._save_db()
                self._refresh_slot_brand_combos()
                refresh()

        def del_brand():
            brand = brand_combo.currentText()
            from copy import deepcopy as _dc
            from collections import OrderedDict as _od
            if brand in DEFAULT_LIBRARY and DEFAULT_LIBRARY.get(brand):
                QMessageBox.critical(dlg, self.t("lib_protected"),
                                     self.t("lib_protected_msg"))
                return
            if QMessageBox.question(dlg, self.t("dlg_del_title"),
                                    self.t("lib_del_brand_msg", b=brand),
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                del self.library[brand]
                self._save_db()
                self._refresh_slot_brand_combos()
                brand_combo.blockSignals(True)
                brand_combo.clear()
                brand_combo.addItems(list(self.library.keys()))
                brand_combo.blockSignals(False)
                refresh()

        def add_fil():
            brand = brand_combo.currentText()
            name, ok = QInputDialog.getText(dlg, self.t("inp_add_title"),
                                            self.t("inp_name"))
            if not ok or not name:
                return
            hex_str, ok2 = QInputDialog.getText(dlg, self.t("inp_color_title"),
                                                self.t("inp_hex"))
            if not ok2 or not hex_str:
                return
            hex_str = hex_str.strip()
            if not hex_str.startswith("#"):
                hex_str = "#" + hex_str
            td_str, ok3 = QInputDialog.getText(dlg, self.t("inp_td_title"),
                                               self.t("inp_td2", td=DEFAULT_TD))
            td = safe_td(td_str) if ok3 and td_str else DEFAULT_TD
            self.library.setdefault(brand, []).append(
                {"name": name.strip(), "hex": hex_str, "td": td})
            self._save_db()
            self._refresh_slot_brand_combos()
            refresh()

        del_fil_btn = QPushButton(self.t("lib_del_fil", n="selected").split('"')[0] + "Delete")
        del_fil_btn.setStyleSheet("background-color: #7f1d1d;")
        del_fil_btn.clicked.connect(del_fil)
        del_brand_btn.clicked.connect(del_brand)

        btm_row = QHBoxLayout()
        add_btn = QPushButton(self.t("lib_add_fil"))
        add_btn.setObjectName("btn_green")
        add_btn.clicked.connect(add_fil)
        close_btn = QPushButton(self.t("lib_close"))
        close_btn.clicked.connect(dlg.accept)
        btm_row.addWidget(add_btn)
        btm_row.addWidget(del_fil_btn)
        btm_row.addStretch()
        btm_row.addWidget(close_btn)
        layout.addLayout(btm_row)

        refresh()
        dlg.exec()

    # ── BATCH DIALOG ──────────────────────────────────────────────────────────

    def _open_batch_dialog(self):
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("batch_title"))
        dlg.resize(480, 460)
        layout = QVBoxLayout(dlg)

        layout.addWidget(QLabel(self.t("batch_title")))
        desc = QLabel(self.t("batch_desc"))
        desc.setObjectName("hint")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        txt_edit = QTextEdit()
        txt_edit.setFont(QFont("Courier New", 11))
        layout.addWidget(txt_edit, 1)

        # SVG import row
        svg_row = QHBoxLayout()
        svg_btn = QPushButton(self.t("batch_import_svg"))
        svg_btn.setFixedHeight(32)
        svg_btn.setToolTip("Parse fill=\"#RRGGBB\" attributes from an SVG file and append to the list.")

        def _import_svg():
            svg_path, _ = QFileDialog.getOpenFileName(
                dlg, self.t("batch_import_svg"), "",
                "SVG Files (*.svg);;All Files (*.*)")
            if not svg_path:
                return
            try:
                with open(svg_path, encoding="utf-8", errors="replace") as _f:
                    svg_text = _f.read()
            except Exception as _e:
                QMessageBox.warning(dlg, self.t("dlg_error"), str(_e))
                return
            found = re.findall(r'fill=["\']#([0-9A-Fa-f]{6})["\']', svg_text)
            # Deduplicate preserving order
            seen = set()
            unique = []
            for c in found:
                key = c.upper()
                if key not in seen:
                    seen.add(key)
                    unique.append("#" + key)
            if unique:
                existing = txt_edit.toPlainText().strip()
                separator = "\n" if existing else ""
                txt_edit.setPlainText(existing + separator + "\n".join(unique))
            else:
                QMessageBox.information(dlg, self.t("dlg_note"),
                                        "No fill=\"#RRGGBB\" attributes found in SVG.")

        svg_btn.clicked.connect(_import_svg)
        svg_row.addWidget(svg_btn)
        svg_row.addStretch()
        layout.addLayout(svg_row)

        prog_lbl = QLabel("")
        prog_lbl.setObjectName("hint")
        layout.addWidget(prog_lbl)

        btn_row = QHBoxLayout()
        calc_btn = QPushButton(self.t("batch_btn_calc"))
        calc_btn.setObjectName("btn_primary")
        calc_btn.setFixedHeight(40)
        cancel_btn = QPushButton(self.t("batch_btn_cancel"))
        cancel_btn.setFixedHeight(40)
        btn_row.addWidget(calc_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

        def do_batch():
            lines = txt_edit.toPlainText().strip().splitlines()
            hexes = []
            for line in lines:
                h = line.strip()
                if not h:
                    continue
                if not h.startswith("#"):
                    h = "#" + h
                if re.match(r'^#[0-9A-Fa-f]{6}$', h):
                    hexes.append(h.upper())
            added = 0
            for i, h in enumerate(hexes):
                if len(self._virtual) >= self._max_virtual:
                    QMessageBox.warning(dlg, self.t("dlg_max_virtual"),
                                        self.t("batch_warn_max", max_v=self._max_virtual))
                    break
                prog_lbl.setText(f"{i + 1}/{len(hexes)}  {h}")
                QApplication.processEvents()
                r = self._calc_for_color(
                    h,
                    optimizer=self._optimizer_check.isChecked(),
                    seq_len=self._len_spin.value() if not self._auto_check.isChecked() else None,
                    auto=self._auto_check.isChecked(),
                    auto_threshold=self._auto_thresh_spin.value())
                if r:
                    self.add_virtual(r)
                    added += 1
            dlg.accept()
            QMessageBox.information(self, self.t("dlg_3mf_title"),
                                    self.t("batch_done", n=added))

        calc_btn.clicked.connect(do_batch)
        cancel_btn.clicked.connect(dlg.reject)
        dlg.exec()

    # ── 3MF ASSISTANT ─────────────────────────────────────────────────────────

    def _open_3mf_assistant(self):
        path, _ = QFileDialog.getOpenFileName(
            self, self.t("open_3mf_title"), "",
            f"{self.t('3mf_filetypes')} (*.3mf);;All Files (*.*)")
        if not path:
            return
        self._open_3mf_with_path(path)

    def _open_3mf_with_path(self, path):
        """Open 3MF assistant with a given file path (used by DnD)."""
        colors, err = _parse_3mf_colors(path)
        if not colors:
            QMessageBox.information(self, self.t("dlg_3mf_title"),
                                    err if err else self.t("dlg_3mf_no_colors_fallback"))
            return
        self._open_3mf_with_colors(path, colors)

    def _open_3mf_with_colors(self, path, colors):
        """Open 3MF assistant dialog with pre-loaded colors."""
        dlg = QDialog(self)
        dlg.setWindowTitle(f"{self.t('dlg_3mf_title')} — {os.path.basename(path)}")
        dlg.resize(860, 680)
        layout = QVBoxLayout(dlg)

        layout.addWidget(QLabel(self.t("3mf_analysis_title", n=len(colors))))

        slots = self._slot_filaments()
        basis_txt = self.t("3mf_basis",
                           t1=slots[0]["hex"], t2=slots[1]["hex"],
                           t3=slots[2]["hex"], t4=slots[3]["hex"])
        basis_lbl = QLabel(basis_txt)
        basis_lbl.setObjectName("hint")
        layout.addWidget(basis_lbl)

        opt_check = QCheckBox(self.t("3mf_optimizer"))
        layout.addWidget(opt_check)

        prog_lbl = QLabel(self.t("3mf_ready"))
        prog_lbl.setObjectName("hint")
        layout.addWidget(prog_lbl)

        # Results table
        table = QTableWidget(len(colors), 5)
        table.setHorizontalHeaderLabels([
            "#", self.t("3mf_col_target"), self.t("3mf_col_seq"),
            self.t("3mf_col_sim"), self.t("3mf_col_quality")])
        table.horizontalHeader().setStretchLastSection(True)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(table, 1)

        results = [None] * len(colors)
        include_checks = []

        def render_rows():
            for idx, hex_c in enumerate(colors[:self._max_virtual]):
                table.setItem(idx, 0, QTableWidgetItem(str(idx + 1)))
                tgt_item = QTableWidgetItem("")
                try:
                    r, g, b = hex_to_rgb(hex_c)
                    tgt_item.setBackground(QColor(r, g, b))
                except Exception:
                    pass
                table.setItem(idx, 1, tgt_item)
                r = results[idx]
                if r:
                    table.setItem(idx, 2, QTableWidgetItem(r["sequence"]))
                    sim_item = QTableWidgetItem("")
                    try:
                        sr, sg, sb = hex_to_rgb(r["sim_hex"])
                        sim_item.setBackground(QColor(sr, sg, sb))
                    except Exception:
                        pass
                    table.setItem(idx, 3, sim_item)
                    de_item = QTableWidgetItem(_de_label_text(r["de"], self.lang))
                    de_item.setForeground(QColor(_de_color(r["de"])))
                    table.setItem(idx, 4, de_item)
                else:
                    table.setItem(idx, 2, QTableWidgetItem(self.t("3mf_not_calc")))
                    table.setItem(idx, 3, QTableWidgetItem(""))
                    table.setItem(idx, 4, QTableWidgetItem(""))

        render_rows()

        # Include checkboxes
        chk_row = QHBoxLayout()
        chk_all = QCheckBox("Select All")
        chk_all.setChecked(True)
        include_vars = [True] * len(colors)

        def on_chk_all(state):
            for i in range(len(include_vars)):
                include_vars[i] = bool(state)
        chk_all.stateChanged.connect(on_chk_all)
        chk_row.addWidget(chk_all)
        layout.addLayout(chk_row)

        def run_all():
            opt = opt_check.isChecked()
            for idx in range(min(len(colors), self._max_virtual)):
                prog_lbl.setText(self.t("3mf_progress", i=idx + 1,
                                         total=min(len(colors), self._max_virtual),
                                         c=colors[idx]))
                QApplication.processEvents()
                r = self._calc_for_color(
                    colors[idx], optimizer=opt,
                    seq_len=self._len_spin.value(),
                    auto=self._auto_check.isChecked(),
                    auto_threshold=self._auto_thresh_spin.value())
                results[idx] = r
            prog_lbl.setText(self.t("3mf_done", n=min(len(colors), self._max_virtual)))
            render_rows()

        def apply_selected():
            added = 0
            for idx, include in enumerate(include_vars):
                if not include or results[idx] is None:
                    continue
                if len(self._virtual) >= self._max_virtual:
                    break
                self.add_virtual(results[idx])
                added += 1
            dlg.accept()
            QMessageBox.information(self, self.t("dlg_3mf_title"),
                                    self.t("dlg_3mf_added", n=added))

        btn_row = QHBoxLayout()
        calc_btn = QPushButton(self.t("3mf_btn_calc"))
        calc_btn.setObjectName("btn_primary")
        calc_btn.setFixedHeight(42)
        calc_btn.clicked.connect(run_all)
        apply_btn = QPushButton(self.t("3mf_btn_apply"))
        apply_btn.setObjectName("btn_green")
        apply_btn.setFixedHeight(42)
        apply_btn.clicked.connect(apply_selected)
        cancel_btn = QPushButton(self.t("3mf_btn_cancel"))
        cancel_btn.setFixedHeight(42)
        cancel_btn.clicked.connect(dlg.reject)
        btn_row.addWidget(calc_btn)
        btn_row.addWidget(apply_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)
        dlg.exec()

    # ── OBJ/MTL ASSISTANT ─────────────────────────────────────────────────────

    def _open_obj_assistant(self, path=None):
        """Open OBJ/MTL color assistant dialog."""
        if path is None:
            path, _ = QFileDialog.getOpenFileName(
                self, self.t("obj_title"), "",
                f"{self.t('obj_filetypes')} (*.obj);;All Files (*.*)")
        if not path:
            return

        colors = parse_obj_mtl_colors(path)
        if not colors:
            if not _HAS_PIL:
                QMessageBox.warning(self, self.t("obj_title"), self.t("obj_no_pil"))
            else:
                QMessageBox.information(self, self.t("obj_title"), self.t("obj_no_colors"))
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(f"{self.t('obj_title')} — {os.path.basename(path)}")
        dlg.resize(860, 680)
        layout = QVBoxLayout(dlg)

        layout.addWidget(QLabel(self.t("3mf_analysis_title", n=len(colors))))

        slots = self._slot_filaments()
        basis_txt = self.t("3mf_basis",
                           t1=slots[0]["hex"], t2=slots[1]["hex"],
                           t3=slots[2]["hex"], t4=slots[3]["hex"])
        basis_lbl = QLabel(basis_txt)
        basis_lbl.setObjectName("hint")
        layout.addWidget(basis_lbl)

        opt_check = QCheckBox(self.t("3mf_optimizer"))
        layout.addWidget(opt_check)

        prog_lbl = QLabel(self.t("3mf_ready"))
        prog_lbl.setObjectName("hint")
        layout.addWidget(prog_lbl)

        # Results table
        table = QTableWidget(len(colors), 5)
        table.setHorizontalHeaderLabels([
            "#", self.t("3mf_col_target"), self.t("3mf_col_seq"),
            self.t("3mf_col_sim"), self.t("3mf_col_quality")])
        table.horizontalHeader().setStretchLastSection(True)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(table, 1)

        results = [None] * len(colors)
        include_vars = [True] * len(colors)

        def render_rows():
            for idx, hex_c in enumerate(colors[:self._max_virtual]):
                table.setItem(idx, 0, QTableWidgetItem(str(idx + 1)))
                tgt_item = QTableWidgetItem("")
                try:
                    rr, gg, bb = hex_to_rgb(hex_c)
                    tgt_item.setBackground(QColor(rr, gg, bb))
                except Exception:
                    pass
                table.setItem(idx, 1, tgt_item)
                r = results[idx]
                if r:
                    table.setItem(idx, 2, QTableWidgetItem(r["sequence"]))
                    sim_item = QTableWidgetItem("")
                    try:
                        sr, sg, sb = hex_to_rgb(r["sim_hex"])
                        sim_item.setBackground(QColor(sr, sg, sb))
                    except Exception:
                        pass
                    table.setItem(idx, 3, sim_item)
                    de_item = QTableWidgetItem(_de_label_text(r["de"], self.lang))
                    de_item.setForeground(QColor(_de_color(r["de"])))
                    table.setItem(idx, 4, de_item)
                else:
                    table.setItem(idx, 2, QTableWidgetItem(self.t("3mf_not_calc")))
                    table.setItem(idx, 3, QTableWidgetItem(""))
                    table.setItem(idx, 4, QTableWidgetItem(""))

        render_rows()

        chk_row = QHBoxLayout()
        chk_all = QCheckBox("Select All")
        chk_all.setChecked(True)

        def on_chk_all(state):
            for i in range(len(include_vars)):
                include_vars[i] = bool(state)
        chk_all.stateChanged.connect(on_chk_all)
        chk_row.addWidget(chk_all)
        layout.addLayout(chk_row)

        def run_all():
            opt = opt_check.isChecked()
            for idx in range(min(len(colors), self._max_virtual)):
                prog_lbl.setText(self.t("3mf_progress", i=idx + 1,
                                         total=min(len(colors), self._max_virtual),
                                         c=colors[idx]))
                QApplication.processEvents()
                r = self._calc_for_color(
                    colors[idx], optimizer=opt,
                    seq_len=self._len_spin.value(),
                    auto=self._auto_check.isChecked(),
                    auto_threshold=self._auto_thresh_spin.value())
                results[idx] = r
            prog_lbl.setText(self.t("3mf_done", n=min(len(colors), self._max_virtual)))
            render_rows()

        def apply_selected():
            added = 0
            for idx, include in enumerate(include_vars):
                if not include or results[idx] is None:
                    continue
                if len(self._virtual) >= self._max_virtual:
                    break
                self.add_virtual(results[idx])
                added += 1
            dlg.accept()
            QMessageBox.information(self, self.t("obj_title"),
                                    self.t("dlg_3mf_added", n=added))

        btn_row = QHBoxLayout()
        calc_btn = QPushButton(self.t("3mf_btn_calc"))
        calc_btn.setObjectName("btn_primary")
        calc_btn.setFixedHeight(42)
        calc_btn.clicked.connect(run_all)
        apply_btn = QPushButton(self.t("3mf_btn_apply"))
        apply_btn.setObjectName("btn_green")
        apply_btn.setFixedHeight(42)
        apply_btn.clicked.connect(apply_selected)
        cancel_btn = QPushButton(self.t("3mf_btn_cancel"))
        cancel_btn.setFixedHeight(42)
        cancel_btn.clicked.connect(dlg.reject)
        btn_row.addWidget(calc_btn)
        btn_row.addWidget(apply_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)
        dlg.exec()

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

    def _open_3mf_wizard(self):
        if self._3mf_wizard is not None and self._3mf_wizard.isVisible():
            self._3mf_wizard.raise_()
            self._3mf_wizard.activateWindow()
            return
        dlg = ThreeMFWizardDialog(self)
        self._3mf_wizard = dlg
        dlg.finished.connect(lambda: setattr(self, "_3mf_wizard", None))
        dlg.show()

    # ── EXPORT DIALOG ─────────────────────────────────────────────────────────

    def _open_export_dialog(self):
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("exp_title"))
        dlg.resize(500, 440)
        layout = QVBoxLayout(dlg)

        layout.addWidget(QLabel(self.t("exp_header")))

        # Format
        from PySide6.QtWidgets import QRadioButton, QButtonGroup
        fmt_group = QButtonGroup(dlg)
        fmt_row = QHBoxLayout()
        json_rb = QRadioButton("JSON (.json)")
        txt_rb = QRadioButton("Text (.txt)")
        json_rb.setChecked(True)
        fmt_group.addButton(json_rb, 0)
        fmt_group.addButton(txt_rb, 1)
        fmt_row.addWidget(json_rb)
        fmt_row.addWidget(txt_rb)
        layout.addLayout(fmt_row)

        # Scope
        scope_group = QButtonGroup(dlg)
        scope_row = QHBoxLayout()
        single_rb = QRadioButton(self.t("exp_scope_single"))
        virtual_rb = QRadioButton(self.t("exp_scope_virtual", n=len(self._virtual)))
        if self._virtual:
            virtual_rb.setChecked(True)
        else:
            single_rb.setChecked(True)
        scope_group.addButton(single_rb, 0)
        scope_group.addButton(virtual_rb, 1)
        scope_row.addWidget(single_rb)
        scope_row.addWidget(virtual_rb)
        layout.addLayout(scope_row)

        # Layer height
        lh_row = QHBoxLayout()
        lh_row.addWidget(QLabel(self.t("exp_lh_label")))
        lh_exp_spin = QDoubleSpinBox()
        lh_exp_spin.setRange(0.01, 1.0)
        lh_exp_spin.setDecimals(3)
        lh_exp_spin.setSingleStep(0.01)
        lh_exp_spin.setValue(self._lh_spin.value())
        lh_row.addWidget(lh_exp_spin)
        lh_row.addWidget(QLabel(self.t("exp_lh_unit")))
        lh_row.addStretch()
        layout.addLayout(lh_row)

        def do_export():
            is_json = fmt_group.checkedId() == 0
            is_virtual = scope_group.checkedId() == 1
            ext = ".json" if is_json else ".txt"
            path, _ = QFileDialog.getSaveFileName(
                dlg, self.t("save_dialog_title"), "",
                f"{'JSON' if is_json else 'Text'} (*{ext})")
            if not path:
                return
            lh = lh_exp_spin.value()

            physical = [
                {"id": i + 1, "brand": s["brand_combo"].currentText(),
                 "filament": s["fil_combo"].currentText(),
                 "hex": s["hex_edit"].text(),
                 "td": s["td_spin"].value()}
                for i, s in enumerate(self._slots)
            ]

            def resolve_cadence(seq):
                if not seq:
                    return 0.0, 0.0
                cad = calc_cadence(seq, lh)
                ids = sorted(cad.keys())
                a = cad.get(ids[0], lh) if ids else lh
                b = cad.get(ids[1], lh) if len(ids) > 1 else lh
                return round(a, 4), round(b, 4)

            try:
                if is_json:
                    seq_now = self._seq_label.text()
                    n_now = len(seq_now) if seq_now and seq_now != "----------" else 10
                    lw = _get_layer_weights(n_now)
                    payload = {
                        "generator": "U1 FullSpectrum Ultimate — PySide6 Edition",
                        "timestamp": datetime.now().isoformat(),
                        "dithering_step_size": lh,
                        "layer_weights": {f"L{i+1}": round(lw[i], 4) for i in range(len(lw))},
                        "physical_slots": physical,
                    }
                    if is_virtual:
                        def vf_entry(v):
                            a_v, b_v = resolve_cadence(v["sequence"])
                            runs = seq_to_runs(v["sequence"])
                            n_f = _seq_filament_count(v["sequence"])
                            pat = ",".join(v["sequence"])
                            mode = "cadence" if n_f <= 2 else "pattern"
                            return {**v, "runs": runs, "filament_count": n_f,
                                    "slicer_input_mode": mode, "pattern_string": pat,
                                    "cadence_a_mm": a_v if mode == "cadence" else None,
                                    "cadence_b_mm": b_v if mode == "cadence" else None,
                                    "dithering_step_size_mm": lh}
                        payload["virtual_filaments"] = [vf_entry(v) for v in self._virtual]
                    else:
                        a_v, b_v = resolve_cadence(seq_now)
                        payload["sequence"] = seq_now
                        payload["target_hex"] = self._target_hex or ""
                        payload["cadence_a_mm"] = a_v
                        payload["cadence_b_mm"] = b_v
                        payload["runs"] = seq_to_runs(seq_now)
                    with open(path, "w", encoding="utf-8") as f:
                        json.dump(payload, f, indent=2, ensure_ascii=False)
                else:
                    lines = [
                        "=" * 56,
                        "  U1 FullSpectrum Ultimate — OrcaSlicer FullSpectrum Export",
                        "=" * 56,
                        f"{self.t('txt_date')}  {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                        f"{self.t('txt_layer_height')}  {lh} mm",
                        "", self.t("txt_physical_heads"), "-" * 36,
                    ]
                    for p in physical:
                        lines.append(
                            f"  T{p['id']}: {p['brand']} — {p['filament']}  ({p['hex']}, TD={p['td']})")
                    if is_virtual and self._virtual:
                        lines += ["", self.t("txt_virtual_heads"), "-" * 56]
                        for v in self._virtual:
                            a_v, b_v = resolve_cadence(v["sequence"])
                            n_f = _seq_filament_count(v["sequence"])
                            runs_str = "  ".join(
                                f"T{fid}x{cnt}" for fid, cnt in seq_to_runs(v["sequence"]))
                            pat_str = ",".join(v["sequence"])
                            if n_f == 1:
                                slicer_hint = self.t("txt_pure")
                            elif n_f == 2:
                                slicer_hint = self.t("txt_cadence", a=a_v, b=b_v, p=pat_str)
                            else:
                                slicer_hint = self.t("txt_pattern", p=pat_str)
                            lines += [
                                f"  V{v['vid']}  [{v.get('label','')}]"
                                f"  Target:{v['target_hex']}  dE={v['de']:.1f}",
                                f"    Seq: {v['sequence']}",
                                f"    Runs: {runs_str}",
                                f"    Slicer: {slicer_hint}",
                            ]
                    with open(path, "w", encoding="utf-8") as f:
                        f.write("\n".join(lines))
                QMessageBox.information(dlg, self.t("dlg_saved"),
                                        self.t("dlg_export_saved", path=path))
                dlg.accept()
            except IOError as e:
                QMessageBox.critical(dlg, self.t("dlg_error"), str(e))

        exp_btn = QPushButton(self.t("exp_btn"))
        exp_btn.setObjectName("btn_primary")
        exp_btn.setFixedHeight(44)
        exp_btn.clicked.connect(do_export)
        layout.addWidget(exp_btn)

        cancel_btn = QPushButton(self.t("exp_cancel"))
        cancel_btn.setFixedHeight(36)
        cancel_btn.clicked.connect(dlg.reject)
        layout.addWidget(cancel_btn)
        dlg.exec()

    # ── ORCA SLICER EXPORT ────────────────────────────────────────────────────

    @staticmethod
    def _find_orca_installations():
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
                for entry in os.scandir(appdata):
                    if not entry.is_dir():
                        continue
                    name = entry.name
                    if any(x in name.lower() for x in
                           ["orca", "snapmaker_orca", "bambu", "orca slicer"]):
                        for sub in ["user/default/filament", "user/filament"]:
                            p = os.path.join(entry.path, sub)
                            add(f"{name}  ({p})", p)
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

        found.sort(key=lambda x: (0 if x["exists"] else 1, x["label"]))
        return found

    def _build_orca_filament_json(self, name, hex_color, notes,
                                   filament_type="PLA", target_path=""):
        hex_clean = hex_color if hex_color.startswith("#") else f"#{hex_color}"
        ft = filament_type.upper()
        if "snapmaker" in target_path.lower():
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

    def _open_orca_export_dialog(self):
        installs = self._find_orca_installations()
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("orca_title"))
        dlg.resize(600, 480)
        layout = QVBoxLayout(dlg)

        layout.addWidget(QLabel(self.t("orca_header")))

        path_edit = QLineEdit(installs[0]["path"] if installs else "")
        path_row = QHBoxLayout()
        path_row.addWidget(path_edit, 1)
        browse_btn = QPushButton(self.t("orca_path_browse"))

        def browse():
            p = QFileDialog.getExistingDirectory(dlg, self.t("orca_path_label"))
            if p:
                path_edit.setText(p)

        browse_btn.clicked.connect(browse)
        path_row.addWidget(browse_btn)
        layout.addLayout(path_row)

        # Scope
        from PySide6.QtWidgets import QRadioButton, QButtonGroup
        scope_group = QButtonGroup(dlg)
        phys_rb = QRadioButton(self.t("orca_scope_phys"))
        virt_rb = QRadioButton(self.t("orca_scope_virt"))
        both_rb = QRadioButton(self.t("orca_scope_both"))
        both_rb.setChecked(True)
        scope_group.addButton(phys_rb, 0)
        scope_group.addButton(virt_rb, 1)
        scope_group.addButton(both_rb, 2)
        layout.addWidget(phys_rb)
        layout.addWidget(virt_rb)
        layout.addWidget(both_rb)

        # Prefix
        prefix_row = QHBoxLayout()
        prefix_row.addWidget(QLabel(self.t("orca_prefix_label")))
        prefix_edit = QLineEdit("U1")
        prefix_edit.setMaximumWidth(100)
        prefix_row.addWidget(prefix_edit)
        hint_lbl = QLabel(self.t("orca_prefix_hint"))
        hint_lbl.setObjectName("hint")
        prefix_row.addWidget(hint_lbl)
        prefix_row.addStretch()
        layout.addLayout(prefix_row)

        # LH
        lh_row = QHBoxLayout()
        lh_row.addWidget(QLabel(self.t("exp_lh_label")))
        lh_orca = QDoubleSpinBox()
        lh_orca.setRange(0.01, 1.0)
        lh_orca.setDecimals(3)
        lh_orca.setSingleStep(0.01)
        lh_orca.setValue(self._lh_spin.value())
        lh_row.addWidget(lh_orca)
        lh_row.addStretch()
        layout.addLayout(lh_row)

        def do_orca_export():
            folder = path_edit.text().strip()
            if not folder:
                QMessageBox.critical(dlg, self.t("dlg_error"), self.t("orca_no_path"))
                return
            scope_id = scope_group.checkedId()
            prefix = prefix_edit.text().strip() or "U1"
            lh = lh_orca.value()
            profiles = []

            if scope_id in (0, 2):
                for i, s in enumerate(self._slots):
                    hex_c = s["hex_edit"].text().strip() or "#888888"
                    brand = s["brand_combo"].currentText()
                    name = s["fil_combo"].currentText()
                    td = s["td_spin"].value()
                    pname = f"{prefix}-T{i+1} {name}" if name and name not in _SLOT_SKIP else f"{prefix}-T{i+1}"
                    notes = self.t("orca_filament_notes_t", i=i+1, brand=brand, name=name, td=td)
                    data = self._build_orca_filament_json(pname, hex_c, notes, "PLA", folder)
                    safe_fn = re.sub(r'[\\/:*?"<>|]', "_", pname) + ".json"
                    profiles.append((safe_fn, data))

            if scope_id in (1, 2):
                if not self._virtual:
                    QMessageBox.warning(dlg, self.t("dlg_saved"), self.t("orca_no_virtual"))
                else:
                    for v in self._virtual:
                        sim_hex = v.get("sim_hex", v.get("target_hex", "#888888"))
                        seq = v["sequence"]
                        n_f = _seq_filament_count(seq)
                        runs_str = "  ".join(f"T{fid}x{cnt}" for fid, cnt in seq_to_runs(seq))
                        cad = calc_cadence(seq, lh)
                        ids = sorted(cad.keys())
                        if n_f == 1:
                            hint = "Pure color"
                        elif n_f == 2:
                            a = round(cad.get(ids[0], lh), 3)
                            b = round(cad.get(ids[1], lh) if len(ids) > 1 else lh, 3)
                            hint = f"Cadence A={a}mm B={b}mm | Step={lh}mm"
                        else:
                            pat = "/".join(seq)
                            hint = f"Pattern={pat} | Step={lh}mm"
                        label = v.get("label", f"V{v['vid']}")
                        pname = f"{prefix}-V{v['vid']} {label}"
                        notes = self.t("orca_filament_notes_v",
                                       seq=runs_str, de=v.get("de", 0.0), hint=hint)
                        data = self._build_orca_filament_json(pname, sim_hex, notes, "PLA", folder)
                        safe_fn = re.sub(r'[\\/:*?"<>|]', "_", pname) + ".json"
                        profiles.append((safe_fn, data))

            if not profiles:
                return
            existing = [fn for fn, _ in profiles if os.path.exists(os.path.join(folder, fn))]
            if existing:
                if QMessageBox.question(dlg, self.t("orca_title"),
                                        self.t("orca_overwrite_confirm", n=len(existing)),
                                        QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
                    return
            os.makedirs(folder, exist_ok=True)
            written = 0
            errors = []
            for fn, data in profiles:
                try:
                    with open(os.path.join(folder, fn), "w", encoding="utf-8") as f:
                        json.dump(data, f, indent=4, ensure_ascii=False)
                    written += 1
                except IOError as e:
                    errors.append(str(e))
            if errors:
                QMessageBox.critical(dlg, self.t("dlg_error"), "\n".join(errors[:3]))
            else:
                QMessageBox.information(dlg, self.t("dlg_saved"),
                                        self.t("orca_success", n=written))
                dlg.accept()

        export_btn = QPushButton(self.t("orca_btn_export"))
        export_btn.setObjectName("btn_primary")
        export_btn.setFixedHeight(44)
        export_btn.clicked.connect(do_orca_export)
        layout.addWidget(export_btn)

        cancel_btn = QPushButton(self.t("orca_btn_cancel"))
        cancel_btn.setFixedHeight(36)
        cancel_btn.clicked.connect(dlg.reject)
        layout.addWidget(cancel_btn)
        dlg.exec()

    # ── 3MF WRITE-BACK ────────────────────────────────────────────────────────

    def _write_3mf_colors(self):
        import xml.etree.ElementTree as ET
        path, _ = QFileDialog.getOpenFileName(
            self, self.t("3mf_write_title"), "",
            "3MF (*.3mf);;All Files (*.*)")
        if not path:
            return

        try:
            with zipfile.ZipFile(path, "r") as zin:
                names = zin.namelist()
                ms_raw = zin.read("Metadata/model_settings.config").decode("utf-8")
                ms_root = ET.fromstring(ms_raw)
                ext_objects = {}
                for obj in ms_root.findall("object"):
                    obj_name = next((m.get("value") for m in obj.findall("metadata")
                                     if m.get("key") == "name"), "?")
                    ext_val = next((m.get("value") for m in obj.findall("metadata")
                                    if m.get("key") == "extruder"), "1")
                    ext_objects.setdefault(ext_val, []).append(obj_name)
        except Exception as e:
            QMessageBox.critical(self, self.t("dlg_error"), str(e))
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("remap_title"))
        dlg.resize(680, 520)
        layout = QVBoxLayout(dlg)
        layout.addWidget(QLabel(self.t("remap_col_hdr")))

        slot_opts = [f"T{i+1} — {self._slots[i]['fil_combo'].currentText()}" for i in range(4)]
        virt_opts = [f"V{vf['vid']} — {vf.get('label','?')} (dE {vf.get('de',0):.1f})"
                     for vf in self._virtual]
        all_opts = slot_opts + virt_opts + [self.t("remap_keep")]

        remap_combos = {}
        sorted_exts = sorted(ext_objects.keys(), key=lambda x: int(x) if x.isdigit() else 99)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_inner = QWidget()
        scroll_layout = QVBoxLayout(scroll_inner)
        scroll.setWidget(scroll_inner)
        layout.addWidget(scroll, 1)

        for ext in sorted_exts:
            row = QHBoxLayout()
            row.addWidget(QLabel(f"Extruder {ext}"))
            obj_list = ", ".join(ext_objects[ext])[:55]
            obj_lbl = QLabel(obj_list)
            obj_lbl.setObjectName("hint")
            row.addWidget(obj_lbl, 1)
            combo = QComboBox()
            combo.addItems(all_opts)
            default_idx = max(0, min(int(ext) - 1, len(all_opts) - 1)) if ext.isdigit() else len(all_opts) - 1
            combo.setCurrentIndex(default_idx)
            remap_combos[ext] = combo
            row.addWidget(combo)
            scroll_layout.addLayout(row)

        lh = self._lh_spin.value()

        def do_write():
            save_path, _ = QFileDialog.getSaveFileName(
                dlg, "Save As", "",
                "3MF (*.3mf)")
            if not save_path:
                return
            try:
                import shutil, tempfile
                with zipfile.ZipFile(path, "r") as zin:
                    with tempfile.NamedTemporaryFile(suffix=".3mf", delete=False) as tmp:
                        tmp_path = tmp.name
                    ext_map = {}
                    vf_for_ext = {}
                    slot_for_ext = {}
                    for ext, combo in remap_combos.items():
                        choice = combo.currentText()
                        old_e = int(ext) if ext.isdigit() else None
                        if old_e is None or choice == self.t("remap_keep"):
                            continue
                        if choice.startswith("T"):
                            try:
                                slot_i = int(choice[1]) - 1
                            except (IndexError, ValueError):
                                continue
                            new_e = slot_i + 1
                            ext_map[old_e] = new_e
                            slot_for_ext[new_e] = slot_i
                        elif choice.startswith("V"):
                            try:
                                vid = int(choice.split()[0][1:])
                            except (IndexError, ValueError):
                                continue
                            new_e = vid
                            ext_map[old_e] = new_e
                            vf_match = next((v for v in self._virtual if v["vid"] == vid), None)
                            vf_for_ext[new_e] = vf_match

                    base_cfg = {"type": "filament", "from": "User", "instantiation": "true",
                                "filament_type": ["PLA"], "filament_diameter": ["1.75"],
                                "filament_flow_ratio": ["1"], "compatible_printers": []}

                    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
                        for item in zin.infolist():
                            data = zin.read(item.filename)
                            fname = item.filename
                            if fname == "Metadata/model_settings.config":
                                try:
                                    root2 = ET.fromstring(data.decode("utf-8"))
                                    for obj in root2.findall("object"):
                                        for meta in obj.findall("metadata"):
                                            if meta.get("key") == "extruder":
                                                old_e = int(meta.get("value", "1"))
                                                if old_e in ext_map:
                                                    meta.set("value", str(ext_map[old_e]))
                                    data = ET.tostring(root2, encoding="unicode").encode("utf-8")
                                except Exception:
                                    pass
                            zout.writestr(item, data)

                        for new_e, vf in vf_for_ext.items():
                            if vf is None:
                                continue
                            sim_hex = vf.get("sim_hex", "#888888")
                            if not sim_hex.startswith("#"):
                                sim_hex = "#" + sim_hex
                            cfg = dict(base_cfg)
                            cfg["filament_settings_id"] = [f"U1-V{vf['vid']}"]
                            cfg["filament_vendor"] = ["U1 FullSpectrum"]
                            cfg["default_filament_colour"] = [sim_hex]
                            seq = vf["sequence"]
                            n_f = _seq_filament_count(seq)
                            cad = calc_cadence(seq, lh)
                            ids_s = sorted(cad.keys())
                            if n_f == 1:
                                note = f"U1 FullSpectrum | Pure T{seq[0]} | dE={vf.get('de',0):.1f}"
                            elif n_f == 2:
                                a = round(cad.get(ids_s[0], lh), 3)
                                b = round(cad.get(ids_s[1], lh) if len(ids_s) > 1 else lh, 3)
                                note = f"U1 FullSpectrum | Seq:{seq} | Cadence A={a}mm B={b}mm | dE={vf.get('de',0):.1f}"
                            else:
                                note = f"U1 FullSpectrum | Seq:{seq} | Pattern | dE={vf.get('de',0):.1f}"
                            cfg["filament_notes"] = [note]
                            zout.writestr(f"Metadata/filament_settings_{new_e}.config",
                                          json.dumps(cfg, indent=4).encode("utf-8"))

                        for new_e, slot_i in slot_for_ext.items():
                            s = self._slots[slot_i]
                            sim_hex = s["hex_edit"].text().strip() or "#888888"
                            if not sim_hex.startswith("#"):
                                sim_hex = "#" + sim_hex
                            cfg = dict(base_cfg)
                            cfg["filament_settings_id"] = [f"U1-T{slot_i+1}"]
                            cfg["filament_vendor"] = [s["brand_combo"].currentText() or "U1"]
                            cfg["default_filament_colour"] = [sim_hex]
                            cfg["filament_notes"] = [f"U1 FullSpectrum | T{slot_i+1} | TD={s['td_spin'].value()}"]
                            zout.writestr(f"Metadata/filament_settings_{new_e}.config",
                                          json.dumps(cfg, indent=4).encode("utf-8"))

                shutil.move(tmp_path, save_path)
                QMessageBox.information(dlg, self.t("dlg_saved"),
                                        self.t("3mf_write_ok", path=save_path))
                dlg.accept()
            except Exception as e:
                QMessageBox.critical(dlg, self.t("dlg_error"), str(e))

        btn_row = QHBoxLayout()
        write_btn = QPushButton(self.t("remap_btn_write"))
        write_btn.setObjectName("btn_primary")
        write_btn.setFixedHeight(40)
        write_btn.clicked.connect(do_write)
        cancel_btn = QPushButton(self.t("exp_cancel"))
        cancel_btn.setFixedHeight(36)
        cancel_btn.clicked.connect(dlg.reject)
        btn_row.addWidget(write_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)
        dlg.exec()

    # ── FULLSPECTRUM DIRECT 3MF EXPORT ────────────────────────────────────────

    def _open_fs_export_dialog(self):
        """Open the FullSpectrum Direct 3MF Export dialog (singleton)."""
        if self._fs_export_dlg is not None and self._fs_export_dlg.isVisible():
            self._fs_export_dlg.raise_()
            self._fs_export_dlg.activateWindow()
            return
        dlg = FullSpectrumExportDialog(self)
        self._fs_export_dlg = dlg
        dlg.exec()

    # ── EXPORT HINTS ─────────────────────────────────────────────────────────

    def _show_export_hint_for_seq(self, seq):
        """Show post-export slicer hint in the status bar."""
        if not seq:
            return
        lh = self._lh_spin.value() if hasattr(self, "_lh_spin") else 0.08
        n_f = _seq_filament_count(seq)
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
            self._set_status(hint, 8000)

    def _show_export_hint_for_virtual(self):
        """Show post-export slicer hint for first virtual head."""
        if not self._virtual:
            return
        vf = self._virtual[0]
        self._show_export_hint_for_seq(vf.get("sequence", ""))

    # ── AUTO SUGGESTION ───────────────────────────────────────────────────────

    def _check_auto_suggestion(self, target_hex):
        """Show status hint if a library filament is close to the target color."""
        target_lab = rgb_to_lab(hex_to_rgb(target_hex.lstrip("#")))
        best_de = float('inf')
        best_fil = None
        best_brand = None
        loaded_hexes = {s["hex_edit"].text().strip().lstrip("#").upper()
                        for s in self._slots}
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
        if best_fil and best_de < 15:
            msg = self.t("auto_suggest_tip",
                         brand=best_brand, name=best_fil.get("name", ""), de=best_de)
            self._set_status(msg, 8000)

    # ── MATERIAL COMPATIBILITY ────────────────────────────────────────────────

    def _check_material_compatibility(self, sequence_slot_indices):
        """Return a warning string if incompatible materials are mixed, else None."""
        used_slots = set(sequence_slot_indices)
        types_used = set()
        for idx in used_slots:
            if idx < 0 or idx >= len(self._slots):
                continue
            slot_name = self._slots[idx]["fil_combo"].currentText()
            slot_brand = self._slots[idx]["brand_combo"].currentText()
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
            if fil_type:
                types_used.add(fil_type)
        if len(types_used) < 2:
            return None
        incompatible_pairs = {("PLA", "ABS"), ("ABS", "PLA"),
                               ("PLA", "ASA"), ("ASA", "PLA"),
                               ("ABS", "PETG"), ("PETG", "ABS")}
        for a, b in incompatible_pairs:
            if a in types_used and b in types_used:
                return self.t("material_warn", a=a, b=b)
        return None

    # ── SLOT COMPARE ─────────────────────────────────────────────────────────

    def open_slot_compare(self):
        """Opens a dialog showing ΔE impact of replacing a slot with an alternative filament."""
        if not self._virtual:
            QMessageBox.information(self, self.t("dlg_note"), self.t("orca_no_virtual"))
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("slot_compare_title"))
        dlg.resize(700, 520)
        layout = QVBoxLayout(dlg)

        title_lbl = QLabel(self.t("slot_compare_title"))
        title_lbl.setStyleSheet("font-size: 13px; font-weight: bold; color: #38bdf8;")
        layout.addWidget(title_lbl)

        # Controls row
        ctrl_row = QHBoxLayout()

        slot_lbl = QLabel(self.t("slot_compare_slot"))
        ctrl_row.addWidget(slot_lbl)
        from PySide6.QtWidgets import QComboBox as _QCB
        slot_combo = _QCB()
        slot_combo.addItems([f"T{i+1}" for i in range(4)])
        slot_combo.setFixedWidth(70)
        ctrl_row.addWidget(slot_combo)

        ctrl_row.addSpacing(16)
        alt_lbl = QLabel(self.t("slot_compare_alt"))
        ctrl_row.addWidget(alt_lbl)
        all_fils = [(brand, fil) for brand, fils in self.library.items() for fil in fils]
        alt_combo = _QCB()
        alt_opts = [f"{b} — {f['name']}" for b, f in all_fils] if all_fils else ["(none)"]
        alt_combo.addItems(alt_opts)
        alt_combo.setMinimumWidth(240)
        ctrl_row.addWidget(alt_combo, 1)

        compare_btn = QPushButton(self.t("compare_run_btn"))
        compare_btn.setFixedHeight(36)
        ctrl_row.addWidget(compare_btn)
        layout.addLayout(ctrl_row)

        # Results table
        table = QTableWidget(0, 4)
        table.setHorizontalHeaderLabels([
            self.t("slot_compare_col_vid"),
            self.t("slot_compare_col_cur"),
            self.t("slot_compare_col_new"),
            self.t("slot_compare_col_delta"),
        ])
        table.horizontalHeader().setStretchLastSection(True)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(table, 1)

        def run_compare():
            table.setRowCount(0)
            slot_idx = int(slot_combo.currentText()[1]) - 1
            alt_label = alt_combo.currentText()
            alt_fil = next((f for b, f in all_fils
                            if f"{b} — {f['name']}" == alt_label), None)
            if alt_fil is None:
                return

            fils_current = self._slot_filaments()
            alt_hex = alt_fil.get("hex", "#888888")
            alt_td = alt_fil.get("td", DEFAULT_TD)
            alt_lab = rgb_to_lab(hex_to_rgb(alt_hex))

            fils_alt = []
            for f in fils_current:
                if f["id"] == slot_idx + 1:
                    fils_alt.append({"id": f["id"], "hex": alt_hex,
                                     "td": alt_td, "lab": alt_lab})
                else:
                    fils_alt.append(f)

            for vf in self._virtual:
                seq = vf.get("sequence", "")
                target_hex = vf.get("target_hex", "")
                if not target_hex:
                    continue
                t_lab = rgb_to_lab(hex_to_rgb(target_hex))
                cur_de = vf.get("de", 0.0)
                seq_ids = [int(c) for c in seq]
                new_sim = self._simulate_mix(seq_ids, fils_alt)
                new_de = delta_e(new_sim, t_lab)
                delta = new_de - cur_de

                row = table.rowCount()
                table.insertRow(row)
                table.setItem(row, 0, QTableWidgetItem(f"V{vf['vid']}"))
                cur_item = QTableWidgetItem(f"{cur_de:.1f}")
                cur_item.setForeground(QColor(_de_color(cur_de)))
                table.setItem(row, 1, cur_item)
                new_item = QTableWidgetItem(f"{new_de:.1f}")
                new_item.setForeground(QColor(_de_color(new_de)))
                table.setItem(row, 2, new_item)
                delta_col = "#4ade80" if delta < 0 else "#f87171" if delta > 0 else "#94a3b8"
                delta_txt = f"+{delta:.1f}" if delta > 0 else f"{delta:.1f}"
                delta_item = QTableWidgetItem(delta_txt)
                delta_item.setForeground(QColor(delta_col))
                table.setItem(row, 3, delta_item)

        compare_btn.clicked.connect(run_compare)

        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.reject)
        layout.addWidget(close_btn)
        dlg.exec()

    # ── PRINT STATISTICS ──────────────────────────────────────────────────────

    def _open_print_stats(self):
        """Filament usage statistics for all virtual heads."""
        if not self._virtual:
            QMessageBox.information(self, self.t("dlg_note"), self.t("orca_no_virtual"))
            return
        from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
            QLabel, QPushButton, QTableWidget, QTableWidgetItem,
            QDoubleSpinBox, QAbstractItemView)
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("print_stats_title"))
        dlg.resize(900, 480)
        layout = QVBoxLayout(dlg)
        layout.addWidget(QLabel(self.t("print_stats_desc")))

        h_row = QHBoxLayout()
        h_row.addWidget(QLabel(self.t("print_stats_height")))
        h_spin = QDoubleSpinBox()
        h_spin.setRange(0.0, 9999.0)
        h_spin.setValue(100.0)
        h_spin.setDecimals(1)
        h_spin.setSuffix(" mm")
        h_row.addWidget(h_spin)
        h_row.addStretch()
        layout.addLayout(h_row)

        # Columns: Head | Label | Seq | T1% | T2% | T3% | T4% | T1 mm | T2 mm | T3 mm | T4 mm | ΔE
        col_headers = [
            self.t("print_stats_col_head"),
            self.t("de_overview_col_label"),
            self.t("print_stats_col_seq"),
            "T1 %", "T2 %", "T3 %", "T4 %",
            "T1 mm", "T2 mm", "T3 mm", "T4 mm",
            self.t("print_stats_col_de"),
        ]
        table = QTableWidget(0, len(col_headers))
        table.setHorizontalHeaderLabels(col_headers)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)  # stretch seq col
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(table, 1)

        fils = self._slot_filaments()
        fil_colors = {str(f["id"]): f.get("hex", "#808080") for f in fils}

        def _refresh_fixed():
            total_h = h_spin.value()
            lh = self._lh_spin.value()
            total_layers = round(total_h / lh) if lh > 0 else 0
            table.setRowCount(0)
            for vf in self._virtual:
                seq = vf.get("sequence", "")
                if not seq:
                    continue
                seq_len = len(seq)
                row = table.rowCount()
                table.insertRow(row)
                table.setItem(row, 0, QTableWidgetItem(f"V{vf['vid']}"))
                table.setItem(row, 1, QTableWidgetItem(vf.get("label", "")))
                table.setItem(row, 2, QTableWidgetItem(seq))
                for i, slot_id in enumerate(["1", "2", "3", "4"]):
                    cnt = seq.count(slot_id)
                    pct = cnt / seq_len * 100
                    pct_item = QTableWidgetItem(f"{pct:.0f}%" if cnt > 0 else "—")
                    if cnt > 0:
                        hx = fil_colors.get(slot_id, "#808080")
                        try:
                            r, g, b = hex_to_rgb(hx)
                            pct_item.setBackground(QColor(r, g, b, 120))
                        except Exception:
                            pass
                    table.setItem(row, 3 + i, pct_item)
                    mm_val = pct / 100 * total_layers * lh if total_layers > 0 else 0
                    mm_item = QTableWidgetItem(f"{mm_val:.0f}" if cnt > 0 else "—")
                    table.setItem(row, 7 + i, mm_item)
                de = vf.get("de", 0.0)
                de_item = QTableWidgetItem(f"{de:.1f}")
                de_item.setForeground(QColor(_de_color(de)))
                table.setItem(row, 11, de_item)

        h_spin.valueChanged.connect(_refresh_fixed)
        _refresh_fixed()

        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.accept)
        layout.addWidget(close_btn)
        dlg.exec()

    # ── LAYER SEQUENCE PREVIEW ────────────────────────────────────────────────

    def _open_layer_preview(self):
        """Visual layer-by-layer color strip preview."""
        from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
            QLabel, QPushButton, QSpinBox, QScrollArea, QWidget, QComboBox)
        from PySide6.QtGui import QPainter, QColor as _QColor
        from PySide6.QtCore import Qt as _Qt

        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("layer_preview_title"))
        dlg.resize(680, 560)
        layout = QVBoxLayout(dlg)

        # Controls
        ctrl_row = QHBoxLayout()

        # Virtual head selector
        head_combo = QComboBox()
        head_combo.addItem(self.t("sec1_title") if hasattr(self, "_result") else "Calculator", "calc")
        for vf in self._virtual:
            head_combo.addItem(f"V{vf['vid']}  {vf.get('label', '')}", f"v{vf['vid']}")
        ctrl_row.addWidget(QLabel(self.t("head_label")))
        ctrl_row.addWidget(head_combo)

        ctrl_row.addSpacing(12)
        ctrl_row.addWidget(QLabel(self.t("layer_preview_layers")))
        n_spin = QSpinBox()
        n_spin.setRange(4, 200)
        n_spin.setValue(40)
        n_spin.setSingleStep(4)
        ctrl_row.addWidget(n_spin)
        ctrl_row.addStretch()
        layout.addLayout(ctrl_row)

        # Canvas widget
        class LayerCanvas(QWidget):
            def __init__(self, parent=None):
                super().__init__(parent)
                self.layers = []  # list of (r,g,b)
                self.setMinimumHeight(300)
            def set_layers(self, layers):
                self.layers = layers
                self.update()
            def paintEvent(self, event):
                if not self.layers:
                    return
                p = QPainter(self)
                w = self.width()
                h = self.height()
                n = len(self.layers)
                lh = max(1, h // n)
                for i, (r, g, b) in enumerate(self.layers):
                    y = i * lh
                    p.fillRect(0, y, w, lh, _QColor(r, g, b))
                p.end()

        canvas = LayerCanvas()
        scroll = QScrollArea()
        scroll.setWidget(canvas)
        scroll.setWidgetResizable(True)
        layout.addWidget(scroll, 1)

        fils = self._slot_filaments()
        by_id = {str(f["id"]): f for f in fils}

        def _get_seq_and_fils():
            key = head_combo.currentData()
            if key == "calc":
                seq = getattr(self, "_last_sequence", None)
                if not seq:
                    return [], fils
                return list(seq), fils
            else:
                vid = int(key[1:])
                vf = next((v for v in self._virtual if v["vid"] == vid), None)
                if vf is None:
                    return [], fils
                return list(vf.get("sequence", "")), fils

        def _refresh():
            seq, cur_fils = _get_seq_and_fils()
            if not seq:
                canvas.set_layers([])
                return
            n = n_spin.value()
            layers = []
            cur_by_id = {str(f["id"]): f for f in cur_fils}
            for i in range(n):
                fid = seq[i % len(seq)]
                fil = cur_by_id.get(str(fid), {})
                hx = fil.get("hex", "#808080")
                try:
                    r, g, b = hex_to_rgb(hx)
                except Exception:
                    r, g, b = 128, 128, 128
                layers.append((r, g, b))
            canvas.set_layers(layers)
            canvas.setMinimumHeight(max(300, n * 8))

        head_combo.currentIndexChanged.connect(_refresh)
        n_spin.valueChanged.connect(_refresh)
        _refresh()

        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.accept)
        layout.addWidget(close_btn)
        dlg.exec()

    # ── COPY ALL CADENCE ──────────────────────────────────────────────────────

    def _open_copy_all_cadence(self):
        if not self._virtual:
            QMessageBox.information(self, self.t("dlg_note"), self.t("orca_no_virtual"))
            return
        lh = self._lh_spin.value()
        lines = []
        for vf in self._virtual:
            seq = vf["sequence"]
            n_f = _seq_filament_count(seq)
            raw = "".join(seq)
            cad = calc_cadence(seq, lh)
            ids = sorted(cad.keys())
            label = vf.get("label", f"V{vf['vid']}")
            if n_f == 1:
                hint = f"Pure T{seq[0]}"
            elif n_f == 2:
                a = round(cad.get(ids[0], lh), 3)
                b = round(cad.get(ids[1], lh) if len(ids) > 1 else lh, 3)
                hint = f"Cadence A={a}mm  B={b}mm  |  Step={lh}mm"
            else:
                hint = f"Pattern: {raw}  |  Step={lh}mm"
            de = vf.get("de", 0.0)
            lines.append(f"V{vf['vid']} [{label}]  Seq:{raw}  dE:{de:.1f}  ->  {hint}")

        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("copy_all_title"))
        dlg.resize(720, 380)
        layout = QVBoxLayout(dlg)
        layout.addWidget(QLabel(self.t("copy_all_title")))
        txt = QTextEdit()
        txt.setFont(QFont("Courier New", 10))
        full = "\n".join(lines)
        txt.setPlainText(full)
        txt.setReadOnly(True)
        layout.addWidget(txt, 1)

        copy_btn = QPushButton(self.t("copy_all_btn"))
        copy_btn.setObjectName("btn_primary")
        copy_btn.clicked.connect(lambda: QApplication.clipboard().setText(full))
        layout.addWidget(copy_btn)
        dlg.exec()

    # ── ΔE OVERVIEW ───────────────────────────────────────────────────────────

    def _open_de_overview(self):
        if not self._virtual:
            QMessageBox.information(self, self.t("dlg_note"), self.t("orca_no_virtual"))
            return
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("de_overview_title"))
        dlg.resize(820, 480)
        layout = QVBoxLayout(dlg)
        layout.addWidget(QLabel(self.t("de_overview_title")))

        table = QTableWidget(len(self._virtual), 7)
        table.setHorizontalHeaderLabels([
            self.t("de_overview_col_id"),
            self.t("de_overview_col_label"),
            self.t("de_overview_col_seq"),
            "Target", "Sim",
            self.t("de_overview_col_de"),
            self.t("de_overview_col_quality"),
        ])
        table.horizontalHeader().setStretchLastSection(True)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        for row, vf in enumerate(self._virtual):
            seq = vf["sequence"]
            de = vf.get("de", 0.0)
            if de < DE_GOOD:
                q_text = "✓ " + self.t("de_quality_excellent")
                q_col = QColor("#4ade80")
            elif de < DE_OK:
                q_text = "~ " + self.t("de_quality_good")
                q_col = QColor("#fbbf24")
            else:
                q_text = "✗ " + self.t("de_quality_visible")
                q_col = QColor("#f87171")

            table.setItem(row, 0, QTableWidgetItem(f"V{vf['vid']}"))
            table.setItem(row, 1, QTableWidgetItem(vf.get("label", "")))
            table.setItem(row, 2, QTableWidgetItem(seq))

            tgt_item = QTableWidgetItem("")
            try:
                r, g, b = hex_to_rgb(vf["target_hex"])
                tgt_item.setBackground(QColor(r, g, b))
            except Exception:
                pass
            table.setItem(row, 3, tgt_item)

            sim_item = QTableWidgetItem("")
            try:
                r, g, b = hex_to_rgb(vf["sim_hex"])
                sim_item.setBackground(QColor(r, g, b))
            except Exception:
                pass
            table.setItem(row, 4, sim_item)

            de_item = QTableWidgetItem(f"{de:.1f}")
            de_item.setForeground(QColor(_de_color(de)))
            table.setItem(row, 5, de_item)

            q_item = QTableWidgetItem(q_text)
            q_item.setForeground(q_col)
            table.setItem(row, 6, q_item)

        layout.addWidget(table, 1)
        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.accept)
        layout.addWidget(close_btn)
        dlg.exec()

    # ── RECIPE EXPORT ─────────────────────────────────────────────────────────

    def _open_recipe_export(self):
        if not self._virtual:
            QMessageBox.information(self, self.t("dlg_note"), self.t("orca_no_virtual"))
            return
        lh = self._lh_spin.value()
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        lines = [
            "=" * 60,
            "  U1 FullSpectrum — Color Recipe",
            f"  Created: {now}   Layer height: {lh} mm",
            "=" * 60, "",
        ]
        for i, s in enumerate(self._slots):
            lines.append(f"  T{i+1}: {s['brand_combo'].currentText()} "
                         f"{s['fil_combo'].currentText()} ({s['hex_edit'].text()})")
        lines += ["", "-" * 60, ""]

        for vf in self._virtual:
            seq = vf["sequence"]
            de = vf.get("de", 0.0)
            n_f = _seq_filament_count(seq)
            runs = seq_to_runs(seq)
            cad = calc_cadence(seq, lh)
            ids = sorted(cad.keys())
            label = vf.get("label", f"V{vf['vid']}")
            q_str = ("excellent" if de < DE_GOOD else "good" if de < DE_OK else "visible") \
                if self.lang == "en" else \
                ("ausgezeichnet" if de < DE_GOOD else "gut" if de < DE_OK else "sichtbar")
            lines.append(f"V{vf['vid']}  \"{label}\"")
            lines.append(f"  Target: {vf['target_hex']}   Simulated: {vf['sim_hex']}   dE = {de:.1f} ({q_str})")
            lines.append(f"  Sequence: {''.join(seq)}   ({len(seq)} layers/cycle)")
            for fid, cnt in runs:
                lines.append(f"    T{fid}  x {cnt} layer{'s' if cnt > 1 else ''}")
            if n_f == 1:
                lines.append("  -> Pure color — no dithering needed")
            elif n_f == 2 and ids:
                a = round(cad.get(ids[0], lh), 3)
                b = round(cad.get(ids[1], lh) if len(ids) > 1 else lh, 3)
                lines.append(f"  -> Cadence A = {a} mm   B = {b} mm   Step = {lh} mm")
            else:
                lines.append(f"  -> Pattern Mode: {''.join(seq)}   Step = {lh} mm")
            lines.append("")
        lines.append("=" * 60)
        full_text = "\n".join(lines)

        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("recipe_title"))
        dlg.resize(700, 500)
        layout = QVBoxLayout(dlg)
        layout.addWidget(QLabel(self.t("recipe_title")))
        txt = QTextEdit()
        txt.setFont(QFont("Courier New", 10))
        txt.setPlainText(full_text)
        txt.setReadOnly(True)
        layout.addWidget(txt, 1)

        copy_btn = QPushButton(self.t("recipe_copy_btn"))
        copy_btn.clicked.connect(lambda: QApplication.clipboard().setText(full_text))
        layout.addWidget(copy_btn)
        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.accept)
        layout.addWidget(close_btn)
        dlg.exec()

    # ── GRADIENT DIALOG ────────────────────────────────────────────────────────

    def _open_gradient_dialog(self):
        from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
            QLabel, QPushButton, QSpinBox, QColorDialog, QFrame, QScrollArea,
            QWidget, QMessageBox)
        from PySide6.QtGui import QColor
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("gradient_title"))
        dlg.resize(540, 420)
        lay = QVBoxLayout(dlg)

        # Color stops row
        stop_count_row = QHBoxLayout()
        stop_count_row.addWidget(QLabel(self.t("grad_stops")))
        stop_spin = QSpinBox()
        stop_spin.setRange(2, 8)
        stop_spin.setValue(3)
        stop_count_row.addWidget(stop_spin)
        stop_count_row.addStretch()
        lay.addLayout(stop_count_row)

        stop_colors = ["#FF0000", "#00FF00", "#0000FF"]
        stop_btns = []
        btn_row = QHBoxLayout()

        def _pick(i):
            c = QColorDialog.getColor(QColor(stop_colors[i]), dlg)
            if c.isValid():
                stop_colors[i] = c.name().upper()
                stop_btns[i].setStyleSheet(f"background:{stop_colors[i]};border-radius:6px;")

        def _rebuild_btns():
            while btn_row.count():
                item = btn_row.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            stop_btns.clear()
            for i in range(stop_spin.value()):
                if i >= len(stop_colors):
                    stop_colors.append("#FFFFFF")
                b = QPushButton()
                b.setFixedSize(48, 32)
                b.setStyleSheet(f"background:{stop_colors[i]};border-radius:6px;")
                b.clicked.connect(lambda checked=False, idx=i: _pick(idx))
                btn_row.addWidget(b)
                stop_btns.append(b)

        _rebuild_btns()
        stop_spin.valueChanged.connect(lambda v: _rebuild_btns())
        lay.addWidget(QLabel(self.t("grad_pick_stops")))
        lay.addLayout(btn_row)

        # Steps
        step_row = QHBoxLayout()
        step_row.addWidget(QLabel(self.t("gradient_steps")))
        step_spin = QSpinBox()
        step_spin.setRange(3, 30)
        step_spin.setValue(8)
        step_row.addWidget(step_spin)
        step_row.addStretch()
        lay.addLayout(step_row)

        # Preview area
        preview_frame = QFrame()
        preview_frame.setFixedHeight(60)
        preview_frame.setFrameStyle(QFrame.StyledPanel)
        lay.addWidget(preview_frame)

        result_colors = []

        def _generate():
            stops = stop_colors[:stop_spin.value()]
            n = step_spin.value()
            result_colors.clear()
            segs = len(stops) - 1
            per_seg = max(1, n // segs)
            for s in range(segs):
                r1, g1, b1 = hex_to_rgb(stops[s])
                r2, g2, b2 = hex_to_rgb(stops[s + 1])
                pts = per_seg if s < segs - 1 else (n - len(result_colors))
                for i in range(pts):
                    t = i / max(pts - 1, 1)
                    r = round(r1 + (r2 - r1) * t)
                    g = round(g1 + (g2 - g1) * t)
                    b = round(b1 + (b2 - b1) * t)
                    result_colors.append(rgb_to_hex(r, g, b))
            # Draw preview
            w = preview_frame.width() or 480
            block = max(1, w // len(result_colors))
            css_parts = []
            for idx, hx in enumerate(result_colors):
                pct_start = round(idx * 100 / len(result_colors))
                pct_end = round((idx + 1) * 100 / len(result_colors))
                css_parts.append(f"{hx} {pct_start}%, {hx} {pct_end}%")
            grad_css = f"background: linear-gradient(to right, {', '.join(css_parts)});"
            preview_frame.setStyleSheet(grad_css)

        btn_gen = QPushButton(self.t("gradient_btn_calc"))
        btn_gen.clicked.connect(_generate)
        lay.addWidget(btn_gen)

        def _add_as_virtual():
            if not result_colors:
                QMessageBox.information(dlg, "", self.t("grad_generate_first"))
                return
            added = 0
            for hx in result_colors:
                if len(self._virtual) >= MAX_VIRTUAL:
                    break
                name = f"Gradient {len(self._virtual)+1}"
                self._virtual.append({"hex": hx, "label": name, "seq": [], "de": 0.0})
                added += 1
            self._recalc_all_virtual()
            self._refresh_virtual_grid()
            QMessageBox.information(dlg, "", self.t("gradient_done").format(n=added))

        row = QHBoxLayout()
        btn_add = QPushButton(self.t("grad_add_virtual"))
        btn_add.clicked.connect(_add_as_virtual)
        row.addWidget(btn_add)
        btn_close = QPushButton(self.t("exp_cancel"))
        btn_close.clicked.connect(dlg.accept)
        row.addWidget(btn_close)
        lay.addLayout(row)
        dlg.exec()

    # ── MULTI-GRADIENT DIALOG (Change 7) ──────────────────────────────────────

    def _open_multi_gradient_dialog(self):
        """Multi-Gradient: weighted blend of all 4 slots as virtual head."""
        fils = self._slot_filaments()
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("multi_gradient_title"))
        dlg.resize(440, 380)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        title_lbl = QLabel(self.t("multi_gradient_title"))
        title_lbl.setStyleSheet("font-size: 14pt; font-weight: bold;")
        layout.addWidget(title_lbl)
        desc_lbl = QLabel(self.t("multi_gradient_desc"))
        desc_lbl.setStyleSheet("color: #64748b;")
        layout.addWidget(desc_lbl)

        weight_spins = []
        for f in fils:
            row = QHBoxLayout()
            sw = QLabel("")
            sw.setFixedSize(28, 28)
            sw.setStyleSheet(f"background-color: {f['hex']}; border-radius: 4px;")
            row.addWidget(sw)
            row.addWidget(QLabel(f"T{f['id']}"))
            spin = QDoubleSpinBox()
            spin.setRange(0, 100)
            spin.setValue(25)
            spin.setSuffix("%")
            spin.setFixedWidth(80)
            weight_spins.append(spin)
            row.addWidget(spin)
            row.addStretch()
            layout.addLayout(row)

        preview_lbl = QLabel("")
        preview_lbl.setFixedSize(80, 32)
        preview_lbl.setStyleSheet("background-color: #808080; border-radius: 6px;")
        layout.addWidget(preview_lbl)

        def _update_preview():
            try:
                ws = [(f["id"], spin.value()) for f, spin in zip(fils, weight_spins)]
                seq = build_weighted_gradient_sequence(ws, max_len=MAX_SEQ_LEN)
                if not seq:
                    return
                sim_lab = self._simulate_mix([str(x) for x in seq], fils)
                preview_lbl.setStyleSheet(
                    f"background-color: {lab_to_hex(sim_lab)}; border-radius: 6px;")
            except Exception:
                pass

        for spin in weight_spins:
            spin.valueChanged.connect(lambda _: _update_preview())
        _update_preview()

        def _auto_balance():
            n = len(weight_spins)
            base = 100.0 / n
            for spin in weight_spins:
                spin.setValue(base)

        auto_btn = QPushButton(self.t("multi_gradient_auto"))
        auto_btn.clicked.connect(_auto_balance)
        layout.addWidget(auto_btn)

        def _add_as_virtual():
            try:
                ws = [(f["id"], spin.value()) for f, spin in zip(fils, weight_spins)]
                seq = build_weighted_gradient_sequence(ws, max_len=MAX_SEQ_LEN)
                if not seq:
                    return
                seq_str = "".join(str(x) for x in seq)
                sim_lab = self._simulate_mix(list(seq_str), fils)
                sim_hex = lab_to_hex(sim_lab)
                result = {
                    "target_hex": sim_hex,
                    "sequence": seq_str,
                    "sim_hex": sim_hex,
                    "de": 0.0,
                    "seq_len": len(seq_str),
                }
                self.add_virtual(result)
                dlg.accept()
            except Exception as e:
                QMessageBox.critical(dlg, self.t("dlg_error"), str(e))

        add_btn = QPushButton(self.t("multi_gradient_add"))
        add_btn.setStyleSheet(
            "QPushButton { background-color: #15803d; color: white; font-size: 13px; "
            "font-weight: bold; border-radius: 6px; } "
            "QPushButton:hover { background-color: #16a34a; }")
        add_btn.setFixedHeight(42)
        add_btn.clicked.connect(_add_as_virtual)
        layout.addWidget(add_btn)

        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.accept)
        layout.addWidget(close_btn)
        dlg.exec()

    # ── HARMONIES DIALOG ───────────────────────────────────────────────────────

    def _open_harmonies_dialog(self):
        from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
            QLabel, QPushButton, QColorDialog, QComboBox, QScrollArea,
            QWidget, QGridLayout, QMessageBox)
        from PySide6.QtGui import QColor
        import colorsys

        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("harm_title"))
        dlg.resize(500, 480)
        lay = QVBoxLayout(dlg)

        base_color = [self._target_hex if hasattr(self, "_target_hex") else "#FF0000"]

        top_row = QHBoxLayout()
        top_row.addWidget(QLabel(self.t("harm_base")))
        base_btn = QPushButton()
        base_btn.setFixedSize(48, 28)
        base_btn.setStyleSheet(f"background:{base_color[0]};border-radius:4px;")

        harmony_combo = QComboBox()
        harmony_combo.addItems([
            self.t("harm_complement"),
            self.t("harm_split"),
            self.t("harm_triadic"),
            self.t("harm_analogous"),
            self.t("harm_tetradic"),
        ])
        top_row.addWidget(base_btn)
        top_row.addWidget(QLabel(self.t("harm_type")))
        top_row.addWidget(harmony_combo)
        top_row.addStretch()
        lay.addLayout(top_row)

        swatch_area = QWidget()
        swatch_lay = QGridLayout(swatch_area)
        lay.addWidget(swatch_area, 1)

        result_hex = []

        def _pick_base():
            c = QColorDialog.getColor(QColor(base_color[0]), dlg)
            if c.isValid():
                base_color[0] = c.name().upper()
                base_btn.setStyleSheet(f"background:{base_color[0]};border-radius:4px;")
                _update()

        base_btn.clicked.connect(_pick_base)

        def _hue_shift(hx, degrees):
            r, g, b = hex_to_rgb(hx)
            h, s, v = colorsys.rgb_to_hsv(r/255, g/255, b/255)
            h = (h + degrees / 360) % 1.0
            nr, ng, nb = colorsys.hsv_to_rgb(h, s, v)
            return rgb_to_hex(round(nr*255), round(ng*255), round(nb*255))

        def _update():
            result_hex.clear()
            while swatch_lay.count():
                item = swatch_lay.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            mode = harmony_combo.currentIndex()
            shifts = {
                0: [180],
                1: [150, 210],
                2: [120, 240],
                3: [30, 60, -30, -60],
                4: [90, 180, 270],
            }.get(mode, [180])
            colors = [base_color[0]] + [_hue_shift(base_color[0], s) for s in shifts]
            result_hex.extend(colors)
            for idx, hx in enumerate(colors):
                col = idx % 5
                row_i = idx // 5
                frame = QWidget()
                fl = QVBoxLayout(frame)
                fl.setContentsMargins(4, 4, 4, 4)
                sw = QLabel()
                sw.setFixedSize(60, 60)
                sw.setStyleSheet(f"background:{hx};border-radius:8px;border:2px solid #334155;")
                fl.addWidget(sw)
                fl.addWidget(QLabel(hx))
                swatch_lay.addWidget(frame, row_i, col)

        harmony_combo.currentIndexChanged.connect(lambda _: _update())
        _update()

        def _add_as_virtual():
            added = 0
            for hx in result_hex:
                if len(self._virtual) >= MAX_VIRTUAL:
                    break
                name = f"Harmony {len(self._virtual)+1}"
                self._virtual.append({"hex": hx, "label": name, "seq": [], "de": 0.0})
                added += 1
            self._recalc_all_virtual()
            self._refresh_virtual_grid()
            QMessageBox.information(dlg, "", self.t("harm_added").format(added))

        btn_row = QHBoxLayout()
        btn_add = QPushButton(self.t("harm_add_virtual"))
        btn_add.clicked.connect(_add_as_virtual)
        btn_row.addWidget(btn_add)
        btn_close = QPushButton(self.t("exp_cancel"))
        btn_close.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_close)
        lay.addLayout(btn_row)
        dlg.exec()

    # ── PALETTE FROM IMAGE ─────────────────────────────────────────────────────

    def _import_palette_from_image(self):
        from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
            QLabel, QPushButton, QSpinBox, QFileDialog, QScrollArea,
            QWidget, QGridLayout, QMessageBox)
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("pal_title"))
        dlg.resize(480, 420)
        lay = QVBoxLayout(dlg)

        path_label = QLabel(self.t("pal_no_file"))
        lay.addWidget(path_label)

        file_path = [None]
        extracted = []

        n_row = QHBoxLayout()
        n_row.addWidget(QLabel(self.t("pal_colors")))
        n_spin = QSpinBox()
        n_spin.setRange(2, 24)
        n_spin.setValue(8)
        n_row.addWidget(n_spin)
        n_row.addStretch()
        lay.addLayout(n_row)

        swatch_area = QScrollArea()
        swatch_widget = QWidget()
        swatch_lay = QGridLayout(swatch_widget)
        swatch_area.setWidget(swatch_widget)
        swatch_area.setWidgetResizable(True)
        lay.addWidget(swatch_area, 1)

        def _browse():
            f, _ = QFileDialog.getOpenFileName(
                dlg, self.t("pal_browse"), "",
                "Images (*.png *.jpg *.jpeg *.bmp *.webp *.tiff)")
            if f:
                file_path[0] = f
                path_label.setText(os.path.basename(f))
                _extract()

        def _extract():
            if not file_path[0]:
                return
            try:
                from PIL import Image
                img = Image.open(file_path[0])
                extracted.clear()
                # FIX 2: Handle indexed/palette mode PNG explicitly
                if img.mode == 'P':
                    palette = img.getpalette()
                    used_indices = set(img.getdata())
                    for idx in sorted(used_indices):
                        r = palette[idx * 3]
                        g = palette[idx * 3 + 1]
                        b = palette[idx * 3 + 2]
                        brightness = (r + g + b) / 3
                        if brightness < 5 or brightness > 250:
                            continue
                        extracted.append(f"#{r:02X}{g:02X}{b:02X}")
                else:
                    img = img.convert("RGB")
                    img.thumbnail((200, 200))
                    n_colors = min(n_spin.value(), 32)
                    quantized = img.quantize(colors=n_colors,
                                             method=Image.Quantize.MEDIANCUT)
                    palette_raw = quantized.getpalette()
                    for i in range(n_colors):
                        r, g, b = palette_raw[i * 3], palette_raw[i * 3 + 1], palette_raw[i * 3 + 2]
                        extracted.append(f"#{r:02X}{g:02X}{b:02X}")
            except Exception as e:
                QMessageBox.warning(dlg, "", str(e))
                return
            while swatch_lay.count():
                item = swatch_lay.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            for idx, hx in enumerate(extracted):
                col = idx % 6
                row_i = idx // 6
                sw = QLabel()
                sw.setFixedSize(56, 56)
                sw.setStyleSheet(f"background:{hx};border-radius:6px;border:2px solid #334155;")
                sw.setToolTip(hx)
                swatch_lay.addWidget(sw, row_i, col)

        browse_btn = QPushButton(self.t("pal_browse"))
        browse_btn.clicked.connect(_browse)
        lay.addWidget(browse_btn)
        n_spin.valueChanged.connect(lambda v: _extract())

        def _add_virtual():
            added = 0
            for hx in extracted:
                if len(self._virtual) >= MAX_VIRTUAL:
                    break
                name = f"Palette {len(self._virtual)+1}"
                self._virtual.append({"hex": hx, "label": name, "seq": [], "de": 0.0})
                added += 1
            self._recalc_all_virtual()
            self._refresh_virtual_grid()
            QMessageBox.information(dlg, "", self.t("pal_added").format(added))

        btn_row = QHBoxLayout()
        btn_add = QPushButton(self.t("pal_add_virtual"))
        btn_add.clicked.connect(_add_virtual)
        btn_row.addWidget(btn_add)
        btn_close = QPushButton(self.t("exp_cancel"))
        btn_close.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_close)
        lay.addLayout(btn_row)
        dlg.exec()

    # ── MULTI-TARGET OPTIMIZER ─────────────────────────────────────────────────

    def _open_multitarget_optimizer(self):
        from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
            QLabel, QPushButton, QColorDialog, QMessageBox, QScrollArea,
            QWidget, QFrame)
        from PySide6.QtGui import QColor
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("mt_title"))
        dlg.resize(520, 480)
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel(self.t("multi_desc")))

        targets = []
        tgt_widget = QWidget()
        tgt_lay = QVBoxLayout(tgt_widget)
        tgt_lay.setSpacing(4)
        scroll = QScrollArea()
        scroll.setWidget(tgt_widget)
        scroll.setWidgetResizable(True)
        lay.addWidget(scroll, 1)

        def _add_target():
            c = QColorDialog.getColor(QColor("#FF0000"), dlg)
            if not c.isValid():
                return
            hx = c.name().upper()
            targets.append(hx)
            row = QHBoxLayout()
            sw = QLabel()
            sw.setFixedSize(32, 32)
            sw.setStyleSheet(f"background:{hx};border-radius:4px;border:1px solid #334155;")
            lbl = QLabel(hx)
            row.addWidget(sw)
            row.addWidget(lbl)
            rm = QPushButton("✕")
            rm.setFixedSize(28, 28)

            def _rm(hx_=hx):
                if hx_ in targets:
                    targets.remove(hx_)
                _refresh_rows()

            rm.clicked.connect(_rm)
            row.addWidget(rm)
            row.addStretch()
            frame = QFrame()
            frame.setLayout(row)
            tgt_lay.addWidget(frame)

        def _refresh_rows():
            while tgt_lay.count():
                item = tgt_lay.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            for hx in targets:
                row = QHBoxLayout()
                sw = QLabel()
                sw.setFixedSize(32, 32)
                sw.setStyleSheet(f"background:{hx};border-radius:4px;border:1px solid #334155;")
                lbl = QLabel(hx)
                row.addWidget(sw)
                row.addWidget(lbl)
                frame = QFrame()
                frame.setLayout(row)
                tgt_lay.addWidget(frame)

        def _optimize():
            if not targets:
                QMessageBox.information(dlg, "", self.t("mt_no_targets"))
                return
            fils = self._slot_filaments()
            best_overall = None
            best_score = float("inf")
            for perm_len in range(1, MAX_SEQ_LEN + 1):
                from itertools import product
                for combo in product(range(len(fils)), repeat=perm_len):
                    seq = [fils[i]["id"] for i in combo]
                    total_de = 0.0
                    for hx in targets:
                        rgb_t = rgb_to_lab(hex_to_rgb(hx))
                        mixed = self._simulate_mix(seq, fils)
                        de = delta_e(rgb_t, mixed)
                        total_de += de
                    avg = total_de / len(targets)
                    if avg < best_score:
                        best_score = avg
                        best_overall = seq
                if best_score < DE_GOOD:
                    break
            if best_overall:
                seq_str = "".join(best_overall)
                QMessageBox.information(
                    dlg, self.t("multi_result"),
                    self.t("multi_best_seq").format(seq_str, round(best_score, 2))
                )
            else:
                QMessageBox.information(dlg, "", self.t("multi_no_result"))

        add_btn = QPushButton(self.t("multi_add_color"))
        add_btn.clicked.connect(_add_target)
        lay.addWidget(add_btn)

        btn_row = QHBoxLayout()
        opt_btn = QPushButton(self.t("multi_optimize"))
        opt_btn.clicked.connect(_optimize)
        btn_row.addWidget(opt_btn)
        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.accept)
        btn_row.addWidget(close_btn)
        lay.addLayout(btn_row)
        dlg.exec()

    # ── SLOT OPTIMIZER ─────────────────────────────────────────────────────────

    def _open_slot_optimizer(self):
        from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
            QLabel, QPushButton, QProgressBar, QTextEdit, QMessageBox)
        from PySide6.QtCore import QThread, Signal as Sig

        class _OptimizerWorker(QThread):
            progress = Sig(int)
            finished = Sig(str)

            def __init__(self, library, virtual, model, lh):
                super().__init__()
                self._lib = library
                self._virtual = virtual
                self._model = model
                self._lh = lh

            def run(self):
                targets = [v["hex"] for v in self._virtual]
                if not targets:
                    self.finished.emit(self.tr("No virtual heads to optimize for."))
                    return
                all_fils = self._lib
                n = len(all_fils)
                best_score = float("inf")
                best_combo = None
                slot_count = 4
                from itertools import combinations
                combos = list(combinations(range(n), slot_count))
                total = len(combos)
                for step, idx_combo in enumerate(combos):
                    if step % max(1, total // 100) == 0:
                        self.progress.emit(int(step * 100 / total))
                    fils = [all_fils[i] for i in idx_combo]
                    total_de = 0.0
                    for hx in targets:
                        rgb_t = hex_to_rgb(hx)
                        # Try all 1-4 length sequences
                        best_local = float("inf")
                        from itertools import product
                        for length in range(1, 5):
                            for seq in product(range(len(fils)), repeat=length):
                                ids = [fils[i]["id"] for i in seq]
                                mixed = self._simulate_mix_static(ids, fils)
                                de = delta_e(rgb_to_lab(rgb_t), rgb_to_lab(mixed))
                                if de < best_local:
                                    best_local = de
                        total_de += best_local
                    score = total_de / len(targets)
                    if score < best_score:
                        best_score = score
                        best_combo = fils
                self.progress.emit(100)
                if best_combo:
                    names = ", ".join(f["name"] for f in best_combo)
                    self.finished.emit(f"Best slot set (avg \u0394E {best_score:.2f}):\n{names}")
                else:
                    self.finished.emit("No result found.")

            def _simulate_mix_static(self, seq, fils_list):
                if not seq:
                    return (128, 128, 128)
                # Simple average for optimizer
                rs, gs, bs = 0, 0, 0
                n = len(seq)
                for fid in seq:
                    fil = next((f for f in fils_list if f["id"] == fid), None)
                    if fil:
                        r, g, b = hex_to_rgb(fil.get("hex", "#808080"))
                        rs += r; gs += g; bs += b
                return (rs // n, gs // n, bs // n)

        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("slotopt_title"))
        dlg.resize(480, 360)
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel(self.t("slotopt_desc")))

        progress = QProgressBar()
        progress.setRange(0, 100)
        progress.setValue(0)
        lay.addWidget(progress)

        result_txt = QTextEdit()
        result_txt.setReadOnly(True)
        result_txt.setFixedHeight(120)
        lay.addWidget(result_txt)

        worker_ref = [None]

        def _start():
            fils = self._build_lib_fils()
            model = self._model_combo.currentText() if hasattr(self, "_model_combo") else "linear"
            lh = self._lh_spin.value() if hasattr(self, "_lh_spin") else 0.2
            w = _OptimizerWorker(fils, self._virtual, model, lh)
            worker_ref[0] = w
            w.progress.connect(progress.setValue)
            w.finished.connect(lambda msg: result_txt.setPlainText(msg))
            result_txt.setPlainText(self.t("slotopt_running"))
            w.start()

        def _build_lib_fils(self=self):
            fils = []
            for brand, entries in self.library.items():
                for entry in entries:
                    fils.append({
                        "id": brand[:2],
                        "name": f"{brand} {entry.get('name', '')}",
                        "hex": entry.get("hex", "#808080"),
                        "td": entry.get("td", DEFAULT_TD),
                        "lab": rgb_to_lab(hex_to_rgb(entry.get("hex", "#808080"))),
                    })
            return fils

        self._build_lib_fils = _build_lib_fils

        btn_row = QHBoxLayout()
        start_btn = QPushButton(self.t("slotopt_start"))
        start_btn.clicked.connect(_start)
        btn_row.addWidget(start_btn)
        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.accept)
        btn_row.addWidget(close_btn)
        lay.addLayout(btn_row)
        dlg.exec()

    # ── TD CALIBRATION ─────────────────────────────────────────────────────────

    def _open_td_calibration(self):
        from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
            QLabel, QPushButton, QDoubleSpinBox, QLineEdit, QMessageBox)
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("tdcal_title"))
        dlg.resize(420, 340)
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel(self.t("tdcal_desc")))

        # Filament hex
        fil_row = QHBoxLayout()
        fil_row.addWidget(QLabel(self.t("tdcal_fil_hex")))
        fil_hex_edit = QLineEdit("#FF0000")
        fil_hex_edit.setFixedWidth(100)
        fil_row.addWidget(fil_hex_edit)
        fil_row.addStretch()
        lay.addLayout(fil_row)

        # Measured printed color
        meas_row = QHBoxLayout()
        meas_row.addWidget(QLabel(self.t("tdcal_measured")))
        meas_hex_edit = QLineEdit("#CC0000")
        meas_hex_edit.setFixedWidth(100)
        meas_row.addWidget(meas_hex_edit)
        meas_row.addStretch()
        lay.addLayout(meas_row)

        # Layer count
        layers_row = QHBoxLayout()
        layers_row.addWidget(QLabel(self.t("tdcal_layers")))
        layers_spin = QDoubleSpinBox()
        layers_spin.setRange(1, 100)
        layers_spin.setValue(4)
        layers_spin.setDecimals(0)
        layers_row.addWidget(layers_spin)
        layers_row.addStretch()
        lay.addLayout(layers_row)

        result_label = QLabel("")
        lay.addWidget(result_label)

        def _calc_td():
            try:
                r1, g1, b1 = hex_to_rgb(fil_hex_edit.text().strip())
                r2, g2, b2 = hex_to_rgb(meas_hex_edit.text().strip())
                layers = int(layers_spin.value())
                # Estimate TD: higher opacity = lower TD
                # Opacity = how much the base "shows through" vs white
                # Simple heuristic: TD ~ -log(1 - de_ratio) / layers
                de_fil = delta_e(rgb_to_lab((r1, g1, b1)), rgb_to_lab((255, 255, 255)))
                de_meas = delta_e(rgb_to_lab((r2, g2, b2)), rgb_to_lab((255, 255, 255)))
                if de_fil < 0.1:
                    result_label.setText(self.t("tdcal_white_warn"))
                    return
                ratio = de_meas / max(de_fil, 0.01)
                import math
                td_estimate = max(0.1, round(-math.log(max(1 - ratio, 0.01)) / max(layers, 1) * 10, 2))
                result_label.setText(self.t("tdcal_result").format(td_estimate))
            except Exception as e:
                result_label.setText(str(e))

        calc_btn = QPushButton(self.t("tdcal_calc"))
        calc_btn.clicked.connect(_calc_td)
        lay.addWidget(calc_btn)

        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.accept)
        lay.addWidget(close_btn)
        dlg.exec()

    # ── WEB UPDATE LIBRARY ────────────────────────────────────────────────────

    def _web_update_library(self):
        """Download community filament DB async — no UI freeze."""
        btn = None
        # find the button and disable it temporarily
        for child in self.findChildren(QPushButton):
            if "online" in child.text().lower() or "update" in child.text().lower():
                btn = child
                break

        self._set_status(self.t("web_update_downloading"))

        class _FetchWorker(QThread):
            done = Signal(dict)   # {"added": N, "error": str|None, "data": [...]}
            def __init__(self, url):
                super().__init__()
                self._url = url
            def run(self):
                import urllib.request
                try:
                    with urllib.request.urlopen(self._url, timeout=15) as resp:
                        data = json.loads(resp.read().decode("utf-8"))
                    self.done.emit({"added": -1, "error": None, "data": data})
                except Exception as e:
                    self.done.emit({"added": 0, "error": str(e), "data": []})

        def _on_done(result):
            if result["error"]:
                self._set_status("❌ " + result["error"], 6000)
                QMessageBox.warning(self, "Community DB", result["error"])
                return
            data = result["data"]
            added = 0
            if isinstance(data, dict):
                # format: {"BrandName": [{name, hex, td}, ...], ...}
                for brand, fils in data.items():
                    if not isinstance(fils, list):
                        continue
                    existing = {f.get("name","").lower() for f in self.library.get(brand, [])}
                    new_fils = [f for f in fils
                                if isinstance(f, dict) and "name" in f and "hex" in f
                                and f["name"].lower() not in existing]
                    if new_fils:
                        self.library.setdefault(brand, []).extend(new_fils)
                        added += len(new_fils)
            elif isinstance(data, list):
                # flat list format
                all_names = {f.get("name","").lower()
                            for fils in self.library.values() for f in fils}
                for entry in data:
                    if not isinstance(entry, dict) or "name" not in entry or "hex" not in entry:
                        continue
                    if entry["name"].lower() not in all_names:
                        brand = entry.get("brand", "Community")
                        self.library.setdefault(brand, []).append(entry)
                        added += 1
            self._save_db()
            msg = self.t("web_update_added", added=added)
            self._set_status(msg, 6000)
            QMessageBox.information(self, "Community DB", msg)

        worker = _FetchWorker(_COMMUNITY_URL)
        worker.done.connect(_on_done)
        # keep reference so GC doesn't kill it
        self._community_worker = worker
        worker.start()

    # ── LAB 3D PLOT ──────────────────────────────────────────────────────────

    def _show_lab_plot(self):
        from PySide6.QtWidgets import QDialog, QVBoxLayout
        try:
            import matplotlib
            matplotlib.use("Qt5Agg")
            from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
            import matplotlib.pyplot as plt
            from mpl_toolkits.mplot3d import Axes3D  # noqa
        except ImportError:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "", "matplotlib not installed.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("lab_plot_title"))
        dlg.resize(680, 560)
        lay = QVBoxLayout(dlg)

        fig = plt.figure(figsize=(7, 5))
        ax = fig.add_subplot(111, projection="3d")

        fils = self._slot_filaments()
        colors_hex = [f.get("hex", "#808080") for f in fils]
        for hx in colors_hex:
            L, a, bv = rgb_to_lab(hex_to_rgb(hx))
            ax.scatter([a], [bv], [L], c=hx, s=80, depthshade=False)

        # Also plot virtual heads
        for v in self._virtual:
            L, a, bv = rgb_to_lab(hex_to_rgb(v.get("hex", "#808080")))
            ax.scatter([a], [bv], [L], c=v.get("hex", "#808080"), s=120, marker="*", depthshade=False)

        ax.set_xlabel("a*")
        ax.set_ylabel("b*")
        ax.set_zlabel("L*")
        ax.set_title("CIE L*a*b* Color Space")

        canvas = FigureCanvasQTAgg(fig)
        lay.addWidget(canvas)

        from PySide6.QtWidgets import QPushButton
        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.accept)
        lay.addWidget(close_btn)
        plt.tight_layout()
        dlg.exec()
        plt.close(fig)

    # ── GAMUT PLOT ───────────────────────────────────────────────────────────

    def _open_gamut_plot(self):
        from PySide6.QtWidgets import QDialog, QVBoxLayout, QPushButton
        try:
            import matplotlib
            matplotlib.use("Qt5Agg")
            from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
            import matplotlib.pyplot as plt
        except ImportError:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "", "matplotlib not installed.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("gamut_plot_title"))
        dlg.resize(600, 520)
        lay = QVBoxLayout(dlg)

        fig, ax = plt.subplots(figsize=(6, 5))
        ax.set_xlim(-128, 128)
        ax.set_ylim(-128, 128)
        ax.set_xlabel("a*")
        ax.set_ylabel("b*")
        ax.set_title("Gamut (a*b* plane)")
        ax.axhline(0, color="gray", lw=0.5)
        ax.axvline(0, color="gray", lw=0.5)

        fils = self._slot_filaments()
        for f in fils:
            L, a, bv = rgb_to_lab(hex_to_rgb(f.get("hex", "#808080")))
            ax.scatter([a], [bv], c=f.get("hex", "#808080"), s=80, zorder=5)
            ax.annotate(f.get("name", "?")[:8], (a, bv), fontsize=7)

        for v in self._virtual:
            L, a, bv = rgb_to_lab(hex_to_rgb(v.get("hex", "#808080")))
            ax.scatter([a], [bv], c=v.get("hex", "#808080"), s=120, marker="*", zorder=6)

        canvas = FigureCanvasQTAgg(fig)
        lay.addWidget(canvas)

        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.accept)
        lay.addWidget(close_btn)
        plt.tight_layout()
        dlg.exec()
        plt.close(fig)

    # ── SAVE SWATCH ───────────────────────────────────────────────────────────

    def _save_swatch(self):
        from PySide6.QtWidgets import QFileDialog, QMessageBox
        path, _ = QFileDialog.getSaveFileName(
            self, self.t("swatch_save_title"), "swatch.png", "PNG (*.png)")
        if not path:
            return
        try:
            from PIL import Image, ImageDraw, ImageFont
            w, h = 400, 200
            img = Image.new("RGB", (w, h), (30, 30, 30))
            draw = ImageDraw.Draw(img)
            # Main swatch
            hx = getattr(self, "_target_hex", "#808080")
            r, g, b = hex_to_rgb(hx)
            draw.rectangle([10, 10, 190, 190], fill=(r, g, b))
            # Result swatch
            if hasattr(self, "_result_hex") and self._result_hex:
                rr, rg, rb = hex_to_rgb(self._result_hex)
                draw.rectangle([210, 10, 390, 190], fill=(rr, rg, rb))
            draw.text((10, 192), f"Target: {hx}", fill=(200, 200, 200))
            if hasattr(self, "_result_hex") and self._result_hex:
                draw.text((210, 192), f"Result: {self._result_hex}", fill=(200, 200, 200))
            img.save(path)
            QMessageBox.information(self, "", self.t("swatch_saved").format(path))
        except Exception as e:
            QMessageBox.warning(self, "", str(e))

    # ── EXPORT PNG SUMMARY ────────────────────────────────────────────────────

    def _export_png_summary(self):
        from PySide6.QtWidgets import QFileDialog, QMessageBox
        path, _ = QFileDialog.getSaveFileName(
            self, self.t("png_export_title"), "summary.png", "PNG (*.png)")
        if not path:
            return
        try:
            from PIL import Image, ImageDraw
            cols = 6
            rows = max(1, (len(self._virtual) + cols - 1) // cols)
            cell = 80
            w = cols * cell + 20
            h = rows * cell + 60
            img = Image.new("RGB", (w, h), (30, 30, 30))
            draw = ImageDraw.Draw(img)
            draw.text((10, 10), f"Virtual Heads ({len(self._virtual)})", fill=(220, 220, 220))
            for idx, v in enumerate(self._virtual):
                col = idx % cols
                row_i = idx // cols
                x = 10 + col * cell
                y = 40 + row_i * cell
                hx = v.get("hex", "#808080")
                r, g, b = hex_to_rgb(hx)
                draw.rectangle([x, y, x + cell - 4, y + cell - 20], fill=(r, g, b))
                lbl = v.get("label", hx)[:8]
                draw.text((x, y + cell - 18), lbl, fill=(200, 200, 200))
            img.save(path)
            QMessageBox.information(self, "", self.t("png_saved").format(path))
        except Exception as e:
            QMessageBox.warning(self, "", str(e))

    # ── SLICER GUIDE ──────────────────────────────────────────────────────────

    def _open_slicer_guide(self):
        from PySide6.QtWidgets import (QDialog, QVBoxLayout, QPushButton,
            QTextBrowser)
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("guide_title"))
        dlg.resize(600, 480)
        lay = QVBoxLayout(dlg)
        txt = QTextBrowser()
        txt.setOpenExternalLinks(True)
        guide_html = """
<h2>Snapmaker U1 Slicer Setup Guide</h2>
<h3>OrcaSlicer (Recommended)</h3>
<ol>
  <li>Open OrcaSlicer and load your model.</li>
  <li>Go to <b>Filament</b> settings and add one profile per slot (T1-T4).</li>
  <li>Assign filament colors matching your loaded filaments.</li>
  <li>In <b>Print Settings &rarr; Layer Height</b>, set the layer height used for calibration.</li>
  <li>Use <b>FullSpectrum</b> export from this tool to generate filament profiles.</li>
  <li>Import the generated profiles into OrcaSlicer via <i>File &rarr; Import Configs</i>.</li>
</ol>
<h3>Layer Dithering Settings</h3>
<ul>
  <li>Enable <b>Layer Dithering</b> in OrcaSlicer print settings.</li>
  <li>Set <b>Dithering Pattern</b> to match the sequence from the calculator.</li>
  <li>Cadence A and B values are shown in the recipe export.</li>
</ul>
<h3>Tips</h3>
<ul>
  <li>Print a test swatch at multiple layer heights to calibrate TD values.</li>
  <li>Use the TD Calibration tool to estimate your filament opacity.</li>
  <li>Dark colors require more layers (lower TD = more opaque).</li>
</ul>
"""
        txt.setHtml(guide_html)
        lay.addWidget(txt, 1)
        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.accept)
        lay.addWidget(close_btn)
        dlg.exec()

    # ── TC ESTIMATOR ──────────────────────────────────────────────────────────

    def _open_tc_estimator(self):
        from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
            QLabel, QPushButton, QDoubleSpinBox, QLineEdit)
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("tc_title"))
        dlg.resize(380, 260)
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel(self.t("tc_desc")))

        lh_row = QHBoxLayout()
        lh_row.addWidget(QLabel(self.t("layer_height")))
        lh_spin = QDoubleSpinBox()
        lh_spin.setRange(0.05, 1.0)
        lh_spin.setSingleStep(0.05)
        lh_spin.setDecimals(2)
        lh_spin.setValue(0.2)
        lh_row.addWidget(lh_spin)
        lh_row.addStretch()
        lay.addLayout(lh_row)

        td_row = QHBoxLayout()
        td_row.addWidget(QLabel(self.t("td_label")))
        td_spin = QDoubleSpinBox()
        td_spin.setRange(0.1, 50.0)
        td_spin.setSingleStep(0.5)
        td_spin.setDecimals(1)
        td_spin.setValue(DEFAULT_TD)
        td_row.addWidget(td_spin)
        td_row.addStretch()
        lay.addLayout(td_row)

        result_lbl = QLabel("")
        lay.addWidget(result_lbl)

        def _calc():
            lh = lh_spin.value()
            td = safe_td(td_spin.value())
            # Transmission coefficient per layer = TD / layer_height
            tc_per_layer = round(td / lh, 4) if lh > 0 else 0
            layers_to_opaque = round(td / lh) if lh > 0 else 0
            result_lbl.setText(
                self.t("tc_result").format(
                    tc=tc_per_layer, n=layers_to_opaque
                )
            )

        calc_btn = QPushButton(self.t("tdcal_calc"))
        calc_btn.clicked.connect(_calc)
        lay.addWidget(calc_btn)

        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.accept)
        lay.addWidget(close_btn)
        dlg.exec()

    # ── FILAMENT MATRIX ───────────────────────────────────────────────────────

    def _open_filament_matrix(self):
        from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
            QLabel, QPushButton, QTableWidget, QTableWidgetItem, QScrollArea)
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("matrix_title"))
        dlg.resize(560, 440)
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel(self.t("matrix_desc")))

        fils = self._slot_filaments()
        n = len(fils)
        if n == 0:
            lay.addWidget(QLabel(self.t("no_filaments")))
            close_btn = QPushButton(self.t("exp_cancel"))
            close_btn.clicked.connect(dlg.accept)
            lay.addWidget(close_btn)
            dlg.exec()
            return

        table = QTableWidget(n, n)
        headers = [f.get("name", f.get("id", "?"))[:10] for f in fils]
        table.setHorizontalHeaderLabels(headers)
        table.setVerticalHeaderLabels(headers)

        for row_i, f1 in enumerate(fils):
            for col_j, f2 in enumerate(fils):
                seq = [f1["id"], f2["id"]]
                mixed = self._simulate_mix(seq, fils)
                hx = lab_to_hex(mixed)
                item = QTableWidgetItem(hx)
                item.setBackground(
                    __import__("PySide6.QtGui", fromlist=["QColor"]).QColor(hx)
                )
                r, g, b = hex_to_rgb(hx)
                lum = 0.299*r + 0.587*g + 0.114*b
                fg = "#000000" if lum > 140 else "#FFFFFF"
                item.setForeground(
                    __import__("PySide6.QtGui", fromlist=["QColor"]).QColor(fg)
                )
                table.setItem(row_i, col_j, item)

        lay.addWidget(table, 1)
        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.accept)
        lay.addWidget(close_btn)
        dlg.exec()

    # ── SEQUENCE EDITOR ───────────────────────────────────────────────────────

    def _open_sequence_editor(self):
        from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
            QLabel, QPushButton, QComboBox, QLineEdit, QScrollArea,
            QWidget, QMessageBox)
        dlg = QDialog(self)
        dlg.setWindowTitle(self.t("seqed_title"))
        dlg.resize(480, 380)
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel(self.t("seqed_desc")))

        fils = self._slot_filaments()
        seq = []

        seq_label = QLabel(self.t("seqed_seq_empty"))
        lay.addWidget(seq_label)

        preview_lbl = QLabel()
        preview_lbl.setFixedHeight(36)
        preview_lbl.setStyleSheet("background:#808080;border-radius:4px;")
        lay.addWidget(preview_lbl)

        def _update_preview():
            if not seq or not fils:
                seq_label.setText(self.t("seqed_seq_empty"))
                preview_lbl.setStyleSheet("background:#808080;border-radius:4px;")
                return
            seq_label.setText(self.t("seqed_seq_prefix") + "".join(seq))
            mixed = self._simulate_mix(seq, fils)
            hx = lab_to_hex(mixed)
            preview_lbl.setStyleSheet(f"background:{hx};border-radius:4px;")

        slot_row = QHBoxLayout()
        slot_combo = QComboBox()
        for f in fils:
            slot_combo.addItem(f.get("name", f.get("id", "?")), f["id"])
        slot_row.addWidget(slot_combo)

        add_btn = QPushButton(self.t("seqed_add"))
        def _add():
            if len(seq) >= MAX_SEQ_LEN:
                QMessageBox.information(dlg, "", self.t("seqed_max"))
                return
            fid = slot_combo.currentData()
            if fid:
                seq.append(fid)
                _update_preview()
        add_btn.clicked.connect(_add)
        slot_row.addWidget(add_btn)

        rm_btn = QPushButton(self.t("seqed_remove"))
        def _rm():
            if seq:
                seq.pop()
                _update_preview()
        rm_btn.clicked.connect(_rm)
        slot_row.addWidget(rm_btn)

        clr_btn = QPushButton(self.t("seqed_clear"))
        def _clr():
            seq.clear()
            _update_preview()
        clr_btn.clicked.connect(_clr)
        slot_row.addWidget(clr_btn)
        lay.addLayout(slot_row)

        def _apply_seq():
            if not seq:
                return
            seq_str = "".join(seq)
            mixed = self._simulate_mix(seq, fils)
            hx = lab_to_hex(mixed)
            lh = self._lh_spin.value() if hasattr(self, "_lh_spin") else 0.2
            cad = calc_cadence(seq, lh)
            de = 0.0
            if hasattr(self, "_target_hex") and self._target_hex:
                t_lab = rgb_to_lab(hex_to_rgb(self._target_hex))
                de = round(delta_e(t_lab, mixed), 2)
            self._virtual.append({
                "hex": hx,
                "label": f"Seq {seq_str}",
                "seq": list(seq),
                "de": de,
            })
            self._recalc_all_virtual()
            self._refresh_virtual_grid()
            QMessageBox.information(dlg, "", self.t("seqed_added").format(seq_str))

        btn_row = QHBoxLayout()
        apply_btn = QPushButton(self.t("seqed_apply"))
        apply_btn.clicked.connect(_apply_seq)
        btn_row.addWidget(apply_btn)
        close_btn = QPushButton(self.t("exp_cancel"))
        close_btn.clicked.connect(dlg.accept)
        btn_row.addWidget(close_btn)
        lay.addLayout(btn_row)
        dlg.exec()

    # ── PICK COLOR FROM IMAGE ─────────────────────────────────────────────────

    def _pick_color_from_image(self):
        from PySide6.QtWidgets import QFileDialog, QMessageBox
        f, _ = QFileDialog.getOpenFileName(
            self, self.t("img_pick_title"), "",
            "Images (*.png *.jpg *.jpeg *.bmp *.webp *.tiff)")
        if not f:
            return
        try:
            from PIL import Image
            img = Image.open(f).convert("RGB")
            img.thumbnail((400, 400))
            # Simple dominant color via quantize
            q = img.quantize(colors=1, method=Image.Quantize.MEDIANCUT)
            pal = q.getpalette()
            r, g, b = pal[0], pal[1], pal[2]
            hx = rgb_to_hex(r, g, b)
            self._apply_target(hx)
        except Exception as e:
            QMessageBox.warning(self, "", str(e))


# ── 3MF FARB-WIZARD (QDialog) ─────────────────────────────────────────────────

class ThreeMFLoaderWorker(QThread):
    """Loads and parses a .3mf file in background — keeps UI responsive."""
    finished = Signal(list, str)   # colors, error_msg

    def __init__(self, path):
        super().__init__()
        self._path = path

    def run(self):
        colors, err = _parse_3mf_colors(self._path)
        self.finished.emit(colors or [], err or "")


class WizardOptimizerWorker(QThread):
    progress = Signal(int, int)   # current, total
    finished = Signal(list, float, list)  # best_4, avg_de, scores

    def __init__(self, target_labs, library_fils):
        super().__init__()
        self._target_labs = target_labs
        self._library_fils = library_fils

    def run(self):
        def progress_cb(i, total):
            self.progress.emit(i, total)
        best4, avg_de, scores = find_best_4_filaments(
            self._target_labs, self._library_fils, progress_cb)
        self.finished.emit(best4, avg_de, scores)


class ThreeMFWizardDialog(QDialog):
    """3-step wizard: load 3MF → optimize 4 filaments → show result & apply."""

    def __init__(self, app):
        super().__init__(app)
        self._app = app
        self.setWindowTitle(app.t("wizard_title"))
        self.resize(720, 580)
        self.setModal(False)

        self._colors = []
        self._best4 = []
        self._avg_de = 0.0
        self._scores = []
        self._worker = None

        self._main_layout = QVBoxLayout(self)
        self._main_layout.setContentsMargins(0, 0, 0, 0)

        self._stack = QStackedWidget()
        self._main_layout.addWidget(self._stack)

        self._page1 = self._build_page1()
        self._page2 = self._build_page2()
        self._page3 = self._build_page3()
        self._stack.addWidget(self._page1)
        self._stack.addWidget(self._page2)
        self._stack.addWidget(self._page3)
        self._stack.setCurrentIndex(0)

    # ── PAGE 1 ─────────────────────────────────────────────────────────────────

    def _build_page1(self):
        app = self._app
        t = app.t
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        title = QLabel(t("wizard_step1"))
        title.setObjectName("section_title")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        load_btn = QPushButton(t("wizard_load_btn"))
        load_btn.setFixedHeight(52)
        load_btn.setStyleSheet("background-color: #0f4c81; font-weight: bold; font-size: 13px;")
        load_btn.clicked.connect(self._load_file)
        layout.addWidget(load_btn)

        self._p1_path_lbl = QLabel(t("wizard_no_file"))
        self._p1_path_lbl.setObjectName("hint")
        self._p1_path_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(self._p1_path_lbl)

        self._p1_count_lbl = QLabel("")
        self._p1_count_lbl.setAlignment(Qt.AlignCenter)
        self._p1_count_lbl.setStyleSheet("color: #4ade80; font-weight: bold; font-size: 12px;")
        layout.addWidget(self._p1_count_lbl)

        # Swatches area
        swatch_scroll = QScrollArea()
        swatch_scroll.setWidgetResizable(True)
        swatch_scroll.setFixedHeight(120)
        swatch_container = QWidget()
        self._p1_swatch_grid = QGridLayout(swatch_container)
        self._p1_swatch_grid.setSpacing(4)
        swatch_scroll.setWidget(swatch_container)
        layout.addWidget(swatch_scroll)

        layout.addStretch()

        self._p1_spinner_lbl = QLabel("")
        self._p1_spinner_lbl.setAlignment(Qt.AlignCenter)
        self._p1_spinner_lbl.setStyleSheet("color: #94a3b8; font-size: 11px;")
        layout.addWidget(self._p1_spinner_lbl)

        self._p1_next_btn = QPushButton(t("wizard_next"))
        self._p1_next_btn.setFixedHeight(42)
        self._p1_next_btn.setEnabled(False)
        self._p1_next_btn.setStyleSheet(
            "background-color: #15803d; font-weight: bold; font-size: 13px;")
        self._p1_next_btn.clicked.connect(self._go_page2)
        layout.addWidget(self._p1_next_btn)

        self._p1_load_btn = None   # filled in after build
        return page

    def _load_file(self):
        app = self._app
        path, _ = QFileDialog.getOpenFileName(
            self,
            app.t("wizard_load_btn"), "",
            f"{app.t('3mf_filetypes')} (*.3mf);;All Files (*.*)")
        if not path:
            return

        # Show loading spinner, disable buttons
        self._p1_path_lbl.setText(os.path.basename(path))
        self._p1_count_lbl.setText("")
        self._p1_spinner_lbl.setText(app.t("wizard_load_spinner"))
        self._p1_next_btn.setEnabled(False)

        # Clear old swatches
        while self._p1_swatch_grid.count():
            item = self._p1_swatch_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # Run in background thread
        self._loader_worker = ThreeMFLoaderWorker(path)
        self._loader_worker.finished.connect(self._on_3mf_loaded)
        self._loader_worker.start()

    def _on_3mf_loaded(self, colors, err):
        # Guard: dialog may have been closed before async load finished
        try:
            spinner_alive = self._p1_spinner_lbl.isVisible()
        except RuntimeError:
            return
        app = self._app
        self._p1_spinner_lbl.setText("")
        if not colors:
            self._p1_count_lbl.setText(
                f"⚠ {err}" if err else app.t("wizard_no_file"))
            return
        self._colors = colors
        self._p1_count_lbl.setText(
            app.t("wizard_colors_found", n=len(colors)))
        # Draw swatches
        for i, hex_c in enumerate(colors[:24]):
            swatch = QLabel()
            swatch.setFixedSize(24, 24)
            try:
                r, g, b = hex_to_rgb(hex_c)
                swatch.setStyleSheet(
                    f"background-color: rgb({r},{g},{b}); border-radius: 3px;")
            except Exception:
                pass
            swatch.setToolTip(hex_c)
            self._p1_swatch_grid.addWidget(swatch, i // 12, i % 12)
        self._p1_next_btn.setEnabled(True)

    def _go_page2(self):
        self._update_page2_info()
        self._stack.setCurrentIndex(1)

    # ── PAGE 2 ─────────────────────────────────────────────────────────────────

    def _build_page2(self):
        app = self._app
        t = app.t
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        title = QLabel(t("wizard_step2"))
        title.setObjectName("section_title")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        self._p2_info_lbl = QLabel("")
        self._p2_info_lbl.setObjectName("hint")
        self._p2_info_lbl.setWordWrap(True)
        self._p2_info_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(self._p2_info_lbl)

        self._p2_progress = QProgressBar()
        self._p2_progress.setRange(0, 100)
        self._p2_progress.setValue(0)
        layout.addWidget(self._p2_progress)

        self._p2_status_lbl = QLabel("")
        self._p2_status_lbl.setObjectName("hint")
        self._p2_status_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(self._p2_status_lbl)

        layout.addStretch()

        btn_row = QHBoxLayout()
        self._p2_start_btn = QPushButton(t("wizard_start"))
        self._p2_start_btn.setFixedHeight(42)
        self._p2_start_btn.setStyleSheet(
            "background-color: #2563eb; font-weight: bold; font-size: 13px;")
        self._p2_start_btn.clicked.connect(self._start_optimization)
        btn_row.addWidget(self._p2_start_btn)

        self._p2_next_btn = QPushButton(t("wizard_next"))
        self._p2_next_btn.setFixedHeight(42)
        self._p2_next_btn.setEnabled(False)
        self._p2_next_btn.setStyleSheet(
            "background-color: #15803d; font-weight: bold; font-size: 13px;")
        self._p2_next_btn.clicked.connect(self._go_page3)
        btn_row.addWidget(self._p2_next_btn)

        layout.addLayout(btn_row)
        return page

    def _update_page2_info(self):
        app = self._app
        lib = app._get_all_library_fils()
        self._lib_fils = lib
        self._p2_info_lbl.setText(
            app.t("wizard_info", n_lib=len(lib), n_col=len(self._colors)))

    def _start_optimization(self):
        app = self._app
        self._p2_start_btn.setEnabled(False)
        self._p2_next_btn.setEnabled(False)
        self._p2_progress.setValue(0)
        target_labs = [rgb_to_lab(hex_to_rgb(h)) for h in self._colors]
        self._worker = WizardOptimizerWorker(target_labs, self._lib_fils)
        self._worker.progress.connect(self._on_worker_progress)
        self._worker.finished.connect(self._on_worker_finished)
        self._worker.start()

    def _on_worker_progress(self, i, total):
        app = self._app
        pct = int(i / total * 100) if total > 0 else 0
        self._p2_progress.setValue(pct)
        self._p2_status_lbl.setText(
            app.t("wizard_checking", i=i, total=total))

    def _on_worker_finished(self, best4, avg_de, scores):
        app = self._app
        self._best4 = best4
        self._avg_de = avg_de
        self._scores = scores
        self._p2_progress.setValue(100)
        self._p2_status_lbl.setText(app.t("wizard_avg_de", de=avg_de))
        self._p2_next_btn.setEnabled(True)

    def _go_page3(self):
        self._rebuild_page3()
        self._stack.setCurrentIndex(2)

    # ── PAGE 3 ─────────────────────────────────────────────────────────────────

    def _build_page3(self):
        # Placeholder; rebuilt dynamically in _rebuild_page3
        page = QWidget()
        page.setLayout(QVBoxLayout())
        return page

    def _rebuild_page3(self):
        app = self._app
        t = app.t
        # Replace old page3 widget
        old_page = self._stack.widget(2)
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(20, 14, 20, 14)
        layout.setSpacing(8)

        title = QLabel(t("wizard_step3"))
        title.setObjectName("section_title")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # 4 filament cards
        cards_row = QHBoxLayout()
        slot_labels = ["T1", "T2", "T3", "T4"]
        best4 = (self._best4 + [None] * 4)[:4]
        for i, fil in enumerate(best4):
            card = QGroupBox(slot_labels[i])
            card_layout = QVBoxLayout(card)
            card_layout.setSpacing(3)
            if fil:
                hex_c = fil.get("hex", "#808080")
                swatch = QLabel()
                swatch.setFixedSize(40, 40)
                try:
                    r, g, b = hex_to_rgb(hex_c)
                    swatch.setStyleSheet(
                        f"background-color: rgb({r},{g},{b}); border-radius: 5px;")
                except Exception:
                    pass
                swatch.setAlignment(Qt.AlignCenter)
                card_layout.addWidget(swatch)
                brand_lbl = QLabel(fil.get("brand", ""))
                brand_lbl.setObjectName("hint")
                brand_lbl.setWordWrap(True)
                brand_lbl.setAlignment(Qt.AlignCenter)
                card_layout.addWidget(brand_lbl)
                name_lbl = QLabel(fil.get("name", ""))
                name_lbl.setWordWrap(True)
                name_lbl.setAlignment(Qt.AlignCenter)
                name_lbl.setStyleSheet("font-weight: bold;")
                card_layout.addWidget(name_lbl)
                td_lbl = QLabel(f"TD={fil.get('td', DEFAULT_TD):.1f}")
                td_lbl.setObjectName("hint")
                td_lbl.setAlignment(Qt.AlignCenter)
                card_layout.addWidget(td_lbl)
            cards_row.addWidget(card)
        layout.addLayout(cards_row)

        # Summary
        de_lbl = QLabel(t("wizard_avg_de", de=self._avg_de))
        de_lbl.setAlignment(Qt.AlignCenter)
        de_lbl.setStyleSheet(
            f"color: {de_color(self._avg_de)}; font-weight: bold; font-size: 13px;")
        layout.addWidget(de_lbl)

        # Coverage table
        cov_title = QLabel(t("wizard_coverage"))
        cov_title.setObjectName("hint")
        cov_title.setAlignment(Qt.AlignCenter)
        layout.addWidget(cov_title)

        table = QTableWidget(len(self._colors), 4)
        table.setHorizontalHeaderLabels([
            t("wizard_cov_color"), t("wizard_cov_hex"),
            t("wizard_cov_de"), t("wizard_cov_quality"),
        ])
        table.horizontalHeader().setStretchLastSection(True)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setFixedHeight(130)
        for i, (hex_c, sc) in enumerate(zip(self._colors, self._scores)):
            color_item = QTableWidgetItem("")
            try:
                r, g, b = hex_to_rgb(hex_c)
                color_item.setBackground(QColor(r, g, b))
            except Exception:
                pass
            table.setItem(i, 0, color_item)
            table.setItem(i, 1, QTableWidgetItem(hex_c))
            de_item = QTableWidgetItem(f"{sc:.1f}")
            de_item.setForeground(QColor(de_color(sc)))
            table.setItem(i, 2, de_item)
            q = ("✓" if sc < DE_GOOD else "~" if sc < DE_OK else "✗")
            table.setItem(i, 3, QTableWidgetItem(q))
        layout.addWidget(table)

        # Add virtual heads checkbox
        self._p3_add_virtual_chk = QCheckBox(t("wizard_add_virtual"))
        layout.addWidget(self._p3_add_virtual_chk)

        # Buttons
        btn_row = QHBoxLayout()
        apply_btn = QPushButton(t("wizard_apply"))
        apply_btn.setFixedHeight(42)
        apply_btn.setStyleSheet(
            "background-color: #15803d; font-weight: bold; font-size: 12px;")
        apply_btn.clicked.connect(self._apply_and_close)
        btn_row.addWidget(apply_btn)

        close_btn = QPushButton(t("wizard_close"))
        close_btn.setFixedHeight(42)
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        # Swap widget in stack
        self._stack.removeWidget(old_page)
        old_page.deleteLater()
        self._page3 = page
        self._stack.insertWidget(2, page)

    def _apply_and_close(self):
        app = self._app
        best4 = self._best4
        for i, fil in enumerate(best4[:4]):
            if fil is None:
                continue
            brand = fil.get("brand", "")
            name = fil.get("name", "")
            hex_c = fil.get("hex", "#808080")
            td = fil.get("td", DEFAULT_TD)
            # Use _apply_slot
            app._apply_slot(i, {
                "brand": brand,
                "filament": name,
                "name": name,
                "hex": hex_c,
                "td": td,
            })

        # Optionally calculate virtual heads for all model colors
        if hasattr(self, "_p3_add_virtual_chk") and self._p3_add_virtual_chk.isChecked():
            for hex_c in self._colors:
                if len(app._virtual) >= app._max_virtual:
                    break
                result = app._calc_for_color(hex_c, auto=True)
                if result:
                    app.add_virtual(result)

        QMessageBox.information(self, app.t("wizard_title"), app.t("wizard_applied"))
        self.accept()


# ── FULLSPECTRUM EXPORT DIALOG ────────────────────────────────────────────────

class FullSpectrumExportDialog(QDialog):
    """Dialog for FullSpectrum Direct 3MF Export."""

    def __init__(self, app):
        super().__init__(app)
        self._app = app
        self.setWindowTitle(app.t("fs_export_title"))
        self.resize(720, 660)
        self._src_path = ""
        self._dst_path = ""
        self._build_ui()

    def _build_ui(self):
        app = self._app
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(16, 16, 16, 12)

        # Title
        title = QLabel(app.t("fs_export_title"))
        title.setObjectName("section_title")
        title.setStyleSheet("font-size: 15px; font-weight: bold; color: #a78bfa;")
        layout.addWidget(title)

        desc = QLabel(app.t("fs_export_desc"))
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #94a3b8; font-size: 10px;")
        layout.addWidget(desc)

        warn = QLabel(app.t("fs_warn_fullspectrum"))
        warn.setWordWrap(True)
        warn.setStyleSheet("color: #fbbf24; font-size: 10px;")
        layout.addWidget(warn)

        # Source file
        src_grp = QGroupBox(app.t("fs_src_label"))
        src_grp_layout = QHBoxLayout(src_grp)
        self._src_edit = QLineEdit()
        self._src_edit.setPlaceholderText("*.3mf …")
        self._src_edit.setReadOnly(False)
        src_grp_layout.addWidget(self._src_edit, 1)
        src_btn = QPushButton("📂 " + app.t("btn_load"))
        src_btn.setFixedWidth(100)
        src_btn.clicked.connect(self._browse_src)
        src_grp_layout.addWidget(src_btn)
        layout.addWidget(src_grp)

        # Output file
        dst_grp = QGroupBox(app.t("fs_dst_label"))
        dst_grp_layout = QVBoxLayout(dst_grp)
        self._radio_overwrite = QPushButton()  # placeholder
        from PySide6.QtWidgets import QRadioButton as _QRB
        self._rb_overwrite = _QRB(app.t("fs_overwrite"))
        self._rb_overwrite.setChecked(True)
        self._rb_save_as = _QRB(app.t("fs_save_as"))
        dst_grp_layout.addWidget(self._rb_overwrite)
        dst_grp_layout.addWidget(self._rb_save_as)
        dst_row = QHBoxLayout()
        self._dst_edit = QLineEdit()
        self._dst_edit.setPlaceholderText("*.3mf …")
        dst_row.addWidget(self._dst_edit, 1)
        dst_btn = QPushButton("📂 " + app.t("fs_save_as"))
        dst_btn.setFixedWidth(160)
        dst_btn.clicked.connect(self._browse_dst)
        dst_row.addWidget(dst_btn)
        dst_grp_layout.addLayout(dst_row)
        layout.addWidget(dst_grp)

        # Settings
        set_grp = QGroupBox(app.t("settings_title"))
        set_layout = QHBoxLayout(set_grp)
        set_layout.addWidget(QLabel(app.t("fs_lh_label")))
        self._lh_spin = QDoubleSpinBox()
        self._lh_spin.setRange(0.04, 0.30)
        self._lh_spin.setSingleStep(0.04)
        self._lh_spin.setDecimals(3)
        # Get current layer height from app
        try:
            self._lh_spin.setValue(app._lh_spin.value())
        except Exception:
            self._lh_spin.setValue(0.08)
        self._lh_spin.setFixedWidth(80)
        self._lh_spin.setToolTip(
            "Optimal: 0.08–0.12 mm für sauberes Farbblending.\n"
            "Optimal: 0.08–0.12 mm for clean color blending.\n"
            "> 0.15 mm → sichtbare Streifen / visible striping.")
        self._lh_spin.valueChanged.connect(self._on_lh_changed)
        set_layout.addWidget(self._lh_spin)
        self._lh_warn_lbl = QLabel("")
        self._lh_warn_lbl.setStyleSheet("color: #fb923c; font-size: 11px;")
        set_layout.addWidget(self._lh_warn_lbl)
        self._on_lh_changed(self._lh_spin.value())
        set_layout.addSpacing(16)
        n_virt = len([v for v in app._virtual if v.get("sequence")])
        self._count_lbl = QLabel(app.t("fs_count", n=n_virt))
        self._count_lbl.setStyleSheet("color: #94a3b8;")
        set_layout.addWidget(self._count_lbl)
        set_layout.addSpacing(16)
        self._local_z_check = QCheckBox(app.t("lbl_local_z"))
        self._local_z_check.setToolTip(app.t("lbl_local_z_tip"))
        set_layout.addWidget(self._local_z_check)
        set_layout.addSpacing(8)
        self._adv_dither_check = QCheckBox(app.t("lbl_adv_dither"))
        self._adv_dither_check.setToolTip(app.t("lbl_adv_dither_tip"))
        set_layout.addWidget(self._adv_dither_check)
        set_layout.addStretch()
        layout.addWidget(set_grp)

        # Preview
        prev_grp = QGroupBox(app.t("fs_preview_groupbox"))
        prev_layout = QVBoxLayout(prev_grp)
        self._preview = QTextEdit()
        self._preview.setReadOnly(True)
        self._preview.setFont(QFont("Courier New", 9))
        self._preview.setMinimumHeight(100)
        prev_layout.addWidget(self._preview)
        layout.addWidget(prev_grp, 1)

        # Buttons
        btn_row = QHBoxLayout()
        prev_btn = QPushButton(app.t("fs_preview_btn"))
        prev_btn.clicked.connect(self._update_preview)
        btn_row.addWidget(prev_btn)

        write_btn = QPushButton(app.t("fs_write_btn"))
        write_btn.setFixedHeight(44)
        write_btn.setStyleSheet(
            "background-color: #7c2d12; color: white; font-size: 13px; font-weight: bold;")
        write_btn.clicked.connect(self._do_write)
        btn_row.addWidget(write_btn)

        close_btn = QPushButton(app.t("exp_cancel"))
        close_btn.clicked.connect(self.reject)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        # Initial preview
        QTimer.singleShot(100, self._update_preview)

    def _on_lh_changed(self, val):
        if val > 0.15:
            self._lh_warn_lbl.setText(self._app.t("fs_lh_warn"))
        else:
            self._lh_warn_lbl.setText("")

    def _browse_src(self):
        p, _ = QFileDialog.getOpenFileName(
            self, self._app.t("fs_src_label"), "",
            f"{self._app.t('3mf_filetypes')} (*.3mf);;All Files (*.*)")
        if p:
            self._src_edit.setText(p)
            if not self._dst_edit.text():
                self._dst_edit.setText(p)

    def _browse_dst(self):
        p, _ = QFileDialog.getSaveFileName(
            self, self._app.t("fs_dst_label"), "",
            f"{self._app.t('3mf_filetypes')} (*.3mf);;All Files (*.*)")
        if p:
            self._dst_edit.setText(p)
            self._rb_save_as.setChecked(True)

    def _update_preview(self):
        lh = self._lh_spin.value()
        defs, extra = build_mixed_filament_definitions(self._app._virtual, lh)
        raw_rows = [row for row in defs.split(";") if row]

        preview_lines = []
        for i, (vf, row) in enumerate(
                zip([v for v in self._app._virtual if v.get("sequence")], raw_rows)):
            seq = vf.get("sequence", "")
            unique_ids = list(dict.fromkeys(int(c) for c in seq)) if seq else []
            n_unique = len(unique_ids)
            label = vf.get("label", f"V{vf.get('vid', i + 5)}")

            if n_unique == 2:
                cad = calc_cadence(seq, lh)
                fid_a, fid_b = unique_ids[0], unique_ids[1]
                ca = cad.get(fid_a, lh)
                cb = cad.get(fid_b, lh)
                cadence_note = f"  # {label}: cad_a={ca}mm cad_b={cb}mm"
            elif n_unique == 1:
                cadence_note = f"  # {label}: pure"
            else:
                cadence_note = f"  # {label}: pattern ({n_unique} fils)"

            preview_lines.append(row + cadence_note)

        preview_lines.append("")
        preview_lines.append(
            f"# global: cad_a={extra.get('mixed_color_layer_height_a', lh)}mm"
            f"  cad_b={extra.get('mixed_color_layer_height_b', lh)}mm  "
            f"(median of {len([v for v in self._app._virtual if v.get('sequence')])} heads)")
        self._preview.setPlainText("\n".join(preview_lines))

    def _do_write(self):
        app = self._app
        src = self._src_edit.text().strip()
        if not src:
            QMessageBox.warning(self, app.t("dlg_note"), app.t("fs_no_src"))
            return

        # ── Version compatibility check ───────────────────────────────────────
        try:
            with zipfile.ZipFile(src, "r") as _zf:
                _cfg_names = [n for n in _zf.namelist()
                              if "project_settings" in n.lower() and n.endswith(".config")]
                _ver_str = None
                for _cfg_name in _cfg_names:
                    try:
                        _cfg_data = json.loads(_zf.read(_cfg_name).decode("utf-8", errors="replace"))
                        _ver_str = (_cfg_data.get("orcaslicer_version")
                                    or _cfg_data.get("version")
                                    or _cfg_data.get("bambu_studio_version"))
                        if _ver_str:
                            break
                    except Exception:
                        pass
                if _ver_str:
                    # Parse version tuple: compare against 2.2.4 (FS v0.7-alpha threshold)
                    _min_ver = (2, 2, 4)
                    try:
                        _parts = [int(x) for x in re.split(r"[.\-]", _ver_str)
                                  if x.isdigit()][:3]
                        while len(_parts) < 3:
                            _parts.append(0)
                        _is_old = tuple(_parts) < _min_ver
                    except Exception:
                        _is_old = False
                    if _is_old:
                        _warn_msg = (
                            f"OrcaSlicer version detected: {_ver_str}\n\n"
                            "This version may predate FullSpectrum v0.7-alpha "
                            "(requires OrcaSlicer ≥ 2.2.4).\n"
                            "Local-Z dithering (dithering_local_z_mode) may not be "
                            "supported.\n\n"
                            "Continue anyway?")
                        if self._local_z_check.isChecked():
                            _warn_msg = (
                                "⚠  Local-Z Dithering is enabled, but the detected "
                                f"OrcaSlicer version ({_ver_str}) may not support it.\n\n"
                                "FullSpectrum v0.7-alpha requires OrcaSlicer ≥ 2.2.4.\n\n"
                                "Continue anyway?")
                        _reply = QMessageBox.warning(
                            self, "Version Compatibility",
                            _warn_msg,
                            QMessageBox.Ok | QMessageBox.Cancel,
                            QMessageBox.Cancel)
                        if _reply != QMessageBox.Ok:
                            return
        except Exception:
            pass  # If version check fails, proceed silently

        lh = self._lh_spin.value()

        if self._rb_overwrite.isChecked():
            out = src
        else:
            out = self._dst_edit.text().strip()
            if not out:
                QMessageBox.warning(self, app.t("dlg_note"), app.t("fs_no_src"))
                return

        # Collect physical filament data
        phys = []
        for s in getattr(app, "_slots", []):
            try:
                hx = s["hex_edit"].text().strip()
                nm = s["fil_combo"].currentText()
                br = s["brand_combo"].currentText()
                td = s["td_spin"].value()
            except Exception:
                hx = "#808080"; nm = ""; br = ""; td = 5.0
            phys.append({"hex": hx, "name": nm, "brand": br, "td": td})

        ok, msg = inject_fullspectrum_into_3mf(
            src, out, app._virtual, phys, lh,
            local_z=self._local_z_check.isChecked(),
            advanced_dithering=self._adv_dither_check.isChecked())
        if ok:
            QMessageBox.information(self, app.t("dlg_saved"),
                                    app.t("fs_success", path=out))
            self.accept()
        else:
            QMessageBox.critical(self, app.t("dlg_error"), msg)


# ── ENTRY POINT ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    app.setApplicationName("U1 FullSpectrum Helper")
    app.setOrganizationName("SnapmakerCommunity")

    # Apply default dark theme before window is shown
    try:
        window = U1App()
        window.show()
        sys.exit(app.exec())
    except Exception as exc:
        import traceback
        traceback.print_exc()
        sys.exit(1)
