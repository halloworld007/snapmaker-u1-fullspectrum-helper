"""
Screenshot generator for U1 FullSpectrum Helper documentation.
Run: python make_screenshots.py
Saves PNGs to docs/screenshots/
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from PySide6.QtWidgets import QApplication, QDialog
from PySide6.QtCore import QTimer

OUT_DIR = os.path.join(os.path.dirname(__file__), "docs", "screenshots")
os.makedirs(OUT_DIR, exist_ok=True)

app = QApplication(sys.argv)

from u1_pyside6 import U1App, ThreeMFWizardDialog, FullSpectrumExportDialog
win = U1App()
win.resize(1440, 900)
win.show()
win.raise_()
win.activateWindow()

_open_dialogs = []   # keep refs alive

def grab(name, widget=None):
    path = os.path.join(OUT_DIR, f"{name}.png")
    app.processEvents()
    target = widget or win
    pixmap = QApplication.primaryScreen().grabWindow(target.winId())
    pixmap.save(path)
    print(f"  Saved: {path}")

def open_and_grab_dialog(dlg_class, name, setup_fn=None, close_after=True):
    dlg = dlg_class(win)
    _open_dialogs.append(dlg)
    if setup_fn:
        setup_fn(dlg)
    dlg.show()
    dlg.raise_()
    def do_grab():
        grab(name, dlg)
        if close_after:
            dlg.close()
            _open_dialogs.discard(dlg) if hasattr(_open_dialogs, 'discard') else None
    QTimer.singleShot(600, do_grab)

def run_sequence():
    steps = [
        # ── Calculator screenshots ───────────────────────────────────────────
        (300,  lambda: win._tabs.setCurrentIndex(0),                   None),
        (400,  lambda: None,                                           "01_start"),

        (200,  lambda: win._apply_target("#E63946"),                   None),
        (600,  lambda: None,                                           "02_target_red"),
        (200,  lambda: win._calc(),                                    None),
        (1400, lambda: None,                                           "03_result_red"),

        (200,  lambda: win._apply_target("#1D3557"),                   None),
        (200,  lambda: win._calc(),                                    None),
        (1400, lambda: None,                                           "04_result_blue"),
        (200,  lambda: win.add_virtual(),                              None),

        (200,  lambda: win._apply_target("#2A9D8F"),                   None),
        (200,  lambda: win._calc(),                                    None),
        (1400, lambda: None,                                           None),
        (200,  lambda: win.add_virtual(),                              None),

        (200,  lambda: win._apply_target("#E9C46A"),                   None),
        (200,  lambda: win._calc(),                                    None),
        (1400, lambda: None,                                           None),
        (200,  lambda: win.add_virtual(),                              None),

        # ── Virtual Heads tab ────────────────────────────────────────────────
        (400,  lambda: win._tabs.setCurrentIndex(1),                   None),
        (600,  lambda: None,                                           "05_virtual_heads"),

        # ── Tools tab ────────────────────────────────────────────────────────
        (200,  lambda: win._tabs.setCurrentIndex(2),                   None),
        (600,  lambda: None,                                           "06_tools_tab"),

        # ── Sidebar with all slots open ───────────────────────────────────────
        (300,  lambda: win._tabs.setCurrentIndex(0),                   None),
        (200,  lambda: [win._toggle_slot(i) for i in range(1, 4)],    None),
        (500,  lambda: None,                                           "07_sidebar_slots"),

        # ── 3MF Wizard dialog (Step 1 — load file) ───────────────────────────
        (400,  lambda: _grab_3mf_wizard(),                             None),

        # ── FullSpectrum Export dialog ────────────────────────────────────────
        (1400, lambda: _grab_fs_export(),                              None),

        # ── 3MF Assistant dialog (inline, open via method with dummy colors) ─
        (1400, lambda: _grab_3mf_assistant(),                          None),

        # ── Filament Search dialog ────────────────────────────────────────────
        (1400, lambda: _grab_filament_search(),                        None),

        # Done
        (1800, lambda: app.quit(),                                     None),
    ]

    total_delay = 0
    for delay, action, name in steps:
        total_delay += delay
        def make_cb(a, n, d):
            def cb():
                a()
                if n:
                    grab(n)
            QTimer.singleShot(d, cb)
        make_cb(action, name, total_delay)


def _grab_3mf_wizard():
    dlg = ThreeMFWizardDialog(win)
    _open_dialogs.append(dlg)
    dlg.show()
    dlg.raise_()
    def do_grab():
        grab("08_3mf_wizard_step1", dlg)
        dlg.close()
    QTimer.singleShot(700, do_grab)


def _grab_fs_export():
    dlg = FullSpectrumExportDialog(win)
    _open_dialogs.append(dlg)
    dlg.show()
    dlg.raise_()
    def do_grab():
        grab("09_fs_3mf_export", dlg)
        dlg.close()
    QTimer.singleShot(700, do_grab)


def _grab_3mf_assistant():
    """Open the 3MF color assistant with dummy colors (no file needed)."""
    dummy_colors = ["#E63946", "#1D3557", "#2A9D8F", "#E9C46A",
                    "#F4A261", "#264653", "#A8DADC", "#457B9D"]
    try:
        win._open_3mf_with_colors("example_model.3mf", dummy_colors)
        # find the newly opened dialog
        for w in app.topLevelWidgets():
            if isinstance(w, QDialog) and w is not win and w.isVisible():
                _open_dialogs.append(w)
                def do_grab(dlg=w):
                    grab("10_3mf_assistant", dlg)
                    QTimer.singleShot(400, dlg.close)
                QTimer.singleShot(700, do_grab)
                break
    except Exception as e:
        print(f"  [skip] 3MF assistant: {e}")


def _grab_filament_search():
    """Open the filament search dialog for slot 0."""
    try:
        win._open_filament_search(0)
        for w in app.topLevelWidgets():
            if isinstance(w, QDialog) and w is not win and w.isVisible():
                _open_dialogs.append(w)
                def do_grab(dlg=w):
                    grab("11_filament_search", dlg)
                    QTimer.singleShot(400, dlg.close)
                QTimer.singleShot(700, do_grab)
                break
    except Exception as e:
        print(f"  [skip] Filament search: {e}")


run_sequence()
sys.exit(app.exec())
