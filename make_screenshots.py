"""
Screenshot generator for U1 FullSpectrum Helper documentation.
Run: python make_screenshots.py
Saves PNGs to docs/screenshots/
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QTimer
from PySide6.QtGui import QScreen

OUT_DIR = os.path.join(os.path.dirname(__file__), "docs", "screenshots")
os.makedirs(OUT_DIR, exist_ok=True)

app = QApplication(sys.argv)

# Import and create the main window
from u1_pyside6 import U1App
win = U1App()
win.resize(1440, 900)
win.show()
win.raise_()
win.activateWindow()

def grab(name):
    path = os.path.join(OUT_DIR, f"{name}.png")
    app.processEvents()
    screen = QApplication.primaryScreen()
    pixmap = screen.grabWindow(win.winId())
    pixmap.save(path)
    print(f"  Saved: {path}")

def run_sequence():
    steps = [
        # Force Calculator tab first (override QSettings restore)
        (300,  lambda: win._tabs.setCurrentIndex(0),                 None),
        (400,  lambda: None,                                         "01_start"),

        # Pick red target → screenshot before calc
        (200,  lambda: win._apply_target("#E63946"),                 None),
        (700,  lambda: None,                                         "02_target_red"),

        # Calculate red
        (200,  lambda: win._calc(),                                  None),
        (1400, lambda: None,                                         "03_result_red"),

        # Blue target + calc
        (200,  lambda: win._apply_target("#1D3557"),                 None),
        (200,  lambda: win._calc(),                                  None),
        (1400, lambda: None,                                         "04_result_blue"),
        (200,  lambda: win.add_virtual(),                            None),

        # Teal target + calc
        (200,  lambda: win._apply_target("#2A9D8F"),                 None),
        (200,  lambda: win._calc(),                                  None),
        (1400, lambda: None,                                         None),
        (200,  lambda: win.add_virtual(),                            None),

        # Yellow target + calc
        (200,  lambda: win._apply_target("#E9C46A"),                 None),
        (200,  lambda: win._calc(),                                  None),
        (1400, lambda: None,                                         None),
        (200,  lambda: win.add_virtual(),                            None),

        # Virtual Heads tab
        (400,  lambda: win._tabs.setCurrentIndex(1),                 None),
        (600,  lambda: None,                                         "05_virtual_heads"),

        # Tools tab
        (200,  lambda: win._tabs.setCurrentIndex(2),                 None),
        (600,  lambda: None,                                         "06_tools_tab"),

        # Back to Calculator, expand all slots
        (300,  lambda: win._tabs.setCurrentIndex(0),                 None),
        (200,  lambda: [win._toggle_slot(i) for i in range(1, 4)],  None),
        (500,  lambda: None,                                         "07_sidebar_slots"),

        # Done
        (500,  lambda: app.quit(),                                   None),
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

run_sequence()
sys.exit(app.exec())
