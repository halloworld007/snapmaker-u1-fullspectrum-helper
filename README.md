# Snapmaker U1 FullSpectrum Helper

**[English](#english) | [Deutsch](#deutsch)**

> **Aktiver Entwicklungszweig / Active development:** `u1_pyside6.py` (PySide6)
> `u1_ultimate.py` (CustomTkinter) wird nicht mehr weiterentwickelt und dient nur noch als Referenz.

---

## English

A GUI tool for the **Snapmaker U1 Toolchanger** that calculates optical color mixing sequences using the [FullSpectrum principle](https://github.com/ratdoux/OrcaSlicer-FullSpectrum).

### What is FullSpectrum?

FullSpectrum is a technique for multi-tool 3D printers that achieves a wider color range by **alternating print layers** instead of mixing filaments physically. For example, alternating red and yellow layers creates the visual impression of orange — similar to how pixels work on a screen, but with layers of filament.

This tool helps you:
- Calculate the optimal layer sequence for any target color
- Preview the simulated mixed color before printing
- Generate up to **20 virtual print heads** (V5–V24) for OrcaSlicer
- Analyze `.3mf` files and find the best matching sequences for all colors in the model

## Screenshots

### Calculator — Single Color Mode
Pick a target color, load your filaments, and calculate the optimal layer sequence.

![Calculator](docs/screenshots/01_start.png)

### Result — Color Match with ΔE
The result card shows Target vs. Simulated color, ΔE quality, sequence string, and layer preview bar.

| Red target (out of gamut) | Blue target |
|---|---|
| ![Result Red](docs/screenshots/03_result_red.png) | ![Result Blue](docs/screenshots/04_result_blue.png) |

### Virtual Print Heads (V5–V24)
Calculated sequences are added as virtual heads — ready for OrcaSlicer export or FS 3MF injection.

![Virtual Heads](docs/screenshots/05_virtual_heads.png)

### Tools Tab
Analysis, color generation, optimization, library management and export — all in one place.

![Tools](docs/screenshots/06_tools_tab.png)

### Sidebar — Physical Slots T1–T4
Load brand, filament, hex color and TD value for each physical tool. The sidebar is resizable by dragging.

![Sidebar Slots](docs/screenshots/07_sidebar_slots.png)

### Features

| Feature | Description |
|---|---|
| **4 Physical Slots** | Load your actual filaments (T1–T4) with brand, color, and TD value |
| **Color Picker** | Choose target color visually or enter a hex code |
| **Sequence Optimizer** | Tests all permutations to find the lowest ΔE sequence |
| **Variable Sequence Length** | Auto-mode or manual (1–10 layers), shorter = fewer tool changes |
| **Top-3 Comparison** | Shows the 3 best sequences side-by-side with one-click add |
| **20 Virtual Print Heads** | V5–V24, each = a FullSpectrum layer sequence |
| **Color Model Toggle** | Switch between Additive (FullSpectrum-compatible), TD-weighted, or Subtractive (pigment) simulation |
| **Gamut Strip** | Visual strip showing achievable colors with current filaments |
| **Gamut Plot** | CIE a*b* diagram with reachable color space and convex hull |
| **Gradient Generator** | Generate smooth color gradients as a batch of virtual heads |
| **Color Harmonies** | Complement, triadic, analogous, split-complement via HSV |
| **3MF Assistant** | Opens a `.3mf` file, extracts all colors, calculates sequences automatically |
| **3MF Extruder Remap** | Remaps original extruder assignments to T1–T4 or virtual heads and writes back |
| **Filament Database** | 226+ presets for Bambu Lab, Prusament, eSUN, Polymaker, Hatchbox, Snapmaker + custom entries |
| **Online Update** | Fetch latest community filament database from GitHub |
| **Slot Optimizer** | Finds the best 4-filament combination from your library for a set of target colors |
| **Slot Presets** | Save and reload your favorite 4-slot configurations |
| **Project Save/Load** | Save and restore complete sessions (slots + virtual heads) as `.u1proj` |
| **TD Calibration** | Estimate TD values from a test print measurement |
| **Palette from Image** | Extract dominant colors from any image file |
| **Direct OrcaSlicer Export** | Write filament profiles directly into OrcaSlicer / Snapmaker_Orca |
| **FullSpectrum Direct 3MF Export** | Inject virtual heads as `mixed_filament_definitions` directly into a `.3mf` file — Cadence and Pattern pre-configured |
| **Export** | Export sequences as `.json` or `.txt` with OrcaSlicer-ready slicer hints |
| **Copy All Cadence Values** | One-click copy of all virtual head dithering settings as formatted text |
| **Tool Change Estimator** | Estimates tool changes, extra print time and purge volume |
| **History** | Last 5 calculated sequences shown inline |

### Color Science

- Simulation uses **gamma-corrected linear RGB averaging** — matches the visual result of alternating thin layers (same model as FilamentMixer / FullSpectrum slicer)
- **ΔE (Delta E)** measures color distance: `< 3` = excellent, `< 6` = good, `≥ 6` = noticeable difference
- **TD (Transmission Distance)** models filament opacity — available as optional weighting in the color model toggle
- Optimization and ΔE calculations run in **CIE Lab color space** (perceptually uniform)
- **Gamut warning** at ΔE > 25 (target color is outside the achievable range)

### Slicer Integration

#### FullSpectrum Slicer (Snapmaker_Orca)

Export **T1–T4 only** (the tool warns you automatically). FullSpectrum generates Mixed Filaments from your physical slots automatically.

| Filament count in sequence | FullSpectrum setting |
|---|---|
| 1 filament | Pure color — assign directly |
| 2 filaments | Mixed Filament → set ratio to match sequence counts |
| 3–4 filaments | Not directly supported — use standard OrcaSlicer instead |

#### Standard OrcaSlicer

Export T1–T4 **and** V5+ virtual heads. Assign virtual heads to objects, then configure under **Others → Dithering**:

| Filament count in sequence | OrcaSlicer setting |
|---|---|
| 1 filament | Pure color — no mixing needed |
| 2 filaments | **Cadence Height** A + B (in mm) |
| 3–4 filaments | **Pattern Mode** — enter the pattern string (e.g. `1121`) |

The Dithering Step Size corresponds to your **layer height** (recommended: **0.08 mm** — color layers become visually invisible at this height).

### FullSpectrum Direct 3MF Export

#### What it does

Writes all virtual print heads (V5–V24) as `mixed_filament_definitions` directly into a `.3mf` project file. After injection, open the file in the FullSpectrum Slicer (Snapmaker_Orca) — Cadence heights and Pattern sequences are already configured. You only need to assign objects to the virtual heads and print.

#### Step by step

1. Configure filaments (T1–T4) in the physical slots
2. Calculate target colors and add them as virtual print heads
3. Click **"💉 FS 3MF Export"** in the Virtual Heads tab toolbar
4. Select the source `.3mf` model file
5. Click **"💉 Write"** — the modified `.3mf` is created
6. Open the `.3mf` in FullSpectrum Slicer (Snapmaker_Orca)
7. Assign objects to the virtual print heads (Mixed Filaments)
8. Print — done!

#### What gets configured automatically

| Setting | Description |
|---|---|
| `mixed_filament_definitions` | All V5–V24 as Mixed Filament rows |
| `mixed_color_layer_height_a` / `_b` | Cadence heights calculated from the sequence |
| `dithering_z_step_size` | Set to 0.0 (local Z mode) |
| `dithering_local_z_mode` | Enabled |
| `filament_colour` | Physical filament colors T1–T4 |
| `mixed_filament_height_lower_bound` / `_upper_bound` | Derived from layer height |

#### Limitation

Object painting (which surfaces get which print head) must still be done in the slicer. This tool configures the Mixed Filament parameters — assigning them to objects and surfaces is the slicer's job.

#### FullSpectrum slicer only

> `mixed_filament_definitions` is a FullSpectrum-specific extension. It is only supported by Snapmaker_Orca (FullSpectrum fork), not by standard OrcaSlicer.

---

### Requirements

```
Python 3.10+
PySide6
```

```bash
pip install PySide6
```

Optional (for extra features):
```bash
pip install matplotlib scipy Pillow
```

### Usage

```bash
python u1_pyside6.py
```

1. **Load filaments** — set brand, filament name, hex color and TD for each of the 4 physical tools
2. **Pick a target color** — use the color picker or type a hex code
3. **Calculate** — click *Berechnen* (or *Optimieren* for full optimizer)
4. **Add to virtual grid** — click *Zu Virtuell* to add the result as V5, V6, etc.
5. **Export** — use the OrcaSlicer export or copy all cadence values with one click

For batch processing a `.3mf` file, click **3MF Assistent**.
To remap a BambuStudio/OrcaSlicer `.3mf` for the U1, click **✏️ 3MF schreiben**.

### File Structure

```
u1_pyside6.py           — main application (single file, actively maintained)
u1_ultimate.py          — legacy CustomTkinter version (no longer developed)
filament_db.json        — your saved custom filaments (auto-created)
presets.json            — your saved slot presets (auto-created)
*.u1proj                — saved project files (slots + virtual heads)
```

---

## Deutsch

Ein GUI-Tool für den **Snapmaker U1 Toolchanger**, das optimale Schichtreihenfolgen für optisches Farbmischen nach dem [FullSpectrum-Prinzip](https://github.com/ratdoux/OrcaSlicer-FullSpectrum) berechnet.

### Was ist FullSpectrum?

FullSpectrum ist eine Technik für 3D-Drucker mit mehreren Werkzeugköpfen, die durch **abwechselnde Druckschichten** eine größere Farbpalette erzielt — ohne physisches Mischen der Filamente. Zum Beispiel entsteht durch abwechselnde rote und gelbe Schichten der optische Eindruck von Orange — ähnlich wie Pixel auf einem Bildschirm, nur mit Filamentschichten.

Dieses Tool hilft dabei:
- Die optimale Schichtsequenz für eine Zielfarbe zu berechnen
- Die simulierte Mischfarbe vorab als Vorschau zu sehen
- Bis zu **20 virtuelle Druckköpfe** (V5–V24) für OrcaSlicer zu erzeugen
- `.3mf`-Dateien zu analysieren und passende Sequenzen für alle enthaltenen Farben zu berechnen

### Funktionen

| Funktion | Beschreibung |
|---|---|
| **4 physische Slots** | Eigene Filamente (T1–T4) mit Marke, Farbe und TD-Wert hinterlegen |
| **Farbwähler** | Zielfarbe visuell wählen oder Hex-Code eingeben |
| **Sequenz-Optimizer** | Testet alle Permutationen für das kleinste ΔE |
| **Variable Sequenzlänge** | Auto-Modus oder manuell (1–10 Schichten) |
| **Top-3-Vergleich** | Die 3 besten Sequenzen nebeneinander, mit Ein-Klick-Hinzufügen |
| **20 virtuelle Druckköpfe** | V5–V24, jeder = eine FullSpectrum-Schichtsequenz |
| **Farbmodell-Toggle** | Umschalten zwischen Additiv (FullSpectrum-kompatibel), TD-gewichtet und Subtraktiv |
| **Gamut-Strip** | Horizontaler Streifen mit erreichbaren Farben der aktuellen Slots |
| **Gamut-Plot** | CIE-a*b*-Diagramm mit erreichbarem Farbraum und konvexer Hülle |
| **Gradient-Generator** | Farbverläufe als Batch virtueller Köpfe generieren |
| **Farbharmonien** | Komplement, Triade, Analog, Split-Komplement per HSV |
| **3MF-Assistent** | Öffnet `.3mf`, extrahiert alle Farben, berechnet Sequenzen automatisch |
| **3MF Extruder-Remap** | Remappt Extruder-Zuweisungen auf T1–T4 oder V-Köpfe und schreibt zurück |
| **Filament-Datenbank** | 226+ Einträge für Bambu Lab, Prusament, eSUN, Polymaker, Hatchbox, Snapmaker + eigene |
| **Online-Update** | Community-Filamentdatenbank von GitHub laden |
| **Slot-Optimizer** | Findet die beste 4-Filament-Kombination aus der Bibliothek für Zielfarben |
| **Slot-Presets** | Bevorzugte 4-Slot-Konfigurationen speichern und laden |
| **Projekt speichern/laden** | Komplette Session (Slots + virtuelle Köpfe) als `.u1proj` sichern |
| **TD-Kalibrierung** | TD-Wert aus Testdruck-Messung schätzen |
| **Palette aus Bild** | Dominante Farben aus einer Bilddatei extrahieren |
| **Direktexport → OrcaSlicer** | Filamentprofile direkt in OrcaSlicer / Snapmaker_Orca schreiben |
| **FullSpectrum Direct 3MF Export** | Virtuelle Köpfe als `mixed_filament_definitions` direkt in eine `.3mf` Datei einschreiben — Cadence und Pattern bereits konfiguriert |
| **Export** | Sequenzen als `.json` oder `.txt` mit OrcaSlicer-Hinweisen exportieren |
| **Alle Cadence-Werte kopieren** | Ein-Klick-Kopie aller Dithering-Werte als formatierten Text |
| **Werkzeugwechsel-Schätzer** | Schätzt Werkzeugwechsel, Zusatzzeit und Purge-Volumen |
| **Verlauf** | Die letzten 5 berechneten Sequenzen werden direkt angezeigt |

### Screenshots

#### Rechner — Einzelfarb-Modus
Zielfarbe wählen, Filamente laden, Sequenz berechnen.

![Calculator](docs/screenshots/01_start.png)

#### Ergebnis — Farbabgleich mit ΔE
Die Ergebniskarte zeigt Ziel- vs. Simulierfarbe, ΔE-Qualität, Sequenz und Schichtvorschau.

| Rotes Ziel (außerhalb Gamut) | Blaues Ziel |
|---|---|
| ![Ergebnis Rot](docs/screenshots/03_result_red.png) | ![Ergebnis Blau](docs/screenshots/04_result_blue.png) |

#### Virtuelle Druckköpfe (V5–V24)
Berechnete Sequenzen werden als virtuelle Köpfe gespeichert — bereit für OrcaSlicer-Export oder FS-3MF-Einschreiben.

![Virtuelle Köpfe](docs/screenshots/05_virtual_heads.png)

#### Tools-Tab
Analyse, Farbgenerierung, Optimierung, Bibliothek und Export — alles an einem Ort.

![Tools](docs/screenshots/06_tools_tab.png)

#### Sidebar — Physische Slots T1–T4
Marke, Filament, Hex-Farbe und TD für jeden physischen Kopf eintragen. Die Sidebar ist durch Ziehen verbreiterbar.

![Slots](docs/screenshots/07_sidebar_slots.png)

### Farbwissenschaft

- Die Simulation verwendet **gamma-korrigierten linearen RGB-Durchschnitt** — entspricht dem visuellen Eindruck alternierender dünner Schichten (identisch mit FilamentMixer / FullSpectrum-Slicer)
- **ΔE (Delta E)** misst den Farbabstand: `< 3` = ausgezeichnet, `< 6` = gut, `≥ 6` = spürbar
- **TD (Transmission Distance)** modelliert die Deckkraft — optional als TD-gewichtetes Farbmodell wählbar
- Optimierung und ΔE-Berechnung laufen im **CIE-Lab-Farbraum** (wahrnehmungsgetreu)
- **Gamut-Warnung** bei ΔE > 25 (Zielfarbe liegt außerhalb des erreichbaren Bereichs)

### Slicer-Integration

#### FullSpectrum-Slicer (Snapmaker_Orca)

Nur **T1–T4 exportieren** (das Tool warnt automatisch). FullSpectrum generiert Mixed Filaments aus den physischen Slots selbst.

| Filamentanzahl in der Sequenz | FullSpectrum-Einstellung |
|---|---|
| 1 Filament | Reine Farbe — direkt zuweisen |
| 2 Filamente | Mixed Filament → Verhältnis aus Sequenz einstellen |
| 3–4 Filamente | Nicht direkt unterstützt — Standard-OrcaSlicer verwenden |

#### Standard-OrcaSlicer

T1–T4 **und** V5+ exportieren. Virtuelle Köpfe den Objekten zuweisen, dann unter **Others → Dithering** konfigurieren:

| Filamentanzahl in der Sequenz | OrcaSlicer-Einstellung |
|---|---|
| 1 Filament | Reine Farbe — kein Mix nötig |
| 2 Filamente | **Cadence Height** A + B (in mm) |
| 3–4 Filamente | **Pattern Mode** — Ziffernkette eingeben (z.B. `1121`) |

Der Dithering Step Size entspricht der **Schichthöhe** (empfohlen: **0,08 mm** — Farbschichten werden bei dieser Höhe unsichtbar).

### FullSpectrum Direct 3MF Export

#### Was es macht

Schreibt alle virtuellen Druckköpfe (V5–V24) als `mixed_filament_definitions` direkt in eine `.3mf` Projektdatei. Nach dem Einschreiben die Datei im FullSpectrum Slicer (Snapmaker_Orca) öffnen — Cadence-Höhen und Pattern-Sequenzen sind bereits konfiguriert. Es müssen nur noch Objekte den virtuellen Köpfen zugewiesen werden.

#### Schritt-für-Schritt

1. Filamente (T1–T4) in den physischen Slots konfigurieren
2. Zielfarben berechnen und als virtuelle Druckköpfe hinzufügen
3. **"💉 FS 3MF Export"** im Toolbar der "Virtuelle Köpfe"-Registerkarte klicken
4. Quelldatei (`.3mf` Modell) auswählen
5. **"💉 Einschreiben"** klicken → modifizierte `.3mf` wird erzeugt
6. `.3mf` im FullSpectrum Slicer (Snapmaker_Orca) öffnen
7. Objekte den virtuellen Druckköpfen (Mixed Filaments) zuweisen
8. Drucken — fertig!

#### Was automatisch konfiguriert wird

| Einstellung | Beschreibung |
|---|---|
| `mixed_filament_definitions` | Alle V5–V24 als Mixed-Filament-Zeilen |
| `mixed_color_layer_height_a` / `_b` | Cadence-Höhen aus der berechneten Sequenz |
| `dithering_z_step_size` | Auf 0.0 gesetzt (lokaler Z-Modus) |
| `dithering_local_z_mode` | Aktiviert |
| `filament_colour` | Physische Filamentfarben T1–T4 |
| `mixed_filament_height_lower_bound` / `_upper_bound` | Abgeleitet aus der Schichthöhe |

#### Einschränkung

Die Objekt-Bemalung (welche Flächen welchen Kopf bekommen) muss weiterhin im Slicer gesetzt werden. Das Tool konfiguriert die Mixed-Filament-Parameter — die Zuweisung zu Objekten/Flächen ist Slicer-Aufgabe.

#### Nur FullSpectrum Slicer

> `mixed_filament_definitions` ist eine FullSpectrum-spezifische Erweiterung und wird nur von Snapmaker_Orca (FullSpectrum Fork) unterstützt, nicht von Standard-OrcaSlicer.

---

### Voraussetzungen

```
Python 3.10+
PySide6
```

```bash
pip install PySide6
```

Optional (für zusätzliche Funktionen):
```bash
pip install matplotlib scipy Pillow
```

### Verwendung

```bash
python u1_pyside6.py
```

1. **Filamente laden** — Marke, Name, Hex-Farbe und TD für jeden der 4 physischen Werkzeugköpfe eintragen
2. **Zielfarbe wählen** — Farbwähler verwenden oder Hex-Code eingeben
3. **Berechnen** — *Berechnen* klicken (oder *Optimieren* für den vollständigen Optimizer)
4. **Zum virtuellen Raster hinzufügen** — *Zu Virtuell* klicken, ergibt V5, V6, usw.
5. **Exportieren** — OrcaSlicer-Export nutzen oder alle Cadence-Werte mit einem Klick kopieren

Für die Stapelverarbeitung einer `.3mf`-Datei den **3MF Assistent** verwenden.
Zum Remappen einer BambuStudio/OrcaSlicer-`.3mf` für den U1: **✏️ 3MF schreiben** klicken.

### Dateistruktur

```
u1_pyside6.py           — Hauptanwendung (single file, aktiv gepflegt)
u1_ultimate.py          — Legacy CustomTkinter-Version (nicht mehr weiterentwickelt)
filament_db.json        — Gespeicherte eigene Filamente (wird automatisch erstellt)
presets.json            — Gespeicherte Slot-Presets (wird automatisch erstellt)
*.u1proj                — Gespeicherte Projekt-Dateien (Slots + virtuelle Köpfe)
```

### Verwandte Projekte

- [OrcaSlicer FullSpectrum](https://github.com/ratdoux/OrcaSlicer-FullSpectrum) — der Slicer-Fork, für den dieses Tool entwickelt wurde
- [Snapmaker Forum](https://forum.snapmaker.com) — Community-Support

## License / Lizenz

MIT
