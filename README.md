# Snapmaker U1 FullSpectrum Helper

**[English](#english) | [Deutsch](#deutsch)**

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

### Features

| Feature | Description |
|---|---|
| **4 Physical Slots** | Load your actual filaments (T1–T4) with brand, color, and TD value |
| **Color Picker** | Choose target color visually or enter a hex code |
| **Sequence Optimizer** | Tests all 24 permutations to find the lowest ΔE sequence |
| **Variable Sequence Length** | Auto-mode or manual (1–10 layers), shorter = fewer tool changes |
| **20 Virtual Print Heads** | V5–V24, each = a FullSpectrum layer sequence |
| **3MF Assistant** | Opens a `.3mf` file, extracts all colors, calculates sequences automatically |
| **Filament Database** | Searchable library with presets for Bambu Lab, Prusament, Anycubic + custom entries |
| **Slot Presets** | Save and reload your favorite 4-slot configurations |
| **Export** | Export sequences as `.json` or `.txt` with OrcaSlicer-ready slicer hints |
| **History** | Last 5 calculated sequences shown inline |

### Color Science

- Colors are calculated in **CIE Lab color space** (perceptual, not RGB)
- **ΔE (Delta E)** measures color distance: `< 3` = excellent, `< 6` = good, `≥ 6` = noticeable difference
- **TD (Transmission Distance)** models filament opacity — lower TD = more opaque = dominates the mix
- **Progressive layer weighting** — top layers weighted 1.5× vs bottom layers (1.0×) to match visual perception
- **Gamut warning** at ΔE > 25 (target color is outside the achievable range)

### Slicer Integration (OrcaSlicer FullSpectrum)

After calculating sequences, use the output in OrcaSlicer under **Others → Dithering**:

| Filament count in sequence | OrcaSlicer setting |
|---|---|
| 1 filament | Pure color — no mixing needed |
| 2 filaments | **Cadence Height** A + B (in mm) |
| 3–4 filaments | **Pattern Mode** — enter the pattern string (e.g. `1/1/2/1/3`) |

The Dithering Step Size corresponds to your **layer height** (default: 0.2 mm — not the nozzle diameter).

### Requirements

```
Python 3.10+
customtkinter
```

```bash
pip install customtkinter
```

### Usage

```bash
python u1_ultimate.py
```

1. **Load filaments** — set brand, filament name, hex color and TD for each of the 4 physical tools
2. **Pick a target color** — use the color picker or type a hex code
3. **Calculate** — click *Berechnen* (or *Optimieren* for full optimizer)
4. **Add to virtual grid** — click *Zu Virtuell* to add the result as V5, V6, etc.
5. **Export** — use the export dialog to generate `.json` or `.txt` for OrcaSlicer

For batch processing a `.3mf` file, click **3MF Assistent**.

### File Structure

```
u1_ultimate.py      — main application (single file)
filament_db.json    — your saved custom filaments (auto-created)
presets.json        — your saved slot presets (auto-created)
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
| **Sequenz-Optimizer** | Testet alle 24 Permutationen für das kleinste ΔE |
| **Variable Sequenzlänge** | Auto-Modus oder manuell (1–10 Schichten) |
| **20 virtuelle Druckköpfe** | V5–V24, jeder = eine FullSpectrum-Schichtsequenz |
| **3MF-Assistent** | Öffnet `.3mf`, extrahiert alle Farben, berechnet Sequenzen automatisch |
| **Filament-Datenbank** | Bibliothek mit Voreinstellungen für Bambu Lab, Prusament, Anycubic + eigene Einträge |
| **Slot-Presets** | Bevorzugte 4-Slot-Konfigurationen speichern und laden |
| **Export** | Sequenzen als `.json` oder `.txt` mit OrcaSlicer-Hinweisen exportieren |
| **Verlauf** | Die letzten 5 berechneten Sequenzen werden direkt angezeigt |

### Farbwissenschaft

- Berechnungen im **CIE-Lab-Farbraum** (wahrnehmungsbasiert, nicht RGB)
- **ΔE (Delta E)** misst den Farbabstand: `< 3` = ausgezeichnet, `< 6` = gut, `≥ 6` = spürbar
- **TD (Transmission Distance)** modelliert die Deckkraft des Filaments — niedriger TD = deckender = dominiert die Mischung
- **Progressive Schichtgewichtung** — obere Schichten werden 1,5× stärker gewichtet als untere (1,0×)
- **Gamut-Warnung** bei ΔE > 25 (Zielfarbe liegt außerhalb des erreichbaren Bereichs)

### Slicer-Integration (OrcaSlicer FullSpectrum)

Nach der Berechnung die Ausgabe in OrcaSlicer unter **Others → Dithering** eingeben:

| Filament-Anzahl in der Sequenz | OrcaSlicer-Einstellung |
|---|---|
| 1 Filament | Reine Farbe — kein Mix nötig |
| 2 Filamente | **Cadence Height** A + B (in mm) |
| 3–4 Filamente | **Pattern Mode** — Pattern-String eingeben (z.B. `1/1/2/1/3`) |

Der Dithering Step Size entspricht der **Schichthöhe** (Standard: 0,2 mm — nicht der Düsendurchmesser).

### Voraussetzungen

```
Python 3.10+
customtkinter
```

```bash
pip install customtkinter
```

### Verwendung

```bash
python u1_ultimate.py
```

1. **Filamente laden** — Marke, Name, Hex-Farbe und TD für jeden der 4 physischen Werkzeugköpfe eintragen
2. **Zielfarbe wählen** — Farbwähler verwenden oder Hex-Code eingeben
3. **Berechnen** — *Berechnen* klicken (oder *Optimieren* für den vollständigen Optimizer)
4. **Zum virtuellen Raster hinzufügen** — *Zu Virtuell* klicken, ergibt V5, V6, usw.
5. **Exportieren** — Export-Dialog öffnen und `.json` oder `.txt` für OrcaSlicer erzeugen

Für die Stapelverarbeitung einer `.3mf`-Datei den **3MF Assistent** verwenden.

### Dateistruktur

```
u1_ultimate.py      — Hauptanwendung (einzelne Datei)
filament_db.json    — Gespeicherte eigene Filamente (wird automatisch erstellt)
presets.json        — Gespeicherte Slot-Presets (wird automatisch erstellt)
```

### Verwandte Projekte

- [OrcaSlicer FullSpectrum](https://github.com/ratdoux/OrcaSlicer-FullSpectrum) — das Slicer-Plugin, für das dieses Tool gedacht ist
- [Snapmaker Forum](https://forum.snapmaker.com) — Community-Support

## License / Lizenz

MIT
