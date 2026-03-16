# Snapmaker U1 FullSpectrum Helper

A GUI tool for the **Snapmaker U1 Toolchanger** that calculates optical color mixing sequences using the [FullSpectrum principle](https://github.com/ratdoux/OrcaSlicer-FullSpectrum).

## What is FullSpectrum?

FullSpectrum is a technique for multi-tool 3D printers that achieves a wider color range by **alternating print layers** instead of mixing filaments physically. For example, alternating red and yellow layers creates the visual impression of orange — similar to how pixels work on a screen, but with layers of filament.

This tool helps you:
- Calculate the optimal layer sequence for any target color
- Preview the simulated mixed color before printing
- Generate up to **20 virtual print heads** (V5–V24) for OrcaSlicer
- Analyze `.3mf` files and find the best matching sequences for all colors in the model

## Features

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

## Color Science

- Colors are calculated in **CIE Lab color space** (perceptual, not RGB)
- **ΔE (Delta E)** measures color distance: `< 3` = excellent, `< 6` = good, `≥ 6` = noticeable difference
- **TD (Transmission Distance)** models filament opacity — lower TD = more opaque = dominates the mix
- **Progressive layer weighting** — top layers are weighted 1.5× vs bottom layers (1.0×) to match visual perception
- **Gamut warning** at ΔE > 25 (target color is outside the achievable range)

## Slicer Integration (OrcaSlicer FullSpectrum)

After calculating sequences, use the output in OrcaSlicer under **Others → Dithering**:

| Filament count in sequence | OrcaSlicer setting |
|---|---|
| 1 filament | Pure color — no mixing needed |
| 2 filaments | **Cadence Height** A + B (in mm) |
| 3–4 filaments | **Pattern Mode** — enter the pattern string (e.g. `1/1/2/1/3`) |

The Dithering Step Size corresponds to your **layer height** (default: 0.2 mm — not the nozzle diameter).

## Requirements

```
Python 3.10+
customtkinter
```

Install dependencies:

```bash
pip install customtkinter
```

## Usage

```bash
python u1_ultimate.py
```

1. **Load filaments** — set brand, filament name, hex color and TD for each of the 4 physical tools
2. **Pick a target color** — use the color picker or type a hex code
3. **Calculate** — click *Berechnen* (or *Optimieren* for full optimizer)
4. **Add to virtual grid** — click *Zu Virtuell* to add the result as V5, V6, etc.
5. **Export** — use the export dialog to generate `.json` or `.txt` for OrcaSlicer

For batch processing a `.3mf` file, click **3MF Assistent**.

## File Structure

```
u1_ultimate.py      — main application (single file)
filament_db.json    — your saved custom filaments (auto-created)
presets.json        — your saved slot presets (auto-created)
```

## Related

- [OrcaSlicer FullSpectrum](https://github.com/ratdoux/OrcaSlicer-FullSpectrum) — the slicer plugin this tool targets
- [Snapmaker Forum](https://forum.snapmaker.com) — community support

## License

MIT
