# FullSpectrum Slicer Integration Guide

## Overview

This guide explains the technical details of the FullSpectrum Direct 3MF Export feature.
It covers the `.3mf` project file format, the `mixed_filament_definitions` token schema,
cadence height calculation, manual pattern encoding, and a complete step-by-step workflow.

The FullSpectrum technique uses **alternating thin print layers** of different filaments to
achieve optical color mixing. Instead of physically mixing filaments, colors emerge from the
visual blending of stacked thin layers — similar to how pixels work on a display.

The Snapmaker_Orca slicer (FullSpectrum fork) natively supports this via the
`mixed_filament_definitions` configuration key, which this tool injects directly into `.3mf`
project files.

---

## Project File Format (.3mf)

A `.3mf` file is a ZIP archive. The relevant file inside is:

```
Metadata/project_settings.config
```

This is a JSON document. A minimal example:

```json
{
    "version": "01.09.04.52",
    "name": "My Print",
    "from": "project",
    "filament_colour": ["#FF0000", "#00FF00", "#0000FF", "#FFFF00"],
    "mixed_filament_definitions": "1,2,1,1,50,0,g,w,m2,d0,o0,u1001"
}
```

### ZIP structure

```
mymodel.3mf
├── 3D/
│   └── 3dmodel.model          (geometry — not modified)
├── Metadata/
│   ├── project_settings.config  (JSON — modified by this tool)
│   ├── model_settings.config    (object assignments — not modified)
│   └── slice_info.config        (slice metadata — not modified)
└── _rels/
    └── .rels
```

The tool reads the ZIP, modifies only `Metadata/project_settings.config`, and writes a new ZIP.
All other entries are passed through unchanged. If `project_settings.config` does not exist,
a minimal one is created.

---

## mixed_filament_definitions Format

The value is a semicolon-separated list of rows. Each row represents one virtual print head
(Mixed Filament slot in the slicer).

### Row schema

```
component_a,component_b,enabled,custom,mix_b_percent,pointillism,g<ids>,w<weights>,m<mode>,d<deleted>,o<origin>,u<stable_id>[,manual_pattern]
```

### Token reference table

| Token | Type | Description |
|---|---|---|
| `component_a` | int | Filament index for the first component (1-based: T1=1, T2=2, T3=3, T4=4) |
| `component_b` | int | Filament index for the second component |
| `enabled` | 0/1 | Whether this Mixed Filament is active (always 1) |
| `custom` | 0/1 | Whether this is a custom mix (always 1 for user-defined) |
| `mix_b_percent` | int | Percentage of component_b in the mix (0–100) |
| `pointillism` | 0/1 | Pointillism mode flag (0 = disabled, use cadence/pattern) |
| `g<ids>` | string | Gradient filament IDs (e.g. `g12` for T1+T2, `g` for 2-filament cadence mode) |
| `w<weights>` | string | Gradient weights as slash-separated percentages (e.g. `w60/40`, `w` for auto) |
| `m<mode>` | int | Mix mode: `m2` = standard cadence/pattern mode |
| `d<deleted>` | 0/1 | Deleted flag (always 0) |
| `o<origin>` | int | Origin flag (always 0 for user-created) |
| `u<stable_id>` | int | Stable ID — unique integer that persists across slicer sessions |
| `manual_pattern` | string (optional) | Comma-separated sequence for 3–4 filament pattern mode (e.g. `1,1,2,1`) |

### Examples

**Pure color (T2 only):**
```
2,2,1,1,0,0,g,w,m2,d0,o0,u1001
```

**2-filament cadence (T1 and T2, 33% T2):**
```
1,2,1,1,33,0,g,w,m2,d0,o0,u1002
```

**3-filament pattern (T1, T2, T3 — pattern 1,1,2,1,3):**
```
1,3,1,1,20,0,g123,w60/20/20,m2,d0,o0,u1003,1,1,2,1,3
```

**4-filament pattern (T1–T4 — pattern 1,2,3,4,1,1):**
```
1,4,1,1,17,0,g1234,w33/17/33/17,m2,d0,o0,u1004,1,2,3,4,1,1
```

### Stable IDs

Stable IDs (`u<id>`) must be unique across all rows. This tool assigns them as `1000 + index`
by default, or uses the `stable_id` field from the virtual head dict if present.

---

## Cadence Height Calculation

For 2-filament Mixed Filaments, the FullSpectrum slicer uses **Cadence Heights** to determine
how many layers of each filament to print per cycle.

### Formula

The slicer formula (v0.92+):

1. Count occurrences of each filament in the sequence string (e.g. `"11121"` → T1=4, T2=1)
2. Sort by count: minority first
3. Minority side anchors to 1 layer
4. Majority side scales: `layers = max(1, round(majority_pct / minority_pct))`
5. Multiply by `layer_height` (in mm) to get the cadence height

### Example

Sequence: `"11121"` with layer height 0.08 mm

- T1 count = 4, T2 count = 1, total = 5
- T2 is minority (1/5 = 20%)
- T1 majority: `round(4/1) = 4` layers → cadence A = `4 × 0.08 = 0.32 mm`
- T2 minority: `1` layer → cadence B = `1 × 0.08 = 0.08 mm`

This means the slicer prints 4 layers of T1, then 1 layer of T2, repeating.

### Global cadence values

When multiple virtual heads are present, the tool calculates the **median** cadence height
across all 2-filament heads and writes it as the global:

- `mixed_color_layer_height_a` — cadence for the majority/first filament
- `mixed_color_layer_height_b` — cadence for the minority/second filament

These are global slicer settings. For best accuracy, use similar sequences across all virtual
heads, or accept the median as a reasonable approximation.

---

## Manual Pattern Format

For sequences with 3 or 4 unique filaments, cadence mode is not sufficient. The slicer uses
**Manual Pattern Mode** instead.

### Encoding

The pattern token is appended as the last field of the row (after `u<stable_id>`):

```
<row...>,u<id>,<p1>,<p2>,<p3>,...
```

Each `<pi>` is a filament index (1–4). The pattern is the complete layer sequence to repeat.

### Example

Sequence string: `"112131"` (T1, T1, T2, T1, T3, T1)

Pattern field: `1,1,2,1,3,1`

Full row:
```
1,3,1,1,17,0,g123,w67/17/17,m2,d0,o0,u1005,1,1,2,1,3,1
```

### Weights

The `w<weights>` token contains the percentage of each filament in the sequence:
- Count each filament: T1=4, T2=1, T3=1 → total=6
- T1: round(4/6×100) = 67%
- T2: round(1/6×100) = 17%
- T3: round(1/6×100) = 17%
- Normalize to sum 100: adjust first weight if off by rounding

Result: `w67/17/17`

---

## Physical Filament Colors

The `filament_colour` array in `project_settings.config` holds the hex colors of the 4
physical filaments (T1–T4). This tool reads the hex colors from the active slots and writes
them:

```json
"filament_colour": ["#FF0000", "#00FF00", "#0000FF", "#FFFF00"]
```

Index 0 = T1, index 1 = T2, index 2 = T3, index 3 = T4.

The slicer uses these colors to display the filament swatches and compute the visual preview.

---

## Additional Injected Settings

Besides `mixed_filament_definitions` and `filament_colour`, these keys are written:

| Key | Value | Description |
|---|---|---|
| `mixed_color_layer_height_a` | float (mm) | Global cadence height for component A |
| `mixed_color_layer_height_b` | float (mm) | Global cadence height for component B |
| `dithering_z_step_size` | `"0.0"` | Use local Z mode (not global step) |
| `dithering_local_z_mode` | `"0"` | Local Z offset mode |
| `dithering_step_painted_zones_only` | `"1"` | Only dither painted zones |
| `mixed_filament_advanced_dithering` | `"0"` | Disable advanced dithering |
| `mixed_filament_gradient_mode` | `"0"` | Disable gradient mode |
| `mixed_filament_height_lower_bound` | `"0.04"` | Minimum cadence height |
| `mixed_filament_height_upper_bound` | float (mm) | Maximum cadence height (4× layer height) |

---

## Step-by-step Workflow

### In the U1 FullSpectrum Helper

1. **Configure physical slots**: Open the app, go to the Sidebar. Set brand, filament name,
   hex color, and TD for T1–T4. These are the actual filaments loaded in your printer.

2. **Select target colors**: Use the Single Color Calculator. Enter a hex code or use the
   color picker. Click "SEQUENZ BERECHNEN" to calculate the optimal layer sequence.

3. **Add virtual heads**: Click "➕ Als virtuellen Druckkopf hinzufügen" to add the result.
   Repeat for each color you need. You can also use the 3MF Assistant or Batch mode.

4. **Open FS 3MF Export**: In the "Virtuelle Köpfe" tab, click "💉 FS 3MF Export" in the
   second toolbar row.

5. **Select source file**: Click "📂 Öffnen" and select your `.3mf` model file. This is the
   file you already prepared in the slicer (with objects, supports, etc.) but without
   Mixed Filament assignments.

6. **Choose output**: Either overwrite the source (a `.bak` backup is created automatically)
   or save as a new file.

7. **Adjust layer height**: The layer height spinner is pre-filled from the main calculator.
   This must match your slicer's layer height setting.

8. **Preview** (optional): Click "🔍 Vorschau" to see the generated `mixed_filament_definitions`
   string before writing.

9. **Write**: Click "💉 Einschreiben". The tool modifies the ZIP and writes the new file.

### In FullSpectrum Slicer (Snapmaker_Orca)

10. **Open the file**: File → Open → select the modified `.3mf`.

11. **Verify Mixed Filaments**: Go to Filament → Mixed Filaments. You should see all virtual
    heads listed with their correct cadence/pattern settings.

12. **Assign to objects**: Select an object, use the "Paint" tool or right-click → "Assign
    Filament" to set which Mixed Filament (virtual head) each surface uses.

13. **Verify dithering settings**: Under Others → Dithering, confirm that the step size
    matches your layer height (0.08 mm recommended).

14. **Slice and print**.

---

## Troubleshooting

### Mixed Filaments not visible in slicer

**Cause**: The slicer did not pick up `mixed_filament_definitions` from the config.

**Solution**:
- Ensure you are using Snapmaker_Orca (FullSpectrum fork), not standard OrcaSlicer.
- Verify that `Metadata/project_settings.config` exists in the ZIP and contains the key.
- Use a ZIP viewer (e.g. 7-Zip) to inspect the `.3mf` contents.

### Backup file not created

**Cause**: The output path differs from the source path (using "Save as new file" mode).

**Note**: Backups (`.bak`) are only created when overwriting the source file. When saving to
a new file, the original is not touched.

### Wrong colors in slicer

**Cause**: The `filament_colour` array was not updated or the hex codes in the slots were
empty/placeholder.

**Solution**: Ensure all 4 physical slots have valid hex colors before exporting.

### Cadence heights seem wrong

**Cause**: The sequence contains unequal filament ratios that cause non-integer scaling.

**Solution**: The formula uses `round()` which may introduce small deviations. For exact
control, use sequences where the ratio of filament counts is a simple integer (e.g. 1:2, 1:3,
1:4 rather than 2:3).

### Pattern mode not working

**Cause**: Some older slicer versions do not support the 3–4 filament pattern token.

**Solution**: Update Snapmaker_Orca to the latest version. The manual pattern field was added
in v0.92+ of the FullSpectrum branch.

### File appears corrupted after export

**Cause**: The source `.3mf` may have been locked or in use by the slicer.

**Solution**: Close the slicer before exporting. The tool cannot write to a file that is
open in another application.

### ΔE values are good but printed colors look wrong

**Cause**: The simulation uses a linearized RGB averaging model. Real printed colors depend
on:
- Actual filament opacity (TD value)
- Print temperature and speed
- Layer adhesion and surface finish

**Solution**: Calibrate your TD values using the built-in TD Calibration tool. Print a test
swatch of alternating T1/T2 layers and measure the resulting color with a colorimeter or
photo comparison.

---

## Technical Notes

### Why 0.08 mm layer height is recommended

At 0.08 mm, individual color layers are thinner than the human eye's resolving power at
normal viewing distance. The colors blend optically rather than appearing as visible stripes.
At 0.12 mm or higher, striping becomes visible.

### Sequence length and tool changes

Each unique filament in the sequence requires a tool change per cycle. For a sequence of
length N with K unique filaments, each print layer requires K tool changes. Total tool changes
= (total layers) × K.

For a 200-layer print with a 5-layer sequence using 2 filaments:
- Cycles = 200 / 5 = 40
- Tool changes per cycle = 2
- Total tool changes = 40 × 2 = 80

The Tool Change Estimator in the app calculates this automatically.

### Stripe risk

FullSpectrum uses a phase-shift mechanism to avoid visible stripe patterns. The slicer
computes `phase_step = (cycle_length / 2) + 1` and shifts the starting position of each
layer by this amount. This works best when `cycle_length` and `phase_step` have no common
factors > 1.

The app warns about stripe risk automatically when it detects unfavorable sequence lengths.

### mixed_filament_definitions vs. OrcaSlicer filaments

In standard OrcaSlicer, virtual heads are separate filament profiles. In FullSpectrum
Slicer, they are Mixed Filaments derived from the physical T1–T4 slots. The
`mixed_filament_definitions` key is FullSpectrum-specific and is silently ignored by
standard OrcaSlicer.

Do not use FullSpectrum Direct 3MF Export if you plan to open the file in standard
OrcaSlicer — use the regular OrcaSlicer export instead.

---

## Related Resources

- [OrcaSlicer FullSpectrum fork](https://github.com/ratdoux/OrcaSlicer-FullSpectrum) — the slicer this tool targets
- [Snapmaker Forum](https://forum.snapmaker.com) — community support and print results
- [3MF Specification](https://3mf.io/specification/) — official 3MF format documentation
