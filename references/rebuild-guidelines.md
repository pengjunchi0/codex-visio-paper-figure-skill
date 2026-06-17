# Rebuild Guidelines

Use this reference when reconstructing a scientific, technical, or academic diagram from a PNG/JPG/screenshot into editable Microsoft Visio `.vsdx` content.

The goal is visual and semantic equivalence using native Visio shapes, text, connectors, groups, and styles. Do not treat any example below as a fixed domain template. Use `.vsdx` as the editable source, then export SVG/PDF/PPTX deliverables from that source.

## Reference Analysis

Capture these properties before drawing:

- Canvas aspect ratio, page margins, and whitespace distribution.
- Major regions: bands, panels, columns, rows, insets, legends, captions.
- Reading order: left-to-right, top-to-bottom, radial, cyclic, or feedback loop.
- Text hierarchy: figure title, panel title, module label, axis label, equation, caption, note.
- Connector semantics: solid arrows, dashed arrows, inhibition lines, braces, loops, feedback paths, summation nodes.
- Repeated motifs: stacked frames, process blocks, neural-network layers, tables, graphs, charts, heatmaps, timelines, molecule-like nodes, icons.
- Style tokens: font family, font size bands, accent colors, fill opacity, line weights, corner radii, dash patterns.
- Output needs: `.vsdx` editability, SVG vector handoff, PDF review/print, PPTX slide handoff.
- Editability risks: tiny text, dense textures, screenshots embedded inside shapes, equations, image-like scientific data.

## Panel Inventory Template

For a complex figure, write a short inventory before coding. Keep it specific to the user's reference image, not to any previous project.

```text
Canvas: landscape or portrait; approximate aspect ratio; main whitespace pattern.
Global structure: top band / bottom band / left workflow / right modules / multi-panel grid.
Panel A: title, accent color, role, main internal objects, incoming/outgoing arrows.
Panel B: title, accent color, role, main internal objects, incoming/outgoing arrows.
Panel C: title, accent color, role, main internal objects, incoming/outgoing arrows.
Shared elements: legends, captions, separators, equations, repeated labels.
Critical text: labels that must be preserved exactly.
Approximation targets: dense details that can be represented by native simplified motifs.
Output formats: vsdx plus png/svg/pdf/pptx as requested.
```

This inventory is the contract for the drawing script. If an item is not in the inventory, it is likely to be omitted or drawn inconsistently.

## Common Figure Archetypes

Choose the closest archetype before drawing:

- **Pipeline or workflow**: ordered process boxes, arrows, inputs, outputs, optional feedback.
- **Model architecture**: repeated blocks, encoders/decoders, latent variables, modules, loss paths.
- **Algorithm schematic**: data structures, operations, equations, constraints, branching logic.
- **Multi-panel methods figure**: several modules with different colors and local legends.
- **Graph or network figure**: nodes, edges, weights, message passing, clusters, highlighted paths.
- **Matrix or heatmap figure**: grids, color scales, row/column labels, aggregation arrows.
- **Chart-heavy figure**: mini plots, axes, trends, bars, curves, distributions, annotations.
- **System diagram**: components, interfaces, data stores, users, protocols, deployment zones.

The archetype only guides layout and helper functions. The actual labels, colors, and structure must come from the user's reference image.

## Coordinate Strategy

Use a reference coordinate system in pixels or normalized units, then convert to Visio page inches:

```powershell
$PageW = 16.0
$PageH = 12.0
$RefW = 1448.0
$RefH = 1086.0
function VX([double]$x) { $PageW * $x / $RefW }
function VY([double]$y) { $PageH - ($PageH * $y / $RefH) }
```

Draw from top-left bounds as `RectTL(x, y, w, h)` so the script remains readable against screenshots.

## Visual Matching Heuristics

- Match structure before decoration. Panel placement, reading order, and flow arrows matter more than texture detail.
- Preserve semantic grouping. Containers, modules, legends, charts, and subgraphs should be separate editable groups.
- Preserve scientific labels exactly when legible. If text is unreadable, use a close placeholder and report the limitation.
- Use simplified native motifs for dense image-like content: stacked rounded rectangles, small dots, mini heatmaps, line charts, bar charts, tables, and graph nodes.
- Keep equations editable as text when practical. Use plain text approximations if Visio equation objects are unavailable.
- Use color as semantic grouping, not decoration. Reuse one accent per panel or module unless the reference clearly uses another scheme.
- Prefer native approximations over screenshots for cubes, graphs, heatmaps, icons, and small charts. The target is editability plus visual equivalence.
- If the reference contains many tiny repeated elements, create 2-3 reusable helper functions rather than hand placing each occurrence.
- Do not force a previous project's palette, module names, or layout onto a new reference.

## Output Format Strategy

- Keep `.vsdx` as the editable master and export all other formats from the saved Visio page.
- Use PNG as a fast visual preview and for compatibility checks.
- Use SVG when the user needs vector graphics for manuscripts, web pages, or post-processing in vector tools.
- Use PDF for review, print, or submission preview. Export from the Visio document after save.
- Use PPTX for presentation handoff. Prefer a PowerPoint slide containing a full-slide SVG render of the Visio page; report that this is a slide render, not guaranteed native PowerPoint shape decomposition.
- Do not regenerate separate drawings per format. Differences between formats should come from export behavior, not divergent source artwork.

## Scientific Figure Style Tokens

Use these defaults unless the reference clearly differs:

- Font: Times New Roman for manuscript-style figures; Arial or Helvetica for clean technical UI-style diagrams.
- Main border: 0.9-1.2 pt.
- Internal border: 0.5-0.8 pt.
- Rounded rectangle radius: small, usually 4-8 px equivalent.
- Arrowheads: consistent filled end arrows for forward flow; dashed lines for feedback, optional paths, or constraints.
- Backgrounds: white or very pale panel tints.
- Captions: black, bold, centered; keep below the relevant panel or figure band.
- Accent colors: muted blue, green, orange, purple, cyan, or gray families; avoid over-saturated fills unless present in the reference.

## Reconstruction Order

1. Page size and canvas background.
2. Major bands, panel boxes, captions, and separators.
3. Main dataflow arrows and module containers.
4. Text-bearing process boxes.
5. Repeated motifs: stacks, cubes, graphs, heatmaps, mini charts, tables, icons.
6. Equations, legends, axis labels, small annotations.
7. Grouping, font normalization, preview export, package inspection, requested SVG/PDF/PPTX export.

Do not optimize small graphics before the panel grid and arrows are correct.

## Avoid These Failure Modes

- Inserting the whole reference image as the final page.
- Copying examples from this skill instead of analyzing the user's actual reference.
- Using a single ungrouped mass of hundreds of shapes with no structure.
- Recoloring by global color replacement when modules share colors with different meanings.
- Leaving Visio locked in the background after a timeout.
- Claiming success after a script times out without checking file timestamp or package contents.
- Letting text overflow boxes after changing fonts.
- Losing semantic editability by converting equations, graph nodes, tables, or charts into pasted crops.
- Making all modules the same palette when the reference uses color to distinguish submodules.
- Exporting SVG/PDF/PPTX from a stale file before saving the final `.vsdx`.

## Verification Rubric

Score the result before delivery:

- Structure: panel locations and flow match the reference.
- Semantics: every named module, caption, and important label exists as editable text.
- Style: colors, line weights, typography, and spacing are consistent.
- Editability: no full-page image; major objects are native shapes.
- Outputs: requested PNG/SVG/PDF/PPTX files exist, are non-empty, and were exported from the same saved `.vsdx`.
- Robustness: target file has a backup; preview or package checks are recorded.

If any category is weak, either fix it or state the limitation explicitly.

## Delivery Expectations

Final response should include:

- Target `.vsdx` path.
- Backup path.
- Preview path if exported.
- SVG/PDF/PPTX paths if requested.
- Whether the `.vsdx` file is native editable Visio shapes.
- Any caveats: unreadable labels, skipped verification, PPTX render limitations, or Visio automation issues.
