---
name: visio-image-rebuilder
description: Rebuild or restyle editable Microsoft Visio diagrams from reference images and existing .vsdx files, then export deliverables as .vsdx, PNG, SVG, PDF, or PPTX. Use when the user asks Codex to open Visio, recreate a diagram from a PNG/JPG/screenshot/reference image, match a scientific model figure, update colors/typography/layout in a .vsdx, or make publication/presentation outputs while preserving Visio native editable shapes rather than embedding the reference image as a flat picture.
---

# Visio Image Rebuilder

## Core Rule

Recreate the reference as editable Visio content. Do not satisfy a rebuild request by inserting the whole reference image into the page. Embedding the reference image is only allowed as a temporary locked tracing layer if it is removed or hidden before delivery and the final `.vsdx` remains native shapes, text, connectors, and groups.

Treat `.vsdx` as the source of truth. Export PNG, SVG, PDF, or PPTX only after the Visio page has been rebuilt or restyled and checked for native editability.

## Workflow

1. Inspect inputs.
   - Confirm paths for the target `.vsdx`, reference image, requested output formats, and output directory.
   - Export the current Visio page to PNG before editing when a target file already exists.
   - Inspect the `.vsdx` package for pages, media entries, and shape counts.
   - Back up the target file before any write.
   - If Visio is already open, close only the target document or ask before terminating a stuck process.

2. Decode the reference image.
   - Identify page orientation, panel boundaries, module colors, captions, text hierarchy, arrows, dashed lines, and repeated motifs.
   - Build an object inventory: containers, titles, process boxes, icons, charts, graphs, equations, connectors, captions.
   - Decide whether the task is a full rebuild, color/style transfer, local edit, or export-only job.
   - For dense scientific figures, first create a coarse panel map, then draw panel internals. Do not start with small decorative details.

3. Prefer Visio automation for native edits.
   - Use COM automation on Windows when Visio is installed.
   - Use `DrawRectangle`, `DrawOval`, `DrawLine`, `Page.Import` only for small source assets, and shape cells such as `FillForegnd`, `LineColor`, `LineWeight`, `Char.Size`, `Char.Color`, `Rounding`.
   - Use explicit coordinates and IDs for fragile edits.
   - Keep grouped structure meaningful: major panels, submodules, repeated blocks, legends.

4. Use package XML edits only for narrow, deterministic changes.
   - XML patching is appropriate for recoloring existing shapes, replacing font tables, or changing known cell values.
   - Preserve Visio XML ordering: shape-level `Cell` nodes should be before `Section`, `Text`, or child `Shapes`.
   - Avoid rebuilding complex geometry by raw XML unless COM automation is unavailable.

5. Export requested formats from the verified Visio source.
   - Use `scripts/visio_page_tools.ps1` for export-only jobs.
   - Use `scripts/visio_rebuild_scaffold.ps1` with `-ExportFormats` for rebuilds that should immediately create deliverables.
   - Prefer SVG for vector web/manuscript handoff, PDF for review/print, and PPTX for presentation decks.
   - For PPTX, use PowerPoint COM when available; the generated slide contains the Visio page render, usually inserted from SVG.

6. Verify without overtrusting a single signal.
   - Export at least one preview after editing when possible.
   - Inspect the `.vsdx` package to confirm that no full-size reference PNG/JPG was left in `visio/media`.
   - Check shape count and representative text labels.
   - Check every requested output file exists and is non-empty.
   - If Visio automation hangs, stop safely, close the document if possible, and report whether the file was actually modified.

## Implementation Pattern

For full rebuilds, generate a script that:

- Opens the target `.vsdx` with Visio COM.
- Saves a timestamped or descriptive backup.
- Clears or duplicates the page depending on user preference.
- Sets page size to match the reference aspect ratio.
- Draws native shapes in top-left reference coordinates converted to Visio coordinates.
- Adds reusable helpers for rectangles, text boxes, ovals, lines, arrows, mini charts, graph nodes, and image-like stacks.
- Saves the document and exports requested formats.

Start from `scripts/visio_rebuild_scaffold.ps1` when building a full reconstruction script. Copy it into the workspace and customize the `Draw-ReferenceFigure` function rather than editing the skill copy.

For style transfer, generate a script that:

- Reads existing shape IDs, text, approximate geometry, fill, and line colors.
- Maps known modules to target palettes by text and group context.
- Applies fills, borders, text colors, line patterns, and font changes to existing shapes.
- Avoids repositioning unless the user asks for layout changes.
- Exports only after the `.vsdx` has been saved and inspected.

For export-only requests, run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\visio_page_tools.ps1 `
  -VsdxPath "C:\path\figure.vsdx" `
  -ExportFormats svg,pdf,pptx `
  -OutputDir "C:\path\exports" `
  -InspectPackage
```

For rebuild plus multi-format export, run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\visio_rebuild_scaffold.ps1 `
  -VsdxPath "C:\path\figure.vsdx" `
  -PageW 16 `
  -PageH 9 `
  -RefW 1600 `
  -RefH 900 `
  -PreviewPath "C:\path\exports\figure.png" `
  -ExportFormats svg,pdf,pptx `
  -OutputDir "C:\path\exports"
```

## Safety Checklist

- Back up before writing outside the workspace.
- Close any open Visio document that locks the target file before direct package edits.
- Never delete or revert unrelated user files.
- If a previous attempt embedded the reference image, restore from backup or remove the image shape before continuing.
- Tell the user clearly whether the final file is native editable shapes or a flat embedded image.
- Tell the user when PPTX export is a rendered slide rather than native PowerPoint shapes.

## Acceptance Criteria

A Visio rebuild is acceptable only when:

- Main panel positions, flow direction, captions, and module hierarchy match the reference at first glance.
- Text remains editable and uses a consistent academic font, usually Times New Roman.
- Repeated motifs are represented with reusable native shapes rather than pasted raster crops.
- The final `.vsdx` package has no full-page raster reference image in `visio/media`.
- Requested PNG/SVG/PDF/PPTX deliverables were exported from the saved `.vsdx` and are non-empty.
- A preview export or package inspection was performed, or the final response explicitly states why verification was skipped.

## Useful Resource

Use `scripts/visio_page_tools.ps1` for common inspection, backup, preview export, multi-format export, and package checks. Use `scripts/visio_export_formats.ps1` when a custom task script needs reusable export helpers. Use `scripts/visio_rebuild_scaffold.ps1` as the starting point for native-shape drawing scripts. Read `references/rebuild-guidelines.md` when a task requires a full figure reconstruction or careful one-to-one scientific diagram matching.
