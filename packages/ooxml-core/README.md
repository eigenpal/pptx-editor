<p align="center">
  <a href="https://pptx-editor.dev/">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-dark.svg" />
      <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-light.svg" />
      <img src="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-light.svg" alt="PPTX Editor by EigenPal" width="320" height="90" />
    </picture>
  </a>
</p>

# @eigenpal/ooxml-core

The format-neutral OOXML layer of the EigenPal editors. It holds the code between the bytes of an OOXML (ECMA-376) file and a bounded XML tree, the text units, and the DrawingML color and geometry math. It has no PresentationML or WordprocessingML knowledge.

This package is private. `@pptx-editor.dev/core` inlines its code and its types at build time (`packages/core/tsup.config.ts`), so package consumers do not need a separate install. Do not import it from `@pptx-editor.dev/react` or from an example.

## What it holds

| Module | Purpose |
| --- | --- |
| `zip.ts` | Bounded zip read (entry count, total size, and ratio caps; refuses overlapping entries and bad limits) and deterministic write. `writeZip`, which PPTX Editor calls, writes in one call; `writeZipSteps` yields between slices of a long entry. Stores already-compressed media, and keeps a saved archive within the ratio cap |
| `zip-write.ts` | The archive layout and the deflate, slice by slice (`writeZipEntriesSteps`, `deflateInSlices`) |
| `opc-names.ts` | OPC part name and relationship target normalization; rejects traversal and encoded separators |
| `xml-reader.ts` | The bounded XML reader: rejects DTDs and entity declarations, caps depth and element counts |
| `xml-fast-read.ts` | A one-pass reader for the plain XML subset. It builds the bounded reader's tree, or gives the part to the bounded reader |
| `xml-text.ts` | Character data for both readers: entity decoding, attribute-value normalization, validity checks |
| `sinks.ts` | Checked XML escaping, `sanitizeHref`, CSS string escaping, inert executable content |
| `limits.ts` | Resource budgets and the hard ceilings no host can raise |
| `sha256.ts` | Synchronous SHA-256 for part and font fingerprints |
| `image-size.ts` | Bounded PNG, JPEG, GIF, and BMP header reads |
| `grapheme.ts` | Grapheme clusters (UAX #29) over `Intl.Segmenter` at the invariant locale, through a replaceable boundary: the unit a caret steps and deletes by |
| `word-segment.ts` | UAX #29 word segments (with dictionary words in Thai, Lao, Khmer, Myanmar, and CJK text) for word steps, word deletion, and double-click |
| `caret-word.ts` | The caret word rule: the word a format at a collapsed caret reaches (`caretWordRange`) |
| `drawingml-color.ts` | The DrawingML color models and transforms (ECMA-376 Part 1, 20.1.2.3): sRGB and linear components, HSL, and `resolveDrawingColor` (`drawingColorBase` for the six color models, then `applyColorTransforms`). It takes a neutral `DrawingColorElement` (local name, attribute reader, transforms), so it works over either editor's tree, and reads both spellings of a percentage (`50000` and `50%`). The caller resolves scheme color names |
| `drawingml-preset-colors.ts` | `PRESET_COLORS`: the `ST_PresetColorVal` names and their sRGB values, in a null-prototype record |
| `drawingml-geometry.ts` | The DrawingML geometry evaluator (ECMA-376 Part 1, 20.1.9): guide formulas (`evaluateFormula`), path commands with arcs as cubic Beziers, the text rectangle, and the caps on guides and path commands. `evaluatePresetGeometry` draws a preset with its adjust values (an unknown name draws a rectangle and is reported), `evaluateGeometry` draws any definition, and `customGeometryDefinition` reads `a:custGeom` through a neutral `DrawingGeometryElement` (namespace, local name, attribute reader, child elements), so it works over either editor's tree. Only `a:moveTo` sets the current point |
| `drawingml-preset-geometry.generated.ts`, `drawingml-preset-geometry-types.ts` | `PRESET_SHAPES`: every ECMA-376 preset as guide formulas and paths, generated from `presetShapeDefinitions.xml` by `scripts/geometry/build-presets.ts` (Ecma material, see `reference/ecma-376/NOTICE`), and its types |

The two DrawingML color modules are PPTX Editor's own (moved from `packages/core/src/cascade/`), not copies. DOCX Editor resolves the same colors in `packages/core/src/store/package/drawing-shape-color.ts`, with other math. The header of `drawingml-color.ts` lists the fixes, measured on reference PDFs, that the DOCX Editor team can take, and the rules of the DOCX Editor module that PPTX Editor does not have yet. `bun run test:docx-shared` does not cover them, because they have no DOCX Editor source to compare. The PresentationML glue (the color map, `phClr`, and the canonical tree) stays in `packages/core/src/cascade/color.ts`.

The DrawingML geometry modules are PPTX Editor's own too (moved from `packages/core/src/layout/geometry/`). DOCX Editor evaluates the same definitions in `packages/core/src/store/package/drawing-geometry-engine.ts` and `drawing-geometry-formula.ts`, and flattens curves to polylines. PPTX Editor keeps exact Beziers, so its module is the base. The header of `drawingml-geometry.ts` lists the fixes that the DOCX Editor team can take (the path-start rule and bounded arc sweeps), and the rules of the DOCX Editor module that PPTX Editor does not have. The PresentationML glue (the cascade's `GeometrySpec` and the canonical tree) stays in `packages/core/src/layout/geometry/geometry.ts`.

The text modules are copies of DOCX Editor's `packages/core/src/layout/grapheme.ts`, `word-segment.ts`, and the format-neutral part of `caret-word.ts` (its layout and revision readers stay in DOCX Editor). Every other module except `image-size.ts` and the DrawingML color and geometry modules is a copy of the DOCX Editor module of the same name (`docx-editor/packages/core/src/store/package/`, and `store/runtime/limits.ts`). The first lines of each file name the DOCX Editor path, its commit, and the IDs of the fixes it holds that DOCX Editor does not have yet. Below that header, the module is the DOCX Editor file at that commit plus those fixes, apart from line wrapping and three recorded differences: `opc-names.ts` leaves out the WordprocessingML constant `WML_MAIN_DOCUMENT_PART`, `sinks.ts` keeps `escapeXmlAttribute` public, and `xml-reader.ts` imports `./limits.ts`. `limits.ts` is copied in full, but PPTX Editor reads only `maxElementCount` from it; the zip caps come from `DEFAULT_ZIP_LIMITS` in `zip.ts`. Comments copied from DOCX Editor can name DOCX Editor files and decisions (for example "design D14" or `font-resource.ts`); they mean the DOCX Editor ones. A fix for the DOCX Editor team goes in the commit message and in the file header.

## Rules

- No imports from `@pptx-editor.dev/*` and no PresentationML names. If a module needs either, it belongs in `packages/core`.
- DOM-free. `bun run check:lane-boundaries` compiles `src/` without the DOM library.
- A fix here is a security fix for both products. Name it as a fix for the DOCX Editor team in the commit message and in the file header.
- Keep each copied module equal to its DOCX Editor file. `bun run test:docx-shared` runs DOCX Editor's tests of these modules against this package and reports any drift (local only; it needs a DOCX Editor checkout).
- Put a test of one module in `src/__tests__/`. Put a test that needs a deck, the package model, or a fixture in `packages/core/src/store/__tests__/`.

## Sharing plan

1. The package is internal to PPTX Editor. The modules keep the exported API of their DOCX Editor sources.
2. The drift is reconciled module by module, and the DOCX Editor tests pass against this package. The fixes that DOCX Editor does not have yet are listed for the DOCX Editor team.
3. The DrawingML color transforms are in the package (`drawingml-color.ts`).
4. The DrawingML geometry evaluator and the preset table are in the package (`drawingml-geometry.ts`).
5. Next: move the HarfBuzz measurer. Then publish the package and make both editors depend on it.
