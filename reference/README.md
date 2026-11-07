# ECMA-376 reference files

This directory holds copies of files from the ECMA-376 standard (Office Open XML File Formats). The engine implements the standard from these files. They are Ecma International material, not EigenPal code, and the repository's Apache-2.0 license does not apply to them.

## Source

All files come from the free downloads at [ECMA-376 on the Ecma International site](https://ecma-international.org/publications-and-standards/standards/ecma-376/), 5th edition, December 2016.

| Directory | Source archive | Inner archive | Changes |
| --- | --- | --- | --- |
| `ecma-376/part1/drawingml-geometries/` | `ECMA-376-1_5th_edition_december_2016.zip` | `OfficeOpenXML-DrawingMLGeometries.zip` | None. Byte-identical. |
| `ecma-376/part4/transitional-schemas/` | `ECMA-376-4_5th_edition_december_2016.zip` | `OfficeOpenXML-XMLSchema-Transitional.zip` | None. Byte-identical. |
| `ecma-376/part1/schemas/` | Copied from DOCX Editor's `reference/ecma-376/part1/schemas/` | Part 1 Strict schemas | None. The seven files are byte-identical to the 5th edition `OfficeOpenXML-XMLSchema-Strict.zip`. DOCX Editor's modified `dml-main.xsd` is not copied: use `part4/transitional-schemas/dml-main.xsd`. |

## Use

| Files | Used by |
| --- | --- |
| `presetShapeDefinitions.xml` | `scripts/geometry/build-presets.ts` generates `packages/ooxml-core/src/drawingml-preset-geometry.generated.ts` from it. The generated file carries the guide formulas as data. |
| `presetTextWarpDefinitions.xml` | `scripts/geometry/build-presets.ts` generates `packages/ooxml-core/src/drawingml-preset-text-warp.generated.ts` from it. Text layout maps glyph outlines between the warp paths. |
| `part4/transitional-schemas/*.xsd` | `scripts/validate-xsd.ts` (`bun run validate:xsd`), a test gate. Never bundled. |
| `part1/schemas/*.xsd` | Lookup only. No gate or build step reads them. |

The published packages contain only the generated geometry data, not these files.

## Quick references

`quick-ref/` holds short summaries written for this project, in the DOCX Editor format. They are EigenPal text, not Ecma material, so the Apache-2.0 license applies to them. Read them before you change parsing, cascade, or layout code.

| File | Covers |
| --- | --- |
| [`quick-ref/presentationml.md`](quick-ref/presentationml.md) | Package parts, slides, layouts, masters, placeholders, the text style cascade, units |
| [`quick-ref/drawingml.md`](quick-ref/drawingml.md) | Transforms, geometry, fills, lines, effects, shape styles, text bodies, pictures, tables |
| [`quick-ref/themes-colors.md`](quick-ref/themes-colors.md) | Theme colors, color maps, color transforms, theme fonts, format scheme (adapted from DOCX Editor) |

When a quick reference and the schema disagree, the schema wins. Fix the quick reference.

## Terms

Ecma International holds the copyright. Ecma's text copyright policy permits copies of a standard, in whole or in part, and derivative works that implement the standard in a conformant product, if the copies carry the Ecma copyright notice and its Copyright License and Disclaimer, and if the content is not modified. The notice is in [`ecma-376/NOTICE`](ecma-376/NOTICE). The published `core` package carries the same notice in its `THIRD_PARTY_NOTICES.md`, and its `LICENSE` says that the Ecma terms, not Apache-2.0, apply to the geometry data. The generated geometry file holds the full notice in a legal comment (`/*!`). Common bundlers keep legal comments or move them to a license file, so the notice also reaches the bundles of apps that use `core`.

Never edit a file in `ecma-376/`. To use a different version, replace the file with an unmodified Ecma copy and record its source in the table above.

## Known defect in the source

The 5th edition `presetShapeDefinitions.xml` defines `upDownArrow` twice and has no `upArrow`. The engine derives `upArrow` as a vertical mirror of `downArrow` (`MIRRORED_PRESETS` in `packages/ooxml-core/src/drawingml-geometry.ts`). The "Known approximations" section of `docs/site/content/pptx-fidelity.mdx` lists this.
