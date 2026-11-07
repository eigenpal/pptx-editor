# @pptx-editor.dev/pptx-to-pdf

## 0.3.0

### Minor Changes

- dfb5df0: EMF and WMF pictures draw as pictures, not as placeholders. The editor converts each one to a PNG in the background with `emf-converter`, which loads only for a deck that has a metafile. The placeholder shows until the PNG is ready, and then the editor repaints. The picture file stays unchanged.

  - `session.picturesReady(index)` waits until the EMF and WMF pictures of a slide, or of every slide, have converted. It resolves `false` when a picture cannot convert or is not ready after 30 seconds. PNG and ZIP export wait for it. Call it before `exportSlideSvg()`.
  - `session.pictureStatus(index)` lists each EMF and WMF picture as `ready`, `pending`, `failed`, or `timed-out`.
  - `picturesTimeoutMs` sets how long PNG and ZIP export wait for the pictures of each slide.
  - An OLE object whose preview is an EMF or WMF picture draws the preview.
  - PDF export draws EMF and WMF pictures with the same converter. On Node.js it draws with `@napi-rs/canvas` when that package is installed, and with its own rasterizer otherwise, so the pixels can differ from the browser's.
  - A WebP picture must pass the same size limits as other pictures before it draws.
  - The natural size of an EMF or WMF picture comes from its picture frame.

  Damaged and unusual files open and save more safely:

  - A save keeps the grid columns of a table whose rows have more cells than columns. Extra cells with text stay in the file.
  - In the editor and in image export, a shape whose picture fill is missing or damaged draws the picture placeholder, not nothing.
  - A damaged notes part, or a damaged theme of the notes master, drops the notes and no longer refuses the deck. A damaged theme of the handout master no longer refuses the deck.
  - Edits, renames, deletes, and layout changes find every copy of a shape in nested `mc:AlternateContent`.

### Patch Changes

- Updated dependencies [8253ddb]
- Updated dependencies [30dc48e]
- Updated dependencies [a90b74e]
- Updated dependencies [a90b74e]
- Updated dependencies [a90b74e]
- Updated dependencies [dfb5df0]
- Updated dependencies [8253ddb]
- Updated dependencies [c47f5ce]
- Updated dependencies [30dc48e]
- Updated dependencies [8253ddb]
- Updated dependencies [a90b74e]
- Updated dependencies [a90b74e]
- Updated dependencies [8253ddb]
- Updated dependencies [9d3e18e]
- Updated dependencies [7666155]
- Updated dependencies [8253ddb]
- Updated dependencies [cbee759]
- Updated dependencies [8253ddb]
- Updated dependencies [8253ddb]
- Updated dependencies [8253ddb]
- Updated dependencies [30dc48e]
  - @pptx-editor.dev/core@0.3.0
  - @pptx-editor.dev/pro@0.3.0

## 0.2.0

### Minor Changes

- db9407d: Draw diagrams (`dgm:` graphic frames) from their saved drawing in the editor and in PDF export. A diagram stays read-only and selects as one object. A drawing past 512 shapes or 8 group levels, or past 4096 shapes and groups on one slide, shows the placeholder. A diagram with no saved drawing still shows a placeholder, and the PDF export still reports `frame.diagram` for it. `ResolvedShape.textXfrm` gives the text rectangle of a diagram shape, and the layout `frame` item gives the laid-out drawing as `drawing`.
- 136b3b5: Draw dashes and line ends as the reference PDFs show them, in the editor and in PDF export. Line ends on a line thinner than 2 pt take the size of a 2 pt line. Triangles, stealth arrows, open arrows, ovals, and diamonds take their measured sizes, and the line stops under each head at its measured point. A preset dash with square caps keeps its flat length, and only the ends of the line get the cap. The PDF export now draws dashes and line ends, and no longer reports `line.dash` or `line.ends`. The `arrowhead-joins` approximation is gone. A custom dash keeps at most 64 dash and space pairs, a negative length counts as 0, and a pattern with no length draws a solid line.
- a3afcdd: Charts now draw from the values cached in their chart parts when `@pptx-editor.dev/pro/drawing` is installed (`drawingModule({ licenseKey })` or `installDrawing()`). Bar, column, line, area, scatter, pie, doughnut, and radar charts draw with titles, legends, axes, gridlines, data labels, and number formats. Other chart types, and every chart without the module, keep the placeholder. `@pptx-editor.dev/pptx-to-pdf` draws charts in its PDFs, and reports `frame.chart` only for the chart types the module leaves out. The core adds the `chart` renderer key to `installDrawingRenderers`, the resolved `chart` part of a chart frame, the laid-out `chart` items of a frame, and `drawingCapabilities().chart`.
- c292536: Glow, inner shadow, reflection, and soft edge now paint through the new `@pptx-editor.dev/pro/drawing` entry, as in DOCX Editor. Register `drawingModule({ licenseKey })` in `modules`, or call `installDrawing()` for headless rendering. Without it, a shape or picture paints without these four effects; outer shadows still paint, and the effect commands and save keep every effect. `@pptx-editor.dev/pptx-to-pdf` always installs the module, and its PDFs now draw reflections with their fade (`effect.reflection-blur` reports the blur that is left out). `@pptx-editor.dev/core/output` no longer exports the internal `glowSpread` and `innerShadowShift`.

### Patch Changes

- 5a03952: Published packages no longer ship source maps, so a stack trace points into `dist/`. Every JavaScript file of `pro`, `pptx-to-pdf`, and `editor-api` starts with its license header, and these packages are minified. The preset shape geometry in `core` carries the Ecma copyright notice in a legal comment, so bundlers that keep legal comments carry the notice into your app bundle.
- 0593988: Load the packaged substitute fonts under Node, so headless layout and PDF export measure and draw text with real font bytes. `fontReport.substituteFailures` and `fontResolution.substituteFailures` list each face that fails to load as `{ family, file, diagnostic }`, and the PDF export reports each one as a document-wide `font.substitute-failed` diagnostic. The PDF export now draws single and double underline and strikethrough, and its `font.missing` diagnostic names the font. The editor also opens a deck in a DOM without the CSS Font Loading API.
- cd341f2: Each package now accepts only the same minor version of `@pptx-editor.dev/core` as a peer, for example `~0.1.0`. The links in the package READMEs now open from npm.
- Updated dependencies [4594311]
- Updated dependencies [98b85e2]
- Updated dependencies [0057839]
- Updated dependencies [db9407d]
- Updated dependencies [7d4b0be]
- Updated dependencies [5a11387]
- Updated dependencies [bfca3d6]
- Updated dependencies [159eb87]
- Updated dependencies [0653436]
- Updated dependencies [5a03952]
- Updated dependencies [136b3b5]
- Updated dependencies [6bed74e]
- Updated dependencies [0593988]
- Updated dependencies [b906153]
- Updated dependencies [a3afcdd]
- Updated dependencies [c292536]
- Updated dependencies [c292536]
- Updated dependencies [60ed234]
- Updated dependencies [60ed234]
- Updated dependencies [84256e4]
- Updated dependencies [0653436]
- Updated dependencies [60c621b]
- Updated dependencies [6bed74e]
- Updated dependencies [db11b83]
- Updated dependencies [14a9bc9]
- Updated dependencies [1358a9d]
- Updated dependencies [aa99dad]
- Updated dependencies [cd341f2]
  - @pptx-editor.dev/core@0.2.0
  - @pptx-editor.dev/pro@0.2.0

## 0.1.0

### Minor Changes

- 457a754: East Asian text measures and draws with each run's `a:ea` font (complex-script text with `a:cs`, symbol characters with `a:sym`), and a theme with an empty `a:ea` uses its script font for the run's language. `eaLnBrk="0"` now breaks East Asian text between any two characters with no kinsoku rule, `latinLnBrk="1"` breaks inside Latin words, Korean text breaks at spaces, `hangingPunct` lets a comma or full stop hang past the line end, the presentation's `p:kinsoku` lists replace the built-in lists, and justified East Asian lines spread between characters, as the reference PDFs show. A missing East Asian font falls back to a host CJK face for the text's language, such as Noto Sans CJK JP, and a character that the run's face lacks uses another loaded face. The canvas shows the text that an input method composes at the caret, and the key that ends a composition runs no shortcut. `@pptx-editor.dev/i18n` adds Japanese (`/ja`), Korean (`/ko`), and Simplified Chinese (`/zh-CN`) catalogs, and `@pptx-editor.dev/pptx-to-pdf` draws East Asian text with Noto Sans CJK JP when the optional `@docx-editor.dev/fonts-cjk` is installed.
- 412d750: Pictures can be cropped, changed, reset, adjusted, styled, sized, and given alt text through new commands, the toolbar's `image` group, and the `imageAltText` and `imageProperties` popups, and picture borders now draw outside the picture.
- 412d750: `openDocumentForExport` and `exportPdfFrom` export one deck more than once from one layout, `ExportResourceError` reports a disposed or aborted session, each PDF result maps its pages to deck slide numbers in `slideNumbers`, and an engine failure inside an export is reported as `PdfEncodingError`.
- 412d750: The Vue adapter, the pro modules for review and collaboration, and PDF export are now published with the other packages.

### Patch Changes

- 412d750: Draw shadow blur, glow, and soft edges at their measured size, put an inner shadow on the side its direction names, and keep a shape's outline inside its effects.
- 412d750: An embedded OLE object now shows the PNG, JPEG, GIF, or BMP picture of its `p:oleObj/p:pic` in the frame's box, in place of a labeled placeholder.
- 412d750: The PDF export draws scaled linear gradients, circle path gradients, and two-stop gradient colors with the same rules as the editor.
- 412d750: `pptx-to-pdf` now draws cropped pictures, linear and path gradients (gradient text too), tables, outer and inner shadows, and glow, and reports each approximation with a diagnostic code; the codes `effects`, `picture.crop`, `path.fill-mode`, `fill.linear`, `fill.path`, and `fill.group` are no longer emitted, and `frame.table` only for a table frame without a grid. Many effects or gradients on one page now stay inside the paint budget and `timeoutMs`.
- 412d750: Clarify package documentation and host responsibilities for collaboration access.
