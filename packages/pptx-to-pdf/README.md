<p align="center">
  <a href="https://pptx-editor.dev/">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-dark.svg" />
      <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-light.svg" />
      <img src="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-light.svg" alt="PPTX Editor by EigenPal" width="320" height="90" />
    </picture>
  </a>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/@pptx-editor.dev/pptx-to-pdf"><img src="https://img.shields.io/npm/v/@pptx-editor.dev/pptx-to-pdf.svg?style=flat-square&color=C43E1C" alt="npm version" /></a>
  <a href="https://www.npmjs.com/package/@pptx-editor.dev/pptx-to-pdf"><img src="https://img.shields.io/npm/dm/@pptx-editor.dev/pptx-to-pdf.svg?style=flat-square&color=C43E1C" alt="npm downloads" /></a>
  <a href="https://github.com/eigenpal/pptx-editor/blob/main/packages/pptx-to-pdf/LICENSE.md"><img src="https://img.shields.io/badge/license-EigenPal_Pro_License-blue.svg?style=flat-square&color=C43E1C" alt="EigenPal Pro License" /></a>
</p>

# @pptx-editor.dev/pptx-to-pdf

Convert a `.pptx` to PDF on Node.js, with searchable text and embedded fonts. The package runs the PPTX Editor engine headlessly: `openPresentation` lays out each slide exactly as the editor does, and this package paints that layout into PDF pages. It is not a second renderer.

The package is a foundation for Node.js. It uses the EigenPal Pro License (`LICENSE.md`). Evaluation is permitted. Production use requires a commercial agreement.

For task guides and API details, see the [developer documentation](https://pptx-editor.dev/docs/export/pdf).

## Use

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { exportPdf } from '@pptx-editor.dev/pptx-to-pdf';

const pptx = await readFile('deck.pptx');
const { bytes, diagnostics } = await exportPdf(pptx, { fidelityPolicy: 'best-effort' });
await writeFile('deck.pdf', bytes);
for (const d of diagnostics) if (d.severity !== 'information') console.warn(d.message);
```

The default `fidelityPolicy` is `strict`. In strict mode, the conversion throws `PdfFidelityError` when a slide has content the PDF cannot reproduce, for example `slide 3: fill.pattern (+4 more)`. The error carries every diagnostic. Use `best-effort` to get the PDF and read `diagnostics` instead.

## Reuse a layout

Use `openDocumentForExport` and `exportPdfFrom` to export one deck more than once from one layout, as with DOCX Editor's `docx-to-pdf`. The session opens the deck and resolves its fonts once, and each slide is laid out once. Dispose the session after all exports, including failed exports.

```ts
import { readFile } from 'node:fs/promises';
import { exportPdfFrom, openDocumentForExport } from '@pptx-editor.dev/pptx-to-pdf';

const pptx = await readFile('deck.pptx');
const opened = await openDocumentForExport(pptx);
if (!opened.ok) throw new Error(`Cannot open the deck: ${opened.reason}`);
try {
  const shown = await exportPdfFrom(opened.session, { fidelityPolicy: 'best-effort' });
  const all = await exportPdfFrom(opened.session, {
    fidelityPolicy: 'best-effort',
    includeHiddenSlides: true,
  });
} finally {
  opened.session.dispose();
}
```

`openDocumentForExport` takes the font options (`fonts`, `substitutes`) and a `signal` that controls the whole session. It returns `{ ok: true, session }`, or `{ ok: false, reason, detail? }` for a deck that core refuses or an aborted signal. `exportPdfFrom` takes the output options: `fidelityPolicy`, `includeHiddenSlides`, `timeoutMs`, `maxPages`, `maxOutputBytes`, and `signal`. An option given to the wrong call throws `TypeError`. A signal given to `exportPdfFrom` stops only that export. A failed or cancelled export does not dispose the session. After `dispose()`, or after the session's signal aborts, an export rejects with `ExportResourceError` (`disposed` or `aborted`). The result's `timings.openMs` is `0`, because the call does not open the deck.

## Errors

Every error class has a stable `code`. Check the class first, then `code`: `PdfWorkLimitError` and `ExportResourceError` both use `aborted`, and two classes extend others.

```ts
import { exportPdf, PdfFidelityError, PdfWorkLimitError } from '@pptx-editor.dev/pptx-to-pdf';

try {
  await exportPdf(pptx, { timeoutMs: 30_000 });
} catch (error) {
  if (error instanceof PdfFidelityError) console.warn(error.diagnostics);
  else if (error instanceof PdfWorkLimitError && error.code === 'timeout') console.warn('Too slow');
  else throw error;
}
```

| Class | `code` | When |
| --- | --- | --- |
| `PdfDocumentOpenError` | `documentOpenFailed` | The deck is not a valid PPTX. `reason` is core's open error reason. |
| `PdfFidelityError` | `fidelityUnsupported` | `strict` policy, and a slide has content the PDF cannot reproduce. `diagnostics` lists it. |
| `PdfPageLimitError` | `pageLimitExceeded` | The export has more pages than `maxPages`. Extends `RangeError`. |
| `PdfOutputLimitError` | `outputTooLarge` | The PDF is larger than `maxOutputBytes`. Extends `PdfEncodingError`, so catch it first. |
| `PdfEncodingError` | `encodingFailed` | An unexpected failure while writing the PDF. `cause` holds the original error. |
| `PdfWorkLimitError` | `timeout`, `aborted`, `operations` | The export passed `timeoutMs`, its `signal` aborted, or it passed the paint budget: paint operations, or content bytes, which include the pictures' file bytes. |
| `ExportResourceError` | `disposed`, `aborted` | The session was disposed, or the session's own `signal` aborted. |
| `TypeError`, `RangeError` | none | An invalid argument or option, before any work starts. `exportPdf` and `openDocumentForExport` ignore option names they do not know; `exportPdfFrom` refuses them. `PdfPageLimitError` is the one `RangeError` that comes later. |

## Diagnostic codes

Each diagnostic has a fixed `code`. Handle an unknown code by its `severity`. The default `strict` policy refuses a deck with any `unsupported` or `approximation` diagnostic, so a deck with a blurred shadow or glow, or a picture with a shadow, needs `best-effort`.

| Code | Severity | Meaning | DOCX Editor code for the same idea |
| --- | --- | --- | --- |
| `fill.pattern`, `fill.picture` | `unsupported` | A pattern fill, or a picture fill on text, is not drawn. |  |
| `line.gradient` | `approximation` | A gradient on a line is drawn in its middle color. |  |
| `fill.outside-box` | `approximation` | A shape's picture fill or rectangular path gradient paints only the shape's box; geometry more than 0.75 pt past it is empty. |  |
| `fill.gradient-stops` | `approximation` | A gradient with more than 64 stops is drawn with 64. |  |
| `picture.missing`, `picture.format`, `picture.unreadable`, `picture.too-large` | `unsupported` | A picture part is missing or external, in a format other than PNG, JPEG, EMF, or WMF, unreadable, or over the pixel limit. |  |
| `picture.metafile` | `unsupported` | An EMF or WMF picture does not convert to a PNG within the limits. |  |
| `picture.svg-fallback` | `information` | An SVG picture is drawn from its PNG fallback, the raster copy that the file stores beside it. |  |
| `picture.tile` | `approximation` | A tiled picture is drawn stretched. |  |
| `picture.color-effects` | `approximation` | Picture color effects are not applied. | `image-effects` |
| `line.compound` | `approximation` | Compound lines are drawn as one line. |  |
| `effect.shadow-blur`, `effect.inner-shadow-blur`, `effect.glow` | `approximation` | The blur is drawn as layered rings. |  |
| `effect.inner-shadow` | `approximation` | An inner shadow starts at the fill, so an outline wider than 1 pt, an unfilled outline, and overlapping paths are shaded differently. |  |
| `effect.silhouette` | `approximation` | A shadow or glow uses an opaque silhouette: picture and fill transparency do not shape it. |  |
| `effect.soft-edge` | `approximation` | Soft edges are drawn hard. |  |
| `effect.reflection-blur` | `approximation` | A reflection is drawn without its blur. |  |
| `effect.not-rendered` | `unsupported` | An effect the engine does not render (`a:presetShdw`, `a:blur`, and others). |  |
| `text.outline` | `unsupported` | Text outlines are not drawn. |  |
| `text.underline-style` | `approximation` | A dotted, dashed, wavy, heavy, or word-only underline is drawn as a single line, as in the editor. |  |
| `text.synthetic-style`, `text.upright`, `text.missing-glyph` | `approximation` | Synthetic bold and italic, upright vertical characters, and characters the face lacks. | `missing-glyph` |
| `font.missing`, `font.not-embedded` | `unsupported` | A face has no bytes, or may not be embedded; its text is not drawn. The message names the font. | `font-embedding` |
| `font.substitute-failed` | `approximation` | A packaged substitute face did not load, so its text measures and draws with another face. The diagnostic is document-wide: it has no slide or page. `fontResolution.substituteFailures` lists each failure as `{ family, file, diagnostic }`. | `font-origin-failed` |
| `frame.chart`, `frame.diagram`, `frame.ole`, `frame.unknown` | `unsupported` | A chart type outside the drawing module (3-D, of-pie, stock, bubble, surface, or a chart with trendlines or error bars), a diagram with no saved drawing, an OLE object with no PNG, JPEG, GIF, or BMP picture, or another graphic frame is not drawn. Tables, OLE pictures, other charts, and diagrams with a saved drawing draw. A `frame.diagram` message says why the diagram is not drawn: it has no saved drawing, its drawing holds content that is not drawn yet, or the drawing is past a shape, nesting, or slide limit. | `drawing` |
| `frame.table` | `unsupported` | A table frame with no grid (no `a:tbl`, or no rows or columns) is not drawn. |  |
| `geometry.unknown-preset` | `approximation` | An unknown preset geometry is drawn as a rectangle. |  |
| `slide.not-laid-out` | `unsupported` | The slide could not be laid out and has no page. |  |
| `layout.approximation`, `font.substituted` | `information` | An engine approximation (see Known approximations on the fidelity page of the docs site) or a substituted font. | Synthetic missing-font deck in export tests (for `font.substituted`) |

## API

The names follow DOCX Editor's `@docx-editor.dev/docx-to-pdf` where the concept is the same.

| Member | What it does |
| --- | --- |
| `exportPdf(source, options?)` | Converts PPTX bytes. Returns `{ bytes, pageCount, slideNumbers, fontResolution, diagnostics, timings }`. |
| `openDocumentForExport(source, options?)` | Opens a reusable session. Returns `OpenPdfDocumentForExportResult`: `{ ok: true, session }` or `{ ok: false, reason, detail? }`. |
| `exportPdfFrom(session, options?)` | Converts from an open session, without reopening the deck. Returns the same result as `exportPdf`. |
| `PdfExportOptions` | `fonts` (one `FontSource` or an array), `substitutes`, `includeHiddenSlides`, `fidelityPolicy`, `timeoutMs`, `maxOutputBytes`, `maxPages`, `signal` |
| `OpenPdfDocumentForExportOptions`, `PdfProjectionOptions` | The session options (`fonts`, `substitutes`, `signal`) and the output options of `exportPdfFrom` |
| `PdfExportSession` | `fontResolution` and `dispose()` |
| `PdfExportResult.slideNumbers` | The one-based deck slide number of each page. Hidden slides (unless `includeHiddenSlides` is true) and slides that cannot be laid out have no page. |
| `PdfDiagnostic` | `code` (a fixed string), `message`, `severity` (`unsupported`, `approximation`, `information`), `slideIndex`, `slideNumber`, and `pageIndex`, `pageNumber` of the output page. A document-wide diagnostic has no slide or page. |
| `PdfFontResolution` | The faces the conversion drew with, the substituted families, the embedded faces the deck may not use, and the substitute faces that failed to load |
| Errors | `PdfDocumentOpenError`, `PdfFidelityError`, `PdfPageLimitError`, `PdfOutputLimitError`, `PdfEncodingError`, `PdfWorkLimitError` (`timeout`, `aborted`, `operations`), `ExportResourceError` (`disposed`, `aborted`) |

One page per slide, at the slide size (EMU ÷ 12700 points). Hidden slides (`show="0"`) are left out, unless `includeHiddenSlides` is true. The output is deterministic: the same deck gives the same bytes. The PDF carries the deck's title, author, subject, and keywords from `docProps/core.xml`, and fixed dates, as in DOCX Editor.

### Differences from DOCX Editor's `docx-to-pdf`

- `substitutes` turns the packaged substitute fonts off. DOCX Editor configures fonts through its ordered font sources and resolvers instead.
- `includeHiddenSlides` has no DOCX Editor equivalent. PPTX has hidden slides; DOCX has no hidden pages. It is an output option of `exportPdfFrom` only. DOCX Editor also accepts `displayMode` at open, as the session's default; this package has no session default, because a session lays out each slide once whatever an export shows.
- `slideNumbers` has no DOCX Editor equivalent: a page is always one slide, but hidden slides have no page.
- `PdfExportSession` exposes only `fontResolution` and `dispose()`. DOCX Editor's session also exposes its Core layout and glyph capabilities, through a Core export contract that its Markdown exporter shares. PPTX Core has no public export-layout contract yet and no second exporter, and its `DocumentSession` is not public API.
- `ExportResourceError` is in this package, not in Core, for the same reason: no second exporter shares it yet.
- `ExportResourceError` has the codes `disposed` and `aborted` only, and takes `(code, options)`: the message follows from the code. DOCX Editor's Core error has more codes and takes a message.
- A cancelled or timed-out export is `PdfWorkLimitError` (`aborted`, `timeout`), for `exportPdf` and `exportPdfFrom` alike, so that both calls report a cancelled export the same way. DOCX Editor reports every abort and timeout as `ExportResourceError` (`aborted`, `timedOut`). Here `ExportResourceError` means only that the session itself is disposed or its own signal aborted.
- DOCX Editor's `displayMode`, `comments`, and `layoutRevision` do not apply to slides.
- The paint modules mirror the editor's SVG paint, one for one (`paint-fill.ts`, `paint-table.ts`, `paint-text.ts`, `paint-effects.ts` beside core's `output/paint-*.ts`). Three pure helpers are shared with core (`adjustForPathMode`, `blurSigma`, `effectMatrix`). The gradient geometry (axis, radius, the trapezoids of a rectangular path gradient) and the picture placement (`fillRect`, `srcRect`) are mirrored, so they have one copy in each painter; `painter-agreement.test.ts` paints slides with both painters and fails when they disagree. DOCX Editor resolves a paint-ready gradient record in Core that both of its painters read; PPTX Core has no such record yet.
- A blurred shadow or glow is up to 13 rings, each a transparency group, so overlapping rings never darken, with σ = `blurRad` × 0.5 from the editor. DOCX Editor draws 6 rings at fixed σ multiples (σ = `blurRad` ÷ 4) over a core at half alpha, without groups.
- Past the paint budget, the export throws `PdfWorkLimitError` with the code `operations`, for operations and content bytes alike, next to `timeout` and `aborted`. DOCX Editor's `PdfWorkLimitError` has the one code `workLimitExceeded` for all of its budgets.
- Each gradient's stop functions are written once per document, and at most 64 stops are drawn. DOCX Editor writes the stops of each gradient as given; a PPTX group, layout, or theme gradient can fill thousands of shapes.
- A reflection is a transparency group of the shape under the reflection transform, through a luminosity soft mask of its fade; its blur is left out. Soft edges and picture brightness and contrast are not drawn yet. DOCX Editor draws them with luminosity masks and the image `Decode` array (`picture-effects.ts`, `images.ts`), which is the pattern to follow.
- Opening an export session always installs the drawing module of `@pptx-editor.dev/pro/drawing`, so a PDF draws glow, inner shadows, reflections, and hard soft edges, as DOCX Editor's PDF export always includes its drawing module.
- Diagnostic codes are dotted by area (`fill.*`, `picture.*`, `line.*`, `effect.*`, `text.*`, `font.*`, `frame.*`). DOCX Editor's codes are kebab-case, as its Markdown warnings are. The table above names the DOCX Editor code where the idea is the same. Handle an unknown code by its `severity`.
- A refused open gives core's open error reason, and the package re-exports core's `OpenError` type, so a host reads one vocabulary with core's `openPresentation`. DOCX Editor re-exports its reason union `HeadlessDocumentRejection`; PPTX Core has no separate union to re-export.

## Limits

| Limit | Value |
| --- | --- |
| `timeoutMs` | 60000 by default. An already aborted `signal` stops before the deck opens. |
| `maxOutputBytes` | 64 MiB at most |
| `maxPages` | 10000 at most |
| Paint operations | 2000000, plus 6000 per page, at most 40000000. DOCX Editor has the same values but takes the larger of the two terms instead of their sum. Counting them also checks `timeoutMs`, and the paint lets the `signal` run every 64 drawn items and between slides. |
| Paint content | Operators and gradient functions before compression, and pictures at their file size: 64 MiB, or 192 KiB per page when that is more, at most 512 MiB (DOCX Editor's budget). Past it, or past the paint operations, the export throws `PdfWorkLimitError` (`operations`). A PNG stays decoded in memory until the save; only the pixel limit below bounds each one. DOCX Editor passes a PNG through without decoding it. |
| Pictures | The header is read before decoding. A picture over 25000000 pixels, or over 20000 pixels on a side, is refused (`picture.too-large`). |

## What it draws

| Content | State |
| --- | --- |
| Slide size, backgrounds (solid, PNG, JPEG) | Drawn |
| Preset and custom geometry, groups, rotation, flips | Drawn, from the engine's evaluated paths |
| Solid fills and lines, with transparency; caps and joins | Drawn |
| PNG and JPEG pictures, identified by their signature: cropped by `srcRect`, placed by `fillRect`, clipped to the geometry, with transparency | Drawn |
| EMF and WMF pictures | Drawn as a PNG at twice their pixel size, at most 4096 pixels on a side, with the deck's fonts that the PDF may embed. The converter draws with the optional peer `@napi-rs/canvas` (`^1.0.9`) when it is installed. Without it, the converter's own rasterizer can refuse a picture with text or an embedded bitmap: that picture reports `picture.metafile` and is not drawn. A picture that runs out of time (10 seconds) converts once more, then reports `picture.metafile` |
| Linear and path gradients on shapes, backgrounds, and table cells; picture fills; group fills; lighten and darken path fills | Drawn, as PDF shadings and clipped pictures |
| Tables: background, cell fills, borders, and cell text | Drawn, from the engine's table layout |
| Text: every run at the position layout measured; bold, italic, size, color, caps, spacing, superscript and subscript, highlight | Drawn, searchable; subset fonts embedded |
| Engine approximations and font substitutions | Reported as `information` |
| Outer and inner shadows, glow | Drawn; the blur is layered rings (`approximation`) |
| Soft edges | Drawn hard (`approximation`) |
| Reflections | Drawn with their fade; the blur is left out (`approximation`) |
| Gradient fills on text | Drawn: the glyphs clip the gradient, which spans the text body |
| Gradient fills on lines | Drawn in the gradient's middle color (`approximation`) |
| Tiled pictures; picture color effects (grayscale, duotone, brightness, black and white) | Drawn stretched; effects not applied (`approximation`) |
| Dashes and line ends | Drawn with the editor's geometry |
| Compound lines | Drawn as one line (`approximation`) |
| Pattern fills | Not drawn (`unsupported`) |
| Underline styles other than single and double | Drawn as a single line (`approximation`) |
| Text outlines | Not drawn (`unsupported`) |
| Upright East Asian vertical text, synthetic bold and italic | Drawn as normal text (`approximation`) |
| Diagrams with a saved drawing | Drawn from the saved drawing, read-only, as in the editor |
| Charts: bar, column, line, area, scatter, pie, doughnut, and radar, from the values cached in the chart part | Drawn, from the engine's layout of the shapes and labels that `@pptx-editor.dev/pro/drawing` gives |
| 3-D, of-pie, stock, bubble, and surface charts; charts with trendlines or error bars | Not drawn (`unsupported`) |
| Diagrams with no saved drawing, OLE objects with no PNG, JPEG, GIF, BMP, EMF, or WMF picture | Not drawn (`unsupported`) |
| Video and audio | The poster picture is drawn; the media is not embedded or played |
| GIF, BMP, and TIFF pictures | Not drawn (`unsupported`) |
| SVG pictures | Drawn from their PNG fallback (`information`); without a fallback, not drawn |
| Hyperlinks, speaker notes, PDF outline | Not written |

## Fonts

- Text uses the same font bytes that layout measured with: fonts passed in `fonts`, fonts embedded in the deck (decoded in memory for this conversion only), and the metric-compatible substitutes from `@docx-editor.dev/fonts`.
- Under Node, the substitutes are read from the `assets/` folder of `@docx-editor.dev/fonts`, and from no other folder. A single-file bundle must ship a copy of that folder and set `DOCX_EDITOR_FONT_ASSET_ROOT` to it before the package loads. Otherwise each face reports `font.substitute-failed`.
- Fonts are embedded with DOCX Editor's own embedder (`fonts.ts`, `pdf-font-embedding.ts`): a subset per face, written as a CID font with a `ToUnicode` map, so text stays searchable. A slide that uses four faces gives a PDF of about 25 KB.
- A face is not embedded when its OS/2 `fsType` forbids embedding or subsetting, permits bitmaps only, or is missing. Its text is not drawn, and the conversion reports `font.not-embedded`. Variable fonts are refused the same way, as in DOCX Editor. DOCX Editor embeds a face that has no OS/2 table; this package refuses it.
- Each character draws the glyph the face's `cmap` gives, as the editor's SVG text does, so there are no ligatures. A character the face lacks is not drawn (`text.missing-glyph`).
- DOCX Editor maps the private-use codepoints of legacy symbol fonts to Unicode for extraction. PPTX Core has no such table yet, so symbol characters extract as their private-use codepoints.

## Dependencies

The dependencies are the same as DOCX Editor's `docx-to-pdf`.

| Package | License | Why |
| --- | --- | --- |
| `pdf-lib` 1.17.1 | MIT | PDF writer |
| `fontkit` 2.0.4 | MIT | Font parsing and subsetting for the embedder copied from DOCX Editor |
| `@pptx-editor.dev/pro` | EigenPal Pro License | Charts, glow, inner shadow, reflection, and soft edge (`@pptx-editor.dev/pro/drawing`), which the export always installs, as DOCX Editor's `docx-to-pdf` installs `@docx-editor.dev/pro/drawing` |
| `pdfjs-dist` 5.4.624 (dev) | Apache-2.0 | Reads the PDF back in tests |
