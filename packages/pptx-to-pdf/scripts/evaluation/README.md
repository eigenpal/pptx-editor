# Evaluation protocol

These commands compare PPTX Editor's PDF with the reference PDF of each deck. They add no public API.

The protocol is a copy of DOCX Editor's (`packages/docx-to-pdf/scripts/evaluation/`, commit `398963c7f67207cebf3c38dcc41d5f0bd1339b5f`), changed for slides. Each copied file names its DOCX Editor source. This repository keeps the candidate export and the measurement primitives. Comparison policy, scoring, caching, and reports belong to the evaluation caller.

Protocol version 1 separates candidate exports, measurements, comparisons, and detailed evidence.

```sh
bun packages/pptx-to-pdf/scripts/evaluation/export.ts input.pptx output.pdf

bun packages/pptx-to-pdf/scripts/evaluation/layout-summary.ts input.pptx summary.json

python packages/pptx-to-pdf/scripts/evaluation/worker.py \
  measure input.pdf measurement.json.gz

python packages/pptx-to-pdf/scripts/evaluation/worker.py \
  compare reference.json.gz result.json --candidate candidate.json.gz

python packages/pptx-to-pdf/scripts/evaluation/worker.py \
  evidence reference.pdf evidence --candidate candidate.pdf --pages 1,2

python packages/pptx-to-pdf/scripts/evaluation/worker.py \
  page input.pdf preview --pages 1
```

Run every command from the root of the checkout, after `bun install`. The Bun commands read the engine from source, so you do not need to build the packages. Use the evaluator's locked Python environment, which supplies PyMuPDF and Pillow. Run each command in a bounded child process: PDF measurement can allocate native memory.

## The contract

[`contract.json`](contract.json) is the machine-readable form of this page. Evaluation callers read it to find the commands and their arguments.

| Field | What it holds |
| --- | --- |
| `protocol` | The protocol version, `1`. `settings.ts` and `worker.py` use the same number. A change that breaks a caller increments it. |
| `runtime` | Bun, Python, and Python package versions. |
| `requiredFiles` | The files that a checkout must have for protocol version 1. |
| `identity.candidate` | The paths to hash for the identity of candidate exports and layout summaries. |
| `identity.measurement` | The paths to hash for the identity of measurements and comparisons. A change here invalidates comparison evidence, not candidate exports. |
| `commands` | Each command: `argv` with `{placeholders}`, optional arguments, inputs, the output and stdout schemas, exit codes, and, for the Bun commands, a timeout. |
| `schemas` | A JSON Schema for each input and output: the font plan, the export report, the error report, the layout summary, the measurement, the comparison, the evidence report, and the page report. |

`contract.test.ts` and `test_contract.py` run every command as the contract writes it and check each output against its schema. When you change an output, change the schema in the same pull request.

## Export

`export.ts input.pptx output.pdf [--fonts plan.json]` writes the candidate PDF and prints one JSON line on stdout:

```json
{
  "protocol": 1,
  "status": "exported",
  "pages": 2,
  "slideNumbers": [1, 3],
  "slideCount": 3,
  "hiddenSlides": [2],
  "diagnostics": [{ "code": "frame.chart", "severity": "unsupported", "message": "…" }],
  "fontResolution": { "faces": [], "substitutions": [], "refusedEmbedded": [] },
  "fontPlan": null,
  "timings": { "openMs": 18, "layoutMs": 4, "paintMs": 9, "saveMs": 2 }
}
```

Exports use best-effort rendering. Diagnostics stay part of the export report, and approximate output is never reported as strict success. A slide that the writer cannot draw completely is still a page, with an `unsupported` or `approximation` diagnostic.

The reference PDF does not contain hidden slides, so the export leaves them out too. `slideNumbers` gives the deck slide number of each page. Use it to pair a candidate page with a reference page: page `i` of the reference PDF is the `i`-th visible slide. A slide that the engine cannot lay out has no page; its number is absent from `slideNumbers`. Give this report to `compare` as `--candidate-report`, and `compare` pairs the pages through it.

On failure, the command prints `{ "protocol": 1, "status": "error", "error": "…", "code": "…" }` and exits with code 1. `code` is always set. It is the typed error code of the package (for example `documentOpenFailed`, `timeout`, or `pageLimitExceeded`), `usage` for the arguments, `fontPlan` for the font plan or one of its face files, `evidenceLimit` for layout text past the record limit, a Node.js file code such as `ENOENT` for the deck or the output folder, or `internal` for an engine failure with no code. The command writes no PDF.

One deadline of 60 s covers the open and the export together. Past it, the command reports `timeout`, in either phase. The engine's open cannot be cancelled, so the deadline is checked when the open ends; an open longer than the contract's timeout of 90 s ends the process with no JSON line.

The export opens the deck through the package's session opener, the same one `openDocumentForExport` uses, and paints it with `exportPdfFrom`, so it lays out each slide once. Its bytes equal those of `exportPdf` with the same settings. `export.test.ts` checks this.

### Fonts

Without `--fonts`, the export uses the deck's embedded fonts and the metric-compatible substitutes packaged in `@docx-editor.dev/fonts` (Carlito for Calibri, Liberation Sans for Arial, and the others in `core/src/editor/document-session.ts`). It reads no installed font, as DOCX Editor's evaluation export sets `useSystemFonts: false`. The result is the same on every machine.

The reference PDFs use the reference fonts. To compare like with like, give the export the same fonts with a font plan:

```json
{
  "protocol": 1,
  "faces": [{ "path": "Calibri.ttf", "family": "Calibri", "bold": false, "italic": false }]
}
```

A relative `path` resolves against the plan's folder. Each file must be one TrueType or OpenType face. The export refuses a font collection (`.ttc`). Plan faces come first, before embedded fonts and substitutes, as the fidelity harness loads the reference fonts. The report's `fontPlan` gives the face count and a SHA-256 for the cache key of the export. The hash runs over one line per face, in plan order: the family, a NUL, `1` or `0` for bold and for italic, a NUL, the hex SHA-256 of the face bytes, and a newline. It changes when a face's bytes or the plan order change (`export.test.ts`). The font files are read at run time. Never commit them.

## Layout summary

`layout-summary.ts input.pptx output.json [--fonts plan.json]` lays out the deck through the same session as the export, without painting a PDF. It writes the page count, the page-to-slide map, the page size in points, the engine's layout notes for each page, and logical text. It prints `{ "protocol": 1, "status": "laid-out", "pages": 2 }`.

Logical text has one entry for each laid-out line, in paint order: its text, the part and engine node of the shape or table that holds it, the layer (slide, layout, or master), the paragraph index, and a bounding box in slide points from the top left. The box is the layout's line box after the shape's transform and the text body rotation. It is not PDF extraction, and paint can differ. Hidden shapes are left out, as paint leaves them out. Text is untrusted document content.

## Measure and compare

`measure` records each page's size, words with boxes (PyMuPDF), a 192 by 144 RGB sketch, image boxes, the drawing count, and font names. `compare` reports page count and page size differences, word coverage, word movement with the word alignment from DOCX Editor's `pdf-visual-diff.py` (`word_movement.py`), and visual screening by local regions. It names the first divergence with a confidence and a coordinate space. Object counts alone never establish missing visible content. Comparisons are screening only: scores belong to the evaluation caller.

`compare` pairs pages by number. With `--candidate-report`, it first renumbers the candidate's pages to reference pages through `slideNumbers` and `hiddenSlides`, so a slide without a candidate page shifts no later page. It refuses a report whose page count differs from the candidate's, and a reference that does not have one page per visible slide, because every later slide would pair with the wrong page. Each `page` in the comparison is a reference page (a slide); each `candidatePage` and `candidateLocation.page` is a page of the candidate PDF. A `pagination` finding names a slide that has no candidate page, or candidate pages with no reference page (`extraCandidatePages`: pages past the reference, or pages of hidden slides). Without the report, it also names a word matched on another page. `firstDivergence` is the earliest finding with high or medium confidence, as in DOCX Editor, so a low-confidence visual finding on an earlier slide does not win. With `--candidate-report`, `evidence` pairs each reference page with its candidate page, and its report gives `candidatePage` (null for a slide without one).

The visual screen compares the two sketches sample by sample. On a 960 by 540 pt slide, one sample is 5 by 3.75 pt. A sample changes when one color channel differs by more than 24 of 255. A page has a `visual-region` finding when more than 1 % of samples change, or when one 8 by 8 sample tile has 3 or more changed samples. A smaller change, for example a thin light border, can stay below this floor: read the evidence, not the screen, before you call a page equal. These are DOCX Editor's values. A matched word that moved more than 2 pt is a `text-position` finding, and page sides that differ by more than 0.5 pt are a `page-geometry` finding.

`evidence` renders at most three pages of two PDFs at up to 144 DPI, with an overlay and a crop of the changed area. `page` renders one page at up to 144 DPI into `page.png`, and its `report.json` gives the page size in points and the PDF rotation matrix. A page too large for the pixel budget renders at a lower resolution, and the report gives the DPI used.

Changes from DOCX Editor's worker:

- The visual sketch is 192 by 144, landscape, with the same number of samples as DOCX Editor's portrait sketch.
- Visual regions are `top`, `body`, and `bottom`. A slide has no running header or footer.
- A page that is too large for the pixel budget renders at a lower resolution instead of failing. ECMA-376 allows slides up to 56 inches on a side (`ST_SlideSizeCoordinate`, at most 51206400 EMU). `evidence` and `page` report the resolution they used.
- Two color sketches of different sizes, for example a DOCX Editor measurement and a PPTX one, are refused instead of compared as grayscale.
- The screening and page-size limits come from `thresholds.json`, generated from `scripts/fidelity/thresholds.ts`.
- `compare` pairs pages by number and takes `--candidate-report`, because a slide can have no candidate page. With the report, a word matches only on its own slide, and a reference that does not have one page per visible slide is refused. The reference PDFs leave hidden slides out.

Location excerpts in comparisons are clipped to 120 characters and marked untrusted. Word text in `movement` samples is untrusted document content too. Failed or missing inputs never count as passes.

## Thresholds

Every limit of the fidelity gates lives in `scripts/fidelity/thresholds.ts`; comparison B's limits are its `COMPARISON_B` block. `thresholds.json` beside the worker is generated from that file (`bun packages/pptx-to-pdf/scripts/evaluation/thresholds-json.ts --write`), and `contract.test.ts` fails when the two disagree. `worker.py` reads its screening limits from it when `compare` first needs them, and the evaluation caller reads the pass limits (`maxPixelMismatch`, `minWordMatchPercent`, `maxWordDisplacementPt`) from it. A missing or incomplete file stops `compare` with one error that names it. Its hash is part of the measurement identity. The worker keeps two sets of DOCX Editor's values in code, because they name report categories and are not gates: the 0.5, 2, and 8 pt movement buckets in `word_movement.py`, and the change level of 28 of 255 that `evidence` crops to.

CI runs no comparison B yet. When the fixture PDFs land, comparison B on the fixtures gates in CI on the raster metric only. The word alignment needs PyMuPDF, which CI does not install, so its limits are a report until CI has it.

## Run the tests

The Bun tests run in `bun run test`:

```sh
bun test packages/pptx-to-pdf/scripts
```

The Python tests need PyMuPDF and Pillow, so CI does not run them. Run them in the evaluator's environment, or with `uv`:

```sh
uv run --no-project --python '>=3.11' --with 'pymupdf>=1.26,<2' --with 'Pillow>=11,<13' \
  python -m unittest discover -s packages/pptx-to-pdf/scripts/evaluation -p 'test_*.py'
```

`test_fixture_protocol.py` exports synthetic fixtures with Bun, so Bun must be on the `PATH`; without it, those tests are skipped and reported as skipped. The tests use only decks under `fixtures/` and PDFs they make. Never use a deck from `demo-presentations/` or the corpus in a test.

## Gaps

- The PDF writer does not draw pattern fills, reflections, charts, and some picture formats yet, and it approximates shadow and glow blur (`packages/pptx-to-pdf/README.md`). Comparison B shows these gaps; it does not hide them.
- There is no Node.js runtime loader yet (DOCX Editor's `runtime-loader.mjs`). Run the Bun commands with Bun, and enforce memory and time limits on the process.
- DOCX Editor's `probe.ts`, `trace.ts`, `layout-worker.ts`, `quick_text.py`, and browser probe have no PPTX counterpart yet. `probe.ts` comes in its own pull request after this protocol: it serves the corpus intake, not comparison B, and it adds a contract command with its own schema.
- The JSON of the reference identity schema waits for the fixture references and their identity file.
- DOCX Editor's export offers packaged glyph fallback faces (`glyphFallbacks`). The PPTX engine has no fallback list yet, so a character that no face covers is not drawn and is reported as `text.missing-glyph`. This affects symbols and Wingdings bullets when the font plan lacks the face.
