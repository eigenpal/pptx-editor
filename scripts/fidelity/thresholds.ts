// Fidelity thresholds and known gaps: the ONE place they live. Tighten over time; never loosen
// to pass. After a change to `THRESHOLDS` or `COMPARISON_B`, regenerate the JSON copy for
// `pptx-eval`: `bun packages/pptx-to-pdf/scripts/evaluation/thresholds-json.ts --write`.
//
// Comparison A is our PDF against the reference PDF: the editor prints each slide with its own
// paint, and pdf.js draws and reads both PDFs the same way (`pdf.ts`). The export lane of
// `bun run fidelity` scores the PDF of `exportPdf` (comparison B) against the same reference PDFs
// with the same code and these same thresholds (`FIXTURE_EXPORT_KNOWN_GAPS`).

export const THRESHOLDS = {
  /** Raster width of both pages (height follows the slide's aspect ratio). */
  renderWidthPx: 1920,
  /** Anti-aliasing-tolerant pixel mismatch per slide, as a fraction of all pixels. */
  maxPixelMismatch: 0.005,
  /** pixelmatch per-pixel color threshold (0–1). */
  pixelThreshold: 0.1,
  /** Each matched line's baseline must be within this many points of the reference line. */
  maxBaselineDeltaPt: 0.5,
  /**
   * Each matched line must start or end within this many points of the reference line (the
   * nearer of its left and right edges, so left, right, and centered text all count).
   */
  maxLineShiftPt: 0.5,
  /** Line breaks must be identical: the fraction of reference lines that our page must hold. */
  minLineBreakMatch: 1,
  /**
   * Lines of our page (with two letters or digits) that the reference page does not hold: text
   * that the editor paints and the reference does not.
   */
  maxExtraLines: 0,
  /**
   * Save, reopen, and paint again: the share of slide markup that may differ from the first
   * paint (`markupOf` in the harness page). 0: the markup must be identical.
   */
  maxSelfConsistencyMismatch: 0,
} as const;

/**
 * Comparison B: the `pptx-to-pdf` PDF against the reference PDF. The pass limits match comparison A
 * where the metric is the same. The screening values only flag a page for review; they are DOCX
 * Editor's (`packages/docx-to-pdf/scripts/evaluation/worker.py`). `worker.py` and `pptx-eval` read
 * these values from `packages/pptx-to-pdf/scripts/evaluation/thresholds.json`, generated from this
 * file. DOCX Editor commit adf0e178. In this repository, the export lane of `bun run fidelity`
 * gates the fixtures on comparison B with the checks of comparison A (`THRESHOLDS`): pixels, every
 * reference line, no extra lines, baseline, and line position. The word metrics below are for
 * `worker.py` and `pptx-eval`, which measure words with PyMuPDF.
 */
export const COMPARISON_B = {
  /** Raster width for the pixel metric, as comparison A. */
  renderWidthPx: THRESHOLDS.renderWidthPx,
  /** Per-pixel YIQ color threshold (0–1), as comparison A. */
  pixelThreshold: THRESHOLDS.pixelThreshold,
  /** Pixel mismatch per slide, as a fraction of all pixels, as comparison A. */
  maxPixelMismatch: THRESHOLDS.maxPixelMismatch,
  /**
   * Every reference word must be matched (percent), in `worker.py`. The export lane gates the
   * same text through lines instead (`THRESHOLDS.minLineBreakMatch`, `maxExtraLines`).
   */
  minWordMatchPercent: 100,
  /**
   * Each matched word may move this far from its reference position (points), in `worker.py`.
   * The export lane gates lines instead (`THRESHOLDS.maxBaselineDeltaPt`, `maxLineShiftPt`).
   */
  maxWordDisplacementPt: 0.5,
  /** Page sides may differ by this many points before the page size is a finding. */
  maxPageSizeDeltaPt: 0.5,
  /** Screening of the 192 × 144 color sketch and of matched words; findings, not a verdict. */
  screening: {
    /** A sketch sample changes when one channel differs by more than this (0–255). */
    colorChannelDelta: 24,
    /** The same for the legacy grayscale sketch. */
    grayDelta: 12,
    /** Side of a sketch tile, in samples. */
    tileSamples: 8,
    /** A tile with this many changed samples is a local change. */
    minChangedSamplesInTile: 3,
    /** More than this fraction of changed samples is a change of the page. */
    maxChangedFraction: 0.01,
    /** A matched word that moved more than this many points is a text-position finding. */
    textShiftPt: 2,
  },
} as const;

/**
 * A PDF page, ours or the reference, may differ from the slide size (`p:sldSz`) by this many
 * points. Chromium prints each page side as a whole number of 1/300 inch, not always the nearest
 * one (measured: up to 0.46 pt off). The reference app makes each page side a whole number of
 * points. Measured: a fraction of 0.375 pt or less
 * rounds down, and 0.45 pt and every fraction from 0.5 pt round up. Fractions from 0.375 to
 * 0.45 pt are not measured; if they round up, a page side is off by less than
 * 1 − 0.375 = 0.625 pt, the bound here.
 */
export const PAGE_SIZE_TOLERANCE_PT = 0.625;

/**
 * The harness files that only read and compare the two PDFs. A pull request whose source changes
 * are only these files, harness tests, and documentation paints and prints the same pages, so it
 * may record the failures that its checks find (`check-known-gaps.ts`). It may not check fewer
 * lines. The ratchet reads this list from the base branch, so a change cannot widen it for itself.
 */
export const COMPARISON_FILES: readonly string[] = [
  'scripts/fidelity/compare.ts',
  'scripts/fidelity/pdf.ts',
];

/**
 * Limits that the harness no longer has, each with the reason. `check-known-gaps.ts` refuses a
 * limit that goes away unless it is listed here, so no limit goes away unseen.
 */
export const RETIRED_LIMITS: Readonly<Record<string, string>> = {
  'margin pixels':
    'known gaps record failing checks, not values: rasters differ slightly between machines',
  'margin baseline':
    'known gaps record failing checks, not values: rasters differ slightly between machines',
  'margin line-breaks':
    'known gaps record failing checks, not values: rasters differ slightly between machines',
  'margin extra-text':
    'known gaps record failing checks, not values: rasters differ slightly between machines',
  rasterScale:
    'the supersampling factor: pdf.js draws both PDFs at renderWidthPx, and no raster is scaled down',
  maxSizeDeltaPx:
    'the raster size difference: pdf.js draws both PDFs onto one raster of the slide size, so it cannot differ; each PDF page must be the slide size within PAGE_SIZE_TOLERANCE_PT instead',
};

/** A threshold that a slide misses (`known-gaps.ts`). */
export type FailureKind = 'pixels' | 'line-breaks' | 'baseline' | 'line-position' | 'extra-text';

/** One known gap: the checks that a slide fails today, in the order of `FailureKind`. */
export type KnownGap = readonly FailureKind[];

/** The known-gap list, with the tools it was measured with. */
export interface KnownGapList {
  /** The pdf.js version that draws and reads both PDFs. */
  readonly pdfjs: string;
  /** Playwright's pinned Chromium revision, which prints our PDF (`pinnedChromium`). */
  readonly chromium: string;
  /** The fixture fonts (`fixtureFontsIdentity`). */
  readonly fonts: string;
  /**
   * How many reference lines each fixture checks, over all its slides. Fewer lines fail the gate,
   * so a comparator or layout change that drops lines shows, also on slides that pass.
   */
  readonly deckLines: Readonly<Record<string, number>>;
  /** By `deck/slide`, for example `text-runs/1`. */
  readonly slides: Readonly<Record<string, KnownGap>>;
}

/**
 * Known gaps, a ratchet: the fixture slides that fail today, with each check they fail. Values
 * are not recorded: rasters differ slightly between machines. `bun run fidelity` fails when a
 * slide fails a check that is not listed, or passes a listed check now. `bun run fidelity
 * --update-gaps` rewrites this list from a run; review the diff, because a new kind is a fix.
 */
export const FIXTURE_KNOWN_GAPS: KnownGapList = {
  pdfjs: '6.5.438',
  chromium: 'chromium-headless-shell 1248',
  fonts:
    '@docx-editor.dev/fonts 2.34.0, @docx-editor.dev/fonts-cjk 2.34.0, fixtures/fonts f68672445aef',
  deckLines: {
    'alignment-anchor': 19,
    'arrange-flip-hidden': 3,
    autofit: 9,
    'bullets-gallery': 17,
    'bullets-numbering': 21,
    'chart-data': 20,
    'chart-sizes': 161,
    charts: 118,
    colors: 0,
    'custom-geometry': 0,
    'diagram-editing': 9,
    diagrams: 20,
    'east-asian-line-breaking': 136,
    'east-asian-theme-fonts': 16,
    'east-asian-typography': 91,
    effects: 4,
    'equation-lines': 20,
    equations: 28,
    'fill-fidelity': 0,
    'fills-lines': 0,
    'follow-ups': 0,
    'font-effects': 11,
    gradients: 0,
    'groups-transforms': 4,
    'header-footer': 3,
    hyperlinks: 13,
    'line-ends-dashes': 0,
    'line-model': 65,
    'lines-arrows': 0,
    'open-font-families': 8,
    'paragraph-format': 27,
    'picture-bullets': 24,
    'picture-editing': 0,
    pictures: 0,
    'placeholders-background': 3,
    'preserve-unsupported': 5,
    'preserved-content': 3,
    'preset-shapes': 0,
    'rtl-text': 4,
    'shape-effects-presets': 24,
    'shape-styles': 24,
    'slide-layouts': 10,
    'slide-size': 0,
    'slide-size-rounding': 0,
    'slide-structure': 7,
    spacing: 18,
    'svg-tab-stops': 13,
    'table-editing': 14,
    tables: 46,
    'text-direction': 47,
    'text-runs': 11,
    'text-warp': 0,
    'three-d-effects': 0,
    'three-d-turned-effects': 0,
    'vertical-columns': 12,
    'vertical-east-asian': 0,
  },
  slides: {
    'chart-sizes/7': ['line-position'],
    'charts/6': ['line-position'],
    'equation-lines/1': ['baseline'],
    'equation-lines/2': ['baseline', 'line-position', 'extra-text'],
    'equations/1': ['baseline', 'line-position'],
    'equations/2': ['baseline', 'line-position'],
    'equations/4': ['baseline', 'line-position'],
    'fill-fidelity/3': ['pixels'],
    'fill-fidelity/4': ['pixels'],
    'fills-lines/1': ['extra-text'],
    'picture-editing/1': ['pixels'],
    'picture-editing/2': ['pixels'],
    'pictures/1': ['pixels'],
    'preserved-content/1': ['pixels'],
    'three-d-effects/1': ['pixels'],
    'three-d-effects/2': ['pixels'],
    'three-d-effects/3': ['pixels'],
    'three-d-effects/4': ['pixels'],
    'three-d-turned-effects/1': ['pixels'],
  },
};

/**
 * Known gaps of the export lane, a ratchet with the rules of `FIXTURE_KNOWN_GAPS`: the fixture
 * slides whose `exportPdf` PDF fails today, with each check they fail. `bun run fidelity
 * --update-gaps` rewrites this list from a run of the export lane.
 */
export const FIXTURE_EXPORT_KNOWN_GAPS: KnownGapList = {
  pdfjs: '6.5.438',
  chromium: 'chromium-headless-shell 1248',
  fonts:
    '@docx-editor.dev/fonts 2.34.0, @docx-editor.dev/fonts-cjk 2.34.0, fixtures/fonts f68672445aef',
  deckLines: {
    'alignment-anchor': 19,
    'arrange-flip-hidden': 3,
    autofit: 9,
    'bullets-gallery': 17,
    'bullets-numbering': 21,
    'chart-data': 20,
    'chart-sizes': 161,
    charts: 118,
    colors: 0,
    'custom-geometry': 0,
    'diagram-editing': 9,
    diagrams: 20,
    'east-asian-line-breaking': 136,
    'east-asian-theme-fonts': 16,
    'east-asian-typography': 91,
    effects: 4,
    'equation-lines': 20,
    equations: 28,
    'fill-fidelity': 0,
    'fills-lines': 0,
    'follow-ups': 0,
    'font-effects': 11,
    gradients: 0,
    'groups-transforms': 4,
    'header-footer': 3,
    hyperlinks: 13,
    'line-ends-dashes': 0,
    'line-model': 65,
    'lines-arrows': 0,
    'open-font-families': 8,
    'paragraph-format': 27,
    'picture-bullets': 24,
    'picture-editing': 0,
    pictures: 0,
    'placeholders-background': 3,
    'preserve-unsupported': 5,
    'preserved-content': 3,
    'preset-shapes': 0,
    'rtl-text': 4,
    'shape-effects-presets': 24,
    'shape-styles': 24,
    'slide-layouts': 10,
    'slide-size': 0,
    'slide-size-rounding': 0,
    'slide-structure': 7,
    spacing: 18,
    'svg-tab-stops': 13,
    'table-editing': 14,
    tables: 46,
    'text-direction': 47,
    'text-runs': 11,
    'text-warp': 0,
    'three-d-effects': 0,
    'three-d-turned-effects': 0,
    'vertical-columns': 12,
    'vertical-east-asian': 0,
  },
  slides: {
    'chart-sizes/7': ['line-position'],
    'charts/6': ['line-position'],
    'equation-lines/1': ['baseline'],
    'equation-lines/2': ['line-breaks', 'baseline', 'line-position', 'extra-text'],
    'equations/1': ['baseline', 'line-position'],
    'equations/2': ['baseline', 'line-position'],
    'equations/4': ['baseline', 'line-position'],
    'fills-lines/1': ['extra-text'],
    'preserved-content/1': ['pixels'],
    'three-d-effects/1': ['pixels'],
    'three-d-effects/2': ['pixels'],
    'three-d-effects/3': ['pixels'],
    'three-d-effects/4': ['pixels'],
    'three-d-effects/6': ['pixels'],
    'three-d-turned-effects/1': ['pixels'],
  },
};
