// Fidelity thresholds and known gaps: the ONE place they live. Tighten over time; never loosen
// to pass. After a change to `THRESHOLDS` or `COMPARISON_B`, regenerate the JSON copy for
// `pptx-eval`: `bun packages/pptx-to-pdf/scripts/evaluation/thresholds-json.ts --write`.
//
// Comparison A is our PDF against the reference PDF: the editor prints each slide with its own
// paint, and pdf.js draws and reads both PDFs the same way (`pdf.ts`).

export const THRESHOLDS = {
  /** Raster width of both pages (height follows the slide's aspect ratio). */
  renderWidthPx: 1920,
  /** Anti-aliasing-tolerant pixel mismatch per slide, as a fraction of all pixels. */
  maxPixelMismatch: 0.005,
  /** pixelmatch per-pixel color threshold (0–1). */
  pixelThreshold: 0.1,
  /** Each matched line's baseline must be within this many points of the reference line. */
  maxBaselineDeltaPt: 0.5,
  /** Line breaks must be identical: the fraction of reference lines that our page must hold. */
  minLineBreakMatch: 1,
} as const;

/**
 * Comparison B: the `pptx-to-pdf` PDF against the reference PDF. The pass limits match comparison A where the metric is the same. The screening
 * values only flag a page for review; they are DOCX Editor's
 * (`docx-to-pdf/scripts/evaluation/worker.py`). `worker.py` and `pptx-eval` read these values
 * from `packages/pptx-to-pdf/scripts/evaluation/thresholds.json`, generated from this file.
 */
export const COMPARISON_B = {
  /** Raster width for the pixel metric, as comparison A. */
  renderWidthPx: THRESHOLDS.renderWidthPx,
  /** Per-pixel YIQ color threshold (0–1), as comparison A. */
  pixelThreshold: THRESHOLDS.pixelThreshold,
  /** Pixel mismatch per slide, as a fraction of all pixels, as comparison A. */
  maxPixelMismatch: THRESHOLDS.maxPixelMismatch,
  /** Every reference word must be matched (percent). A report only in CI until CI has PyMuPDF. */
  minWordMatchPercent: 100,
  /**
   * Each matched word may move this far from its reference position (points). A report only in
   * CI until CI has PyMuPDF.
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
 * A reference PDF page may differ from the slide size (`p:sldSz`) by this many points. The reference
 * app makes each page side a whole number of points. Measured: a fraction of 0.375 pt or less
 * rounds down, and 0.45 pt and every fraction from 0.5 pt round up. Fractions from 0.375 to
 * 0.45 pt are not measured; if they round up, a page side is off by less than
 * 1 − 0.375 = 0.625 pt, the bound here.
 */
export const PAGE_SIZE_TOLERANCE_PT = 0.625;

/** A threshold that a slide misses (`known-gaps.ts`). */
export type FailureKind = 'pixels' | 'line-breaks' | 'baseline';

/**
 * The noise margins of a known gap's recorded value (`known-gaps.ts`). The same tools give the
 * same scores on every machine, so the margins only absorb the rounding of the recorded value:
 * 0.01 % of the pixels, 0.01 pt of baseline, and 0.01 % for line breaks.
 */
export const GAP_MARGINS = { pixels: 0.0001, baseline: 0.01, 'line-breaks': 0.0001 } as const;

/**
 * One known gap: each threshold that a slide misses today, with its measured value. The gate
 * fails when the slide scores worse than the value by more than its margin (`GAP_MARGINS`), and
 * also when it scores better, until the better value is recorded.
 */
export interface KnownGap {
  /** The pixel mismatch, as a fraction of all pixels. */
  readonly pixels?: number;
  /** The share of reference lines that our page holds (0 to 1). */
  readonly 'line-breaks'?: number;
  /** The largest baseline distance of a matched line, in points. */
  readonly baseline?: number;
  /**
   * How many reference lines the slide checks. Another count fails the gate until it is
   * recorded.
   */
  readonly lines?: number;
}

/** The known-gap list, with the tools it was measured with. */
export interface KnownGapList {
  /** The pdf.js version that draws and reads both PDFs. */
  readonly pdfjs: string;
  /** Playwright's Chromium, which prints our PDF. */
  readonly chromium: string;
  /** The fixture fonts (`fixtureFontsIdentity`). */
  readonly fonts: string;
  /** By `deck/slide`, for example `text-runs/1`. */
  readonly slides: Readonly<Record<string, KnownGap>>;
}

/**
 * Known gaps, a ratchet: the fixture slides that fail today, with each threshold they miss and
 * its measured value. `bun run fidelity` fails when a slide fails in a way that is not listed,
 * scores worse or better than its value, or passes now. `bun run fidelity --update-gaps`
 * rewrites this list from a run; review the diff, because a worse value must be a fix instead.
 */
export const FIXTURE_KNOWN_GAPS: KnownGapList = {
  pdfjs: '5.4.624',
  chromium: '156.0.8078.4',
  fonts:
    '@docx-editor.dev/fonts 2.34.0, @docx-editor.dev/fonts-cjk 2.34.0, fixtures/fonts 602beb9f174c',
  slides: {
    'autofit/1': { pixels: 0.0094, lines: 9 },
    'east-asian-line-breaking/1': { pixels: 0.0062, lines: 52 },
    'east-asian-line-breaking/2': { pixels: 0.0068, lines: 52 },
    'east-asian-line-breaking/4': { pixels: 0.0094, 'line-breaks': 0.8, lines: 10 },
    'east-asian-theme-fonts/1': { pixels: 0.0219, 'line-breaks': 0.75, lines: 16 },
    'east-asian-typography/2': { pixels: 0.0148, 'line-breaks': 0.7647, lines: 17 },
    'fill-fidelity/1': { pixels: 0.2808 },
    'fill-fidelity/2': { pixels: 0.2103 },
    'fill-fidelity/3': { pixels: 0.1836 },
    'fill-fidelity/4': { pixels: 0.0148 },
    'fill-fidelity/5': { pixels: 0.0992 },
    'fill-fidelity/6': { pixels: 0.8086 },
    'fills-lines/1': { pixels: 0.0784 },
    'fills-lines/2': { pixels: 0.0116 },
    'follow-ups/1': { pixels: 0.0284 },
    'follow-ups/2': { pixels: 0.0736 },
    'follow-ups/3': { pixels: 0.1153 },
    'font-effects/1': { pixels: 0.008, lines: 10 },
    'font-substitution/1': { pixels: 0.0141, lines: 8 },
    'gradients/1': { pixels: 0.0894 },
    'line-model/2': { pixels: 0.0051, lines: 31 },
    'paragraph-format/1': { pixels: 0.0095, 'line-breaks': 0.8889, lines: 27 },
    'picture-editing/1': { pixels: 0.0076 },
    'picture-editing/2': { pixels: 0.0072 },
    'pictures/1': { pixels: 0.0096 },
    'placeholders-background/1': { pixels: 0.139, lines: 3 },
    'rtl-text/1': { pixels: 0.0292, lines: 1 },
    'rtl-text/2': { pixels: 0.0221, lines: 2 },
    'rtl-text/3': { pixels: 0.0175 },
    'rtl-text/4': { pixels: 0.0226 },
    'shape-effects-presets/2': { 'line-breaks': 0.6667, lines: 9 },
    'shape-effects-presets/3': { 'line-breaks': 0.6667, lines: 6 },
    'slide-structure/1': { pixels: 0.0492, lines: 1 },
    'slide-structure/2': { pixels: 0.0057, lines: 1 },
    'spacing/1': { pixels: 0.0053, lines: 18 },
    'tables/3': { pixels: 0.0324, 'line-breaks': 0.8824, baseline: 2.97, lines: 17 },
    'text-direction/1': { pixels: 0.0198, 'line-breaks': 0.7872, lines: 47 },
    'text-runs/1': { pixels: 0.006, lines: 11 },
    'vertical-columns/1': { pixels: 0.0106, lines: 12 },
    'vertical-east-asian/1': { pixels: 0.0098 },
  },
};
