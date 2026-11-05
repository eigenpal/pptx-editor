// Fidelity thresholds and tolerances: the ONE place they live. Tighten over time; never loosen
// to pass. After a change, regenerate the JSON copy for `pptx-eval`:
// `bun packages/pptx-to-pdf/scripts/evaluation/thresholds-json.ts --write`.
//
// Geometry is checked by pixels only. A shape-bounds check needs vector paths from the PDF.

export const THRESHOLDS = {
  /** Reference raster width for 16:9 and other aspect ratios (height follows). */
  renderWidthPx: 1920,
  /**
   * Both images are rasterized at this many device pixels per comparison pixel, and each is
   * box-filtered down to `renderWidthPx` before any comparison (`supersample.ts`). Chromium rounds
   * a glyph's y to a whole device pixel and Poppler floors it, so at scale 1 an exact text line
   * is 1 px low whenever its PDF baseline has a pixel fraction of 0.5 or more. At scale 4 the
   * snapping moves a glyph by at most 0.25 px. It is not a tolerance: every limit below applies
   * to the filtered images. It changed once, from 1 to 4, and does not change again
   * (`RASTER_SCALE_MOVE`, `bun run fidelity:check-gaps`).
   */
  rasterScale: 4,
  /**
   * Where pixelmatch judges which differing pixels are anti-aliasing: `device`, on the
   * device-scale images before the box filter (`deviceAntialiased` in `supersample.ts`). On the
   * filtered images, every edge is a ramp, and a stroke moved by 1 px would count as
   * anti-aliasing. It changed once, from `images` to `device` (`ANTIALIASING_MOVE`,
   * `bun run fidelity:check-gaps`).
   */
  antialiasing: 'device',
  /**
   * The render may differ from its size at `renderWidthPx` (`renderSizePx`), and the PDF raster
   * from its expected size (`expectedRasterSizePx`: the render size, cut where the reference PDF rounded a
   * page side down), by this many pixels. A raster that is short where its page ends compares its
   * overlap. A larger difference fails as a slide size mismatch, with no pixel score.
   */
  maxSizeDeltaPx: 0,
  /** Anti-aliasing-tolerant pixel mismatch per slide, as a fraction of all pixels. */
  maxPixelMismatch: 0.005,
  /** pixelmatch per-pixel color threshold (0–1). */
  pixelThreshold: 0.1,
  /** Each line's baseline must be within this many points of the PDF position. */
  maxBaselineDeltaPt: 0.5,
  /** Line breaks must be identical on every text body: fraction of lines that must match. */
  minLineBreakMatch: 1,
  /** Save → reopen → re-render must be pixel-identical to the first render. */
  maxSelfConsistencyMismatch: 0,
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
 * rounds down, and 0.45 pt and every fraction from 0.5 pt round up (`CONTRIBUTING.md`,
 * "Fidelity harness"). Fractions from 0.375 to 0.45 pt are not measured; if they round up, a page side is
 * off by less than 1 − 0.375 = 0.625 pt, the bound here.
 */
export const PAGE_SIZE_TOLERANCE_PT = 0.625;

/** A threshold that a slide misses (`known-gaps.ts`). */
export type FailureKind = 'size' | 'pixels' | 'line-breaks' | 'baseline';

/**
 * The font environments that have a known-gap list. Scores depend on the fonts and the
 * rasterizer: `linux` is Linux without the reference fonts, as in CI (Blacksmith Ubuntu 24.04);
 * `macos-reference` is macOS with the reference fonts (the Mac that exports
 * the references). A new
 * environment comes with its detection (`fidelityEnvironment`) and its list in one change.
 */
export const FIDELITY_ENVIRONMENTS = ['linux', 'macos-reference'] as const;
export type FidelityEnvironment = (typeof FIDELITY_ENVIRONMENTS)[number];

/**
 * The largest baseline offset of the 80 % box estimate seen on the 11 fixture slides whose lines
 * all match (`CONTRIBUTING.md`, "Fidelity harness"). It is an observation, not a proven bound: the report
 * says a larger delta is above it and does not name the estimate as its cause. It is a report
 * rule, not a gate threshold.
 */
export const BASELINE_ESTIMATE_OFFSET_PT = 1.02;

/**
 * The noise margins of a known gap's recorded value (`known-gaps.ts`). Each environment renders
 * the same scores on every run (CI reproduces its list exactly), so the margins only absorb the
 * rounding of the recorded value and a small rasterizer change: 0.01 % of the pixels (about
 * 200 px at 1920 × 1080), 0.01 pt of baseline, and 0.01 % for line breaks, which are counts.
 */
export const GAP_MARGINS = { pixels: 0.0001, baseline: 0.01, 'line-breaks': 0.0001 } as const;

/**
 * One known gap: each threshold that a slide misses today, with its measured value. The gate
 * fails when the slide scores worse than the value by more than its margin (`GAP_MARGINS`), and
 * also when it scores better by more than its margin, until the better value is recorded.
 */
export interface KnownGap {
  /** The render or the PDF raster is not its expected size, so the pixels are not compared. */
  readonly size?: true;
  /** The pixel mismatch, as a fraction of all pixels. */
  readonly pixels?: number;
  /** The share of laid-out lines that appear as PDF lines (0 to 1). */
  readonly 'line-breaks'?: number;
  /** The largest baseline distance of a matched line, in points. */
  readonly baseline?: number;
  /**
   * How many laid-out lines the slide checks, for every slide with checked lines except a size
   * entry, which measures none (`gapEntry`).
   * Another count fails the gate until it is recorded, and fewer lines are never a fix.
   */
  readonly lines?: number;
}

/** One environment's list, with the tools it was measured with. */
export interface KnownGapList {
  /** `pdftoppm -v`: Poppler rasterizes the reference PDF. */
  readonly poppler: string;
  /** Playwright's Chromium renders the editor. */
  readonly chromium: string;
  /**
   * The `THRESHOLDS.rasterScale` that the scores were measured at. A test keeps it equal to the
   * constant. `bun run fidelity:check-gaps` let each list take new values once, when it moved
   * from scale 1 to 4 (`RASTER_SCALE_MOVE`).
   */
  readonly rasterScale: number;
  /** The `THRESHOLDS.antialiasing` that the scores were measured with; a test keeps it equal. */
  readonly antialiasing: 'images' | 'device';
  /** By `deck/slide`, for example `text-runs/1`. */
  readonly slides: Readonly<Record<string, KnownGap>>;
}

/**
 * Known gaps, a ratchet: the fixture slides that fail against their reference PDF, by
 * environment, with each threshold they miss and its measured value. `bun run fidelity` fails
 * when a slide fails in a way that is not listed, when a listed slide scores worse or better
 * than its value, and when a listed failure is gone; it prints the entry to record. A new entry,
 * a new kind, or a looser value is allowed only for a new fixture, or after a changed deck or a
 * changed environment (app version, macOS, font set, export script); `bun run fidelity:check-gaps`
 * checks this against the base branch in CI. The thresholds above stay the same for every
 * slide; a listed slide still reports "fail". `bun run fidelity -- --print-gaps` prints this run's
 * entries.
 */
export const FIXTURE_KNOWN_GAPS: Readonly<Record<FidelityEnvironment, KnownGapList>> = {
  // Recorded 31 October 2026 against the reference PDFs: `linux` in the CI run of
  // eigenpal/pptx-editor-miro#29 at 6ff03b3, `macos-reference` on the Mac that exported the references.
  // Both: 7 of 33 slides pass; the `slide-size` and `slide-size-rounding` fixtures came later
  // and pass too. Why each slide fails: `CONTRIBUTING.md`, "Fidelity harness".
  // Tightened the same day:
  // baselines are glyph origins from the PDF, and the line model is calibrated. `linux` values
  // come from the CI run of eigenpal/pptx-editor-miro#52 at da1ed45. Whole-point baselines are off, so
  // seven slides miss the 0.5 pt baseline limit by 0.05 pt or less. `line-model` is a new fixture
  // (the gate allows entries for one). Why each slide fails: `CONTRIBUTING.md`, "Fidelity harness".
  // 31 October 2026, fill-fidelity (eigenpal/pptx-editor-miro#51; `linux` from its CI run at a542094):
  // `gradients/1` passes, `fills-lines/1`, `tables/3`, and on `linux` `placeholders-background/1`
  // score better, and the new `fill-fidelity` export adds 6 slides, of which 1, 2,
  // and 5 pass. Slides 3, 4, and 6 fail on the Splash raster of tiles and pictures. Combined with eigenpal/pptx-editor-miro#52: 18 of 43 slides pass on `linux`,
  // 21 of 43 on `macos-reference`.
  // 1 November 2026, the new `picture-editing` fixture (eigenpal/pptx-editor-miro#54; `linux` from its
  // CI run at ed1f83c): both slides fail on pixels, with the same
  // values in both environments, from the "Blur", "Glow and soft edge", and "Picture recolor"
  // approximations (the gate allows entries for a new fixture).
  // 1 November 2026, the shape-format fixtures (eigenpal/pptx-editor-miro#53; `linux` from its
  // CI run at 41bd712): `arrange-flip-hidden`, `shape-styles`, and
  // slide 3 of `shape-effects-presets` pass in both environments. Slides 1 and 2 of
  // `shape-effects-presets` fail on pixels, from the "Blur" and "Glow and soft edge"
  // approximations (the gate allows entries for a new fixture).
  // 1 November 2026, `paragraph-format` (eigenpal/pptx-editor-miro#30,
  // `macos-reference` only): line breaks match, but slide 1 fails on baseline, from the "Line model"
  // approximation (the gate allows entries for a new fixture). `linux` from the CI run of eigenpal/pptx-editor-miro#43 at 8e0a51f.
  // 1 November 2026, `text-direction` (eigenpal/pptx-editor-miro#34,
  // `macos-reference` only): slide 1 fails on pixels, line breaks, and baseline, from the "Stacked
  // text (`wordArtVert`)" and "Other vertical text types" approximations, a line break in the
  // three-column box, and repeated column lines that `pdftotext` joins across two columns (the
  // gate allows entries for a new fixture). `linux` from the CI run of eigenpal/pptx-editor-miro#43 at 8e0a51f.
  // 1 November 2026, `font-effects` and `bullets-gallery` (eigenpal/pptx-editor-miro#43,
  // `macos-reference` only): `bullets-gallery` passes. `font-effects` slide 1
  // fails on baseline: the subscript line is 2.83 pt off. Both decks changed here, so these
  // references and the `font-effects` entry replace the ones of eigenpal/pptx-editor-miro#32 and eigenpal/pptx-editor-miro#36 (the gate allows entries
  // for a changed deck). `linux` from the CI run of eigenpal/pptx-editor-miro#43 at 8e0a51f.
  // 5 November 2026, the new `rtl-text`, `east-asian-line-breaking`, and `follow-ups` fixtures:
  // all 11 slides fail. Right-to-left lines and CJK line breaking miss on line breaks and
  // baselines ("Right-to-left text", "East Asian, complex-script, and symbol fonts", "Anchoring
  // and alignment variants"). The moving paths and the gray transform miss on pixels ("Paths
  // without a move", "Gray transform"), and so do pattern fills ("Pattern bitmaps"). The gate
  // allows entries for a new fixture. `linux` from the CI run of eigenpal/pptx-editor-miro#160 at
  // 69c253a.
  // 5 November 2026, measured blur, glow, and soft edge kernels and the inner shadow side:
  // `effects/1` and slides 1 and 2 of `shape-effects-presets` pass, and both `picture-editing`
  // slides score better (`macos-reference` on the Mac with Poppler 26.01.0, every other slide
  // unchanged against main on that Mac; `linux` from the CI runs of eigenpal/pptx-editor-miro#157,
  // the same list at every commit of the change, every other slide unchanged against main).
  // 5 November 2026, supersampled rasters (`THRESHOLDS.rasterScale` 4, `supersample.ts`): both
  // images are rasterized at 4 device pixels for each comparison pixel and box-filtered down, so
  // the glyph snapping of Chromium (round) and Poppler (floor) no longer puts an exact line 1 px
  // off. Every pixel score was measured again once, in both environments, and each list records
  // `rasterScale: 4` (`RASTER_SCALE_MOVE`). The thresholds did not change.
  // 5 November 2026, anti-aliasing judged at the device scale (`THRESHOLDS.antialiasing`,
  // `deviceAntialiased`): judged on the filtered images, a stroke or a line moved by 1 px counted
  // as anti-aliasing, so the scale-4 lists missed such a move. Every pixel score was measured
  // again once, in both environments, and each list records `antialiasing: 'device'`
  // (`bun run fidelity:check-gaps` allows the new values for that move only,
  // `ANTIALIASING_MOVE`). The thresholds did not change. Measured on `macos-reference` (Poppler
  // 26.10.0): 39 of 73 slides pass, as before. Most pixel entries are higher, because a 1 px
  // move of a stroke now counts. `autofit/1`, `fills-lines/1`, `paragraph-format/1`, and
  // `vertical-columns/1` already fail on lines, and now fail on pixels too. `linux` from the CI
  // report of this change: 35 of 73 slides pass. With the substitute fonts, glyph outlines differ
  // from the reference PDFs, and the stricter test counts more of them: `alignment-anchor/1` and
  // `line-model/1` (which failed at scale 1) and `slide-layouts/1` fail on pixels again or now.
  // 5 November 2026, whole-point baselines and each face's own line metrics
  // (`WHOLE_POINT_BASELINES`, `face-line-metrics.ts`): on `macos-reference`, 46 of 73 slides
  // pass, against 39. Seven slides pass that failed on baseline, and 11 listed slides score
  // better. No listed slide scores worse. `linux` from the CI report of this change at
  // 68aca79: 41 of 73 slides pass, against 35. Owner exception (2026-11-05, PR eigenpal/pptx-editor-miro#156,
  // `OWNER_GAP_EXCEPTION` in `check-known-gaps.ts`): `rtl-text/4` on `linux` goes from 3.40 % to
  // 3.41 % of pixels, the value the CI run measured. Each of its 12 lines is within 0.12 pt of its
  // PDF glyph origin, against up to 0.49 pt without the rule. CI draws it with substitute glyph
  // outlines, and no line breaks at the same words as the PDF (`line-breaks` 0).
  linux: {
    poppler: '24.02.0',
    chromium: '156.0.8078.4',
    rasterScale: 4,
    antialiasing: 'device',
    slides: {
      'autofit/1': { pixels: 0.0061, lines: 9 },
      'bullets-numbering/1': { pixels: 0.0091, lines: 9 },
      'east-asian-line-breaking/1': { pixels: 0.0764, 'line-breaks': 0.1154, lines: 52 },
      'east-asian-line-breaking/2': { pixels: 0.0594, 'line-breaks': 0.6154, lines: 52 },
      'east-asian-line-breaking/3': { pixels: 0.0307, 'line-breaks': 0.3929, lines: 28 },
      'east-asian-line-breaking/4': { pixels: 0.0197, 'line-breaks': 0.3077, lines: 13 },
      'fill-fidelity/3': { pixels: 0.0091 },
      'fill-fidelity/4': { pixels: 0.0177 },
      'fill-fidelity/6': { pixels: 0.0513 },
      'fills-lines/1': { pixels: 0.0096, 'line-breaks': 0, lines: 1 },
      'fills-lines/2': { pixels: 0.0062 },
      'follow-ups/1': { pixels: 0.0295 },
      'follow-ups/2': { pixels: 0.0859 },
      'follow-ups/3': { pixels: 0.1456 },
      'font-effects/1': { pixels: 0.0051, lines: 11 },
      'font-substitution/1': { pixels: 0.0278, lines: 8 },
      'paragraph-format/1': { pixels: 0.0062, lines: 27 },
      'picture-editing/1': { pixels: 0.008 },
      'picture-editing/2': { pixels: 0.0083 },
      'pictures/1': { pixels: 0.0063 },
      'placeholders-background/1': { pixels: 0.0126, lines: 3 },
      'rtl-text/1': { pixels: 0.0417, 'line-breaks': 0.0833, lines: 12 },
      'rtl-text/2': { pixels: 0.037, 'line-breaks': 0.0769, lines: 13 },
      'rtl-text/3': { pixels: 0.0258, 'line-breaks': 0, lines: 18 },
      'rtl-text/4': { pixels: 0.0341, 'line-breaks': 0, lines: 12 },
      'slide-layouts/1': { pixels: 0.0058, lines: 2 },
      'slide-structure/1': { pixels: 0.0502, lines: 1 },
      'tables/3': { pixels: 0.024 },
      'text-direction/1': { pixels: 0.0133, 'line-breaks': 0.6154, baseline: 91.08, lines: 52 },
      'vertical-columns/1': { pixels: 0.0079, lines: 12 },
      'vertical-east-asian/1': { pixels: 0.0119 },
    },
  },
  'macos-reference': {
    poppler: '26.10.0',
    chromium: '156.0.8078.4',
    rasterScale: 4,
    antialiasing: 'device',
    slides: {
      'autofit/1': { pixels: 0.0053, lines: 9 },
      'east-asian-line-breaking/1': { pixels: 0.0082, 'line-breaks': 0.4615, lines: 52 },
      'east-asian-line-breaking/2': { 'line-breaks': 0.6154, lines: 52 },
      'east-asian-line-breaking/3': { pixels: 0.0158, 'line-breaks': 0.6786, lines: 28 },
      'east-asian-line-breaking/4': { pixels: 0.0169, 'line-breaks': 0.3077, lines: 13 },
      'fill-fidelity/3': { pixels: 0.0091 },
      'fill-fidelity/4': { pixels: 0.0177 },
      'fill-fidelity/6': { pixels: 0.0513 },
      'fills-lines/1': { pixels: 0.0078, 'line-breaks': 0, lines: 1 },
      'fills-lines/2': { pixels: 0.0062 },
      'follow-ups/1': { pixels: 0.0295 },
      'follow-ups/2': { pixels: 0.0859 },
      'follow-ups/3': { pixels: 0.1456 },
      'font-substitution/1': { pixels: 0.0166, lines: 8 },
      'picture-editing/1': { pixels: 0.008 },
      'picture-editing/2': { pixels: 0.0084 },
      'pictures/1': { pixels: 0.0063 },
      'rtl-text/1': { pixels: 0.0346, 'line-breaks': 0.0833, lines: 12 },
      'rtl-text/2': { pixels: 0.0303, 'line-breaks': 0.1538, lines: 13 },
      'rtl-text/3': { pixels: 0.0204, 'line-breaks': 0, lines: 17 },
      'rtl-text/4': { pixels: 0.0267, 'line-breaks': 0, lines: 12 },
      'slide-structure/1': { pixels: 0.0499, lines: 1 },
      'tables/3': { pixels: 0.0238 },
      'text-direction/1': { pixels: 0.012, 'line-breaks': 0.6154, baseline: 91.08, lines: 52 },
      'vertical-columns/1': { pixels: 0.0065, lines: 12 },
      'vertical-east-asian/1': { pixels: 0.0127 },
    },
  },
};
