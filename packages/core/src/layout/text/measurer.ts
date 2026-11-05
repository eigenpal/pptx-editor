// The text measurer port. Layout is DOM-free; the host injects a measurer. The HarfBuzz
// measurer shapes with the real font bytes; the fixed measurer is a deterministic stand-in for
// tests. Both answer the same two questions: which face draws this request, and how wide is
// this text in that face.

/** A font request from a run. */
export interface FontRequest {
  readonly family: string;
  readonly bold: boolean;
  readonly italic: boolean;
}

/** Face metrics in em units (font size = 1). Ascent and descent are positive. */
export interface FaceMetrics {
  readonly ascent: number;
  readonly descent: number;
  readonly lineGap: number;
  readonly underlinePosition: number;
  readonly underlineThickness: number;
}

/** The face that measures and paints a request. */
export interface ResolvedFace {
  /** The internal family name paint must use, so paint draws exactly the measured bytes. */
  readonly id: string;
  /** The family the face actually is (for diagnostics and the fidelity report). */
  readonly family: string;
  readonly bold: boolean;
  readonly italic: boolean;
  /** The request asked for bold/italic the face does not have: paint synthesizes it. */
  readonly synthBold: boolean;
  readonly synthItalic: boolean;
  /** True when the requested family was not available and a substitute answered. */
  readonly substituted: boolean;
  readonly metrics: FaceMetrics;
}

/** The measurer port. */
export interface TextMeasurer {
  resolve(request: FontRequest): ResolvedFace;
  /**
   * Advances in em per UTF-16 code unit. A cluster's whole advance sits on its first unit and
   * the rest of the cluster's units are 0, so offsets map to positions without a cluster table.
   */
  advances(
    face: ResolvedFace,
    text: string,
    options?: { readonly kerning?: boolean }
  ): Float64Array;
}

/**
 * Metric-compatible substitutes for faces that cannot ship (from `@docx-editor.dev/fonts`), and
 * close open alternatives for common presentation faces. Keys are lower case.
 */
export const FONT_SUBSTITUTES: Readonly<Record<string, readonly string[]>> = Object.freeze({
  calibri: ['Carlito'],
  'calibri light': ['Carlito'],
  cambria: ['Caladea'],
  arial: ['Liberation Sans', 'Arial'],
  helvetica: ['Liberation Sans'],
  'helvetica neue': ['Liberation Sans'],
  'arial narrow': ['Liberation Sans'],
  'times new roman': ['Liberation Serif'],
  times: ['Liberation Serif'],
  'courier new': ['Liberation Mono'],
  courier: ['Liberation Mono'],
  'century gothic': ['TeX Gyre Adventor'],
  aptos: ['Carlito'],
  'aptos display': ['Carlito'],
  'aptos narrow': ['Liberation Sans'],
  'segoe ui': ['Liberation Sans'],
  'open sans': ['Liberation Sans'],
  'myriad pro': ['Liberation Sans'],
  'myriad pro light': ['Liberation Sans'],
});

/**
 * Substitutes built to match the original's advance widths (`@docx-editor.dev/fonts`): line
 * breaks match the original font, only glyph shapes differ. Every other substitution can change
 * line breaks. Keys are lower case.
 */
export const METRIC_COMPATIBLE: Readonly<Record<string, string>> = Object.freeze({
  calibri: 'carlito',
  cambria: 'caladea',
  arial: 'liberation sans',
  'times new roman': 'liberation serif',
  'courier new': 'liberation mono',
  'century gothic': 'tex gyre adventor',
});

/** Whether measuring `family` with `substitute` keeps the original's advance widths. */
export function isMetricCompatible(family: string, substitute: string): boolean {
  return METRIC_COMPATIBLE[family.toLowerCase()] === substitute.toLowerCase();
}

/** The face used when nothing else matches. */
export const DEFAULT_FALLBACK_FAMILY = 'Liberation Sans';

/**
 * A deterministic measurer for tests and for hosts without fonts: every code unit is 0.5 em,
 * spaces 0.25 em. Never use it for fidelity work.
 */
export function createFixedMeasurer(): TextMeasurer {
  const metrics: FaceMetrics = {
    ascent: 0.9,
    descent: 0.25,
    lineGap: 0,
    underlinePosition: -0.1,
    underlineThickness: 0.05,
  };
  return {
    resolve: (r) => ({
      id: r.family,
      family: r.family,
      bold: r.bold,
      italic: r.italic,
      synthBold: false,
      synthItalic: false,
      substituted: false,
      metrics,
    }),
    advances: (_face, text) => {
      const out = new Float64Array(text.length);
      for (let i = 0; i < text.length; i += 1) {
        const c = text.charCodeAt(i);
        out[i] = c >= 0xdc00 && c <= 0xdfff ? 0 : c === 32 ? 0.25 : 0.5;
      }
      return out;
    },
  };
}
