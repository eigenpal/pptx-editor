// The text measurer port. Layout is DOM-free; the host injects a measurer. The HarfBuzz
// measurer shapes with the real font bytes; the fixed measurer is a deterministic stand-in for
// tests. Both answer the same two questions: which face draws this request, and how wide is
// this text in that face.

import type { PathCommand } from '@eigenpal/ooxml-core';
import { eastAsianScriptOf } from '../../cascade/script-fonts.ts';

/** A font request from a run. */
export interface FontRequest {
  readonly family: string;
  readonly bold: boolean;
  readonly italic: boolean;
  /**
   * The script of the text the face draws. `eastAsian` text (the `a:ea` slot) whose family is
   * missing falls back to an East Asian face (`eastAsianFallbackFamilies`) before
   * `DEFAULT_FALLBACK_FAMILY`. Absent for Latin text.
   */
  readonly script?: 'eastAsian';
  /**
   * The language of East Asian text (a BCP 47 tag, `lang` or `altLang`): it orders the East
   * Asian fallback faces, so Chinese and Korean text draw with their own glyph forms.
   */
  readonly lang?: string;
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
  /**
   * Another face that draws a character `face` lacks: the East Asian fallback faces first for
   * East Asian characters (in the order of `lang`), then each other face in registration order. Undefined when `face`
   * has the character or no face has it. Optional: a measurer without it measures a missing
   * character with an estimated width (the "Missing glyphs" approximation).
   */
  fallbackFace?(face: ResolvedFace, codePoint: number, lang?: string): ResolvedFace | undefined;
  /**
   * The outline of the glyph that the face's character map gives a character, in em with y up
   * and the origin on the baseline at the glyph's start. Undefined when the face lacks the
   * character or its outline. Optional: without it, warped text (`a:prstTxWarp`) draws flat.
   */
  glyphOutline?(face: ResolvedFace, codePoint: number): readonly PathCommand[] | undefined;
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

/**
 * The substitutes of `family`. A deck names any family, so only own keys count: a family named
 * `constructor` has none.
 */
export function fontSubstitutes(family: string): readonly string[] {
  const key = family.toLowerCase();
  return Object.hasOwn(FONT_SUBSTITUTES, key) ? FONT_SUBSTITUTES[key]! : [];
}

/** Whether measuring `family` with `substitute` keeps the original's advance widths. */
export function isMetricCompatible(family: string, substitute: string): boolean {
  return METRIC_COMPATIBLE[family.toLowerCase()] === substitute.toLowerCase();
}

/** The face used when nothing else matches. */
export const DEFAULT_FALLBACK_FAMILY = 'Liberation Sans';

const JAPANESE_FACES = ['Noto Sans CJK JP', 'Noto Sans JP'];
const SIMPLIFIED_FACES = ['Noto Sans CJK SC', 'Noto Sans SC'];
const TRADITIONAL_FACES = ['Noto Sans CJK TC', 'Noto Sans TC'];
const KOREAN_FACES = ['Noto Sans CJK KR', 'Noto Sans KR'];

/**
 * The faces that draw East Asian text whose own family is missing, in order, when the host
 * supplies one, for Japanese text or text with no language: the pan-CJK face of
 * `@docx-editor.dev/fonts-cjk` (Noto Sans CJK JP), the regional Noto Sans faces of
 * `googleFonts({ cjk: true })`, and the Noto Sans CJK faces a system may have. They cover the
 * text, not the original's metrics. `eastAsianFallbackFamilies` gives the order for a language.
 */
export const EAST_ASIAN_FALLBACK_FAMILIES: readonly string[] = Object.freeze([
  ...JAPANESE_FACES,
  ...SIMPLIFIED_FACES,
  ...TRADITIONAL_FACES,
  ...KOREAN_FACES,
]);

const FALLBACKS_BY_SCRIPT: Readonly<Record<string, readonly string[]>> = {
  Hans: Object.freeze([
    ...SIMPLIFIED_FACES,
    ...TRADITIONAL_FACES,
    ...JAPANESE_FACES,
    ...KOREAN_FACES,
  ]),
  Hant: Object.freeze([
    ...TRADITIONAL_FACES,
    ...SIMPLIFIED_FACES,
    ...JAPANESE_FACES,
    ...KOREAN_FACES,
  ]),
  Hang: Object.freeze([
    ...KOREAN_FACES,
    ...JAPANESE_FACES,
    ...SIMPLIFIED_FACES,
    ...TRADITIONAL_FACES,
  ]),
};

/**
 * The East Asian fallback faces in the order for a language (a BCP 47 tag): the faces of its
 * script first (Simplified Chinese for `Hans`, Traditional Chinese for `Hant`, Korean for
 * `Hang`), so the text draws with its own glyph forms. Japanese and no language:
 * `EAST_ASIAN_FALLBACK_FAMILIES`.
 */
export function eastAsianFallbackFamilies(lang?: string): readonly string[] {
  const script = eastAsianScriptOf(lang);
  return (script && FALLBACKS_BY_SCRIPT[script]) || EAST_ASIAN_FALLBACK_FAMILIES;
}

/**
 * A deterministic measurer for tests and for hosts without fonts: every code unit is 0.5 em,
 * spaces 0.25 em, and every glyph outline is a box. Never use it for fidelity work.
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
    // A box for every character but a space: 0.4 em wide and 0.7 em tall.
    glyphOutline: (_face, cp) =>
      cp === 32
        ? []
        : [
            { op: 'M', x: 0.05, y: 0 },
            { op: 'L', x: 0.45, y: 0 },
            { op: 'L', x: 0.45, y: 0.7 },
            { op: 'L', x: 0.05, y: 0.7 },
            { op: 'Z' },
          ],
  };
}
