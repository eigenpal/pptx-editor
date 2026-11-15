// The stand-in measurer of an open that does not wait for its fonts (`OpenOptions.standIn`).
//
// Until the text engine and the packaged substitutes arrive, a session cannot shape. The fixed
// measurer gives every character 0.5 em, but paint draws each character with the browser's own
// face at the measured position, so the letters of a slide drew far apart or on top of each other
// ("PPTX Edi t or"). This measurer asks the browser canvas for the width of each character in the
// font stack that paint uses (`${face.id}, sans-serif`), so the characters draw where layout put
// them. The text shows in the browser's face for the family until the substitutes arrive, as in
// DOCX Editor, and lays out again then.
//
// The family name is from the file. It goes into a quoted CSS family name, with its quotes and
// backslashes escaped and its control characters removed, so it can name a face and nothing else.
// The measurer registers no face, so a family name never changes a font of the host page.

import type { FaceMetrics, FontRequest, ResolvedFace, TextMeasurer } from '../layout/index.ts';

/** The slice of a 2D canvas context that measures text. */
interface MeasureContext {
  font: string;
  measureText(text: string): {
    readonly width: number;
    readonly fontBoundingBoxAscent?: number;
    readonly fontBoundingBoxDescent?: number;
  };
}

/** The size, in pixels, at which the canvas measures. Advances are in em, so any size works. */
const SIZE = 100;
/** A family name longer than this is cut: a font name has no use for more. */
const MAX_FAMILY = 128;

/**
 * A quoted CSS family name for `family`, or `sans-serif` when nothing of it is left. Quotes and
 * backslashes are escaped and control characters are removed, so the value names one family.
 */
export function cssFamily(family: string): string {
  const clean = family.replace(/[\u0000-\u001f\u007f-\u009f]/g, '').slice(0, MAX_FAMILY);
  if (clean.trim() === '') return 'sans-serif';
  return `"${clean.replace(/["\\]/g, (c) => `\\${c}`)}"`;
}

/** A 2D context of an `OffscreenCanvas`, or null where the runtime has none (a server, tests). */
function offscreenContext(): MeasureContext | null {
  try {
    const Canvas = (globalThis as { OffscreenCanvas?: typeof OffscreenCanvas }).OffscreenCanvas;
    const context = Canvas ? new Canvas(1, 1).getContext('2d') : null;
    return context && typeof context.measureText === 'function' ? context : null;
  } catch {
    return null;
  }
}

/**
 * A measurer that measures with the browser canvas, or null where no canvas can measure: the
 * caller then keeps the fixed measurer. `context` is for tests. @internal
 */
export function createCanvasMeasurer(
  context: MeasureContext | null = offscreenContext()
): TextMeasurer | null {
  if (!context) return null;
  const ctx = context;
  const faces = new Map<string, ResolvedFace>();
  const widths = new Map<string, Map<string, number>>();
  const fontOf = (face: ResolvedFace): string =>
    `${face.italic ? 'italic ' : ''}${face.bold ? 'bold ' : ''}${SIZE}px ${face.id}, sans-serif`;
  const metricsOf = (font: string): FaceMetrics => {
    ctx.font = font;
    const m = ctx.measureText('Hxgy');
    const ascent = (m.fontBoundingBoxAscent ?? 0) / SIZE;
    const descent = (m.fontBoundingBoxDescent ?? 0) / SIZE;
    const valid = ascent > 0 && descent >= 0 && ascent + descent < 4;
    return {
      ascent: valid ? ascent : 0.9,
      descent: valid ? descent : 0.25,
      lineGap: 0,
      underlinePosition: -0.1,
      underlineThickness: 0.05,
    };
  };
  return {
    resolve(r: FontRequest): ResolvedFace {
      const id = cssFamily(r.family);
      const key = `${r.bold ? 'b' : ''}${r.italic ? 'i' : ''}${id}`;
      let face = faces.get(key);
      if (!face) {
        const font = `${r.italic ? 'italic ' : ''}${r.bold ? 'bold ' : ''}${SIZE}px ${id}, sans-serif`;
        face = {
          id,
          family: r.family,
          bold: r.bold,
          italic: r.italic,
          // Paint asks the browser for the weight and the slant, as the canvas measured them.
          synthBold: r.bold,
          synthItalic: r.italic,
          substituted: false,
          metrics: metricsOf(font),
        };
        faces.set(key, face);
      }
      return face;
    },
    advances(face: ResolvedFace, text: string): Float64Array {
      const font = fontOf(face);
      let cache = widths.get(font);
      if (!cache) widths.set(font, (cache = new Map()));
      const out = new Float64Array(text.length);
      let set = false;
      for (let i = 0; i < text.length;) {
        const cp = text.codePointAt(i) ?? 0;
        const ch = String.fromCodePoint(cp);
        let width = cache.get(ch);
        if (width === undefined) {
          if (!set) ((ctx.font = font), (set = true));
          width = ctx.measureText(ch).width / SIZE;
          cache.set(ch, width);
        }
        // A character of two code units has its whole advance on the first.
        out[i] = width;
        i += ch.length;
      }
      return out;
    },
  };
}
