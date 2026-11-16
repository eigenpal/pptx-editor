// The font styles that parts of a deck ask for: regular, and bold, italic, and bold italic when
// their text properties set them. The session loads only these faces of the opening slide's
// substitutes before its first paint, and the host's font origins hear them as
// `FontResolutionRequest.faces`. The scan reads no cascade, so it can miss a face; that face
// loads with the other faces after the open, and the text lays out again when it arrives.

import { attr, NS, type Element, type XmlPart } from '../store/index.ts';

/** One face of a family: its weight and style, as `@docx-editor.dev/fonts` takes them. */
export interface FaceStyle {
  readonly weight: 400 | 700;
  readonly style: 'normal' | 'italic';
}

export const REGULAR: FaceStyle = { weight: 400, style: 'normal' };
const BOLD: FaceStyle = { weight: 700, style: 'normal' };
const ITALIC: FaceStyle = { weight: 400, style: 'italic' };
const BOLD_ITALIC: FaceStyle = { weight: 700, style: 'italic' };

/** The four faces of each packaged substitute family. */
export const ALL_FACE_STYLES: readonly FaceStyle[] = [REGULAR, BOLD, ITALIC, BOLD_ITALIC];

/** What the text properties of one part set. */
export interface PartStyles {
  readonly bold: boolean;
  readonly italic: boolean;
}

/**
 * Elements whose `b` and `i` set the face of text: run properties (ECMA-376 Part 1, 21.1.2.3.9
 * `rPr`, 21.1.2.3.2 `defRPr`, 21.1.2.2.3 `endParaRPr`) and table cell text (20.1.4.2.29
 * `tcTxStyle`, whose values are `on` and `off`).
 */
const STYLED = new Set(['rPr', 'defRPr', 'endParaRPr', 'tcTxStyle']);
const ON = new Set(['1', 'true', 'on']);
const cache = new WeakMap<XmlPart, PartStyles>();

/**
 * The styles `part` sets. A table or a chart counts as bold: the built-in table styles make the
 * header row bold, and a chart title is bold unless its part says otherwise.
 */
export function partStyles(part: XmlPart): PartStyles {
  const hit = cache.get(part);
  if (hit) return hit;
  let bold = false;
  let italic = false;
  const walk = (el: Element, depth: number): void => {
    if (depth > 256 || (bold && italic)) return;
    if (el.ns === NS.a && STYLED.has(el.local)) {
      if (ON.has(attr(el, 'b') ?? '')) bold = true;
      if (ON.has(attr(el, 'i') ?? '')) italic = true;
    } else if ((el.ns === NS.a && el.local === 'tbl') || (el.ns === NS.c && el.local === 'chart'))
      bold = true;
    for (const c of el.children) if (c.type === 'element') walk(c, depth + 1);
  };
  walk(part.root, 0);
  const out = { bold, italic };
  cache.set(part, out);
  return out;
}

/** The faces that parts with these styles ask for. Regular is always one of them. */
export function faceStylesOf(styles: readonly PartStyles[]): FaceStyle[] {
  const bold = styles.some((s) => s.bold);
  const italic = styles.some((s) => s.italic);
  return [
    REGULAR,
    ...(bold ? [BOLD] : []),
    ...(italic ? [ITALIC] : []),
    ...(bold && italic ? [BOLD_ITALIC] : []),
  ];
}

/** A key for one face of one family, case-folded. */
export function faceKey(family: string, face: FaceStyle): string {
  return `${family.toLowerCase()}\u0000${face.weight}\u0000${face.style}`;
}
