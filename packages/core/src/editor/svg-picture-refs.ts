// The text rules of the SVG picture check (svg-picture.ts): how a value reads after the
// browser's own decoding, which `data:` URIs can hold a document, which ids a value links to,
// and which style sheet rules can set a marker or link to an id. Each rule errs on the side of
// a refusal or a larger count: a refused picture draws its raster alternative.

/** CSS escapes (`\64` or `\:`). */
const CSS_ESCAPE = /\\(?:([0-9a-f]{1,6})[ \t\n\r\f]?|([\s\S]))/gi;

/** `value` with its CSS escapes decoded. */
export function cssUnescape(value: string): string {
  if (!value.includes('\\')) return value;
  return value.replace(CSS_ESCAPE, (_, hex: string | undefined, ch: string | undefined) => {
    if (!hex) return ch ?? '';
    const code = parseInt(hex, 16);
    return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : '�';
  });
}

/** Spaces and control characters: a URL parser drops them, or a type reads past them. */
const BLANKS = /[\u0000- \u007f]+/g;

/** The `data:` media types that hold no document: raster pictures and fonts. */
const SAFE_DATA =
  /^(?:image\/(?:png|jpeg|jpg|pjpeg|gif|webp|bmp|avif)|font\/[\w.+-]+|application\/(?:x-)?font-[\w.+-]+|application\/vnd\.ms-fontobject)$/;

/**
 * Whether `value` (an attribute value or a style sheet) holds a `data:` URI that can hold a
 * document: any media type but a raster picture or a font. A URI with no comma is not a
 * `data:` URI and loads nothing. `application/octet-stream` is a font only in the `src` of an
 * `@font-face` rule of a style sheet (`sheet`).
 */
export function loadsDocument(value: string, sheet = false): boolean {
  if (!value.includes(':') && !value.includes('\\')) return false;
  const plain = cssUnescape(value).replace(BLANKS, '').toLowerCase();
  const faces = sheet ? [...plain.matchAll(/@font-face\{[^}]*\}/g)] : [];
  for (let at = plain.indexOf('data:'); at >= 0; at = plain.indexOf('data:', at + 5)) {
    const rest = plain.slice(at + 5, at + 5 + 256);
    const comma = rest.indexOf(',');
    if (comma < 0) continue;
    const type = rest.slice(0, comma).split(';')[0]!;
    if (SAFE_DATA.test(type)) continue;
    const inFace = faces.some((f) => at > f.index && at < f.index + f[0].length);
    if (type === 'application/octet-stream' && inFace) continue;
    return true;
  }
  return false;
}

/** A fragment id as the browser reads it: percent-decoded. */
function fragmentId(raw: string): string {
  try {
    return decodeURIComponent(raw);
  } catch {
    return raw;
  }
}

/** The same-document id of a link (`href="#id"`), or undefined. */
export function hrefId(value: string): string | undefined {
  const plain = value.replace(BLANKS, '');
  return plain.startsWith('#') && plain.length > 1 ? fragmentId(plain.slice(1)) : undefined;
}

/** The ids of every `url(#id)` in CSS text or an attribute value, escapes decoded. */
export function urlIds(value: string): string[] {
  if (!/url/i.test(value) && !value.includes('\\')) return [];
  const plain = cssUnescape(value);
  const ids: string[] = [];
  for (const m of plain.matchAll(/url\(([^)]*)\)/gi)) {
    const target = m[1]!.replace(BLANKS, '').replace(/^["']|["']$/g, '');
    if (target.startsWith('#') && target.length > 1) ids.push(fragmentId(target.slice(1)));
  }
  return ids;
}

/** Whether CSS text (a style or a sheet) sets a marker property. */
export const setsMarker = (css: string): boolean =>
  /marker/i.test(cssUnescape(css.replace(/\/\*[\s\S]*?\*\//g, '')));

/**
 * One style sheet rule that links to ids, read by its rightmost compound selector only (its
 * type, classes, and id): every element the whole selector matches, it matches too.
 */
export interface SheetRule {
  readonly type: string;
  readonly classes: readonly string[];
  readonly id: string | undefined;
  /** The ids its declarations link to, and whether they set a marker. */
  readonly ids: readonly string[];
  readonly marker: boolean;
}

/**
 * The rules of the style sheets `css`, or null when this check cannot read them: a custom
 * property (`var(`), an escape, or an at-rule other than `@font-face` (`@import`, `@media`).
 */
export function sheetRules(css: string): SheetRule[] | null {
  const sheet = css.replace(/\/\*[\s\S]*?\*\//g, '');
  if (/var\(|\\|@(?!font-face)/i.test(sheet)) return null;
  const rules: SheetRule[] = [];
  for (const rule of sheet.matchAll(/([^{}]*)\{([^{}]*)\}/g)) {
    const ids = urlIds(rule[2]!);
    if (ids.length === 0) continue;
    const marker = /marker/i.test(rule[2]!);
    for (const selector of rule[1]!.split(',')) {
      const compound = (
        selector
          .trim()
          .split(/[\s>+~]+/)
          .pop() ?? ''
      ).replace(/^[^|]*\|/, '');
      const type = /^([a-z][\w-]*)/i.exec(compound)?.[1]?.toLowerCase() ?? '';
      const classes = [...compound.matchAll(/\.([\w-]+)/g)].map((m) => m[1]!);
      const id = /#([\w-]+)/.exec(compound)?.[1];
      rules.push({ type, classes, id, ids, marker });
    }
  }
  return rules;
}
