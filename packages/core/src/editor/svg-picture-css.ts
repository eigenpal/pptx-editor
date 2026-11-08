// The value rules of the SVG icon subset (svg-picture.ts): which attribute and style values
// an icon may hold, and a linear reader of `style` attributes and `<style>` sheets. A value
// outside these rules draws the picture's raster alternative.

const NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i;

/** A finite plain number. */
const isNumber = (v: string): boolean => NUMBER.test(v) && Number.isFinite(Number(v));

/** A length: a finite number with an optional `px`. */
export const isLength = (v: string): boolean => isNumber(v.trim().replace(/px$/i, ''));

/** A list of finite numbers, by spaces or commas. */
export function numbers(v: string): number[] | null {
  const words = v.trim().split(/[\s,]+/);
  if (words.length === 1 && words[0] === '') return [];
  return words.every(isNumber) ? words.map(Number) : null;
}

/** A solid color: `none`, `currentColor`, a hex color, `rgb()`/`rgba()`, or a color name. */
export function isColor(v: string): boolean {
  const s = v.trim();
  if (/^#(?:[0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(s)) return true;
  if (/^[a-z]{1,30}$/i.test(s)) return true;
  const m = /^rgba?\(([^()]*)\)$/i.exec(s);
  return !!m && (numbers(m[1]!.replace(/%/g, '')) ?? []).length >= 3;
}

/** The transform functions an icon may use, with finite numbers. */
export function isTransform(v: string): boolean {
  const rest = v
    .trim()
    .replace(
      /(?:matrix|translate|scale|rotate|skewX|skewY)\(([^()]*)\)[\s,]*/g,
      (_, args: string) => (numbers(args) ? '' : '!')
    );
  return rest === '';
}

/** The `url(#id)` of a clip path, or undefined. */
export function clipId(v: string): string | undefined {
  return /^url\(\s*["']?#([\w.:-]+)["']?\s*\)$/.exec(v.trim())?.[1];
}

const KEYWORDS: Readonly<Record<string, RegExp>> = {
  'stroke-linecap': /^(?:butt|round|square)$/,
  'stroke-linejoin': /^(?:miter|round|bevel)$/,
  display: /^[a-z-]{1,20}$/,
  visibility: /^(?:visible|hidden|collapse)$/,
  overflow: /^(?:visible|hidden|auto|scroll)$/,
};

/**
 * Whether `value` is valid for the paint property `name`. `found` receives the id of a clip
 * path, an opacity below 1, and the even-odd rule. Undefined for a name that is not a paint
 * property.
 */
export function paintValue(
  name: string,
  value: string,
  found: { clips: string[]; opacity: boolean; evenOdd: boolean }
): boolean | undefined {
  const v = value.trim();
  switch (name) {
    case 'fill':
    case 'stroke':
      return isColor(v);
    case 'stroke-width':
      return isLength(v);
    case 'stroke-miterlimit':
      return isNumber(v) && Number(v) <= 10;
    case 'opacity':
      if (!isNumber(v)) return false;
      if (Number(v) < 1) found.opacity = true;
      return true;
    case 'fill-opacity':
    case 'stroke-opacity':
      return isNumber(v);
    case 'clip-path': {
      const id = clipId(v);
      if (id === undefined) return v === 'none';
      found.clips.push(id);
      return true;
    }
    case 'enable-background':
      return !/[()\\]/.test(v);
    case 'fill-rule':
    case 'clip-rule':
      if (v === 'evenodd') found.evenOdd = true;
      return v === 'evenodd' || v === 'nonzero';
    default:
      return KEYWORDS[name]?.test(v);
  }
}

/** Declarations (`name: value; ...`) with valid paint properties only. */
export function declarations(
  css: string,
  found: { clips: string[]; opacity: boolean; evenOdd: boolean }
): boolean {
  for (const part of css.split(';')) {
    if (part.trim() === '') continue;
    const colon = part.indexOf(':');
    if (colon < 0) return false;
    const name = part.slice(0, colon).trim().toLowerCase();
    const value = part.slice(colon + 1).replace(/!important\s*$/i, '');
    if (paintValue(name, value, found) !== true) return false;
  }
  return true;
}

/** A selector list of simple selectors: a type, classes, and ids, with no combinator. */
const SELECTORS = /^\s*[\w*.#-]+(?:\s*,\s*[\w*.#-]+)*\s*$/;

/**
 * Whether a style sheet holds only rules of simple selectors and valid paint declarations: no
 * at-rule, escape, custom property, or nested block. Comments are read as nothing.
 */
export function styleSheet(
  css: string,
  found: { clips: string[]; opacity: boolean; evenOdd: boolean }
): boolean {
  if (/[@\\]|var\(/i.test(css)) return false;
  let text = '';
  let from = 0;
  for (let at = css.indexOf('/*'); at >= 0; at = css.indexOf('/*', from)) {
    text += `${css.slice(from, at)} `;
    const end = css.indexOf('*/', at + 2);
    from = end < 0 ? css.length : end + 2;
  }
  text += css.slice(from);
  let at = 0;
  for (;;) {
    const open = text.indexOf('{', at);
    if (open < 0) return text.slice(at).trim() === '';
    const close = text.indexOf('}', open);
    if (close < 0) return false;
    const body = text.slice(open + 1, close);
    if (body.includes('{') || !SELECTORS.test(text.slice(at, open))) return false;
    if (!declarations(body, found)) return false;
    at = close + 1;
  }
}
