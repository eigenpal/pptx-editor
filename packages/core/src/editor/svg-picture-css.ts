// The text rules of the SVG picture subset (svg-picture.ts): one linear pass over a style
// sheet, a `style` attribute, or a presentation attribute value, and the links and lengths
// they hold. Each pass reads its text a bounded number of times, with no backtracking
// pattern, so a hostile sheet costs what its length costs.

/** What a pass found. */
export interface CssScan {
  /** The ids that `url(#id)` names. */
  readonly ids: string[];
  /** Whether it sets `opacity`: each element it reaches draws in a layer of its own. */
  opacity: boolean;
  /** Whether it can scale shapes: a scaling transform, or a geometry property. */
  scales: boolean;
  /** The widest stroke and the largest font it sets, in CSS pixels. */
  stroke: number;
  font: number;
}

/** Spaces and control characters: a URL parser drops them. */
const isBlank = (c: number): boolean => c <= 0x20 || c === 0x7f;

const isNameChar = (c: number): boolean =>
  (c >= 0x61 && c <= 0x7a) || (c >= 0x41 && c <= 0x5a) || (c >= 0x30 && c <= 0x39) || c === 0x2d;

/** `text` with no spaces or control characters. */
export function squeeze(text: string): string {
  let out = '';
  let from = 0;
  for (let i = 0; i < text.length; i += 1)
    if (isBlank(text.charCodeAt(i))) {
      out += text.slice(from, i);
      from = i + 1;
    }
  return from === 0 ? text : out + text.slice(from);
}

/** A fragment id as the browser reads it: percent-decoded. */
export function fragmentId(raw: string): string {
  try {
    return decodeURIComponent(raw);
  } catch {
    return raw;
  }
}

/** CSS pixels for each unit. A percent or a font-relative unit reads as unbounded. */
const UNITS: Readonly<Record<string, number>> = {
  '': 1,
  px: 1,
  pt: 4 / 3,
  pc: 16,
  in: 96,
  cm: 96 / 2.54,
  mm: 96 / 25.4,
  q: 96 / 101.6,
};
const LENGTH = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?/i;

/** The largest length among the words of `value`, in CSS pixels; 0 for none. */
export function largestLength(value: string): number {
  let largest = 0;
  for (const word of value.split(/[\s,/]+/)) {
    const m = LENGTH.exec(word);
    if (!m) continue;
    const unit = word.slice(m[0].length).toLowerCase();
    const scale = UNITS[unit];
    largest = Math.max(largest, scale === undefined ? Infinity : Math.abs(Number(m[0])) * scale);
  }
  return largest;
}

/** Whether a transform can make a shape larger: a scale, a matrix, a skew, or a 3D one. */
export const scalingTransform = (value: string): boolean =>
  /scale|matrix|skew|3d|perspective/i.test(value);

/**
 * Properties that pull in what the subset leaves out (filters, masks, markers), or whose cost
 * grows with the drawn length or with a shape a style sets (`stroke-dasharray`, `d`).
 */
const refusedProperty = (name: string): boolean =>
  name.includes('filter') ||
  name.includes('mask') ||
  name.includes('marker') ||
  name === 'stroke-dasharray' ||
  name === 'd';
/** Geometry properties: a style can size a shape past its attributes. */
const GEOMETRY: ReadonlySet<string> = new Set(['width', 'height', 'r', 'rx', 'ry']);

/** CSS text with each comment as one space; an open comment runs to the end. */
function withoutComments(css: string): string {
  let out = '';
  let from = 0;
  for (let at = css.indexOf('/*'); at >= 0; at = css.indexOf('/*', from)) {
    out += `${css.slice(from, at)} `;
    const end = css.indexOf('*/', at + 2);
    from = end < 0 ? css.length : end + 2;
  }
  return from === 0 ? css : out + css.slice(from);
}

/**
 * Scan CSS text: a style sheet (`sheet`), or the declarations of a `style` attribute. Returns
 * null when the text uses what the subset leaves out: an escape, a custom property, an at-rule
 * other than `@font-face` and `@media`, a refused property, a `data:` URI outside an
 * `@font-face` rule, or a `url()` other than `url(#id)` and a font in an `@font-face` rule.
 */
export function scanCss(css: string, sheet: boolean): CssScan | null {
  if (css.includes('\\')) return null;
  const clean = withoutComments(css);
  const text = clean.toLowerCase();
  // Custom properties, and pictures from a list or a string, which `url()` does not name.
  if (text.includes('var(') || text.includes('image-set') || text.includes('src(')) return null;
  const scan: CssScan = { ids: [], opacity: false, scales: false, stroke: 0, font: 0 };
  /** For each open block, whether it is an `@font-face` rule. */
  const faces: boolean[] = [];
  let face = false;
  // A declaration starts the text, or follows `{` or `;`.
  let declaration = true;
  for (let i = 0; i < text.length; i += 1) {
    const c = text.charCodeAt(i);
    if (c === 0x40) {
      let j = i + 1;
      while (j < text.length && isNameChar(text.charCodeAt(j))) j += 1;
      const rule = text.slice(i + 1, j);
      if (!sheet || (rule !== 'font-face' && rule !== 'media')) return null;
      face = rule === 'font-face';
      i = j - 1;
      continue;
    }
    if (c === 0x7b) {
      if (!sheet) return null;
      faces.push(face || (faces.length > 0 && faces[faces.length - 1]!));
      face = false;
      declaration = true;
      continue;
    }
    if (c === 0x7d) {
      faces.pop();
      declaration = false;
      continue;
    }
    if (c === 0x3b) {
      declaration = true;
      continue;
    }
    if (isBlank(c)) continue;
    if (declaration && isNameChar(c)) {
      let j = i;
      while (j < text.length && isNameChar(text.charCodeAt(j))) j += 1;
      let k = j;
      while (k < text.length && isBlank(text.charCodeAt(k))) k += 1;
      // A property name, not a selector: a colon follows it.
      if (text.charCodeAt(k) === 0x3a) {
        const name = text.slice(i, j);
        if (refusedProperty(name)) return null;
        let end = k + 1;
        while (end < text.length && !';{}'.includes(text[end]!)) end += 1;
        const value = text.slice(k + 1, end);
        if (name === 'opacity') scan.opacity = true;
        if (name === 'transform' && scalingTransform(value)) scan.scales = true;
        if (GEOMETRY.has(name)) scan.scales = true;
        if (name === 'stroke-width') scan.stroke = Math.max(scan.stroke, largestLength(value));
        if (name === 'font-size' || name === 'font')
          scan.font = Math.max(scan.font, largestLength(value));
      }
      declaration = false;
      i = j - 1;
      continue;
    }
    declaration = false;
    // A `data:` URI outside `url()`: in a string, or a list of pictures.
    if (c === 0x64 && text.startsWith('data:', i) && !faces[faces.length - 1]) return null;
    if (c === 0x75 && text.startsWith('url(', i)) {
      const end = text.indexOf(')', i + 4);
      if (end < 0) return null;
      const target = squeeze(clean.slice(i + 4, end)).replace(/^["']|["']$/g, '');
      if (target.startsWith('#') && target.length > 1) scan.ids.push(fragmentId(target.slice(1)));
      else if (!(faces[faces.length - 1] && target.toLowerCase().startsWith('data:'))) return null;
      i = end;
    }
  }
  return scan;
}

/** The `data:` media types of the pictures an `<image>` may draw. */
const RASTER = /^data:image\/(?:png|jpeg|jpg|gif|webp|bmp)(?:;[^,]*)?;base64,/;

/** The base64 payload of a raster picture `data:` URI, or undefined for any other link. */
export function rasterData(href: string): string | undefined {
  const plain = squeeze(href);
  const head = RASTER.exec(plain.slice(0, 128).toLowerCase());
  return head ? plain.slice(head[0].length) : undefined;
}

/**
 * Whether an attribute value holds a link that the subset leaves out: a `data:` URI (a comma
 * makes it one; "Sales data: 2024" is text) or an external URL. `href` and `url()` values are
 * checked where they are read; this catches the rest.
 */
export function foreignLink(value: string): boolean {
  if (!value.includes(':') && !value.includes('//')) return false;
  const plain = squeeze(value).toLowerCase();
  const data = plain.indexOf('data:');
  if (data >= 0 && plain.indexOf(',', data) >= 0) return true;
  return /(?:^|[("'=])(?:https?:|file:|\/\/)/.test(plain);
}
