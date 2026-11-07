// Whether an SVG picture part is safe for the browser to draw. The browser draws an SVG
// picture on the page's main thread, and some elements multiply the work of others: `<use>`
// copies its target, so ten levels of ten copies each draw 10^10 shapes, and a marker draws
// its content at each vertex of a path. This check reads the part with the bounded XML reader
// and counts the elements that the picture draws, with every `<use>` expanded and every marker
// drawn at each vertex of the elements that a marker property reaches. A part over the limit
// draws the picture's raster alternative instead. So does a part that the reader refuses (an
// internal DTD subset, too deep, too many elements), one with a `<use>` cycle, one with a
// marker inside marker content, and one that links to a `data:` URI that can hold a document
// (an SVG, XML, HTML, or unknown type): a nested document there escapes this count. Raster
// pictures and fonts in `data:` URIs draw. The check needs no DOM.

import { readXml, type XmlNode } from '@eigenpal/ooxml-core';

/**
 * The most elements an SVG picture may draw, with each `<use>` expanded and each marker drawn
 * at each vertex. Icons and diagrams draw a few thousand at most.
 */
export const MAX_SVG_DRAWN_ELEMENTS = 200_000;
/** The most `<use>` links that an expansion may follow in a chain. */
const MAX_USE_CHAIN = 16;

/** Elements whose content draws only where a `<use>` or a marker property links to it. */
const UNDRAWN: ReadonlySet<string> = new Set(['defs', 'symbol', 'marker']);
/** The elements that draw markers (SVG 2, 11.6.2 "markable elements"). */
const MARKABLE: ReadonlySet<string> = new Set(['path', 'line', 'polyline', 'polygon']);

type Element = Extract<XmlNode, { type: 'element' }>;

const localName = (name: string): string => name.slice(name.indexOf(':') + 1);

/** The `#id` that a `<use>` links to, from `href` or any prefixed `href` (`xlink:href`). */
function linkTarget(el: Element): string | undefined {
  for (const [name, value] of Object.entries(el.attributes))
    if (localName(name) === 'href' && value.trim().startsWith('#')) return value.trim().slice(1);
  return undefined;
}

/** A document type declaration without an internal subset: it declares nothing to expand. */
const PLAIN_DOCTYPE = /<!DOCTYPE[^[>]*>/;

// --- data: URIs ----------------------------------------------------------------------------

/** CSS escapes (`\64` or `\:`): a style can spell `data:` with them. */
const CSS_ESCAPE = /\\(?:([0-9a-f]{1,6})\s?|(.))/gi;
/** The `data:` media types that hold no document: raster pictures and fonts. */
const SAFE_DATA =
  /^(?:image\/(?:png|jpe?g|pjpeg|gif|webp|bmp|avif)|font\/[\w.+-]+|application\/(?:x-)?font-[\w.+-]+)$/;

/** A value with CSS escapes decoded, in lower case, with no space or quote. */
function plain(value: string): string {
  return value
    .replace(CSS_ESCAPE, (_, hex: string | undefined, ch: string | undefined) =>
      hex ? String.fromCodePoint(Math.min(parseInt(hex, 16), 0x10ffff)) : (ch ?? '')
    )
    .toLowerCase()
    .replace(/[\s"']/g, '');
}

/** Whether the `data:` URI that starts `uri` (in plain form) can hold a document. */
function documentData(uri: string): boolean {
  const type = /^data:([^;,]*)/.exec(uri)?.[1];
  return type !== undefined && !SAFE_DATA.test(type);
}

/** Whether CSS text (a style sheet, a style, or a presentation attribute) loads a document. */
function cssLoadsDocument(css: string): boolean {
  if (!/data|\\/i.test(css)) return false;
  for (const m of plain(css).matchAll(/(?:url\(|@import)(data:[^)]*)/g))
    if (documentData(m[1]!)) return true;
  return false;
}

/** Whether the attribute `name` (a link, a style, or a property) loads a document. */
function attributeLoadsDocument(name: string, value: string): boolean {
  const local = localName(name);
  // A link is the whole value: only its start says what it is, so a long picture is cheap.
  if (local === 'href' || local === 'src') return documentData(plain(value.slice(0, 256)));
  return /url|\\/i.test(value) && cssLoadsDocument(value);
}

// --- Markers -------------------------------------------------------------------------------

const MARKER_ATTRIBUTES = ['marker-start', 'marker-mid', 'marker-end'];
const MARKER_DECLARATION = /marker(?:-start|-mid|-end)?\s*:([^;}]*)/gi;
const URL_ID = /url\(\s*["']?#([^"')\s]+)/g;

const urlIds = (value: string): string[] => [...value.matchAll(URL_ID)].map((m) => m[1]!);

/** The ids of the markers that `el` names in its marker attributes and its style. */
function ownMarkerIds(el: Element): string[] {
  const ids: string[] = [];
  for (const name of MARKER_ATTRIBUTES) {
    const v = el.attributes[name];
    if (v !== undefined) ids.push(...urlIds(v));
  }
  const style = el.attributes.style;
  if (style) for (const m of style.matchAll(MARKER_DECLARATION)) ids.push(...urlIds(m[1]!));
  return ids;
}

/**
 * A style sheet rule that sets a marker: its selector, read as a type, an id, and classes, or
 * null for a selector this check does not read (a combinator, an attribute, a pseudo-class),
 * which counts as matching every element.
 */
interface MarkerRule {
  readonly selector: null | {
    readonly type: string;
    readonly id?: string;
    readonly classes: readonly string[];
  };
  readonly ids: readonly string[];
}

const COMPOUND = /^([a-z][\w-]*|\*)?((?:[.#][\w-]+)*)$/i;

/** The rules of the style sheets `css` that set a marker. */
function markerRules(css: string): MarkerRule[] {
  const rules: MarkerRule[] = [];
  const sheet = css.replace(/\/\*[\s\S]*?\*\//g, '');
  for (const rule of sheet.matchAll(/([^{}]*)\{([^{}]*)\}/g)) {
    const ids = [...rule[2]!.matchAll(MARKER_DECLARATION)].flatMap((m) => urlIds(m[1]!));
    if (ids.length === 0) continue;
    for (const text of rule[1]!.split(',')) {
      const m = COMPOUND.exec(text.trim());
      if (!m || !text.trim()) {
        rules.push({ selector: null, ids });
        continue;
      }
      const parts = m[2]!.match(/[.#][\w-]+/g) ?? [];
      const id = parts.find((p) => p.startsWith('#'))?.slice(1);
      rules.push({
        selector: {
          type: m[1] && m[1] !== '*' ? m[1] : '',
          ...(id ? { id } : {}),
          classes: parts.filter((p) => p.startsWith('.')).map((p) => p.slice(1)),
        },
        ids,
      });
    }
  }
  return rules;
}

function matches(rule: MarkerRule, el: Element): boolean {
  const s = rule.selector;
  if (!s) return true;
  if (s.type && s.type !== localName(el.name)) return false;
  if (s.id !== undefined && el.attributes.id !== s.id) return false;
  if (s.classes.length === 0) return true;
  const own = new Set((el.attributes.class ?? '').split(/\s+/));
  return s.classes.every((c) => own.has(c));
}

// --- Vertices ------------------------------------------------------------------------------

const PATH_TOKEN = /([a-z])|[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?/gi;
/**
 * The numbers of one segment of each path command. An arc counts 4, not 7: its two flags can
 * be written with no space between them and the next number.
 */
const SEGMENT: Readonly<Record<string, number>> = {
  m: 2,
  l: 2,
  t: 2,
  h: 1,
  v: 1,
  c: 6,
  s: 4,
  q: 4,
  a: 4,
};

/** At most how many vertices a markable element has: each one can draw a marker. */
function vertices(el: Element): number {
  const kind = localName(el.name);
  if (kind === 'line') return 2;
  if (kind !== 'path') {
    let numbers = 0;
    for (const _ of (el.attributes.points ?? '').matchAll(PATH_TOKEN)) numbers += 1;
    return Math.ceil(numbers / 2) + 1;
  }
  // One vertex for each segment of each command, and one for each `z`.
  let total = 0;
  let per = 2;
  let run = 0;
  const end = (): void => void (total += Math.ceil(run / per));
  for (const m of (el.attributes.d ?? '').matchAll(PATH_TOKEN)) {
    const letter = m[1]?.toLowerCase();
    if (letter === undefined) {
      run += 1;
      continue;
    }
    end();
    run = 0;
    if (letter === 'z') total += 1;
    else per = SEGMENT[letter] ?? 1;
  }
  end();
  return total + 1;
}

// --- The check -----------------------------------------------------------------------------

const results = new WeakMap<Uint8Array, boolean>();

/** Whether the SVG document `bytes` draws within the limits (see the module comment). */
export function svgDrawable(bytes: Uint8Array): boolean {
  const hit = results.get(bytes);
  if (hit !== undefined) return hit;
  const ok = check(bytes);
  results.set(bytes, ok);
  return ok;
}

function check(bytes: Uint8Array): boolean {
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
  } catch {
    return false;
  }
  const read = readXml(text.replace(PLAIN_DOCTYPE, ''), {
    maxBytes: bytes.length + 1,
    maxElements: MAX_SVG_DRAWN_ELEMENTS,
  });
  if (!read.ok) return false;
  const roots: Element[] = [];
  for (const n of read.nodes) if (n.type === 'element') roots.push(n);
  // The tree is at most `MAX_XML_DEPTH` deep: walks on an explicit stack.
  const ids = new Map<string, Element>();
  let sheets = '';
  for (const stack = [...roots]; stack.length;) {
    const el = stack.pop()!;
    const id = el.attributes.id;
    if (id !== undefined && !ids.has(id)) ids.set(id, el);
    for (const [name, value] of Object.entries(el.attributes))
      if (attributeLoadsDocument(name, value)) return false;
    if (localName(el.name) === 'style')
      for (const c of el.children) if (c.type === 'text') sheets += `${c.value}\n`;
    for (const c of el.children) if (c.type === 'element') stack.push(c);
  }
  if (cssLoadsDocument(sheets)) return false;

  // The markers that reach each element: its own, its style sheet rules', and its ancestors'.
  // Markable elements and `<use>` elements keep theirs. A marker that reaches marker content
  // draws inside each marker it is in: refused.
  const rules = /marker/i.test(sheets) ? markerRules(sheets) : [];
  const reach = new Map<Element, readonly string[]>();
  type Entry = { el: Element; inMarker: boolean; inherited: readonly string[] };
  const stack: Entry[] = roots.map((el) => ({ el, inMarker: false, inherited: [] }));
  while (stack.length) {
    const { el, inMarker, inherited } = stack.pop()!;
    const kind = localName(el.name);
    const own = ownMarkerIds(el);
    for (const rule of rules) if (matches(rule, el)) own.push(...rule.ids);
    const reached = own.length ? [...inherited, ...own] : inherited;
    const marker = inMarker || kind === 'marker';
    if (reached.length && (MARKABLE.has(kind) || kind === 'use')) {
      if (marker) return false;
      reach.set(el, reached);
    }
    for (const c of el.children)
      if (c.type === 'element') stack.push({ el: c, inMarker: marker, inherited: reached });
  }

  // The drawn size of each element: itself, its children, the target of its `<use>`, and the
  // largest marker that reaches it at each vertex. The vertices that each element draws count
  // too: a `<use>` that a marker reaches passes it to the markable elements it copies.
  type Size = { readonly drawn: number; readonly vertices: number };
  const sizes = new Map<Element, Size>();
  const active = new Set<Element>();
  const INFINITE: Size = { drawn: Infinity, vertices: Infinity };
  let chain = 0;
  const largestMarker = (names: readonly string[]): number => {
    let max = 0;
    for (const name of names) {
      const m = ids.get(name);
      if (m && localName(m.name) === 'marker') max = Math.max(max, size(m).drawn);
    }
    return max;
  };
  const size = (el: Element): Size => {
    const known = sizes.get(el);
    if (known) return known;
    if (active.has(el)) return INFINITE; // a cycle
    active.add(el);
    const kind = localName(el.name);
    let drawn = 1;
    let verts = MARKABLE.has(kind) ? vertices(el) : 0;
    for (const c of el.children) {
      if (c.type !== 'element') continue;
      // Definitions draw only through a link, which counts them there.
      if (UNDRAWN.has(localName(c.name))) drawn += 1;
      else {
        const s = size(c);
        drawn += s.drawn;
        verts += s.vertices;
      }
      if (drawn > MAX_SVG_DRAWN_ELEMENTS) break;
    }
    const target = kind === 'use' ? linkTarget(el) : undefined;
    const linked = target === undefined ? undefined : ids.get(target);
    let copied = 0;
    if (linked && drawn <= MAX_SVG_DRAWN_ELEMENTS) {
      chain += 1;
      const s = chain > MAX_USE_CHAIN ? INFINITE : size(linked);
      chain -= 1;
      drawn += s.drawn;
      verts += s.vertices;
      copied = s.vertices;
    }
    const markers = reach.get(el);
    if (markers && drawn <= MAX_SVG_DRAWN_ELEMENTS) {
      const at = (MARKABLE.has(kind) ? vertices(el) : 0) + copied;
      if (at > 0) drawn += at * largestMarker(markers);
    }
    active.delete(el);
    const s = { drawn, vertices: verts };
    sizes.set(el, s);
    return s;
  };
  let total = 0;
  for (const root of roots) total += size(root).drawn;
  return total <= MAX_SVG_DRAWN_ELEMENTS;
}
