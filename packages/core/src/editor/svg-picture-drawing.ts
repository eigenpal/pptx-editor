// An SVG picture as drawing records, for a writer with no browser (the PDF export). The editor
// lets the browser draw an SVG picture only when it is a plain icon (`svg-picture.ts`). This
// module reads the same icons into paths, solid paints, transforms, clip paths, and group
// opacity, so the export draws the same vector art as the editor. A picture outside the icon
// subset, or an icon feature this reader does not draw, gives `undefined`, and the caller
// draws the picture's raster alternative.
//
// The part is untrusted input. It goes through the icon check first, which bounds its bytes,
// elements, shapes, path vertices, and `<use>` chains, and refuses scripts, animation, external
// links, and style sheets with at-rules. Here, `<use>` links resolve only to elements of the
// same document, a cycle or a chain longer than `MAX_USE_DEPTH` refuses, and the drawing holds
// at most `MAX_DRAWING_COMMANDS` path commands. Nothing is fetched or run.

import { PRESET_COLORS, readXml, type PathCommand, type XmlNode } from '@eigenpal/ooxml-core';
import type { Matrix } from '../layout/index.ts';
import {
  drawingNumber,
  drawingPath,
  IDENTITY,
  MAX_DRAWING_COMMANDS,
  MAX_STYLE_RULES,
  multiply,
  num,
  numberList,
  parseTransform,
  rectPath,
  shapePath,
  SvgUnsupported,
  transformPath,
  viewBoxMatrix,
} from './svg-picture-geometry.ts';
import { svgDrawable, MAX_SVG_BYTES } from './svg-picture.ts';

/** A solid paint: red, green, and blue in 0–255, and its opacity in 0–1. @internal */
export interface SvgPictureColor {
  readonly r: number;
  readonly g: number;
  readonly b: number;
  readonly a: number;
}

/**
 * One shape of an SVG picture: its path in its own user space and the matrix to the view box.
 * @internal
 */
export interface SvgPictureShape {
  readonly kind: 'shape';
  readonly matrix: Matrix;
  readonly path: readonly PathCommand[];
  readonly fill?: { readonly color: SvgPictureColor; readonly evenOdd: boolean };
  readonly stroke?: {
    readonly color: SvgPictureColor;
    readonly width: number;
    readonly cap: 'butt' | 'round' | 'square';
    readonly join: 'miter' | 'round' | 'bevel';
    readonly miterLimit: number;
  };
}

/**
 * Items drawn together: with an opacity, as one layer, and inside a clip path in view box space.
 * @internal
 */
export interface SvgPictureGroup {
  readonly kind: 'group';
  readonly opacity: number;
  readonly clip?: { readonly path: readonly PathCommand[]; readonly evenOdd: boolean };
  readonly items: readonly SvgPictureItem[];
}

/** A shape or a group of an SVG picture. @internal */
export type SvgPictureItem = SvgPictureShape | SvgPictureGroup;

/** An SVG picture as drawing records in the space of its view box. @internal */
export interface SvgPictureDrawing {
  /** The view box: `x`, `y`, width, and height. The picture stretches it over its frame. */
  readonly viewBox: readonly [number, number, number, number];
  readonly items: readonly SvgPictureItem[];
}

/** The longest `<use>` chain the reader follows. */
const MAX_USE_DEPTH = 8;
/** The most elements the reader visits, with each `<use>` expanded. */
const MAX_VISITS = 50_000;

const SVG_NS = 'http://www.w3.org/2000/svg';
const XLINK_NS = 'http://www.w3.org/1999/xlink';

type Element = Extract<XmlNode, { type: 'element' }>;

/** CSS color names, from the DrawingML preset colors without their `dk`, `lt`, `med` short forms. */
const COLOR_NAMES: ReadonlyMap<string, string> = new Map([
  ...Object.keys(PRESET_COLORS)
    .filter((k) => !/^(?:dk|lt|med)[A-Z]/.test(k))
    .map((k) => [k.toLowerCase(), PRESET_COLORS[k]!] as [string, string]),
  ['rebeccapurple', '663399'],
]);

/** A paint value: a color, `none`, or undefined when the value is not one (CSS ignores it). */
function parseColor(value: string): SvgPictureColor | 'none' | undefined {
  const s = value.trim().toLowerCase();
  if (s === 'none') return 'none';
  if (s === 'transparent') return { r: 0, g: 0, b: 0, a: 0 };
  // No `color` property reaches an icon, so the current color is the initial black.
  if (s === 'currentcolor') return { r: 0, g: 0, b: 0, a: 1 };
  const hex = /^#([0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/.exec(s)?.[1];
  if (hex) {
    const full = hex.length <= 4 ? [...hex].map((c) => c + c).join('') : hex;
    const byte = (i: number) => parseInt(full.slice(i * 2, i * 2 + 2), 16);
    return { r: byte(0), g: byte(1), b: byte(2), a: full.length === 8 ? byte(3) / 255 : 1 };
  }
  const fn = /^rgba?\(([^()]*)\)$/.exec(s)?.[1];
  if (fn !== undefined) {
    const parts = fn.trim().split(/\s*[,/]\s*|\s+/);
    if (parts.length < 3 || parts.length > 4) return undefined;
    const channel = (p: string): number | undefined => {
      const n = num(p.replace(/%$/, ''));
      if (n === undefined) return undefined;
      return Math.max(0, Math.min(255, p.endsWith('%') ? (n * 255) / 100 : n));
    };
    const [r, g, b] = parts.slice(0, 3).map(channel);
    let a = 1;
    if (parts[3] !== undefined) {
      const n = num(parts[3].replace(/%$/, ''));
      if (n === undefined) return undefined;
      a = Math.max(0, Math.min(1, parts[3].endsWith('%') ? n / 100 : n));
    }
    if (r === undefined || g === undefined || b === undefined) return undefined;
    return { r, g, b, a };
  }
  const named = COLOR_NAMES.get(s);
  if (named) return parseColor(`#${named}`);
  return undefined;
}

// --- Style --------------------------------------------------------------------------------

const PROPERTIES = new Set([
  'fill',
  'fill-opacity',
  'fill-rule',
  'stroke',
  'stroke-width',
  'stroke-opacity',
  'stroke-linecap',
  'stroke-linejoin',
  'stroke-miterlimit',
  'opacity',
  'clip-path',
  'clip-rule',
  'display',
  'visibility',
  'overflow',
]);
const INHERITED = new Set([
  'fill',
  'fill-opacity',
  'fill-rule',
  'stroke',
  'stroke-width',
  'stroke-opacity',
  'stroke-linecap',
  'stroke-linejoin',
  'stroke-miterlimit',
  'clip-rule',
  'visibility',
]);

interface Rule {
  readonly type: string | undefined;
  readonly classes: readonly string[];
  readonly id: string | undefined;
  /** Ids, classes, and types in the selector, for the cascade order. */
  readonly specificity: readonly [number, number, number];
  /** The rule's place in the document, and after `index` its place in the cascade order. */
  order: number;
  readonly decls: readonly (readonly [string, string, boolean])[];
}

/**
 * The rules of a picture, sorted once in cascade order and filed by one key of each selector:
 * its id, else its first class, else its type, else none. An element looks only in the files
 * of its own id, classes, and type, and in the file of rules without a key.
 */
interface RuleIndex {
  readonly byId: Map<string, Rule[]>;
  readonly byClass: Map<string, Rule[]>;
  readonly byType: Map<string, Rule[]>;
  readonly any: Rule[];
}

function indexRules(rules: Rule[]): RuleIndex {
  const sorted = [...rules].sort(
    (a, b) =>
      a.specificity[0] - b.specificity[0] ||
      a.specificity[1] - b.specificity[1] ||
      a.specificity[2] - b.specificity[2] ||
      a.order - b.order
  );
  const out: RuleIndex = { byId: new Map(), byClass: new Map(), byType: new Map(), any: [] };
  const file = (map: Map<string, Rule[]>, key: string, rule: Rule) => {
    const list = map.get(key);
    if (list) list.push(rule);
    else map.set(key, [rule]);
  };
  sorted.forEach((rule, at) => {
    rule.order = at;
    if (rule.id !== undefined) file(out.byId, rule.id, rule);
    else if (rule.classes.length > 0) file(out.byClass, rule.classes[0]!, rule);
    else if (rule.type !== undefined) file(out.byType, rule.type, rule);
    else out.any.push(rule);
  });
  return out;
}

/** Whether a paint property's value is one a browser reads; an invalid one is dropped. */
function validValue(name: string, value: string): boolean {
  if (value === 'inherit') return true;
  if (name === 'fill' || name === 'stroke') return parseColor(value) !== undefined;
  return true;
}

function declarationList(css: string): [string, string, boolean][] {
  const out: [string, string, boolean][] = [];
  for (const part of css.split(';')) {
    if (part.trim() === '') continue;
    const colon = part.indexOf(':');
    if (colon < 0) continue;
    const name = part.slice(0, colon).trim().toLowerCase();
    let value = part.slice(colon + 1).trim();
    const important = /!\s*important$/i.test(value);
    if (important) value = value.replace(/!\s*important$/i, '').trim();
    if (PROPERTIES.has(name)) out.push([name, value, important]);
  }
  return out;
}

function parseSheet(css: string, rules: Rule[]): void {
  const text = css.replace(/\/\*[\s\S]*?(?:\*\/|$)/g, ' ');
  let at = 0;
  for (;;) {
    const open = text.indexOf('{', at);
    if (open < 0) return;
    const close = text.indexOf('}', open);
    if (close < 0) return;
    const decls = declarationList(text.slice(open + 1, close));
    for (const sel of text.slice(at, open).split(',')) {
      const m = /^\s*(\*|[A-Za-z][\w-]*)?((?:[.#][\w-]+)*)\s*$/.exec(sel);
      if (!m) throw new SvgUnsupported('selector');
      const parts = m[2]!.match(/[.#][\w-]+/g) ?? [];
      const ids = parts.filter((p) => p[0] === '#').map((p) => p.slice(1));
      if (ids.length > 1) throw new SvgUnsupported('selector');
      const classes = parts.filter((p) => p[0] === '.').map((p) => p.slice(1));
      const type = m[1] && m[1] !== '*' ? m[1] : undefined;
      if (rules.length >= MAX_STYLE_RULES) throw new SvgUnsupported('too many style rules');
      const specificity = [ids.length, classes.length, type ? 1 : 0] as const;
      rules.push({ type, classes, id: ids[0], specificity, order: rules.length, decls });
    }
    at = close + 1;
  }
}

type Style = Readonly<Record<string, string>>;

const declaredCache = new WeakMap<Element, Map<string, string>>();

/**
 * The declared values of an element: presentation attributes, then sheet rules, then `style`,
 * without values a browser drops as invalid. Read once per element of one picture.
 */
function declared(el: Element, local: string, rules: RuleIndex): Map<string, string> {
  const known = declaredCache.get(el);
  if (known) return known;
  const out = new Map<string, string>();
  const set = (name: string, value: string) => {
    if (validValue(name, value)) out.set(name, value);
  };
  for (const name in el.attributes)
    if (PROPERTIES.has(name)) set(name, el.attributes[name]!.trim());
  const classes = [...new Set((el.attributes.class ?? '').trim().split(/\s+/))].filter(Boolean);
  const id = el.attributes.id;
  const candidates = [
    ...(id !== undefined ? (rules.byId.get(id) ?? []) : []),
    ...classes.flatMap((c) => rules.byClass.get(c) ?? []),
    ...(rules.byType.get(local) ?? []),
    ...rules.any,
  ];
  const matching = candidates
    .filter(
      (r) =>
        (r.type === undefined || r.type === local) &&
        (r.id === undefined || r.id === id) &&
        r.classes.every((c) => classes.includes(c))
    )
    .sort((a, b) => a.order - b.order);
  const inline = declarationList(el.attributes.style ?? '');
  for (const pass of [false, true]) {
    for (const r of matching)
      for (const [name, value, important] of r.decls) if (important === pass) set(name, value);
    for (const [name, value, important] of inline) if (important === pass) set(name, value);
  }
  declaredCache.set(el, out);
  return out;
}

const INITIAL: Style = {
  fill: 'black',
  'fill-opacity': '1',
  'fill-rule': 'nonzero',
  stroke: 'none',
  'stroke-width': '1',
  'stroke-opacity': '1',
  'stroke-linecap': 'butt',
  'stroke-linejoin': 'miter',
  'stroke-miterlimit': '4',
  'clip-rule': 'nonzero',
  visibility: 'visible',
};

function computed(parent: Style, own: Map<string, string>): Style {
  const out: Record<string, string> = Object.create(null);
  for (const name of INHERITED) out[name] = parent[name]!;
  for (const [name, value] of own) {
    if (value === 'inherit') {
      if (parent[name] !== undefined) out[name] = parent[name]!;
      continue;
    }
    out[name] = value;
  }
  return out;
}

const opacity = (v: string | undefined): number => {
  const n = num(v);
  return n === undefined ? 1 : Math.max(0, Math.min(1, n));
};

// --- Reader -------------------------------------------------------------------------------

/**
 * The drawing of an SVG picture part, or `undefined` when the editor draws its raster
 * alternative (`svgDrawable`) or the picture uses an icon feature this reader does not draw:
 * a nested `<svg>`, a clip path in `objectBoundingBox` units, a clip path of more than one
 * shape, or a clip path that is itself clipped. `pptx-to-pdf` draws it.
 * @internal
 */
export function svgPictureDrawing(bytes: Uint8Array): SvgPictureDrawing | undefined {
  if (bytes.length === 0 || bytes.length > MAX_SVG_BYTES || !svgDrawable(bytes)) return undefined;
  try {
    return read(bytes);
  } catch (e) {
    if (e instanceof SvgUnsupported) return undefined;
    throw e;
  }
}

function read(bytes: Uint8Array): SvgPictureDrawing | undefined {
  const decoded = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
  // The icon check refuses a document type with an internal subset; another one is dropped.
  const doctype = decoded.indexOf('<!DOCTYPE');
  const text =
    doctype < 0
      ? decoded
      : decoded.slice(0, doctype) + decoded.slice(decoded.indexOf('>', doctype) + 1);
  const xml = readXml(text, { maxBytes: bytes.length + 1, maxElements: 10_000 });
  if (!xml.ok) return undefined;
  const root = xml.nodes.find((n): n is Element => n.type === 'element');
  if (!root) return undefined;

  // Namespaces: an element is SVG when its prefix maps to the SVG namespace.
  type Scope = ReadonlyMap<string, string>;
  const scopeOf = (el: Element, parent: Scope): Scope => {
    let next: Map<string, string> | undefined;
    for (const name in el.attributes)
      if (name === 'xmlns' || name.startsWith('xmlns:'))
        (next ??= new Map(parent)).set(name === 'xmlns' ? '' : name.slice(6), el.attributes[name]!);
    return next ?? parent;
  };
  const svgLocal = (el: Element, scope: Scope): string | undefined => {
    const colon = el.name.indexOf(':');
    const ns = scope.get(colon < 0 ? '' : el.name.slice(0, colon));
    return ns === SVG_NS ? el.name.slice(colon + 1) : undefined;
  };
  const hrefOf = (el: Element, scope: Scope): string | undefined => {
    for (const name in el.attributes) {
      const colon = name.indexOf(':');
      if (name.slice(colon + 1) !== 'href') continue;
      if (colon >= 0 && scope.get(name.slice(0, colon)) !== XLINK_NS) continue;
      return el.attributes[name]!.trim();
    }
    return undefined;
  };

  // Ids and style sheets, in document order.
  const ids = new Map<string, { el: Element; scope: Scope }>();
  const sheetRules: Rule[] = [];
  const index = (el: Element, scope: Scope): void => {
    const inner = scopeOf(el, scope);
    const local = svgLocal(el, inner);
    if (local === undefined) return;
    const id = el.attributes.id;
    if (id !== undefined && !ids.has(id)) ids.set(id, { el, scope: inner });
    if (local === 'style') {
      for (const c of el.children) if (c.type === 'text') parseSheet(c.value, sheetRules);
      return;
    }
    for (const c of el.children) if (c.type === 'element') index(c, inner);
  };
  const rootScope = scopeOf(root, new Map());
  index(root, new Map());
  const rules = indexRules(sheetRules);

  const ra = root.attributes;
  const vbList = numberList(ra.viewBox);
  const w = num(ra.width);
  const h = num(ra.height);
  const viewBox: [number, number, number, number] | undefined =
    vbList && vbList.length === 4 && vbList[2]! > 0 && vbList[3]! > 0
      ? (vbList as [number, number, number, number])
      : w !== undefined && h !== undefined && w > 0 && h > 0 && !/%/.test(ra.width + ra.height)
        ? [0, 0, w, h]
        : undefined;
  if (!viewBox) return undefined;
  const box = viewBox.map(drawingNumber) as [number, number, number, number];
  if (!(box[2] > 0 && box[3] > 0)) return undefined;

  let commands = 0;
  let visited = 0;
  const count = (path: readonly PathCommand[]) => {
    commands += path.length;
    if (commands > MAX_DRAWING_COMMANDS) throw new SvgUnsupported('too many commands');
  };

  /** The clip path an element names, in view box space, or undefined for none. */
  const clipFor = (
    value: string | undefined,
    m: Matrix,
    style: Style,
    uses: readonly Element[]
  ): SvgPictureGroup['clip'] | undefined => {
    if (value === undefined || value === 'none') return undefined;
    const id = /^url\(\s*["']?#([\w.:-]+)["']?\s*\)$/.exec(value)?.[1];
    const target = id === undefined ? undefined : ids.get(id);
    if (!target || svgLocal(target.el, target.scope) !== 'clipPath')
      throw new SvgUnsupported('clip path');
    const units = target.el.attributes.clipPathUnits?.trim();
    if (units === 'objectBoundingBox') throw new SvgUnsupported('clip units');
    const own = declared(target.el, 'clipPath', rules);
    if (own.has('clip-path') && own.get('clip-path') !== 'none')
      throw new SvgUnsupported('nested clip');
    const items: SvgPictureItem[] = [];
    const clipM = target.el.attributes.transform
      ? multiply(m, parseTransform(target.el.attributes.transform) ?? throwUnsupported())
      : m;
    const clipStyle = computed(style, own);
    for (const c of target.el.children)
      if (c.type === 'element') walk(c, target.scope, clipM, clipStyle, uses, items, true);
    const shapes = items.filter((i): i is SvgPictureShape => i.kind === 'shape');
    if (shapes.length !== items.length || shapes.length !== 1)
      throw new SvgUnsupported('clip shape');
    const shape = shapes[0]!;
    return { path: transformPath(shape.path, shape.matrix), evenOdd: shape.fill?.evenOdd ?? false };
  };

  /** Read `el` and its children into `out`. `clipping` reads the children of a clip path. */
  const walk = (
    el: Element,
    parentScope: Scope,
    parentM: Matrix,
    parentStyle: Style,
    uses: readonly Element[],
    out: SvgPictureItem[],
    clipping: boolean,
    useBox?: { x: number; y: number; w?: number; h?: number }
  ): void => {
    if (++visited > MAX_VISITS) throw new SvgUnsupported('too many elements');
    const scope = scopeOf(el, parentScope);
    const local = svgLocal(el, scope);
    // Elements of other namespaces, and the ones that draw nothing.
    if (local === undefined) return;
    if (['title', 'desc', 'metadata', 'style', 'defs', 'clipPath'].includes(local)) return;
    if (local === 'symbol' && !useBox) return;
    if (local === 'svg') throw new SvgUnsupported('nested svg');
    const own = declared(el, local, rules);
    if (own.get('display') === 'none') return;
    const style = computed(parentStyle, own);
    let m = parentM;
    if (el.attributes.transform !== undefined) {
      const t = parseTransform(el.attributes.transform);
      if (!t) throw new SvgUnsupported('transform');
      m = multiply(m, t);
    }
    const groupOpacity = clipping ? 1 : opacity(own.get('opacity'));
    const clipValue = own.get('clip-path');

    if (local === 'use') {
      const href = hrefOf(el, scope);
      const id = href?.startsWith('#') ? href.slice(1) : undefined;
      const target = id === undefined ? undefined : ids.get(id);
      if (!target) return;
      if (uses.includes(target.el) || uses.length >= MAX_USE_DEPTH)
        throw new SvgUnsupported('use chain');
      const x = num(el.attributes.x) ?? 0;
      const y = num(el.attributes.y) ?? 0;
      const inner: SvgPictureItem[] = [];
      const um = multiply(m, [1, 0, 0, 1, x, y]);
      const box = {
        x: 0,
        y: 0,
        w: num(el.attributes.width),
        h: num(el.attributes.height),
      };
      walk(target.el, target.scope, um, style, [...uses, target.el], inner, clipping, box);
      emit(inner, groupOpacity, clipValue, m, style, uses, out, clipping);
      return;
    }
    if (local === 'g' || local === 'symbol') {
      const inner: SvgPictureItem[] = [];
      let cm = m;
      let viewportClip: SvgPictureGroup['clip'] | undefined;
      if (local === 'symbol') {
        const vb = numberList(el.attributes.viewBox);
        const vw = useBox?.w ?? viewBox[2];
        const vh = useBox?.h ?? viewBox[3];
        if (vb && vb.length === 4 && vb[2]! > 0 && vb[3]! > 0) {
          if (!(vw > 0 && vh > 0)) return;
          cm = multiply(m, viewBoxMatrix(vb, 0, 0, vw, vh, el.attributes.preserveAspectRatio));
        }
        // A symbol's viewport clips, as its overflow is hidden.
        const ov = own.get('overflow');
        if (vw > 0 && vh > 0 && (ov === undefined || ov === 'hidden' || ov === 'scroll'))
          viewportClip = {
            path: transformPath(rectPath(0, 0, vw, vh), m),
            evenOdd: false,
          };
      }
      for (const c of el.children)
        if (c.type === 'element') walk(c, scope, cm, style, uses, inner, clipping);
      if (viewportClip && !clipping && inner.length > 0)
        emit(
          [{ kind: 'group', opacity: 1, clip: viewportClip, items: inner }],
          groupOpacity,
          clipValue,
          m,
          style,
          uses,
          out,
          clipping
        );
      else emit(inner, groupOpacity, clipValue, m, style, uses, out, clipping);
      return;
    }
    if (!['path', 'rect', 'circle', 'ellipse', 'line', 'polyline', 'polygon'].includes(local))
      throw new SvgUnsupported(local);
    const path = shapePath(local, el.attributes);
    if (!path || path.length === 0) return;
    count(path);
    if (style.visibility !== 'visible') return;
    if (clipping) {
      if (clipValue !== undefined && clipValue !== 'none')
        throw new SvgUnsupported('clip inside a clip path');
      out.push({
        kind: 'shape',
        matrix: m,
        path,
        fill: { color: { r: 0, g: 0, b: 0, a: 1 }, evenOdd: style['clip-rule'] === 'evenodd' },
      });
      return;
    }
    const fillColor = local === 'line' ? 'none' : (parseColor(style.fill!) ?? 'none');
    const strokeColor = parseColor(style.stroke!) ?? 'none';
    const strokeWidth = num(style['stroke-width']) ?? 1;
    const fa = opacity(style['fill-opacity']);
    const sa = opacity(style['stroke-opacity']);
    const fill: SvgPictureShape['fill'] =
      fillColor !== 'none' && fillColor.a * fa > 0
        ? {
            color: { ...fillColor, a: fillColor.a * fa },
            evenOdd: style['fill-rule'] === 'evenodd',
          }
        : undefined;
    const cap = style['stroke-linecap'];
    const join = style['stroke-linejoin'];
    const stroke: SvgPictureShape['stroke'] =
      strokeColor !== 'none' && strokeColor.a * sa > 0 && strokeWidth > 0
        ? {
            color: { ...strokeColor, a: strokeColor.a * sa },
            width: strokeWidth,
            cap: cap === 'round' || cap === 'square' ? cap : ('butt' as const),
            join: join === 'round' || join === 'bevel' ? join : ('miter' as const),
            miterLimit: Math.max(1, num(style['stroke-miterlimit']) ?? 4),
          }
        : undefined;
    if (!fill && !stroke) return;
    const shape: SvgPictureShape = {
      kind: 'shape',
      matrix: m,
      path,
      ...(fill ? { fill } : {}),
      ...(stroke ? { stroke } : {}),
    };
    // One paint takes the element's opacity itself; a fill and a stroke draw as one layer.
    if (groupOpacity < 1 && !(fill && stroke) && clipValue === undefined) {
      const a = (c: SvgPictureColor) => ({ ...c, a: c.a * groupOpacity });
      out.push({
        ...shape,
        ...(fill ? { fill: { ...fill, color: a(fill.color) } } : {}),
        ...(stroke ? { stroke: { ...stroke, color: a(stroke.color) } } : {}),
      });
      return;
    }
    emit([shape], groupOpacity, clipValue, m, style, uses, out, clipping);
  };

  /** Add `items` to `out`, as a group when they have an opacity or a clip path. */
  const emit = (
    items: SvgPictureItem[],
    groupOpacity: number,
    clipValue: string | undefined,
    m: Matrix,
    style: Style,
    uses: readonly Element[],
    out: SvgPictureItem[],
    clipping: boolean
  ): void => {
    if (items.length === 0 || groupOpacity <= 0) return;
    const clip = clipping ? undefined : clipFor(clipValue, m, style, uses);
    if (clipping && clipValue !== undefined && clipValue !== 'none')
      throw new SvgUnsupported('clip inside a clip path');
    if (groupOpacity >= 1 && !clip) out.push(...items);
    else out.push({ kind: 'group', opacity: groupOpacity, ...(clip ? { clip } : {}), items });
  };

  // The root: its own style and transform-free children, in view box space.
  const own = declared(root, 'svg', rules);
  if (own.get('display') === 'none') return { viewBox, items: [] };
  const style = computed(INITIAL, own);
  const items: SvgPictureItem[] = [];
  for (const c of root.children)
    if (c.type === 'element') walk(c, rootScope, IDENTITY, style, [], items, false);
  const rootOpacity = opacity(own.get('opacity'));
  const top: SvgPictureItem[] = [];
  emit(items, rootOpacity, own.get('clip-path'), IDENTITY, style, [], top, false);
  return { viewBox: box, items: top.map(checked) };
}

/** An item with every number checked by `drawingNumber`, or a throw for one out of range. */
function checked(item: SvgPictureItem): SvgPictureItem {
  if (item.kind === 'group')
    return {
      ...item,
      ...(item.clip ? { clip: { ...item.clip, path: drawingPath(item.clip.path) } } : {}),
      items: item.items.map(checked),
    };
  return {
    ...item,
    matrix: item.matrix.map(drawingNumber) as unknown as Matrix,
    path: drawingPath(item.path),
    ...(item.stroke
      ? {
          stroke: {
            ...item.stroke,
            width: drawingNumber(item.stroke.width),
            miterLimit: drawingNumber(item.stroke.miterLimit),
          },
        }
      : {}),
  };
}

function throwUnsupported(): never {
  throw new SvgUnsupported('value');
}
