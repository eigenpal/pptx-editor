// Shapes from a DrawingML clipboard package: what presentation applications put on the system
// clipboard for copied shapes. Browsers do not give that format to a web page, so a host that
// can read the system clipboard (a desktop wrapper) passes its bytes in.
//
// The package is an OPC package. Its main part is a drawing (`a:graphic`) that holds a locked
// canvas (`lc:lockedCanvas`, ECMA-376 Part 1, 20.3): the copied shapes as DrawingML shapes
// (`a:sp`, `a:pic`, `a:cxnSp`, `a:grpSp`, `a:graphicFrame`), at their positions on the source
// slide. Pictures are image relationships of the drawing. This module turns the canvas into
// this editor's shape clipboard payload, and the `paste` command validates it as it validates
// every payload. Theme colors and theme fonts stay references, so the shapes take the theme
// of the slide they are pasted on.

import type { ClipboardPayload, ClipboardRelationship } from '../contracts/index.ts';
import {
  CONTENT_TYPES_PART,
  contentTypeOf,
  decodeXml,
  NS,
  parseXmlPart,
  readZip,
  relationshipsOf,
  specOf,
  type NodeSpec,
  type PackagePart,
  type ZipLimits,
} from '../store/index.ts';
import { toBase64 } from '../automation/commands/base64.ts';

/**
 * The system clipboard types that hold a DrawingML clipboard package, by platform. A desktop
 * host reads one of them and passes the bytes to `drawingClipboardPayload`. @public
 */
export const DRAWING_CLIPBOARD_TYPES = Object.freeze({
  /** The pasteboard type on macOS. */
  macos: 'com.microsoft.Art--GVML-ClipFormat',
  /** The registered clipboard format name on Windows. */
  windows: 'Art::GVML ClipFormat',
});

const DRAWING_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const LOCKED_CANVAS = 'http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas';
const TABLE = 'http://schemas.openxmlformats.org/drawingml/2006/table';

const MAX_PACKAGE = 64 * 1024 * 1024;
const LIMITS: ZipLimits = Object.freeze({
  maxEntries: 1000,
  maxTotalBytes: 128 * 1024 * 1024,
  maxRatio: 200,
});
const MAX_DEPTH = 32;

/** Each DrawingML shape and the names of its non-visual parts. */
const KINDS: Readonly<Record<string, { readonly nv: string; readonly cnv: string }>> = {
  sp: { nv: 'nvSpPr', cnv: 'cNvSpPr' },
  pic: { nv: 'nvPicPr', cnv: 'cNvPicPr' },
  cxnSp: { nv: 'nvCxnSpPr', cnv: 'cNvCxnSpPr' },
  grpSp: { nv: 'nvGrpSpPr', cnv: 'cNvGrpSpPr' },
  graphicFrame: { nv: 'nvGraphicFramePr', cnv: 'cNvGraphicFramePr' },
};
/** Shape properties that keep their content and move to the PresentationML namespace. */
const PROPERTIES = new Set(['spPr', 'grpSpPr', 'style', 'blipFill', 'xfrm']);

type Child = NodeSpec | { readonly text: string };
const elements = (s: NodeSpec): NodeSpec[] =>
  (s.children ?? []).filter((c): c is NodeSpec => !('text' in c));
const isA = (s: NodeSpec, local: string): boolean => s.ns === NS.a && s.local === local;
/** The same element in the PresentationML namespace. */
const inP = (s: NodeSpec, children: readonly Child[] = s.children ?? []): NodeSpec => ({
  ns: NS.p,
  local: s.local,
  prefix: 'p',
  attrs: s.attrs ?? [],
  children,
});
const pEl = (local: string, children: readonly Child[] = []): NodeSpec => ({
  ns: NS.p,
  local,
  prefix: 'p',
  attrs: [],
  children,
});

/** `p:nvSpPr` and its kin: the drawing properties, the locks, and empty application data. */
function nonVisual(nv: NodeSpec, local: string, cnv: string): NodeSpec | null {
  const props = elements(nv).find((c) => isA(c, 'cNvPr'));
  if (!props) return null;
  const locks = elements(nv).find((c) => isA(c, cnv));
  return pEl(local, [inP(props), locks ? inP(locks) : pEl(cnv), pEl('nvPr')]);
}

/** Whether a graphic frame holds a table: its other content needs parts that do not travel. */
function holdsTable(frame: NodeSpec): boolean {
  const graphic = elements(frame).find((c) => isA(c, 'graphic'));
  const data = graphic && elements(graphic).find((c) => isA(c, 'graphicData'));
  return data?.attrs?.some((a) => a.local === 'uri' && a.value === TABLE) ?? false;
}

/** A DrawingML shape as the PresentationML shape, or null for one that cannot paste. */
function convert(shape: NodeSpec, depth: number): NodeSpec | null {
  const kind = shape.ns === NS.a ? KINDS[shape.local] : undefined;
  if (!kind || depth > MAX_DEPTH) return null;
  if (shape.local === 'graphicFrame' && !holdsTable(shape)) return null;
  const out: NodeSpec[] = [];
  for (const c of elements(shape)) {
    if (isA(c, kind.nv)) {
      const nv = nonVisual(c, kind.nv, kind.cnv);
      if (!nv) return null;
      out.push(nv);
    } else if (c.ns === NS.a && PROPERTIES.has(c.local)) out.push(inP(c));
    else if (isA(c, 'txSp')) {
      // The text of a DrawingML shape: its body is the PresentationML `p:txBody`.
      const body = elements(c).find((k) => isA(k, 'txBody'));
      if (body) out.push(inP(body));
    } else if (isA(c, 'graphic')) out.push(c);
    else if (shape.local === 'grpSp' && c.ns === NS.a && KINDS[c.local]) {
      const member = convert(c, depth + 1);
      if (member) out.push(member);
    }
    // Anything else (an extension list, an unknown element) stays behind.
  }
  if (out[0]?.local !== kind.nv) return null;
  return pEl(shape.local, out);
}

/** The `a:xfrm` of a canvas or a group. */
const xfrmOf = (props: NodeSpec | undefined): NodeSpec | undefined =>
  props && elements(props).find((c) => isA(c, 'xfrm'));
const num = (el: NodeSpec | undefined, name: string): number => {
  const v = Number(el?.attrs?.find((a) => a.ns === '' && a.local === name)?.value);
  return Number.isFinite(v) ? v : 0;
};

/**
 * Scale a top-level shape from the canvas's child space to the canvas, when the two differ:
 * `at` maps a position, `by` scales a size. The canvas sits where the shapes were copied from.
 */
function placed(shape: NodeSpec, at: (v: number, axis: 'x' | 'y') => number, by: number[]) {
  const move = (xfrm: NodeSpec): NodeSpec => ({
    ...xfrm,
    children: (xfrm.children ?? []).map((c) => {
      if ('text' in c || !(isA(c, 'off') || isA(c, 'ext'))) return c;
      return {
        ...c,
        attrs: (c.attrs ?? []).map((a) => {
          const v = Number(a.value);
          if (a.ns !== '' || !Number.isFinite(v)) return a;
          if (a.local === 'x' || a.local === 'y')
            return { ...a, value: `${Math.round(at(v, a.local))}` };
          if (a.local === 'cx') return { ...a, value: `${Math.round(v * by[0]!)}` };
          if (a.local === 'cy') return { ...a, value: `${Math.round(v * by[1]!)}` };
          return a;
        }),
      };
    }),
  });
  return {
    ...shape,
    children: (shape.children ?? []).map((c) => {
      if ('text' in c || c.ns !== NS.p) return c;
      if (c.local === 'xfrm') return move(c);
      if (c.local !== 'spPr' && c.local !== 'grpSpPr') return c;
      return {
        ...c,
        children: (c.children ?? []).map((x) => ('text' in x || !isA(x, 'xfrm') ? x : move(x))),
      };
    }),
  };
}

/** The relationship ids a subtree uses (`r:*` attributes). */
function relIds(spec: NodeSpec, out: Set<string>, depth = 0): Set<string> {
  if (depth > 2 * MAX_DEPTH) return out;
  for (const a of spec.attrs ?? []) if (a.ns === NS.r) out.add(a.value);
  for (const c of elements(spec)) relIds(c, out, depth + 1);
  return out;
}

/**
 * Turn a DrawingML clipboard package into this editor's shape clipboard payload, to paste
 * with the `paste` command. The package is the bytes of the system clipboard type in
 * {@link DRAWING_CLIPBOARD_TYPES}, which presentation applications write for copied shapes.
 * Shapes, pictures, connectors, groups, and tables convert; charts, diagrams, and embedded
 * objects need parts that do not travel and stay behind. Returns null when the bytes are not
 * such a package or hold no shape that converts. The bytes are untrusted: the package reads
 * within size limits, and the paste validates the payload again. @public
 */
export function drawingClipboardPayload(bytes: Uint8Array): ClipboardPayload | null {
  if (!(bytes instanceof Uint8Array) || bytes.length === 0 || bytes.length > MAX_PACKAGE)
    return null;
  const zip = readZip(bytes, LIMITS);
  if (!zip.ok) return null;
  const parts = new Map<string, PackagePart>();
  const xml = (partName: string): boolean => {
    const data = zip.entries.get(partName);
    if (!data) return false;
    const tree = parseXmlPart(partName, decodeXml(data));
    if (tree.ok) parts.set(partName, { kind: 'xml', partName, xml: tree.part });
    return tree.ok;
  };
  xml(CONTENT_TYPES_PART);
  for (const name of zip.entries.keys()) if (name.endsWith('.rels')) xml(name);
  const pkg = { parts };
  const main = relationshipsOf(pkg, '/').find((r) => r.type === DRAWING_REL && r.target);
  if (!main?.target || !xml(main.target)) return null;
  const drawing = parts.get(main.target);
  if (drawing?.kind !== 'xml') return null;
  // `a:graphic` → `a:graphicData` (locked canvas) → `lc:lockedCanvas`.
  const root = specOf(drawing.xml.root);
  const data = isA(root, 'graphic') ? elements(root).find((c) => isA(c, 'graphicData')) : undefined;
  const canvas =
    data && elements(data).find((c) => c.ns === LOCKED_CANVAS && c.local === 'lockedCanvas');
  if (!canvas) return null;

  // The canvas maps its child space (the source slide's positions) onto itself.
  const frame = xfrmOf(elements(canvas).find((c) => isA(c, 'grpSpPr')));
  const ext = frame && elements(frame).find((c) => isA(c, 'ext'));
  const chOff = frame && elements(frame).find((c) => isA(c, 'chOff'));
  const chExt = frame && elements(frame).find((c) => isA(c, 'chExt'));
  const sx = num(chExt, 'cx') > 0 && num(ext, 'cx') > 0 ? num(ext, 'cx') / num(chExt, 'cx') : 1;
  const sy = num(chExt, 'cy') > 0 && num(ext, 'cy') > 0 ? num(ext, 'cy') / num(chExt, 'cy') : 1;
  const origin = { x: num(chOff, 'x'), y: num(chOff, 'y') };
  const at = (v: number, axis: 'x' | 'y'): number =>
    origin[axis] + (v - origin[axis]) * (axis === 'x' ? sx : sy);

  const shapes: NodeSpec[] = [];
  for (const c of elements(canvas)) {
    const shape = c.ns === NS.a && KINDS[c.local] ? convert(c, 0) : null;
    if (shape) shapes.push(sx === 1 && sy === 1 ? shape : placed(shape, at, [sx, sy]));
  }
  if (!shapes.length) return null;

  // The pictures the shapes use travel with them; other relationships stay behind.
  const used = new Set<string>();
  for (const s of shapes) relIds(s, used);
  const relationships: ClipboardRelationship[] = [];
  for (const rel of relationshipsOf(pkg, main.target)) {
    if (!used.has(rel.id) || rel.external || rel.type !== IMAGE_REL || !rel.target) continue;
    const media = zip.entries.get(rel.target);
    const contentType = contentTypeOf(pkg, rel.target);
    if (!media || !contentType?.startsWith('image/')) continue;
    relationships.push({
      id: rel.id,
      type: rel.type,
      part: rel.target,
      contentType,
      base64: toBase64(media),
    });
  }
  return { version: 1, document: 'drawing-clipboard', slidePart: '', shapes, relationships };
}
