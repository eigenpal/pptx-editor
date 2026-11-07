// Slide layouts (ECMA-376 Part 1, §19.3.1.39 `sldLayout`): the layouts a deck offers, and the
// commands that put slides on a layout.
//
// Change layout re-points the slide's layout relationship and re-maps its placeholders by
// `type` and `idx` (`mapPlaceholders`): a placeholder that finds a
// place takes that place's `p:ph` and position and keeps its content; an empty one with no place
// goes; one with content and no place stays where it was (with a new `idx` when the new layout
// uses its own); a place no placeholder took gets an empty placeholder. An empty placeholder that
// an animation names, and a placeholder in `mc:AlternateContent`, stay where they are too. Reset re-applies the slide's own layout and also removes the placeholders'
// own formatting. Shapes that are not placeholders never change. Each command is one
// transaction, so one undo step.

import {
  animatedShapeIds,
  attr,
  child,
  elements,
  NS,
  path,
  presentationInfo,
  relationshipsOf,
  relsPartName,
  REL,
  shapeTreeOf,
  textOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import {
  alternateBranch,
  contentElements,
  isAlternateContent,
  resolveSlide,
  type Xfrm,
} from '../../cascade/index.ts';
import type { SlideLayoutSummary, SlideLayoutKind } from '../../contracts/neutral.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { relativeTarget, shapeIds } from './package-edits.ts';
import { specFromXml } from './shape-commands.ts';
import { removeShapes } from './shape-removal.ts';
import { emptyPlaceholderXml, FOOTER_PLACEHOLDERS, slideRef } from './slide-commands.ts';

/** One layout of the deck, in the order the Layout gallery lists them. */
export interface DeckLayout {
  /** The layout part: its identity in the open document. */
  readonly part: string;
  /** The slide master part. */
  readonly master: string;
  /** `p:cSld/@name` ("Title and Content"). */
  readonly name: string;
  /** `p:sldLayout/@type` (`ST_SlideLayoutType`, §19.7.15); `cust` when absent, as the schema says. */
  readonly type: SlideLayoutKind;
  /** The master's `p:cSld/@name`, or its theme's `a:theme/@name`. */
  readonly masterName: string;
}

/** `ST_SlideLayoutType` (§19.7.15): a file value outside it reads as `cust`. */
const LAYOUT_TYPES: ReadonlySet<string> = new Set<SlideLayoutKind>([
  'title',
  'tx',
  'twoColTx',
  'tbl',
  'txAndChart',
  'chartAndTx',
  'dgm',
  'chart',
  'txAndClipArt',
  'clipArtAndTx',
  'titleOnly',
  'blank',
  'txAndObj',
  'objAndTx',
  'objOnly',
  'obj',
  'txAndMedia',
  'mediaAndTx',
  'objOverTx',
  'txOverObj',
  'txAndTwoObj',
  'twoObjAndTx',
  'twoObjOverTx',
  'fourObj',
  'vertTx',
  'clipArtAndVertTx',
  'vertTitleAndTx',
  'vertTitleAndTxOverChart',
  'twoObj',
  'objAndTwoObj',
  'twoObjAndObj',
  'cust',
  'secHead',
  'twoTxTwoObj',
  'objTx',
  'picTx',
]);

function layoutType(value: string | undefined): SlideLayoutKind {
  return value !== undefined && LAYOUT_TYPES.has(value) ? (value as SlideLayoutKind) : 'cust';
}

const layoutCache = new WeakMap<PptxPackage, readonly DeckLayout[]>();

/**
 * The deck's layouts: masters in `p:sldMasterIdLst` order, and each master's layouts in its
 * `p:sldLayoutIdLst` order (§19.3.1.41). Memoized per package identity.
 */
export function deckLayouts(pkg: PptxPackage): readonly DeckLayout[] {
  const hit = layoutCache.get(pkg);
  if (hit) return hit;
  const presentation = presentationInfo(pkg).presentationPart;
  const out: DeckLayout[] = [];
  const seen = new Set<string>();
  const target = (source: string, rId: string | undefined): string | undefined =>
    relationshipsOf(pkg, source).find((r) => r.id === rId && !r.external)?.target;
  const masters = elements(path(xmlPartOf(pkg, presentation)?.root, 'p:sldMasterIdLst'));
  for (const m of masters) {
    const master = target(presentation, attr(m, 'r:id'));
    const root = xmlPartOf(pkg, master)?.root;
    if (!master || !root) continue;
    const theme = relationshipsOf(pkg, master).find((r) => r.type === REL.theme && !r.external);
    const masterName =
      attr(child(root, 'p:cSld'), 'name') ||
      attr(xmlPartOf(pkg, theme?.target)?.root, 'name') ||
      '';
    for (const l of elements(child(root, 'p:sldLayoutIdLst'))) {
      const part = target(master, attr(l, 'r:id'));
      const layout = xmlPartOf(pkg, part)?.root;
      if (!part || !layout || seen.has(part) || layout.local !== 'sldLayout') continue;
      seen.add(part);
      out.push({
        part,
        master,
        name: attr(child(layout, 'p:cSld'), 'name') ?? '',
        type: layoutType(attr(layout, 'type')),
        masterName,
      });
    }
  }
  layoutCache.set(pkg, out);
  return out;
}

/**
 * Layout summaries for the snapshot, reference-stable while the deck's layouts, their names,
 * and their types stay the same.
 */
export function createLayoutSummaries(): (
  pkg: PptxPackage | undefined
) => readonly SlideLayoutSummary[] {
  let last: readonly SlideLayoutSummary[] = [];
  return (pkg) => {
    const next = pkg ? deckLayouts(pkg) : [];
    const same =
      next.length === last.length &&
      next.every(
        (l, i) =>
          l.part === last[i].id &&
          l.name === last[i].name &&
          l.type === last[i].type &&
          l.master === last[i].masterId &&
          l.masterName === last[i].masterName
      );
    if (!same)
      last = next.map((l) => ({
        id: l.part,
        name: l.name,
        type: l.type,
        masterId: l.master,
        masterName: l.masterName,
      }));
    return last;
  };
}

/** The layout part a command names, or a refusal when the deck has no such layout. */
export function deckLayout(pkg: PptxPackage, id: string): string {
  if (!deckLayouts(pkg).some((l) => l.part === id))
    throw new CommandRefusal('errors.layoutNotFound');
  return id;
}

/**
 * The layout New Slide gives a slide after slide `after`: the layout of that slide,
 * except after a title slide (`type="title"`), where it takes the first Title and Content layout
 * (`obj`) of the same master. In a deck with no slides: the first master's Title Slide layout,
 * or its first layout. Layouts count in `p:sldLayoutIdLst` order.
 */
export function newSlideLayout(pkg: PptxPackage, after: number): string | undefined {
  const slides = presentationInfo(pkg).slides;
  const all = deckLayouts(pkg);
  if (!slides.length) {
    const first = all.filter((l) => l.master === all[0]?.master);
    return (first.find((l) => l.type === 'title') ?? first[0])?.part;
  }
  const source = slides[Math.max(0, Math.min(after, slides.length - 1))]?.layoutPart;
  const own = all.find((l) => l.part === source);
  if (own?.type !== 'title') return source;
  return all.find((l) => l.master === own.master && l.type === 'obj')?.part ?? source;
}

// ---------------------------------------------------------------------- placeholder mapping

/** What a slide placeholder holds, for the places that accept it. */
type Content = 'empty' | 'text' | 'picture' | 'table' | 'chart' | 'diagram' | 'media' | 'object';

/** One placeholder: its element, its `p:ph`, and the `p:ph` defaults filled in (§19.3.1.36). */
interface Placeholder {
  readonly el: Element;
  readonly ph: Element;
  readonly type: string;
  readonly idx: string | undefined;
}

const TITLES: ReadonlySet<string> = new Set(['title', 'ctrTitle']);
const TEXT_PLACES: ReadonlySet<string> = new Set(['body', 'subTitle', 'obj']);

/** The layout placeholder types that take each kind of content. `obj` takes any content. */
const ACCEPTS: Record<Exclude<Content, 'empty'>, ReadonlySet<string>> = {
  text: TEXT_PLACES,
  picture: new Set(['pic', 'clipArt', 'obj']),
  table: new Set(['tbl', 'obj']),
  chart: new Set(['chart', 'obj']),
  diagram: new Set(['dgm', 'obj']),
  media: new Set(['media', 'obj']),
  object: new Set(['obj']),
};

/** Placeholder types that are one family: an empty placeholder moves only inside its family. */
function family(type: string): string {
  if (TITLES.has(type)) return 'title';
  if (TEXT_PLACES.has(type)) return 'text';
  return type;
}

function phOf(el: Element): Element | undefined {
  const nv = elements(el).find((c) => c.ns === NS.p && c.local.startsWith('nv'));
  return path(nv, 'p:nvPr', 'p:ph');
}

function placeholderOf(el: Element): Placeholder | undefined {
  const ph = phOf(el);
  return ph ? { el, ph, type: attr(ph, 'type') ?? 'obj', idx: attr(ph, 'idx') } : undefined;
}

const inBranch = (branch: Element | undefined): Placeholder[] =>
  elements(branch)
    .map(placeholderOf)
    .filter((p): p is Placeholder => !!p);

/**
 * The placeholders the cascade reads in `branch`: a nested `mc:AlternateContent` reads as its
 * own branch (`contentElements`), so a placeholder nested in it holds its place.
 */
const readIn = (branch: Element | undefined, root: Element | undefined): Placeholder[] =>
  contentElements(branch, root)
    .map(placeholderOf)
    .filter((p): p is Placeholder => !!p);

/** The deepest `mc:AlternateContent` nesting that the branch walk follows: the XML depth cap. */
const MAX_NESTING = 256;

/** Every `mc:Choice` and `mc:Fallback` of `wrapper`, and those of each one nested in them. */
function branchesOf(wrapper: Element, depth = 0, out: Element[] = []): Element[] {
  for (const b of elements(wrapper)) {
    if (b.ns !== NS.mc || (b.local !== 'Choice' && b.local !== 'Fallback')) continue;
    out.push(b);
    if (depth < MAX_NESTING)
      for (const inner of elements(b))
        if (isAlternateContent(inner)) branchesOf(inner, depth + 1, out);
  }
  return out;
}

/**
 * The placeholders at the top of a shape tree, in tree order. With `unwrap`, a placeholder in
 * `mc:AlternateContent` counts as the cascade reads it (the layout's places).
 */
function placeholdersOf(tree: Element | undefined, root?: Element, unwrap = false): Placeholder[] {
  const out: Placeholder[] = [];
  for (const el of elements(tree)) {
    if (unwrap && isAlternateContent(el))
      for (const p of readIn(alternateBranch(el, root), root)) out.push(p);
    else {
      const p = placeholderOf(el);
      if (p) out.push(p);
    }
  }
  return out;
}

/** A slide placeholder in `mc:AlternateContent`: the wrapper, and its copy in each branch. */
interface Wrapped {
  readonly wrapper: Element;
  /** The copy the cascade reads. */
  readonly read: Placeholder;
  /**
   * Its copies in every `mc:Choice` and `mc:Fallback`. In each branch, every read copy first takes
   * the copy with the same `p:cNvPr/@id`; the reads that are left then take the copies that are
   * left, in order. No copy goes to two reads.
   */
  readonly copies: readonly Placeholder[];
}

/**
 * The slide placeholders in top-level `mc:AlternateContent`, in tree order, nested
 * `mc:AlternateContent` in its branches included.
 */
function wrappedPlaceholdersOf(tree: Element | undefined, root?: Element): Wrapped[] {
  const out: Wrapped[] = [];
  for (const wrapper of elements(tree)) {
    if (!isAlternateContent(wrapper)) continue;
    const reads = readIn(alternateBranch(wrapper, root), root);
    // A nested `mc:AlternateContent` holds copies in its own branches.
    const paired = branchesOf(wrapper).map((b) => pairCopies(reads, inBranch(b)));
    reads.forEach((read, i) => {
      out.push({
        wrapper,
        read,
        copies: paired.map((b) => b[i]).filter((p): p is Placeholder => !!p),
      });
    });
  }
  return out;
}

/**
 * The copy in one branch for each read copy, by index of `reads`. Pass 1 pairs by `p:cNvPr/@id`;
 * pass 2 gives the reads that are left the copies that are left, in tree order.
 */
function pairCopies(
  reads: readonly Placeholder[],
  branch: readonly Placeholder[]
): (Placeholder | undefined)[] {
  const taken = new Set<Placeholder>();
  const out: (Placeholder | undefined)[] = reads.map((read) => {
    const id = shapeIdOf(read.el);
    const copy =
      id !== undefined ? branch.find((p) => !taken.has(p) && shapeIdOf(p.el) === id) : undefined;
    if (copy) taken.add(copy);
    return copy;
  });
  const left = branch.filter((p) => !taken.has(p));
  let next = 0;
  for (let i = 0; i < out.length; i += 1) if (!out[i]) out[i] = left[next++];
  return out;
}

const shapeIdOf = (el: Element): string | undefined => {
  const nv = elements(el).find((c) => c.ns === NS.p && c.local.startsWith('nv'));
  return attr(child(nv, 'p:cNvPr'), 'id');
};

/** Whether a text body holds any character (a field counts). */
function hasText(txBody: Element | undefined): boolean {
  const walk = (el: Element, depth: number): boolean =>
    depth < 64 &&
    elements(el).some(
      (c) => (c.local === 't' && textOf(c).length > 0) || c.local === 'fld' || walk(c, depth + 1)
    );
  return !!txBody && walk(txBody, 0);
}

function contentOf(p: Placeholder): Content {
  const el = p.el;
  if (el.kind === 'sp') return hasText(child(el, 'p:txBody')) ? 'text' : 'empty';
  if (el.kind === 'pic') {
    const nvPr = path(el, 'p:nvPicPr', 'p:nvPr');
    return elements(nvPr).some((m) => m.local === 'videoFile' || m.local === 'audioFile')
      ? 'media'
      : 'picture';
  }
  const uri = attr(path(el, 'a:graphic', 'a:graphicData'), 'uri') ?? '';
  if (uri.endsWith('/table')) return 'table';
  if (uri.endsWith('/chart')) return 'chart';
  if (uri.endsWith('/diagram')) return 'diagram';
  return 'object';
}

/** Whether slide placeholder `s` may take a layout place of type `type`. */
function fitsType(s: Placeholder, content: Content, type: string): boolean {
  if (FOOTER_PLACEHOLDERS.has(s.type) || FOOTER_PLACEHOLDERS.has(type)) return s.type === type;
  if (content === 'empty') return family(s.type) === family(type);
  if (TITLES.has(s.type)) return TITLES.has(type);
  return ACCEPTS[content].has(type);
}

/**
 * Finds layout places for slide placeholders, by rule. Every rule and `fitsType` read only a
 * place's `type` and `idx`. So the first free place that a rule accepts is the earliest head
 * among the place lists, by type, of one scope: all places, or the places with one `idx`. A
 * deck with thousands of places matches in linear time, with the result of a scan of the
 * layout in order.
 */
class PlaceFinder {
  readonly #scopes = new Map<string, Map<string, { list: Placeholder[]; at: number }>>();
  readonly #order = new Map<Placeholder, number>();

  constructor(
    layout: readonly Placeholder[],
    private readonly taken: (t: Placeholder) => boolean
  ) {
    layout.forEach((t, i) => {
      this.#order.set(t, i);
      this.#add(ALL_PLACES, t);
      this.#add(idxScope(t.idx), t);
    });
  }

  #add(scope: string, t: Placeholder): void {
    let byType = this.#scopes.get(scope);
    if (!byType) this.#scopes.set(scope, (byType = new Map()));
    let q = byType.get(t.type);
    if (!q) byType.set(t.type, (q = { list: [], at: 0 }));
    q.list.push(t);
  }

  /** The first free place of `scope` whose type `accepts` takes and that fits `s`. */
  #first(
    scope: string,
    s: Placeholder,
    content: Content,
    accepts: (type: string) => boolean
  ): Placeholder | undefined {
    let best: Placeholder | undefined;
    for (const [type, q] of this.#scopes.get(scope) ?? []) {
      if (!accepts(type) || !fitsType(s, content, type)) continue;
      // A place once taken stays taken, so the list start moves past it for good.
      while (q.at < q.list.length && this.taken(q.list[q.at]!)) q.at += 1;
      const head = q.list[q.at];
      if (head && (!best || this.#order.get(head)! < this.#order.get(best)!)) best = head;
    }
    return best;
  }

  /**
   * The identity rules, from the closest: the same `type` and `idx`, then title to title, then
   * the same `idx`. `reservedPlaces` uses them too.
   */
  readonly identity: readonly ((s: Placeholder, c: Content) => Placeholder | undefined)[] = [
    (s, c) => this.#first(idxScope(s.idx), s, c, (type) => type === s.type),
    (s, c) =>
      TITLES.has(s.type) ? this.#first(ALL_PLACES, s, c, (t) => TITLES.has(t)) : undefined,
    (s, c) =>
      s.idx !== undefined && s.idx !== '0'
        ? this.#first(idxScope(s.idx), s, c, () => true)
        : undefined,
  ];

  /** The first free place of the same type. */
  sameType = (s: Placeholder, c: Content) => this.#first(ALL_PLACES, s, c, (t) => t === s.type);

  /** The first free place that takes the content. */
  any = (s: Placeholder, c: Content) => this.#first(ALL_PLACES, s, c, () => true);
}

const ALL_PLACES = '*';
const idxScope = (idx: string | undefined): string => (idx === undefined ? '\u0000' : `#${idx}`);

/**
 * Match slide placeholders to layout places. Each place takes at most one placeholder; slide
 * placeholders take places in tree order, in passes from the closest match: the identity rules
 * of `PlaceFinder`, then the same `type`, then any place that accepts the content (never a
 * footer place). In the last two passes, placeholders with content choose before empty ones,
 * so an empty placeholder never takes the only place that content could go to. In the
 * identity passes, an empty placeholder and a filled one choose in tree order: this order is
 * not checked against reference output (an approximation on the public fidelity page).
 */
export function mapPlaceholders(
  slide: readonly Placeholder[],
  layout: readonly Placeholder[]
): Map<Placeholder, Placeholder> {
  const out = new Map<Placeholder, Placeholder>();
  const used = new Set<Placeholder>();
  const content = new Map(slide.map((s) => [s, contentOf(s)]));
  const find = new PlaceFinder(layout, (t) => used.has(t));
  const pass = (
    pick: (s: Placeholder, c: Content) => Placeholder | undefined,
    which: (c: Content) => boolean = () => true
  ): void => {
    for (const s of slide) {
      const c = content.get(s)!;
      if (out.has(s) || !which(c)) continue;
      const t = pick(s, c);
      if (t) {
        out.set(s, t);
        used.add(t);
      }
    }
  };
  for (const rule of find.identity) pass(rule);
  const filled = (c: Content): boolean => c !== 'empty';
  const empty = (c: Content): boolean => c === 'empty';
  pass(find.sameType, filled);
  pass(find.any, filled);
  pass(find.sameType, empty);
  pass(find.any, empty);
  return out;
}

// ---------------------------------------------------------------------- the edit

/** Run attributes reset keeps: language and proofing state, as clear formatting does. */
const KEEP_RUN_ATTRS: ReadonlySet<string> = new Set([
  'lang',
  'altLang',
  'dirty',
  'err',
  'noProof',
  'smtClean',
]);
/** Run children reset keeps: links. */
const KEEP_RUN_CHILDREN: ReadonlySet<string> = new Set(['hlinkClick', 'hlinkMouseOver', 'extLst']);
const RUN_PROPS: ReadonlySet<string> = new Set(['rPr', 'endParaRPr', 'defRPr']);

function setAttr(batch: OpBatch, part: string, node: string, name: string, value: string | null) {
  batch.push({ op: 'setAttribute', part, node, name: { ns: '', local: name }, value });
}

/** Remove a text body's own formatting: body and list styles, paragraph and run properties. */
function clearTextFormatting(batch: OpBatch, part: string, txBody: Element): void {
  const bodyPr = child(txBody, 'a:bodyPr');
  if (bodyPr) {
    for (const a of bodyPr.attrs) setAttr(batch, part, bodyPr.id, a.local, null);
    for (const c of elements(bodyPr)) batch.push({ op: 'removeNode', part, node: c.id });
  }
  for (const c of elements(child(txBody, 'a:lstStyle')))
    batch.push({ op: 'removeNode', part, node: c.id });
  const walk = (el: Element, depth: number): void => {
    if (depth > 32) return;
    for (const c of elements(el)) {
      // The body and list styles are cleared above: their nodes are gone.
      if (c.ns === NS.a && (c.local === 'bodyPr' || c.local === 'lstStyle')) continue;
      if (c.ns === NS.a && c.local === 'pPr') {
        for (const a of c.attrs) if (a.local !== 'lvl') setAttr(batch, part, c.id, a.local, null);
        for (const k of elements(c)) batch.push({ op: 'removeNode', part, node: k.id });
      } else if (c.ns === NS.a && RUN_PROPS.has(c.local)) {
        for (const a of c.attrs)
          if (!KEEP_RUN_ATTRS.has(a.local))
            batch.push({
              op: 'setAttribute',
              part,
              node: c.id,
              name: { ns: a.ns, local: a.local },
              value: null,
            });
        for (const k of elements(c))
          if (!KEEP_RUN_CHILDREN.has(k.local)) batch.push({ op: 'removeNode', part, node: k.id });
      } else walk(c, depth + 1);
    }
  };
  walk(txBody, 0);
}

/**
 * `<a:xfrm>` for a resolved box with its rotation and flips (§20.1.7.6 `CT_Transform2D`: `rot`
 * in 60000ths of a degree), written first in `p:spPr` (`CT_ShapeProperties` order).
 */
function xfrmXml(x: Xfrm): string {
  const n = (v: number): string => String(Math.round(v));
  const rot = Math.round((((x.rot % 360) + 360) % 360) * 60_000) % 21_600_000;
  const attrs =
    (rot ? ` rot="${rot}"` : '') + (x.flipH ? ' flipH="1"' : '') + (x.flipV ? ' flipV="1"' : '');
  return `<a:xfrm${attrs}><a:off x="${n(x.x)}" y="${n(x.y)}"/><a:ext cx="${n(x.cx)}" cy="${n(x.cy)}"/></a:xfrm>`;
}

/**
 * The layout places that slide placeholders in `mc:AlternateContent` hold: the first place by
 * the identity rules of `PlaceFinder` that fit the placeholder, with the footer rules of `mapPlaceholders`.
 */
function reservedPlaces(
  wrapped: readonly Wrapped[],
  places: readonly Placeholder[]
): Map<Placeholder, Wrapped> {
  const out = new Map<Placeholder, Wrapped>();
  const done = new Set<Wrapped>();
  const find = new PlaceFinder(places, (t) => out.has(t));
  for (const rule of find.identity)
    for (const w of wrapped) {
      if (done.has(w)) continue;
      const t = rule(w.read, contentOf(w.read));
      if (t) {
        out.set(t, w);
        done.add(w);
      }
    }
  return out;
}

/** Give an element that keeps its place its resolved box as its own `a:xfrm`. */
export function keepBox(batch: OpBatch, part: string, el: Element, box: Xfrm | undefined): void {
  const spPr = child(el, 'p:spPr');
  if (spPr && box && !child(spPr, 'a:xfrm') && (el.kind === 'sp' || el.kind === 'pic'))
    batch.insert({
      op: 'insertNode',
      part,
      parent: spPr.id,
      index: 0,
      spec: specFromXml(xfrmXml(box)),
    });
}

/**
 * Put slide `index` on `layoutPart` and re-map its placeholders (see the file comment). With
 * `reset`, the placeholders also lose their own shape and text formatting. Returns the node ids
 * of the text placeholders whose box or text style changed, for the autofit re-fit.
 *
 * Three cases keep a placeholder where it is, with its own box, instead of re-mapping it:
 * content with no place; an empty placeholder with no place that an animation names
 * (`p:timing`), so no `p:spTgt` or `p:bldP` points at a removed shape; and a placeholder in
 * `mc:AlternateContent`, whose layout place stays taken.
 */
export function applyLayout(
  batch: OpBatch,
  index: number,
  layoutPart: string,
  reset = false
): string[] {
  const ref = slideRef(batch.pkg, index);
  const part = ref.partName;
  const tree = shapeTreeOf(batch.pkg, part);
  if (!tree) throw new CommandRefusal('errors.malformed');
  // Where each placeholder is now, through the layout it is on now.
  const before = new Map(
    (resolveSlide(batch.pkg, index)?.elements ?? [])
      .filter((e) => e.layer === 'slide')
      .map((e) => [e.nodeId, e.xfrm])
  );
  const animated = animatedShapeIds(xmlPartOf(batch.pkg, part)?.root);
  const slide = placeholdersOf(tree);
  const wrapped = wrappedPlaceholdersOf(tree, xmlPartOf(batch.pkg, part)?.root);
  const allPlaces = placeholdersOf(
    shapeTreeOf(batch.pkg, layoutPart),
    xmlPartOf(batch.pkg, layoutPart)?.root,
    true
  );
  const reserved = reservedPlaces(wrapped, allPlaces);
  const places = allPlaces.filter((t) => !reserved.has(t));
  const map = mapPlaceholders(slide, places);
  const refit: string[] = [];

  // A wrapped placeholder keeps its `p:ph` and its box, in every branch.
  for (const w of ref.layoutPart === layoutPart ? [] : wrapped) {
    const box = before.get(w.read.el.id);
    for (const c of w.copies) keepBox(batch, part, c.el, box);
  }

  // ECMA-376 §19.3.1.36: `idx` names one placeholder. A placeholder that stays without a place
  // takes a new `idx` when a placeholder of the new layout uses its own: the smallest one that
  // no placeholder of the slide or the layout uses (`ST_PlaceholderIndex`, xsd:unsignedInt).
  const placeIdx = new Set(
    [...allPlaces.map((t) => t.idx), ...wrapped.map((w) => w.read.idx)].filter(
      (i): i is string => i !== undefined
    )
  );
  const usedIdx = new Set(
    [...allPlaces, ...slide, ...wrapped.flatMap((w) => w.copies)]
      .map((p) => Number(p.idx))
      .filter((n) => Number.isInteger(n))
  );
  // The loop ends after at most `usedIdx.size + 1` steps.
  const freeIdx = (): string | undefined => {
    for (let i = 1; i <= 0xffff_ffff; i += 1)
      if (!usedIdx.has(i)) {
        usedIdx.add(i);
        return String(i);
      }
    return undefined;
  };
  // A wrapped placeholder that holds no place takes a new `idx` too, in every branch.
  const layoutIdx = new Set(allPlaces.map((t) => t.idx));
  layoutIdx.delete(undefined);
  const holding = new Set(reserved.values());
  for (const w of wrapped) {
    if (holding.has(w) || !layoutIdx.has(w.read.idx)) continue;
    const free = freeIdx();
    if (free !== undefined) for (const c of w.copies) setAttr(batch, part, c.ph.id, 'idx', free);
  }

  for (const s of slide) {
    const t = map.get(s);
    const spPr = child(s.el, 'p:spPr');
    if (!t) {
      const id = shapeIdOf(s.el);
      if (contentOf(s) === 'empty' && !(id !== undefined && animated.has(id))) {
        removeShapes(batch, part, [s.el.id]);
        continue;
      }
      // Content with no place (or an animated placeholder) stays where it is: its box becomes
      // its own.
      keepBox(batch, part, s.el, before.get(s.el.id));
      const free = s.idx !== undefined && placeIdx.has(s.idx) ? freeIdx() : undefined;
      if (free !== undefined) setAttr(batch, part, s.ph.id, 'idx', free);
      continue;
    }
    // ECMA-376 §19.3.1.36: the slide placeholder takes the place's identity.
    for (const name of ['type', 'idx', 'orient', 'sz']) {
      const want = attr(t.ph, name) ?? null;
      if ((attr(s.ph, name) ?? null) !== want) setAttr(batch, part, s.ph.id, name, want);
    }
    if (s.el.kind !== 'sp') continue;
    if (reset && spPr)
      for (const c of elements(spPr)) batch.push({ op: 'removeNode', part, node: c.id });
    else {
      const xfrm = child(spPr, 'a:xfrm');
      if (xfrm) batch.push({ op: 'removeNode', part, node: xfrm.id });
    }
    const txBody = child(s.el, 'p:txBody');
    if (reset && txBody) clearTextFormatting(batch, part, txBody);
    if (hasText(txBody)) refit.push(s.el.id);
  }

  // A place no placeholder took gets an empty placeholder, after the one before it in layout
  // order, so placeholders keep the layout's stacking order.
  // In a collaboration session new ids come from the actor's stripe (`shapeIds`).
  const nextId = shapeIds(batch, tree);
  const taken = new Map([...map].map(([s, t]) => [t, s.el]));
  for (const [t, w] of reserved) taken.set(t, w.wrapper);
  let previous: Element | undefined;
  for (const t of allPlaces) {
    const s = taken.get(t);
    if (s) {
      previous = s;
      continue;
    }
    if (FOOTER_PLACEHOLDERS.has(t.type) || t.el.kind !== 'sp') continue;
    const now = shapeTreeOf(batch.pkg, part)!;
    const kids = now.children;
    const firstPlaceholder = elements(now).find((el) => phOf(el));
    const at = previous
      ? kids.findIndex((c) => c.id === previous!.id) + 1
      : firstPlaceholder
        ? kids.findIndex((c) => c.id === firstPlaceholder.id)
        : kids.findIndex((c) => c.type === 'element' && c.local === 'grpSpPr') + 1;
    const id = nextId();
    const node = batch.insert({
      op: 'insertNode',
      part,
      parent: now.id,
      index: Math.max(0, at),
      spec: specFromXml(emptyPlaceholderXml(t.el, id)),
    });
    previous = batch.element(part, node);
  }

  // Re-point the layout relationship (§19.3.1.38 `sld`: a slide has one slideLayout target).
  if (ref.layoutPart !== layoutPart) {
    const rels = xmlPartOf(batch.pkg, relsPartName(part));
    const rel = elements(rels?.root).find(
      (r) => attr(r, 'Type') === REL.slideLayout && attr(r, 'TargetMode') !== 'External'
    );
    if (!rel) throw new CommandRefusal('errors.malformed');
    setAttr(batch, relsPartName(part), rel.id, 'Target', relativeTarget(part, layoutPart));
  }
  return refit;
}
