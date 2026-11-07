// Formatted text on the clipboard. Copy takes the selected text as DrawingML paragraphs with
// their own run formatting (`copyTextParagraphs`), and as the formatting the slide shows
// (`richFromSelection`, for HTML). The `pasteText` command puts paragraphs in at the caret, or
// in a new text box when no text is selected. Pasted paragraphs are untrusted input: they are
// checked and cut down to what text formatting needs (`cleanParagraphs`), so a paste never
// brings in a relationship, a link, or a picture.

import type {
  ClipboardNode,
  EditorCommand,
  EditorSelection,
  TextPoint,
} from '../contracts/index.ts';
import type { LayoutItem } from '../layout/index.ts';
import type { RunProps } from '../cascade/index.ts';
import { child, NS, specOf, type Element, type NodeSpec } from '../store/index.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import { insertTextBox } from '../automation/commands/shape-commands.ts';
import {
  comparePoints,
  deleteRange,
  paragraphLength,
  paragraphsOf,
  slideNumberOf,
  slotsOf,
  splitParagraph,
} from '../automation/commands/text-model.ts';
import { cleanText } from '../automation/commands/text-units.ts';
import { sanitizeNode } from './clipboard.ts';
import type { RichParagraph, RichRun } from './clipboard-rich.ts';
import type { CommandContext, CommandEffect } from './commands.ts';
import { selectionBody } from './table-cells.ts';
import { refitShape, slidePart } from './text-targets.ts';

/** The clipboard type for formatted text from this editor. */
export const TEXT_MIME = 'application/x-pptx-editor-text+json';

/** The most paragraphs one paste puts in. */
const MAX_PARAGRAPHS = 5_000;
const MAX_NODES = 200_000;

type ShapeItem = Extract<LayoutItem, { kind: 'shape' }>;
type TextSelection = Extract<EditorSelection, { kind: 'text' }>;
type PasteText = Extract<EditorCommand, { type: 'pasteText' }>;

const ordered = (sel: TextSelection): [TextPoint, TextPoint] =>
  comparePoints(sel.anchor, sel.focus) <= 0 ? [sel.anchor, sel.focus] : [sel.focus, sel.anchor];

/** Elements of a paragraph a paste keeps, and the ones in its properties it drops. */
const INLINE = new Set(['pPr', 'r', 'br', 'endParaRPr']);
/** Elements that name a relationship or carry another part's content. */
const DROPPED = new Set([
  'hlinkClick',
  'hlinkMouseOver',
  'hlinkHover',
  'blipFill',
  'buBlip',
  'extLst',
]);

/** Text that XML 1.0 can hold: a control character, such as a NUL at the end of HTML, goes. */
const xmlText = (text: string): string => cleanText(text).replace(/[\n\u000B]/g, ' ');

/**
 * A checked node without links, relationships, or pictures: only DrawingML elements and
 * unqualified attributes, and text only in `a:t`.
 */
function stripped(spec: NodeSpec): NodeSpec {
  return {
    ...spec,
    prefix: 'a',
    attrs: (spec.attrs ?? [])
      .filter((x) => x.ns === '')
      .map((x) => ({ ...x, value: xmlText(x.value) })),
    nsDecls: [],
    children: (spec.children ?? []).flatMap((c): (NodeSpec | { readonly text: string })[] =>
      'text' in c
        ? spec.local === 't'
          ? [{ text: xmlText(c.text) }]
          : []
        : c.ns !== NS.a || DROPPED.has(c.local)
          ? []
          : [stripped(c)]
    ),
  };
}

/** The children each inline element of a paragraph may hold. */
const INLINE_CHILDREN: Readonly<Record<string, ReadonlySet<string>>> = {
  r: new Set(['rPr', 't']),
  fld: new Set(['rPr', 't']),
  br: new Set(['rPr']),
};

/** An inline element with only the children it may hold. */
function inline(spec: NodeSpec): NodeSpec {
  const clean = stripped(spec);
  const allowed = INLINE_CHILDREN[spec.local];
  if (!allowed) return clean;
  return {
    ...clean,
    children: (clean.children ?? []).filter((c) => !('text' in c) && allowed.has(c.local)),
  };
}

/**
 * Check untrusted pasted paragraphs: `a:p` elements holding paragraph properties, runs, and
 * breaks. A field becomes a run with the text it showed. Null when they are not paragraphs.
 */
export function cleanParagraphs(raw: unknown): NodeSpec[] | null {
  if (!Array.isArray(raw) || raw.length === 0 || raw.length > MAX_PARAGRAPHS) return null;
  const budget = { n: MAX_NODES };
  const out: NodeSpec[] = [];
  for (const item of raw) {
    const spec = sanitizeNode(item, 0, budget);
    if (!spec || spec.ns !== NS.a || spec.local !== 'p') return null;
    const kids = (spec.children ?? []).flatMap((c): NodeSpec[] => {
      if ('text' in c || c.ns !== NS.a) return [];
      if (c.local === 'fld') return [{ ...inline(c), local: 'r', attrs: [] }];
      if (c.local === 'pPr' || c.local === 'endParaRPr') return [stripped(c)];
      return INLINE.has(c.local) ? [inline(c)] : [];
    });
    out.push({ ...spec, attrs: [], nsDecls: [], children: kids });
  }
  return out;
}

/**
 * The selected text as DrawingML paragraphs, with each run's own formatting (`a:rPr`) and each
 * paragraph's properties. A field becomes a run with the text it shows.
 */
export function copyTextParagraphs(
  batch: OpBatch,
  part: string,
  sel: TextSelection
): ClipboardNode[] {
  const body = selectionBody(batch, part, sel);
  if (!body) return [];
  const [from, to] = ordered(sel);
  const n = slideNumberOf(batch, part);
  const out: NodeSpec[] = [];
  paragraphsOf(body).forEach((p, i) => {
    if (i < from.paragraph || i > to.paragraph) return;
    const lo = i === from.paragraph ? from.offset : 0;
    const hi = i === to.paragraph ? to.offset : Infinity;
    const kids: NodeSpec[] = [];
    const pPr = child(p, 'a:pPr');
    if (pPr) kids.push(specOf(pPr));
    for (const s of slotsOf(p, n)) {
      if (s.end <= lo || s.start >= hi) continue;
      if (s.kind === 'break') {
        kids.push(specOf(s.el));
        continue;
      }
      const text = s.text.slice(Math.max(0, lo - s.start), Math.min(s.text.length, hi - s.start));
      if (!text) continue;
      const rPr = child(s.el, 'a:rPr');
      kids.push({
        ns: NS.a,
        local: 'r',
        prefix: 'a',
        attrs: [],
        children: [
          ...(rPr ? [specOf(rPr)] : []),
          { ns: NS.a, local: 't', prefix: 'a', attrs: [], children: [{ text }] },
        ],
      });
    }
    const end = child(p, 'a:endParaRPr');
    if (end) kids.push(specOf(end));
    out.push({ ns: NS.a, local: 'p', prefix: 'a', attrs: [], children: kids });
  });
  return cleanParagraphs(out) ?? [];
}

const hex = (c: { r: number; g: number; b: number }): string =>
  [c.r, c.g, c.b]
    .map((v) =>
      Math.round(Math.max(0, Math.min(255, v)))
        .toString(16)
        .padStart(2, '0')
    )
    .join('')
    .toUpperCase();

/** A run as the slide shows it. */
function richRun(text: string, props: RunProps): RichRun {
  return {
    text,
    size: props.size / 100,
    ...(props.bold ? { bold: true } : {}),
    ...(props.italic ? { italic: true } : {}),
    ...(props.underline && props.underline !== 'none' ? { underline: true } : {}),
    ...(props.strike && props.strike !== 'noStrike' ? { strike: true } : {}),
    ...(props.baseline ? { baseline: props.baseline > 0 ? 30000 : -25000 } : {}),
    ...(props.fill.type === 'solid' ? { color: hex(props.fill.color) } : {}),
    ...(props.latin ? { font: props.latin } : {}),
  };
}

/** The selected text with the formatting the slide shows, for HTML. */
export function richFromSelection(item: ShapeItem, sel: TextSelection): RichParagraph[] {
  const text = item.element.text;
  if (!text) return [];
  const [from, to] = ordered(sel);
  const out: RichParagraph[] = [];
  text.paragraphs.forEach((p, i) => {
    if (i < from.paragraph || i > to.paragraph) return;
    const lo = i === from.paragraph ? from.offset : 0;
    const hi = i === to.paragraph ? to.offset : Infinity;
    const runs: RichRun[] = [];
    let at = 0;
    for (const r of p.runs) {
      const len = r.kind === 'break' ? 1 : r.text.length;
      const start = at;
      at += len;
      if (at <= lo || start >= hi) continue;
      if (r.kind === 'break') runs.push({ text: '', br: true });
      else {
        const piece = r.text.slice(Math.max(0, lo - start), Math.min(len, hi - start));
        if (piece) runs.push(richRun(piece, r.props));
      }
    }
    const b = p.bullet;
    out.push({
      runs,
      ...(p.level ? { level: p.level } : {}),
      ...(p.align === 'ctr' || p.align === 'r' || p.align === 'just' ? { align: p.align } : {}),
      ...(b?.type === 'char'
        ? { bullet: { kind: 'char', char: b.char, ...(b.font ? { font: b.font } : {}) } }
        : b?.type === 'autonum'
          ? { bullet: { kind: 'number', scheme: b.scheme, startAt: b.startAt } }
          : {}),
    });
  });
  return out;
}

/** The index of `a:endParaRPr` in a paragraph, or its end. */
const inlineEnd = (p: Element): number => {
  const i = p.children.findIndex((c) => c.type === 'element' && c.local === 'endParaRPr');
  return i < 0 ? p.children.length : i;
};

/** Paragraph `spec`'s inline items: its runs and breaks. */
const inlineOf = (spec: NodeSpec): NodeSpec[] =>
  (spec.children ?? []).filter(
    (c): c is NodeSpec => !('text' in c) && (c.local === 'r' || c.local === 'br')
  );

/**
 * Put paragraphs in at a text selection, replacing it. The first paragraph's runs join the
 * caret's paragraph, and the last one's runs go before the text after the caret. A paragraph in
 * between keeps its own properties, or takes the caret paragraph's. Returns the caret after the
 * pasted text.
 */
export function pasteParagraphs(
  batch: OpBatch,
  part: string,
  txBodyId: string,
  sel: TextSelection,
  specs: readonly NodeSpec[]
): TextPoint {
  let at = sel.focus;
  if (comparePoints(sel.anchor, sel.focus) !== 0)
    at = deleteRange(batch, part, txBodyId, sel.anchor, sel.focus);
  const n = slideNumberOf(batch, part);
  // The text after the caret moves to a paragraph of its own; it joins the last pasted one.
  splitParagraph(batch, part, txBodyId, at);
  const split = paragraphsOf(batch.element(part, txBodyId));
  const first = split[at.paragraph]!;
  const firstPPr = child(first, 'a:pPr');
  // A paste into an empty paragraph takes the first pasted paragraph's properties, and so does
  // one at the start of a paragraph when more paragraphs follow. One pasted paragraph that joins
  // the text after the caret keeps that text's paragraph properties.
  const sourcePPr = (specs[0]?.children ?? []).find(
    (c): c is NodeSpec => !('text' in c) && c.local === 'pPr'
  );
  const after = split[at.paragraph + 1];
  const tailEmpty = !after || paragraphLength(after, n) === 0;
  if (paragraphLength(first, n) === 0 && (tailEmpty || specs.length > 1) && sourcePPr) {
    if (firstPPr) batch.push({ op: 'removeNode', part, node: firstPPr.id });
    batch.insert({ op: 'insertNode', part, parent: first.id, index: 0, spec: sourcePPr });
  }
  let lastId = first.id;
  specs.forEach((spec, i) => {
    if (i > 0) {
      const body = batch.element(part, txBodyId);
      const index = body.children.findIndex((c) => c.id === lastId) + 1;
      const own = (spec.children ?? []).find(
        (c): c is NodeSpec => !('text' in c) && c.local === 'pPr'
      );
      const pPr = own ?? (firstPPr ? specOf(firstPPr) : undefined);
      const end = (spec.children ?? []).find(
        (c): c is NodeSpec => !('text' in c) && c.local === 'endParaRPr'
      );
      lastId = batch.insert({
        op: 'insertNode',
        part,
        parent: txBodyId,
        index,
        spec: { ...spec, children: [...(pPr ? [pPr] : []), ...(end ? [end] : [])] },
      });
    }
    let index = inlineEnd(batch.element(part, lastId));
    for (const item of inlineOf(spec)) {
      batch.insert({ op: 'insertNode', part, parent: lastId, index, spec: item });
      index += 1;
    }
  });
  const paragraphs = paragraphsOf(batch.element(part, txBodyId));
  const lastIndex = paragraphs.findIndex((p) => p.id === lastId);
  const caret = { paragraph: lastIndex, offset: paragraphLength(paragraphs[lastIndex]!, n) };
  // Join the text after the caret back on. That text keeps its own paragraph's properties.
  const tail = paragraphs[lastIndex + 1];
  if (lastIndex > at.paragraph && tail && paragraphLength(tail, n) > 0) {
    const own = child(paragraphs[lastIndex]!, 'a:pPr');
    if (own) batch.push({ op: 'removeNode', part, node: own.id });
    if (firstPPr)
      batch.insert({ op: 'insertNode', part, parent: lastId, index: 0, spec: specOf(firstPPr) });
  }
  deleteRange(batch, part, txBodyId, caret, { paragraph: lastIndex + 1, offset: 0 });
  return caret;
}

/** The share of the slide width a new text box of pasted text takes. */
const BOX_WIDTH = 0.6;

/**
 * Compile `pasteText`: at a text selection, the paragraphs replace it; otherwise they go in a
 * new text box in the middle of the slide, selected.
 */
export function compilePasteText(
  ctx: CommandContext,
  batch: OpBatch,
  command: PasteText
): CommandEffect {
  const specs = cleanParagraphs(command.paragraphs);
  if (!specs) throw new CommandRefusal('errors.clipboard');
  if (ctx.slideSorter) throw new CommandRefusal('disabled.slideSorter');
  const part = slidePart(ctx);
  const sel = ctx.selection;
  if (sel.kind === 'text') {
    const body = selectionBody(batch, part, sel);
    if (!body) throw new CommandRefusal('disabled.noText');
    const at = pasteParagraphs(batch, part, body.id, sel, specs);
    refitShape(batch, ctx, part, sel.nodeId);
    return { label: 'paste', selection: { ...sel, anchor: at, focus: at } };
  }
  const { width, height } = ctx.session.info;
  const id = insertTextBox(batch, part, width, height);
  // A wider box than Insert Text Box's: pasted text is often a sentence or more.
  const sp = batch.element(part, id);
  const xfrm = child(child(sp, 'p:spPr'), 'a:xfrm');
  const off = child(xfrm, 'a:off');
  const ext = child(xfrm, 'a:ext');
  if (off && ext) {
    const cx = Math.round(width * BOX_WIDTH);
    batch.push({
      op: 'setAttribute',
      part,
      node: ext.id,
      name: { ns: '', local: 'cx' },
      value: String(cx),
    });
    batch.push({
      op: 'setAttribute',
      part,
      node: off.id,
      name: { ns: '', local: 'x' },
      value: String(Math.round((width - cx) / 2)),
    });
  }
  const txBody = child(batch.element(part, id), 'p:txBody');
  if (!txBody) throw new CommandRefusal('errors.clipboard');
  const caret = { paragraph: 0, offset: 0 };
  pasteParagraphs(
    batch,
    part,
    txBody.id,
    { kind: 'text', nodeId: id, anchor: caret, focus: caret },
    specs
  );
  refitShape(batch, ctx, part, id);
  return { label: 'paste', selection: { kind: 'shapes', nodeIds: [id] } };
}
