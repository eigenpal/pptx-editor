// Text frames and text ranges.
//
// A shape's text is addressed by UTF-16 offset over its paragraphs joined with `\r`. A line
// break (`a:br`) reads as `\v` and a field as its displayed text, so offsets are the editor's
// own (`text-model.ts`: a break is one unit, a field its text length). A write splits
// paragraphs at `\r` or `\n` and writes a line break for `\v`.

import type {
  ResolvedParagraph,
  ResolvedRun,
  ResolvedTextBody,
  RunProps,
} from '../cascade/index.ts';
import { toHex } from '../cascade/index.ts';
import { attr, child, elements, intAttr, type Element } from '../store/index.ts';
import type { TextMeasurer } from '../layout/index.ts';
import { refit } from './commands/autofit.ts';
import { setRunTypeface } from './commands/font-format.ts';
import { OpBatch } from './commands/op-batch.ts';
import { refitTable } from './commands/table-fit.ts';
import { cellBodyAsSaved, joinCellBodies, readTable, tblOf } from './commands/table-model.ts';
import { setListKind } from './commands/list-commands.ts';
import {
  addExternalRelationship,
  removeUnusedHyperlink,
  xmlAttr,
} from './commands/package-edits.ts';
import { specFromXml } from './commands/shape-commands.ts';
import {
  aSpec,
  deleteRange,
  ensureParagraphProps,
  insertBreak,
  insertText,
  paragraphLength,
  paragraphsOf,
  paragraphsForFormatting,
  runsForFormatting,
  setRunAttr,
  setRunChild,
  slideNumberOf,
  slotsOf,
  splitParagraph,
} from './commands/text-model.ts';
import { isValidXmlText, REL, sanitizeHref } from '../store/index.ts';
import type {
  AutomationCellAddress,
  AutomationFontWrite,
  AutomationParagraphWrite,
  AutomationTextFrameWrite,
  AutomationTextSpan,
} from './operations.ts';
import type {
  AutomationFontRead,
  AutomationParagraphRead,
  AutomationRef,
  AutomationValue,
} from './protocol.ts';
import { resolvedShape } from './plan-reads.ts';
import {
  AutomationFailure,
  badArgument,
  edit,
  numberIn,
  ownerSlideIndex,
  shapeAt,
  toEmu,
  toPoints,
  type PlanContext,
  type ShapeLocation,
} from './targets.ts';
import type { TextPoint } from '../contracts/neutral.ts';

/** What a text plan needs besides the batch. */
export interface TextPlanContext extends PlanContext {
  /** Present when the host can re-fit text bodies by their autofit rules. */
  readonly measurer?: TextMeasurer;
  /** Set in a collaboration room: a table structure edit marks the table. */
  readonly collaborating?: boolean;
}

/** A shape's text body, its element, and its location. */
interface TextTarget {
  readonly at: ShapeLocation;
  readonly txBody: Element;
  readonly slideNumber: number | undefined;
  /** Set when the body is a table cell's `a:txBody`. */
  readonly cell?: boolean;
}

/** The longest text one write may carry. */
const MAX_TEXT = 1_000_000;

/**
 * The text body of a table cell (`a:tc/a:txBody`, ECMA-376 Part 1, 21.1.3.16): the anchor cell
 * of the block that covers the grid position. A cell can hold more than one body (a file, or
 * two peers who type at once into a cell with none): an edit joins them first
 * (`joinCellBodies`), and a read (`read`) writes nothing and reads them as the save joins them.
 */
function cellTarget(
  ctx: TextPlanContext,
  at: ShapeLocation,
  cell: AutomationCellAddress,
  read: boolean
): TextTarget {
  if (typeof cell !== 'object' || cell === null) badArgument('cell');
  if (!tblOf(at.el))
    throw new AutomationFailure('unsupported-content', 'that shape is not a table', 'not-table');
  if (!read) edit(() => joinCellBodies(ctx.batch, at.part, at.el.id));
  const m = edit(() => readTable(ctx.batch, at.part, at.el.id));
  const row = numberIn(cell.row, 'row', 0, m.nR - 1);
  const column = numberIn(cell.column, 'column', 0, m.nC - 1);
  if (!Number.isInteger(row) || !Number.isInteger(column)) badArgument('cell');
  const anchor = m.anchors[row]![column]!;
  const txBody = cellBodyAsSaved(m.tbl, m.tcs[anchor.row]![anchor.column]!);
  if (!txBody)
    throw new AutomationFailure('unsupported-content', 'a cell with no text body', 'no-text');
  return { at, txBody, slideNumber: slideNumberOf(ctx.batch, at.part), cell: true };
}

/**
 * The text body of a shape; one is added to a `p:sp` that has none when `create` is set. `read`
 * marks a read, which writes nothing (`cellTarget`).
 */
function textTarget(
  ctx: TextPlanContext,
  ref: AutomationRef,
  create = false,
  cell?: AutomationCellAddress,
  read = false
): TextTarget {
  const at = shapeAt(ctx, ref);
  if (cell !== undefined) return cellTarget(ctx, at, cell, read);
  let txBody = child(at.el, 'p:txBody');
  if (!txBody && create && at.el.kind === 'sp') {
    const id = edit(() => addTextBody(ctx.batch, at));
    txBody = ctx.batch.element(at.part, id);
  }
  if (!txBody)
    throw new AutomationFailure('unsupported-content', 'that shape has no text frame', 'no-text');
  return { at, txBody, slideNumber: slideNumberOf(ctx.batch, at.part) };
}

/**
 * The text body a `p:sp` without one gets: the body a read describes and a write adds. Plain
 * body properties, so the body, list styles, and run properties come from the cascade
 * (placeholder, layout, master, `p:style`) as for any other body.
 */
const NEW_TEXT_BODY = '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>';

/** Add {@link NEW_TEXT_BODY} to a `p:sp`; the id of the new `p:txBody`. */
function addTextBody(batch: TextPlanContext['batch'], at: ShapeLocation): string {
  // CT_Shape (ECMA-376 Part 1, 19.3.1.43): nvSpPr, spPr, style, txBody, extLst.
  const ext = at.el.children.findIndex((c) => c.type === 'element' && c.local === 'extLst');
  return batch.insert({
    op: 'insertNode',
    part: at.part,
    parent: at.el.id,
    index: ext < 0 ? at.el.children.length : ext,
    spec: specFromXml(NEW_TEXT_BODY),
  });
}

/**
 * The text body a read looks at, with the context to resolve it in. A `p:sp` without
 * `p:txBody` can hold text (CT_Shape, ECMA-376 Part 1, 19.3.1.43): it reads as the body a write
 * would add, resolved through the cascade in a copy of the package that has it, so a
 * placeholder reads its inherited body properties and fonts. Any other shape (a picture, a
 * connector, a graphic frame, a group) has no text frame: the read refuses, or with `orNull`
 * answers null, so a load through a collection's items skips it instead of failing the sync.
 */
function readTarget(
  ctx: TextPlanContext,
  ref: AutomationRef,
  cell: AutomationCellAddress | undefined,
  orNull: boolean | undefined
): { t: TextTarget; ctx: TextPlanContext } | null {
  const at = shapeAt(ctx, ref);
  if (cell === undefined && !child(at.el, 'p:txBody')) {
    if (at.el.kind === 'sp') {
      const batch = new OpBatch(ctx.batch.pkg, { reservedPartNames: ctx.batch.reservedPartNames });
      const id = addTextBody(batch, at);
      const virtual: TextPlanContext = { ...ctx, batch };
      return {
        t: {
          at: shapeAt(virtual, ref),
          txBody: batch.element(at.part, id),
          slideNumber: slideNumberOf(batch, at.part),
        },
        ctx: virtual,
      };
    }
    if (orNull) return null;
  }
  return { t: textTarget(ctx, ref, false, cell, true), ctx };
}

/** The text of one paragraph in the object model's form. */
function paragraphString(p: Element, n: number | undefined): string {
  return slotsOf(p, n)
    .map((s) => (s.kind === 'break' ? '\v' : s.text))
    .join('');
}

/** The whole text of a body. */
function bodyText(t: TextTarget): string {
  return paragraphsOf(t.txBody)
    .map((p) => paragraphString(p, t.slideNumber))
    .join('\r');
}

/** A global offset as a paragraph point. */
function pointAt(t: TextTarget, offset: number): TextPoint {
  const paragraphs = paragraphsOf(t.txBody);
  let base = 0;
  for (let i = 0; i < paragraphs.length; i += 1) {
    const len = paragraphLength(paragraphs[i]!, t.slideNumber);
    if (offset <= base + len) return { paragraph: i, offset: offset - base };
    base += len + 1;
  }
  badArgument('start');
}

/** A span as two points, checked against the text length. */
function rangeOf(
  t: TextTarget,
  span: AutomationTextSpan | undefined
): { from: TextPoint; to: TextPoint } | undefined {
  if (span === undefined) return undefined;
  const total = bodyText(t).length;
  const start = numberIn(span.start, 'start', 0, total);
  const length = numberIn(span.length, 'length', 0, total - start);
  if (!Number.isInteger(start) || !Number.isInteger(length)) badArgument('span');
  return { from: pointAt(t, start), to: pointAt(t, start + length) };
}

/**
 * Re-fit a slide shape after its text changed (ECMA-376 Part 1, 21.1.2.1.1 bodyPr), or the rows
 * of a slide table to their text.
 */
function refitText(ctx: TextPlanContext, at: ShapeLocation): void {
  const index = ownerSlideIndex(ctx.batch.pkg, at.owner);
  if (index === undefined || !ctx.measurer) return;
  const measurer = ctx.measurer;
  if (tblOf(at.el)) edit(() => refitTable(ctx.batch, index, at.part, at.el.id, measurer));
  else edit(() => refit(ctx.batch, index, at.part, at.el.id, measurer));
}

export function readText(
  context: TextPlanContext,
  ref: AutomationRef,
  span?: AutomationTextSpan,
  cell?: AutomationCellAddress,
  orNull?: boolean
): AutomationValue {
  const target = readTarget(context, ref, cell, orNull);
  if (!target) return { kind: 'null' };
  const { t } = target;
  const text = bodyText(t);
  if (span === undefined) return { kind: 'text', text };
  rangeOf(t, span);
  return { kind: 'text', text: text.slice(span.start, span.start + span.length) };
}

/** Write `text` at a point: paragraphs at `\r` or `\n`, line breaks at `\v`. */
function writeAt(ctx: TextPlanContext, t: TextTarget, at: TextPoint, text: string): void {
  const { part } = t.at;
  const id = t.txBody.id;
  let point = at;
  const paragraphs = text.replace(/\r\n/g, '\r').split(/[\r\n]/);
  paragraphs.forEach((piece, i) => {
    if (i > 0) point = splitParagraph(ctx.batch, part, id, point);
    piece.split('\v').forEach((segment, k) => {
      if (k > 0) point = insertBreak(ctx.batch, part, id, point);
      if (segment) point = insertText(ctx.batch, part, id, point, segment);
    });
  });
}

export function setText(
  ctx: TextPlanContext,
  ref: AutomationRef,
  span: AutomationTextSpan | undefined,
  text: string,
  cell?: AutomationCellAddress
): AutomationValue {
  if (typeof text !== 'string' || text.length > MAX_TEXT) badArgument('text');
  xmlTextCheck(text);
  const t = textTarget(ctx, ref, true, cell);
  const all = { start: 0, length: bodyText(t).length };
  const range = rangeOf(t, span ?? all)!;
  edit(() => {
    const at = deleteRange(ctx.batch, t.at.part, t.txBody.id, range.from, range.to);
    const fresh: TextTarget = { ...t, txBody: ctx.batch.element(t.at.part, t.txBody.id) };
    writeAt(ctx, fresh, at, text);
  });
  refitText(ctx, t.at);
  return { kind: 'applied' };
}

/** Replace the whole text of any text body (a shape's `p:txBody` or a cell's `a:txBody`). */
export function replaceBodyText(
  ctx: TextPlanContext,
  at: ShapeLocation,
  txBodyId: string,
  text: string
): void {
  if (typeof text !== 'string' || text.length > MAX_TEXT) badArgument('text');
  xmlTextCheck(text);
  const txBody = ctx.batch.element(at.part, txBodyId);
  const t: TextTarget = { at, txBody, slideNumber: slideNumberOf(ctx.batch, at.part) };
  const paragraphs = paragraphsOf(txBody);
  const last = paragraphs.length - 1;
  edit(() => {
    const point = deleteRange(
      ctx.batch,
      at.part,
      txBodyId,
      { paragraph: 0, offset: 0 },
      {
        paragraph: last,
        offset: paragraphLength(paragraphs[last]!, t.slideNumber),
      }
    );
    writeAt(ctx, { ...t, txBody: ctx.batch.element(at.part, txBodyId) }, point, text);
  });
}

/** The text of any text body, in the object model's form. */
export function bodyTextOf(txBody: Element, slideNumber: number | undefined): string {
  return paragraphsOf(txBody)
    .map((p) => paragraphString(p, slideNumber))
    .join('\r');
}

/** Refuse text with characters XML 1.0 cannot carry, lone surrogates included. */
function xmlTextCheck(text: string): void {
  // `\v` becomes a line break; every other control character but tab, CR, and LF is refused.
  if (!isValidXmlText(text.replace(/\v/g, ' '))) badArgument('text');
}

const ANCHORS: Readonly<Record<string, [string, boolean]>> = {
  t: ['t', false],
  ctr: ['ctr', false],
  b: ['b', false],
  tCentered: ['t', true],
  ctrCentered: ['ctr', true],
  bCentered: ['b', true],
};

export function readTextFrame(
  context: TextPlanContext,
  ref: AutomationRef,
  orNull?: boolean
): AutomationValue {
  const target = readTarget(context, ref, undefined, orNull);
  if (!target) return { kind: 'null' };
  const { t, ctx } = target;
  const resolved = resolvedShape(ctx, t.at);
  const body: ResolvedTextBody | undefined = resolved?.kind === 'shape' ? resolved.text : undefined;
  const bodyPr = child(t.txBody, 'a:bodyPr');
  const ins = (name: string, fallback: number): number =>
    toPoints(intAttr(bodyPr, name) ?? fallback);
  const anchor = body?.body.anchor ?? attr(bodyPr, 'anchor') ?? 't';
  const centered = body?.body.anchorCtr ?? attr(bodyPr, 'anchorCtr') === '1';
  const autofit = body?.body.autofit.kind;
  return {
    kind: 'textFrame',
    textFrame: {
      autoSize:
        autofit ??
        (child(bodyPr, 'a:spAutoFit')
          ? 'shape'
          : child(bodyPr, 'a:normAutofit')
            ? 'normal'
            : 'none'),
      leftMargin: body ? toPoints(body.body.lIns) : ins('lIns', 91_440),
      rightMargin: body ? toPoints(body.body.rIns) : ins('rIns', 91_440),
      topMargin: body ? toPoints(body.body.tIns) : ins('tIns', 45_720),
      bottomMargin: body ? toPoints(body.body.bIns) : ins('bIns', 45_720),
      verticalAlignment: centered ? `${anchor}Centered` : anchor,
      wordWrap: (body?.body.wrap ?? attr(bodyPr, 'wrap') ?? 'square') !== 'none',
      hasText: bodyText(t).length > 0,
    },
  };
}

/** The `a:bodyPr` autofit element of each `autoSize` value (EG_TextAutofit, 21.1.2.1.1). */
const AUTOSIZES: Readonly<Record<string, string>> = Object.assign(Object.create(null), {
  none: 'noAutofit',
  normal: 'normAutofit',
  shape: 'spAutoFit',
});

/** The autofit children of `a:bodyPr` (EG_TextAutofit). */
const AUTOFITS = ['noAutofit', 'normAutofit', 'spAutoFit'];
/** `a:bodyPr` children that come before the autofit choice. */
const BEFORE_AUTOFIT = ['prstTxWarp'];

export function setTextFrame(
  ctx: TextPlanContext,
  ref: AutomationRef,
  props: AutomationTextFrameWrite
): AutomationValue {
  if (typeof props !== 'object' || props === null) badArgument('props');
  const t = textTarget(ctx, ref, true);
  const { part } = t.at;
  let bodyPr = child(t.txBody, 'a:bodyPr');
  if (!bodyPr) {
    const id = ctx.batch.insert({
      op: 'insertNode',
      part,
      parent: t.txBody.id,
      index: 0,
      spec: aSpec('bodyPr'),
    });
    bodyPr = ctx.batch.element(part, id);
  }
  const bodyPrId = bodyPr.id;
  const set = (name: string, value: string | null): void =>
    ctx.batch.push({
      op: 'setAttribute',
      part,
      node: bodyPrId,
      name: { ns: '', local: name },
      value,
    });
  const margin = (name: string, value: number | undefined, key: string): void => {
    if (value === undefined) return;
    set(name, String(toEmu(numberIn(value, key, 0, 4032))));
  };
  margin('lIns', props.leftMargin, 'leftMargin');
  margin('rIns', props.rightMargin, 'rightMargin');
  margin('tIns', props.topMargin, 'topMargin');
  margin('bIns', props.bottomMargin, 'bottomMargin');
  if (props.verticalAlignment !== undefined) {
    const anchor = Object.prototype.hasOwnProperty.call(ANCHORS, props.verticalAlignment)
      ? ANCHORS[props.verticalAlignment]
      : badArgument('verticalAlignment');
    set('anchor', anchor![0]);
    set('anchorCtr', anchor![1] ? '1' : null);
  }
  if (props.wordWrap !== undefined) {
    if (typeof props.wordWrap !== 'boolean') badArgument('wordWrap');
    set('wrap', props.wordWrap ? 'square' : 'none');
  }
  if (props.autoSize !== undefined) {
    const local = Object.hasOwn(AUTOSIZES, props.autoSize)
      ? AUTOSIZES[props.autoSize]
      : badArgument('autoSize');
    for (const c of elements(ctx.batch.element(part, bodyPrId)))
      if (AUTOFITS.includes(c.local)) ctx.batch.push({ op: 'removeNode', part, node: c.id });
    let index = 0;
    ctx.batch.element(part, bodyPrId).children.forEach((c, i) => {
      if (c.type === 'element' && BEFORE_AUTOFIT.includes(c.local)) index = i + 1;
    });
    ctx.batch.insert({ op: 'insertNode', part, parent: bodyPrId, index, spec: aSpec(local) });
  }
  refitText(ctx, t.at);
  return { kind: 'applied' };
}

/** The resolved paragraphs of a body, by `a:p` node id. */
function resolvedParagraphsOf(ctx: TextPlanContext, t: TextTarget): Map<string, ResolvedParagraph> {
  const resolved = resolvedShape(ctx, t.at);
  const out = new Map<string, ResolvedParagraph>();
  if (resolved?.kind === 'shape' && resolved.text)
    for (const p of resolved.text.paragraphs) out.set(p.nodeId, p);
  if (resolved?.kind === 'graphicFrame' && resolved.table)
    for (const c of resolved.table.cells) for (const p of c.text.paragraphs) out.set(p.nodeId, p);
  return out;
}

/**
 * The format that text typed at a point takes: the run that ends at or holds the point, or at
 * the start of a paragraph the run that starts there, or the end properties of a paragraph with
 * no text.
 */
function caretProps(t: TextTarget, p: Element, resolved: ResolvedParagraph, at: number): RunProps {
  const runs = new Map<string, ResolvedRun>(resolved.runs.map((r) => [r.nodeId, r]));
  const text = slotsOf(p, t.slideNumber).filter((s) => s.kind !== 'break' && runs.has(s.el.id));
  const hit = text.find((s) => s.start < at && at <= s.end) ?? text.find((s) => s.start === at);
  return hit ? runs.get(hit.el.id)!.props : resolved.endProps;
}

/** Each paragraph's runs in the range, with their resolved properties. */
function runsIn(
  ctx: TextPlanContext,
  t: TextTarget,
  range: { from: TextPoint; to: TextPoint } | undefined
): RunProps[] {
  const byId = resolvedParagraphsOf(ctx, t);
  const paragraphs = paragraphsOf(t.txBody);
  if (
    range &&
    range.from.paragraph === range.to.paragraph &&
    range.from.offset === range.to.offset
  ) {
    // An empty range reads the format that text typed there takes.
    const p = paragraphs[range.from.paragraph];
    const resolved = p ? byId.get(p.id) : undefined;
    return p && resolved ? [caretProps(t, p, resolved, range.from.offset)] : [];
  }
  const out: RunProps[] = [];
  paragraphs.forEach((p, i) => {
    const resolved = byId.get(p.id);
    if (!resolved) return;
    if (range && (i < range.from.paragraph || i > range.to.paragraph)) return;
    // A range that ends at the start of a later paragraph holds none of its text.
    if (range && i === range.to.paragraph && i > range.from.paragraph && range.to.offset === 0)
      return;
    const lo = range && i === range.from.paragraph ? range.from.offset : 0;
    const hi = range && i === range.to.paragraph ? range.to.offset : Infinity;
    const slots = slotsOf(p, t.slideNumber);
    const runs = new Map<string, ResolvedRun>(resolved.runs.map((r) => [r.nodeId, r]));
    let any = false;
    for (const s of slots) {
      const run = runs.get(s.el.id);
      if (!run || s.kind === 'break') continue;
      if (s.start < hi && s.end > lo) {
        out.push(run.props);
        any = true;
      }
    }
    if (!any && (!range || slots.length === 0)) out.push(resolved.endProps);
  });
  if (out.length === 0 && range) {
    // Only a paragraph break: the format of the end of the paragraph it ends.
    const p = paragraphs[range.from.paragraph];
    const resolved = p ? byId.get(p.id) : undefined;
    if (p && resolved) out.push(caretProps(t, p, resolved, range.from.offset));
  }
  return out;
}

/** The value all items share, or null when they differ or there are none. */
function shared<T>(items: readonly T[]): T | null {
  if (items.length === 0) return null;
  const first = items[0]!;
  return items.every((v) => v === first) ? first : null;
}

/** One run's character format as a read. */
function fontOf(p: RunProps): AutomationFontRead {
  return {
    name: p.latin,
    size: p.size / 100,
    bold: p.bold,
    italic: p.italic,
    underline: p.underline || 'none',
    color: p.fill.type === 'solid' ? toHex({ ...p.fill.color, a: 1 }).toUpperCase() : null,
    strikethrough: p.strike === 'sngStrike',
    doubleStrikethrough: p.strike === 'dblStrike',
    allCaps: p.caps === 'all',
    smallCaps: p.caps === 'small',
    superscript: p.baseline > 0,
    subscript: p.baseline < 0,
  };
}

const FONT_FIELDS: readonly (keyof AutomationFontRead)[] = [
  'name',
  'size',
  'bold',
  'italic',
  'underline',
  'color',
  'strikethrough',
  'doubleStrikethrough',
  'allCaps',
  'smallCaps',
  'superscript',
  'subscript',
];

/** A format that nothing resolved: every field unknown. */
const UNKNOWN_FONT = Object.fromEntries(
  FONT_FIELDS.map((field) => [field, null])
) as unknown as AutomationFontRead;

export function readFont(
  context: TextPlanContext,
  ref: AutomationRef,
  span?: AutomationTextSpan,
  cell?: AutomationCellAddress,
  orNull?: boolean
): AutomationValue {
  const target = readTarget(context, ref, cell, orNull);
  if (!target) return { kind: 'null' };
  const { t, ctx } = target;
  const fonts = runsIn(ctx, t, rangeOf(t, span)).map(fontOf);
  const font: Record<string, unknown> = {};
  for (const field of FONT_FIELDS) font[field] = shared(fonts.map((f) => f[field]));
  return { kind: 'font', font: font as unknown as AutomationFontRead };
}

/**
 * The runs of a text body in order, each with its effective format. Neighbours with the same
 * format read as one run. A line break reads as `\v`, and a paragraph break as `\r` at the end
 * of the run before it, so the run texts join to the body text.
 */
export function readTextRuns(
  ctx: TextPlanContext,
  ref: AutomationRef,
  cell?: AutomationCellAddress
): AutomationValue {
  let t: TextTarget;
  try {
    t = textTarget(ctx, ref, false, cell, true);
  } catch (error) {
    // A table cell may have no text body (CT_TableCell, ECMA-376 Part 1, 21.1.3.16): no runs.
    if (cell !== undefined && error instanceof AutomationFailure && error.detail === 'no-text')
      return { kind: 'textRuns', runs: [] };
    throw error;
  }
  const byId = resolvedParagraphsOf(ctx, t);
  const runs: { text: string; font: AutomationFontRead }[] = [];
  const add = (text: string, props: RunProps | undefined): void => {
    const last = runs[runs.length - 1];
    // A slot the cascade did not resolve keeps the format of the run before it.
    const font = props ? fontOf(props) : (last?.font ?? UNKNOWN_FONT);
    if (last && FONT_FIELDS.every((field) => last.font[field] === font[field])) last.text += text;
    else runs.push({ text, font });
  };
  const paragraphs = paragraphsOf(t.txBody);
  paragraphs.forEach((p, i) => {
    const resolved = byId.get(p.id);
    const props = new Map<string, RunProps>(resolved?.runs.map((r) => [r.nodeId, r.props]));
    for (const s of slotsOf(p, t.slideNumber)) {
      const text = s.kind === 'break' ? '\v' : s.text;
      if (text.length > 0) add(text, props.get(s.el.id) ?? resolved?.endProps);
    }
    if (i === paragraphs.length - 1) return;
    const last = runs[runs.length - 1];
    if (last) last.text += '\r';
    else add('\r', resolved?.endProps);
  });
  return { kind: 'textRuns', runs };
}

/** `ST_TextUnderlineType` (ECMA-376 Part 1, 20.1.10.82). */
const UNDERLINES = new Set([
  'none',
  'words',
  'sng',
  'dbl',
  'heavy',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dotDashHeavy',
  'dotDotDash',
  'dotDotDashHeavy',
  'wavy',
  'wavyHeavy',
  'wavyDbl',
]);

/** The fill children of `a:rPr` (EG_FillProperties). */
const FILLS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];

/** A color write as a DrawingML color element. */
export function colorSpec(value: unknown, name: string, transparency?: number): string {
  if (typeof value !== 'string') badArgument(name);
  const hex = /^#?([0-9a-fA-F]{6})$/.exec(value.trim());
  const alpha =
    transparency === undefined || transparency === 0
      ? ''
      : `<a:alpha val="${Math.round((1 - numberIn(transparency, 'transparency', 0, 1)) * 100_000)}"/>`;
  if (hex) return `<a:srgbClr val="${hex[1]!.toUpperCase()}">${alpha}</a:srgbClr>`;
  const preset = PRESET_NAMES.get(value.trim().toLowerCase());
  if (preset) return `<a:prstClr val="${preset}">${alpha}</a:prstClr>`;
  badArgument(name);
}

/** A few HTML color names the object model accepts, as `ST_PresetColorVal` values. */
const PRESET_NAMES = new Map(
  [
    'black',
    'white',
    'red',
    'green',
    'blue',
    'yellow',
    'orange',
    'purple',
    'gray',
    'grey',
    'silver',
    'maroon',
    'navy',
    'teal',
    'olive',
    'lime',
    'aqua',
    'fuchsia',
    'cyan',
    'magenta',
    'pink',
    'brown',
    'gold',
    'indigo',
    'violet',
  ].map((name) => [name, name === 'grey' ? 'gray' : name])
);

export function setFont(
  ctx: TextPlanContext,
  ref: AutomationRef,
  span: AutomationTextSpan | undefined,
  font: AutomationFontWrite,
  cell?: AutomationCellAddress
): AutomationValue {
  if (typeof font !== 'object' || font === null) badArgument('font');
  const t = textTarget(ctx, ref, false, cell);
  const range = rangeOf(t, span);
  const { part } = t.at;
  edit(() => {
    const runs = runsForFormatting(ctx.batch, part, t.txBody.id, range, false);
    for (const run of runs) {
      const set = (name: string, value: string | null): void =>
        setRunAttr(ctx.batch, part, run, name, value);
      const flag = (value: unknown, name: string): boolean => {
        if (typeof value !== 'boolean') badArgument(name);
        return value;
      };
      if (font.name !== undefined) {
        if (typeof font.name !== 'string' || !/^[\p{L}\p{N}\p{M} \-.+_]{1,64}$/u.test(font.name))
          badArgument('name');
        setRunTypeface(ctx.batch, part, run, font.name.trim());
      }
      if (font.size !== undefined)
        set('sz', String(Math.round(numberIn(font.size, 'size', 1, 4000) * 100)));
      if (font.bold !== undefined) set('b', flag(font.bold, 'bold') ? '1' : '0');
      if (font.italic !== undefined) set('i', flag(font.italic, 'italic') ? '1' : '0');
      if (font.underline !== undefined) {
        if (!UNDERLINES.has(font.underline)) badArgument('underline');
        set('u', font.underline);
      }
      if (font.color !== undefined)
        setRunChild(
          ctx.batch,
          part,
          run,
          FILLS,
          specFromXml(`<a:solidFill>${colorSpec(font.color, 'color')}</a:solidFill>`)
        );
      if (font.strikethrough !== undefined)
        set('strike', flag(font.strikethrough, 'strikethrough') ? 'sngStrike' : 'noStrike');
      if (font.doubleStrikethrough !== undefined)
        set(
          'strike',
          flag(font.doubleStrikethrough, 'doubleStrikethrough') ? 'dblStrike' : 'noStrike'
        );
      if (font.allCaps !== undefined) set('cap', flag(font.allCaps, 'allCaps') ? 'all' : 'none');
      if (font.smallCaps !== undefined)
        set('cap', flag(font.smallCaps, 'smallCaps') ? 'small' : 'none');
      // 30 % and -25 %: the superscript and subscript offsets the editor writes.
      if (font.superscript !== undefined)
        set('baseline', flag(font.superscript, 'superscript') ? '30000' : '0');
      if (font.subscript !== undefined)
        set('baseline', flag(font.subscript, 'subscript') ? '-25000' : '0');
    }
  });
  refitText(ctx, t.at);
  return { kind: 'applied' };
}

export function readParagraph(
  context: TextPlanContext,
  ref: AutomationRef,
  span?: AutomationTextSpan,
  cell?: AutomationCellAddress,
  orNull?: boolean
): AutomationValue {
  const target = readTarget(context, ref, cell, orNull);
  if (!target) return { kind: 'null' };
  const { t, ctx } = target;
  const range = rangeOf(t, span);
  const byId = resolvedParagraphsOf(ctx, t);
  const ids = paragraphsOf(t.txBody).map((p) => p.id);
  const lo = range ? range.from.paragraph : 0;
  const hi = range ? range.to.paragraph : ids.length - 1;
  const chosen = ids.slice(lo, hi + 1).map((id) => byId.get(id));
  const files = paragraphsOf(t.txBody).slice(lo, hi + 1);
  const resolvedAll = chosen.every((p) => p !== undefined);
  const paragraph: AutomationParagraphRead = resolvedAll
    ? {
        alignment: shared(chosen.map((p) => p!.align)),
        bulletVisible: shared(chosen.map((p) => p!.bullet !== null)),
        bulletType: shared(chosen.map((p) => p!.bullet?.type ?? 'none')),
        bulletStyle: shared(
          chosen.map((p) => (p!.bullet?.type === 'autonum' ? p!.bullet.scheme : null))
        ),
        indentLevel: shared(chosen.map((p) => p!.level + 1)),
      }
    : {
        alignment: shared(files.map((p) => attr(child(p, 'a:pPr'), 'algn') ?? 'l')),
        bulletVisible: null,
        bulletType: null,
        bulletStyle: null,
        indentLevel: shared(files.map((p) => (intAttr(child(p, 'a:pPr'), 'lvl') ?? 0) + 1)),
      };
  return { kind: 'paragraph', paragraph };
}

/** `ST_TextAlignType` (ECMA-376 Part 1, 20.1.10.59). */
const ALIGNMENTS = new Set(['l', 'ctr', 'r', 'just', 'justLow', 'dist', 'thaiDist']);

export function setParagraph(
  ctx: TextPlanContext,
  ref: AutomationRef,
  span: AutomationTextSpan | undefined,
  format: AutomationParagraphWrite,
  cell?: AutomationCellAddress
): AutomationValue {
  if (typeof format !== 'object' || format === null) badArgument('paragraph');
  const t = textTarget(ctx, ref, false, cell);
  const range = rangeOf(t, span);
  const { part } = t.at;
  const slide = ownerSlideIndex(ctx.batch.pkg, t.at.owner);
  edit(() => {
    const ids = paragraphsForFormatting(ctx.batch, part, t.txBody.id, range);
    const all = paragraphsOf(ctx.batch.element(part, t.txBody.id)).map((p) => p.id);
    for (const pId of ids) {
      const pPrId = ensureParagraphProps(ctx.batch, part, pId);
      const set = (name: string, value: string | null): void =>
        ctx.batch.push({
          op: 'setAttribute',
          part,
          node: pPrId,
          name: { ns: '', local: name },
          value,
        });
      if (format.alignment !== undefined) {
        if (!ALIGNMENTS.has(format.alignment)) badArgument('alignment');
        set('algn', format.alignment);
      }
      if (format.indentLevel !== undefined) {
        const level = numberIn(format.indentLevel, 'indentLevel', 1, 9);
        if (!Number.isInteger(level)) badArgument('indentLevel');
        set('lvl', level === 1 ? null : String(level - 1));
      }
    }
    const hidden = format.bulletVisible === false || format.bulletType === 'none';
    const kind = hidden
      ? 'none'
      : format.bulletType === 'autonum' || format.bulletStyle !== undefined
        ? 'number'
        : format.bulletVisible === true || format.bulletType === 'char'
          ? 'bullet'
          : undefined;
    if (kind) {
      const target = {
        shapeId: t.at.el.id,
        txBodyId: t.txBody.id,
        paragraphs: ids.map((id) => all.indexOf(id)),
      };
      setListKind(ctx.batch, part, slide ?? 0, target, kind, slide !== undefined);
    }
    if (format.bulletStyle !== undefined) {
      if (!/^[a-zA-Z]{1,32}$/.test(format.bulletStyle)) badArgument('bulletStyle');
      for (const pId of ids) {
        const pPr = child(ctx.batch.element(part, pId), 'a:pPr');
        const auto = child(pPr, 'a:buAutoNum');
        if (auto)
          ctx.batch.push({
            op: 'setAttribute',
            part,
            node: auto.id,
            name: { ns: '', local: 'type' },
            value: format.bulletStyle,
          });
      }
    }
  });
  refitText(ctx, t.at);
  return { kind: 'applied' };
}

/** The runs a hyperlink covers, as `a:hlinkClick` holders. */
export function setHyperlink(
  ctx: TextPlanContext,
  ref: AutomationRef,
  span: AutomationTextSpan | undefined,
  address: string | null,
  screenTip?: string
): AutomationValue {
  const t = textTarget(ctx, ref);
  // An empty range holds no text to link: a link on it would be an empty run.
  if (span !== undefined && span.length === 0)
    throw new AutomationFailure('invalid-argument', 'no text to link', 'span');
  const range = rangeOf(t, span);
  const { part } = t.at;
  let href: string | null = null;
  if (address !== null) {
    if (typeof address !== 'string' || address.length > 2048) badArgument('address');
    const projected = sanitizeHref(address);
    if (!projected.ok) badArgument('address');
    href = projected.href;
  }
  if (
    screenTip !== undefined &&
    (typeof screenTip !== 'string' || screenTip.length > 1024 || !isValidXmlText(screenTip))
  )
    badArgument('screenTip');
  edit(() => {
    const runs = runsForFormatting(ctx.batch, part, t.txBody.id, range, false).filter(
      (r) => r.local === 'r' || r.local === 'fld'
    );
    if (runs.length === 0)
      throw new AutomationFailure('invalid-argument', 'no text to link', 'span');
    // The relationships the old links named go once nothing names them: another run may share one.
    const old = new Set<string>();
    for (const run of runs) {
      const rId = attr(child(child(run, 'a:rPr'), 'a:hlinkClick'), 'r:id');
      if (rId) old.add(rId);
    }
    const rId = href ? addExternalRelationship(ctx.batch, part, REL.hyperlink, href) : null;
    for (const run of runs) {
      if (!rId) {
        setRunChild(ctx.batch, part, run, ['hlinkClick'], null);
        continue;
      }
      const tip = screenTip ? ` tooltip="${xmlAttr(screenTip, 'screenTip')}"` : '';
      setRunChild(
        ctx.batch,
        part,
        run,
        ['hlinkClick'],
        specFromXml(`<a:hlinkClick r:id="${xmlAttr(rId, 'rId')}"${tip}/>`)
      );
    }
    for (const id of old) removeUnusedHyperlink(ctx.batch, part, id);
  });
  return { kind: 'applied' };
}

/** Where each text link of a body sits, by global offset. */
export function linkSpans(
  txBody: Element,
  slideNumber: number | undefined
): { rId: string; tooltip: string | null; start: number; length: number }[] {
  const out: { rId: string; tooltip: string | null; start: number; length: number }[] = [];
  let base = 0;
  for (const p of paragraphsOf(txBody)) {
    for (const s of slotsOf(p, slideNumber)) {
      const link = child(child(s.el, 'a:rPr'), 'a:hlinkClick');
      const rId = attr(link, 'r:id');
      if (!rId) continue;
      const last = out[out.length - 1];
      if (last && last.rId === rId && last.start + last.length === base + s.start)
        out[out.length - 1] = { ...last, length: last.length + (s.end - s.start) };
      else
        out.push({
          rId,
          tooltip: attr(link, 'tooltip') ?? null,
          start: base + s.start,
          length: s.end - s.start,
        });
    }
    base += paragraphLength(p, slideNumber) + 1;
  }
  return out;
}
