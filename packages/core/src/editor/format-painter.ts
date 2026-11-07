// The format painter: the Format Painter button, Ctrl/Cmd+Shift+C, and Ctrl/Cmd+Shift+V.
//
// A copy holds node specs, never node ids, so it stays valid after edits. A paste is one store
// transaction, so it is one undo step. Text copies the run properties at the start of the
// selection (and paragraph alignment and spacing when whole paragraphs are selected); a shape
// copies its fill, outline, and effects, its `p:style`, and the run properties of its text.
// A copy adds the bold, italic, underline, strike, and size the source shows when it inherits
// them; an inherited font or color is not copied (the target keeps its own).
//
// The press logic follows DOCX Editor's painter (EigenPal,
// docx-editor/packages/core/src/editor/surface-format-painter.ts): one press arms the painter
// for one paint, a second press inside the double-press window keeps it on until Escape, and a
// press while it is on turns it off. The window lives here, so a host binds `click` only.

import { attr, child, elements, NS, specOf, type Element, type NodeSpec } from '../store/index.ts';
import type {
  EditorCommand,
  EditorSelection,
  FormatPainterMode,
  TextPoint,
} from '../contracts/index.ts';
import { refitAll } from '../automation/commands/autofit.ts';
import { selectedIds, type CommandContext, type CommandEffect } from './commands.ts';
import { selectionBody } from './table-cells.ts';
import { tableTextTargets } from './table-commands.ts';
import { tblOf } from '../automation/commands/table-model.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import { formatAt } from './selection-state.ts';
import {
  comparePoints,
  ensureParagraphProps,
  ensureRunProps,
  paragraphLength,
  paragraphsForFormatting,
  paragraphsOf,
  wordTextOf,
  runsForFormatting,
  setRunChild,
  slideNumberOf,
  slotsOf,
} from '../automation/commands/text-model.ts';
import { textWordSegments, wordSegmentAt } from '../automation/commands/text-units.ts';

/** Two presses closer than this are one double-press (the common double-click time). */
const DOUBLE_PRESS_MS = 500;

/** Run attributes that are not formatting: they stay on the target and are never copied. */
const RUN_STATE_ATTRS = new Set([
  'lang',
  'altLang',
  'dirty',
  'err',
  'noProof',
  'smtClean',
  'smtId',
  'bmk',
]);
/** Run children that are not formatting: links and extensions. */
const RUN_STATE_CHILDREN = new Set(['hlinkClick', 'hlinkMouseOver', 'extLst']);
/** `p:spPr` children a shape copy carries, in schema order (CT_ShapeProperties). */
const SHAPE_PROPS = [
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'ln',
  'effectLst',
  'effectDag',
  'scene3d',
  'sp3d',
];
const GEOMETRY = ['xfrm', 'custGeom', 'prstGeom'];
/** Paragraph spacing children, in schema order (CT_TextParagraphProperties). */
const SPACING = ['lnSpc', 'spcBef', 'spcAft'];

/** Formatting copied from a selection. */
export interface CopiedFormat {
  /** The slide part it came from: a picture fill names a relationship of that part only. */
  readonly part: string;
  /** Character formatting: `a:rPr` without language, proofing, or links. */
  readonly run: NodeSpec | null;
  /** Paragraph alignment and spacing, copied when whole paragraphs are selected. */
  readonly paragraph: { readonly algn: string | null; readonly spacing: NodeSpec[] } | null;
  /** Shape formatting: `p:spPr` fill, outline, and effects, and `p:style`. */
  readonly shape: { readonly spPr: NodeSpec[]; readonly style: NodeSpec | null } | null;
}

const isA = (el: Element, names: readonly string[]): boolean =>
  el.ns === NS.a && names.includes(el.local);

/** Whether a spec names a relationship (`r:embed`, `r:id`): it is valid in its own part only. */
function namesRelationship(spec: NodeSpec): boolean {
  if (spec.attrs?.some((a) => a.ns === NS.r)) return true;
  return (spec.children ?? []).some((c) => 'local' in c && namesRelationship(c));
}

/** The formatting of a run's `a:rPr` (or `a:endParaRPr`), as an `a:rPr` spec. */
function runSpec(rPr: Element | undefined, resolved: ReturnType<typeof formatAt>): NodeSpec {
  const spec = rPr ? specOf(rPr) : { ns: NS.a, local: 'rPr', attrs: [], children: [] };
  const attrs = (spec.attrs ?? []).filter((a) => a.ns === '' && !RUN_STATE_ATTRS.has(a.local));
  // The toggles and size the user sees, also when they come from the placeholder or master,
  // so the painted text looks like its source wherever it sits.
  const set = (local: string, value: string | undefined): void => {
    if (value !== undefined && !attrs.some((a) => a.local === local))
      attrs.push({ ns: '', local, value });
  };
  if (resolved) {
    set('b', resolved.bold ? '1' : '0');
    set('i', resolved.italic ? '1' : '0');
    set('u', resolved.underline ? 'sng' : 'none');
    set('strike', resolved.strike ? 'sngStrike' : 'noStrike');
    if (resolved.fontSize) set('sz', String(Math.round(resolved.fontSize * 100)));
  }
  const children = (spec.children ?? []).filter(
    (c): c is NodeSpec => 'local' in c && !RUN_STATE_CHILDREN.has(c.local) && !namesRelationship(c)
  );
  return {
    ns: NS.a,
    local: 'rPr',
    ...(spec.prefix ? { prefix: spec.prefix } : {}),
    attrs,
    children,
  };
}

/** The text range of a text selection, start first. */
function ordered(sel: Extract<EditorSelection, { kind: 'text' }>): [TextPoint, TextPoint] {
  return comparePoints(sel.anchor, sel.focus) <= 0
    ? [sel.anchor, sel.focus]
    : [sel.focus, sel.anchor];
}

/** Copy the formatting of a text selection: the character at its start. */
function copyText(
  batch: OpBatch,
  part: string,
  ctx: CommandContext,
  sel: Extract<EditorSelection, { kind: 'text' }>
): CopiedFormat {
  const txBody = selectionBody(batch, part, sel);
  if (!txBody) throw new CommandRefusal('disabled.noText');
  const [from, to] = ordered(sel);
  const p = paragraphsOf(txBody)[from.paragraph];
  if (!p) throw new CommandRefusal('disabled.noText');
  const n = slideNumberOf(batch, part);
  const collapsed = comparePoints(from, to) === 0;
  // A caret takes the character before it, as typing does; a range takes its first character.
  const at = collapsed ? Math.max(0, from.offset - 1) : from.offset;
  const slots = slotsOf(p, n).filter((s) => s.kind !== 'break');
  const slot = slots.find((s) => at >= s.start && at < s.end) ?? slots[slots.length - 1];
  const rPr = slot ? child(slot.el, 'a:rPr') : child(p, 'a:endParaRPr');
  const layout = ctx.session.layout(ctx.slide);
  const whole =
    !collapsed &&
    from.offset === 0 &&
    (to.paragraph > from.paragraph || to.offset >= paragraphLength(p, n));
  const pPr = child(p, 'a:pPr');
  return {
    part,
    run: runSpec(rPr, formatAt(layout, sel)),
    paragraph: whole
      ? {
          algn: attr(pPr, 'algn') ?? null,
          spacing: pPr
            ? elements(pPr)
                .filter((c) => isA(c, SPACING))
                .map(specOf)
            : [],
        }
      : null,
    shape: null,
  };
}

/** Copy the formatting of the first selected shape. */
function copyShape(batch: OpBatch, part: string, ctx: CommandContext, id: string): CopiedFormat {
  const el = batch.element(part, id);
  if (el.kind !== 'sp' && el.kind !== 'cxnSp' && el.kind !== 'pic')
    throw new CommandRefusal('disabled.unsupportedElement');
  const spPr = child(el, 'p:spPr');
  if (!spPr) throw new CommandRefusal('errors.malformed');
  const style = child(el, 'p:style');
  const txBody = child(el, 'p:txBody');
  const first = txBody
    ? paragraphsOf(txBody)
        .flatMap((p) => slotsOf(p, slideNumberOf(batch, part)))
        .find((s) => s.kind === 'run')
    : undefined;
  const shapes: EditorSelection = { kind: 'shapes', nodeIds: [id] };
  return {
    part,
    run: txBody
      ? runSpec(child(first?.el, 'a:rPr'), formatAt(ctx.session.layout(ctx.slide), shapes))
      : null,
    paragraph: null,
    shape: {
      spPr: elements(spPr)
        .filter((c) => isA(c, SHAPE_PROPS))
        .map(specOf),
      style: style ? specOf(style) : null,
    },
  };
}

/** Copy the formatting of the selection. Throws CommandRefusal. */
export function copyFormat(ctx: CommandContext): CopiedFormat {
  const part = ctx.session.info.slides[ctx.slide]?.partName;
  if (!part) throw new CommandRefusal('disabled.noEditor');
  const batch = ctx.session.batch();
  const sel = ctx.selection;
  if (sel.kind === 'text') return copyText(batch, part, ctx, sel);
  if (sel.kind === 'shapes' && sel.nodeIds.length)
    return copyShape(batch, part, ctx, sel.nodeIds[0]);
  throw new CommandRefusal('disabled.noSelection');
}

/** Replace the formatting of a run (or a paragraph mark) with a copied `a:rPr`. */
function paintRun(batch: OpBatch, part: string, runLike: Element, run: NodeSpec): void {
  const rPrId = ensureRunProps(batch, part, runLike);
  const rPr = batch.element(part, rPrId);
  for (const a of rPr.attrs)
    if (!RUN_STATE_ATTRS.has(a.local) && !run.attrs?.some((b) => b.local === a.local))
      batch.push({
        op: 'setAttribute',
        part,
        node: rPrId,
        name: { ns: a.ns, local: a.local },
        value: null,
      });
  for (const a of run.attrs ?? [])
    batch.push({
      op: 'setAttribute',
      part,
      node: rPrId,
      name: { ns: '', local: a.local },
      value: a.value,
    });
  for (const c of elements(batch.element(part, rPrId)))
    if (!RUN_STATE_CHILDREN.has(c.local)) batch.push({ op: 'removeNode', part, node: c.id });
  for (const c of run.children ?? [])
    if ('local' in c) setRunChild(batch, part, runLike, [c.local], c);
}

/** Paint paragraph alignment and spacing. */
function paintParagraph(
  batch: OpBatch,
  part: string,
  pId: string,
  paragraph: NonNullable<CopiedFormat['paragraph']>
): void {
  const pPrId = ensureParagraphProps(batch, part, pId);
  batch.push({
    op: 'setAttribute',
    part,
    node: pPrId,
    name: { ns: '', local: 'algn' },
    value: paragraph.algn,
  });
  for (const c of elements(batch.element(part, pPrId)))
    if (isA(c, SPACING)) batch.push({ op: 'removeNode', part, node: c.id });
  paragraph.spacing.forEach((spec, index) =>
    batch.insert({ op: 'insertNode', part, parent: pPrId, index, spec })
  );
}

/** Paint shape formatting onto a shape, a connector, or a picture. False for other elements. */
function paintShape(
  batch: OpBatch,
  part: string,
  id: string,
  copied: CopiedFormat,
  shape: NonNullable<CopiedFormat['shape']>
): boolean {
  const el = batch.element(part, id);
  if (el.kind !== 'sp' && el.kind !== 'cxnSp' && el.kind !== 'pic') return false;
  const spPr = child(el, 'p:spPr');
  if (!spPr) return false;
  for (const c of elements(spPr))
    if (isA(c, SHAPE_PROPS)) batch.push({ op: 'removeNode', part, node: c.id });
  let index = 0;
  batch.element(part, spPr.id).children.forEach((c, i) => {
    if (c.type === 'element' && isA(c, GEOMETRY)) index = i + 1;
  });
  // A picture fill names a relationship of the slide it came from; elsewhere it is left out.
  for (const spec of shape.spPr)
    if (part === copied.part || !namesRelationship(spec))
      batch.insert({ op: 'insertNode', part, parent: spPr.id, index: index++, spec });
  const old = child(batch.element(part, id), 'p:style');
  if (old) batch.push({ op: 'removeNode', part, node: old.id });
  if (shape.style) {
    const at = batch.element(part, id).children.findIndex((c) => c.id === spPr.id) + 1;
    batch.insert({ op: 'insertNode', part, parent: id, index: at, spec: shape.style });
  }
  return true;
}

/**
 * The word a caret paints: the word a double-click there selects (UAX #29, `wordSegmentAt`),
 * also when the caret touches the word at its edge. Null when the caret is not in or next to a
 * word. A field reads as object characters (`wordTextOf`), so it ends a word.
 */
function wordAt(p: Element, offset: number, n: number | undefined): [number, number] | null {
  const seg = wordSegmentAt(textWordSegments(wordTextOf(p, n)), offset);
  return seg?.word ? [seg.start, seg.end] : null;
}

/**
 * Apply copied formatting to the selection: shape formatting to the selected shapes, text
 * formatting to the selected text (a caret paints its word) or to all text of the selected
 * shapes. Returns the shapes whose text changed. Throws CommandRefusal.
 */
export function pasteFormat(
  batch: OpBatch,
  part: string,
  copied: CopiedFormat,
  selection: EditorSelection
): string[] {
  const ids = selectedIds(selection);
  if (ids.length === 0) throw new CommandRefusal('disabled.noSelection');
  const n = slideNumberOf(batch, part);
  const painted: string[] = [];
  if (copied.shape) {
    const shapes = ids.filter((id) => paintShape(batch, part, id, copied, copied.shape!));
    if (shapes.length === 0) throw new CommandRefusal('disabled.unsupportedElement');
    for (const id of shapes) {
      const txBody = child(batch.element(part, id), 'p:txBody');
      if (!txBody || !copied.run) continue;
      for (const r of runsForFormatting(batch, part, txBody.id))
        paintRun(batch, part, r, copied.run);
      painted.push(id);
    }
    return painted;
  }
  for (const id of ids) {
    // A cell range or a selected table: all text of each cell, as Bold formats it.
    const table = selection.kind !== 'text' && tblOf(batch.element(part, id));
    if (table) {
      for (const t of tableTextTargets(batch, part, selection, id)) {
        if (copied.run)
          for (const r of runsForFormatting(batch, part, t.txBodyId))
            paintRun(batch, part, r, copied.run);
        if (copied.paragraph)
          for (const pId of paragraphsForFormatting(batch, part, t.txBodyId))
            paintParagraph(batch, part, pId, copied.paragraph);
      }
      painted.push(id);
      continue;
    }
    const txBody =
      selection.kind === 'text'
        ? selectionBody(batch, part, selection)
        : child(batch.element(part, id), 'p:txBody');
    if (!txBody) continue;
    let range: { from: TextPoint; to: TextPoint } | undefined;
    if (selection.kind === 'text') {
      const [from, to] = ordered(selection);
      range = { from, to };
      if (comparePoints(from, to) === 0) {
        const p = paragraphsOf(txBody)[from.paragraph];
        const word = p ? wordAt(p, from.offset, n) : null;
        if (!word) throw new CommandRefusal('disabled.noWord');
        range = {
          from: { paragraph: from.paragraph, offset: word[0] },
          to: { paragraph: from.paragraph, offset: word[1] },
        };
      }
    }
    if (copied.run)
      for (const r of runsForFormatting(batch, part, txBody.id, range))
        paintRun(batch, part, r, copied.run);
    if (copied.paragraph)
      for (const pId of paragraphsForFormatting(batch, part, txBody.id, range))
        paintParagraph(batch, part, pId, copied.paragraph);
    painted.push(id);
  }
  if (painted.length === 0) throw new CommandRefusal('disabled.noText');
  return painted;
}

/**
 * The painter's state: the copied formatting and whether the painter is armed. It lives in the
 * editor, not in the adapter, and the document never holds it.
 */
export class FormatPainter {
  #mode: FormatPainterMode = 'off';
  #copied: CopiedFormat | null = null;
  #lastPress: number | null = null;
  #dismissed: number | null = null;
  /** A new object on every change, so caches keyed on it see painter changes. */
  state: object = {};

  get mode(): FormatPainterMode {
    return this.#mode;
  }

  #set(mode: FormatPainterMode, copied: CopiedFormat | null = this.#copied): void {
    this.#mode = mode;
    this.#copied = copied;
    this.state = {};
  }

  /**
   * Run a painter command. Returns null for any other command, `state` when only the painter
   * changed, or the compiled paste. A probe (`Editor.can`) changes nothing. Throws
   * CommandRefusal.
   */
  compile(
    ctx: CommandContext,
    command: EditorCommand
  ): null | 'state' | { batch: OpBatch; effect: CommandEffect } {
    switch (command.type) {
      case 'copyFormat': {
        const copied = copyFormat(ctx);
        if (!ctx.probe) this.#set(this.#mode, copied);
        return 'state';
      }
      case 'toggleFormatPainter':
        if (ctx.readOnly) throw new CommandRefusal('disabled.viewing');
        if (!ctx.probe) this.#press(ctx, command.sticky === true);
        else if (this.#mode === 'off') copyFormat(ctx);
        return 'state';
      case 'pasteFormat': {
        if (ctx.readOnly) throw new CommandRefusal('disabled.viewing');
        if (!this.#copied) throw new CommandRefusal('disabled.noFormatCopied');
        const part = ctx.session.info.slides[ctx.slide]?.partName;
        if (!part) throw new CommandRefusal('disabled.noEditor');
        const batch = ctx.session.batch();
        const painted = pasteFormat(batch, part, this.#copied, ctx.selection);
        if (!ctx.probe) {
          refitAll(batch, ctx.slide, part, painted, ctx.session.measurer);
          // A paste ends a painter armed for one paint.
          if (this.#mode === 'once') this.#set('off');
        }
        return { batch, effect: { label: 'pasteFormat' } };
      }
      default:
        return null;
    }
  }

  /** A press of the Format Painter button. */
  #press(ctx: CommandContext, sticky: boolean): void {
    const now = Date.now();
    const double = this.#lastPress !== null && now - this.#lastPress < DOUBLE_PRESS_MS;
    const previous = this.#mode;
    this.#lastPress = now;
    try {
      if (sticky || (previous === 'once' && double)) {
        this.#set('sticky', copyFormat(ctx));
      } else if (previous !== 'off') {
        // A press while it is on turns it off. Remember a sticky one, so the second half of
        // the double-click that put it away does not arm it again.
        this.#dismissed = previous === 'sticky' ? now : null;
        this.#set('off');
      } else if (this.#dismissed !== null && now - this.#dismissed < DOUBLE_PRESS_MS) {
        this.#dismissed = null;
      } else {
        this.#set('once', copyFormat(ctx));
        this.#dismissed = null;
      }
    } catch (e) {
      // A refused press opens no double-press window and leaves the painter off.
      this.#lastPress = null;
      if (previous !== 'off') this.#set('off');
      throw e;
    }
  }

  /** Turn the painter off (Escape). True when it was on. */
  cancel(): boolean {
    if (this.#mode === 'off') return false;
    this.#set('off');
    return true;
  }

  /** Turn the painter off; `forget` also drops the copied formatting (a new document). */
  reset(forget = false): void {
    if (this.#mode !== 'off' || (forget && this.#copied))
      this.#set('off', forget ? null : this.#copied);
  }
}
