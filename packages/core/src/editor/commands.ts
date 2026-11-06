// The command table: every EditorCommand compiles to ONE store transaction, or is refused
// with an i18n reason key. `can` runs the same compiler against a throwaway batch, so enabled
// state and execution can never disagree (docx-editor's `Editor.can` contract). A probe skips
// only the autofit re-measurement (`refitShape`), which never refuses (its sizes are bounded
// by `XFRM_BOUND`, and a missing `a:off` or `a:ext` is added).

import {
  attr,
  child,
  elements,
  NS,
  REL,
  sanitizeHref,
  type Element,
  type NodeSpec,
  type PptxPackage,
} from '../store/index.ts';
import type { LayoutItem } from '../layout/index.ts';
import type { EditorCommand, EditorSelection, TransformItem } from '../contracts/index.ts';
import { refit } from '../automation/commands/autofit.ts';
import { refitTable, resizeTable } from '../automation/commands/table-fit.ts';
import { tblOf } from '../automation/commands/table-model.ts';
import { writeCopiedTableMarkers } from '../automation/commands/table-repair.ts';
import { pasteShapes } from './clipboard.ts';
import type { DocumentSession } from './document-session.ts';
import {
  fileXfrm,
  groupShapes,
  refitGroup,
  setTransforms,
  shapesOf,
  ungroupShapes,
} from '../automation/commands/group-commands.ts';
import { CommandRefusal, OpBatch } from '../automation/commands/op-batch.ts';
import { addExternalRelationship, parentOf } from '../automation/commands/package-edits.ts';
import {
  boxOf,
  deltaToParent,
  findItem,
  toFile,
  toParent,
} from '../automation/commands/placement.ts';
import {
  alignBoxes,
  deleteShapes,
  duplicateShapes,
  insertImage,
  insertLine,
  insertShape,
  insertTextBox,
  orderShapes,
  specFromXml,
} from '../automation/commands/shape-commands.ts';
import {
  colorXml,
  setLineColor,
  setLineDash,
  setLineWidth,
  setShapeFill,
  setShapeGradient,
  setTextFitting,
  setLineEnds,
  setVerticalAlign,
} from '../automation/commands/style-commands.ts';
import { setShapeEffect, setShapeStyle } from '../automation/commands/effect-commands.ts';
import {
  distributeBoxes,
  flipped,
  parentFrame,
  renameShape,
  rotated,
  setLockAspectRatio,
  setShapesVisible,
} from '../automation/commands/arrange-commands.ts';
import { deleteSlides, duplicateSlides, moveSlides } from '../automation/commands/slide-multi.ts';
import { setSlideNumbers } from '../automation/commands/slide-numbers.ts';
import { notesShown, setNotes } from '../automation/commands/notes-commands.ts';
import { compilePictureCommand, isPictureCommand } from './picture-commands.ts';
import { compileSlideCommand, isSlideCommand } from './slide-command-table.ts';
import { compileTextFormatCommand, isTextFormatCommand } from './text-format-commands.ts';
import {
  forParagraphs,
  forRuns,
  listTarget,
  refitShape,
  refitTargets,
  selectedIds,
  slidePart,
  textTargets,
} from './text-targets.ts';

export { selectedIds };
import { normalizePresentation } from '../automation/commands/slide-sections.ts';
import { deckLayout, newSlideLayout } from '../automation/commands/slide-layouts.ts';
import { findMatches, replaceMatches } from '../automation/commands/find.ts';
import {
  removeBulletInPlace,
  setListKind,
  shiftListLevel,
} from '../automation/commands/list-commands.ts';
import { selectionBody, textItemOf } from './table-cells.ts';
import { coversWholeBody } from './text-drag.ts';
import { commandWordAt } from '../automation/commands/command-word.ts';
import { clearCells, compileTable, isTableSelection, setTableAnchor } from './table-commands.ts';
import { MAX_MULTIPLE, setLineSpacingMultiple } from '../automation/commands/paragraph-format.ts';
import { isLatinTypeface } from '../automation/commands/font-format.ts';
import {
  addSlide,
  deleteSlide,
  duplicateSlide,
  moveSlide,
  setSlideBackground,
} from '../automation/commands/slide-commands.ts';
import {
  aSpec,
  paragraphsOf,
  runsForFormatting,
  setRunAttr,
  setRunChild,
  slideNumberOf,
  slotsOf,
} from '../automation/commands/text-model.ts';

/** What a compiled command does to editor state besides the document. */
export interface CommandEffect {
  readonly selection?: EditorSelection;
  readonly slide?: number;
  /** The slides the rail selects after the command; `slide` is the current one among them. */
  readonly slides?: readonly number[];
  readonly label: string;
}

/** Editor state a command compiles against. */
export interface CommandContext {
  readonly session: DocumentSession;
  readonly slide: number;
  /** The slides selected in the rail, ascending; always holds `slide`. */
  readonly selectedSlides?: readonly number[];
  /** The host's translator, for text a command writes into the deck (a section name). */
  readonly translate?: (key: string) => string;
  readonly selection: EditorSelection;
  readonly readOnly: boolean;
  /** The Slide Sorter view is on: it shows no slide to put a new shape on. */
  readonly slideSorter?: boolean;
  /**
   * Only find out whether the command can run (`Editor.can`): skip the autofit re-measurement,
   * which adds ops but never refuses. The ops of a probe are never committed.
   */
  readonly probe?: boolean;
  /**
   * The editor is in a collaboration room. A merge then keeps a hidden copy of the text it moves
   * in each covered cell, so a peer's concurrent delete of the top-left cell cannot lose it.
   */
  readonly collaborating?: boolean;
  /**
   * The package when the pointer gesture that sends this command started (a handle drag sends
   * one `setTransforms` per move). A table resize scales from it, so the path does not count.
   * @internal
   */
  readonly gestureBase?: PptxPackage;
  /** How many times in a row this payload has been pasted, counting this paste (1 = first). */
  readonly pasteRepeat?: number;
  /** The package to compile against instead of the store's: the preview of an armed format. */
  readonly base?: PptxPackage;
  /**
   * Compile a format that is already armed at the caret (`caret-arm.ts`): a caret inside a word
   * arms too, as the arm was decided when it started. The editor sets it for its replays.
   */
  readonly caretArm?: boolean;
  /** Formatting state for toggles (true when the selection is already bold, …). */
  readonly active: {
    bold: boolean;
    italic: boolean;
    underline: boolean;
    strike: boolean;
    bullets: boolean;
    numbering: boolean;
    fontSize?: number;
    /** `fontSize` as drawn under `a:normAutofit/@fontScale` (smaller when the text is shrunk). */
    renderedFontSize?: number;
    /** The selected text has more than one size: the size steppers are off, as in DOCX Editor. */
    fontSizeMixed?: boolean;
    superscript?: boolean;
    subscript?: boolean;
  };
}

/** Commands that edit the slide list or the sections: they act on the repaired part. */
const PRESENTATION_EDITS: ReadonlySet<EditorCommand['type']> = new Set([
  'addSlide',
  'duplicateSlide',
  'duplicateSlides',
  'deleteSlide',
  'deleteSlides',
  'moveSlide',
  'moveSlides',
  'addSection',
  'renameSection',
  'removeSection',
  'removeAllSections',
  'moveSection',
]);

/** Commands that put a shape on the current slide. */
const CANVAS_INSERTS: ReadonlySet<EditorCommand['type']> = new Set([
  'insertTextBox',
  'insertShape',
  'insertLine',
  'insertImage',
  'insertTable',
  'paste',
]);

const FONT_SIZES = [
  8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96,
];

/** The group scope of a shape selection. */
const scopeOf = (selection: EditorSelection): string | undefined =>
  selection.kind === 'shapes' ? selection.group : undefined;

/** A shapes selection in the current scope. */
function shapesSelection(ids: readonly string[], group: string | undefined): EditorSelection {
  return { kind: 'shapes', nodeIds: ids, ...(group ? { group } : {}) };
}

/** Laid-out items of the selected shapes. */
function selectedItems(ctx: CommandContext, ids: readonly string[]): LayoutItem[] {
  const all = ctx.session.layout(ctx.slide)?.items ?? [];
  return ids.map((id) => {
    const item = findItem(all, id);
    if (!item) throw new CommandRefusal('disabled.unsupportedElement');
    return item;
  });
}

/**
 * Refuse a Distribute or Rotate that moves and turns nothing (a turn of 0° or 360°, boxes that
 * are already evenly spaced), as the lock, hide, and rename commands refuse: no undo step.
 */
function refuseUnchanged(
  batch: OpBatch,
  ctx: CommandContext,
  part: string,
  items: readonly TransformItem[]
): void {
  const layout = ctx.session.layout(ctx.slide)?.items ?? [];
  const turn = (deg: number): number =>
    ((Math.round(deg * 60_000) % 21_600_000) + 21_600_000) % 21_600_000;
  const same = items.every((t) => {
    const old = fileXfrm(batch, part, t.nodeId, layout);
    return (
      Math.round(old.x) === Math.round(t.x) &&
      Math.round(old.y) === Math.round(t.y) &&
      Math.round(old.cx) === Math.round(t.cx) &&
      Math.round(old.cy) === Math.round(t.cy) &&
      turn(old.rot) === turn(t.rot) &&
      (t.flipH === undefined || t.flipH === old.flipH) &&
      (t.flipV === undefined || t.flipV === old.flipV)
    );
  });
  if (same) throw new CommandRefusal('disabled.noArrangeChange');
}

/** Write transforms, re-fit autofit text, and refit the groups around them. */
function writeTransforms(
  batch: OpBatch,
  ctx: CommandContext,
  part: string,
  items: readonly TransformItem[]
): void {
  const layout = ctx.session.layout(ctx.slide)?.items ?? [];
  const start = batch.pkg;
  const resized = items.filter((t) => {
    const old = fileXfrm(batch, part, t.nodeId, layout);
    return Math.abs(old.cx - t.cx) > 1 || Math.abs(old.cy - t.cy) > 1;
  });
  setTransforms(batch, part, items, layout);
  // A move keeps the text layout; a resize re-fits autofit text. A table scales its columns
  // and rows from the drag's start (`gestureBase`), also on a move that comes back to the
  // frame's box, so one box gives one table whatever path the pointer took; the frame then
  // takes the laid-out table size.
  for (const t of items) {
    if (tblOf(batch.element(part, t.nodeId))) {
      resizeTable(
        batch,
        ctx.slide,
        part,
        t.nodeId,
        t,
        ctx.gestureBase ?? start,
        ctx.session.measurer
      );
      if (!ctx.probe) refitTable(batch, ctx.slide, part, t.nodeId, ctx.session.measurer);
    } else if (resized.includes(t) && child(batch.element(part, t.nodeId), 'p:txBody'))
      refitShape(batch, ctx, part, t.nodeId);
  }
  // Autofit can change a height; refit the groups again so their bounds stay tight.
  const groups = new Set(
    items.map((t) => parentOf(batch, part, t.nodeId)).filter((p) => p.kind === 'grpSp')
  );
  for (const g of groups) refitGroup(batch, part, g.id, layout);
}

/**
 * The runs of the command word at a text caret (split at the word's edges), or none between
 * words. Insert link at a caret links them.
 */
function commandWordRuns(
  batch: OpBatch,
  part: string,
  sel: Extract<EditorSelection, { kind: 'text' }>
): Element[] {
  const body = selectionBody(batch, part, sel);
  const range = body ? commandWordAt(batch, part, body.id, sel.focus) : null;
  return body && range ? runsForFormatting(batch, part, body.id, range) : [];
}

/** The runs of the link under a caret (the run before it carries the link), or none. */
function linkedRunsAt(
  batch: OpBatch,
  part: string,
  sel: Extract<EditorSelection, { kind: 'text' }>
): Element[] {
  const txBody = selectionBody(batch, part, sel);
  if (!txBody) return [];
  const n = slideNumberOf(batch, part);
  const slots = paragraphsOf(txBody).flatMap((p) => slotsOf(p, n));
  const para = paragraphsOf(txBody)[sel.focus.paragraph];
  const here = para
    ? slotsOf(para, n).find((sl) => sel.focus.offset > sl.start && sel.focus.offset <= sl.end)
    : undefined;
  const rIdOf = (el: Element): string | undefined =>
    attr(child(child(el, 'a:rPr'), 'a:hlinkClick'), 'r:id');
  const rId = here ? rIdOf(here.el) : undefined;
  if (!rId) return [];
  return slots.map((sl) => sl.el).filter((el) => rIdOf(el) === rId);
}

/** What clear formatting keeps on a run: its language and proofing state, and its link. */
const KEEP_RUN_ATTRS = new Set(['lang', 'altLang', 'dirty', 'err', 'noProof', 'smtClean']);
const KEEP_RUN_CHILDREN = new Set(['hlinkClick', 'hlinkMouseOver', 'extLst']);

function setPPrAttr(
  batch: OpBatch,
  part: string,
  pPrId: string,
  name: string,
  value: string | null
): void {
  batch.push({ op: 'setAttribute', part, node: pPrId, name: { ns: '', local: name }, value });
}

/**
 * Compile a command against a fresh batch over the store's package (or `ctx.base`). Throws
 * CommandRefusal. A batch marked
 * `armedCaret` holds a format armed at a caret for the text typed next: the editor keeps it as
 * session state (`caret-arm.ts`) and commits it only with that text, so a host that compiles
 * commands itself must not commit such a batch alone.
 */
export function compile(
  ctx: CommandContext,
  command: EditorCommand
): { batch: OpBatch; effect: CommandEffect } {
  if (
    command.type === 'undo' ||
    command.type === 'redo' ||
    command.type === 'toggleSmartGuides' ||
    command.type === 'setPresentationView' ||
    command.type === 'setSectionsCollapsed' ||
    command.type === 'setTableBorderPen'
  )
    throw new Error('history and view commands do not compile');
  if (ctx.readOnly) throw new CommandRefusal('disabled.viewing');
  // In the Slide Sorter, the Insert commands are off: a new shape waits for the Normal view.
  if (ctx.slideSorter && CANVAS_INSERTS.has(command.type))
    throw new CommandRefusal('disabled.slideSorter');
  const batch = ctx.probe ? ctx.session.probeBatch(ctx.base) : ctx.session.batch(ctx.base);
  if (PRESENTATION_EDITS.has(command.type)) normalizePresentation(batch);
  if (isSlideCommand(command)) return { batch, effect: compileSlideCommand(ctx, batch, command) };
  if (isPictureCommand(command))
    return { batch, effect: compilePictureCommand(ctx, batch, slidePart(ctx), command) };
  if (isTextFormatCommand(command))
    return { batch, effect: compileTextFormatCommand(ctx, batch, command) };
  const info = ctx.session.info;
  const { width: slideW, height: slideH } = info;
  switch (command.type) {
    case 'toggleBold':
    case 'toggleItalic':
    case 'toggleUnderline':
    case 'toggleStrike': {
      const part = slidePart(ctx);
      const [name, on, off, active] =
        command.type === 'toggleBold'
          ? ['b', '1', '0', ctx.active.bold]
          : command.type === 'toggleItalic'
            ? ['i', '1', '0', ctx.active.italic]
            : command.type === 'toggleUnderline'
              ? ['u', 'sng', 'none', ctx.active.underline]
              : ['strike', 'sngStrike', 'noStrike', ctx.active.strike];
      forRuns(batch, part, ctx, (run) => setRunAttr(batch, part, run, name, active ? off : on));
      return { batch, effect: { label: command.type } };
    }
    case 'setFontSize':
    case 'stepFontSize': {
      const part = slidePart(ctx);
      let points: number;
      if (command.type === 'setFontSize') points = command.points;
      else {
        if (ctx.active.fontSizeMixed) throw new CommandRefusal('disabled.mixedSize');
        // Increase steps from the stored size. Decrease on whole shrunk text bodies (a shape
        // selection, or a range over all the text) steps from the drawn size: the whole body
        // takes the new size and fits again from 100 %, so the size shown goes down. A word or
        // a caret in shrunk text keeps the body's scale, so it steps from the stored size, as
        // the scale applies once more.
        const stored = ctx.active.fontSize ?? 18;
        const layout = ctx.session.layout(ctx.slide);
        const wholeBodies = coversWholeBody(
          ctx.selection,
          layout && ctx.selection.kind === 'text' ? textItemOf(layout, ctx.selection) : undefined
        );
        const current =
          command.direction < 0 && wholeBodies
            ? Math.min(stored, ctx.active.renderedFontSize ?? stored)
            : stored;
        points =
          command.direction > 0
            ? (FONT_SIZES.find((s) => s > current) ?? current + 10)
            : ([...FONT_SIZES].reverse().find((s) => s < current) ?? Math.max(1, current - 1));
      }
      if (!Number.isFinite(points) || points < 1 || points > 4000)
        throw new CommandRefusal('errors.unsupported', 'size');
      forRuns(batch, part, ctx, (run) =>
        setRunAttr(batch, part, run, 'sz', String(Math.round(points * 100)))
      );
      return { batch, effect: { label: 'fontSize' } };
    }
    case 'setFontFamily': {
      const part = slidePart(ctx);
      const family = command.family.trim();
      // The same bound as the font catalog: a family name can reach a CSS declaration. A Latin
      // theme font reference (`+mj-lt`, `+mn-lt`) is written as is, so the text follows the theme.
      if (!isLatinTypeface(family)) throw new CommandRefusal('errors.unsupported', 'family');
      forRuns(batch, part, ctx, (run) =>
        setRunChild(batch, part, run, ['latin'], aSpec('latin', { typeface: family }))
      );
      return { batch, effect: { label: 'fontFamily' } };
    }
    case 'setTextColor':
    case 'setHighlight': {
      const part = slidePart(ctx);
      const color = command.color;
      const highlight = command.type === 'setHighlight';
      const group = highlight
        ? ['highlight']
        : ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
      const clr = color !== null ? specFromXml(colorXml(color)) : null;
      const spec = clr
        ? highlight
          ? aSpec('highlight', {}, [clr])
          : aSpec('solidFill', {}, [clr])
        : null;
      // Highlight is a pen, as in DOCX Editor: at a caret it marks the text typed next.
      forRuns(batch, part, ctx, (run) => setRunChild(batch, part, run, group, spec), !highlight);
      return { batch, effect: { label: command.type } };
    }
    case 'setLink': {
      const part = slidePart(ctx);
      const sel = ctx.selection;
      const collapsed =
        sel.kind === 'text' &&
        sel.anchor.paragraph === sel.focus.paragraph &&
        sel.anchor.offset === sel.focus.offset;
      // With only a caret, the command acts on the link under it: every run with its `r:id`.
      const linked = collapsed && sel.kind === 'text' ? linkedRunsAt(batch, part, sel) : [];
      // Else the word at or next to the caret, as Change Case reaches it (`command-word.ts`).
      const word =
        collapsed && linked.length === 0 && sel.kind === 'text'
          ? commandWordRuns(batch, part, sel)
          : [];
      if (collapsed && linked.length === 0 && word.length === 0)
        throw new CommandRefusal('disabled.selectText');
      let spec: NodeSpec | null = null;
      if (command.url !== null) {
        const raw = command.url.trim();
        if (!raw || raw.length > 2048) throw new CommandRefusal('errors.unsafeLink');
        // A bare host ("example.com") is a web address, as in Google Slides.
        const withScheme = /^[a-z][a-z0-9+.-]*:/i.test(raw) ? raw : `https://${raw}`;
        const safe = sanitizeHref(withScheme);
        if (!safe.ok || !/^[a-z][a-z0-9+.-]*:/i.test(safe.href))
          throw new CommandRefusal('errors.unsafeLink');
        const rId = addExternalRelationship(batch, part, REL.hyperlink, safe.href);
        spec = {
          ns: NS.a,
          local: 'hlinkClick',
          attrs: [{ ns: NS.r, local: 'id', prefix: 'r', value: rId }],
        };
      }
      if (linked.length)
        for (const run of linked) setRunChild(batch, part, run, ['hlinkClick'], spec);
      else if (word.length)
        for (const run of word) setRunChild(batch, part, run, ['hlinkClick'], spec);
      else forRuns(batch, part, ctx, (run) => setRunChild(batch, part, run, ['hlinkClick'], spec));
      return { batch, effect: { label: 'link' } };
    }
    case 'clearFormatting': {
      const part = slidePart(ctx);
      forRuns(batch, part, ctx, (run) => {
        // A paragraph mark (`a:endParaRPr`) is its own properties: an empty paragraph's caret.
        const el = batch.element(part, run.id);
        const rPr = el.local === 'endParaRPr' ? el : child(el, 'a:rPr');
        if (!rPr) return;
        for (const a of rPr.attrs)
          if (!KEEP_RUN_ATTRS.has(a.local))
            batch.push({
              op: 'setAttribute',
              part,
              node: rPr.id,
              name: { ns: a.ns, local: a.local },
              value: null,
            });
        for (const c of elements(batch.element(part, rPr.id)))
          if (!KEEP_RUN_CHILDREN.has(c.local)) batch.push({ op: 'removeNode', part, node: c.id });
      });
      return { batch, effect: { label: 'clearFormatting' } };
    }
    case 'setAlign': {
      const part = slidePart(ctx);
      forParagraphs(batch, part, ctx, (pPr) => setPPrAttr(batch, part, pPr, 'algn', command.align));
      return { batch, effect: { label: 'align' } };
    }
    case 'toggleBullets':
    case 'toggleNumbering': {
      const part = slidePart(ctx);
      const active = command.type === 'toggleBullets' ? ctx.active.bullets : ctx.active.numbering;
      const kind = active ? 'none' : command.type === 'toggleBullets' ? 'bullet' : 'number';
      const targets = textTargets(batch, part, ctx.selection);
      for (const t of targets)
        setListKind(batch, part, ctx.slide, listTarget(batch, part, t), kind, !ctx.probe);
      refitTargets(batch, ctx, part, targets);
      return { batch, effect: { label: command.type } };
    }
    case 'indent': {
      const part = slidePart(ctx);
      forParagraphs(batch, part, ctx, (pPr) => {
        shiftListLevel(batch, part, pPr, command.direction);
      });
      return { batch, effect: { label: 'indent' } };
    }
    case 'removeBullet': {
      const part = slidePart(ctx);
      const targets = textTargets(batch, part, ctx.selection);
      let removed = false;
      for (const t of targets)
        removed =
          removeBulletInPlace(batch, part, ctx.slide, listTarget(batch, part, t)) || removed;
      if (!removed) throw new CommandRefusal('disabled.noList');
      refitTargets(batch, ctx, part, targets);
      return { batch, effect: { label: 'removeBullet' } };
    }
    case 'setLineSpacing': {
      const part = slidePart(ctx);
      // The range of the Paragraph dialog's Multiple.
      if (!(command.multiple >= 0.01 && command.multiple <= MAX_MULTIPLE))
        throw new CommandRefusal('errors.unsupported', 'spacing');
      forParagraphs(batch, part, ctx, (pPr) =>
        setLineSpacingMultiple(batch, part, pPr, command.multiple)
      );
      return { batch, effect: { label: 'lineSpacing' } };
    }
    case 'insertTextBox': {
      const id = insertTextBox(batch, slidePart(ctx), slideW, slideH);
      return {
        batch,
        effect: {
          label: 'insertTextBox',
          selection: {
            kind: 'text',
            nodeId: id,
            anchor: { paragraph: 0, offset: 0 },
            focus: { paragraph: 0, offset: 0 },
          },
        },
      };
    }
    case 'insertShape': {
      const id = insertShape(batch, slidePart(ctx), command.preset, slideW, slideH);
      return {
        batch,
        effect: { label: 'insertShape', selection: { kind: 'shapes', nodeIds: [id] } },
      };
    }
    case 'insertImage': {
      const id = insertImage(
        batch,
        slidePart(ctx),
        command.bytes,
        command.mime,
        command.width,
        command.height,
        slideW,
        slideH
      );
      return {
        batch,
        effect: { label: 'insertImage', selection: { kind: 'shapes', nodeIds: [id] } },
      };
    }
    case 'duplicate': {
      const ids = selectedIds(ctx.selection);
      if (!ids.length) throw new CommandRefusal('disabled.noSelection');
      const part = slidePart(ctx);
      const out = duplicateShapes(batch, part, ids);
      writeCopiedTableMarkers(batch, part, out);
      const group = scopeOf(ctx.selection);
      if (group) refitGroup(batch, part, group, ctx.session.layout(ctx.slide)?.items ?? []);
      return { batch, effect: { label: 'duplicate', selection: shapesSelection(out, group) } };
    }
    case 'delete': {
      // Selected cells lose their text; the table stays.
      if (ctx.selection.kind === 'cells') return clearCells(ctx, batch);
      const ids = selectedIds(ctx.selection);
      if (!ids.length) throw new CommandRefusal('disabled.noSelection');
      const part = slidePart(ctx);
      let group = scopeOf(ctx.selection);
      deleteShapes(batch, part, ids);
      // A group left empty goes too; one with children left keeps tight bounds.
      while (group) {
        const g = batch.element(part, group);
        if (shapesOf(g).length > 0) {
          refitGroup(batch, part, group, ctx.session.layout(ctx.slide)?.items ?? []);
          break;
        }
        const parent = parentOf(batch, part, group);
        deleteShapes(batch, part, [group]);
        group = parent.kind === 'grpSp' ? parent.id : undefined;
      }
      return { batch, effect: { label: 'delete', selection: { kind: 'none' } } };
    }
    case 'order': {
      const ids = selectedIds(ctx.selection);
      if (!ids.length) throw new CommandRefusal('disabled.noSelection');
      orderShapes(batch, slidePart(ctx), ids, command.move);
      return { batch, effect: { label: 'order' } };
    }
    case 'alignObjects': {
      const ids = selectedIds(ctx.selection);
      if (!ids.length) throw new CommandRefusal('disabled.noSelection');
      const items = selectedItems(ctx, ids);
      // One element aligns to the slide, in its parent's space.
      const a = toParent(items[0], 0, 0);
      const b = toParent(items[0], slideW, slideH);
      const space = {
        x: Math.min(a.x, b.x),
        y: Math.min(a.y, b.y),
        w: Math.abs(b.x - a.x),
        h: Math.abs(b.y - a.y),
      };
      const boxes = items.map((item) => ({ id: item.element.nodeId, ...boxOf(item) }));
      const moved = alignBoxes(
        boxes.map((x) => ({ ...x, x: x.x - space.x, y: x.y - space.y })),
        command.align,
        space.w,
        space.h
      );
      writeTransforms(
        batch,
        ctx,
        slidePart(ctx),
        moved.map((m, i) => toFile(items[i], { ...boxes[i], x: m.x + space.x, y: m.y + space.y }))
      );
      return { batch, effect: { label: 'alignObjects' } };
    }
    case 'setTransform': {
      writeTransforms(batch, ctx, slidePart(ctx), [command]);
      return { batch, effect: { label: 'transform' } };
    }
    case 'setTransforms': {
      if (!command.items.length) throw new CommandRefusal('disabled.noSelection');
      writeTransforms(batch, ctx, slidePart(ctx), command.items);
      return { batch, effect: { label: 'transform' } };
    }
    case 'nudge': {
      const ids = selectedIds(ctx.selection);
      if (!ids.length) throw new CommandRefusal('disabled.noSelection');
      const items = selectedItems(ctx, ids).map((item) => {
        const d = deltaToParent(item, command.dx, command.dy);
        const box = boxOf(item);
        return toFile(item, { ...box, x: box.x + d.x, y: box.y + d.y });
      });
      writeTransforms(batch, ctx, slidePart(ctx), items);
      return { batch, effect: { label: 'nudge' } };
    }
    case 'group': {
      const ids = selectedIds(ctx.selection);
      const group = groupShapes(
        batch,
        slidePart(ctx),
        ids,
        ctx.session.layout(ctx.slide)?.items ?? []
      );
      return {
        batch,
        effect: {
          label: 'group',
          selection: shapesSelection([group], scopeOf(ctx.selection)),
        },
      };
    }
    case 'ungroup': {
      const ids = selectedIds(ctx.selection);
      if (!ids.length) throw new CommandRefusal('disabled.notAGroup');
      const out = ungroupShapes(
        batch,
        slidePart(ctx),
        ids,
        ctx.session.layout(ctx.slide)?.items ?? []
      );
      return {
        batch,
        effect: { label: 'ungroup', selection: shapesSelection(out, scopeOf(ctx.selection)) },
      };
    }
    case 'paste': {
      const part = slidePart(ctx);
      const p = command.payload as unknown as { document?: unknown; slidePart?: unknown };
      const here = p?.document === ctx.session.id && p.slidePart === part;
      // Each repeated paste onto the source slide moves 0.25 in further.
      const out = pasteShapes(
        batch,
        part,
        command.payload,
        here ? 228_600 * (ctx.pasteRepeat ?? 1) : 0
      );
      if (!out.length) throw new CommandRefusal('errors.clipboard');
      writeCopiedTableMarkers(batch, part, out);
      return { batch, effect: { label: 'paste', selection: shapesSelection(out, undefined) } };
    }
    case 'setShapeFill': {
      // In a table, Shape Fill shades the cells.
      if (isTableSelection(ctx))
        return compileTable(ctx, batch, { type: 'setCellFill', color: command.color });
      setShapeFill(batch, slidePart(ctx), selectedIds(ctx.selection), command.color);
      return { batch, effect: { label: 'shapeFill' } };
    }
    case 'setShapeGradient': {
      setShapeGradient(batch, slidePart(ctx), selectedIds(ctx.selection), command.gradient);
      return { batch, effect: { label: 'shapeFill' } };
    }
    case 'setLineColor': {
      setLineColor(batch, slidePart(ctx), selectedIds(ctx.selection), command.color);
      return { batch, effect: { label: 'lineColor' } };
    }
    case 'setLineWidth': {
      const items = ctx.session.layout(ctx.slide)?.items ?? [];
      const visible = (id: string): boolean => {
        const el = findItem(items, id)?.element;
        const line = el && 'line' in el ? el.line : undefined;
        return !!line && line.fill.type !== 'none';
      };
      setLineWidth(batch, slidePart(ctx), selectedIds(ctx.selection), command.points, visible);
      return { batch, effect: { label: 'lineWidth' } };
    }
    case 'setLineDash': {
      setLineDash(batch, slidePart(ctx), selectedIds(ctx.selection), command.dash);
      return { batch, effect: { label: 'lineDash' } };
    }
    case 'setVerticalAlign': {
      if (!setTableAnchor(ctx, batch, command.anchor))
        setVerticalAlign(batch, slidePart(ctx), selectedIds(ctx.selection), command.anchor);
      return { batch, effect: { label: 'verticalAlign' } };
    }
    case 'setLineEnds': {
      setLineEnds(batch, slidePart(ctx), selectedIds(ctx.selection), command.head, command.tail);
      return { batch, effect: { label: 'lineEnds' } };
    }
    case 'setSlideBackground': {
      const slides = command.allSlides ? info.slides.map((_, i) => i) : [ctx.slide];
      for (const i of slides) setSlideBackground(batch, i, command.background);
      return { batch, effect: { label: 'slideBackground' } };
    }
    case 'insertLine': {
      const id = insertLine(batch, slidePart(ctx), command.kind, slideW, slideH);
      return {
        batch,
        effect: { label: 'insertLine', selection: { kind: 'shapes', nodeIds: [id] } },
      };
    }
    case 'setAutofit': {
      const part = slidePart(ctx);
      for (const id of setTextFitting(batch, part, selectedIds(ctx.selection), command.mode))
        refitShape(batch, ctx, part, id);
      return { batch, effect: { label: 'textFitting' } };
    }
    case 'addSlide': {
      const layout =
        command.layoutId === undefined
          ? newSlideLayout(batch.pkg, ctx.slide)
          : deckLayout(batch.pkg, command.layoutId);
      const slide = addSlide(batch, ctx.slide, layout);
      return { batch, effect: { label: 'addSlide', slide, selection: { kind: 'none' } } };
    }
    case 'duplicateSlide': {
      const slide = duplicateSlide(batch, command.index);
      return { batch, effect: { label: 'duplicateSlide', slide, selection: { kind: 'none' } } };
    }
    case 'deleteSlide': {
      deleteSlide(batch, command.index);
      const remaining = info.slides.length - 1;
      return {
        batch,
        effect: {
          label: 'deleteSlide',
          slide: Math.min(command.index, remaining - 1),
          selection: { kind: 'none' },
        },
      };
    }
    case 'moveSlide': {
      moveSlide(batch, command.from, command.to, command.section);
      // Node ids are per part: a selection on the old current slide means nothing after it.
      return {
        batch,
        effect: {
          label: 'moveSlide',
          slide: Math.max(0, Math.min(command.to, info.slides.length - 1)),
          selection: { kind: 'none' },
        },
      };
    }
    case 'setSlideNumbers': {
      setSlideNumbers(batch, command.show);
      return { batch, effect: { label: 'slideNumbers', selection: { kind: 'none' } } };
    }
    case 'deleteSlides': {
      const slide = deleteSlides(batch, command.indices);
      return { batch, effect: { label: 'deleteSlides', slide, selection: { kind: 'none' } } };
    }
    case 'duplicateSlides':
    case 'moveSlides': {
      const slides =
        command.type === 'duplicateSlides'
          ? duplicateSlides(batch, command.indices)
          : moveSlides(batch, command.indices, command.to, command.section);
      // The current slide stays itself when it moves, and becomes its copy when duplicated.
      const k = [...new Set(command.indices)].sort((x, y) => x - y).indexOf(ctx.slide);
      const slide = slides[Math.max(0, k)];
      return {
        batch,
        effect: { label: command.type, slide, slides, selection: { kind: 'none' } },
      };
    }
    case 'setNotes': {
      // One commit of the notes field is one undo step: two edits never merge.
      setNotes(batch, command.index, command.text, notesShown(command.paragraphs));
      return { batch, effect: { label: 'notes' } };
    }
    case 'replace':
    case 'replaceAll': {
      const all = command.type === 'replaceAll';
      const matches = all
        ? findMatches(batch.pkg, command.query, command.options)
        : [command.match];
      if (!matches.length) throw new CommandRefusal('disabled.noMatch');
      replaceMatches(batch, matches, command.replacement);
      // Autofit text re-fits, as after typing, and a table frame takes its laid-out size;
      // notes have no autofit.
      const shapes = new Set(matches.filter((m) => m.where === 'slide').map((m) => m.nodeId));
      for (const m of matches) {
        const kind = shapes.delete(m.nodeId) ? batch.element(m.part, m.nodeId).kind : undefined;
        if ((kind === 'sp' || kind === 'graphicFrame') && !ctx.probe)
          refit(batch, m.slide, m.part, m.nodeId, ctx.session.measurer);
      }
      // The selection clears and the focus stays where it is (in the find bar): a text
      // selection here would move the focus to the slide. Find selects the next match itself.
      const effect = { label: command.type, selection: { kind: 'none' } } as const;
      return { batch, effect: all ? effect : { ...effect, slide: command.match.slide } };
    }
    case 'setShapeEffect': {
      setShapeEffect(batch, slidePart(ctx), selectedIds(ctx.selection), command.effect);
      return { batch, effect: { label: 'shapeEffect' } };
    }
    case 'setShapeStyle': {
      setShapeStyle(batch, slidePart(ctx), selectedIds(ctx.selection), command.style);
      return { batch, effect: { label: 'shapeStyle' } };
    }
    case 'distributeObjects': {
      const ids = selectedIds(ctx.selection);
      if (!ids.length) throw new CommandRefusal('disabled.noSelection');
      const items = selectedItems(ctx, ids);
      // One element centers on the slide, in its parent's space, as with align.
      const a = toParent(items[0], 0, 0);
      const b = toParent(items[0], slideW, slideH);
      const origin = { x: Math.min(a.x, b.x), y: Math.min(a.y, b.y) };
      const size = command.axis === 'vertical' ? Math.abs(b.y - a.y) : Math.abs(b.x - a.x);
      const boxes = items.map((item) => ({ id: item.element.nodeId, ...boxOf(item) }));
      const moved = distributeBoxes(
        boxes.map((x) => ({ ...x, x: x.x - origin.x, y: x.y - origin.y })),
        command.axis,
        size
      );
      const out = moved.map((m, i) =>
        toFile(items[i], { ...boxes[i], x: m.x + origin.x, y: m.y + origin.y })
      );
      refuseUnchanged(batch, ctx, slidePart(ctx), out);
      writeTransforms(batch, ctx, slidePart(ctx), out);
      return { batch, effect: { label: 'distributeObjects' } };
    }
    case 'flipObjects':
    case 'rotateObjects': {
      const ids = selectedIds(ctx.selection);
      if (!ids.length) throw new CommandRefusal('disabled.noSelection');
      // Tables do not flip or turn.
      const items = selectedItems(ctx, ids).filter((item) => item.kind !== 'frame');
      if (!items.length) throw new CommandRefusal('disabled.unsupportedElement');
      const out = items.map((item): TransformItem => {
        const box = boxOf(item);
        const x = item.element.xfrm;
        // Flip and turn act on the slide, also inside a turned or mirrored group.
        const parent = parentFrame(item.parentMatrix);
        if (command.type === 'rotateObjects')
          return toFile(item, { ...box, rot: rotated(box.rot, command.degrees, parent) });
        const f = flipped({ rot: box.rot, flipH: x.flipH, flipV: x.flipV }, command.axis, parent);
        return { ...toFile(item, { ...box, rot: f.rot }), flipH: f.flipH, flipV: f.flipV };
      });
      refuseUnchanged(batch, ctx, slidePart(ctx), out);
      writeTransforms(batch, ctx, slidePart(ctx), out);
      return { batch, effect: { label: command.type } };
    }
    case 'setLockAspectRatio': {
      setLockAspectRatio(batch, slidePart(ctx), selectedIds(ctx.selection), command.lock);
      return { batch, effect: { label: 'lockAspectRatio' } };
    }
    case 'setShapesVisible': {
      const part = slidePart(ctx);
      setShapesVisible(batch, part, command.nodeIds, command.visible);
      // A hidden element leaves the selection, and so does everything inside a hidden group:
      // it no longer draws, so it cannot be edited.
      const hidden = new Set(command.visible ? [] : command.nodeIds);
      const inHidden = (id: string): boolean => {
        for (
          let el = batch.element(part, id);
          el.kind !== 'spTree';
          el = parentOf(batch, part, el.id)
        )
          if (hidden.has(el.id)) return true;
        return false;
      };
      const kept = selectedIds(ctx.selection).filter((id) => !inHidden(id));
      const selection =
        kept.length === selectedIds(ctx.selection).length
          ? undefined
          : kept.length
            ? shapesSelection(kept, scopeOf(ctx.selection))
            : ({ kind: 'none' } as const);
      return {
        batch,
        effect: { label: 'shapesVisible', ...(selection ? { selection } : {}) },
      };
    }
    case 'renameShape': {
      renameShape(batch, slidePart(ctx), command.nodeId, command.name);
      return { batch, effect: { label: 'renameShape' } };
    }
    case 'toggleFormatPainter':
    case 'copyFormat':
    case 'pasteFormat':
    case 'toggleCrop':
      // The painter and crop mode hold editor state: the facade runs them (format-painter.ts,
      // crop-mode.ts).
      throw new CommandRefusal('errors.unsupported', command.type);
    default:
      return compileTable(ctx, batch, command);
  }
}

/**
 * The `Editor.can` cache key of a command without bytes. NaN and Infinity serialize as null in
 * JSON, so they are marked, and strings that start with `#` are escaped, so `NaN` and `"#NaN"`
 * never share a key. @internal
 */
export function commandCacheKey(command: EditorCommand): string {
  return JSON.stringify(command, (_, v: unknown) =>
    (typeof v === 'number' && !isFinite(v)) || (typeof v === 'string' && v[0] === '#')
      ? `#${String(v)}`
      : v
  );
}
