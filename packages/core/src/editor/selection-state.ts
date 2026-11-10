// Selection state for the snapshot: the text format, the element geometry and style, the
// selected text, and the color names the chrome shows. Pure functions over a slide layout.

import type { Autofit, ResolvedFill, ResolvedParagraph, RunProps } from '../cascade/index.ts';
import type { Element } from '../store/index.ts';
import { autofitSizePt, type SlideLayout } from '../layout/index.ts';
import type {
  EditorSelection,
  GradientFill,
  FontFormatRead,
  LineEndType,
  ParagraphFormatRead,
  SelectionFormat,
  SelectionGeometry,
  SlideBackgroundState,
  PatternFill,
  PatternPreset,
  TextFitting,
  TextWarpPreset,
  VerticalAlign,
} from '../contracts/index.ts';
import { selectedIds, type CommandContext } from './commands.ts';
import { boxOf, findItem, toFile } from '../automation/commands/placement.ts';
import {
  addressOf,
  cellItem,
  cellRangeOf,
  expandLayoutRange,
  formatItemOf,
  type ShapeItem,
  tableItem,
  textItemOf,
} from './table-cells.ts';
import { itemById } from './surface.ts';
import { comparePoints } from '../automation/commands/text-model.ts';
import { PATTERN_PRESETS } from '../automation/commands/fill-commands.ts';
import { commandWordRange } from '../automation/commands/command-word.ts';
import type { FormatKeyTarget } from './format-keys.ts';
import { effectsStateOf, shapeStyleOf } from '../automation/commands/effect-commands.ts';
import { aspectLocked } from '../automation/commands/arrange-commands.ts';
import { paragraphFormatRead } from '../automation/commands/paragraph-format.ts';
import { fontFormatRead } from '../automation/commands/font-format.ts';
import { bulletFormatMixedOf, bulletFormatOf } from './bullet-format.ts';
import { textDirectionOf } from './text-body-commands.ts';

const EMPTY_FORMAT = null;

/** The standard picker families (DOCX Editor's set plus common presentation fonts). */
export const STANDARD_FONT_FAMILIES: readonly string[] = [
  'Aptos',
  'Arial',
  'Calibri',
  'Calibri Light',
  'Cambria',
  'Century Gothic',
  'Consolas',
  'Courier New',
  'Garamond',
  'Georgia',
  'Helvetica',
  'Open Sans',
  'Roboto',
  'Times New Roman',
  'Verdana',
];
export { FONT_NAME } from '../automation/commands/font-format.ts';

/** The plain text of a text selection (line breaks as newlines). */
export function selectedText(layout: SlideLayout | undefined, sel: EditorSelection): string {
  if (sel.kind !== 'text' || !layout) return '';
  const item = textItemOf(layout, sel);
  if (!item?.element.text) return '';
  const [from, to] =
    comparePoints(sel.anchor, sel.focus) <= 0 ? [sel.anchor, sel.focus] : [sel.focus, sel.anchor];
  const out: string[] = [];
  item.element.text.paragraphs.forEach((p, i) => {
    if (i < from.paragraph || i > to.paragraph) return;
    const text = p.runs.map((r) => (r.kind === 'break' ? '\n' : r.text)).join('');
    out.push(
      text.slice(i === from.paragraph ? from.offset : 0, i === to.paragraph ? to.offset : undefined)
    );
  });
  return out.join('\n');
}

/**
 * The text Change Case acts on: the selected text, or at a caret the word that the command
 * reaches (`commandWordRange`). Shift+F3 reads it to pick the next case.
 */
export function caseTargetText(layout: SlideLayout | undefined, sel: EditorSelection): string {
  const text = selectedText(layout, sel);
  if (text || sel.kind !== 'text' || !layout) return text;
  const p = textItemOf(layout, sel)?.element.text?.paragraphs[sel.focus.paragraph];
  const plain = p?.runs.map((r) => (r.kind === 'break' ? '\n' : r.text)).join('') ?? '';
  const word = commandWordRange(plain, sel.focus.offset, new Set());
  return word ? plain.slice(word.start, word.end) : '';
}

/** Whether a layout item holds text: a shape's text body, a table, or a group with either. */
function hasText(item: ReturnType<typeof itemById>): boolean {
  if (item?.kind === 'group') return item.children.some(hasText);
  return (
    (item?.kind === 'shape' && !!item.element.text) || (item?.kind === 'frame' && !!item.table)
  );
}

/** Where the format keys act (format-keys.ts): text, selected shapes, and the case target. */
export function formatKeyTarget(
  layout: SlideLayout | undefined,
  sel: EditorSelection
): FormatKeyTarget {
  return {
    inText: sel.kind === 'text',
    // Only a shape with text takes the paragraph and font keys; a picture leaves them to the browser.
    shapes: sel.kind === 'shapes' && sel.nodeIds.some((id) => hasText(itemById(layout, id))),
    selectedText: () => caseTargetText(layout, sel),
  };
}

/** A text body's stored `a:normAutofit/@fontScale` (1 when the text is not shrunk). */
function fontScaleOf(body: { readonly autofit: Autofit }): number {
  return body.autofit.kind === 'normal' ? body.autofit.fontScale : 1;
}

/**
 * The size every text run of `items` is drawn at, under each body's own font scale; null when
 * they draw at different sizes (shapes shrunk by different scales), undefined with no text run.
 */
function drawnSize(items: readonly ShapeItem[]): number | null | undefined {
  let size: number | undefined;
  for (const item of items) {
    const text = item.element.text;
    if (!text) continue;
    const scale = fontScaleOf(text.body);
    for (const p of text.paragraphs)
      for (const r of p.runs) {
        if (r.kind === 'break' || r.text.length === 0) continue;
        const drawn = autofitSizePt(r.props.size / 100, scale);
        if (size === undefined) size = drawn;
        else if (drawn !== size) return null;
      }
  }
  return size;
}

/** The shapes' common drawn size; with no text run in them, the paragraph mark's (`own`). */
const renderedOf = (shapes: number | null | undefined, own: number): number | null =>
  shapes === undefined ? own : shapes;

/** `a:bodyPr/@anchor` as the chrome's three choices (`just` and `dist` read as top). */
function verticalOf(anchor: string): VerticalAlign {
  return anchor === 'ctr' || anchor === 'b' ? anchor : 't';
}

/** A resolved line end as the chrome's types. */
function endOf(end: { type: string } | undefined): LineEndType {
  const t = end?.type;
  return t === 'triangle' || t === 'arrow' || t === 'stealth' || t === 'oval' || t === 'diamond'
    ? t
    : 'none';
}

/** The `link` of a run with a link: its URL, or the `#` form of a slide link. */
function linkFieldOf(
  rId: string | undefined,
  action: string | undefined,
  linkOf: ((rId: string | undefined, action: string | undefined) => string | undefined) | undefined
): { link?: string } {
  if (!rId && !action) return {};
  const link = linkOf?.(rId, action);
  return link ? { link } : {};
}

export function formatAt(
  layout: SlideLayout | undefined,
  sel: EditorSelection,
  linkOf?: (rId: string | undefined, action: string | undefined) => string | undefined
): SelectionFormat | null {
  const ids = selectedIds(sel);
  if (!layout || ids.length === 0) return EMPTY_FORMAT;
  // A table cell reads as a text item; a selected table or cell range reads its first cell. Of
  // selected shapes, the first one with text: a shape without text first in the selection does
  // not hide the text shapes after it.
  const item =
    sel.kind === 'shapes'
      ? sel.nodeIds
          .map((id) => formatItemOf(layout, { ...sel, nodeIds: [id] }))
          .find((it) => it?.element.text)
      : formatItemOf(layout, sel);
  if (!item?.element.text) return EMPTY_FORMAT;
  const paragraphs = item.element.text.paragraphs;
  const pi = sel.kind === 'text' ? Math.min(sel.focus.paragraph, sel.anchor.paragraph) : 0;
  const p = paragraphs[pi] ?? paragraphs[0];
  if (!p) return EMPTY_FORMAT;
  let props = p.endProps;
  if (sel.kind === 'text') {
    const [from, to] =
      comparePoints(sel.anchor, sel.focus) <= 0 ? [sel.anchor, sel.focus] : [sel.focus, sel.anchor];
    const caret = comparePoints(from, to) === 0;
    const at = caret ? Math.max(0, from.offset - 1) : from.offset;
    // At a caret, an empty run there carries the format armed for the next typed text.
    const armed = caret ? armedRunAt(p.runs, from.offset) : undefined;
    let pos = 0;
    if (armed) props = armed.props;
    else
      for (const r of p.runs) {
        const len = r.kind === 'break' ? 1 : r.text.length;
        if (r.kind !== 'break' && at >= pos && at < pos + Math.max(1, len)) {
          props = r.props;
          break;
        }
        pos += len;
        if (r.kind !== 'break') props = r.props;
      }
  } else props = p.runs.find((r) => r.kind !== 'break')?.props ?? p.endProps;
  const color = props.fill.type === 'solid' ? props.fill.color : undefined;
  const hex = color
    ? `#${[color.r, color.g, color.b].map((v) => Math.round(v).toString(16).padStart(2, '0')).join('')}`
    : undefined;
  return {
    bold: props.bold,
    italic: props.italic,
    underline: props.underline !== 'none',
    strike: props.strike !== 'noStrike',
    fontSize: props.size / 100,
    renderedFontSize: renderedOf(
      sel.kind === 'text' ? undefined : drawnSize(bodyItems(layout, ids, sel)),
      autofitSizePt(props.size / 100, fontScaleOf(item.element.text.body))
    ),
    fontFamily: props.latin,
    ...(props.latinTheme ? { fontFamilyTheme: props.latinTheme } : {}),
    ...linkFieldOf(props.hyperlink, props.hyperlinkAction, linkOf),
    ...(hex ? { color: hex } : {}),
    align: p.align === 'ctr' || p.align === 'r' || p.align === 'just' ? p.align : 'l',
    verticalAlign: verticalOf(item.element.text.body.anchor),
    bullets: p.bullet?.type === 'char' || p.bullet?.type === 'picture',
    numbering: p.bullet?.type === 'autonum',
    ...paragraphField(layout, ids, paragraphs, sel),
    superscript: props.baseline > 0,
    subscript: props.baseline < 0,
    ...fontField(layout, ids, paragraphs, sel, props),
    ...bulletField(selectedParagraphs(layout, ids, paragraphs, sel)),
    ...bodyFields(layout, ids),
  };
}

/**
 * The bullet format of the selected paragraphs: absent when they differ, so no gallery row is
 * checked, with what they share, so the dialog opens on their list's tab.
 */
function bulletField(
  selected: readonly ResolvedParagraph[]
): Pick<SelectionFormat, 'bulletFormat' | 'bulletFormatMixed'> {
  const reads = selected.map(bulletFormatOf);
  const key = JSON.stringify(reads[0] ?? null);
  if (reads.every((r) => JSON.stringify(r) === key)) return { bulletFormat: reads[0] ?? null };
  return { bulletFormatMixed: bulletFormatMixedOf(reads) };
}

/**
 * Text Direction and Columns over every selected shape with text. A field is absent when the
 * shapes differ, so no menu row is checked.
 */
function bodyFields(
  layout: SlideLayout,
  ids: readonly string[]
): Pick<SelectionFormat, 'textDirection' | 'columns'> {
  const bodies: { vert: string; numCol: number; spcCol: number }[] = [];
  for (const id of ids) {
    const it = findItem(layout.items, id);
    if (it?.kind === 'shape' && it.element.text) bodies.push(it.element.text.body);
  }
  const first = bodies[0];
  if (!first) return {};
  const direction = textDirectionOf(first.vert);
  const sameDirection = bodies.every((b) => textDirectionOf(b.vert) === direction);
  const sameColumns = bodies.every((b) => b.numCol === first.numCol && b.spcCol === first.spcCol);
  return {
    ...(sameDirection ? { textDirection: direction } : {}),
    ...(sameColumns ? { columns: { count: first.numCol, gapEmu: first.spcCol } } : {}),
  };
}

/** The formatting state the toggles read (`CommandContext.active`), from a format read. */
export function activeOf(f: SelectionFormat | null): CommandContext['active'] {
  return {
    bold: f?.bold ?? false,
    italic: f?.italic ?? false,
    underline: f?.underline ?? false,
    strike: f?.strike ?? false,
    bullets: f?.bullets ?? false,
    numbering: f?.numbering ?? false,
    // A step starts from the stored size, so Increase always asks for a larger size; the
    // refit may then draw it at the same size (shrunk text keeps filling its box).
    ...(f?.fontSize ? { fontSize: f.fontSize } : {}),
    ...(f?.renderedFontSize ? { renderedFontSize: f.renderedFontSize } : {}),
    ...(f?.font?.disagrees.fontSizePt ? { fontSizeMixed: true } : {}),
    superscript: f?.superscript ?? false,
    subscript: f?.subscript ?? false,
  };
}

/** An empty text run that starts at `offset`: a format armed for the text typed there. */
function armedRunAt(runs: ResolvedParagraph['runs'], offset: number) {
  let pos = 0;
  for (const r of runs) {
    if (r.kind === 'text' && r.text.length === 0 && pos === offset) return r;
    pos += r.kind === 'break' ? 1 : r.text.length;
  }
  return undefined;
}

/**
 * The selected text's font: the runs a non-empty text range covers, the run at a caret (or the
 * format armed there), or every run of every selected shape with text, the runs
 * `setFontFormat` writes.
 */
function fontField(
  layout: SlideLayout,
  ids: readonly string[],
  paragraphs: readonly ResolvedParagraph[],
  sel: EditorSelection,
  caret: RunProps
): { font?: FontFormatRead } {
  const runs: RunProps[] = [];
  const textRuns = (p: ResolvedParagraph, from = 0, to = Number.MAX_SAFE_INTEGER) => {
    let pos = 0;
    for (const r of p.runs) {
      const len = r.kind === 'break' ? 1 : r.text.length;
      if (r.kind !== 'break' && len > 0 && pos < to && pos + len > from) runs.push(r.props);
      pos += len;
    }
  };
  if (sel.kind === 'text' && comparePoints(sel.anchor, sel.focus) !== 0) {
    const [from, to] =
      comparePoints(sel.anchor, sel.focus) < 0 ? [sel.anchor, sel.focus] : [sel.focus, sel.anchor];
    for (let pi = from.paragraph; pi <= to.paragraph; pi += 1) {
      const p = paragraphs[pi];
      if (!p) continue;
      textRuns(
        p,
        pi === from.paragraph ? from.offset : 0,
        pi === to.paragraph ? to.offset : Number.MAX_SAFE_INTEGER
      );
    }
  } else if (sel.kind !== 'text')
    for (const item of bodyItems(layout, ids, sel))
      for (const p of item.element.text?.paragraphs ?? []) textRuns(p);
  const read = fontFormatRead(runs.length ? runs : [caret]);
  return read ? { font: read } : {};
}

/**
 * The selected paragraphs' format: the text selection's paragraphs, or every paragraph of
 * every selected shape with text, the paragraphs `setParagraphFormat` writes.
 */
function paragraphField(
  layout: SlideLayout,
  ids: readonly string[],
  paragraphs: readonly ResolvedParagraph[],
  sel: EditorSelection
): { paragraph?: ParagraphFormatRead } {
  const read = paragraphFormatRead(selectedParagraphs(layout, ids, paragraphs, sel));
  return read ? { paragraph: read } : {};
}

/**
 * The paragraphs a paragraph command writes: the text selection's paragraphs, or every
 * paragraph of every selected shape with text.
 */
function selectedParagraphs(
  layout: SlideLayout,
  ids: readonly string[],
  paragraphs: readonly ResolvedParagraph[],
  sel: EditorSelection
): readonly ResolvedParagraph[] {
  if (sel.kind === 'text') {
    const from = Math.min(sel.anchor.paragraph, sel.focus.paragraph);
    const to = Math.max(sel.anchor.paragraph, sel.focus.paragraph);
    return paragraphs.slice(from, to + 1);
  }
  return bodyItems(layout, ids, sel).flatMap((item) => item.element.text?.paragraphs ?? []);
}

/**
 * The text items a command writes for a selection that is not a text range, as `textTargets`
 * finds them: each selected shape with text, every cell of a selected table, and the cells of
 * a cell range.
 */
function bodyItems(
  layout: SlideLayout,
  ids: readonly string[],
  sel: EditorSelection
): readonly ShapeItem[] {
  const out: ShapeItem[] = [];
  for (const id of ids) {
    const frame = tableItem(layout, id);
    if (frame) {
      const r = sel.kind === 'cells' ? expandLayoutRange(frame.table, cellRangeOf(sel)) : null;
      for (const cl of frame.table.cells) {
        const a = addressOf(cl);
        if (r && (a.row < r.r0 || a.row > r.r1 || a.column < r.c0 || a.column > r.c1)) continue;
        const cell = cellItem(frame, a);
        if (cell?.element.text) out.push(cell);
      }
      continue;
    }
    const item = findItem(layout.items, id);
    if (item?.kind === 'shape' && item.element.text) out.push(item);
  }
  return out;
}

export const hexOf = (c: { r: number; g: number; b: number }): string =>
  `#${[c.r, c.g, c.b]
    .map((v) =>
      Math.round(Math.max(0, Math.min(255, v)))
        .toString(16)
        .padStart(2, '0')
    )
    .join('')
    .toUpperCase()}`;

/** A fill as the chrome shows it: `#RRGGBB`, `none`, `gradient`, or `other`. */
export function fillName(fill: ResolvedFill | undefined): string {
  if (!fill || fill.type === 'none') return 'none';
  if (fill.type === 'linear' || (fill.type === 'path' && fill.path === 'circle')) return 'gradient';
  return fill.type === 'solid' ? hexOf(fill.color) : 'other';
}

/**
 * A shape fill as the chrome shows it: a `fillName`, or `pattern` or `picture`. A pattern with
 * a preset name outside ECMA-376 is `other`: the pattern dialog cannot show or keep it.
 */
export function shapeFillName(fill: ResolvedFill | undefined): string {
  if (fill?.type === 'pattern') return patternOf(fill) ? 'pattern' : 'other';
  return fill?.type === 'picture' ? fill.type : fillName(fill);
}

/**
 * A pattern fill as the chrome edits it, with its colors as `#RRGGBB`. Undefined for any other
 * fill, and for a preset name outside ECMA-376.
 */
export function patternOf(fill: ResolvedFill): PatternFill | undefined {
  if (fill.type !== 'pattern' || !PATTERN_PRESETS.has(fill.preset)) return undefined;
  return {
    preset: fill.preset as PatternPreset,
    foreground: hexOf(fill.fg),
    background: hexOf(fill.bg),
  };
}

/** A linear or circular gradient as the chrome edits it. */
export function gradientOf(fill: ResolvedFill): GradientFill | undefined {
  if (fill.type !== 'linear' && !(fill.type === 'path' && fill.path === 'circle')) return undefined;
  return {
    kind: fill.type === 'linear' ? 'linear' : 'radial',
    angle: fill.type === 'linear' ? fill.angle : 90,
    stops: fill.stops.map((s) => ({ pos: s.pos, color: hexOf(s.color) })),
    rotWithShape: fill.rotWithShape,
  };
}

/**
 * The current slide's background as the chrome shows it. The function it returns keeps the
 * object it returned last while the background does not change (the snapshot contract).
 */
export function createBackgroundSource(): (lay: SlideLayout | undefined) => SlideBackgroundState {
  let last: SlideBackgroundState = { kind: 'none', inherited: true };
  return (lay) => {
    if (!lay) return last;
    const fill = lay.slide.background;
    const gradient = gradientOf(fill);
    const pattern = patternOf(fill);
    const next: SlideBackgroundState = {
      kind:
        fill.type === 'solid'
          ? 'solid'
          : gradient
            ? 'gradient'
            : pattern
              ? 'pattern'
              : fill.type === 'picture'
                ? 'picture'
                : fill.type === 'none'
                  ? 'none'
                  : 'other',
      ...(fill.type === 'solid' ? { color: hexOf(fill.color) } : {}),
      ...(gradient ? { gradient } : {}),
      ...(pattern ? { pattern } : {}),
      ...(fill.type === 'picture' && fill.mode === 'tile' ? { tile: true } : {}),
      inherited: lay.slide.backgroundPart !== lay.slide.partName,
    };
    if (JSON.stringify(next) !== JSON.stringify(last)) last = next;
    return last;
  };
}

/**
 * The selected element's geometry and style. `source` returns the element's XML, for what the
 * layout does not carry: the aspect lock and the `p:style` of the Shape Styles gallery.
 */
export function geometryOf(
  layout: SlideLayout | undefined,
  sel: EditorSelection,
  source?: (nodeId: string) => Element | undefined
): SelectionGeometry | null {
  const ids = selectedIds(sel);
  const item = layout && ids.length === 1 ? itemById(layout, ids[0]) : undefined;
  if (!item) return null;
  const e = item.element;
  const kind =
    e.kind === 'shape'
      ? 'shape'
      : e.kind === 'connector'
        ? 'connector'
        : e.kind === 'picture'
          ? 'picture'
          : e.kind === 'group'
            ? 'group'
            : 'graphicFrame';
  // Position and size in the file's coordinates, so the panel can write them back.
  const box = toFile(item, boxOf(item));
  const line = 'line' in e ? e.line : undefined;
  const lineVisible = !!line && line.fill.type !== 'none';
  const autofit = e.kind === 'shape' ? e.text?.body.autofit.kind : undefined;
  const el = source?.(e.nodeId);
  const shapeStyle = e.kind === 'shape' ? shapeStyleOf(el) : undefined;
  const textFitting: TextFitting | undefined = !autofit
    ? undefined
    : autofit === 'normal'
      ? 'shrink'
      : autofit === 'shape'
        ? 'shape'
        : 'none';
  return {
    kind,
    name: e.name,
    nodeId: e.nodeId,
    x: box.x,
    y: box.y,
    cx: box.cx,
    cy: box.cy,
    rot: box.rot,
    ...(e.kind === 'shape' ? { fill: shapeFillName(e.fill) } : {}),
    ...(e.kind === 'shape' && gradientOf(e.fill) ? { gradient: gradientOf(e.fill) } : {}),
    ...(e.kind === 'shape' && patternOf(e.fill) ? { pattern: patternOf(e.fill) } : {}),
    ...(e.kind === 'shape' || e.kind === 'connector' || e.kind === 'picture'
      ? {
          lineColor: lineVisible ? fillName(line.fill) : 'none',
          ...(line ? { lineWidth: line.width / 12_700 } : {}),
          ...(lineVisible ? { lineDash: line.dash } : {}),
          ...(e.kind === 'connector' ||
          (e.kind === 'shape' &&
            e.geometry.kind === 'preset' &&
            /^(line|lineInv|straightConnector1|bentConnector[2-5]|curvedConnector[2-5]|arc)$/.test(
              e.geometry.preset
            ))
            ? { lineHead: endOf(line?.head), lineTail: endOf(line?.tail) }
            : {}),
        }
      : {}),
    ...(textFitting ? { textFitting } : {}),
    ...(e.kind === 'shape' && e.text
      ? { textWarp: (e.text.body.warp?.preset ?? 'textNoShape') as TextWarpPreset }
      : {}),
    flipH: e.xfrm.flipH,
    flipV: e.xfrm.flipV,
    lockAspectRatio: aspectLocked(el),
    ...(e.kind === 'shape' || e.kind === 'connector' || e.kind === 'picture'
      ? { effects: effectsStateOf(e.effects, el) }
      : {}),
    ...(shapeStyle ? { shapeStyle } : {}),
  };
}
