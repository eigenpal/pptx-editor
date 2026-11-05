// The Paragraph dialog (Format > Paragraph) as store ops: alignment, indents, spacing before and
// after, and line spacing, on every selected paragraph, in one transaction.
//
// ECMA-376 Part 1, DrawingML `CT_TextParagraphProperties` (`a:pPr`): the `algn`
// (`ST_TextAlignType`), `marL` (`ST_TextMargin`), and `indent` (`ST_TextIndent`) attributes, and
// the `a:lnSpc`, `a:spcBef`, and `a:spcAft` children (`CT_TextSpacing`: one `a:spcPct` or
// `a:spcPts`), which come first in the schema's child order. `a:spcPts` is in hundredths of a
// point (`ST_TextSpacingPoint`, 0 to 158400); `a:spcPct` in thousandths of a percent
// (`ST_TextSpacingPercent`, 0 to 13200000).
//
// A value the command leaves out keeps the paragraph's own value; `null` removes it, so the
// paragraph inherits again (layout, master, list style), through the text style cascade.

import { attr, child, elements, NS } from '../../store/index.ts';
import type { ResolvedParagraph, Spacing } from '../../cascade/index.ts';
import type {
  ParagraphAlignment,
  ParagraphFormatCommand,
  ParagraphFormatRead,
  ParagraphLineSpacing,
} from '../../contracts/neutral.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { aSpec } from './text-model.ts';

/** The dialog's alignments and their `algn` values. */
const ALGN: Readonly<Record<ParagraphAlignment, string>> = {
  left: 'l',
  center: 'ctr',
  right: 'r',
  justify: 'just',
  distributed: 'dist',
};
/** The widest margin and indent the schema allows, EMU. */
export const MAX_MARGIN = 51_206_400;
/** `ST_TextSpacingPoint` as points. */
export const MAX_POINTS = 1584;
/** `ST_TextSpacingPercent` as lines (13200000 thousandths of a percent). */
export const MAX_LINES = 132;
/** The widest line spacing multiple. */
export const MAX_MULTIPLE = 9.99;
/** The spacing children, in `CT_TextParagraphProperties` order; they precede all others. */
const SPACING_ORDER = ['lnSpc', 'spcBef', 'spcAft'] as const;
type SpacingChild = (typeof SPACING_ORDER)[number];

const FIELDS: readonly (keyof ParagraphFormatCommand)[] = [
  'alignment',
  'spaceBeforePt',
  'spaceBeforeLines',
  'spaceAfterPt',
  'spaceAfterLines',
  'lineSpacing',
  'indentLeftEmu',
  'indentFirstLineEmu',
];

const finite = (v: unknown): v is number => typeof v === 'number' && Number.isFinite(v);
const inRange = (v: unknown, min: number, max: number): boolean =>
  finite(v) && v >= min && v <= max;

/** `a:spcPts` for points: hundredths of a point. */
const pointsVal = (points: number): number => Math.round(points * 100);
/** `a:spcPct` for lines: thousandths of a percent of one line. */
const linesVal = (lines: number): number => Math.round(lines * 100_000);

function validLineSpacing(s: unknown): s is ParagraphLineSpacing {
  if (typeof s !== 'object' || s === null) return false;
  const { rule, value } = s as { rule?: unknown; value?: unknown };
  // The smallest line spacing is 0.01 (lines or points), so lines never collapse.
  if (rule === 'multiple') return inRange(value, 0.01, MAX_MULTIPLE);
  return rule === 'exact' && inRange(value, 0.01, MAX_POINTS);
}

/**
 * Check a command's fields, or refuse it: an unknown alignment, a margin or spacing the
 * schema does not allow, a space given both in points and in lines, or no field at all. The
 * refusal's detail names the field.
 */
export function checkParagraphFormat(format: ParagraphFormatCommand): void {
  const given = (k: keyof ParagraphFormatCommand): boolean => format[k] !== undefined;
  if (!FIELDS.some(given)) throw new CommandRefusal('errors.unsupported', 'paragraph');
  const bad = (name: keyof ParagraphFormatCommand): never => {
    throw new CommandRefusal('errors.unsupported', name);
  };
  const set = (k: keyof ParagraphFormatCommand): boolean => format[k] != null;
  if (given('alignment') && !Object.hasOwn(ALGN, format.alignment as string)) bad('alignment');
  for (const k of ['spaceBeforePt', 'spaceAfterPt'] as const)
    if (set(k) && !inRange(format[k], 0, MAX_POINTS)) bad(k);
  for (const k of ['spaceBeforeLines', 'spaceAfterLines'] as const)
    if (set(k) && !inRange(format[k], 0, MAX_LINES)) bad(k);
  // One `a:spcBef` holds points or lines, never both.
  if (given('spaceBeforePt') && given('spaceBeforeLines')) bad('spaceBeforeLines');
  if (given('spaceAfterPt') && given('spaceAfterLines')) bad('spaceAfterLines');
  if (set('lineSpacing') && !validLineSpacing(format.lineSpacing)) bad('lineSpacing');
  if (set('indentLeftEmu') && !inRange(format.indentLeftEmu, 0, MAX_MARGIN)) bad('indentLeftEmu');
  if (set('indentFirstLineEmu') && !inRange(format.indentFirstLineEmu, -MAX_MARGIN, MAX_MARGIN))
    bad('indentFirstLineEmu');
}

/** The spacing value element (`a:spcPct` or `a:spcPts`) a command asks for, or null. */
type SpacingValue = { readonly local: 'spcPts' | 'spcPct'; readonly val: number } | null;

function setAttr(batch: OpBatch, part: string, node: string, name: string, value: string | null) {
  batch.push({ op: 'setAttribute', part, node, name: { ns: '', local: name }, value });
}

/**
 * Set (or, with null, remove) one spacing child of `a:pPr`, at its schema position. When the
 * child already holds the same kind of value, only its `val` changes: two collaborators who
 * set the same field then merge as one attribute write, not as two elements.
 */
function setSpacingChild(
  batch: OpBatch,
  part: string,
  pPrId: string,
  local: SpacingChild,
  value: SpacingValue
): void {
  const existing = child(batch.element(part, pPrId), `a:${local}`);
  const inner = existing ? elements(existing).filter((c) => c.ns === NS.a) : [];
  if (value && inner.length === 1 && inner[0].local === value.local) {
    const val = String(value.val);
    if (attr(inner[0], 'val') !== val) setAttr(batch, part, inner[0].id, 'val', val);
    return;
  }
  if (existing) batch.push({ op: 'removeNode', part, node: existing.id });
  if (!value) return;
  const rank = SPACING_ORDER.indexOf(local);
  let index = 0;
  for (const [i, c] of batch.element(part, pPrId).children.entries()) {
    if (c.type !== 'element' || c.ns !== NS.a) continue;
    const r = SPACING_ORDER.indexOf(c.local as SpacingChild);
    if (r >= 0 && r < rank) index = i + 1;
  }
  const spec = aSpec(local, {}, [aSpec(value.local, { val: String(value.val) })]);
  batch.insert({ op: 'insertNode', part, parent: pPrId, index, spec });
}

/** A space command's value element: points when given, lines otherwise. */
function spaceValue(
  points: number | null | undefined,
  lines: number | null | undefined
): SpacingValue | undefined {
  if (points !== undefined)
    return points === null ? null : { local: 'spcPts', val: pointsVal(points) };
  if (lines !== undefined) return lines === null ? null : { local: 'spcPct', val: linesVal(lines) };
  return undefined;
}

/** Write the command's fields on one paragraph's `a:pPr`. */
export function applyParagraphFormat(
  batch: OpBatch,
  part: string,
  pPrId: string,
  format: ParagraphFormatCommand
): void {
  if (format.alignment !== undefined) setAttr(batch, part, pPrId, 'algn', ALGN[format.alignment]);
  const emu = (v: number | null): string | null => (v === null ? null : String(Math.round(v)));
  if (format.indentLeftEmu !== undefined)
    setAttr(batch, part, pPrId, 'marL', emu(format.indentLeftEmu));
  if (format.indentFirstLineEmu !== undefined)
    setAttr(batch, part, pPrId, 'indent', emu(format.indentFirstLineEmu));
  const { lineSpacing } = format;
  if (lineSpacing !== undefined)
    setSpacingChild(
      batch,
      part,
      pPrId,
      'lnSpc',
      lineSpacing === null
        ? null
        : lineSpacing.rule === 'multiple'
          ? { local: 'spcPct', val: linesVal(lineSpacing.value) }
          : { local: 'spcPts', val: pointsVal(lineSpacing.value) }
    );
  const before = spaceValue(format.spaceBeforePt, format.spaceBeforeLines);
  if (before !== undefined) setSpacingChild(batch, part, pPrId, 'spcBef', before);
  const after = spaceValue(format.spaceAfterPt, format.spaceAfterLines);
  if (after !== undefined) setSpacingChild(batch, part, pPrId, 'spcAft', after);
}

/** Whether a command only removes values, so a paragraph with no `a:pPr` needs none. */
export function onlyRemoves(format: ParagraphFormatCommand): boolean {
  return FIELDS.every((k) => format[k] == null);
}

/** `algn` as the dialog's alignment. `justLow` and `thaiDist` read as their nearest choice. */
function alignmentOf(algn: ResolvedParagraph['align']): ParagraphAlignment {
  if (algn === 'ctr') return 'center';
  if (algn === 'r') return 'right';
  if (algn === 'just' || algn === 'justLow') return 'justify';
  if (algn === 'dist' || algn === 'thaiDist') return 'distributed';
  return 'left';
}

const lineSpacingOf = (s: Spacing): ParagraphLineSpacing =>
  'pts' in s ? { rule: 'exact', value: s.pts / 100 } : { rule: 'multiple', value: s.pct };

const sameSpacing = (a: Spacing, b: Spacing): boolean =>
  'pts' in a ? 'pts' in b && a.pts === b.pts : 'pct' in b && a.pct === b.pct;

/**
 * The Paragraph dialog's view of the selected paragraphs: each field holds the value they
 * share, or null when they disagree. Null for no paragraph.
 */
export function paragraphFormatRead(
  paragraphs: readonly ResolvedParagraph[]
): ParagraphFormatRead | null {
  const first = paragraphs[0];
  if (!first) return null;
  const differs = <T>(of: (p: ResolvedParagraph) => T, eq = (a: T, b: T) => a === b): boolean =>
    paragraphs.some((p) => !eq(of(p), of(first)));
  const disagrees = {
    alignment: differs((p) => alignmentOf(p.align)),
    spaceBefore: differs((p) => p.spaceBefore, sameSpacing),
    spaceAfter: differs((p) => p.spaceAfter, sameSpacing),
    lineSpacing: differs((p) => p.lineSpacing, sameSpacing),
    indentLeft: differs((p) => p.marL),
    indentFirstLine: differs((p) => p.indent),
  };
  const pts = (s: Spacing, mixed: boolean): number | null =>
    !mixed && 'pts' in s ? s.pts / 100 : null;
  const lines = (s: Spacing, mixed: boolean): number | null =>
    !mixed && 'pct' in s ? s.pct : null;
  return {
    alignment: disagrees.alignment ? null : alignmentOf(first.align),
    spaceBeforePt: pts(first.spaceBefore, disagrees.spaceBefore),
    spaceBeforeLines: lines(first.spaceBefore, disagrees.spaceBefore),
    spaceAfterPt: pts(first.spaceAfter, disagrees.spaceAfter),
    spaceAfterLines: lines(first.spaceAfter, disagrees.spaceAfter),
    lineSpacing: disagrees.lineSpacing ? null : lineSpacingOf(first.lineSpacing),
    indentLeftEmu: disagrees.indentLeft ? null : first.marL,
    indentFirstLineEmu: disagrees.indentFirstLine ? null : first.indent,
    disagrees,
  };
}

/** Whether two reads show the same thing, for adapters that keep the last read. @internal */
export function sameParagraphFormat(
  a: ParagraphFormatRead | null,
  b: ParagraphFormatRead | null
): boolean {
  return a === b || (a !== null && b !== null && JSON.stringify(a) === JSON.stringify(b));
}

/**
 * The toolbar's Line Spacing value (`setLineSpacing`): a multiple, written in place as the
 * dialog writes it, so two collaborators who change it merge as one attribute write.
 */
export function setLineSpacingMultiple(
  batch: OpBatch,
  part: string,
  pPrId: string,
  multiple: number
): void {
  setSpacingChild(batch, part, pPrId, 'lnSpc', { local: 'spcPct', val: linesVal(multiple) });
}
