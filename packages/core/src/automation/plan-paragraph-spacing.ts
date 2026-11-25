// Paragraph spacing and indents of the automation `readParagraph` and `setParagraph`
// operations: line spacing (`a:lnSpc`, ECMA-376 Part 1, 21.1.2.2.5), space before and after
// (`a:spcBef`, 21.1.2.2.10; `a:spcAft`, 21.1.2.2.9), each of type CT_TextSpacing with one
// `a:spcPct` (21.1.2.2.11) or `a:spcPts` (21.1.2.2.12), and the `marL` (ST_TextMargin) and
// `indent` (ST_TextIndent) attributes of `a:pPr` (CT_TextParagraphProperties, 21.1.2.2.7).
//
// The `lineSpacing` field is in points, with a rule: `exactly` is `a:spcPts`, and under
// `multiple` (`a:spcPct`) 12 points is one line. Writes go through the Paragraph dialog's
// `applyParagraphFormat`, so a value of the same kind changes in place.

import { textSpacing, type ResolvedParagraph, type Spacing } from '../cascade/index.ts';
import type { ParagraphFormatCommand } from '../contracts/neutral.ts';
import { child, intAttr, type Element } from '../store/index.ts';
import { applyParagraphFormat } from './commands/paragraph-format.ts';
import type { OpBatch } from './commands/op-batch.ts';
import type { AutomationParagraphWrite } from './operations.ts';
import type { AutomationParagraphRead } from './protocol.ts';
import { badArgument, numberIn, toEmu, toPoints } from './targets.ts';

/** The value all items share, or null. */
function shared<T>(items: readonly T[]): T | null {
  if (items.length === 0) return null;
  const first = items[0]!;
  return items.every((v) => v === first) ? first : null;
}

/**
 * Single spacing in points, for the `lineSpacing` field: 12 points is one line, so a multiple
 * of 1.5 reads 18. This is the convention of the DOCX Editor API for the same member.
 */
const POINTS_PER_LINE = 12;

/** A spacing as the points of the `lineSpacing` field. */
const spacingPoints = (s: Spacing): number =>
  'pts' in s
    ? s.pts / 100
    : Math.round((Math.round(s.pct * 100_000) * POINTS_PER_LINE) / 1000) / 100;

/** A spacing as lines, rounded to the precision of `a:spcPct` (thousandths of a percent). */
const spacingLines = (s: { readonly pct: number }): number => Math.round(s.pct * 100_000) / 100_000;

/** The spacing and indents of one paragraph; a field the paragraph does not hold is absent. */
export interface SpacingOf {
  readonly lineSpacing?: Spacing;
  readonly spaceBefore?: Spacing;
  readonly spaceAfter?: Spacing;
  readonly marL?: number;
  readonly indent?: number;
}

/**
 * The paragraph's own spacing children (`a:lnSpc`, `a:spcBef`, `a:spcAft`, CT_TextSpacing,
 * ECMA-376 Part 1, 21.1.2.2.5, 21.1.2.2.10, 21.1.2.2.9) and indents (`marL`, `indent`).
 */
export function ownSpacing(p: Element): SpacingOf {
  const pPr = child(p, 'a:pPr');
  const spacing = (local: string): Spacing | undefined => textSpacing(child(pPr, `a:${local}`));
  const emu = (name: string): number | undefined => {
    const v = intAttr(pPr, name);
    return v === undefined ? undefined : Math.max(-51_206_400, Math.min(v, 51_206_400));
  };
  return {
    lineSpacing: spacing('lnSpc'),
    spaceBefore: spacing('spcBef'),
    spaceAfter: spacing('spcAft'),
    marL: emu('marL'),
    indent: emu('indent'),
  };
}

/**
 * The spacing fields of a read: each value the paragraphs share, or null. A space in lines
 * reads null in points and a space in points reads null in lines, except 0, which is 0 in both.
 */
export function spacingRead(
  list: readonly SpacingOf[]
): Pick<
  AutomationParagraphRead,
  | 'lineSpacing'
  | 'lineSpacingRule'
  | 'spaceBefore'
  | 'spaceAfter'
  | 'lineUnitBefore'
  | 'lineUnitAfter'
  | 'leftIndent'
  | 'firstLineIndent'
> {
  const each = <T>(of: (p: SpacingOf) => T | null | undefined): T | null =>
    shared(list.map((p) => of(p) ?? null));
  const points = (s: Spacing | undefined): number | null =>
    !s ? null : 'pts' in s ? s.pts / 100 : s.pct === 0 ? 0 : null;
  const lines = (s: Spacing | undefined): number | null =>
    !s ? null : 'pct' in s ? spacingLines(s) : s.pts === 0 ? 0 : null;
  const emu = (v: number | undefined): number | null => (v === undefined ? null : toPoints(v));
  return {
    lineSpacing: each((p) => p.lineSpacing && spacingPoints(p.lineSpacing)),
    lineSpacingRule: each((p) =>
      p.lineSpacing ? ('pts' in p.lineSpacing ? 'exactly' : 'multiple') : null
    ),
    spaceBefore: each((p) => points(p.spaceBefore)),
    spaceAfter: each((p) => points(p.spaceAfter)),
    lineUnitBefore: each((p) => lines(p.spaceBefore)),
    lineUnitAfter: each((p) => lines(p.spaceAfter)),
    leftIndent: each((p) => emu(p.marL)),
    firstLineIndent: each((p) => emu(p.indent)),
  };
}

/** The checked spacing fields of a paragraph write. */
export interface SpacingWrite {
  /** The fields that do not depend on the paragraph, as a Paragraph dialog command. */
  readonly command?: ParagraphFormatCommand;
  /** The line spacing for a paragraph with `inForce`, or undefined to leave it. */
  readonly lineSpacing?: (inForce: Spacing) => ParagraphFormatCommand['lineSpacing'];
}

/**
 * Check the spacing and indent fields of a write before any change, so a bad value refuses the
 * whole write. The limits are the schema's: `ST_TextSpacingPoint` (ECMA-376 Part 1, 20.1.10.78)
 * is 0 to 158400 hundredths of a point; `a:spcPct` is at most 13200000 thousandths of a percent
 * (132 lines), as `ST_TextSpacingPercent` (ECMA-376 Part 4, 17.2.2.11) allows; `ST_TextMargin`
 * (Part 1, 20.1.10.72) is 0 to 51206400 EMU, and `ST_TextIndent` (Part 1, 20.1.10.70) is
 * -51206400 to 51206400 EMU.
 */
export function spacingWrite(format: AutomationParagraphWrite): SpacingWrite {
  const given = (name: keyof AutomationParagraphWrite): boolean => format[name] !== undefined;
  const points = (name: keyof AutomationParagraphWrite, min: number, max: number) =>
    given(name) ? numberIn(format[name], name, min, max) : undefined;
  // One `a:spcBef` holds points or lines, never both.
  if (given('spaceBefore') && given('lineUnitBefore')) badArgument('lineUnitBefore');
  if (given('spaceAfter') && given('lineUnitAfter')) badArgument('lineUnitAfter');
  const rule = format.lineSpacingRule;
  if (rule !== undefined && rule !== 'multiple' && rule !== 'exactly')
    badArgument('lineSpacingRule');
  // 1584 points: the most `a:spcPts` holds, and 132 lines of 12 points as a multiple.
  const value = points('lineSpacing', 0.01, 1584);
  const fields: [keyof ParagraphFormatCommand, number | undefined][] = [
    ['spaceBeforePt', points('spaceBefore', 0, 1584)],
    ['spaceAfterPt', points('spaceAfter', 0, 1584)],
    ['spaceBeforeLines', points('lineUnitBefore', 0, 132)],
    ['spaceAfterLines', points('lineUnitAfter', 0, 132)],
    ['indentLeftEmu', mapDefined(points('leftIndent', 0, 4032), toEmu)],
    ['indentFirstLineEmu', mapDefined(points('firstLineIndent', -4032, 4032), toEmu)],
  ];
  const set = fields.filter(([, v]) => v !== undefined);
  const command = set.length ? (Object.fromEntries(set) as ParagraphFormatCommand) : undefined;
  if (rule === undefined && value === undefined) return { command };
  return {
    command,
    lineSpacing: (inForce) => {
      let s = inForce;
      // The rule keeps the points in force: 1.5 lines become exactly 18 points, and back.
      if (rule === 'exactly' && 'pct' in s) s = { pts: spacingPoints(s) * 100 };
      else if (rule === 'multiple' && 'pts' in s) s = { pct: s.pts / 100 / POINTS_PER_LINE };
      // A rule the paragraph already has, with no value, adds no formatting.
      else if (value === undefined) return undefined;
      if (value !== undefined)
        s = 'pts' in s ? { pts: value * 100 } : { pct: value / POINTS_PER_LINE };
      // A converted spacing keeps the smallest line spacing a write takes, 0.01 points.
      return 'pts' in s
        ? { rule: 'exact', value: Math.max(s.pts / 100, 0.01) }
        : { rule: 'multiple', value: Math.max(s.pct, 0.01 / POINTS_PER_LINE) };
    },
  };
}

const mapDefined = <T, U>(v: T | undefined, f: (v: T) => U): U | undefined =>
  v === undefined ? undefined : f(v);

/**
 * Write a checked spacing on one paragraph. `before` holds the paragraphs as the cascade
 * resolves them before the write, for the line spacing in force; a paragraph it lacks (on a
 * layout or a master) uses its own spacing, or single spacing.
 */
export function applySpacing(
  batch: OpBatch,
  part: string,
  pId: string,
  pPrId: string,
  spacing: SpacingWrite,
  before: ReadonlyMap<string, ResolvedParagraph> | undefined
): void {
  if (!spacing.lineSpacing && !spacing.command) return;
  const inForce = before?.get(pId)?.lineSpacing ??
    ownSpacing(batch.element(part, pId)).lineSpacing ?? { pct: 1 };
  const lineSpacing = spacing.lineSpacing?.(inForce);
  applyParagraphFormat(batch, part, pPrId, {
    ...spacing.command,
    ...(lineSpacing ? { lineSpacing } : {}),
  });
}
