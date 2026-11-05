// Table format through automation: the borders and margins of a cell (`a:tcPr`, ECMA-376
// Part 1, 21.1.3.17) and the style options and style of a table (`a:tblPr`, 21.1.3.15). The
// writes are the editor's own (`commands/table-format.ts`), so an outside edge also changes the
// cell across it, as the editor's Borders menu does.

import { child, intAttr } from '../store/index.ts';
import type { TableStyleOption } from '../contracts/neutral.ts';
import { DEFAULT_TABLE_STYLE_ID, tableStyleList } from '../cascade/table-styles.ts';
import {
  clearTableStyle,
  setCellEdge,
  setCellMargins,
  setStyleOptions,
  setTableStyle as setStyleId,
  type CellLine,
} from './commands/table-format.ts';
import type {
  AutomationCellAddress,
  AutomationCellBorderSide,
  AutomationCellMarginsWrite,
  AutomationLineWrite,
  AutomationTableStyleWrite,
} from './operations.ts';
import type {
  AutomationLineRead,
  AutomationRef,
  AutomationTableCellFormatRead,
  AutomationValue,
} from './protocol.ts';
import { lineElementRead, lineRead, resolvedShape } from './plan-reads.ts';
import { COMPOUNDS, DASHES } from './plan-shapes.ts';
import { cellBlock, refit, tableAt } from './plan-tables.ts';
import { colorSpec, type TextPlanContext } from './plan-text.ts';
import { AutomationFailure, badArgument, edit, numberIn, toEmu, toPoints } from './targets.ts';

/** The `a:tcPr` line of each side, and the resolved border it reads from. */
const SIDES: Readonly<
  Record<AutomationCellBorderSide, { readonly line: CellLine; readonly resolved: string }>
> = Object.freeze({
  __proto__: null,
  left: { line: 'lnL', resolved: 'left' },
  right: { line: 'lnR', resolved: 'right' },
  top: { line: 'lnT', resolved: 'top' },
  bottom: { line: 'lnB', resolved: 'bottom' },
  diagonalDown: { line: 'lnTlToBr', resolved: 'tlToBr' },
  diagonalUp: { line: 'lnBlToTr', resolved: 'blToTr' },
} as Record<AutomationCellBorderSide, { line: CellLine; resolved: string }>);

/** CT_TableCellProperties margin defaults (ECMA-376 Part 1, 21.1.3.17), EMU. */
const MARGINS = [
  ['left', 'marL', 91_440],
  ['right', 'marR', 91_440],
  ['top', 'marT', 45_720],
  ['bottom', 'marB', 45_720],
] as const;

/** The widest line a write accepts, in points (`ST_LineWidth`, 20 116 800 EMU). */
const MAX_WEIGHT = 1584;

function sideOf(side: unknown): { line: CellLine; resolved: string } {
  if (typeof side !== 'string' || !Object.prototype.hasOwnProperty.call(SIDES, side))
    badArgument('side');
  return SIDES[side as AutomationCellBorderSide];
}

export function readTableCellFormat(
  ctx: TextPlanContext,
  ref: AutomationRef,
  cell: AutomationCellAddress
): AutomationValue {
  const { at, m } = tableAt(ctx, ref, true);
  const block = cellBlock(m, cell);
  const tc = m.tcs[block.r0]![block.c0]!;
  const tcPr = child(tc, 'a:tcPr');
  const resolved = resolvedShape(ctx, at);
  const hit =
    resolved?.kind === 'graphicFrame'
      ? resolved.table?.cells.find((c) => c.nodeId === tc.id)
      : undefined;
  const border = (side: AutomationCellBorderSide): AutomationLineRead => {
    const { line, resolved: edge } = SIDES[side];
    if (hit) return lineRead(hit.borders[edge as keyof typeof hit.borders]);
    return lineElementRead(child(tcPr, `a:${line}`));
  };
  const margin = (name: string, fallback: number): number =>
    toPoints(intAttr(tcPr, name) ?? fallback);
  const format: AutomationTableCellFormatRead = {
    borders: {
      left: border('left'),
      right: border('right'),
      top: border('top'),
      bottom: border('bottom'),
      diagonalDown: border('diagonalDown'),
      diagonalUp: border('diagonalUp'),
    },
    margins: {
      left: margin('marL', MARGINS[0][2]),
      right: margin('marR', MARGINS[1][2]),
      top: margin('marT', MARGINS[2][2]),
      bottom: margin('marB', MARGINS[3][2]),
    },
  };
  return { kind: 'tableCellFormat', format };
}

/**
 * Set one line of a cell's block. The fields the write leaves out keep the line the cell shows
 * now (its own, or the table style's); a cell with no line starts from a black solid line of
 * 1 point. Any color, weight, dash, or transparency makes the line visible.
 */
export function setTableCellBorder(
  ctx: TextPlanContext,
  ref: AutomationRef,
  cell: AutomationCellAddress,
  side: AutomationCellBorderSide,
  line: AutomationLineWrite
): AutomationValue {
  const { line: name } = sideOf(side);
  if (typeof line !== 'object' || line === null) badArgument('line');
  const { m } = tableAt(ctx, ref);
  const block = cellBlock(m, cell);
  const before = readTableCellFormat(ctx, ref, cell);
  const current = before.kind === 'tableCellFormat' ? before.format.borders[side] : null;
  const styled =
    line.color !== undefined ||
    line.weight !== undefined ||
    line.dashStyle !== undefined ||
    line.transparency !== undefined ||
    line.style !== undefined;
  if (line.visible !== undefined && typeof line.visible !== 'boolean') badArgument('visible');
  const visible = line.visible ?? (styled ? true : (current?.visible ?? false));
  // A value the write names is checked; a value kept from the file that is out of range or
  // unknown falls back to the default, so a write never fails on what it did not set.
  const kept = <T>(value: T | null | undefined, ok: (v: T) => boolean, fallback: T): T =>
    value !== null && value !== undefined && ok(value) ? value : fallback;
  const weight =
    line.weight !== undefined
      ? numberIn(line.weight, 'weight', 0, MAX_WEIGHT)
      : kept(current?.weight, (w) => w >= 0 && w <= MAX_WEIGHT, 1);
  if (line.dashStyle !== undefined && !DASHES.has(line.dashStyle)) badArgument('dashStyle');
  const dash = line.dashStyle ?? kept(current?.dashStyle, (d) => DASHES.has(d), 'solid');
  if (line.style !== undefined && !COMPOUNDS.has(line.style)) badArgument('style');
  const compound = line.style ?? kept(current?.style, (c) => COMPOUNDS.has(c), 'sng');
  const transparency = numberIn(
    line.transparency ?? current?.transparency ?? 0,
    'transparency',
    0,
    1
  );
  // Check the color now: an invisible line still keeps no color, but a bad one is refused.
  const color = colorSpec(line.color ?? current?.color ?? '#000000', 'color', transparency);
  const fill = visible ? `<a:solidFill>${color}</a:solidFill>` : '<a:noFill/>';
  // CT_LineProperties children in schema order: fill, dash, join, head and tail ends.
  const xml = (el: CellLine): string =>
    `<a:${el} w="${toEmu(weight)}" cap="flat" cmpd="${compound}" algn="ctr">${fill}` +
    `<a:prstDash val="${dash}"/><a:round/><a:headEnd type="none" w="med" len="med"/>` +
    `<a:tailEnd type="none" w="med" len="med"/></a:${el}>`;
  edit(() => setCellEdge(ctx.batch, m, block, name, xml));
  return { kind: 'applied' };
}

/** Set the margins of a cell's block, in points. Null removes a margin, back to its default. */
export function setTableCellMargins(
  ctx: TextPlanContext,
  ref: AutomationRef,
  cell: AutomationCellAddress,
  margins: AutomationCellMarginsWrite
): AutomationValue {
  if (typeof margins !== 'object' || margins === null) badArgument('margins');
  const { at, m } = tableAt(ctx, ref);
  const block = cellBlock(m, cell);
  const emu: { -readonly [K in keyof AutomationCellMarginsWrite]: number | null } = {};
  for (const [key] of MARGINS) {
    const value = margins[key];
    if (value === undefined) continue;
    emu[key] = value === null ? null : toEmu(numberIn(value, key, 0, 4032));
  }
  if (Object.keys(emu).length === 0) return { kind: 'applied' };
  edit(() => setCellMargins(ctx.batch, m, block, emu));
  refit(ctx, at);
  return { kind: 'applied' };
}

/** The `a:tblPr` flag of each style option a write names. */
const OPTIONS: readonly (readonly [keyof AutomationTableStyleWrite, TableStyleOption])[] = [
  ['firstRow', 'firstRow'],
  ['lastRow', 'lastRow'],
  ['firstColumn', 'firstCol'],
  ['lastColumn', 'lastCol'],
  ['bandedRows', 'bandRow'],
  ['bandedColumns', 'bandCol'],
];

/** The name of the built-in default style (`cascade/table-styles.ts`). */
const DEFAULT_STYLE_NAME = 'Medium Style 2 - Accent 1';

/** The id of the style named `name`: the presentation's own, or the built-in default. */
function styleIdOf(ctx: TextPlanContext, name: string): string {
  const own = tableStyleList(ctx.batch.pkg).find((s) => !s.builtIn && s.name === name);
  if (own) return own.id;
  if (name === DEFAULT_STYLE_NAME) return DEFAULT_TABLE_STYLE_ID;
  throw new AutomationFailure(
    'unsupported-content',
    'the presentation does not define that table style',
    'table-style-undefined'
  );
}

/**
 * Set style options (`a:tblPr` flags) and the style a table uses. A style is found by its
 * `styleName` in `ppt/tableStyles.xml`; the file format holds no other definition of a style,
 * so only the built-in default can be used without one.
 */
export function setTableStyle(
  ctx: TextPlanContext,
  ref: AutomationRef,
  style: AutomationTableStyleWrite
): AutomationValue {
  if (typeof style !== 'object' || style === null) badArgument('style');
  const { at, m } = tableAt(ctx, ref);
  const options: Partial<Record<TableStyleOption, boolean>> = {};
  for (const [key, flag] of OPTIONS) {
    const value = style[key];
    if (value === undefined) continue;
    if (typeof value !== 'boolean') badArgument(key);
    options[flag] = value;
  }
  const name = style.styleName;
  if (name !== undefined && name !== null && typeof name !== 'string') badArgument('styleName');
  const id = typeof name === 'string' ? styleIdOf(ctx, name) : undefined;
  edit(() => {
    setStyleOptions(ctx.batch, m, options);
    if (name === null) clearTableStyle(ctx.batch, m);
    else if (id !== undefined) setStyleId(ctx.batch, m, id);
  });
  refit(ctx, at);
  return { kind: 'applied' };
}
