// Table formatting on the canonical tree: cell properties (`a:tcPr`, ECMA-376 Part 1, 21.1.3.17:
// fill, borders, margins, anchor) and table properties (`a:tblPr`, 21.1.3.15: the style id and
// the style options). Each edit compiles to ops on one batch.

import { attr, child, elements, NS, textOf, type Element } from '../../store/index.ts';
import type {
  ColorChoice,
  LineDash,
  TableBorderSpec,
  TableBorderTarget,
  TableStyleOption,
  VerticalAlign,
} from '../../contracts/neutral.ts';
import { STYLE_ID } from '../../cascade/table-styles.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { specFromXml } from './shape-commands.ts';
import { colorXml, isColorChoice, MAX_LINE_WIDTH } from './style-commands.ts';
import {
  aSpec,
  anchorsIn,
  blockOf,
  setAttr,
  type CellRange,
  type TableModel,
} from './table-model.ts';

/** `a:tcPr` children in schema order (CT_TableCellProperties). */
const TCPR_ORDER = [
  'lnL',
  'lnR',
  'lnT',
  'lnB',
  'lnTlToBr',
  'lnBlToTr',
  'cell3D',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'headers',
  'extLst',
];
const FILLS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];

/** The largest cell margin the editor writes: the cascade's length cap. */
const MAX_LENGTH = 51_206_400;
const DASHES: readonly LineDash[] = [
  'solid',
  'sysDot',
  'sysDash',
  'dash',
  'dashDot',
  'lgDash',
  'lgDashDot',
];

/** The cell's `a:tcPr`, created after its text body when missing. Returns its id. */
function ensureTcPr(batch: OpBatch, part: string, tcId: string): string {
  const tc = batch.element(part, tcId);
  const tcPr = child(tc, 'a:tcPr');
  if (tcPr) return tcPr.id;
  const body = tc.children.findIndex((n) => n.type === 'element' && n.local === 'txBody');
  return batch.insert({
    op: 'insertNode',
    part,
    parent: tcId,
    index: body + 1,
    spec: aSpec('tcPr'),
  });
}

/** Replace the `a:tcPr` children named in `group` with `xml` (none when null), in schema order. */
function replaceTcPrChild(
  batch: OpBatch,
  part: string,
  tcId: string,
  group: readonly string[],
  xml: string | null
): void {
  const id = ensureTcPr(batch, part, tcId);
  for (const c of elements(batch.element(part, id)))
    if (c.ns === NS.a && group.includes(c.local))
      batch.push({ op: 'removeNode', part, node: c.id });
  if (!xml) return;
  const spec = specFromXml(xml);
  const rank = TCPR_ORDER.indexOf(spec.local);
  let index = 0;
  batch.element(part, id).children.forEach((c, i) => {
    if (c.type === 'element' && TCPR_ORDER.indexOf(c.local) <= rank) index = i + 1;
  });
  batch.insert({ op: 'insertNode', part, parent: id, index, spec });
}

/** The top-left `a:tc` of every block in a range. */
function blockCells(m: TableModel, range: CellRange): Element[] {
  return anchorsIn(m, range).map((a) => m.tcs[a.row][a.column]);
}

/** Cell shading: a solid fill, or `a:noFill` for No Fill (the Shading menu). */
export function setCellFill(
  batch: OpBatch,
  m: TableModel,
  range: CellRange,
  color: ColorChoice | null
): void {
  if (color !== null && !isColorChoice(color))
    throw new CommandRefusal('errors.unsupported', 'color');
  const xml = color === null ? '<a:noFill/>' : `<a:solidFill>${colorXml(color)}</a:solidFill>`;
  setCellFillXml(batch, m, range, xml);
}

/** Replace the fill of each block in a range with one EG_FillProperties element. */
export function setCellFillXml(batch: OpBatch, m: TableModel, range: CellRange, xml: string): void {
  for (const tc of blockCells(m, range)) replaceTcPrChild(batch, m.part, tc.id, FILLS, xml);
}

type Side = 'lnL' | 'lnR' | 'lnT' | 'lnB';
const OPPOSITE: Record<Side, Side> = { lnL: 'lnR', lnR: 'lnL', lnT: 'lnB', lnB: 'lnT' };

/** One border line for a cell edge (`a:lnL`, `a:lnR`, `a:lnT`, `a:lnB`). */
function borderXml(side: Side, spec: TableBorderSpec | null): string {
  if (!spec) return `<a:${side} w="12700" cmpd="sng"><a:noFill/></a:${side}>`;
  const w = Math.round(spec.width * 12_700);
  return (
    `<a:${side} w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${colorXml(spec.color)}</a:solidFill>` +
    `<a:prstDash val="${spec.dash ?? 'solid'}"/><a:round/><a:headEnd type="none" w="med" len="med"/>` +
    `<a:tailEnd type="none" w="med" len="med"/></a:${side}>`
  );
}

/** Check a border spec from the chrome. */
function checkSpec(spec: TableBorderSpec | undefined): TableBorderSpec {
  if (!spec || !isColorChoice(spec.color)) throw new CommandRefusal('errors.unsupported', 'border');
  const w = spec.width * 12_700;
  if (!Number.isFinite(w) || w <= 0 || w > MAX_LINE_WIDTH)
    throw new CommandRefusal('errors.unsupported', 'width');
  if (spec.dash !== undefined && !DASHES.includes(spec.dash))
    throw new CommandRefusal('errors.unsupported', 'dash');
  return spec;
}

/**
 * Borders of a cell range. `all`, `outside`, `inside`, and one side follow the
 * Borders menu; `none` removes every border of the range. An edge on the range's outside is
 * also written on the neighbor across it, so the two cells agree on the line between them.
 */
export function setTableBorders(
  batch: OpBatch,
  m: TableModel,
  range: CellRange,
  scope: TableBorderTarget,
  input: TableBorderSpec | undefined
): void {
  const spec = scope === 'none' ? null : checkSpec(input);
  const writes = new Map<string, { tc: Element; side: Side }>();
  const write = (r: number, c: number, side: Side): void => {
    if (r < 0 || c < 0 || r >= m.nR || c >= m.nC) return;
    const b = blockOf(m, { row: r, column: c });
    const tc = m.tcs[b.r0][b.c0];
    writes.set(`${tc.id}:${side}`, { tc, side });
  };
  for (const a of anchorsIn(m, range)) {
    const b = blockOf(m, a);
    // Each edge with the grid positions across it: a tall or wide block has several neighbors.
    const across = (r0: number, r1: number, c0: number, c1: number): [number, number][] => {
      const out: [number, number][] = [];
      for (let r = r0; r <= r1; r += 1) for (let c = c0; c <= c1; c += 1) out.push([r, c]);
      return out;
    };
    const edges: [Side, boolean, [number, number][]][] = [
      ['lnL', b.c0 === range.c0, across(b.r0, b.r1, b.c0 - 1, b.c0 - 1)],
      ['lnR', b.c1 === range.c1, across(b.r0, b.r1, b.c1 + 1, b.c1 + 1)],
      ['lnT', b.r0 === range.r0, across(b.r0 - 1, b.r0 - 1, b.c0, b.c1)],
      ['lnB', b.r1 === range.r1, across(b.r1 + 1, b.r1 + 1, b.c0, b.c1)],
    ];
    for (const [side, outer, neighbors] of edges) {
      const wanted =
        scope === 'all' ||
        scope === 'none' ||
        (scope === 'outside' && outer) ||
        (scope === 'inside' && !outer) ||
        (outer &&
          ((scope === 'left' && side === 'lnL') ||
            (scope === 'right' && side === 'lnR') ||
            (scope === 'top' && side === 'lnT') ||
            (scope === 'bottom' && side === 'lnB')));
      if (!wanted) continue;
      write(a.row, a.column, side);
      if (outer) for (const [nr, nc] of neighbors) write(nr, nc, OPPOSITE[side]);
    }
  }
  for (const { tc, side } of writes.values())
    replaceTcPrChild(batch, m.part, tc.id, [side], borderXml(side, spec));
}

/** Cell margins (`marL`, `marR`, `marT`, `marB`), EMU. Null removes one, back to its default. */
export function setCellMargins(
  batch: OpBatch,
  m: TableModel,
  range: CellRange,
  margins: {
    left?: number | null;
    right?: number | null;
    top?: number | null;
    bottom?: number | null;
  }
): void {
  const sides: [string, number | null | undefined][] = [
    ['marL', margins.left],
    ['marR', margins.right],
    ['marT', margins.top],
    ['marB', margins.bottom],
  ];
  if (sides.every(([, v]) => v === undefined))
    throw new CommandRefusal('errors.unsupported', 'margins');
  for (const [, v] of sides)
    if (v !== undefined && v !== null && (!Number.isFinite(v) || v < 0 || v > MAX_LENGTH))
      throw new CommandRefusal('errors.unsupported', 'margin');
  for (const tc of blockCells(m, range)) {
    const id = ensureTcPr(batch, m.part, tc.id);
    for (const [name, v] of sides)
      if (v !== undefined)
        setAttr(batch, m.part, id, name, v === null ? null : String(Math.round(v)));
  }
}

/** A cell line element of `a:tcPr` (CT_TableCellProperties, ECMA-376 Part 1, 21.1.3.17). */
export type CellLine = Side | 'lnTlToBr' | 'lnBlToTr';

/**
 * Write one line of a cell block: `xml(name)` gives the element for a line name. An outside
 * edge is also written on each cell across it, as the opposite edge, so the two cells agree on
 * the line between them (as `setTableBorders` does). A diagonal has no neighbor.
 */
export function setCellEdge(
  batch: OpBatch,
  m: TableModel,
  block: CellRange,
  line: CellLine,
  xml: (name: CellLine) => string
): void {
  const writes = new Map<string, { tc: Element; line: CellLine }>();
  const write = (r: number, c: number, name: CellLine): void => {
    if (r < 0 || c < 0 || r >= m.nR || c >= m.nC) return;
    const b = blockOf(m, { row: r, column: c });
    const tc = m.tcs[b.r0][b.c0];
    writes.set(`${tc.id}:${name}`, { tc, line: name });
  };
  write(block.r0, block.c0, line);
  if (line === 'lnL' || line === 'lnR') {
    const c = line === 'lnL' ? block.c0 - 1 : block.c1 + 1;
    for (let r = block.r0; r <= block.r1; r += 1) write(r, c, OPPOSITE[line]);
  } else if (line === 'lnT' || line === 'lnB') {
    const r = line === 'lnT' ? block.r0 - 1 : block.r1 + 1;
    for (let c = block.c0; c <= block.c1; c += 1) write(r, c, OPPOSITE[line]);
  }
  for (const w of writes.values()) replaceTcPrChild(batch, m.part, w.tc.id, [w.line], xml(w.line));
}

/**
 * Vertical alignment of cell text (`a:tcPr/@anchor`). `center` also sets `@anchorCtr`, which
 * centers the text block horizontally (ECMA-376 Part 1, 21.1.3.17); absent leaves it.
 */
export function setCellAnchor(
  batch: OpBatch,
  m: TableModel,
  range: CellRange,
  anchor: VerticalAlign,
  center?: boolean
): void {
  if (anchor !== 't' && anchor !== 'ctr' && anchor !== 'b')
    throw new CommandRefusal('errors.unsupported', 'anchor');
  for (const tc of blockCells(m, range)) {
    const tcPr = ensureTcPr(batch, m.part, tc.id);
    setAttr(batch, m.part, tcPr, 'anchor', anchor);
    if (center !== undefined) setAttr(batch, m.part, tcPr, 'anchorCtr', center ? '1' : null);
  }
}

/** The table's `a:tblPr`, created first in `a:tbl` when missing. Returns its id. */
function ensureTblPr(batch: OpBatch, m: TableModel): string {
  const tblPr = child(batch.element(m.part, m.tbl.id), 'a:tblPr');
  if (tblPr) return tblPr.id;
  return batch.insert({
    op: 'insertNode',
    part: m.part,
    parent: m.tbl.id,
    index: 0,
    spec: aSpec('tblPr'),
  });
}

/** Use the table style `styleId` (`a:tblPr/a:tableStyleId`); an inline style gives way. */
export function setTableStyle(batch: OpBatch, m: TableModel, styleId: string): void {
  if (!STYLE_ID.test(styleId)) throw new CommandRefusal('errors.unsupported', 'style');
  const id = ensureTblPr(batch, m);
  const tblPr = batch.element(m.part, id);
  const current = child(tblPr, 'a:tableStyleId');
  if (current && textOf(current).trim() === styleId && !child(tblPr, 'a:tableStyle')) return;
  for (const c of elements(tblPr))
    if (c.local === 'tableStyle' || c.local === 'tableStyleId')
      batch.push({ op: 'removeNode', part: m.part, node: c.id });
  // CT_TableProperties: fill and effect properties, then the style choice, then extLst.
  const now = batch.element(m.part, id);
  const ext = now.children.findIndex((n) => n.type === 'element' && n.local === 'extLst');
  batch.insert({
    op: 'insertNode',
    part: m.part,
    parent: id,
    index: ext < 0 ? now.children.length : ext,
    spec: aSpec('tableStyleId', {}, [{ text: styleId }]),
  });
}

/** Whether a style option is on (`a:tblPr` flags default to off). */
export function styleOption(tblPr: Element | undefined, option: TableStyleOption): boolean {
  const v = attr(tblPr, option);
  return v === '1' || v === 'true';
}

/** Set style options on (true) or off (false); an absent option keeps its value. */
export function setStyleOptions(
  batch: OpBatch,
  m: TableModel,
  options: Partial<Readonly<Record<TableStyleOption, boolean>>>
): void {
  const entries = Object.entries(options).filter(([, on]) => on !== undefined);
  if (entries.length === 0) return;
  const id = ensureTblPr(batch, m);
  for (const [option, on] of entries) setAttr(batch, m.part, id, option, on ? '1' : null);
}

/** Remove the table's style choice (`a:tableStyleId` or an inline `a:tableStyle`). */
export function clearTableStyle(batch: OpBatch, m: TableModel): void {
  const tblPr = child(batch.element(m.part, m.tbl.id), 'a:tblPr');
  for (const c of elements(tblPr))
    if (c.local === 'tableStyle' || c.local === 'tableStyleId')
      batch.push({ op: 'removeNode', part: m.part, node: c.id });
}

/** Turn a style option on or off. */
export function toggleStyleOption(batch: OpBatch, m: TableModel, option: TableStyleOption): void {
  const on = styleOption(child(m.tbl, 'a:tblPr'), option);
  setAttr(batch, m.part, ensureTblPr(batch, m), option, on ? null : '1');
}
