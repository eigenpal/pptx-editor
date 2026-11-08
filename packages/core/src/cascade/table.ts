// Resolve `a:tbl` (design D2–D4 of `table-rendering`): the grid, merges, table style parts in
// the reference order with per-region borders, cell properties, and cell text through the same
// text cascade as a text box.

import {
  attr,
  boolAttr,
  child,
  children,
  elements,
  intAttr,
  tableAsDrawn,
  TABLE_MAX_COLUMNS,
  TABLE_MAX_ROWS,
  type Element,
  type PptxPackage,
  type TableGridBlock,
} from '../store/index.ts';
import { colorChild, resolveColor, type ColorContext, type Rgba } from './color.ts';
import {
  fillElementOf,
  resolveFillElement,
  resolveLine,
  styleRefContext,
  themeFill,
  type ResolvedFill,
  type ResolvedLine,
} from './paint-props.ts';
import { tableStyleFor } from './table-styles.ts';
import { entryParts } from './picture-bullet.ts';
import { colorElement, levelEntries, resolveParagraph, syntheticElement } from './text-style.ts';
import type { Theme } from './theme.ts';
import type {
  BodyProps,
  CellBorders,
  KinsokuRules,
  ResolvedTable,
  ResolvedTableCell,
  ResolvedTextBody,
} from './types.ts';

/** What table resolution needs about the slide. */
export interface TableScope {
  readonly pkg: PptxPackage;
  readonly theme: Theme;
  readonly ctx: ColorContext;
  readonly slideNumber: number;
  readonly defaultTextStyle?: Element;
  /** The presentation's custom kinsoku lists (`p:kinsoku`). */
  readonly kinsoku?: KinsokuRules;
  /** The part of the table, whose relationships the picture bullets of its cells name. */
  readonly partName?: string;
  /** The presentation part, whose relationships the picture bullets of `defaultTextStyle` name. */
  readonly presentationPart?: string;
}

const MAX_LENGTH = 51_206_400;

type Side = 'left' | 'right' | 'top' | 'bottom' | 'insideH' | 'insideV' | 'tl2br' | 'tr2bl';
const SIDES: readonly Side[] = [
  'left',
  'right',
  'top',
  'bottom',
  'insideH',
  'insideV',
  'tl2br',
  'tr2bl',
];

/** The text properties a style part sets. */
interface TextBits {
  bold?: boolean;
  italic?: boolean;
  latin?: string;
  ea?: string;
  cs?: string;
  color?: Rgba;
}

/** One style part (`wholeTbl`, `firstRow`, …), resolved. A null border is "no line". */
interface StylePart {
  readonly fill?: ResolvedFill;
  readonly text: TextBits;
  readonly borders: ReadonlyMap<Side, ResolvedLine | null>;
}

/** The rows and columns a style part covers, inclusive. */
interface Region {
  readonly r0: number;
  readonly r1: number;
  readonly c0: number;
  readonly c1: number;
}

const length = (v: number | undefined, d: number): number =>
  v === undefined ? d : Math.max(0, Math.min(v, MAX_LENGTH));

const onOff = (v: string | undefined): boolean | undefined =>
  v === 'on' ? true : v === 'off' ? false : undefined;

/** A line from `a:ln` or a theme `a:lnRef`; null when it draws nothing. */
function lineOf(holder: Element | undefined, scope: TableScope): ResolvedLine | null | undefined {
  if (!holder) return undefined;
  const ln = child(holder, 'a:ln');
  const ref = child(holder, 'a:lnRef');
  let line: ResolvedLine | undefined;
  if (ln) line = resolveLine(ln, undefined, scope.ctx, scope.ctx);
  else if (ref) {
    const idx = intAttr(ref, 'idx') ?? 0;
    const themed = idx > 0 ? scope.theme.lineStyles[idx - 1] : undefined;
    line = themed
      ? resolveLine(undefined, themed, scope.ctx, styleRefContext(ref, scope.ctx))
      : undefined;
  } else return undefined;
  return line && line.fill.type !== 'none' && line.width > 0 ? line : null;
}

/** A fill from `a:fill` (holding a fill element) or a theme `a:fillRef`. */
function fillOf(holder: Element | undefined, scope: TableScope): ResolvedFill | undefined {
  const wrapped = child(holder, 'a:fill');
  if (wrapped) return resolveFillElement(fillElementOf(wrapped), scope.ctx) ?? { type: 'none' };
  const ref = child(holder, 'a:fillRef');
  if (!ref) return undefined;
  const themed = themeFill(intAttr(ref, 'idx') ?? 0, scope.theme);
  return themed
    ? (resolveFillElement(themed, styleRefContext(ref, scope.ctx)) ?? { type: 'none' })
    : { type: 'none' };
}

function readPart(el: Element | undefined, scope: TableScope): StylePart | undefined {
  if (!el) return undefined;
  const tcStyle = child(el, 'a:tcStyle');
  const borders = new Map<Side, ResolvedLine | null>();
  const tcBdr = child(tcStyle, 'a:tcBdr');
  for (const side of SIDES) {
    const line = lineOf(child(tcBdr, `a:${side}`), scope);
    if (line !== undefined) borders.set(side, line);
  }
  const text: TextBits = {};
  const tx = child(el, 'a:tcTxStyle');
  if (tx) {
    const b = onOff(attr(tx, 'b'));
    const i = onOff(attr(tx, 'i'));
    if (b !== undefined) text.bold = b;
    if (i !== undefined) text.italic = i;
    const font = child(tx, 'a:font');
    const ref = child(tx, 'a:fontRef');
    if (font) {
      for (const slot of ['latin', 'ea', 'cs'] as const) {
        const face = attr(child(font, `a:${slot}`), 'typeface');
        if (face) text[slot] = face;
      }
    } else if (ref) {
      const m = attr(ref, 'idx') === 'major' ? 'mj' : 'mn';
      text.latin = `+${m}-lt`;
      text.ea = `+${m}-ea`;
      text.cs = `+${m}-cs`;
    }
    const color = resolveColor(colorChild(tx) ?? colorChild(ref), scope.ctx);
    if (color) text.color = color;
  }
  const fill = fillOf(tcStyle, scope);
  return { ...(fill ? { fill } : {}), text, borders };
}

/** The text style a cell's parts set, as a chain entry. */
function textEntry(bits: TextBits): Element | undefined {
  const attrs: Record<string, string> = {};
  if (bits.bold !== undefined) attrs.b = bits.bold ? '1' : '0';
  if (bits.italic !== undefined) attrs.i = bits.italic ? '1' : '0';
  const kids: Element[] = [];
  if (bits.color) kids.push(syntheticElement('solidFill', {}, [colorElement(bits.color)]));
  for (const slot of ['latin', 'ea', 'cs'] as const) {
    const face = bits[slot];
    if (face) kids.push(syntheticElement(slot, { typeface: face }));
  }
  if (!kids.length && !Object.keys(attrs).length) return undefined;
  return syntheticElement('defPPr', {}, [syntheticElement('defRPr', attrs, kids)]);
}

function cellBody(tcPr: Element | undefined): BodyProps {
  const anchor = attr(tcPr, 'anchor');
  return {
    lIns: length(intAttr(tcPr, 'marL'), 91_440),
    rIns: length(intAttr(tcPr, 'marR'), 91_440),
    tIns: length(intAttr(tcPr, 'marT'), 45_720),
    bIns: length(intAttr(tcPr, 'marB'), 45_720),
    anchor:
      anchor === 'ctr' || anchor === 'b' || anchor === 'just' || anchor === 'dist' ? anchor : 't',
    anchorCtr: boolAttr(tcPr, 'anchorCtr') ?? false,
    wrap: 'square',
    vert: attr(tcPr, 'vert') ?? 'horz',
    rot: 0,
    upright: false,
    numCol: 1,
    spcCol: 0,
    autofit: { kind: 'none' },
    spcFirstLastPara: false,
  };
}

function cellText(
  tc: Element,
  tcPr: Element | undefined,
  entry: Element | undefined,
  scope: TableScope
): ResolvedTextBody {
  const txBody = child(tc, 'a:txBody');
  const lstStyle = child(txBody, 'a:lstStyle');
  const chainFor = (level: number): Element[] => [
    ...levelEntries(scope.defaultTextStyle, level),
    ...(entry ? [entry] : []),
    ...levelEntries(lstStyle, level),
  ];
  // A cell's own list style and paragraphs are in the table's part; the deck's default text
  // style is in the presentation part. The table style's text entry holds no bullet.
  const pictures = {
    pkg: scope.pkg,
    partOf: entryParts(
      [
        [scope.defaultTextStyle, scope.presentationPart],
        [lstStyle, scope.partName],
      ],
      scope.partName
    ),
  };
  const paragraphs = elements(txBody)
    .filter((p) => p.kind === 'paragraph')
    .map((p) =>
      resolveParagraph({
        paragraph: p,
        chainFor,
        ctx: scope.ctx,
        theme: scope.theme,
        slideNumber: scope.slideNumber,
        ...(scope.kinsoku ? { kinsoku: scope.kinsoku } : {}),
        pictures,
      })
    );
  return { ...(txBody ? { nodeId: txBody.id } : {}), body: cellBody(tcPr), paragraphs };
}

const CELL_LINES: readonly [string, keyof CellBorders][] = [
  ['a:lnL', 'left'],
  ['a:lnR', 'right'],
  ['a:lnT', 'top'],
  ['a:lnB', 'bottom'],
  ['a:lnTlToBr', 'tlToBr'],
  ['a:lnBlToTr', 'blToTr'],
];

/** Resolve the `a:tbl` inside a frame's `a:graphicData`. */
export function resolveTable(
  graphicData: Element | undefined,
  scope: TableScope
): ResolvedTable | undefined {
  return tableOf(graphicData, scope, undefined);
}

/**
 * Resolve the table of a graphic frame by the frame's merge grid (`tableGridOf`): a table with
 * a structure mark draws by the grid its save writes. @internal
 */
export function resolveFrameTable(frame: Element, scope: TableScope): ResolvedTable | undefined {
  return tableOf(child(child(frame, 'a:graphic'), 'a:graphicData'), scope, frame);
}

function tableOf(
  graphicData: Element | undefined,
  scope: TableScope,
  frame: Element | undefined
): ResolvedTable | undefined {
  const tbl = child(graphicData, 'a:tbl');
  if (!tbl) return undefined;
  const tblPr = child(tbl, 'a:tblPr');
  // Grid positions and merge coverage (`tableAsDrawn`). A valid row holds one `a:tc` per grid
  // column. Coverage comes from the spans only (`resolveTableGrid`): a cell with an orphan
  // `hMerge` or `vMerge` (no span reaches it) is drawn with its own text. A table with a
  // structure mark draws as its save writes it: cell counts repaired, so a row that concurrent
  // edits left short draws the blank cells the save adds, and the merge grid the save writes.
  const saved = tableAsDrawn(frame, TABLE_MAX_COLUMNS, tbl);
  if (!saved) return undefined;
  const { grid, cells: rowsOfCells } = saved;
  const gridCols = children(child(tbl, 'a:tblGrid'), 'a:gridCol');
  const columns = Array.from({ length: saved.columns }, (_, c) =>
    length(c < gridCols.length ? intAttr(gridCols[c], 'w') : saved.addedColumnWidth, 0)
  );
  const trs = children(tbl, 'a:tr').slice(0, TABLE_MAX_ROWS);
  const nC = columns.length;
  const nR = trs.length;
  const flag = (name: string): boolean => boolAttr(tblPr, name) ?? false;
  const flags = {
    firstRow: flag('firstRow'),
    lastRow: flag('lastRow'),
    firstCol: flag('firstCol'),
    lastCol: flag('lastCol'),
    bandRow: flag('bandRow'),
    bandCol: flag('bandCol'),
  };
  const style = tableStyleFor(scope.pkg, tblPr);
  const parts = new Map<string, StylePart | undefined>();
  const part = (name: string): StylePart | undefined => {
    if (!parts.has(name)) parts.set(name, readPart(child(style, `a:${name}`), scope));
    return parts.get(name);
  };

  const tcs: (Element | undefined)[] = [];
  for (let r = 0; r < nR; r += 1) for (let c = 0; c < nC; c += 1) tcs.push(rowsOfCells[r][c]);
  const blockAt = (i: number): TableGridBlock | undefined => {
    const b = grid.blocks[grid.owner[i]];
    return b && b.row * nC + b.col === i ? b : undefined;
  };

  const band = (index: number, count: number, first: boolean, last: boolean): number => {
    const start = first ? 1 : 0;
    const end = last ? count - 2 : count - 1;
    return index >= start && index <= end ? (index - start) % 2 : -1;
  };

  const cells: ResolvedTableCell[] = tcs.map((tc, i) => {
    const r = Math.floor(i / nC);
    const c = i % nC;
    const block = blockAt(i);
    const g = block?.gridSpan ?? 1;
    const k = block?.rowSpan ?? 1;
    const r2 = r + k - 1;
    const c2 = c + g - 1;
    const isFirstRow = flags.firstRow && r === 0;
    const isLastRow = flags.lastRow && r2 === nR - 1;
    const isFirstCol = flags.firstCol && c === 0;
    const isLastCol = flags.lastCol && c2 === nC - 1;
    const applied: [StylePart | undefined, Region][] = [
      [part('wholeTbl'), { r0: 0, r1: nR - 1, c0: 0, c1: nC - 1 }],
    ];
    const bandV = flags.bandCol ? band(c, nC, flags.firstCol, flags.lastCol) : -1;
    if (bandV >= 0)
      applied.push([part(bandV ? 'band2V' : 'band1V'), { r0: 0, r1: nR - 1, c0: c, c1: c }]);
    const bandH = flags.bandRow ? band(r, nR, flags.firstRow, flags.lastRow) : -1;
    if (bandH >= 0)
      applied.push([part(bandH ? 'band2H' : 'band1H'), { r0: r, r1: r, c0: 0, c1: nC - 1 }]);
    if (isLastCol) applied.push([part('lastCol'), { r0: 0, r1: nR - 1, c0: nC - 1, c1: nC - 1 }]);
    if (isFirstCol) applied.push([part('firstCol'), { r0: 0, r1: nR - 1, c0: 0, c1: 0 }]);
    if (isLastRow) applied.push([part('lastRow'), { r0: nR - 1, r1: nR - 1, c0: 0, c1: nC - 1 }]);
    if (isFirstRow) applied.push([part('firstRow'), { r0: 0, r1: 0, c0: 0, c1: nC - 1 }]);
    const one: Region = { r0: r, r1: r2, c0: c, c1: c2 };
    if (isLastRow && isLastCol) applied.push([part('seCell'), one]);
    if (isLastRow && isFirstCol) applied.push([part('swCell'), one]);
    if (isFirstRow && isLastCol) applied.push([part('neCell'), one]);
    if (isFirstRow && isFirstCol) applied.push([part('nwCell'), one]);

    let fill: ResolvedFill = { type: 'none' };
    const text: TextBits = {};
    const borders: { -readonly [K in keyof CellBorders]: ResolvedLine | null } = {};
    for (const [p, region] of applied) {
      if (!p) continue;
      if (p.fill) fill = p.fill;
      Object.assign(text, p.text);
      const edges: [keyof CellBorders, Side][] = [
        ['left', c === region.c0 ? 'left' : 'insideV'],
        ['right', c2 === region.c1 ? 'right' : 'insideV'],
        ['top', r === region.r0 ? 'top' : 'insideH'],
        ['bottom', r2 === region.r1 ? 'bottom' : 'insideH'],
        ['tlToBr', 'tl2br'],
        ['blToTr', 'tr2bl'],
      ];
      for (const [edge, side] of edges)
        if (p.borders.has(side)) borders[edge] = p.borders.get(side)!;
    }
    const tcPr = child(tc, 'a:tcPr');
    const ownFill = fillElementOf(tcPr);
    if (ownFill) fill = resolveFillElement(ownFill, scope.ctx) ?? { type: 'none' };
    for (const [name, edge] of CELL_LINES) {
      const ln = child(tcPr, name);
      if (!ln) continue;
      const line = resolveLine(ln, undefined, scope.ctx, scope.ctx);
      borders[edge] = line && line.fill.type !== 'none' && line.width > 0 ? line : null;
    }
    const resolvedBorders: { -readonly [K in keyof CellBorders]: ResolvedLine } = {};
    for (const [edge, line] of Object.entries(borders) as [
      keyof CellBorders,
      ResolvedLine | null,
    ][])
      if (line) resolvedBorders[edge] = line;
    return {
      nodeId: tc?.id ?? '',
      row: r,
      col: c,
      gridSpan: g,
      rowSpan: k,
      covered: !block,
      fill,
      borders: resolvedBorders,
      text: tc
        ? cellText(tc, tcPr, textEntry(text), scope)
        : { body: cellBody(undefined), paragraphs: [] },
    };
  });

  const ownBg = fillElementOf(tblPr);
  const background: ResolvedFill = ownBg
    ? (resolveFillElement(ownBg, scope.ctx) ?? { type: 'none' })
    : (fillOf(child(style, 'a:tblBg'), scope) ?? { type: 'none' });
  const styleId = attr(style, 'styleId');
  return {
    columns,
    rows: trs.map((tr) => ({ nodeId: tr.id, height: length(intAttr(tr, 'h'), 0) })),
    cells,
    background,
    rtl: boolAttr(tblPr, 'rtl') ?? false,
    ...(styleId ? { styleId } : {}),
  };
}
