// Paint a laid-out table (design D3 of `table-rendering`): the table background, every cell's
// fill, then every border, then cell text. Borders paint after all fills, so a neighbor's fill
// never covers a border. Coordinates are the frame's local EMU.

import type { ResolvedLine } from '../cascade/index.ts';
import type { CellLayout, TableLayout } from '../layout/index.ts';
import { fillPaint, strokeAttrs, type PaintContext } from './paint-fill.ts';
import { paintText, textShows, type Shows } from './paint-text.ts';
import { num, svg } from './svg.ts';

function edge(
  doc: Document,
  line: ResolvedLine | undefined,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  ctx: PaintContext
): SVGElement | null {
  if (!line) return null;
  const stroke = strokeAttrs(line, Math.abs(x2 - x1), Math.abs(y2 - y1), ctx);
  if (stroke.stroke === 'none') return null;
  return svg(doc, 'line', {
    x1: num(x1),
    y1: num(y1),
    x2: num(x2),
    y2: num(y2),
    ...stroke,
    // Solid cell borders meet square (measured on reference PDFs). A dashed border keeps its own cap: a
    // forced square cap lengthens every dash by one line width and closes a dotted border's gaps.
    'stroke-linecap': stroke['stroke-dasharray'] ? stroke['stroke-linecap'] : 'square',
  });
}

function borders(doc: Document, c: CellLayout, ctx: PaintContext): SVGElement[] {
  const b = c.paintBorders ?? c.cell.borders;
  const r = c.x + c.width;
  const bottom = c.y + c.height;
  return [
    edge(doc, b.tlToBr, c.x, c.y, r, bottom, ctx),
    edge(doc, b.blToTr, c.x, bottom, r, c.y, ctx),
    edge(doc, b.top, c.x, c.y, r, c.y, ctx),
    edge(doc, b.bottom, c.x, bottom, r, bottom, ctx),
    edge(doc, b.left, c.x, c.y, c.x, bottom, ctx),
    edge(doc, b.right, r, c.y, r, bottom, ctx),
  ].filter((e): e is SVGElement => e !== null);
}

/**
 * Whether cell `c` can show, with its borders and a margin of its widest border on every side.
 * Its text is checked line by line, since it can overflow the cell.
 */
function cellShows(c: CellLayout, shows: Shows): boolean {
  const b = c.cell.borders;
  let pad = 0;
  for (const line of [b.top, b.bottom, b.left, b.right, b.tlToBr, b.blToTr])
    if (line) pad = Math.max(pad, line.width);
  const x0 = c.x - pad;
  const y0 = c.y - pad;
  const x1 = c.x + c.width + pad;
  const y1 = c.y + c.height + pad;
  return shows([
    { x: x0, y: y0 },
    { x: x1, y: y0 },
    { x: x1, y: y1 },
    { x: x0, y: y1 },
  ]);
}

/**
 * The table's SVG nodes, in the frame's local space. With `shows`, a cell that cannot show
 * paints no fill and no border, and a text line that cannot show is left out.
 */
export function paintTable(
  doc: Document,
  table: TableLayout,
  background: Parameters<typeof fillPaint>[0],
  ctx: PaintContext,
  shows?: Shows
): SVGElement[] {
  const out: SVGElement[] = [];
  const cells = shows ? table.cells.filter((c) => cellShows(c, shows)) : table.cells;
  const bg = fillPaint(background, table.width, table.height, ctx);
  if (bg.paint !== 'none')
    out.push(
      svg(doc, 'rect', {
        width: num(table.width),
        height: num(table.height),
        fill: bg.paint,
        'fill-opacity': bg.opacity < 1 ? num(bg.opacity) : undefined,
      })
    );
  for (const c of cells) {
    const fill = fillPaint(c.cell.fill, c.width, c.height, ctx);
    if (fill.paint === 'none') continue;
    // Gradients and pictures fill the cell's own box.
    out.push(
      svg(doc, 'g', { transform: `translate(${num(c.x)} ${num(c.y)})` }, [
        svg(doc, 'rect', {
          width: num(c.width),
          height: num(c.height),
          fill: fill.paint,
          'fill-opacity': fill.opacity < 1 ? num(fill.opacity) : undefined,
        }),
      ])
    );
  }
  const lines = svg(doc, 'g', { 'data-kind': 'table-borders' });
  for (const c of cells) for (const node of borders(doc, c, ctx)) lines.appendChild(node);
  out.push(lines);
  const shown = new Set(cells);
  for (const c of table.cells) {
    if (!c.text.lines.length) continue;
    // A cell out of view paints its text only when the text overflows into view.
    if (shows && !shown.has(c) && !textShows(c.text, shows)) continue;
    out.push(paintText(doc, c.text, ctx, shows));
  }
  return out;
}
