// The editing chrome around a table on the slide, after DOCX Editor's table chrome: the select
// handle at the table's top-left corner (a click selects the table, a drag moves it), a line on
// the grid line under the pointer (a drag there resizes, `surface-table.ts`), and the "+"
// buttons on the table's left and top edges that insert a row below or a column right of a
// grid line. The chrome shows for the table the selection is in or on, and for a hovered table.
// Moved out of `surface.ts` (the 1000-line cap); the surface paints these in its overlay.

import type { SlideLayout } from '../layout/index.ts';
import { apply, invert } from '../layout/index.ts';
import type { EditorCommand, EditorSelection, TableCellAddress } from '../contracts/index.ts';
import { SVG_NS } from '../output/index.ts';
import { ICON_PATHS } from './icon-paths.ts';
import { gridLineAt, tableItem } from './table-cells.ts';
import { parentGroupOf } from '../automation/commands/placement.ts';

const EMU_PER_PX = 9_525;
/**
 * The select handle: a 16 px square, 6 px out from the table's top-left corner, so it stays
 * clear of the corner resize handle (8 px, centered on the corner).
 */
const HANDLE_PX = 16;
const HANDLE_GAP_PX = 6;
/** An insert button: a 16 px circle whose center is 14 px out from the table's edge. */
const INSERT_RADIUS_PX = 8;
const INSERT_OFFSET_PX = 14;
/** How far out from an edge (and in over it) the pointer shows that edge's insert buttons. */
const INSERT_BAND_OUT_PX = 26;
const INSERT_BAND_IN_PX = 8;
/** A grid line within this distance of the pointer is the one a press drags (`surface-table.ts`). */
const LINE_PX = 4;

type TableFrame = NonNullable<ReturnType<typeof tableItem>>;
type Point = { readonly x: number; readonly y: number };

/** What the table chrome needs from the surface. */
export interface TableChromeHost {
  readonly selection: () => EditorSelection;
  readonly readOnly: () => boolean;
  readonly zoom: () => number;
  readonly label: (key: string) => string;
  /** The format painter is armed: it only selects, so the chrome hides. */
  readonly painting?: () => boolean;
  select(selection: EditorSelection): void;
  run(command: EditorCommand): { readonly ok: boolean } | void;
}

/** An insert button: a row below grid line `index`, or a column right of it. */
type Insert = { readonly axis: 'row' | 'column'; readonly index: number };

/** What the pointer is over, for one table. */
type Hover = {
  readonly id: string;
  readonly insert?: Insert;
  readonly line?: { readonly axis: 'column' | 'row'; readonly index: number };
};

/**
 * A press on the chrome: `move` asks the surface to select table `id` (in its group, `scope`)
 * and drag it.
 */
export type TableChromePress =
  | { readonly kind: 'move'; readonly id: string; readonly scope: { readonly group?: string } }
  | { readonly kind: 'insert' };

/** The table frame the selection is in or on. */
export function selectedTableId(layout: SlideLayout, sel: EditorSelection): string | undefined {
  const id =
    sel.kind === 'text' && sel.cell
      ? sel.nodeId
      : sel.kind === 'cells'
        ? sel.nodeId
        : sel.kind === 'shapes' && sel.nodeIds.length === 1
          ? sel.nodeIds[0]
          : undefined;
  return id && tableItem(layout, id) ? id : undefined;
}

/** The scale of a frame's matrix: slide units for one local unit. */
function scaleOf(frame: TableFrame): number {
  const [a, b] = frame.matrix;
  return Math.hypot(a, b) || 1;
}

/** The grid line index a button at local position `at` stands for: the nearest inner or end line. */
function nearestLine(lines: readonly number[], at: number): number {
  let best = 1;
  for (let i = 1; i < lines.length; i += 1)
    if (Math.abs(lines[i] - at) < Math.abs(lines[best] - at)) best = i;
  return best;
}

/** The cell a caret goes into so that `insertRow below` or `insertColumn right` lands at the line. */
function insertCell(frame: TableFrame, insert: Insert): TableCellAddress | undefined {
  // A cell whose last row (or column) ends at the line: inserting below (right of) it inserts
  // at the line, also when the cell is merged. A line that every merged cell crosses has none.
  const cell = frame.table.cells.find((c) =>
    insert.axis === 'row'
      ? c.cell.row + c.cell.rowSpan === insert.index
      : c.cell.col + c.cell.gridSpan === insert.index
  );
  return cell ? { row: cell.cell.row, column: cell.cell.col } : undefined;
}

/** The table chrome of one surface. */
export class TableChrome {
  #hover: Hover | null = null;
  #key = '';

  constructor(readonly host: TableChromeHost) {}

  /** The chrome is off: in view mode and while the format painter is armed. */
  #off(): boolean {
    return this.host.readOnly() || (this.host.painting?.() ?? false);
  }

  /** The table the chrome shows for: the selection's table, else the hovered one. */
  #tableId(layout: SlideLayout): string | undefined {
    if (this.#off()) return undefined;
    return selectedTableId(layout, this.host.selection()) ?? this.#hover?.id;
  }

  /**
   * The pointer moved to slide point `p` over element `hit` (none outside the slide's shapes).
   * Returns true when the chrome changed, so the surface paints its overlay again.
   */
  hover(layout: SlideLayout, p: Point, hit: string | undefined): boolean {
    const next = this.#off() ? null : this.#hoverAt(layout, p, hit);
    const key = next
      ? `${next.id}|${next.insert ? `${next.insert.axis}${next.insert.index}` : ''}|${next.line ? `${next.line.axis}${next.line.index}` : ''}`
      : '';
    this.#hover = next;
    if (key === this.#key) return false;
    this.#key = key;
    return true;
  }

  #hoverAt(layout: SlideLayout, p: Point, hit: string | undefined): Hover | null {
    const px = EMU_PER_PX / this.host.zoom();
    const selected = selectedTableId(layout, this.host.selection());
    const ids = [hit, selected, this.#hover?.id];
    for (const id of ids) {
      const frame = id ? tableItem(layout, id) : undefined;
      if (!id || !frame) continue;
      // The select handle and the way to it, out from the top-left corner, keep the chrome.
      const corner = apply(frame.matrix, 0, 0);
      const reach = (HANDLE_PX + HANDLE_GAP_PX + 2) * px;
      if (p.x >= corner.x - reach && p.x <= corner.x && p.y >= corner.y - reach && p.y <= corner.y)
        return { id };
      const lp = px / scaleOf(frame);
      const local = apply(invert(frame.matrix), p.x, p.y);
      const t = frame.table;
      const out = INSERT_BAND_OUT_PX * lp;
      const over = INSERT_BAND_IN_PX * lp;
      // The pointer is near the table: on it, or in the band of its left or top edge.
      if (local.x < -out || local.y < -out || local.x > t.width + over || local.y > t.height + over)
        continue;
      const onTable = local.x >= 0 && local.y >= 0 && local.x <= t.width && local.y <= t.height;
      const line = onTable ? gridLineAt(frame, p, LINE_PX * px) : undefined;
      if (line) return { id, line };
      const inLeftBand = local.x <= over && local.y >= 0 && local.y <= t.height;
      const inTopBand = local.y <= over && local.x >= 0 && local.x <= t.width;
      if (inLeftBand && !inTopBand && t.rowY.length > 1)
        return { id, insert: { axis: 'row', index: nearestLine(t.rowY, local.y) } };
      // The grid of a right-to-left table runs the other way; its columns insert from the menus.
      if (inTopBand && !inLeftBand && !t.rtl && t.columnX.length > 1)
        return { id, insert: { axis: 'column', index: nearestLine(t.columnX, local.x) } };
      if (onTable || id === hit) return { id };
      // Outside the table and its bands, the selected table keeps its handle; a hovered one does not.
      if (id === selected) return null;
    }
    return null;
  }

  /** The chrome to paint for the selection and the pointer, in slide units. */
  elements(doc: Document, layout: SlideLayout): Element[] {
    const id = this.#tableId(layout);
    const frame = id ? tableItem(layout, id) : undefined;
    if (!id || !frame) return [];
    const px = EMU_PER_PX / this.host.zoom();
    const parts: Element[] = [this.#handle(doc, frame, id, px)];
    const hover = this.#hover?.id === id ? this.#hover : null;
    // A remote edit can remove the hovered line before the pointer moves again.
    const lines = (axis: 'row' | 'column') =>
      axis === 'row' ? frame.table.rowY.length : frame.table.columnX.length;
    if (hover?.line && hover.line.index < lines(hover.line.axis))
      parts.push(this.#lineMark(doc, frame, hover.line, px));
    if (
      hover?.insert &&
      hover.insert.index < lines(hover.insert.axis) &&
      insertCell(frame, hover.insert)
    )
      parts.push(...this.#insertButton(doc, frame, id, hover.insert, px));
    return parts;
  }

  /** The select handle: the `open_with` symbol in a box out from the table's top-left corner. */
  #handle(doc: Document, frame: TableFrame, id: string, px: number): Element {
    const corner = apply(frame.matrix, 0, 0);
    const size = HANDLE_PX * px;
    const off = (HANDLE_PX + HANDLE_GAP_PX) * px;
    const g = doc.createElementNS(SVG_NS, 'g');
    g.setAttribute('class', 'pptx-surface__table-handle');
    g.setAttribute('data-table-handle', id);
    // Pointer chrome: keyboard users select with Escape, the menus, and the toolbar.
    g.setAttribute('aria-hidden', 'true');
    g.setAttribute('transform', `translate(${corner.x - off} ${corner.y - off})`);
    const title = doc.createElementNS(SVG_NS, 'title');
    title.textContent = this.host.label('table.region.table');
    const box = doc.createElementNS(SVG_NS, 'rect');
    box.setAttribute('class', 'pptx-surface__table-handle-box');
    box.setAttribute('width', String(size));
    box.setAttribute('height', String(size));
    box.setAttribute('rx', String(2 * px));
    box.setAttribute('stroke-width', String(px));
    // The 960-unit symbol at 12 px, centered in the box.
    const icon = doc.createElementNS(SVG_NS, 'g');
    const k = (12 * px) / 960;
    icon.setAttribute('transform', `translate(${2 * px} ${14 * px}) scale(${k})`);
    icon.setAttribute('class', 'pptx-surface__table-handle-icon');
    for (const d of ICON_PATHS.open_with ?? []) {
      const path = doc.createElementNS(SVG_NS, 'path');
      path.setAttribute('d', d);
      icon.append(path);
    }
    g.append(title, box, icon);
    return g;
  }

  /** The grid line under the pointer, across the whole table: a drag there resizes. */
  #lineMark(
    doc: Document,
    frame: TableFrame,
    line: { readonly axis: 'column' | 'row'; readonly index: number },
    px: number
  ): Element {
    const t = frame.table;
    const at = line.axis === 'column' ? t.columnX[line.index] : t.rowY[line.index];
    const [x1, y1, x2, y2] = line.axis === 'column' ? [at, 0, at, t.height] : [0, at, t.width, at];
    const a = apply(frame.matrix, x1, y1);
    const b = apply(frame.matrix, x2, y2);
    const mark = doc.createElementNS(SVG_NS, 'line');
    mark.setAttribute('class', 'pptx-surface__table-line');
    mark.setAttribute('x1', String(a.x));
    mark.setAttribute('y1', String(a.y));
    mark.setAttribute('x2', String(b.x));
    mark.setAttribute('y2', String(b.y));
    mark.setAttribute('stroke-width', String(2 * px));
    return mark;
  }

  /** The "+" button of grid line `insert.index`, and a guide on the line where the row or column goes. */
  #insertButton(
    doc: Document,
    frame: TableFrame,
    id: string,
    insert: Insert,
    px: number
  ): Element[] {
    const t = frame.table;
    const lp = px / scaleOf(frame);
    const row = insert.axis === 'row';
    const at = row ? t.rowY[insert.index] : t.columnX[insert.index];
    const guideA = apply(frame.matrix, row ? 0 : at, row ? at : 0);
    const guideB = apply(frame.matrix, row ? t.width : at, row ? at : t.height);
    const guide = doc.createElementNS(SVG_NS, 'line');
    guide.setAttribute('class', 'pptx-surface__table-insert-guide');
    guide.setAttribute('x1', String(guideA.x));
    guide.setAttribute('y1', String(guideA.y));
    guide.setAttribute('x2', String(guideB.x));
    guide.setAttribute('y2', String(guideB.y));
    guide.setAttribute('stroke-width', String(2 * px));
    guide.setAttribute('stroke-dasharray', `${3 * px} ${2 * px}`);
    const c = apply(
      frame.matrix,
      row ? -INSERT_OFFSET_PX * lp : at,
      row ? at : -INSERT_OFFSET_PX * lp
    );
    const g = doc.createElementNS(SVG_NS, 'g');
    g.setAttribute('class', 'pptx-surface__table-insert');
    g.setAttribute('data-table-insert', `${id}|${insert.axis}|${insert.index}`);
    g.setAttribute('aria-hidden', 'true');
    const title = doc.createElementNS(SVG_NS, 'title');
    title.textContent = this.host.label(row ? 'table.insertRowBelow' : 'table.insertColumnRight');
    const dot = doc.createElementNS(SVG_NS, 'circle');
    dot.setAttribute('class', 'pptx-surface__table-insert-dot');
    dot.setAttribute('cx', String(c.x));
    dot.setAttribute('cy', String(c.y));
    dot.setAttribute('r', String(INSERT_RADIUS_PX * px));
    dot.setAttribute('stroke-width', String(px));
    const plus = doc.createElementNS(SVG_NS, 'path');
    plus.setAttribute('class', 'pptx-surface__table-insert-plus');
    const arm = 4 * px;
    plus.setAttribute('d', `M${c.x - arm} ${c.y}H${c.x + arm}M${c.x} ${c.y - arm}V${c.y + arm}`);
    plus.setAttribute('stroke-width', String(1.5 * px));
    g.append(title, dot, plus);
    return [guide, g];
  }

  /**
   * A press on `target`. Returns what it was for when it was on the chrome, else null: an insert
   * button inserts its row or column here; the select handle asks the surface to select and drag.
   */
  press(layout: SlideLayout, target: Element): TableChromePress | null {
    if (this.#off()) return null;
    const handle = target.closest?.('[data-table-handle]')?.getAttribute('data-table-handle');
    if (handle && tableItem(layout, handle)) {
      const group = parentGroupOf(layout.items, handle);
      return { kind: 'move', id: handle, scope: group ? { group } : {} };
    }
    const insert = target.closest?.('[data-table-insert]')?.getAttribute('data-table-insert');
    if (!insert) return null;
    const [id, axis, index] = insert.split('|');
    const frame = tableItem(layout, id);
    const at: Insert = { axis: axis === 'column' ? 'column' : 'row', index: Number(index) };
    const cell = frame ? insertCell(frame, at) : undefined;
    if (!cell) return { kind: 'insert' };
    const point = { paragraph: 0, offset: 0 };
    const before = this.host.selection();
    this.host.select({ kind: 'text', nodeId: id, cell, anchor: point, focus: point });
    const result = this.host.run(
      at.axis === 'row'
        ? { type: 'insertRow', where: 'below' }
        : { type: 'insertColumn', where: 'right' }
    );
    // A refused insert leaves the selection as it was. After an insert, the caret goes into the
    // new row or column, in the row or column of the cell it was inserted from: the next key
    // types there.
    if (result && !result.ok) this.host.select(before);
    else {
      const added =
        at.axis === 'row'
          ? { row: at.index, column: cell.column }
          : { row: cell.row, column: at.index };
      this.host.select({ kind: 'text', nodeId: id, cell: added, anchor: point, focus: point });
    }
    return { kind: 'insert' };
  }
}
