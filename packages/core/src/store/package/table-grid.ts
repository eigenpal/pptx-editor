// The merge grid of a DrawingML table (`a:tbl`): which grid positions each drawn cell covers.
// One rule for the cascade, the table commands, find, and the store part checks, so what the
// slide draws, what a command edits, and what a commit accepts never disagree. Format-neutral:
// DrawingML only, no PresentationML names.
//
// ECMA-376 Part 1, 21.1.3.16 tc (`CT_TableCell`): a cell spans `gridSpan` columns and `rowSpan`
// rows (both default 1), and the cells it covers carry `hMerge` or `vMerge`. A valid file holds
// one rectangle per merged block, every covered cell marked. The schema does not say what a
// file means when the markup disagrees, so this module picks one rule:
//
// - Coverage comes from spans only. Positions are read row by row; the first position that no
//   earlier cell covers is a drawn cell, and its `gridSpan` x `rowSpan` rectangle covers the
//   rest. A merge marker never hides a cell by itself.
// - An orphan marker (an `hMerge` cell that no span from the left covers, a `vMerge` cell that
//   no span from above covers) is a drawn cell with its own text, properties, and spans.
// - A span that runs past the grid ends at the grid's last row or column.
// - A span that runs into a position an earlier cell covers ends before that position: the
//   earlier cell (in reading order) keeps it, so drawn cells never overlap.
// - A position with no `a:tc` (a short row) is not drawn and covers nothing.

import { NS } from './namespaces.ts';
import { child, elements, textOf } from './query.ts';
import type { Attr, Element } from './tree.ts';

/** The most rows of a table the cascade draws. @internal */
export const TABLE_MAX_ROWS = 1000;
/** The most grid columns of a table the cascade draws. @internal */
export const TABLE_MAX_COLUMNS = 256;

/** One drawn cell of a table and the grid rectangle it covers. @internal */
export interface TableGridBlock {
  /** The drawn cell's grid row. */
  readonly row: number;
  /** The drawn cell's grid column. */
  readonly col: number;
  /** Rows covered, after clipping: at least 1. */
  readonly rowSpan: number;
  /** Columns covered, after clipping: at least 1. */
  readonly gridSpan: number;
  /** The cell's own `gridSpan` or `rowSpan` runs past the grid. */
  readonly pastGrid: boolean;
  /** The cell's span runs into a position that an earlier cell covers. */
  readonly overlaps: boolean;
}

/** The merge grid of a table. @internal */
export interface TableGrid {
  /** Grid rows read. */
  readonly rows: number;
  /** Grid columns read. */
  readonly columns: number;
  /** Drawn cells, in reading order. */
  readonly blocks: readonly TableGridBlock[];
  /** For each position (`row * columns + col`), its block's index, or -1 (no `a:tc`). */
  readonly owner: Int32Array;
}

/**
 * A cell as the grid reads it: its attributes. An `a:tc` element is one; so is a clipboard
 * spec's attribute list. @internal
 */
export interface TableGridCell {
  readonly attrs: readonly Attr[];
}

/**
 * A span attribute as the grid reads it (`xsd:int`, as `intAttr` reads one): an absent or
 * unreadable value, or one below 1, is 1.
 */
const tableSpan = (tc: TableGridCell, name: 'gridSpan' | 'rowSpan'): number => {
  const raw = tc.attrs.find((a) => a.ns === '' && a.local === name)?.value.trim();
  if (raw === undefined || !/^-?\d{1,15}$/.test(raw)) return 1;
  const v = Number(raw);
  return Number.isSafeInteger(v) && v > 1 ? v : 1;
};

/**
 * The merge grid of a table. `cells[r][c]` is the `a:tc` at grid position (r, c), or undefined
 * where a row is short; positions past `columns` are ignored. The caller bounds `cells.length`
 * and `columns` (`TABLE_MAX_ROWS`, `TABLE_MAX_COLUMNS`, or one `a:tc` per position), because
 * the grid allocates one entry per position. @internal
 */
export function resolveTableGrid(
  cells: readonly (readonly (TableGridCell | undefined)[])[],
  columns: number
): TableGrid {
  const nR = cells.length;
  const nC = Math.max(0, columns);
  const owner = new Int32Array(nR * nC).fill(-1);
  const blocks: TableGridBlock[] = [];
  for (let r = 0; r < nR; r += 1)
    for (let c = 0; c < nC; c += 1) {
      if (owner[r * nC + c] !== -1) continue;
      const tc = cells[r][c];
      if (!tc) continue;
      const wantG = tableSpan(tc, 'gridSpan');
      const wantK = tableSpan(tc, 'rowSpan');
      let g = Math.min(wantG, nC - c);
      let k = Math.min(wantK, nR - r);
      const fits = { g, k };
      for (let dc = 1; dc < g; dc += 1)
        if (owner[r * nC + c + dc] !== -1) {
          g = dc;
          break;
        }
      for (let dr = 1; dr < k; dr += 1) {
        let free = true;
        for (let dc = 0; dc < g && free; dc += 1) free = owner[(r + dr) * nC + c + dc] === -1;
        if (!free) {
          k = dr;
          break;
        }
      }
      const index = blocks.length;
      blocks.push({
        row: r,
        col: c,
        rowSpan: k,
        gridSpan: g,
        pastGrid: wantG > nC - c || wantK > nR - r,
        overlaps: g < fits.g || k < fits.k,
      });
      for (let dr = 0; dr < k; dr += 1)
        for (let dc = 0; dc < g; dc += 1) owner[(r + dr) * nC + c + dc] = index;
    }
  return { rows: nR, columns: nC, blocks, owner };
}

/**
 * How a cell's merge markup disagrees with the grid (21.1.3.16 tc). `past-grid` and `overlap`
 * are about a drawn cell's spans; `orphan` is a drawn cell with `hMerge` or `vMerge`;
 * `unmarked` is a covered cell with neither; `mismarked` is a covered cell whose markers do
 * not match its place in its block (`hMerge` off the block's first column, `vMerge` off its
 * first row). @internal
 */
export type TableMergeFault = 'past-grid' | 'overlap' | 'orphan' | 'unmarked' | 'mismarked';

/** One fault of a table's merge markup, at a grid position. @internal */
export interface TableMergeFaultAt {
  /** The cell's grid row. */
  readonly row: number;
  /** The cell's grid column. */
  readonly col: number;
  /** What disagrees. */
  readonly fault: TableMergeFault;
}

/** A merge marker as the grid reads it (`xsd:boolean`, as `boolAttr` reads one). */
const tableFlag = (tc: TableGridCell, name: 'hMerge' | 'vMerge'): boolean => {
  const raw = tc.attrs.find((a) => a.ns === '' && a.local === name)?.value;
  return raw === '1' || raw === 'true';
};

/**
 * The faults of a table's merge markup against `grid`, the `resolveTableGrid` of the same
 * `cells`. A table without faults holds one rectangle per merged block, every covered cell
 * marked and no drawn cell marked. Positions without an `a:tc` are skipped. @internal
 */
export function tableMergeFaults(
  cells: readonly (readonly (TableGridCell | undefined)[])[],
  grid: TableGrid
): TableMergeFaultAt[] {
  const out: TableMergeFaultAt[] = [];
  for (const b of grid.blocks) {
    if (b.pastGrid) out.push({ row: b.row, col: b.col, fault: 'past-grid' });
    if (b.overlaps) out.push({ row: b.row, col: b.col, fault: 'overlap' });
  }
  for (let r = 0; r < grid.rows; r += 1)
    for (let c = 0; c < grid.columns; c += 1) {
      const owner = grid.owner[r * grid.columns + c];
      const tc = cells[r][c];
      if (owner === -1 || !tc) continue;
      const h = tableFlag(tc, 'hMerge');
      const v = tableFlag(tc, 'vMerge');
      const b = grid.blocks[owner];
      if (b.row === r && b.col === c) {
        if (h || v) out.push({ row: r, col: c, fault: 'orphan' });
      } else if (!h && !v) out.push({ row: r, col: c, fault: 'unmarked' });
      else if (h !== c > b.col || v !== r > b.row) out.push({ row: r, col: c, fault: 'mismarked' });
    }
  return out;
}

/**
 * Whether a cell holds text: a run or a field with text, or a line break, in a paragraph of
 * its `a:txBody` (21.1.3.16 tc, 21.1.2.2.6 p). The table commands, the save, and the editor's
 * grid read text with this one rule. @internal
 */
export function tableCellHasText(tc: Element | undefined): boolean {
  for (const p of elements(child(tc, 'a:txBody'))) {
    if (p.ns !== NS.a || p.local !== 'p') continue;
    for (const el of elements(p)) {
      if (el.ns !== NS.a) continue;
      if (el.local === 'br') return true;
      if ((el.local === 'r' || el.local === 'fld') && textOf(child(el, 'a:t')).length > 0)
        return true;
    }
  }
  return false;
}

/** One span that a table command ends early (`tableSpanTrims`). @internal */
export interface TableSpanTrim {
  /** The cell's grid row. */
  readonly row: number;
  /** The cell's grid column. */
  readonly col: number;
  /** The span attribute that ends early. */
  readonly name: 'gridSpan' | 'rowSpan';
  /** The new span. 1 removes the attribute. */
  readonly value: number;
}

/**
 * The spans that a table command ends early. A span ends before the first cell it claims that
 * has text and no merge marker on that axis (`hMerge` for `gridSpan`, `vMerge` for `rowSpan`):
 * a merge and a concurrent row or column delete leave a span over a cell that no one merged,
 * and that cell stays visible. A blank cell without a marker stays covered, as a merge that
 * writes only spans leaves it. A cell with a merge marker is never trimmed. Every trim reads
 * `cells` as given, so the trims apply in any order, and a second pass finds none. A table
 * without faults (`tableMergeFaults`) has no trims. @internal
 */
export function tableSpanTrims<T extends TableGridCell>(
  cells: readonly (readonly (T | undefined)[])[],
  hasText: (tc: T) => boolean
): TableSpanTrim[] {
  const out: TableSpanTrim[] = [];
  const shows = (tc: T | undefined, marker: 'hMerge' | 'vMerge'): boolean =>
    !!tc && !tableFlag(tc, marker) && hasText(tc);
  cells.forEach((row, r) =>
    row.forEach((tc, c) => {
      if (!tc || tableFlag(tc, 'hMerge') || tableFlag(tc, 'vMerge')) return;
      const g = tableSpan(tc, 'gridSpan');
      for (let k = 1; k < g && c + k < row.length; k += 1)
        if (shows(row[c + k], 'hMerge')) {
          out.push({ row: r, col: c, name: 'gridSpan', value: k });
          break;
        }
      const h = tableSpan(tc, 'rowSpan');
      for (let j = 1; j < h && r + j < cells.length; j += 1)
        if (shows(cells[r + j][c], 'vMerge')) {
          out.push({ row: r, col: c, name: 'rowSpan', value: j });
          break;
        }
    })
  );
  return out;
}

/**
 * The merge grid that a table command writes for `cells`: `resolveTableGrid` after the
 * `tableSpanTrims`. A table without faults has no trims, so its grid is `resolveTableGrid`. The
 * save writes this grid into a table with a structure mark, and the editor draws and edits such
 * a table by it (`tableGridOf`), so the canvas shows what the save writes. @internal
 */
export function editedTableGrid<T extends TableGridCell>(
  cells: readonly (readonly (T | undefined)[])[],
  columns: number,
  hasText: (tc: T) => boolean
): TableGrid {
  const trims = tableSpanTrims(cells, hasText);
  if (trims.length === 0) return resolveTableGrid(cells, columns);
  const next: (TableGridCell | undefined)[][] = cells.map((row) => [...row]);
  for (const t of trims) {
    const attrs = next[t.row][t.col]!.attrs.filter((a) => !(a.ns === '' && a.local === t.name));
    if (t.value > 1) attrs.push({ ns: '', local: t.name, value: String(t.value) });
    next[t.row][t.col] = { attrs };
  }
  return resolveTableGrid(next, columns);
}
