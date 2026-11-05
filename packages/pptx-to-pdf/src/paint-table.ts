/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A laid-out table as PDF paint, in the order of the editor's SVG table paint
// (`packages/core/src/output/paint-table.ts`): the table background, every cell's fill, then
// every border, then cell text. Borders paint after all fills, so a neighbor's fill never covers
// a border. Coordinates are the frame's local EMU.

import {
  concatTransformationMatrix,
  LineCapStyle,
  lineTo,
  moveTo,
  popGraphicsState,
  pushGraphicsState,
  setLineCap,
  stroke,
} from 'pdf-lib';
import type { ResolvedFill, ResolvedLine } from '@pptx-editor.dev/core/cascade';
import type { CellLayout, TableLayout } from '@pptx-editor.dev/core/layout';
import { paintFill, strokeStyle, type FillScope } from './paint-fill.ts';
import { drawText } from './paint-text.ts';
import { traceRect, type PageOps } from './page-ops.ts';

function edge(
  line: ResolvedLine | undefined,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  p: PageOps,
  scope: FillScope
): void {
  if (!line) return;
  p.push(pushGraphicsState());
  const color = strokeStyle(line, p, scope);
  if (color) {
    p.alpha(1, color.a);
    // Solid cell borders meet square, as in the SVG paint (measured on reference PDFs).
    if (line.dash === 'solid') p.push(setLineCap(LineCapStyle.Projecting));
    p.push(moveTo(x1, y1), lineTo(x2, y2), stroke());
  }
  p.push(popGraphicsState());
}

function borders(c: CellLayout, p: PageOps, scope: FillScope): void {
  const b = c.cell.borders;
  const r = c.x + c.width;
  const bottom = c.y + c.height;
  edge(b.tlToBr, c.x, c.y, r, bottom, p, scope);
  edge(b.blToTr, c.x, bottom, r, c.y, p, scope);
  edge(b.top, c.x, c.y, r, c.y, p, scope);
  edge(b.bottom, c.x, bottom, r, bottom, p, scope);
  edge(b.left, c.x, c.y, c.x, bottom, p, scope);
  edge(b.right, r, c.y, r, bottom, p, scope);
}

/** Draw a table in its frame's local space (the caller applies the frame's matrix). */
export async function drawTable(
  table: TableLayout,
  background: ResolvedFill,
  p: PageOps,
  scope: FillScope
): Promise<void> {
  const box = (w: number, h: number) => (): void => traceRect(0, 0, w, h, p);
  await paintFill(
    background,
    table.width,
    table.height,
    box(table.width, table.height),
    p,
    scope,
    'tables'
  );
  for (const c of table.cells) {
    // Gradients and pictures fill the cell's own box.
    p.push(pushGraphicsState(), concatTransformationMatrix(1, 0, 0, 1, c.x, c.y));
    await paintFill(
      c.cell.fill,
      c.width,
      c.height,
      box(c.width, c.height),
      p,
      scope,
      'table cells'
    );
    p.push(popGraphicsState());
  }
  for (const c of table.cells) borders(c, p, scope);
  for (const c of table.cells) if (c.text.lines.length) await drawText(c.text, p, scope);
}
