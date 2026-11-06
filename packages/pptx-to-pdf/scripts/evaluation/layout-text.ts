/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor `packages/docx-to-pdf/scripts/evaluation/layout-text.ts`, commit
// 398963c7f67207cebf3c38dcc41d5f0bd1339b5f. PPTX has no semantic span walker, so this walks
// the slide layout: shapes, groups, and table cells, in paint order.

import type { ResolvedShape } from '@pptx-editor.dev/core/cascade';
import {
  apply,
  multiply,
  type LayoutItem,
  type Matrix,
  type SlideLayout,
  type TextLayout,
} from '@pptx-editor.dev/core/layout';
import { ProtocolError } from './settings.ts';

const EMU_PER_PT = 12_700;
const MAX_LINES = 250_000;
const MAX_CHARACTERS = 8_000_000;

/** One laid-out line: its text and its box on the slide, in points from the top left. */
export interface LayoutTextLine {
  readonly text: string;
  /** The part and the engine node of the shape or table frame that holds the line. */
  readonly partName: string;
  readonly nodeId: string;
  /** Where the shape comes from: the slide, its layout, or its master. */
  readonly layer: 'master' | 'layout' | 'slide';
  /** Zero-based paragraph index in its text body. */
  readonly paragraph: number;
  /** `[x0, y0, x1, y1]`, the bounding box of the line after rotation, rounded to 0.01 pt. */
  readonly bbox: readonly [number, number, number, number];
}

const translate = (x: number, y: number): Matrix => [1, 0, 0, 1, x, y];

/** The text frame → slide matrix paint uses: the box offset, then the body rotation. */
function frameMatrix(outer: Matrix, text: TextLayout): Matrix {
  const { box } = text;
  let m = multiply(outer, translate(box.x, box.y));
  if (text.rotation !== 0) {
    const r = (text.rotation * Math.PI) / 180;
    const [cx, cy] = [box.width / 2, box.height / 2];
    m = multiply(m, translate(cx, cy));
    m = multiply(m, [Math.cos(r), Math.sin(r), -Math.sin(r), Math.cos(r), 0, 0]);
    m = multiply(m, translate(-cx, -cy));
  }
  return m;
}

const round = (emu: number): number => Math.round((emu / EMU_PER_PT) * 100) / 100;

/** Logical text evidence of one slide: not PDF extraction, and before paint. */
export function recordSlideText(layout: SlideLayout, budget = { lines: 0, characters: 0 }) {
  const lines: LayoutTextLine[] = [];
  const record = (
    text: TextLayout,
    outer: Matrix,
    source: Pick<ResolvedShape, 'partName' | 'nodeId' | 'layer'>
  ): void => {
    const m = frameMatrix(outer, text);
    for (const line of text.lines) {
      const runs = line.runs.filter((r) => r.text.length > 0);
      if (runs.length === 0) continue;
      const value = runs.map((r) => r.text).join('');
      budget.lines += 1;
      budget.characters += value.length;
      if (budget.lines > MAX_LINES || budget.characters > MAX_CHARACTERS)
        throw new ProtocolError('evidenceLimit', 'Layout text exceeds the evidence limit');
      let [x0, x1] = [Infinity, -Infinity];
      for (const r of runs) {
        x0 = Math.min(x0, r.x[0]!);
        x1 = Math.max(x1, r.x[r.x.length - 1]!);
      }
      const corners = [
        apply(m, x0, line.top),
        apply(m, x1, line.top),
        apply(m, x0, line.top + line.height),
        apply(m, x1, line.top + line.height),
      ];
      const xs = corners.map((c) => c.x);
      const ys = corners.map((c) => c.y);
      lines.push({
        text: value,
        partName: source.partName,
        nodeId: source.nodeId,
        layer: source.layer,
        paragraph: line.paragraph,
        bbox: [
          round(Math.min(...xs)),
          round(Math.min(...ys)),
          round(Math.max(...xs)),
          round(Math.max(...ys)),
        ],
      });
    }
  };
  // Hidden elements are not painted, so their text is not recorded either.
  const walk = (item: LayoutItem): void => {
    if (item.element.hidden) return;
    if (item.kind === 'group') for (const child of item.children) walk(child);
    else if (item.kind === 'shape' && item.text && item.textMatrix)
      record(item.text, item.textMatrix, item.element);
    else if (item.kind === 'frame' && item.table)
      for (const cell of item.table.cells) record(cell.text, item.matrix, item.element);
  };
  for (const item of layout.items) walk(item);
  return lines;
}
