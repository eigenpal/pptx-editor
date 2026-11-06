/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A chart is laid out once, by the Pro chart renderer, and drawn by both the editor's SVG
// paint and the PDF writer. These tests paint each chart of `fixtures/charts.pptx` both ways
// and fail when the two disagree on the colors they fill or the text they set.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { Window } from 'happy-dom';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { openPresentation } from '@pptx-editor.dev/core/editor';
import { paintSlide } from '@pptx-editor.dev/core/output';
import { installDrawing } from '@pptx-editor.dev/pro/drawing';
import { exportPdf } from '../index.ts';

const deck = new Uint8Array(
  readFileSync(join(import.meta.dir, '../../../../fixtures/charts.pptx'))
);
const BEST = { fidelityPolicy: 'best-effort' } as const;

interface Painted {
  readonly fills: Set<string>;
  readonly text: string[];
}

const words = (text: string): string[] => text.split(/\s+/).filter(Boolean);

/** The colors the editor's SVG fills inside each chart, and the words it sets. */
async function svgCharts(): Promise<Painted[]> {
  installDrawing();
  const opened = await openPresentation(deck);
  if (!opened.ok) throw new Error('deck');
  const { document } = new Window();
  const out: Painted[] = [];
  for (let index = 0; index < 8; index += 1) {
    const root = paintSlide(document as unknown as Document, opened.session.layout(index)!, {
      images: () => 'blob:picture',
      idPrefix: `c${index}`,
      hitTargets: false,
    }) as unknown as Element;
    const chart = root.querySelector('[data-kind="chart"]');
    if (!chart) throw new Error(`slide ${index + 1}: no chart group`);
    const fills = new Set<string>();
    for (const path of chart.querySelectorAll('path')) {
      const fill = path.getAttribute('fill');
      if (fill && fill.startsWith('#')) fills.add(fill.toLowerCase());
    }
    const text = [...chart.querySelectorAll('text')].flatMap((t) => words(t.textContent ?? ''));
    out.push({ fills, text });
  }
  opened.session.dispose();
  return out;
}

const hex = (r: number, g: number, b: number) =>
  `#${[r, g, b].map((v) => Math.round(v).toString(16).padStart(2, '0')).join('')}`;

/** The colors each PDF page fills and the words it sets. */
async function pdfPages(): Promise<Painted[]> {
  const { bytes } = await exportPdf(deck, BEST);
  const pdf = await getDocument({ data: bytes.slice(), useSystemFonts: false }).promise;
  const out: Painted[] = [];
  for (let page = 1; page <= pdf.numPages; page += 1) {
    const p = await pdf.getPage(page);
    const ops = await p.getOperatorList();
    const fills = new Set<string>();
    ops.fnArray.forEach((fn, i) => {
      if (fn !== OPS.setFillRGBColor) return;
      const args = ops.argsArray[i] as unknown[];
      if (typeof args[0] === 'string') fills.add(String(args[0]).toLowerCase());
      else fills.add(hex(Number(args[0]), Number(args[1]), Number(args[2])));
    });
    const content = await p.getTextContent();
    const text = content.items.flatMap((item) => ('str' in item ? words(item.str) : []));
    out.push({ fills, text });
  }
  return out;
}

describe('the PDF writer and the SVG paint draw one chart layout', () => {
  test('on the fill colors and the text of every fixture chart', async () => {
    const [svg, pdf] = await Promise.all([svgCharts(), pdfPages()]);
    expect(pdf).toHaveLength(svg.length);
    for (let index = 0; index < svg.length; index += 1) {
      const { fills, text } = svg[index]!;
      expect(fills.size).toBeGreaterThan(1);
      expect(text.length).toBeGreaterThan(3);
      for (const fill of fills) expect([index, pdf[index]!.fills.has(fill)]).toEqual([index, true]);
      // The PDF page also holds the slide title; every chart word is on it, as often.
      const counts = new Map<string, number>();
      for (const word of pdf[index]!.text) counts.set(word, (counts.get(word) ?? 0) + 1);
      for (const word of text) {
        const left = counts.get(word) ?? 0;
        expect([index, word, left > 0]).toEqual([index, word, true]);
        counts.set(word, left - 1);
      }
    }
  });
});
