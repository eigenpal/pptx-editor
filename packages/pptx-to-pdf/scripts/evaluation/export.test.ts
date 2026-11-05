/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor `packages/docx-to-pdf/scripts/evaluation/export.test.ts` and
// `layout-summary.test.ts`, commit 398963c7f67207cebf3c38dcc41d5f0bd1339b5f: the evaluation
// export and the layout summary agree with the public API on the synthetic fixtures.

import { describe, expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { openPresentation } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { exportPdf } from '../../src/index.ts';
import { exportForEvaluation } from './export.ts';
import { summarizePages } from './layout-summary.ts';
import { recordSlideText } from './layout-text.ts';
import {
  errorCode,
  EXPORT_SETTINGS,
  loadFontPlan,
  openFailure,
  parseArguments,
  SESSION_SETTINGS,
} from './settings.ts';

const FIXTURES = join(import.meta.dir, '../../../../fixtures');
const deck = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(FIXTURES, `${name}.pptx`)));
const IN = 914_400;

/** Three text slides; the second is hidden. No fixture deck has a hidden slide. */
const withHidden = (): Uint8Array =>
  buildPptx({
    slides: [1, 2, 3].map((n) => ({
      layout: 'blank' as const,
      hidden: n === 2,
      shapes: shape({
        id: 2,
        x: IN,
        y: IN,
        cx: 6 * IN,
        cy: IN,
        text: para(run(`Slide ${n} words`)),
      }),
    })),
  });

describe('evaluation export', () => {
  test('equals exportPdf with the protocol settings, byte for byte', async () => {
    for (const name of ['text-runs', 'preset-shapes', 'pictures']) {
      const { bytes, report } = await exportForEvaluation(deck(name));
      const direct = await exportPdf(deck(name), { ...SESSION_SETTINGS, ...EXPORT_SETTINGS });
      expect(Buffer.from(bytes).equals(Buffer.from(direct.bytes))).toBe(true);
      expect(report.diagnostics).toEqual(direct.diagnostics);
      expect(report.pages).toBe(direct.pageCount);
      expect(report.timings.openMs).toBeGreaterThan(0);
    }
  });

  test('maps pages to slides and leaves hidden slides out, as the reference PDF does', async () => {
    const { report } = await exportForEvaluation(withHidden());
    expect(report.slideCount).toBe(3);
    expect(report.hiddenSlides).toEqual([2]);
    expect(report.slideNumbers).toEqual([1, 3]);
    expect(report.pages).toBe(2);
  });

  test('approximate output is reported, never hidden', async () => {
    const { report } = await exportForEvaluation(deck('effects'));
    const severity = (code: string) => report.diagnostics.find((d) => d.code === code)?.severity;
    expect(severity('effect.shadow-blur')).toBe('approximation');
    expect(severity('effect.reflection')).toBe('unsupported');
  });
});

describe('layout summary', () => {
  test('agrees with the export on pages and slides', async () => {
    for (const bytes of [deck('slide-structure'), deck('tables'), withHidden()]) {
      const summary = await summarizePages(bytes);
      const { report } = await exportForEvaluation(bytes);
      expect(summary.pageCount).toBe(report.pages);
      expect(summary.slideNumbers).toEqual([...report.slideNumbers]);
      expect(summary.hiddenSlides).toEqual(report.hiddenSlides);
      expect(summary.fontResolution).toEqual(report.fontResolution);
      for (const page of summary.pages) {
        expect(page.width).toBe(960);
        expect(page.height).toBe(540);
      }
    }
  });

  test('records logical text with line boxes on the slide, table cells included', async () => {
    const summary = await summarizePages(deck('tables'));
    const lines = summary.text.pages[0]!.lines;
    const texts = lines.map((l) => l.text);
    expect(texts).toContain('Region');
    expect(texts).toContain('North');
    for (const line of lines) {
      const [x0, y0, x1, y1] = line.bbox;
      expect(x0).toBeGreaterThanOrEqual(0);
      expect(y0).toBeGreaterThanOrEqual(0);
      expect(x1).toBeGreaterThan(x0);
      expect(y1).toBeGreaterThan(y0);
      expect(x1).toBeLessThanOrEqual(960);
      expect(y1).toBeLessThanOrEqual(540);
      expect(line.partName.startsWith('/ppt/')).toBe(true);
    }
    // A cell to the right of another cell on the same row has a larger x.
    const region = lines.find((l) => l.text === 'Region')!;
    const q1 = lines.find((l) => l.text === 'Q1')!;
    expect(q1.bbox[0]).toBeGreaterThan(region.bbox[2]);
    expect(Math.abs(q1.bbox[1] - region.bbox[1])).toBeLessThan(1);
  });

  test('a shape line box sits where the shape is', async () => {
    const summary = await summarizePages(withHidden());
    const [first] = summary.text.pages[0]!.lines;
    expect(first!.text).toBe('Slide 1 words');
    // The shape is at one inch (72 pt) from the top left, inside its default insets.
    expect(first!.bbox[0]).toBeGreaterThanOrEqual(72);
    expect(first!.bbox[0]).toBeLessThan(90);
    expect(first!.bbox[1]).toBeGreaterThanOrEqual(72);
    expect(first!.bbox[3]).toBeLessThan(144);
    expect(summary.text.pages.map((p) => p.slideNumber)).toEqual([1, 3]);
  });

  test('refuses an invalid deck with the typed open error', async () => {
    await expect(summarizePages(new Uint8Array([1, 2, 3]))).rejects.toMatchObject({
      code: 'documentOpenFailed',
    });
  });

  test('a vertical text body records its line box turned with it', async () => {
    const summary = await summarizePages(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              x: IN,
              y: IN,
              cx: IN,
              cy: 4 * IN,
              bodyPr: '<a:bodyPr vert="vert" wrap="square"/>',
              text: para(run('Vertical words')),
            }),
          },
        ],
      })
    );
    const [line] = summary.text.pages[0]!.lines;
    const [x0, y0, x1, y1] = line!.bbox;
    // Turned 90°: the line runs down the tall shape, inside its one-inch-wide box.
    expect(y1 - y0).toBeGreaterThan(x1 - x0);
    expect(x0).toBeGreaterThanOrEqual(72);
    expect(x1).toBeLessThanOrEqual(144);
    expect(y0).toBeGreaterThanOrEqual(72);
  });

  test('reports a protocol deadline as a timeout in every phase', async () => {
    await expect(exportForEvaluation(deck('slide-structure'), undefined, 1)).rejects.toMatchObject({
      code: 'timeout',
    });
    await expect(summarizePages(deck('slide-structure'), undefined, 1)).rejects.toMatchObject({
      code: 'timeout',
    });
  });

  test('reports an open past the protocol deadline as a timeout', () => {
    expect(openFailure({ ok: false, reason: 'aborted' })).toMatchObject({ code: 'timeout' });
    expect(openFailure({ ok: false, reason: 'zip', detail: 'bad' })).toMatchObject({
      code: 'documentOpenFailed',
      reason: 'zip',
    });
  });
});

/** A packaged substitute family's regular and bold faces, standing in for the reference fonts. */
async function substituteFaces(): Promise<{ regular: Uint8Array; bold: Uint8Array }> {
  const opened = await openPresentation(deck('text-runs'));
  if (!opened.ok) throw new Error('fixture');
  const carlito = opened.session.faces.filter((f) => f.family === 'Carlito' && !f.italic);
  const regular = carlito.find((f) => !f.bold)?.bytes;
  const bold = carlito.find((f) => f.bold)?.bytes;
  opened.session.dispose();
  if (!regular || !bold) throw new Error('no substitute faces');
  return { regular, bold };
}

describe('protocol settings', () => {
  test('arguments: two paths and at most one font plan, with clear errors', () => {
    const usage = 'in out [--fonts plan.json]';
    expect(parseArguments(['a', 'b', '--fonts', 'p.json'], usage)).toEqual({
      input: 'a',
      output: 'b',
      fonts: 'p.json',
    });
    expect(() => parseArguments(['a', 'b', '--fonts', 'p', '--fonts', 'q'], usage)).toThrow(
      '--fonts is given twice'
    );
    expect(() => parseArguments(['a', 'b', '--fonts'], usage)).toThrow('--fonts needs a path');
    expect(() => parseArguments(['a', 'b', '--size'], usage)).toThrow('Unknown option --size');
    expect(() => parseArguments(['a'], usage)).toThrow('Expected');
    try {
      parseArguments(['a'], usage);
    } catch (error) {
      expect(errorCode(error)).toBe('usage');
    }
  });

  test('every error has a code: font plan errors are fontPlan, engine failures internal', async () => {
    const folder = mkdtempSync(join(tmpdir(), 'pptx-plan-'));
    try {
      const refusal = async (text: string | null) => {
        const path = join(folder, 'plan.json');
        if (text === null) rmSync(path, { force: true });
        else writeFileSync(path, text);
        return loadFontPlan(path).then(
          () => 'loaded',
          (error: unknown) => errorCode(error)
        );
      };
      const face = (extra: object) =>
        JSON.stringify({
          protocol: 1,
          faces: [{ path: 'none.ttf', family: 'Calibri', bold: false, italic: false, ...extra }],
        });
      expect(await refusal(null)).toBe('fontPlan'); // missing plan: not the deck's ENOENT
      expect(await refusal('{')).toBe('fontPlan');
      expect(await refusal(JSON.stringify({ protocol: 2, faces: [] }))).toBe('fontPlan');
      expect(await refusal(JSON.stringify({ protocol: 1, faces: [] }))).toBe('fontPlan');
      expect(await refusal(face({ bold: 'yes' }))).toBe('fontPlan');
      expect(await refusal(face({}))).toBe('fontPlan'); // missing face file
      writeFileSync(join(folder, 'big.json'), ' '.repeat(1024 * 1024 + 1));
      expect(
        await loadFontPlan(join(folder, 'big.json')).then(
          () => 'loaded',
          (error: unknown) => errorCode(error)
        )
      ).toBe('fontPlan');
    } finally {
      rmSync(folder, { recursive: true, force: true });
    }
    expect(errorCode(new TypeError('engine'))).toBe('internal');
    expect(errorCode(Object.assign(new Error('missing deck'), { code: 'ENOENT' }))).toBe('ENOENT');
  });

  test('layout text past the evidence limit is refused with its code', async () => {
    const opened = await openPresentation(deck('text-runs'));
    if (!opened.ok) throw new Error('fixture');
    const layout = opened.session.layout(0)!;
    opened.session.dispose();
    let code: string | undefined;
    try {
      recordSlideText(layout, { lines: 250_000, characters: 0 });
    } catch (error) {
      code = errorCode(error);
    }
    expect(code).toBe('evidenceLimit');
  });

  test('the font plan identity follows the face bytes and the plan order', async () => {
    const folder = mkdtempSync(join(tmpdir(), 'pptx-plan-'));
    try {
      const faces = await substituteFaces();
      writeFileSync(join(folder, 'a.ttf'), faces.regular);
      writeFileSync(join(folder, 'b.ttf'), faces.bold);
      const plan = (faces: [string, boolean][]) =>
        JSON.stringify({
          protocol: 1,
          faces: faces.map(([path, bold]) => ({ path, family: 'Calibri', bold, italic: false })),
        });
      const identity = async (text: string) => {
        writeFileSync(join(folder, 'plan.json'), text);
        return (await loadFontPlan(join(folder, 'plan.json'))).identity.sha256;
      };
      const first = await identity(
        plan([
          ['a.ttf', false],
          ['b.ttf', true],
        ])
      );
      expect(
        await identity(
          plan([
            ['a.ttf', false],
            ['b.ttf', true],
          ])
        )
      ).toBe(first);
      expect(
        await identity(
          plan([
            ['b.ttf', true],
            ['a.ttf', false],
          ])
        )
      ).not.toBe(first);
      writeFileSync(join(folder, 'a.ttf'), faces.bold);
      expect(
        await identity(
          plan([
            ['a.ttf', false],
            ['b.ttf', true],
          ])
        )
      ).not.toBe(first);
      // A folder is not a font file.
      await expect(identity(plan([['.', false]]))).rejects.toThrow('not a regular file');
    } finally {
      rmSync(folder, { recursive: true, force: true });
    }
  });
});
