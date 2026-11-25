/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// TextFrame.getTextFit on a server runtime: the engine's own layout, measured with the runtime's
// measurer and no canvas, against the text area (ECMA-376 Part 1, 21.1.2.1.1 bodyPr).

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { createPptxEditor } from '@pptx-editor.dev/core/editor';
import { PptxEditor as BrowserEditor } from '../browser.ts';
import type { FontSource } from '@pptx-editor.dev/core/layout';
import {
  buildPptx,
  cell,
  group,
  para,
  run,
  shape,
  table,
} from '../../../../scripts/fixtures/builder.ts';
import { isPptxEditorError, PptxEditor, type RequestContext, type TextFrameFit } from '../index.ts';
import { IN, open, rejection } from './support.ts';

const LONG = 'The quick brown fox jumps over the lazy dog. '.repeat(12);
// Text that overflows the box at full size and fits at a shrink step.
const SHRINKS = 'The quick brown fox jumps over the lazy dog. '.repeat(3);
const NO_AUTOFIT = '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>';
const NORM_AUTOFIT = '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>';
const SP_AUTOFIT = '<a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr>';

/** One slide of three-inch by one-inch text boxes, one for each body, ids from 2. */
function deckOf(...boxes: { text: string; bodyPr: string }[]): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: boxes
          .map((b, i) =>
            shape({
              id: i + 2,
              x: IN,
              y: IN + i * 1.5 * IN,
              cx: 3 * IN,
              cy: IN,
              text: b.text,
              bodyPr: b.bodyPr,
            })
          )
          .join(''),
      },
    ],
  });
}

const failure = (error: unknown) =>
  isPptxEditorError(error) ? [error.code, error.debugInfo.innerError?.code] : [String(error)];

/** The fit of each shape id on the first slide, in one sync. */
async function fitsOf(bytes: Uint8Array, ids: readonly string[]): Promise<TextFrameFit[]> {
  const runtime = await open(bytes);
  try {
    return await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      const results = ids.map((id) => shapes.getItem(id).textFrame.getTextFit());
      await context.sync();
      return results.map((r) => r.value);
    });
  } finally {
    runtime.dispose();
  }
}

describe('TextFrame.getTextFit', () => {
  test('reports short and long text against the text area, in points', async () => {
    const [short, long] = await fitsOf(
      deckOf(
        { text: para(run('Short', 'sz="1800"')), bodyPr: NO_AUTOFIT },
        { text: para(run(LONG, 'sz="1800"')), bodyPr: NO_AUTOFIT }
      ),
      ['2', '3']
    );
    // One inch less the default top and bottom insets of 0.05 inch each.
    expect(short!.availableHeight).toBe(64.8);
    expect(short!.overflows).toBe(false);
    expect(short!.textHeight).toBeGreaterThan(0);
    expect(short!.textHeight).toBeLessThan(short!.availableHeight);
    expect(long!.overflows).toBe(true);
    expect(long!.textHeight).toBeGreaterThan(long!.availableHeight);
    expect(Object.isFrozen(long)).toBe(true);
  });

  test('a write in the same batch re-fits first; each autofit kind has its answer', async () => {
    const runtime = await open(
      deckOf(
        { text: para(run('a')), bodyPr: SP_AUTOFIT },
        { text: para(run('a')), bodyPr: NORM_AUTOFIT },
        { text: para(run('a')), bodyPr: NORM_AUTOFIT },
        { text: para(run('a')), bodyPr: NO_AUTOFIT }
      )
    );
    const fits = await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      const write = (id: string, text: string) => {
        const frame = shapes.getItem(id).textFrame;
        frame.textRange.text = text;
        return frame.getTextFit();
      };
      // Text that the grown shape holds, text that a shrink step makes fit, text that the last
      // step does not fit, and text that overflows with no autofit.
      const results = [
        write('2', LONG),
        write('3', SHRINKS),
        write('4', LONG.repeat(4)),
        write('5', LONG),
      ];
      await context.sync();
      return results.map((r) => r.value);
    });
    const [grown, shrunk, tooLong, raw] = fits;
    expect(grown!.overflows).toBe(false);
    expect(grown!.availableHeight).toBeGreaterThan(64.8);
    expect(shrunk!.overflows).toBe(false);
    expect(tooLong!.overflows).toBe(true);
    expect(raw!.overflows).toBe(true);
    runtime.dispose();
  });

  test('empty paragraphs at the end do not count', async () => {
    const lines = para(run('One', 'sz="2400"')) + para(run('Two', 'sz="2400"'));
    const empty = '<a:p><a:endParaRPr lang="en-US" sz="4000"/></a:p>';
    const [plain, trailing] = await fitsOf(
      deckOf(
        { text: lines, bodyPr: NO_AUTOFIT },
        { text: lines + empty + empty, bodyPr: NO_AUTOFIT }
      ),
      ['2', '3']
    );
    expect(trailing).toEqual(plain!);
  });

  test('vertical text measures across its lines, along the shape width', async () => {
    const [vertical] = await fitsOf(
      deckOf({
        text: para(run('Up', 'sz="1800"')),
        bodyPr: '<a:bodyPr vert="vert" wrap="square"><a:noAutofit/></a:bodyPr>',
      }),
      ['2']
    );
    // Three inches less the left and right insets of 0.1 inch each.
    expect(vertical!.availableHeight).toBe(201.6);
    expect(vertical!.overflows).toBe(false);
  });

  test('a shape in a group, and a shape with no text body, answer', async () => {
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            group(
              { id: 10, x: IN, y: IN, cx: 6 * IN, cy: 2 * IN },
              shape({
                id: 11,
                x: IN,
                y: IN,
                cx: 3 * IN,
                cy: IN,
                text: para(run(LONG)),
                bodyPr: NO_AUTOFIT,
              })
            ) + shape({ id: 12, x: IN, y: 4 * IN, cx: 3 * IN, cy: IN }),
        },
      ],
    });
    const runtime = await open(bytes);
    const [inGroup, bare] = await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      const a = shapes.getItem('10').group.shapes.getItem('11').textFrame.getTextFit();
      const b = shapes.getItem('12').textFrame.getTextFit();
      await context.sync();
      return [a.value, b.value];
    });
    expect(inGroup!.overflows).toBe(true);
    expect(bare).toMatchObject({ overflows: false, availableHeight: 64.8 });
    runtime.dispose();
  });

  test('lists the families that another face measured', async () => {
    const bytes = deckOf({
      text: para(run('Named', 'sz="1800"', '<a:latin typeface="No Such Family"/>')),
      bodyPr: NO_AUTOFIT,
    });
    const [fit] = await fitsOf(bytes, ['2']);
    expect(fit!.substitutedFonts).toContain('No Such Family');
  });

  test('a face that the runtime has is not listed', async () => {
    const path = join(import.meta.dir, '../../../../fixtures/fonts/STIXTwoMath-Regular.otf');
    const font: FontSource = {
      family: 'STIX Two Math',
      bold: false,
      italic: false,
      bytes: new Uint8Array(readFileSync(path)),
      origin: 'system',
    };
    const runtime = await PptxEditor.createServer(
      deckOf({
        text: para(run('Named', 'sz="1800"', '<a:latin typeface="STIX Two Math"/>')),
        bodyPr: NO_AUTOFIT,
      }),
      { substitutes: false, fonts: font }
    );
    const fit = await runtime.run(async (context) => {
      const result = context.presentation.slides
        .getItemAt(0)
        .shapes.getItem('2')
        .textFrame.getTextFit();
      await context.sync();
      return result.value;
    });
    expect(fit.substitutedFonts).not.toContain('STIX Two Math');
    runtime.dispose();
  });

  test('a table, a hidden shape, and a layout shape refuse with NotSupported', async () => {
    const bytes = buildPptx({
      slides: [
        {
          layout: 'title',
          shapes:
            table({
              id: 2,
              x: IN,
              y: IN,
              cx: 2 * IN,
              cy: IN,
              columns: [2 * IN],
              tblPr: '<a:tblPr/>',
              rows: `<a:tr h="${IN}">${cell('Cell')}</a:tr>`,
            }) +
            shape({ id: 3, x: IN, y: 3 * IN, cx: IN, cy: IN, text: para(run('x')) }).replace(
              '<p:cNvPr id="3" name="Shape 3"/>',
              '<p:cNvPr id="3" name="Shape 3" hidden="1"/>'
            ),
        },
      ],
    });
    const cases: [(c: RequestContext) => unknown, string][] = [
      [
        (c) => c.presentation.slides.getItemAt(0).shapes.getItem('2').textFrame.getTextFit(),
        'edit.no-text-frame',
      ],
      [
        (c) => c.presentation.slides.getItemAt(0).shapes.getItem('3').textFrame.getTextFit(),
        'text-fit.not-drawn',
      ],
      [
        (c) => c.presentation.slides.getItemAt(0).layout.shapes.getItemAt(0).textFrame.getTextFit(),
        'text-fit.not-on-slide',
      ],
    ];
    for (const [read, inner] of cases) {
      const runtime = await open(bytes);
      const error = await rejection(
        runtime.run(async (context) => {
          read(context);
          await context.sync();
        })
      );
      expect(failure(error)).toEqual(['NotSupported', inner]);
      runtime.dispose();
    }
  });

  test('a runtime over a live editor reads the fit and records no undo step', async () => {
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(deckOf({ text: para(run(LONG)), bodyPr: NO_AUTOFIT }))).ok).toBe(true);
    const runtime = BrowserEditor.createBrowser(live);
    const fit = await runtime.run(async (context) => {
      const result = context.presentation.slides
        .getItemAt(0)
        .shapes.getItem('2')
        .textFrame.getTextFit();
      await context.sync();
      return result.value;
    });
    expect(fit.overflows).toBe(true);
    expect(live.snapshot().canUndo).toBe(false);
  });

  test('a null text frame skips the read, as other reads of a null object', async () => {
    const runtime = await open(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: group(
              { id: 10, x: IN, y: IN, cx: 2 * IN, cy: IN },
              shape({ id: 11, x: IN, y: IN, cx: IN, cy: IN, text: para(run('x')) }) +
                shape({ id: 12, x: 2 * IN, y: IN, cx: IN, cy: IN, text: para(run('y')) })
            ),
          },
        ],
      })
    );
    const [isNull, loaded] = await runtime.run(async (context) => {
      const frame = context.presentation.slides
        .getItemAt(0)
        .shapes.getItem('10')
        .getTextFrameOrNullObject();
      const fit = frame.getTextFit();
      await context.sync();
      return [frame.isNullObject, fit.isLoaded];
    });
    expect([isNull, loaded]).toEqual([true, false]);
    runtime.dispose();
  });
});
