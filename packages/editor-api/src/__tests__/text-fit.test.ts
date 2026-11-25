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

  test('empty paragraphs at the end count only when they move the text up', async () => {
    const lines = para(run('One', 'sz="2400"')) + para(run('Two', 'sz="2400"'));
    const empty = '<a:p><a:endParaRPr lang="en-US" sz="4000"/></a:p>';
    const anchored = (anchor: string) =>
      `<a:bodyPr wrap="square" anchor="${anchor}"><a:noAutofit/></a:bodyPr>`;
    const [plain, top, bottom, middle, columns] = await fitsOf(
      deckOf(
        { text: lines, bodyPr: NO_AUTOFIT },
        { text: lines + empty + empty, bodyPr: NO_AUTOFIT },
        { text: lines + empty + empty, bodyPr: anchored('b') },
        { text: lines + empty + empty, bodyPr: anchored('ctr') },
        {
          text: lines + empty + empty,
          bodyPr: '<a:bodyPr wrap="square" anchor="b" numCol="2"><a:noAutofit/></a:bodyPr>',
        }
      ),
      ['2', '3', '4', '5', '6']
    );
    // With columns, anchoring does not move the text: the empty paragraphs do not count.
    expect(columns!.overflows).toBe(false);
    expect(plain!.overflows).toBe(false);
    // Below top-anchored text, the empty paragraphs draw nothing.
    expect(top).toEqual(plain!);
    // Bottom or middle anchoring pushes "One" above the shape: the picture shows the overflow.
    expect(bottom!.overflows).toBe(true);
    expect(middle!.overflows).toBe(true);
    expect(bottom!.textHeight).toBeGreaterThan(plain!.textHeight);
  });

  test('a font write in the same sync re-fits the frame before the read', async () => {
    const runtime = await open(deckOf({ text: para(run('Grow')), bodyPr: SP_AUTOFIT }));
    const fit = await runtime.run(async (context) => {
      const frame = context.presentation.slides.getItemAt(0).shapes.getItem('2').textFrame;
      frame.textRange.font.size = 96;
      const result = frame.getTextFit();
      await context.sync();
      return result.value;
    });
    expect(fit.overflows).toBe(false);
    expect(fit.availableHeight).toBeGreaterThan(64.8);
    runtime.dispose();
  });

  test('shrunk bottom-anchored text with empty paragraphs at the end can overflow', async () => {
    const empty = '<a:p><a:endParaRPr lang="en-US" sz="4000"/></a:p>';
    const runtime = await open(
      deckOf({
        text: para(run(SHRINKS)) + empty + empty,
        bodyPr: '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>',
      })
    );
    const fit = await runtime.run(async (context) => {
      const frame = context.presentation.slides.getItemAt(0).shapes.getItem('2').textFrame;
      // The frame write re-fits: the shrink step fits the text without the empty paragraphs,
      // and the picture draws them below it, which pushes the text above the shape.
      frame.verticalAlignment = 'Bottom';
      const result = frame.getTextFit();
      await context.sync();
      return result.value;
    });
    expect(fit.overflows).toBe(true);
    runtime.dispose();
  });

  test('a placeholder with no text body reads as an empty frame', async () => {
    const runtime = await open(
      buildPptx({
        slides: [
          {
            layout: 'title',
            shapes:
              '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/>' +
              '<p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>',
          },
        ],
      })
    );
    const fit = await runtime.run(async (context) => {
      const result = context.presentation.slides
        .getItemAt(0)
        .shapes.getItem('2')
        .textFrame.getTextFit();
      await context.sync();
      return result.value;
    });
    expect(fit.overflows).toBe(false);
    expect(fit.textHeight).toBeGreaterThan(0);
    expect(fit.availableHeight).toBeGreaterThan(fit.textHeight);
    runtime.dispose();
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

  test('a table, a group, a hidden shape, and a layout shape refuse with NotSupported', async () => {
    const hidden = (xml: string, id: number) =>
      xml.replace(`name="Shape ${id}"/>`, `name="Shape ${id}" hidden="1"/>`);
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
            hidden(shape({ id: 3, x: IN, y: 3 * IN, cx: IN, cy: IN, text: para(run('x')) }), 3) +
            group(
              { id: 4, x: IN, y: 5 * IN, cx: IN, cy: IN },
              shape({ id: 5, x: IN, y: 5 * IN, cx: IN, cy: IN, text: para(run('y')) })
            ).replace('name="Group 4"/>', 'name="Group 4" hidden="1"/>'),
        },
      ],
    });
    const cases: [(c: RequestContext) => unknown, string][] = [
      [
        (c) => c.presentation.slides.getItemAt(0).shapes.getItem('2').textFrame.getTextFit(),
        'edit.no-text-frame',
      ],
      [
        (c) => c.presentation.slides.getItemAt(0).shapes.getItem('4').textFrame.getTextFit(),
        'edit.no-text-frame',
      ],
      [
        (c) => c.presentation.slides.getItemAt(0).shapes.getItem('3').textFrame.getTextFit(),
        'text-fit.not-drawn',
      ],
      [
        (c) =>
          c.presentation.slides
            .getItemAt(0)
            .shapes.getItem('4')
            .group.shapes.getItem('5')
            .textFrame.getTextFit(),
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
    const live = createPptxEditor({ substitutes: false });
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

  test('fixed metrics, which model no font, refuse the read', async () => {
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(deckOf({ text: para(run('a')), bodyPr: NO_AUTOFIT }))).ok).toBe(true);
    const runtime = BrowserEditor.createBrowser(live);
    const error = await rejection(
      runtime.run(async (context) => {
        context.presentation.slides.getItemAt(0).shapes.getItem('2').textFrame.getTextFit();
        await context.sync();
      })
    );
    expect(failure(error)).toEqual(['NotSupported', 'host.no-font-measurement']);
  });

  test('columns: text that flows into the second column reports the column height', async () => {
    const columns = '<a:bodyPr wrap="square" numCol="2"><a:noAutofit/></a:bodyPr>';
    const [two, over] = await fitsOf(
      deckOf(
        {
          text: para(
            run('The quick brown fox jumps over the lazy dog. A lazy dog sleeps.', 'sz="1400"')
          ),
          bodyPr: columns,
        },
        { text: para(run(LONG, 'sz="1800"')), bodyPr: columns }
      ),
      ['2', '3']
    );
    expect(two!.overflows).toBe(false);
    expect(two!.textHeight).toBe(two!.availableHeight);
    expect(over!.overflows).toBe(true);
    expect(over!.textHeight).toBeGreaterThan(over!.availableHeight);
  });

  test('a metric-compatible substitute is not listed', async () => {
    const runtime = await PptxEditor.createServer(
      deckOf({
        text:
          para(run('Named', 'sz="1800"', '<a:latin typeface="Calibri"/>')) +
          para(run('Other', 'sz="1800"', '<a:latin typeface="No Such Family"/>')),
        bodyPr: NO_AUTOFIT,
      })
    );
    const fit = await runtime.run(async (context) => {
      const result = context.presentation.slides
        .getItemAt(0)
        .shapes.getItem('2')
        .textFrame.getTextFit();
      await context.sync();
      return result.value;
    });
    expect(fit.substitutedFonts).toEqual(['No Such Family']);
    runtime.dispose();
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
