/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The paragraph spacing and indent members of `ParagraphFormat`, SDK extensions over `a:pPr`
// (CT_TextParagraphProperties, ECMA-376 Part 1, 21.1.2.2.7): line spacing (`a:lnSpc`,
// 21.1.2.2.5), space before and after (`a:spcBef`, 21.1.2.2.10; `a:spcAft`, 21.1.2.2.9), each
// one `a:spcPct` (21.1.2.2.11) or `a:spcPts` (21.1.2.2.12), and the `marL` and `indent`
// attributes. Reads go through the text style cascade; writes go through `setParagraph`.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { FontSource } from '@pptx-editor.dev/core/layout';
import { createStoreAutomationHost, type AutomationHost } from '@pptx-editor.dev/core/automation';
import { createPptxEditor } from '@pptx-editor.dev/core/editor';
import { PptxDocumentStore, readPptx, readZip } from '@pptx-editor.dev/core/store';
import { buildPptx, cell, para, run, shape, table } from '../../../../scripts/fixtures/builder.ts';
import {
  isPptxEditorError,
  PptxEditor,
  type ParagraphFormat,
  type RequestContext,
} from '../browser.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { IN, open, refusedAtSync, reopen } from './support.ts';

const FONT: FontSource = {
  family: 'STIX Two Math',
  bold: false,
  italic: false,
  bytes: new Uint8Array(
    readFileSync(join(import.meta.dir, '../../../../fixtures/fonts/STIXTwoMath-Regular.otf'))
  ),
  origin: 'system',
};

const EXT =
  '<a:extLst><a:ext uri="{3C5E3F21-0D1A-4C8B-9C3E-2C3F5A1B7E10}"><x:keep xmlns:x="urn:example:keep"/></a:ext></a:extLst>';

/**
 * Slide 1: a box with three paragraphs (1.5 lines and 6 points before with a hanging indent;
 * exactly 24 points and half a line after; nothing of its own), and a 2 x 1 table. Slide 2:
 * the body placeholder of the title layout, which inherits from the master body style.
 */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({
            id: 2,
            name: 'Box',
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: 2 * IN,
            text:
              para(
                run('One'),
                `<a:pPr marL="457200" indent="-228600" algn="r"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef><a:buNone/><a:defRPr b="1"/>${EXT}</a:pPr>`
              ) +
              para(
                run('Two'),
                '<a:pPr><a:lnSpc><a:spcPts val="2400"/></a:lnSpc><a:spcAft><a:spcPct val="50000"/></a:spcAft></a:pPr>'
              ) +
              para(run('Three')),
          }) +
          table({
            id: 3,
            x: IN,
            y: 4 * IN,
            cx: 4 * IN,
            cy: IN,
            columns: [2 * IN, 2 * IN],
            tblPr: '<a:tblPr/>',
            rows: `<a:tr h="370840">${cell('A')}${cell('B\nC')}</a:tr>`,
          }),
      },
      {
        layout: 'title',
        shapes: shape({
          id: 2,
          name: 'Body',
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: 3 * IN,
          ph: '<p:ph idx="1"/>',
          text: para(run('Inherited')),
        }),
      },
    ],
  });

const NAMES =
  'lineSpacing,lineSpacingRule,spaceBefore,spaceAfter,lineUnitBefore,lineUnitAfter,leftIndent,firstLineIndent';

type Read = Pick<
  ParagraphFormat,
  | 'lineSpacing'
  | 'lineSpacingRule'
  | 'spaceBefore'
  | 'spaceAfter'
  | 'lineUnitBefore'
  | 'lineUnitAfter'
  | 'leftIndent'
  | 'firstLineIndent'
>;

const values = (f: ParagraphFormat): Read => ({
  lineSpacing: f.lineSpacing,
  lineSpacingRule: f.lineSpacingRule,
  spaceBefore: f.spaceBefore,
  spaceAfter: f.spaceAfter,
  lineUnitBefore: f.lineUnitBefore,
  lineUnitAfter: f.lineUnitAfter,
  leftIndent: f.leftIndent,
  firstLineIndent: f.firstLineIndent,
});

const box = (context: RequestContext) =>
  context.presentation.slides.getItemAt(0).shapes.getItem('2');

/** The range of paragraph `i` (0-based) of "One\rTwo\rThree". */
const paragraph = (context: RequestContext, i: number) =>
  box(context).textFrame.textRange.getSubstring([0, 4, 8][i]!, [3, 3, 5][i]!);

/** The XML of slide `n` in saved bytes. */
function slideXml(bytes: Uint8Array, n = 1): string {
  const zip = readZip(bytes);
  if (!zip.ok) throw new Error('saved');
  return new TextDecoder().decode(new Map(zip.entries).get(`/ppt/slides/slide${n}.xml`));
}

/** The `a:pPr` elements of a slide, in document order. */
const paragraphProps = (xml: string): string[] =>
  xml.match(/<a:pPr\b[^>]*?(?:\/>|>.*?<\/a:pPr>)/g) ?? [];

async function readAll(runtime: Awaited<ReturnType<typeof open>>): Promise<Read[]> {
  return runtime.run(async (context) => {
    const formats = [0, 1, 2].map((i) => paragraph(context, i).paragraphFormat.load(NAMES));
    await context.sync();
    return formats.map(values);
  });
}

describe('ParagraphFormat spacing and indents', () => {
  test('reads each paragraph through the cascade, in points and in lines', async () => {
    expect(await readAll(await open(deck()))).toEqual([
      {
        lineSpacing: 18,
        lineSpacingRule: 'Multiple',
        spaceBefore: 6,
        spaceAfter: 0,
        lineUnitBefore: null,
        lineUnitAfter: 0,
        leftIndent: 36,
        firstLineIndent: -18,
      },
      {
        lineSpacing: 24,
        lineSpacingRule: 'Exactly',
        spaceBefore: 0,
        spaceAfter: null,
        lineUnitBefore: 0,
        lineUnitAfter: 0.5,
        leftIndent: 0,
        firstLineIndent: 0,
      },
      // The master's other style gives a text box 90 % line spacing.
      {
        lineSpacing: 10.8,
        lineSpacingRule: 'Multiple',
        spaceBefore: 0,
        spaceAfter: 0,
        lineUnitBefore: 0,
        lineUnitAfter: 0,
        leftIndent: 0,
        firstLineIndent: 0,
      },
    ]);
  });

  test('a range over paragraphs that differ reads null, and a shared value reads it', async () => {
    const runtime = await open(deck());
    const read = await runtime.run(async (context) => {
      const all = box(context).textFrame.textRange.paragraphFormat.load(NAMES);
      const lastTwo = box(context)
        .textFrame.textRange.getSubstring(4, 9)
        .paragraphFormat.load(NAMES);
      await context.sync();
      return [values(all), values(lastTwo)];
    });
    expect(read[0]).toEqual({
      lineSpacing: null,
      lineSpacingRule: null,
      spaceBefore: null,
      spaceAfter: null,
      lineUnitBefore: null,
      lineUnitAfter: null,
      leftIndent: null,
      firstLineIndent: null,
    });
    // Paragraphs 2 and 3 share the zero space before and the indents only.
    expect(read[1]).toEqual({
      lineSpacing: null,
      lineSpacingRule: null,
      spaceBefore: 0,
      spaceAfter: null,
      lineUnitBefore: 0,
      lineUnitAfter: null,
      leftIndent: 0,
      firstLineIndent: 0,
    });
  });

  test('a placeholder reads the values its layout and master give', async () => {
    const runtime = await open(deck());
    const read = await runtime.run(async (context) => {
      const format = context.presentation.slides
        .getItemAt(1)
        .shapes.getItem('2')
        .textFrame.textRange.paragraphFormat.load(NAMES);
      await context.sync();
      return values(format);
    });
    // The master body style, level 1: 90 % line spacing, 10 points before, a hanging bullet.
    expect(read).toEqual({
      lineSpacing: 10.8,
      lineSpacingRule: 'Multiple',
      spaceBefore: 10,
      spaceAfter: 0,
      lineUnitBefore: null,
      lineUnitAfter: 0,
      leftIndent: 18,
      firstLineIndent: -18,
    });
  });

  test('writes each field on the range only, in schema order, and keeps other content', async () => {
    const runtime = await open(deck());
    await runtime.run(async (context) => {
      const f = paragraph(context, 0).paragraphFormat;
      f.lineSpacing = 24;
      f.spaceBefore = 12;
      f.spaceAfter = 3;
      f.leftIndent = 72;
      f.firstLineIndent = 9;
      await context.sync();
    });
    const props = paragraphProps(slideXml(await runtime.save()));
    expect(props[0]).toBe(
      `<a:pPr marL="914400" indent="114300" algn="r"><a:lnSpc><a:spcPct val="200000"/></a:lnSpc><a:spcBef><a:spcPts val="1200"/></a:spcBef><a:spcAft><a:spcPts val="300"/></a:spcAft><a:buNone/><a:defRPr b="1"/>${EXT}</a:pPr>`
    );
    // The other paragraphs keep their properties.
    expect(props[1]).toBe(
      '<a:pPr><a:lnSpc><a:spcPts val="2400"/></a:lnSpc><a:spcAft><a:spcPct val="50000"/></a:spcAft></a:pPr>'
    );
    const reads = await readAll(await reopen(runtime));
    expect(reads[0]).toEqual({
      lineSpacing: 24,
      lineSpacingRule: 'Multiple',
      spaceBefore: 12,
      spaceAfter: 3,
      lineUnitBefore: null,
      lineUnitAfter: null,
      leftIndent: 72,
      firstLineIndent: 9,
    });
  });

  test('a line spacing write keeps each paragraph rule', async () => {
    const runtime = await open(deck());
    const read = await runtime.run(async (context) => {
      const f = box(context).textFrame.textRange.paragraphFormat;
      f.lineSpacing = 15;
      const each = [0, 1, 2].map((i) =>
        paragraph(context, i).paragraphFormat.load('lineSpacing,lineSpacingRule')
      );
      await context.sync();
      return each.map((p) => [p.lineSpacing, p.lineSpacingRule]);
    });
    expect(read).toEqual([
      [15, 'Multiple'],
      [15, 'Exactly'],
      [15, 'Multiple'],
    ]);
    const xml = slideXml(await runtime.save());
    // 15 points as a multiple is 1.25 lines; as an exact spacing it is 1500.
    expect(xml).toContain('<a:lnSpc><a:spcPct val="125000"/></a:lnSpc>');
    expect(xml).toContain('<a:lnSpc><a:spcPts val="1500"/></a:lnSpc>');
  });

  test('a rule write keeps the points in force, and commutes with a spacing write', async () => {
    for (const order of ['rule first', 'spacing first']) {
      const runtime = await open(deck());
      const read = await runtime.run(async (context) => {
        const f = box(context).textFrame.textRange.paragraphFormat;
        if (order === 'rule first') f.lineSpacingRule = 'Exactly';
        f.lineSpacing = 20;
        if (order === 'spacing first') f.lineSpacingRule = 'Exactly';
        f.load('lineSpacing,lineSpacingRule');
        await context.sync();
        return [f.lineSpacing, f.lineSpacingRule];
      });
      expect(read).toEqual([20, 'Exactly']);
    }
    // 1.5 lines become exactly 18 points; exactly 24 points become 2 lines.
    const runtime = await open(deck());
    const read = await runtime.run(async (context) => {
      paragraph(context, 0).paragraphFormat.lineSpacingRule = 'Exactly';
      paragraph(context, 1).paragraphFormat.lineSpacingRule = 'Multiple';
      const each = [0, 1].map((i) =>
        paragraph(context, i).paragraphFormat.load('lineSpacing,lineSpacingRule')
      );
      await context.sync();
      return each.map((p) => [p.lineSpacing, p.lineSpacingRule]);
    });
    expect(read).toEqual([
      [18, 'Exactly'],
      [24, 'Multiple'],
    ]);
  });

  test('a rule write that a paragraph already has adds no formatting', async () => {
    const runtime = await open(deck());
    await runtime.run(async (context) => {
      paragraph(context, 2).paragraphFormat.lineSpacingRule = 'Multiple';
      await context.sync();
    });
    expect(paragraphProps(slideXml(await runtime.save()))[2]).toBe('<a:pPr/>');
  });

  test('a space in points replaces a space in lines, and back', async () => {
    const runtime = await open(deck());
    const read = await runtime.run(async (context) => {
      const two = paragraph(context, 1).paragraphFormat;
      two.spaceAfter = 4;
      two.lineUnitBefore = 0.25;
      two.load('spaceBefore,spaceAfter,lineUnitBefore,lineUnitAfter');
      await context.sync();
      return [two.spaceBefore, two.spaceAfter, two.lineUnitBefore, two.lineUnitAfter];
    });
    expect(read).toEqual([null, 4, 0.25, null]);
    expect(paragraphProps(slideXml(await runtime.save()))[1]).toBe(
      '<a:pPr><a:lnSpc><a:spcPts val="2400"/></a:lnSpc><a:spcBef><a:spcPct val="25000"/></a:spcBef><a:spcAft><a:spcPts val="400"/></a:spcAft></a:pPr>'
    );
  });

  test('a master shape reads its own values only, and a write keeps single spacing', async () => {
    const runtime = await open(deck());
    const read = await runtime.run(async (context) => {
      // The body placeholder of the master: its paragraph sets no spacing of its own.
      const f = context.presentation.slideMasters.getItemAt(0).shapes.getItem('3').textFrame
        .textRange.paragraphFormat;
      const before = f.load(NAMES);
      await context.sync();
      const first = values(before);
      f.lineSpacing = 18;
      f.leftIndent = 9;
      f.load('lineSpacing,lineSpacingRule,leftIndent');
      await context.sync();
      return [first, [f.lineSpacing, f.lineSpacingRule, f.leftIndent]];
    });
    expect(read[0]).toMatchObject({
      lineSpacing: null,
      lineSpacingRule: null,
      spaceBefore: null,
      leftIndent: null,
    });
    expect(read[1]).toEqual([18, 'Multiple', 9]);
    // The master title sets no spacing of its own: a rule write starts from single spacing.
    const title = await runtime.run(async (context) => {
      const f = context.presentation.slideMasters.getItemAt(0).shapes.getItem('2').textFrame
        .textRange.paragraphFormat;
      f.lineSpacingRule = 'Exactly';
      f.load('lineSpacing,lineSpacingRule');
      await context.sync();
      return [f.lineSpacing, f.lineSpacingRule];
    });
    expect(title).toEqual([12, 'Exactly']);
  });

  test('a placeholder write overrides the master value on the slide only', async () => {
    const runtime = await open(deck());
    const read = await runtime.run(async (context) => {
      const f = context.presentation.slides.getItemAt(1).shapes.getItem('2').textFrame
        .textRange.paragraphFormat;
      f.spaceBefore = 0;
      f.lineSpacingRule = 'Exactly';
      f.load('spaceBefore,lineSpacing,lineSpacingRule');
      await context.sync();
      return [f.spaceBefore, f.lineSpacing, f.lineSpacingRule];
    });
    // 90 % of 12 points is exactly 10.8 points.
    expect(read).toEqual([0, 10.8, 'Exactly']);
    const bytes = await runtime.save();
    expect(slideXml(bytes, 2)).toContain(
      '<a:pPr><a:lnSpc><a:spcPts val="1080"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef></a:pPr>'
    );
    const zip = readZip(bytes);
    if (!zip.ok) throw new Error('saved');
    const master = new TextDecoder().decode(
      new Map(zip.entries).get('/ppt/slideMasters/slideMaster1.xml')
    );
    expect(master).toContain('<a:spcBef><a:spcPts val="1000"/></a:spcBef>');
  });

  test('works on a table cell through TableCell.paragraphFormat', async () => {
    const runtime = await open(deck());
    await runtime.run(async (context) => {
      const t = context.presentation.slides.getItemAt(0).shapes.getItem('3').getTable();
      const f = t.getCellOrNullObject(0, 1).paragraphFormat;
      f.lineSpacing = 6;
      f.spaceAfter = 2;
      f.leftIndent = 4;
      await context.sync();
    });
    const again = await reopen(runtime);
    const read = await again.run(async (context) => {
      const t = context.presentation.slides.getItemAt(0).shapes.getItem('3').getTable();
      const changed = t.getCellOrNullObject(0, 1).paragraphFormat.load(NAMES);
      const other = t.getCellOrNullObject(0, 0).paragraphFormat.load('lineSpacing,leftIndent');
      const cellLoad = t.getCellOrNullObject(0, 1).load('paragraphFormat/spaceAfter');
      await context.sync();
      return [
        values(changed),
        [other.lineSpacing, other.leftIndent],
        cellLoad.paragraphFormat.spaceAfter,
      ];
    });
    expect(read).toEqual([
      {
        lineSpacing: 6,
        lineSpacingRule: 'Multiple',
        spaceBefore: 0,
        spaceAfter: 2,
        lineUnitBefore: 0,
        lineUnitAfter: null,
        leftIndent: 4,
        firstLineIndent: 0,
      },
      [10.8, 0],
      2,
    ]);
  });

  test('more space reports overflow through TextFrame.getTextFit', async () => {
    // A box that does not grow, measured with an open face: a runtime with no face refuses.
    const fixed = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 3 * IN,
            cy: IN,
            text: para(run('One')) + para(run('Two')),
            bodyPr: '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>',
          }),
        },
      ],
    });
    const runtime = await PptxEditor.createServer(fixed, { substitutes: false, fonts: FONT });
    const read = await runtime.run(async (context) => {
      const frame = box(context).textFrame;
      const fits = [frame.getTextFit()];
      await context.sync();
      // The space before the first paragraph does not apply, so the second one moves down.
      frame.textRange.paragraphFormat.spaceBefore = 200;
      fits.push(frame.getTextFit());
      await context.sync();
      frame.textRange.paragraphFormat.spaceBefore = 0;
      frame.textRange.paragraphFormat.lineSpacingRule = 'Exactly';
      frame.textRange.paragraphFormat.lineSpacing = 60;
      fits.push(frame.getTextFit());
      await context.sync();
      return fits.map((fit) => fit.value.overflows);
    });
    expect(read).toEqual([false, true, true]);
  });

  test('a percent string line spacing lays out as its thousandths form', async () => {
    const boxOf = (id: number, val: string) =>
      shape({
        id,
        x: IN,
        y: id * IN,
        cx: 3 * IN,
        cy: IN,
        text: ['One', 'Two', 'Three']
          .map((t) => para(run(t), `<a:pPr><a:lnSpc><a:spcPct val="${val}"/></a:lnSpc></a:pPr>`))
          .join(''),
        bodyPr: '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>',
      });
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: boxOf(2, '150%') + boxOf(3, '150000') + boxOf(4, '100000'),
        },
      ],
    });
    const runtime = await PptxEditor.createServer(bytes, { substitutes: false, fonts: FONT });
    const heights = await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      const fits = ['2', '3', '4'].map((id) => shapes.getItem(id).textFrame.getTextFit());
      await context.sync();
      return fits.map((fit) => fit.value.textHeight);
    });
    expect(heights[0]).toBe(heights[1]);
    expect(heights[0]).toBeGreaterThan(heights[2]!);
  });

  test('refuses an invalid value with InvalidArgument and writes nothing', async () => {
    const runtime = await open(deck());
    const original = slideXml(await (await open(deck())).save());
    const cases: [keyof Read, unknown][] = [
      ['lineSpacing', 0],
      ['lineSpacing', 1585],
      ['lineSpacing', Number.NaN],
      ['lineSpacing', '12'],
      ['lineSpacingRule', 'AtLeast'],
      ['spaceBefore', -1],
      ['spaceAfter', 1585],
      ['lineUnitBefore', 133],
      ['lineUnitAfter', -0.5],
      ['leftIndent', -1],
      ['leftIndent', 4033],
      ['firstLineIndent', -4033],
    ];
    for (const [name, value] of cases) {
      const error = await runtime.run((context) =>
        refusedAtSync(context, () => {
          const f = box(context).textFrame.textRange.paragraphFormat;
          f.spaceAfter = 9;
          (f as unknown as Record<string, unknown>)[name] = value;
        })
      );
      expect([name, isPptxEditorError(error) && error.code]).toEqual([name, 'InvalidArgument']);
    }
    expect(slideXml(await runtime.save())).toBe(original);
  });

  test('the server and the browser runtime write the same paragraph properties', async () => {
    const edit = async (context: RequestContext) => {
      const f = box(context).textFrame.textRange.paragraphFormat;
      f.lineSpacingRule = 'Exactly';
      f.lineSpacing = 14;
      f.lineUnitBefore = 0.5;
      f.spaceAfter = 6;
      f.leftIndent = 10;
      f.firstLineIndent = -10;
      await context.sync();
    };
    const server = await open(deck());
    await server.run(edit);
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(deck())).ok).toBe(true);
    await PptxEditor.createBrowser(editor).run(edit);
    const browserXml = slideXml(editor.save()!);
    expect(paragraphProps(browserXml)).toEqual(paragraphProps(slideXml(await server.save())));
    expect(browserXml).toContain('<a:lnSpc><a:spcPts val="1400"/></a:lnSpc>');
  });

  test('one sync is one undo step in the editor', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(deck())).ok).toBe(true);
    const before = slideXml(editor.save()!);
    await PptxEditor.createBrowser(editor).run(async (context) => {
      const f = box(context).textFrame.textRange.paragraphFormat;
      f.lineSpacing = 30;
      f.spaceBefore = 8;
      f.leftIndent = 20;
      await context.sync();
    });
    expect(slideXml(editor.save()!)).not.toBe(before);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(slideXml(editor.save()!)).toBe(before);
  });

  // The collaboration test in `packages/pro` sends this exact sequence to two peers.
  test('emits one setParagraph operation for each write', async () => {
    const read = readPptx(deck());
    if (!read.ok) throw new Error('fixture');
    const inner = createStoreAutomationHost(new PptxDocumentStore(read.pkg));
    const sent: unknown[] = [];
    const host: AutomationHost = {
      ...inner,
      execute: (request) => {
        for (const operation of request.operations)
          if (operation.op === 'setParagraph') sent.push(operation);
        return inner.execute(request);
      },
    };
    await createRuntime({ host, save: true }).run(async (context) => {
      const f = box(context).textFrame.textRange.getSubstring(0, 3).paragraphFormat;
      f.lineSpacing = 18;
      f.lineSpacingRule = 'Exactly';
      f.spaceBefore = 6;
      f.lineUnitAfter = 0.5;
      f.leftIndent = 12;
      f.firstLineIndent = -6;
      await context.sync();
    });
    const op = (paragraph: object) => ({
      op: 'setParagraph',
      shape: expect.anything(),
      span: { start: 0, length: 3 },
      paragraph,
    });
    expect(sent).toEqual([
      op({ lineSpacing: 18 }),
      op({ lineSpacingRule: 'exactly' }),
      op({ spaceBefore: 6 }),
      op({ lineUnitAfter: 0.5 }),
      op({ leftIndent: 12 }),
      op({ firstLineIndent: -6 }),
    ]);
  });
});
