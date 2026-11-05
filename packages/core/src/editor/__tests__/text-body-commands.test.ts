// Home > Text Direction and Home > Columns: `a:bodyPr/@vert`, `@numCol`, and `@spcCol` on the
// selected shapes. Each value lays out, undoes as one step, and survives save and reopen; the
// two vertical types that layout does not draw correctly are refused with a typed reason.

import { describe, expect, test } from 'bun:test';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  table,
} from '../../../../../scripts/fixtures/builder.ts';
import { attr, elements, find, readPptx, xmlPartOf, type Element } from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide } from '../../layout/index.ts';
import { MAX_COORDINATE } from '../../automation/commands/package-edits.ts';
import { columnFlowHeight } from '../../automation/commands/autofit.ts';
import { COLUMNS_CONTROL, TEXT_DIRECTION_CONTROL } from '../chrome-paragraph-controls.ts';
import {
  createPptxEditor,
  runToolbarCommand,
  toolbarCommandState,
  type PptxEditorInstance,
} from '../index.ts';

const IN = 914_400;
const LOREM = 'The quick brown fox jumps over the lazy dog. '.repeat(6);

function deck(): Uint8Array {
  return buildPptx({
    slides: [
      {
        shapes:
          shape({
            id: 2,
            name: 'T',
            x: IN,
            y: IN,
            cx: 6 * IN,
            cy: 1.5 * IN,
            text: para(run(LOREM)),
          }) +
          '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="L"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>',
      },
    ],
  });
}

async function open(bytes = deck()): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const root = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
const byName = (e: PptxEditorInstance, name: string): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === name)!.id;
const bodyPr = (e: PptxEditorInstance): Element => find(root(e), 'a:bodyPr')!;
const textLayout = (e: PptxEditorInstance) => {
  const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.text)!;
  return item.kind === 'shape' ? item.text! : undefined;
};
const selectShape = (e: PptxEditorInstance, name = 'T') =>
  e.setSelection({ kind: 'shapes', nodeIds: [byName(e, name)] });

describe('setTextDirection', () => {
  test('writes vert and lays the text out rotated, as one undo step', async () => {
    const e = await open();
    selectShape(e);
    expect(textLayout(e)!.rotation).toBe(0);
    expect(runToolbarCommand(e, 'paragraph.textDirection', 'vertical270').ok).toBe(true);
    expect(attr(bodyPr(e), 'vert')).toBe('vert270');
    expect(textLayout(e)!.rotation).toBe(270);
    expect(e.snapshot().format?.textDirection).toBe('vertical270');
    expect(toolbarCommandState(e, 'paragraph.textDirection').value).toBe('vertical270');
    expect(e.exec({ type: 'setTextDirection', direction: 'horizontal' }).ok).toBe(true);
    expect(textLayout(e)!.rotation).toBe(0);
    e.exec({ type: 'undo' });
    expect(attr(bodyPr(e), 'vert')).toBe('vert270');
  });

  test('refuses the types layout does not draw, and file tokens, with the field as the detail', async () => {
    const e = await open();
    selectShape(e);
    for (const direction of ['mongolianVertical', 'wordArtVerticalRtl', 'vert', 'sideways'])
      expect(e.exec({ type: 'setTextDirection', direction: direction as never })).toMatchObject({
        ok: false,
        reason: 'errors.unsupported',
        detail: 'direction',
      });
    expect(attr(bodyPr(e), 'vert')).toBeUndefined();
  });

  test('needs a shape with text', async () => {
    const e = await open();
    selectShape(e, 'L');
    expect(e.can({ type: 'setTextDirection', direction: 'vertical' })).toEqual({
      ok: false,
      reason: 'disabled.noText',
    });
  });

  test('in a table cell, Text Direction and Columns say they are not available in a table', async () => {
    const e = await open(
      buildPptx({
        slides: [
          {
            shapes: table({
              id: 2,
              x: IN,
              y: IN,
              cx: 4 * IN,
              cy: IN,
              columns: [2 * IN, 2 * IN],
              tblPr: '<a:tblPr/>',
              rows: `<a:tr h="370840">${cell('A')}${cell('B')}</a:tr>`,
            }),
          },
        ],
      })
    );
    const frame = elements(find(root(e), 'p:spTree')!).find((c) => c.local === 'graphicFrame')!;
    const at = { paragraph: 0, offset: 0 };
    e.setSelection({
      kind: 'text',
      nodeId: frame.id,
      cell: { row: 0, column: 0 },
      anchor: at,
      focus: at,
    });
    const refusal = { ok: false, reason: 'disabled.tableText' };
    expect(e.can({ type: 'setTextDirection', direction: 'vertical' })).toEqual(refusal);
    expect(e.can({ type: 'setColumns', count: 2 })).toEqual(refusal);
    expect(toolbarCommandState(e, 'paragraph.columns').disabledReason).toBe('disabled.tableText');
  });
});

describe('setColumns', () => {
  test('writes numCol and spcCol, and lines flow into the next column', async () => {
    const e = await open();
    selectShape(e);
    expect(textLayout(e)!.lines.every((l) => l.column === 0)).toBe(true);
    expect(e.exec({ type: 'setColumns', count: 2, gapEmu: IN / 4 }).ok).toBe(true);
    expect(attr(bodyPr(e), 'numCol')).toBe('2');
    expect(attr(bodyPr(e), 'spcCol')).toBe(String(IN / 4));
    expect(textLayout(e)!.lines.some((l) => l.column === 1)).toBe(true);
    expect(e.snapshot().format?.columns).toEqual({ count: 2, gapEmu: IN / 4 });
    // The menu keeps the gap.
    expect(runToolbarCommand(e, 'paragraph.columns', 3).ok).toBe(true);
    expect(attr(bodyPr(e), 'spcCol')).toBe(String(IN / 4));
    expect(toolbarCommandState(e, 'paragraph.columns').value).toBe('3');
    // Only a number or a numeric string runs the slot.
    for (const value of [true, [2], null, {}])
      expect(runToolbarCommand(e, 'paragraph.columns', value as never).ok).toBe(false);
  });

  test('refuses counts and gaps the schema does not allow', async () => {
    const e = await open();
    selectShape(e);
    for (const [count, gapEmu, field] of [
      [0, undefined, 'count'],
      [17, undefined, 'count'],
      [1.5, undefined, 'count'],
      [2, -1, 'gapEmu'],
      [2, 60_000_000, 'gapEmu'],
      [2, Number.NaN, 'gapEmu'],
    ] as const)
      expect(
        e.exec({ type: 'setColumns', count, ...(gapEmu !== undefined ? { gapEmu } : {}) })
      ).toMatchObject({ ok: false, detail: field });
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('direction and columns survive save and reopen', async () => {
    const e = await open();
    selectShape(e);
    e.exec({ type: 'setColumns', count: 3, gapEmu: 100_000 });
    e.exec({ type: 'setTextDirection', direction: 'wordArtVertical' });
    const saved = e.save()!;
    expect(readPptx(saved).ok).toBe(true);
    const again = await open(saved);
    selectShape(again);
    const f = again.snapshot().format!;
    expect([f.textDirection, f.columns]).toEqual([
      'wordArtVertical',
      { count: 3, gapEmu: 100_000 },
    ]);
  });
});

describe('the read', () => {
  test('a file value outside the schema reads as layout draws it; shapes that differ read none', async () => {
    const odd = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            shape({
              id: 2,
              name: 'A',
              x: 0,
              y: 0,
              cx: 2 * IN,
              cy: 2 * IN,
              bodyPr: '<a:bodyPr vert="bogus" numCol="2.5"/>',
              text: para(run('a')),
            }) +
            shape({
              id: 3,
              name: 'B',
              x: 3 * IN,
              y: 0,
              cx: 2 * IN,
              cy: 2 * IN,
              bodyPr: '<a:bodyPr vert="vert" numCol="3"/>',
              text: para(run('b')),
            }),
        },
      ],
    });
    const e = await open(odd);
    selectShape(e, 'A');
    expect(e.snapshot().format?.textDirection).toBe('horizontal');
    expect(e.snapshot().format?.columns).toEqual({ count: 2, gapEmu: 0 });
    e.setSelection({ kind: 'shapes', nodeIds: [byName(e, 'B'), byName(e, 'A')] });
    expect(e.snapshot().format?.textDirection).toBeUndefined();
    expect(e.snapshot().format?.columns).toBeUndefined();
    expect(toolbarCommandState(e, 'paragraph.textDirection').value).toBeUndefined();
  });
});

describe('autofit after a direction or column change', () => {
  const one = (bodyPr: string, cx: number, cy: number, text = para(run(LOREM))) =>
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({ id: 2, name: 'T', x: IN, y: IN, cx, cy, bodyPr, text }),
        },
      ],
    });
  const ext = (e: PptxEditorInstance) => {
    const x = find(
      elements(find(root(e), 'p:spTree')!).find((c) => c.id === byName(e, 'T'))!,
      'a:ext'
    )!;
    return [Number(attr(x, 'cx')), Number(attr(x, 'cy'))];
  };
  const norm = (e: PptxEditorInstance) => find(bodyPr(e), 'a:normAutofit');

  test('shrink on overflow measures rotated text against the shape width', async () => {
    const e = await open(
      one(
        '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>',
        8 * IN,
        IN,
        para(run('A short line'))
      )
    );
    selectShape(e);
    expect(e.exec({ type: 'setTextDirection', direction: 'vertical' }).ok).toBe(true);
    // The rotated line is 1 in long and its height is far less than 8 in: no shrink.
    expect(attr(norm(e), 'fontScale')).toBeUndefined();
  });

  test('resize shape to fit text grows a rotated shape across its lines, the width', async () => {
    const e = await open(
      one('<a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr>', 4 * IN, 0.4 * IN)
    );
    selectShape(e);
    expect(e.exec({ type: 'setTextDirection', direction: 'vertical' }).ok).toBe(true);
    const [cx, cy] = ext(e);
    expect(cy).toBe(Math.round(0.4 * IN));
    const laid = textLayout(e)!;
    expect(laid.contentHeight).toBeLessThanOrEqual(laid.box.height + 1);
    expect(cx).not.toBe(4 * IN);
  });

  test('both kinds see the last column overflow', async () => {
    const lots = [1, 2, 3, 4, 5, 6].map(() => para(run(LOREM))).join('');
    const shrink = await open(
      one('<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>', 4 * IN, 2 * IN, lots)
    );
    selectShape(shrink);
    expect(shrink.exec({ type: 'setColumns', count: 2 }).ok).toBe(true);
    expect(Number(attr(norm(shrink), 'fontScale'))).toBeLessThan(100_000);
    const grow = await open(
      one('<a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr>', 4 * IN, IN, lots)
    );
    selectShape(grow);
    expect(grow.exec({ type: 'setColumns', count: 2 }).ok).toBe(true);
    const laid = textLayout(grow)!;
    // Every line sits inside the box, and both columns hold text.
    expect(laid.contentHeight).toBeLessThanOrEqual(laid.box.height + 1);
    expect(laid.lines.some((l) => l.column === 1)).toBe(true);
  });

  test('typing into a two-column box that resizes never leaves the last line below it', async () => {
    const e = await open(
      one('<a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr>', 4 * IN, 0.4 * IN, para(run('Go')))
    );
    selectShape(e);
    expect(e.exec({ type: 'setColumns', count: 2 }).ok).toBe(true);
    const nodeId = byName(e, 'T');
    e.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 2 },
      focus: { paragraph: 0, offset: 2 },
    });
    for (const word of ' the quick brown fox jumps over the lazy dog again and again'.split(
      /(?= )/
    )) {
      e.input('text', word);
      const laid = textLayout(e)!;
      const inner = laid.box.height;
      for (const l of laid.lines) expect(l.top + l.height).toBeLessThanOrEqual(inner + 1);
    }
  });

  test('an ellipse that resizes grows until its text rectangle holds every line', async () => {
    for (const columns of [1, 2]) {
      const e = await open(
        buildPptx({
          slides: [
            {
              layout: 'blank',
              shapes: shape({
                id: 2,
                name: 'T',
                x: IN,
                y: IN,
                cx: 4 * IN,
                cy: IN,
                prst: 'ellipse',
                bodyPr: '<a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr>',
                text: para(run('Go')),
              }),
            },
          ],
        })
      );
      selectShape(e);
      if (columns > 1) expect(e.exec({ type: 'setColumns', count: columns }).ok).toBe(true);
      e.setSelection({
        kind: 'text',
        nodeId: byName(e, 'T'),
        anchor: { paragraph: 0, offset: 2 },
        focus: { paragraph: 0, offset: 2 },
      });
      for (const word of LOREM.split(/(?= )/).slice(0, 30)) {
        e.input('text', word);
        const laid = textLayout(e)!;
        for (const l of laid.lines)
          expect(l.top + l.height).toBeLessThanOrEqual(laid.box.height + 1);
      }
      // The ellipse grew, and not more than the text needs.
      const laid = textLayout(e)!;
      expect(ext(e)[1]).toBeGreaterThan(IN);
      expect(laid.box.height - laid.contentHeight).toBeLessThan(laid.lines[0]!.height);
    }
  });

  test('a paste into a rounded rectangle or an octagon that resizes leaves no line below', async () => {
    for (const [prst, columns] of [
      ['roundRect', 1],
      ['octagon', 1],
      ['octagon', 2],
      ['plaque', 1],
    ] as const) {
      for (const width of [2, 3.5, 4.5, 6]) {
        const e = await open(
          buildPptx({
            slides: [
              {
                layout: 'blank',
                shapes: shape({
                  id: 2,
                  name: 'T',
                  x: IN,
                  y: IN,
                  cx: width * IN,
                  cy: IN,
                  prst,
                  bodyPr: `<a:bodyPr wrap="square" numCol="${columns}"><a:spAutoFit/></a:bodyPr>`,
                  text: para(run('Go')),
                }),
              },
            ],
          })
        );
        e.setSelection({
          kind: 'text',
          nodeId: byName(e, 'T'),
          anchor: { paragraph: 0, offset: 2 },
          focus: { paragraph: 0, offset: 2 },
        });
        e.input('text', ` ${LOREM}`);
        const laid = textLayout(e)!;
        for (const l of laid.lines)
          expect(l.top + l.height, `${prst} ${columns} ${width}`).toBeLessThanOrEqual(
            laid.box.height + 1
          );
      }
    }
  });

  test('a placeholder that inherits a rotation keeps it when it resizes', async () => {
    const e = await open(
      buildPptx({
        slides: [
          {
            layout: 'title',
            shapes:
              '<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
              `<p:spPr/><p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${para(run('Go'))}</p:txBody></p:sp>`,
          },
        ],
      })
    );
    // The layout's title placeholder turns 30° and flips; the slide's title has no a:xfrm.
    const layoutPart = e.session!.info.slides[0]!.layoutPart!;
    const layoutRoot = xmlPartOf(e.session!.store.package, layoutPart)!.root;
    const spPr = find(
      elements(find(layoutRoot, 'p:spTree')!).find(
        (c) => find(c, 'p:ph') && attr(find(c, 'p:ph'), 'type') === 'title'
      )!,
      'p:spPr'
    )!;
    const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    expect(
      e.session!.store.transact([
        {
          op: 'insertNode',
          part: layoutPart,
          parent: spPr.id,
          index: 0,
          spec: {
            ns: A,
            local: 'xfrm',
            prefix: 'a',
            attrs: [
              { ns: '', local: 'rot', value: '1800000' },
              { ns: '', local: 'flipH', value: '1' },
            ],
            children: [
              {
                ns: A,
                local: 'off',
                prefix: 'a',
                attrs: [
                  { ns: '', local: 'x', value: '914400' },
                  { ns: '', local: 'y', value: '914400' },
                ],
              },
              {
                ns: A,
                local: 'ext',
                prefix: 'a',
                attrs: [
                  { ns: '', local: 'cx', value: '4572000' },
                  { ns: '', local: 'cy', value: '914400' },
                ],
              },
            ],
          },
        },
      ]).ok
    ).toBe(true);
    const drawn = () => {
      const item = e
        .session!.layout(0)!
        .items.find((i) => i.kind === 'shape' && i.element.name === 'T')!;
      return item.element.xfrm;
    };
    const before = drawn();
    expect([before.rot, before.flipH]).toEqual([30, true]);
    e.setSelection({
      kind: 'text',
      nodeId: byName(e, 'T'),
      anchor: { paragraph: 0, offset: 2 },
      focus: { paragraph: 0, offset: 2 },
    });
    e.input('text', ` ${LOREM}`);
    expect([drawn().rot, drawn().flipH, drawn().flipV]).toEqual([
      before.rot,
      before.flipH,
      before.flipV,
    ]);
  });

  test('a text rectangle that does not grow with the shape keeps the size', async () => {
    const e = await open(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes:
              '<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
              '<p:spPr><a:xfrm rot="5400000"><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm>' +
              '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="3657600" b="457200"/>' +
              '<a:pathLst><a:path w="3657600" h="914400"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="3657600" y="914400"/></a:lnTo></a:path></a:pathLst></a:custGeom></p:spPr>' +
              `<p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${para(run('Go'))}</p:txBody></p:sp>`,
          },
        ],
      })
    );
    e.setSelection({
      kind: 'text',
      nodeId: byName(e, 'T'),
      anchor: { paragraph: 0, offset: 2 },
      focus: { paragraph: 0, offset: 2 },
    });
    for (const ch of 'abc') e.input('text', ch);
    expect(ext(e)).toEqual([3_657_600, 914_400]);
  });

  test('a transform without a:off or a:ext still resizes, with the child added in place', async () => {
    for (const [xfrm, kids] of [
      ['<a:xfrm><a:ext cx="3657600" cy="91440"/></a:xfrm>', ['off', 'ext']],
      ['<a:xfrm><a:off x="914400" y="914400"/></a:xfrm>', ['off', 'ext']],
      ['<a:xfrm rot="5400000"/>', ['off', 'ext']],
    ] as const) {
      const bytes = buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes:
              '<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
              `<p:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
              `<p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${para(run('Go'))}</p:txBody></p:sp>`,
          },
        ],
      });
      const e = await open(bytes);
      selectShape(e);
      expect(e.exec({ type: 'setColumns', count: 2 }).ok).toBe(true);
      const shapeEl = elements(find(root(e), 'p:spTree')!).find((c) => c.id === byName(e, 'T'))!;
      expect(elements(find(shapeEl, 'a:xfrm')).map((c) => c.local)).toEqual([...kids]);
      expect(Number(attr(find(shapeEl, 'a:ext'), 'cy'))).toBeGreaterThan(0);
      // A refit never changes the rotation, also one the cascade could not read.
      expect(attr(find(shapeEl, 'a:xfrm'), 'rot')).toBe(
        xfrm.includes('rot') ? '5400000' : undefined
      );
    }
  });

  test('a rotated box that does not wrap keeps its size', async () => {
    const e = await open(
      one(
        '<a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr>',
        5 * IN,
        0.4 * IN,
        para(run('One long line of text'))
      )
    );
    selectShape(e);
    expect(e.exec({ type: 'setTextDirection', direction: 'vertical' }).ok).toBe(true);
    expect(ext(e)).toEqual([5 * IN, Math.round(0.4 * IN)]);
  });

  test('a crafted line spacing cannot make the column refit run without end', async () => {
    const huge = '<a:pPr><a:lnSpc><a:spcPct val="999999999999999"/></a:lnSpc></a:pPr>';
    const text = [1, 2, 3, 4].map(() => para(run('Big text', 'sz="20000"'), huge)).join('');
    const e = await open(one('<a:bodyPr numCol="2"><a:spAutoFit/></a:bodyPr>', 4 * IN, IN, text));
    const nodeId = byName(e, 'T');
    e.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 3 },
      focus: { paragraph: 0, offset: 3 },
    });
    const start = performance.now();
    expect(e.input('text', 'x').ok).toBe(true);
    expect(performance.now() - start).toBeLessThan(2_000);
    expect(ext(e)[1]).toBeLessThanOrEqual(MAX_COORDINATE);
  });

  test('a new a:normAutofit goes after a:prstTxWarp, in schema order', async () => {
    const e = await open(
      buildPptx({
        slides: [
          {
            layout: 'title',
            shapes: shape({
              id: 2,
              name: 'T',
              x: IN,
              y: IN,
              cx: 4 * IN,
              cy: IN,
              ph: '<p:ph type="title"/>',
              bodyPr:
                '<a:bodyPr><a:prstTxWarp prst="textPlain"><a:avLst/></a:prstTxWarp></a:bodyPr>',
              text: para(run('Title')),
            }),
          },
        ],
      })
    );
    selectShape(e);
    e.exec({ type: 'setTextDirection', direction: 'vertical' });
    const kids = elements(bodyPr(e)).map((c) => c.local);
    expect(kids).toEqual(['prstTxWarp', 'normAutofit']);
  });
});

describe('column flow height', () => {
  test('is the smallest height at which every line fits, by the layout rule', () => {
    const lines = [0, 10, 20, 30, 40].map((top) => ({ top, height: 10 }));
    expect(columnFlowHeight(lines, 1)).toBe(50);
    expect(columnFlowHeight(lines, 2)).toBe(30);
    expect(columnFlowHeight(lines, 5)).toBe(10);
    expect(columnFlowHeight([], 2)).toBe(0);
    // A line taller than the rest decides the height of its column.
    expect(
      columnFlowHeight(
        [
          { top: 0, height: 5 },
          { top: 5, height: 40 },
        ],
        2
      )
    ).toBe(40);
  });
});

describe('approximations', () => {
  test('mongolianVert and wordArtVertRtl report "other-vert"', () => {
    for (const vert of ['mongolianVert', 'wordArtVertRtl']) {
      const bytes = buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              x: 0,
              y: 0,
              cx: IN,
              cy: 3 * IN,
              bodyPr: `<a:bodyPr vert="${vert}"/>`,
              text: para(run('text')),
            }),
          },
        ],
      });
      const r = readPptx(bytes);
      if (!r.ok) throw new Error('open failed');
      const found = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer()).diagnostics;
      expect(found).toContain('approximation: other-vert');
    }
  });
});

describe('menu rows', () => {
  test('every Text Direction and Columns row has an icon for the row and the button', () => {
    for (const control of [TEXT_DIRECTION_CONTROL, COLUMNS_CONTROL]) {
      expect(control.options!.length).toBeGreaterThan(0);
      for (const option of control.options!.filter((o) => !o.dialog))
        expect(option.paths?.length).toBeGreaterThan(0);
    }
  });
});
