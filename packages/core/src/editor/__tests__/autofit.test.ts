// Shrink on overflow: other text searches the order inferred from saved decks (line-spacing
// reduction first, then font scale with the 20 % reduction); a title takes one 90 % step at
// most, as measured (see `autofit.ts`).

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { find, xmlPartOf } from '../../store/index.ts';
import { autofitStepsFor } from '../../layout/autofit.ts';
import {
  createPptxEditor,
  runToolbarCommand,
  toolbarCommandState,
  type PptxEditorInstance,
} from '../index.ts';

const IN = 914_400;
const PT = 12_700;

const SINGLE = '<a:pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr>';
/** A one-line paragraph of 18 pt: 21.6 pt at 100 %. */
const LINE = para(run('x', 'sz="1800"'), SINGLE);
/** An empty paragraph whose mark is 18 pt: also 21.6 pt at 100 %. */
const EMPTY = `<a:p>${SINGLE}<a:endParaRPr lang="en-US" sz="1800"/></a:p>`;

/**
 * Five one-line paragraphs (`line`, 18 pt at 100 % by default: 21.6 pt each) in a box
 * `heightPt` tall, set to shrink on overflow. `ph` makes the box a placeholder; `after` is XML
 * after the fifth line.
 */
async function shrinkEditor(
  heightPt: number,
  ph = '',
  after = '',
  line = LINE
): Promise<PptxEditorInstance> {
  const lines = line.repeat(5) + after;
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: heightPt * PT,
          text: lines,
          bodyPr: '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>',
          ph,
        }),
      },
    ],
  });
  const e: PptxEditorInstance = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  e.setSelection({ kind: 'shapes', nodeIds: [e.session!.layout(0)!.items[0].element.nodeId] });
  expect(e.exec({ type: 'setAutofit', mode: 'shrink' }).ok).toBe(true);
  return e;
}

function normAutofit(e: PptxEditorInstance): Record<string, string> {
  const part = xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!;
  const norm = find(part.root, 'a:normAutofit');
  expect(norm).toBeDefined();
  return Object.fromEntries(norm!.attrs.map((a) => [a.local, a.value]));
}

const shrink = async (
  heightPt: number,
  ph = '',
  after = '',
  line = LINE
): Promise<Record<string, string>> => normAutofit(await shrinkEditor(heightPt, ph, after, line));

describe('shrink on overflow', () => {
  test('a small overflow reduces line spacing and keeps the font size', async () => {
    // 5 × 21.6 + 7.2 insets = 115.2 pt; at 90 % it is 104.4 pt, which fits 110 pt.
    expect(await shrink(110)).toEqual({ lnSpcReduction: '10000' });
  });

  test('then 92.5 % with the 10 % reduction, then with 20 %: no 20 % step at full size', async () => {
    // 92.5 % gives 17 pt runs: 5 × 1.2 × 17 × 0.9 + 7.2 = 99 pt fits 100 pt; at 95 pt the 20 %
    // reduction follows (88.8 pt). 100 % with 20 % (93.6 pt) would fit both, but it is no step.
    expect(await shrink(100)).toEqual({ fontScale: '92500', lnSpcReduction: '10000' });
    expect(await shrink(95)).toEqual({ fontScale: '92500', lnSpcReduction: '20000' });
  });

  test('a large overflow scales the font in 7.5 % steps with the 20 % reduction', async () => {
    // 62.5 % gives 11 pt runs: 5 × 1.2 × 11 × 0.8 + 7.2 = 60 pt, which fits 61 pt; 70 % does not.
    expect(await shrink(61)).toEqual({ fontScale: '62500', lnSpcReduction: '20000' });
  });

  test('a title takes one step at most: 90 % with no line spacing reduction', async () => {
    // At 90 % (16 pt runs) the text is 5 × 1.2 × 16 + 7.2 = 103.2 pt: it fits 110 pt, where
    // other text only reduces its line spacing.
    expect(await shrink(110, '<p:ph type="title"/>')).toEqual({ fontScale: '90000' });
    // Text that overflows at 90 % stays at 90 % and overflows the box.
    expect(await shrink(61, '<p:ph type="title"/>')).toEqual({ fontScale: '90000' });
    expect(await shrink(61, '<p:ph type="ctrTitle"/>')).toEqual({ fontScale: '90000' });
  });

  test('a title that fits keeps 100 %', async () => {
    expect(await shrink(130, '<p:ph type="title"/>')).toEqual({});
  });

  test('empty paragraphs at the end do not count', async () => {
    // The five lines fit 130 pt at 100 % (115.2 pt). Two empty lines more would need 158.4 pt.
    expect(await shrink(130, '', EMPTY + EMPTY)).toEqual({});
    // A body of empty paragraphs only measures the first: 21.6 + 7.2 = 28.8 pt fits 30 pt.
    expect(await shrink(30, '', EMPTY.repeat(3), '')).toEqual({});
    // An empty paragraph before text counts: seven lines need 158.4 pt, 143.3 pt with the 10 %
    // reduction, and 135.7 pt at 92.5 % (17 pt runs) with it; with the 20 % reduction they
    // need 7 × 1.2 × 17 × 0.8 + 7.2 = 121.4 pt, which fits.
    expect(await shrink(130, '', EMPTY + LINE)).toEqual({
      fontScale: '92500',
      lnSpcReduction: '20000',
    });
  });

  test('a paragraph at the end with a line break or a field counts', async () => {
    // Each makes the text seven lines (158.4 pt at 100 %), as in the test above.
    const shrunk = { fontScale: '92500', lnSpcReduction: '20000' };
    const br = `<a:p>${SINGLE}<a:br><a:rPr lang="en-US" sz="1800"/></a:br><a:endParaRPr lang="en-US" sz="1800"/></a:p>`;
    expect(await shrink(130, '', br)).toEqual(shrunk);
    // A field counts even when its stored text is empty.
    const field = `<a:p>${SINGLE}<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="datetime1"><a:rPr lang="en-US" sz="1800"/><a:t></a:t></a:fld><a:endParaRPr lang="en-US" sz="1800"/></a:p>`;
    expect(await shrink(130, '', EMPTY + field)).toEqual(shrunk);
  });

  test('with spacing in points, 85 % takes the 10 % reduction before the 20 %', async () => {
    // 20 pt before lines 2 to 5 (80 pt) does not scale. 92.5 % with 20 % needs
    // 5 × 1.2 × 17 × 0.8 + 80 + 7.2 = 168.8 pt; 85 % (15 pt runs) with 10 % needs 168.2 pt.
    const spaced = para(
      run('x', 'sz="1800"'),
      '<a:pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="2000"/></a:spcBef></a:pPr>'
    );
    expect(await shrink(168.5, '', '', spaced)).toEqual({
      fontScale: '85000',
      lnSpcReduction: '10000',
    });
  });

  test('typing a new empty line at the end keeps the scale', async () => {
    const e = await shrinkEditor(130);
    const nodeId = e.session!.layout(0)!.items[0].element.nodeId;
    const end = { paragraph: 4, offset: 1 };
    e.setSelection({ kind: 'text', nodeId, anchor: end, focus: end });
    expect(e.input('paragraph').ok).toBe(true);
    expect(e.input('paragraph').ok).toBe(true);
    const el = e.session!.layout(0)!.items[0].element;
    expect(el.kind === 'shape' && el.text?.paragraphs.length).toBe(7);
    expect(normAutofit(e)).toEqual({});
  });
});

describe('the shrink steps', () => {
  test('85 % takes the 10 % reduction before the 20 %; smaller scales take 20 % only', () => {
    const steps = autofitStepsFor(undefined).map((s) => `${s.scale}/${s.reduction}`);
    expect(steps.slice(0, 7)).toEqual([
      '1/0',
      '1/0.1',
      '0.925/0.1',
      '0.925/0.2',
      '0.85/0.1',
      '0.85/0.2',
      '0.775/0.2',
    ]);
    expect(steps.filter((s) => s.endsWith('/0.1'))).toEqual(['1/0.1', '0.925/0.1', '0.85/0.1']);
  });
});

describe('the size that the selection reports', () => {
  test('shrunk text reports its drawn size, and the size box shows it', async () => {
    const e = await shrinkEditor(61);
    // 62.5 % of 18 pt is 11.25 pt, drawn at 11 pt (whole points, half up).
    expect(e.snapshot().format?.fontSize).toBe(18);
    expect(e.snapshot().format?.renderedFontSize).toBe(11);
    expect(toolbarCommandState(e, 'font.size').value).toBe(11);
  });

  test('a size set on shrunk text is written as given, and the text fits again', async () => {
    const e = await shrinkEditor(61);
    expect(e.exec({ type: 'setFontSize', points: 10 }).ok).toBe(true);
    // 5 × 1.2 × 10 + 7.2 = 67.2 pt overflows 61 pt, and 90 % line spacing gives 61.2 pt; at
    // 92.5 % the runs draw at 9 pt: 5 × 1.2 × 9 × 0.9 + 7.2 = 55.8 pt fits.
    expect(normAutofit(e)).toEqual({ fontScale: '92500', lnSpcReduction: '10000' });
    expect(e.snapshot().format?.fontSize).toBe(10);
    expect(toolbarCommandState(e, 'font.size').value).toBe(9);
  });
});

describe('Increase font size on shrunk text', () => {
  const sizes = (e: PptxEditorInstance) => [
    e.snapshot().format?.fontSize,
    toolbarCommandState(e, 'font.size').value,
  ];

  test('a shrunk title asks for a larger size, and its drawn size never goes down', async () => {
    const e = await shrinkEditor(110, '<p:ph type="title"/>');
    expect(sizes(e)).toEqual([18, 16]);
    expect(e.exec({ type: 'stepFontSize', direction: 1 }).ok).toBe(true);
    // 20 pt overflows at 100 % and at 90 % (18 pt), so the title stays at 90 % and overflows.
    expect(sizes(e)).toEqual([20, 18]);
  });

  test('shrunk body text asks for a larger size and keeps filling its box', async () => {
    const e = await shrinkEditor(61);
    expect(sizes(e)).toEqual([18, 11]);
    expect(e.exec({ type: 'stepFontSize', direction: 1 }).ok).toBe(true);
    // 20 pt fits 61 pt at 55 % (11 pt) with the 20 % reduction: the drawn size stays.
    expect(sizes(e)).toEqual([20, 11]);
  });

  test('Decrease steps down from the drawn size, so the size shown always goes down', async () => {
    const e = await shrinkEditor(110, '<p:ph type="title"/>');
    expect(sizes(e)).toEqual([18, 16]);
    expect(e.exec({ type: 'stepFontSize', direction: -1 }).ok).toBe(true);
    // From 16 drawn, not 18 stored: 14 pt fits at 100 %.
    expect(sizes(e)).toEqual([14, 14]);
  });

  test('Decrease over all the text of a shrunk body steps from the drawn size', async () => {
    const e = await shrinkEditor(110, '<p:ph type="title"/>');
    const id = e.session!.layout(0)!.items[0].element.nodeId;
    e.setSelection({
      kind: 'text',
      nodeId: id,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 4, offset: 1 },
    });
    expect(sizes(e)).toEqual([18, 16]);
    expect(e.exec({ type: 'stepFontSize', direction: -1 }).ok).toBe(true);
    expect(sizes(e)).toEqual([14, 14]);
  });

  test('Decrease on a word in shrunk text steps from the stored size, one step', async () => {
    const e = await shrinkEditor(61);
    expect(sizes(e)).toEqual([18, 11]);
    const id = e.session!.layout(0)!.items[0].element.nodeId;
    // A caret in the first paragraph: the body keeps its scale for the other paragraphs.
    e.setSelection({
      kind: 'text',
      nodeId: id,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 1 },
    });
    expect(e.exec({ type: 'stepFontSize', direction: -1 }).ok).toBe(true);
    expect(e.snapshot().format?.fontSize).toBe(16);
  });
});

describe('shapes shrunk by different scales', () => {
  test('the size box shows no size', async () => {
    const box = (id: number, y: number, bodyPr: string) =>
      shape({
        id,
        x: IN,
        y,
        cx: 4 * IN,
        cy: IN,
        text: para(run('x', 'sz="2800"')),
        bodyPr,
      });
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            box(2, IN, '<a:bodyPr wrap="square"><a:normAutofit fontScale="62500"/></a:bodyPr>') +
            box(3, 3 * IN, '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>'),
        },
      ],
    });
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(bytes)).ok).toBe(true);
    const ids = e.session!.layout(0)!.items.map((it) => it.element.nodeId);
    e.setSelection({ kind: 'shapes', nodeIds: ids });
    expect(e.snapshot().format?.fontSize).toBe(28);
    expect(e.snapshot().format?.renderedFontSize).toBeNull();
    expect(toolbarCommandState(e, 'font.size').value).toBeUndefined();
    // One shape alone shows its own drawn size.
    e.setSelection({ kind: 'shapes', nodeIds: [ids[0]!] });
    expect(toolbarCommandState(e, 'font.size').value).toBe(18);
  });

  test('an empty text box shows the size of its paragraph mark', async () => {
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: IN,
            text: '<a:p><a:endParaRPr lang="en-US" sz="2400"/></a:p>',
            bodyPr: '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>',
          }),
        },
      ],
    });
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(bytes)).ok).toBe(true);
    e.setSelection({ kind: 'shapes', nodeIds: [e.session!.layout(0)!.items[0].element.nodeId] });
    expect(e.snapshot().format?.renderedFontSize).toBe(24);
    expect(toolbarCommandState(e, 'font.size').value).toBe(24);
  });
});

describe('the text fitting control (paragraph.autofit)', () => {
  test('shows the fitting of the selected shape and writes the chosen one', async () => {
    const e = await shrinkEditor(61);
    expect(toolbarCommandState(e, 'paragraph.autofit')).toMatchObject({
      enabled: true,
      value: 'shrink',
    });
    const part = () => xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!;
    runToolbarCommand(e, 'paragraph.autofit', 'shape');
    expect(find(part().root, 'a:spAutoFit')).toBeDefined();
    expect(find(part().root, 'a:normAutofit')).toBeUndefined();
    expect(toolbarCommandState(e, 'paragraph.autofit').value).toBe('shape');
    runToolbarCommand(e, 'paragraph.autofit', 'none');
    expect(find(part().root, 'a:noAutofit')).toBeDefined();
    expect(toolbarCommandState(e, 'paragraph.autofit').value).toBe('none');
  });
});
