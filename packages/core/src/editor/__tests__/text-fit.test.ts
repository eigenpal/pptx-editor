// `readTextFit` against the layout that a shape picture paints (`DocumentSession.layout`, which
// `renderImageWith` and `renderImagePng` draw): the same numbers for each autofit kind of
// `a:bodyPr` (ECMA-376 Part 1, 21.1.2.1.1), after a text write re-fits the body.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  createAutomationHost,
  createStoreAutomationHost,
  storeAutomationPort,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
} from '../../automation/index.ts';
import { createFixedMeasurer, type LayoutItem } from '../../layout/index.ts';
import { openPresentation, type DocumentSession } from '../index.ts';

const IN = 914_400;
const EMU_PER_PT = 12_700;
const LONG = 'The quick brown fox jumps over the lazy dog. '.repeat(12);
// Text that overflows the box at full size and fits at a shrink step.
const SHRINKS = 'The quick brown fox jumps over the lazy dog. '.repeat(3);
const BODIES = [
  '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>',
  '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>',
  '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>',
  '<a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr>',
  '<a:bodyPr vert="vert270" wrap="square" lIns="0" rIns="182880"><a:noAutofit/></a:bodyPr>',
];

async function open(): Promise<DocumentSession> {
  const opened = await openPresentation(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: BODIES.map((bodyPr, i) =>
            shape({
              id: i + 2,
              x: IN,
              y: IN + i * IN,
              cx: 3 * IN,
              cy: IN,
              text: para(run('a')),
              bodyPr,
            })
          ).join(''),
        },
      ],
    }),
    { substitutes: false }
  );
  if (!opened.ok) throw new Error('fixture');
  return opened.session;
}

function ok(host: AutomationHost, operations: AutomationOperation[]): AutomationValue[] {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.map((r) => (r.status === 'ok' ? r.value : { kind: 'null' }));
}

const shapeOp = (id: string): AutomationOperation => ({
  op: 'getShape',
  container: { result: 0 },
  key: { id },
});

/** The painted item of the shape `id` on slide 0. */
function painted(session: DocumentSession, id: string): Extract<LayoutItem, { kind: 'shape' }> {
  const item = session
    .layout(0)!
    .items.find((i) => i.element.layer === 'slide' && i.element.drawingId === id);
  if (item?.kind !== 'shape' || !item.text) throw new Error(`no text item ${id}`);
  return item;
}

const points = (emu: number): number => Math.round((emu / EMU_PER_PT) * 100) / 100;

describe('readTextFit', () => {
  test('answers what the painted layout of each shape holds', async () => {
    const session = await open();
    const host = createStoreAutomationHost(session.store, { measurer: session.measurer });
    const texts = [LONG, SHRINKS, LONG.repeat(4), LONG, LONG];
    const ops: AutomationOperation[] = [{ op: 'getSlide', key: { index: 0 } }];
    texts.forEach((text, i) => {
      ops.push(shapeOp(String(i + 2)));
      ops.push({ op: 'setText', shape: { result: ops.length - 1 }, text });
    });
    // The reads come after every write of the batch: each reads the re-fit body.
    const firstShape = 1;
    texts.forEach((_, i) => ops.push({ op: 'readTextFit', shape: { result: firstShape + 2 * i } }));
    const results = ok(host, ops).slice(-texts.length);
    const overflows: boolean[] = [];
    results.forEach((value, i) => {
      if (value.kind !== 'textFit') throw new Error('not a fit');
      const item = painted(session, String(i + 2));
      const text = item.text!;
      // Horizontal text keeps its insets; `vert270` takes the left and right insets across.
      const [top, bottom] = i === 4 ? [0, 182_880] : [45_720, 45_720];
      expect(value.textFit).toEqual({
        overflows: text.contentHeight > text.box.height + 1,
        textHeight: points(text.contentHeight - top - bottom),
        availableHeight: points(text.box.height - top - bottom),
        substitutedFonts: [...text.substitutions],
      });
      overflows.push(value.textFit.overflows);
    });
    // Raw overflow, a shrink step that fits, the last step that does not, a grown shape, and
    // vertical text across a three-inch frame.
    expect(overflows).toEqual([true, false, true, false, true]);
    // The fitting body took a shrink step; the other stopped at the last step (25 %).
    const autofitOf = (id: string) => painted(session, id).element.text!.body.autofit;
    const shrunk = autofitOf('3');
    expect(
      shrunk.kind === 'normal' && shrunk.fontScale * (1 - shrunk.lineSpaceReduction)
    ).toBeLessThan(1);
    expect(autofitOf('4')).toMatchObject({ kind: 'normal', fontScale: 0.25 });
    session.dispose();
  });

  test('a host with no measurer, the fixed measurer, or a stand-in measurer refuses', async () => {
    const session = await open();
    // An editor's port while its canvas measurer stands in for the fonts.
    const standIn = createAutomationHost({
      port: {
        ...storeAutomationPort(session.store, { measurer: session.measurer }),
        standInMeasurer: () => true,
      },
      capabilities: {
        document: true,
        save: true,
        events: true,
        selection: false,
        images: false,
        autofit: true,
      },
    });
    for (const host of [
      createStoreAutomationHost(session.store),
      createStoreAutomationHost(session.store, { measurer: createFixedMeasurer() }),
      standIn,
    ]) {
      const response = host.execute({
        operations: [
          { op: 'getSlide', key: { index: 0 } },
          shapeOp('2'),
          { op: 'readTextFit', shape: { result: 1 } },
        ],
      });
      expect(response.ok).toBe(false);
      expect(response.results[2]).toMatchObject({
        status: 'error',
        error: { code: 'unsupported-capability', detail: 'text-fit-measurer' },
      });
      host.dispose();
    }
    session.dispose();
  });
});
