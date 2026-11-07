import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const deck = (warp: string): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 3 * IN,
            cy: 2 * IN,
            fill: solid('4472C4'),
            bodyPr: `<a:bodyPr wrap="none" fromWordArt="1" rtlCol="0">${warp}<a:noAutofit/></a:bodyPr>`,
            text: para(run('Warp')),
          }) + shape({ id: 3, x: 5 * IN, y: IN, cx: IN, cy: IN, fill: solid('ED7D31') }),
      },
    ],
  });

const slideXml = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!);
const itemOf = (e: PptxEditorInstance, i: number) => e.session!.layout(0)!.items[i];
const warpedOf = (e: PptxEditorInstance) => {
  const item = itemOf(e, 0) as { text?: { warped?: readonly unknown[] } };
  return item.text?.warped;
};

describe('text warp', () => {
  test('a warp draws the glyphs as paths, and the command sets and removes it', async () => {
    const e = await open(deck(''));
    const id = itemOf(e, 0).element.nodeId;
    expect(warpedOf(e)).toBeUndefined();
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    expect(e.snapshot().geometry?.textWarp).toBe('textNoShape');

    expect(e.exec({ type: 'setTextWarp', preset: 'textWave1', adjust: { adj1: 99_999 } }).ok).toBe(
      true
    );
    // The adjust value stays in the handle range; bodyPr keeps its attributes and the autofit.
    expect(slideXml(e)).toContain(
      '<a:bodyPr wrap="none" fromWordArt="1" rtlCol="0"><a:prstTxWarp prst="textWave1"><a:avLst><a:gd name="adj1" fmla="val 20000"/></a:avLst></a:prstTxWarp><a:noAutofit/></a:bodyPr>'
    );
    expect(e.snapshot().geometry?.textWarp).toBe('textWave1');
    expect(warpedOf(e)?.length).toBeGreaterThan(0);

    // The same preset keeps the adjust values that the command does not pass.
    expect(e.exec({ type: 'setTextWarp', preset: 'textWave1', adjust: { adj2: -5000 } }).ok).toBe(
      true
    );
    expect(slideXml(e)).toContain(
      '<a:avLst><a:gd name="adj1" fmla="val 20000"/><a:gd name="adj2" fmla="val -5000"/></a:avLst>'
    );

    expect(e.exec({ type: 'setTextWarp', preset: null }).ok).toBe(true);
    expect(slideXml(e)).not.toContain('prstTxWarp');
    expect(warpedOf(e)).toBeUndefined();
    expect(e.exec({ type: 'setTextWarp', preset: null })).toMatchObject({
      ok: false,
      reason: 'disabled.noEffectChange',
    });

    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(e.snapshot().geometry?.textWarp).toBe('textWave1');
  });

  test('the command refuses unknown names, unknown guides, and shapes without text', async () => {
    const e = await open(deck(''));
    const [text, rect] = [0, 1].map((i) => itemOf(e, i).element.nodeId);
    e.setSelection({ kind: 'shapes', nodeIds: [text] });
    const bad = { type: 'setTextWarp', preset: 'textNope' } as unknown as Parameters<
      PptxEditorInstance['exec']
    >[0];
    expect(e.exec(bad)).toMatchObject({ ok: false, reason: 'errors.unsupported' });
    expect(
      e.exec({ type: 'setTextWarp', preset: 'textWave1', adjust: { 'x"/><a:bad': 1 } })
    ).toMatchObject({ ok: false, reason: 'errors.unsupported' });
    expect(slideXml(e)).not.toContain('prstTxWarp');
    e.setSelection({ kind: 'shapes', nodeIds: [rect] });
    expect(e.exec({ type: 'setTextWarp', preset: 'textArchUp' })).toMatchObject({
      ok: false,
      reason: 'disabled.noText',
    });
  });

  test('a kept adjust value outside the handle range comes back in range', async () => {
    const e = await open(
      deck(
        '<a:prstTxWarp prst="textWave1"><a:avLst><a:gd name="adj1" fmla="val 99999999999999999999999"/></a:avLst></a:prstTxWarp>'
      )
    );
    e.setSelection({ kind: 'shapes', nodeIds: [itemOf(e, 0).element.nodeId] });
    expect(e.exec({ type: 'setTextWarp', preset: 'textWave1', adjust: { adj2: 0 } }).ok).toBe(true);
    expect(slideXml(e)).toContain(
      '<a:avLst><a:gd name="adj1" fmla="val 20000"/><a:gd name="adj2" fmla="val 0"/></a:avLst>'
    );
    expect(slideXml(e)).not.toContain('e+');
  });

  test('typing in warped text keeps the warp', async () => {
    const warp =
      '<a:prstTxWarp prst="textArchUp"><a:avLst><a:gd name="adj" fmla="val 10800000"/></a:avLst></a:prstTxWarp>';
    const e = await open(deck(warp));
    const nodeId = itemOf(e, 0).element.nodeId;
    const at = { paragraph: 0, offset: 0 };
    e.setSelection({ kind: 'text', nodeId, anchor: at, focus: at });
    for (const ch of 'New ') expect(e.input('text', ch).ok).toBe(true);
    expect(slideXml(e)).toContain(`${warp}<a:noAutofit/>`);
    expect(slideXml(e)).toContain('<a:t>New Warp</a:t>');
    expect(warpedOf(e)?.length).toBeGreaterThan(0);
  });

  test('removing a warp that a layout placeholder gives writes textNoShape', async () => {
    const ph = '<p:ph type="body" idx="1"/>';
    const warp =
      '<a:prstTxWarp prst="textWave1"><a:avLst><a:gd name="adj1" fmla="val 15000"/></a:avLst></a:prstTxWarp>';
    const e = await open(
      buildPptx({
        layouts: [
          {
            name: 'Warped',
            shapes: shape({
              id: 2,
              x: IN,
              y: IN,
              cx: 3 * IN,
              cy: 2 * IN,
              ph,
              bodyPr: `<a:bodyPr wrap="none">${warp}</a:bodyPr>`,
              text: para(run('Prompt')),
            }),
          },
        ],
        slides: [
          {
            layoutName: 'Warped',
            shapes: shape({
              id: 2,
              x: IN,
              y: IN,
              cx: 3 * IN,
              cy: 2 * IN,
              ph,
              bodyPr: '<a:bodyPr/>',
              text: para(run('Inherited')),
            }),
          },
        ],
      })
    );
    const id = itemOf(e, 0).element.nodeId;
    expect(warpedOf(e)?.length).toBeGreaterThan(0);
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    expect(e.snapshot().geometry?.textWarp).toBe('textWave1');
    // The same preset keeps the inherited adjust values.
    expect(e.exec({ type: 'setTextWarp', preset: 'textWave1', adjust: { adj2: 1000 } }).ok).toBe(
      true
    );
    expect(slideXml(e)).toContain(
      '<a:avLst><a:gd name="adj1" fmla="val 15000"/><a:gd name="adj2" fmla="val 1000"/></a:avLst>'
    );
    expect(e.exec({ type: 'setTextWarp', preset: null }).ok).toBe(true);
    expect(slideXml(e)).toContain(
      '<a:bodyPr><a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp></a:bodyPr>'
    );
    expect(warpedOf(e)).toBeUndefined();
    expect(e.snapshot().geometry?.textWarp).toBe('textNoShape');
    expect(e.exec({ type: 'setTextWarp', preset: null })).toMatchObject({
      ok: false,
      reason: 'disabled.noEffectChange',
    });
  });

  test('a saved warp parses, and an unknown name or textNoShape draws flat', async () => {
    const arch = await open(
      deck(
        '<a:prstTxWarp prst="textArchUp"><a:avLst><a:gd name="adj" fmla="val 10800000"/></a:avLst></a:prstTxWarp>'
      )
    );
    expect(warpedOf(arch)?.length).toBeGreaterThan(0);
    for (const prst of ['textNoShape', 'textNope']) {
      const e = await open(deck(`<a:prstTxWarp prst="${prst}"><a:avLst/></a:prstTxWarp>`));
      expect(warpedOf(e)).toBeUndefined();
      // The unknown element stays in the saved part.
      expect(slideXml(e)).toContain(`<a:prstTxWarp prst="${prst}">`);
    }
  });
});
