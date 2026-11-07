import { describe, expect, test } from 'bun:test';
import { buildPptx, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx, serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

/** A shape with an outline and an effect list, and one with its own 3-D properties. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 2 * IN,
            cy: IN,
            fill: solid('4472C4'),
            line: '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
            effects:
              '<a:effectLst><a:glow rad="63500"><a:srgbClr val="FFC000"/></a:glow></a:effectLst>',
          }) +
          shape({
            id: 3,
            x: 4 * IN,
            y: IN,
            cx: 2 * IN,
            cy: IN,
            fill: solid('ED7D31'),
            effects:
              '<a:scene3d><a:camera prst="orthographicFront" zoom="100000"/><a:lightRig rig="balanced" dir="tl"/></a:scene3d>' +
              '<a:sp3d extrusionH="12700" prstMaterial="metal"><a:contourClr><a:srgbClr val="C00000"/></a:contourClr></a:sp3d>',
          }),
      },
    ],
  });

const slideXml = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!);
const ids = (e: PptxEditorInstance): string[] =>
  e.session!.layout(0)!.items.map((i) => i.element.nodeId);
const spPrOf = (xml: string, name: string): string =>
  new RegExp(`name="${name}".*?<p:spPr>(.*?)</p:spPr>`).exec(xml)?.[1] ?? '';
const fx = (e: PptxEditorInstance) => e.snapshot().geometry!.effects!;

describe('bevel and 3-D rotation', () => {
  test('a bevel writes a front scene and a top bevel after the effect list', async () => {
    const e = await open(deck());
    e.setSelection({ kind: 'shapes', nodeIds: [ids(e)[0]!] });
    expect(fx(e)).toMatchObject({ bevel: 'none', rotation3d: 'none' });
    expect(e.exec({ type: 'setShapeEffect', effect: { kind: 'bevel', preset: 'circle' } }).ok).toBe(
      true
    );
    const spPr = spPrOf(slideXml(e), 'Shape 2');
    expect(spPr).toMatch(
      /<\/a:ln><a:effectLst>.*<\/a:effectLst><a:scene3d><a:camera prst="orthographicFront"\/><a:lightRig rig="threePt" dir="t"\/><\/a:scene3d><a:sp3d><a:bevelT w="76200" h="76200"\/><\/a:sp3d>$/
    );
    expect(fx(e).bevel).toBe('circle');
    // A second preset changes the bevel in place.
    e.exec({ type: 'setShapeEffect', effect: { kind: 'bevel', preset: 'angle' } });
    expect(spPrOf(slideXml(e), 'Shape 2')).toContain(
      '<a:bevelT w="76200" h="76200" prst="angle"/>'
    );
    expect(fx(e).bevel).toBe('angle');
    e.exec({ type: 'setShapeEffect', effect: { kind: 'bevel', preset: null } });
    expect(spPrOf(slideXml(e), 'Shape 2')).not.toContain('bevelT');
    expect(fx(e).bevel).toBe('none');
    expect(
      e.exec({ type: 'setShapeEffect', effect: { kind: 'bevel', preset: null } })
    ).toMatchObject({
      ok: false,
      reason: 'disabled.noEffectChange',
    });
  });

  test('the camera preset changes in place, and its own attributes and children stay', async () => {
    const e = await open(deck());
    e.setSelection({ kind: 'shapes', nodeIds: [ids(e)[1]!] });
    e.exec({ type: 'setShapeEffect', effect: { kind: 'rotation3d', preset: 'isometricTopUp' } });
    e.exec({ type: 'setShapeEffect', effect: { kind: 'bevel', preset: 'softRound' } });
    const spPr = spPrOf(slideXml(e), 'Shape 3');
    expect(spPr).toContain(
      '<a:scene3d><a:camera prst="isometricTopUp" zoom="100000"/><a:lightRig rig="balanced" dir="tl"/></a:scene3d>'
    );
    expect(spPr).toContain(
      '<a:sp3d extrusionH="12700" prstMaterial="metal"><a:bevelT w="76200" h="76200" prst="softRound"/><a:contourClr>'
    );
    expect(fx(e)).toMatchObject({ bevel: 'softRound', rotation3d: 'isometricTopUp' });
    e.exec({ type: 'setShapeEffect', effect: { kind: 'rotation3d', preset: null } });
    expect(spPrOf(slideXml(e), 'Shape 3')).toContain(
      '<a:camera prst="orthographicFront" zoom="100000"/>'
    );
    expect(fx(e).rotation3d).toBe('none');
    // Undo restores the turned camera.
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(fx(e).rotation3d).toBe('isometricTopUp');
  });

  test('blur sets and clears in the effect list', async () => {
    const e = await open(deck());
    e.setSelection({ kind: 'shapes', nodeIds: [ids(e)[0]!] });
    e.exec({ type: 'setShapeEffect', effect: { kind: 'blur', size: 4 } });
    expect(spPrOf(slideXml(e), 'Shape 2')).toContain('<a:effectLst><a:blur rad="50800"/><a:glow');
    expect(fx(e).blur).toBe(4);
    e.exec({ type: 'setShapeEffect', effect: { kind: 'blur', size: null } });
    expect(fx(e).blur).toBe(0);
  });

  test('the saved copy keeps the 3-D properties and opens again', async () => {
    const e = await open(deck());
    e.setSelection({ kind: 'shapes', nodeIds: [ids(e)[0]!] });
    e.exec({ type: 'setShapeEffect', effect: { kind: 'bevel', preset: 'convex' } });
    e.exec({ type: 'setShapeEffect', effect: { kind: 'rotation3d', preset: 'obliqueTopLeft' } });
    const reopened = readPptx(e.save()!);
    if (!reopened.ok) throw new Error('reopen');
    const xml = serializeXmlPart(xmlPartOf(reopened.pkg, '/ppt/slides/slide1.xml')!);
    expect(spPrOf(xml, 'Shape 2')).toContain('<a:camera prst="obliqueTopLeft"/>');
    expect(spPrOf(xml, 'Shape 2')).toContain('<a:bevelT w="76200" h="76200" prst="convex"/>');
  });
});
