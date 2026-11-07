import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx, para, run, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx, serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

/** The fixture theme's third effect style is an outer shadow (the default theme's). */
const STYLE_REFS = (effectIdx: number): string =>
  '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>' +
  '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
  `<a:effectRef idx="${effectIdx}"><a:schemeClr val="accent1"/></a:effectRef>` +
  '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>';

const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, fill: solid('4472C4') }) +
          shape({ id: 3, x: 4 * IN, y: IN, cx: 2 * IN, cy: IN, style: STYLE_REFS(3) }) +
          shape({
            id: 4,
            x: 7 * IN,
            y: IN,
            cx: 2 * IN,
            cy: IN,
            fill: solid('ED7D31'),
            text: para(run('Text')),
          }),
      },
    ],
  });

const slideXml = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!);
const ids = (e: PptxEditorInstance): string[] =>
  e.session!.layout(0)!.items.map((i) => i.element.nodeId);
const effectsOf = (e: PptxEditorInstance, id: string) => {
  const item = e.session!.layout(0)!.items.find((i) => i.element.nodeId === id)!;
  return 'effects' in item.element ? item.element.effects : undefined;
};

describe('shape effects', () => {
  test('a shadow preset writes an effect list after the outline and shows in the snapshot', async () => {
    const e = await open(deck());
    const [rect] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [rect] });
    expect(e.snapshot().geometry?.effects?.shadow).toBe('none');
    expect(
      e.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'outerBottomRight' } }).ok
    ).toBe(true);
    expect(slideXml(e)).toContain(
      '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst>'
    );
    expect(e.snapshot().geometry?.effects?.shadow).toBe('outerBottomRight');
    expect(effectsOf(e, rect)?.outerShadow?.dist).toBe(38_100);
    // An inner shadow replaces the outer one: a shape has one shadow.
    expect(
      e.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'innerTop' } }).ok
    ).toBe(true);
    expect(e.snapshot().geometry?.effects?.shadow).toBe('innerTop');
    expect(effectsOf(e, rect)?.outerShadow).toBeUndefined();
    expect(e.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: null } }).ok).toBe(
      true
    );
    expect(e.snapshot().geometry?.effects?.shadow).toBe('none');
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(e.snapshot().geometry?.effects?.shadow).toBe('outerBottomRight');
    expect(readPptx(e.save()!).ok).toBe(true);
  });

  test('shadow off on a theme effect style writes an empty effect list', async () => {
    const e = await open(deck());
    const [, styled] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [styled] });
    // The theme's shadow is not a gallery preset.
    expect(e.snapshot().geometry?.effects?.shadow).toBe('other');
    expect(effectsOf(e, styled)?.outerShadow).toBeDefined();
    expect(e.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: null } }).ok).toBe(
      true
    );
    expect(slideXml(e)).toContain('<a:effectLst/></p:spPr>');
    expect(effectsOf(e, styled)?.outerShadow).toBeUndefined();
    expect(e.snapshot().geometry?.effects?.shadow).toBe('none');
  });

  test('an effect on a theme effect style that uses a:effectDag is refused', async () => {
    const files = unzipSync(deck());
    const k = 'ppt/theme/theme1.xml';
    const theme = strFromU8(files[k]!);
    const dag = theme.replace(
      /<a:effectStyle><a:effectLst><a:outerShdw[\s\S]*?<\/a:effectLst><\/a:effectStyle>/,
      '<a:effectStyle><a:effectDag><a:cont/><a:glow rad="101600"><a:srgbClr val="FF0000"/></a:glow></a:effectDag></a:effectStyle>'
    );
    expect(dag).not.toBe(theme);
    files[k] = strToU8(dag);
    const e = await open(zipSync(files));
    const [, styled] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [styled] });
    const xml = slideXml(e);
    // An own effect list would replace the theme's DAG effects with no message.
    expect(
      e.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'outerBottom' } })
    ).toMatchObject({
      ok: false,
      reason: 'errors.unsupported',
    });
    expect(slideXml(e)).toBe(xml);
  });

  test('No shadow on an own a:effectDag says the DAG is not supported', async () => {
    const files = unzipSync(deck());
    const k = 'ppt/slides/slide1.xml';
    const dag =
      '<a:effectDag><a:cont/><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:prstClr val="black"/></a:outerShdw></a:effectDag></p:spPr>';
    files[k] = strToU8(strFromU8(files[k]!).replace('</p:spPr>', dag));
    const e = await open(zipSync(files));
    e.setSelection({ kind: 'shapes', nodeIds: [ids(e)[0]!] });
    expect(
      e.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: null } })
    ).toMatchObject({ ok: false, reason: 'errors.unsupported' });
  });

  test('another effect on a theme effect style keeps the theme shadow', async () => {
    const e = await open(deck());
    const [, styled] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [styled] });
    expect(
      e.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 8, color: '#FF0000' } }).ok
    ).toBe(true);
    const xml = slideXml(e);
    // Schema order: glow before outerShdw; the theme's shadow is copied into the shape.
    expect(xml).toContain(
      '<a:effectLst><a:glow rad="101600"><a:srgbClr val="FF0000"><a:alpha val="40000"/></a:srgbClr></a:glow><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">'
    );
    const fx = e.snapshot().geometry?.effects;
    expect(fx?.glow).toBe(8);
    expect(fx?.glowColor).toBe('#FF0000');
    expect(fx?.shadow).toBe('other');
  });

  test('a soft edge on a theme effect style keeps the theme shadow after it', async () => {
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, style: STYLE_REFS(3) }),
        },
      ],
    });
    const e = await open(bytes);
    e.setSelection({ kind: 'shapes', nodeIds: [ids(e)[0]] });
    expect(e.exec({ type: 'setShapeEffect', effect: { kind: 'softEdge', size: 5 } }).ok).toBe(true);
    expect(slideXml(e)).toContain('</a:outerShdw><a:softEdge rad="63500"/></a:effectLst>');
    expect(e.snapshot().geometry?.effects?.softEdge).toBe(5);
  });

  test('glow size keeps the glow color; reflection and soft edge set and clear', async () => {
    const e = await open(deck());
    const [rect] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [rect] });
    expect(e.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 5 } }).ok).toBe(true);
    expect(slideXml(e)).toContain(
      '<a:glow rad="63500"><a:schemeClr val="accent1"><a:satMod val="175000"/><a:alpha val="40000"/></a:schemeClr></a:glow>'
    );
    expect(e.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 18 } }).ok).toBe(true);
    expect(slideXml(e)).toContain('<a:glow rad="228600"><a:schemeClr val="accent1">');
    expect(
      e.exec({ type: 'setShapeEffect', effect: { kind: 'reflection', preset: 'half4pt' } }).ok
    ).toBe(true);
    expect(e.exec({ type: 'setShapeEffect', effect: { kind: 'softEdge', size: 2.5 } }).ok).toBe(
      true
    );
    const fx = e.snapshot().geometry!.effects!;
    expect([fx.glow, fx.reflection, fx.softEdge]).toEqual([18, 'half4pt', 2.5]);
    expect(slideXml(e)).toMatch(/<a:glow [^>]*>.*<\/a:glow><a:reflection [^>]*\/><a:softEdge/);
    for (const kind of ['glow', 'softEdge'] as const)
      expect(e.exec({ type: 'setShapeEffect', effect: { kind, size: null } }).ok).toBe(true);
    expect(
      e.exec({ type: 'setShapeEffect', effect: { kind: 'reflection', preset: null } }).ok
    ).toBe(true);
    expect(e.snapshot().geometry!.effects).toEqual({
      shadow: 'none',
      reflection: 'none',
      glow: 0,
      softEdge: 0,
      bevel: 'none',
      rotation3d: 'none',
      blur: 0,
    });
  });

  test('invalid effects and targets are refused with a reason', async () => {
    const e = await open(deck());
    const [rect] = ids(e);
    const shadow = { type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'outerTop' } };
    expect(e.can(shadow as never)).toEqual({ ok: false, reason: 'disabled.noSelection' });
    e.setSelection({ kind: 'shapes', nodeIds: [rect] });
    const bad = (effect: unknown) => e.exec({ type: 'setShapeEffect', effect } as never);
    expect(bad({ kind: 'shadow', preset: 'sideways' })).toMatchObject({
      ok: false,
      reason: 'errors.unsupported',
    });
    expect(bad({ kind: 'glow', size: 1000 })).toMatchObject({ ok: false });
    expect(bad({ kind: 'glow', size: Number.NaN })).toMatchObject({ ok: false });
    expect(bad({ kind: 'blur', size: -3 })).toMatchObject({ ok: false });
    expect(bad({ kind: 'bevel', preset: 'round' })).toMatchObject({
      ok: false,
      reason: 'errors.unsupported',
    });
    expect(bad({ kind: 'rotation3d', preset: 'perspectiveHeroic' })).toMatchObject({
      ok: false,
      reason: 'errors.unsupported',
    });
    expect(bad({ kind: 'glow', size: 5, color: 'red' })).toMatchObject({ ok: false });
    expect(slideXml(e)).not.toContain('effectLst');
  });
});

describe('shape styles gallery', () => {
  test('a theme style writes p:style and removes the own fill, outline, and effects', async () => {
    const e = await open(deck());
    const [rect, , text] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [rect, text] });
    expect(
      e.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'outerTop' } }).ok
    ).toBe(true);
    expect(
      e.exec({ type: 'setShapeStyle', style: { row: 'intenseEffect', color: 'accent2' } }).ok
    ).toBe(true);
    const xml = slideXml(e);
    expect(xml).not.toContain('4472C4');
    expect(xml).not.toContain('outerShdw blurRad="50800"');
    expect(xml).toContain(
      '</p:spPr><p:style><a:lnRef idx="0"><a:schemeClr val="accent2"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent2"/></a:fillRef><a:effectRef idx="3"><a:schemeClr val="accent2"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody>'
    );
    e.setSelection({ kind: 'shapes', nodeIds: [rect] });
    const g = e.snapshot().geometry!;
    expect(g.shapeStyle).toEqual({ row: 'intenseEffect', color: 'accent2' });
    // The fill and the shadow now come from the theme's style matrix.
    expect(g.fill).not.toBe('none');
    expect(g.effects?.shadow).toBe('other');
    expect(e.exec({ type: 'setShapeStyle', style: { row: 'coloredFill', color: 'dk1' } }).ok).toBe(
      true
    );
    expect(e.snapshot().geometry!.shapeStyle).toEqual({ row: 'coloredFill', color: 'dk1' });
    expect(slideXml(e)).toContain(
      '<a:lnRef idx="2"><a:schemeClr val="dk1"><a:shade val="15000"/></a:schemeClr></a:lnRef>'
    );
    expect(readPptx(e.save()!).ok).toBe(true);
  });

  test('a new shape matches the Colored Fill row; bad cells are refused', async () => {
    const e = await open(deck());
    expect(e.exec({ type: 'insertShape', preset: 'rect' }).ok).toBe(true);
    expect(e.snapshot().geometry?.shapeStyle).toEqual({ row: 'coloredFill', color: 'accent1' });
    const bad = (style: unknown) => e.exec({ type: 'setShapeStyle', style } as never);
    expect(bad({ row: 'loud', color: 'accent1' })).toMatchObject({ ok: false });
    expect(bad({ row: 'coloredFill', color: 'accent9' })).toMatchObject({ ok: false });
  });
});

describe('Shape Format fixtures', () => {
  const fixture = (name: string): Uint8Array =>
    new Uint8Array(readFileSync(join(import.meta.dir, '../../../../../fixtures', `${name}.pptx`)));
  const geometryAt = (e: PptxEditorInstance, slide: number, index: number) => {
    e.goToSlide(slide);
    const id = e.session!.layout(slide)!.items[index]!.element.nodeId;
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    return e.snapshot().geometry!;
  };

  test('the effect presets read back as the gallery choices', async () => {
    const e = await open(fixture('shape-effects-presets'));
    const shadows = [
      'outerBottomRight',
      'outerBottom',
      'outerTopLeft',
      'outerLeft',
      'outerCenter',
      'innerTopLeft',
      'innerBottom',
      'innerRight',
      'innerCenter',
    ];
    shadows.forEach((s, i) => expect(geometryAt(e, 0, i).effects?.shadow).toBe(s as never));
    const second = [0, 1, 2, 3, 4, 5, 6, 7].map((i) => geometryAt(e, 1, i).effects!);
    expect(second.slice(0, 3).map((f) => f.reflection)).toEqual([
      'tightTouching',
      'half4pt',
      'full8pt',
    ]);
    expect(second.slice(3, 5).map((f) => f.glow)).toEqual([5, 18]);
    expect(second.slice(5, 7).map((f) => f.softEdge)).toEqual([2.5, 10]);
    expect([second[7]!.glow, second[7]!.shadow]).toEqual([8, 'outerBottomRight']);
    const third = [0, 1, 2, 3, 4].map((i) => geometryAt(e, 2, i).effects!.shadow);
    expect(third).toEqual(['none', 'none', 'other', 'none', 'other']);
  });

  test('the Shape Styles rows read back as gallery cells', async () => {
    const e = await open(fixture('shape-styles'));
    const rows = [
      'coloredOutline',
      'coloredFill',
      'lightOutline',
      'subtleEffect',
      'moderateEffect',
      'intenseEffect',
    ];
    const colors = ['accent1', 'accent2', 'accent6', 'dk1'];
    rows.forEach((row, r) =>
      colors.forEach((color, c) =>
        expect(geometryAt(e, 0, r * 4 + c).shapeStyle).toEqual({ row, color } as never)
      )
    );
  });

  test('flips, the lock, the hidden shape, and the distributed boxes read back', async () => {
    const e = await open(fixture('arrange-flip-hidden'));
    const names = e.snapshot().shapes!.map((s) => [s.name, s.hidden]);
    expect(names).toContainEqual(['Hidden shape', true]);
    expect(e.session!.layout(0)!.items.map((i) => i.element.name)).not.toContain('Hidden shape');
    const turned = geometryAt(e, 0, 3);
    expect([turned.rot, turned.flipH]).toEqual([330, true]);
    expect(geometryAt(e, 0, 5).lockAspectRatio).toBe(true);
    const boxes = e
      .session!.layout(0)!
      .items.slice(-3)
      .map((i) => i.element.nodeId);
    e.setSelection({ kind: 'shapes', nodeIds: boxes });
    const before = slideXml(e);
    // The fixture's boxes are already evenly spaced: nothing to do, and no undo step.
    expect(e.exec({ type: 'distributeObjects', axis: 'horizontal' })).toMatchObject({
      ok: false,
      reason: 'disabled.noArrangeChange',
    });
    expect(slideXml(e)).toBe(before);
  });
});

describe('effect state and refusals', () => {
  test('an outer preset keeps its slide direction on a turned shape, as the gallery writes it', async () => {
    // The Shadow gallery, Offset: Bottom Right, on a shape at rot 90° (r53-1):
    // rotWithShape="0", and the shadow falls to the slide's bottom right.
    const turned = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 3 * IN,
            cy: IN,
            rot: 90 * 60_000,
            fill: solid('4472C4'),
          }),
        },
      ],
    });
    const e = await open(turned);
    e.setSelection({ kind: 'shapes', nodeIds: [ids(e)[0]!] });
    expect(
      e.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'outerBottomRight' } }).ok
    ).toBe(true);
    expect(slideXml(e)).toContain(
      '<a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0">'
    );
    expect(effectsOf(e, ids(e)[0]!)?.outerShadow?.rotWithShape).toBe(false);
    expect(e.snapshot().geometry?.effects?.shadow).toBe('outerBottomRight');
    // Inner shadows have no rotWithShape (CT_InnerShadowEffect).
    e.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'innerTop' } });
    expect(slideXml(e)).toContain('<a:innerShdw ');
    expect(slideXml(e)).not.toMatch(/<a:innerShdw[^>]*rotWithShape/);
  });

  test('an outer and an inner shadow together read as other', async () => {
    const files = unzipSync(deck());
    const k = 'ppt/slides/slide1.xml';
    const both =
      '<a:effectLst><a:innerShdw blurRad="63500" dist="50800" dir="2700000"><a:prstClr val="black"><a:alpha val="50000"/></a:prstClr></a:innerShdw>' +
      '<a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst></p:spPr>';
    files[k] = strToU8(strFromU8(files[k]!).replace('</p:spPr>', both));
    const e = await open(zipSync(files));
    e.setSelection({ kind: 'shapes', nodeIds: [ids(e)[0]!] });
    expect(effectsOf(e, ids(e)[0]!)?.innerShadow).toBeDefined();
    expect(e.snapshot().geometry?.effects?.shadow).toBe('other');
  });

  test('a reflection or shadow that is not a gallery preset reads as other', async () => {
    const e = await open(
      new Uint8Array(readFileSync(join(import.meta.dir, '../../../../../fixtures/effects.pptx')))
    );
    e.goToSlide(1);
    const items = e.session!.layout(1)!.items;
    const at = (i: number) => {
      e.setSelection({ kind: 'shapes', nodeIds: [items[i]!.element.nodeId] });
      return e.snapshot().geometry!.effects!;
    };
    // blurRad 38100 and stA 38000: not a gallery reflection, though endPos and dist match one.
    expect(at(4).reflection).toBe('other');
    // A skewed perspective shadow is not an Offset preset.
    expect(at(0).shadow).toBe('other');
  });

  test('removing an effect the shape lacks changes nothing; unknown kinds are refused', async () => {
    const e = await open(deck());
    const [rect] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [rect] });
    const before = slideXml(e);
    expect(
      e.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: null } })
    ).toMatchObject({ ok: false, reason: 'disabled.noEffectChange', noop: true });
    expect(slideXml(e)).toBe(before);
    for (const kind of ['constructor', 'toString', '__proto__'])
      expect(e.exec({ type: 'setShapeEffect', effect: { kind, size: 5 } } as never)).toMatchObject({
        ok: false,
        reason: 'errors.unsupported',
      });
  });

  test('a glow color change keeps the glow size', async () => {
    const e = await open(deck());
    const [rect] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [rect] });
    // 7.3333 pt: the chrome shows 7.33.
    e.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 7.3333 } });
    const rad = /<a:glow rad="(\d+)"/.exec(slideXml(e))![1];
    const shown = e.snapshot().geometry!.effects!.glow;
    expect(
      e.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: shown, color: '#FF0000' } }).ok
    ).toBe(true);
    expect(slideXml(e)).toContain(`<a:glow rad="${rad}"><a:srgbClr val="FF0000">`);
  });
});

describe('effects a placeholder inherits', () => {
  const TITLE =
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4000000" cy="1000000"/></a:xfrm></p:spPr>' +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>T</a:t></a:r></a:p></p:txBody></p:sp>';
  const SHADOW =
    '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst>';

  async function titleWithLayoutShadow() {
    const files = unzipSync(buildPptx({ slides: [{ layout: 'title', shapes: TITLE }] }));
    const k = 'ppt/slideLayouts/slideLayout1.xml';
    files[k] = strToU8(strFromU8(files[k]).replace('<p:spPr/>', `<p:spPr>${SHADOW}</p:spPr>`));
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(zipSync(files))).ok).toBe(true);
    const id = editor.snapshot().shapes!.find((s) => s.name === 'Title 1')!.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [id] });
    return editor;
  }

  test('No shadow removes a shadow from the layout placeholder', async () => {
    const editor = await titleWithLayoutShadow();
    expect(editor.snapshot().geometry?.effects?.shadow).not.toBe('none');
    const r = editor.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: null } });
    expect(r.ok).toBe(true);
    expect(editor.snapshot().geometry?.effects?.shadow).toBe('none');
  });

  test('a new effect keeps the inherited shadow', async () => {
    const editor = await titleWithLayoutShadow();
    const before = editor.snapshot().geometry?.effects?.shadow;
    expect(editor.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 5 } }).ok).toBe(
      true
    );
    expect(editor.snapshot().geometry?.effects).toMatchObject({ shadow: before, glow: 5 });
  });
});
