import { afterAll, afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { buildPptx, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import {
  drawingCapabilities,
  drawingRenderersVersion,
  installDrawingRenderers,
  subscribeDrawingRenderers,
  type DrawingEffectsRenderer,
} from '../../layout/index.ts';
import { chromeOptionNeedsEffects } from '../chrome-offered.ts';
import {
  CHROME_GROUPS,
  chromeControlForSlot,
  chromeOfferedControl,
  commandForSlotValue,
  createPptxEditor,
  valueSlotOptions,
  type ChromeControl,
  type ChromeSlotId,
  type PptxEditorInstance,
} from '../index.ts';

// Without the `effects` drawing renderer (`@pptx-editor.dev/pro/drawing`), the chrome does not
// offer glow, inner shadow, reflection, or soft edge. The registry, the commands, and save keep
// every effect.

/** A renderer that paints nothing: only its presence matters here. */
const STUB: DrawingEffectsRenderer = {
  body: () => null,
  glow: () => null,
  reflection: () => null,
};

// Core installs no renderer, so each test returns to that. The late-install warning is the
// subject of its own test below; other files in this process may have painted already.
let warn: ReturnType<typeof spyOn>;
beforeAll(() => (warn = spyOn(console, 'warn').mockImplementation(() => {})));
afterAll(() => warn.mockRestore());
afterEach(() => installDrawingRenderers({ effects: undefined, chart: undefined }));

const control = (slot: ChromeSlotId): ChromeControl => chromeControlForSlot(slot)!;
const values = (c: ChromeControl): string[] => (c.options ?? []).map((o) => o.value);

describe('drawingCapabilities', () => {
  test('follows the installed effects renderer', () => {
    installDrawingRenderers({ effects: undefined });
    expect(drawingCapabilities()).toEqual({ effects: false, chart: false });
    installDrawingRenderers({ effects: STUB });
    expect(drawingCapabilities()).toEqual({ effects: true, chart: false });
    installDrawingRenderers({ effects: undefined });
    expect(drawingCapabilities()).toEqual({ effects: false, chart: false });
  });

  test('follows the installed chart renderer, with one frozen answer per combination', () => {
    const none = drawingCapabilities();
    installDrawingRenderers({ chart: () => null });
    const chart = drawingCapabilities();
    expect(chart).toEqual({ effects: false, chart: true });
    expect(Object.isFrozen(chart)).toBe(true);
    installDrawingRenderers({ effects: STUB });
    expect(drawingCapabilities()).toEqual({ effects: true, chart: true });
    installDrawingRenderers({ effects: undefined });
    expect(drawingCapabilities()).toBe(chart);
    installDrawingRenderers({ chart: undefined });
    expect(drawingCapabilities()).toBe(none);
  });

  test('an install raises the version and calls each subscriber', () => {
    const calls: number[] = [];
    const before = drawingRenderersVersion();
    const stop = subscribeDrawingRenderers(() => calls.push(drawingRenderersVersion()));
    const failing = subscribeDrawingRenderers(() => {
      throw new Error('boom');
    });
    installDrawingRenderers({ effects: STUB });
    installDrawingRenderers({ effects: undefined });
    expect(calls).toEqual([before + 1, before + 2]);
    stop();
    failing();
    installDrawingRenderers({ effects: STUB });
    expect(calls).toHaveLength(2);
  });

  test('a query is not a use: an install after it logs no warning', () => {
    // A fresh process, because the warning fires once per process after the first paint.
    const url = new URL('../../layout/drawing-renderers.ts', import.meta.url).href;
    const script = `
      const m = await import(${JSON.stringify(url)});
      let warned = 0;
      console.warn = () => { warned += 1; };
      m.drawingCapabilities();
      m.drawingRenderersVersion();
      m.subscribeDrawingRenderers(() => {});
      m.installDrawingRenderers({ effects: { body: () => null, glow: () => null, reflection: () => null } });
      const quiet = warned;
      m.drawingRenderers();
      m.installDrawingRenderers({ effects: { body: () => null, glow: () => null, reflection: () => null } });
      console.log(JSON.stringify({ quiet, after: warned }));
    `;
    const run = Bun.spawnSync([process.execPath, '-e', script]);
    expect(run.exitCode).toBe(0);
    expect(JSON.parse(run.stdout.toString())).toEqual({ quiet: 0, after: 1 });
  });
});

describe('chromeOfferedControl', () => {
  test('Shape Effects offers only the outer shadows without the renderer', () => {
    installDrawingRenderers({ effects: undefined });
    const full = control('shape.effects');
    const shown = chromeOfferedControl('shape.effects', full);
    expect(shown.optionGroups?.map((g) => g.id)).toEqual(['shadow']);
    expect(values(shown)).toContain('shadow.none');
    expect(values(shown)).toContain('shadow.outerBottomRight');
    expect(values(shown).some((v) => v.startsWith('shadow.inner'))).toBe(false);
    expect(values(shown).every((v) => v.startsWith('shadow.'))).toBe(true);
    // The same object each time, so a memoized adapter part keeps its rows.
    expect(chromeOfferedControl('shape.effects', full)).toBe(shown);
  });

  test('useEditorValueCommand options follow the same rule', () => {
    expect(valueSlotOptions('shape.effects')).toContain('shadow.outerTop');
    expect(valueSlotOptions('shape.effects').some((v) => v.startsWith('glow.'))).toBe(false);
    installDrawingRenderers({ effects: STUB });
    expect(valueSlotOptions('shape.effects')).toContain('glow.8.accent2');
  });

  test('Picture Styles offers only the styles whose effects paint without the renderer', () => {
    installDrawingRenderers({ effects: undefined });
    const shown = chromeOfferedControl('image.styles', control('image.styles'));
    expect(values(shown)).toEqual(['simpleFrameWhite', 'simpleFrameBlack']);
  });

  test('with the renderer, every slot shows its registry control', () => {
    installDrawingRenderers({ effects: STUB });
    for (const slot of ['shape.effects', 'image.styles'] as const)
      expect(chromeOfferedControl(slot, control(slot))).toBe(control(slot));
    installDrawingRenderers({ effects: undefined });
    for (const slot of ['shape.effects', 'image.styles'] as const)
      expect(chromeOfferedControl(slot, control(slot))).not.toBe(control(slot));
  });

  test('other slots, the registry, and the commands do not change', () => {
    installDrawingRenderers({ effects: undefined });
    for (const slot of ['shape.styles', 'paragraph.align', 'image.color'] as const)
      expect(chromeOfferedControl(slot, control(slot))).toBe(control(slot));
    const effects = CHROME_GROUPS.find((g) => g.id === 'shape')!.controls.find(
      (c) => c.id === 'effects'
    )!;
    expect(effects.optionGroups?.map((g) => g.id)).toEqual([
      'shadow',
      'reflection',
      'glow',
      'softEdge',
      'bevel',
      'rotation3d',
    ]);
    expect(commandForSlotValue('shape.effects', 'glow.8')).not.toBeNull();
    expect(commandForSlotValue('image.styles', 'softEdgeOval')).not.toBeNull();
  });

  test('the rule names each Pro effect option', () => {
    for (const v of [
      'shadow.innerTop',
      'reflection.none',
      'glow.8.accent2',
      'softEdge.5',
      'bevel.circle',
      'rotation3d.none',
    ])
      expect(chromeOptionNeedsEffects('shape.effects', v)).toBe(true);
    for (const v of ['shadow.none', 'shadow.outerTop', 'shadow.toString', 'nope'])
      expect(chromeOptionNeedsEffects('shape.effects', v)).toBe(false);
    for (const v of ['reflectedRoundedRectangle', 'softEdgeRectangle', 'softEdgeOval'])
      expect(chromeOptionNeedsEffects('image.styles', v)).toBe(true);
    for (const v of ['simpleFrameWhite', 'simpleFrameBlack', 'toString'])
      expect(chromeOptionNeedsEffects('image.styles', v)).toBe(false);
  });
});

describe('effects without the renderer', () => {
  test('save keeps glow, inner shadow, reflection, and soft edge byte for byte', async () => {
    installDrawingRenderers({ effects: undefined });
    const effects =
      '<a:effectLst><a:glow rad="63500"><a:srgbClr val="FF0000"/></a:glow>' +
      '<a:innerShdw blurRad="63500" dist="50800" dir="5400000"><a:srgbClr val="000000"/></a:innerShdw>' +
      '<a:reflection blurRad="6350" stA="50000" endPos="55000" dist="5000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>' +
      '<a:softEdge rad="12700"/></a:effectLst>';
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            shape({ id: 2, x: 0, y: 0, cx: 914_400, cy: 914_400, fill: solid('4472C4'), effects }) +
            shape({ id: 3, x: 2_000_000, y: 0, cx: 914_400, cy: 914_400, fill: solid('ED7D31') }),
        },
      ],
    });
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(bytes)).ok).toBe(true);
    // An edit on the same slide, so the slide part is written again.
    const other = editor.session!.layout(0)!.items[1]!.element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [other] });
    expect(
      editor.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'outerTop' } }).ok
    ).toBe(true);
    const slide = strFromU8(unzipSync(editor.save()!)['ppt/slides/slide1.xml']!);
    expect(slide).toContain('<a:outerShdw');
    expect(slide).toContain(effects);
  });
});

// Removing an effect needs no renderer. Without it, the `none` row of an effect stays while the
// selection has that effect, so a community build can remove an effect from the file.

const IN = 914_400;
const GLOW = '<a:glow rad="63500"><a:srgbClr val="FF0000"/></a:glow>';
const INNER =
  '<a:innerShdw blurRad="63500" dist="50800" dir="5400000"><a:srgbClr val="000000"/></a:innerShdw>';
const REFLECTION =
  '<a:reflection blurRad="6350" stA="50000" endPos="55000" dist="5000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>';
const SOFT = '<a:softEdge rad="12700"/>';
/** An unknown extension inside the effect list: save keeps it as it is. */
const UNKNOWN = '<a:extLst><a:ext uri="{00000000-0000-0000-0000-000000000001}"/></a:extLst>';
const list = (effects: string): string => `<a:effectLst>${effects}</a:effectLst>`;
/** An own empty list stays: it says "no effects" over an inherited one. */
const EMPTY = '<a:effectLst/>';

const box = (id: number, x: number, effects?: string): string =>
  shape({
    id,
    x,
    y: 0,
    cx: IN / 2,
    cy: IN / 2,
    fill: solid('4472C4'),
    ...(effects ? { effects: `<a:effectLst>${effects}</a:effectLst>` } : {}),
  });

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const nodeIds = (editor: PptxEditorInstance): string[] =>
  editor.session!.layout(0)!.items.map((i) => i.element.nodeId);

/** Shapes: glow and soft edge, plain, reflection, inner shadow, and a group with a glow inside. */
async function effectsDeck(): Promise<{ editor: PptxEditorInstance; ids: string[] }> {
  const group =
    '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="2743200"/><a:ext cx="914400" cy="457200"/><a:chOff x="0" y="2743200"/><a:chExt cx="914400" cy="457200"/></a:xfrm></p:grpSpPr>' +
    box(11, 0, GLOW) +
    '</p:grpSp>';
  const editor = await open(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            box(2, 0, GLOW + SOFT) +
            box(3, IN) +
            box(4, 2 * IN, REFLECTION) +
            box(5, 3 * IN, INNER) +
            group,
        },
      ],
    })
  );
  return { editor, ids: nodeIds(editor) };
}

const slideXml = (editor: PptxEditorInstance): string =>
  strFromU8(unzipSync(editor.save()!)['ppt/slides/slide1.xml']!);

const offeredEffects = (editor: PptxEditorInstance | null): ChromeControl =>
  chromeOfferedControl('shape.effects', control('shape.effects'), editor);

/** The Pro rows the menu offers: the options that a build without the renderer hides. */
const proRows = (editor: PptxEditorInstance): string[] =>
  values(offeredEffects(editor)).filter((v) => chromeOptionNeedsEffects('shape.effects', v));

const run = (editor: PptxEditorInstance, value: string): boolean =>
  editor.exec(commandForSlotValue('shape.effects', value)!).ok;

describe('removing an effect without the renderer', () => {
  test('a selection offers the No row of each effect it has, and no preset', async () => {
    installDrawingRenderers({ effects: undefined });
    const { editor, ids } = await effectsDeck();
    const [glowSoft, plain, reflected] = ids;
    editor.setSelection({ kind: 'shapes', nodeIds: [plain!] });
    expect(proRows(editor)).toEqual([]);
    expect(offeredEffects(editor)).toBe(offeredEffects(null));
    editor.setSelection({ kind: 'shapes', nodeIds: [glowSoft!] });
    expect(proRows(editor)).toEqual(['glow.none', 'softEdge.none']);
    const shown = offeredEffects(editor);
    expect(shown.optionGroups?.map((g) => g.id)).toEqual(['shadow', 'glow', 'softEdge']);
    // The same choices give the same object, so a memoized adapter part keeps its rows.
    expect(offeredEffects(editor)).toBe(shown);
    editor.setSelection({ kind: 'shapes', nodeIds: [reflected!] });
    expect(proRows(editor)).toEqual(['reflection.none']);
    // Picture Styles only adds effects: it offers no Pro style for any selection.
    expect(values(chromeOfferedControl('image.styles', control('image.styles'), editor))).toEqual([
      'simpleFrameWhite',
      'simpleFrameBlack',
    ]);
  });

  test('a mixed multi-selection offers the No row of each effect that a shape has', async () => {
    installDrawingRenderers({ effects: undefined });
    const { editor, ids } = await effectsDeck();
    const [glowSoft, plain, reflected] = ids;
    editor.setSelection({ kind: 'shapes', nodeIds: [plain!, reflected!] });
    expect(proRows(editor)).toEqual(['reflection.none']);
    editor.setSelection({ kind: 'shapes', nodeIds: [glowSoft!, reflected!] });
    expect(proRows(editor)).toEqual(['reflection.none', 'glow.none', 'softEdge.none']);
    expect(valueSlotOptions('shape.effects', editor)).toContain('reflection.none');
    expect(valueSlotOptions('shape.effects')).not.toContain('reflection.none');
    // No Reflection removes the reflection, and nothing else.
    const before = slideXml(editor);
    expect(run(editor, 'reflection.none')).toBe(true);
    expect(slideXml(editor)).toBe(before.replace(list(REFLECTION), EMPTY));
    expect(proRows(editor)).toEqual(['glow.none', 'softEdge.none']);
  });

  test('a group carries no effect: it offers no Pro row, and its children keep theirs', async () => {
    installDrawingRenderers({ effects: undefined });
    const { editor, ids } = await effectsDeck();
    const [, , reflected, , group] = ids;
    editor.setSelection({ kind: 'shapes', nodeIds: [group!] });
    expect(proRows(editor)).toEqual([]);
    // A group with a shape: the rows follow the shape, and the group's child keeps its glow.
    editor.setSelection({ kind: 'shapes', nodeIds: [group!, reflected!] });
    expect(proRows(editor)).toEqual(['reflection.none']);
    expect(run(editor, 'reflection.none')).toBe(true);
    expect(slideXml(editor)).toContain(GLOW);
  });

  test('No Shadow is always offered, and it removes an inner shadow', async () => {
    installDrawingRenderers({ effects: undefined });
    const { editor, ids } = await effectsDeck();
    editor.setSelection({ kind: 'shapes', nodeIds: [ids[3]!] });
    expect(editor.snapshot().geometry?.effects?.shadow).not.toBe('none');
    const shown = values(offeredEffects(editor));
    expect(shown).toContain('shadow.none');
    expect(shown.some((v) => v.startsWith('shadow.inner'))).toBe(false);
    const before = slideXml(editor);
    expect(run(editor, 'shadow.none')).toBe(true);
    expect(slideXml(editor)).toBe(before.replace(list(INNER), EMPTY));
    expect(editor.snapshot().geometry?.effects?.shadow).toBe('none');
  });

  test('No Glow saves without the glow, keeps the rest byte for byte, and undoes', async () => {
    installDrawingRenderers({ effects: undefined });
    const editor = await open(
      buildPptx({
        slides: [{ layout: 'blank', shapes: box(2, 0, GLOW + SOFT + UNKNOWN) + box(3, IN) }],
      })
    );
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeIds(editor)[0]!] });
    const original = slideXml(editor);
    expect(proRows(editor)).toEqual(['glow.none', 'softEdge.none']);
    expect(run(editor, 'glow.none')).toBe(true);
    const removed = slideXml(editor);
    expect(removed).toBe(original.replace(GLOW, ''));
    expect(removed).toContain(`<a:effectLst>${SOFT}${UNKNOWN}</a:effectLst>`);
    // The menu follows the edit: No Glow goes, and No Soft Edges stays.
    expect(proRows(editor)).toEqual(['softEdge.none']);
    // The saved deck reopens with the same slide.
    const reopened = await open(editor.save()!);
    expect(slideXml(reopened)).toBe(removed);
    // Undo brings the glow and its row back; redo removes them again.
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(slideXml(editor)).toBe(original);
    expect(proRows(editor)).toEqual(['glow.none', 'softEdge.none']);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(slideXml(editor)).toBe(removed);
    expect(proRows(editor)).toEqual(['softEdge.none']);
    // With the soft edge removed too, the shape offers no Pro row.
    expect(run(editor, 'softEdge.none')).toBe(true);
    expect(proRows(editor)).toEqual([]);
    expect(slideXml(editor)).toContain(`<a:effectLst>${UNKNOWN}</a:effectLst>`);
  });

  test('text editing in a shape, a picture, and view mode follow the command', async () => {
    installDrawingRenderers({ effects: undefined });
    const { editor, ids } = await effectsDeck();
    // A caret in the shape's text: the rows follow that shape.
    editor.setSelection({
      kind: 'text',
      nodeId: ids[0]!,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 0 },
    });
    expect(proRows(editor)).toEqual(['glow.none', 'softEdge.none']);
    // View mode removes nothing, so it offers no Pro row.
    editor.setSelection({ kind: 'shapes', nodeIds: [ids[0]!] });
    editor.setMode('view');
    expect(proRows(editor)).toEqual([]);
    editor.setMode('edit');
    expect(proRows(editor)).toEqual(['glow.none', 'softEdge.none']);
    // A picture with a soft edge.
    const PNG = Uint8Array.from(
      atob(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
      ),
      (c) => c.charCodeAt(0)
    );
    const pictures = await open(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            rels: [
              [
                'rIdImg',
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                '../media/image1.png',
              ],
            ],
            shapes:
              '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
              '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
              `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${list(SOFT)}</p:spPr></p:pic>`,
          },
        ],
        media: [{ name: 'image1.png', bytes: PNG, contentType: 'image/png' }],
      })
    );
    pictures.setSelection({ kind: 'shapes', nodeIds: [nodeIds(pictures)[0]!] });
    expect(proRows(pictures)).toEqual(['softEdge.none']);
    expect(run(pictures, 'softEdge.none')).toBe(true);
    expect(slideXml(pictures)).not.toContain('<a:softEdge');
    expect(proRows(pictures)).toEqual([]);
  });

  test('an effect inherited from the layout placeholder offers its No row', async () => {
    installDrawingRenderers({ effects: undefined });
    const ph = (spPr: string): string =>
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
      `<p:spPr>${spPr}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hi</a:t></a:r></a:p></p:txBody></p:sp>`;
    const editor = await open(
      buildPptx({
        layouts: [
          {
            name: 'Glowing',
            shapes: ph(
              `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${list(GLOW)}`
            ),
          },
        ],
        slides: [{ layoutName: 'Glowing', shapes: ph('') }],
      })
    );
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeIds(editor)[0]!] });
    expect(proRows(editor)).toEqual(['glow.none']);
    expect(run(editor, 'glow.none')).toBe(true);
    // The slide shape gets its own empty list over the layout's glow; the layout keeps it.
    expect(slideXml(editor)).toContain(EMPTY);
    expect(strFromU8(unzipSync(editor.save()!)['ppt/slideLayouts/slideLayout3.xml']!)).toContain(
      GLOW
    );
    expect(proRows(editor)).toEqual([]);
  });

  test('with the renderer, every row shows for any selection', async () => {
    const { editor, ids } = await effectsDeck();
    editor.setSelection({ kind: 'shapes', nodeIds: [ids[1]!] });
    installDrawingRenderers({ effects: STUB });
    expect(offeredEffects(editor)).toBe(control('shape.effects'));
  });
});
