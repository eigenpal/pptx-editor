import { afterAll, afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { buildPptx, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import {
  drawingCapabilities,
  installDrawingRenderers,
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
afterEach(() => installDrawingRenderers({ effects: undefined }));

const control = (slot: ChromeSlotId): ChromeControl => chromeControlForSlot(slot)!;
const values = (c: ChromeControl): string[] => (c.options ?? []).map((o) => o.value);

describe('drawingCapabilities', () => {
  test('follows the installed effects renderer', () => {
    installDrawingRenderers({ effects: undefined });
    expect(drawingCapabilities()).toEqual({ effects: false });
    installDrawingRenderers({ effects: STUB });
    expect(drawingCapabilities()).toEqual({ effects: true });
    installDrawingRenderers({ effects: undefined });
    expect(drawingCapabilities()).toEqual({ effects: false });
  });

  test('a query is not a use: an install after it logs no warning', () => {
    // A fresh process, because the warning fires once per process after the first paint.
    const url = new URL('../../layout/drawing-renderers.ts', import.meta.url).href;
    const script = `
      const m = await import(${JSON.stringify(url)});
      let warned = 0;
      console.warn = () => { warned += 1; };
      m.drawingCapabilities();
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
    ]);
    expect(commandForSlotValue('shape.effects', 'glow.8')).not.toBeNull();
    expect(commandForSlotValue('image.styles', 'softEdgeOval')).not.toBeNull();
  });

  test('the rule names each Pro effect option', () => {
    for (const v of ['shadow.innerTop', 'reflection.none', 'glow.8.accent2', 'softEdge.5'])
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
