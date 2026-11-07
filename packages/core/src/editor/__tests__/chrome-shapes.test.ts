import { describe, expect, spyOn, test } from 'bun:test';
import { buildPptx, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import {
  chromeControlForSlot,
  chromeSlotIds,
  commandForSlotValue,
  formattingBarChromeGroups,
  createPptxEditor,
  listenPaneRequests,
  requestPane,
  resetDevWarnings,
  runToolbarCommand,
  slotOptionState,
  toolbarCommandState,
} from '../index.ts';
import { chromeOptionChecked, readableText, shapeStylePreview } from '../chrome-shapes.ts';
import type { ShapeEffectsValue } from '../../contracts/index.ts';

const IN = 914_400;

describe('Shape Format slots', () => {
  test('each value compiles to its command; a bad value is refused', () => {
    expect(commandForSlotValue('shape.styles', 'intenseEffect.accent3')).toEqual({
      type: 'setShapeStyle',
      style: { row: 'intenseEffect', color: 'accent3' },
    });
    expect(commandForSlotValue('shape.styles', { row: 'coloredOutline', color: 'dk1' })).toEqual({
      type: 'setShapeStyle',
      style: { row: 'coloredOutline', color: 'dk1' },
    });
    expect(commandForSlotValue('shape.styles', 'loud.accent1')).toBeNull();
    expect(commandForSlotValue('shape.styles', 'coloredFill.accent1.junk')).toBeNull();
    expect(commandForSlotValue('shape.effects', 'shadow.outerTop')).toEqual({
      type: 'setShapeEffect',
      effect: { kind: 'shadow', preset: 'outerTop' },
    });
    expect(commandForSlotValue('shape.effects', 'shadow.none')).toEqual({
      type: 'setShapeEffect',
      effect: { kind: 'shadow', preset: null },
    });
    expect(commandForSlotValue('shape.effects', 'reflection.full8pt')).toEqual({
      type: 'setShapeEffect',
      effect: { kind: 'reflection', preset: 'full8pt' },
    });
    expect(commandForSlotValue('shape.effects', 'glow.11')).toEqual({
      type: 'setShapeEffect',
      effect: { kind: 'glow', size: 11 },
    });
    expect(commandForSlotValue('shape.effects', 'softEdge.2.5')).toEqual({
      type: 'setShapeEffect',
      effect: { kind: 'softEdge', size: 2.5 },
    });
    for (const bad of ['shadow.drop', 'softEdge.-2', 'glow.big', 'blur.3', 'outerTop', 7])
      expect(commandForSlotValue('shape.effects', bad)).toBeNull();
    expect(commandForSlotValue('arrange.rotate', 'left90')).toEqual({
      type: 'rotateObjects',
      degrees: -90,
    });
    expect(commandForSlotValue('arrange.rotate', 'flipHorizontal')).toEqual({
      type: 'flipObjects',
      axis: 'horizontal',
    });
    expect(commandForSlotValue('arrange.align', 'distributeVertical')).toEqual({
      type: 'distributeObjects',
      axis: 'vertical',
    });
  });

  test('the slots are public but not in the default formatting bar', () => {
    const bar = formattingBarChromeGroups().flatMap((g) =>
      g.controls.map((c) => `${g.id}.${c.id}`)
    );
    for (const slot of [
      'shape.styles',
      'shape.effects',
      'shape.properties',
      'arrange.rotate',
      'arrange.selectionPane',
    ] as const) {
      expect(chromeSlotIds()).toContain(slot);
      expect(bar).not.toContain(slot);
    }
    expect(bar).toContain('shape.fill');
  });

  test('every Shape Effects option is in one of its groups', () => {
    const control = chromeControlForSlot('shape.effects')!;
    const groups = new Set(control.optionGroups!.map((g) => g.id));
    expect([...groups]).toEqual(['shadow', 'reflection', 'glow', 'softEdge']);
    for (const o of control.options!) {
      expect(groups.has(o.group!)).toBe(true);
      expect(o.value.startsWith(`${o.group}.`)).toBe(true);
      expect(commandForSlotValue('shape.effects', o.value)).not.toBeNull();
    }
    expect(chromeOptionChecked('shadow.none glow.8', 'glow.8')).toBe(true);
    expect(chromeOptionChecked('shadow.none glow.8', 'glow.5')).toBe(false);
    expect(chromeOptionChecked('left', 'left')).toBe(true);
    expect(chromeOptionChecked(undefined, 'left')).toBe(false);
  });

  test('state follows the selection', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, fill: solid('4472C4') }),
        },
      ],
    });
    expect((await editor.load(bytes)).ok).toBe(true);
    expect(toolbarCommandState(editor, 'shape.effects').enabled).toBe(false);
    const id = editor.session!.layout(0)!.items[0].element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [id] });
    expect(toolbarCommandState(editor, 'shape.effects')).toMatchObject({
      enabled: true,
      value: 'shadow.none reflection.none glow.none softEdge.none',
    });
    editor.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 8 } });
    expect(toolbarCommandState(editor, 'shape.effects').value).toBe(
      'shadow.none reflection.none glow.8.accent1 softEdge.none'
    );
    expect(toolbarCommandState(editor, 'arrange.rotate').enabled).toBe(true);
    // Off while no side panel or host listens for pane requests; on while one does.
    expect(toolbarCommandState(editor, 'arrange.selectionPane')).toMatchObject({
      enabled: false,
      disabledReason: 'disabled.noSidePanel',
    });
    const stop = listenPaneRequests(editor, () => {});
    expect(toolbarCommandState(editor, 'arrange.selectionPane')).toMatchObject({
      enabled: true,
      disabledReason: null,
    });
    stop();
    expect(editor.exec(commandForSlotValue('arrange.rotate', 'right90')!).ok).toBe(true);
    expect(editor.snapshot().geometry?.rot).toBe(90);
  });
});

describe('Shape Format slot values', () => {
  test('glow values take an optional accent; other effects refuse one', () => {
    expect(commandForSlotValue('shape.effects', 'glow.8.accent2')).toEqual({
      type: 'setShapeEffect',
      effect: { kind: 'glow', size: 8, color: { theme: 'accent2' } },
    });
    expect(commandForSlotValue('shape.effects', 'glow.2.5')).toEqual({
      type: 'setShapeEffect',
      effect: { kind: 'glow', size: 2.5 },
    });
    for (const bad of ['glow.8.accent9', 'glow.none.accent1', 'softEdge.5.accent1', 'glow.0'])
      expect(commandForSlotValue('shape.effects', bad)).toBeNull();
  });

  test('a pane slot asks for its side panel tab; a bad value warns once', () => {
    const editor = createPptxEditor();
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      // Nobody listens: the slot says that a side panel must show the pane.
      expect(runToolbarCommand(editor, 'arrange.selectionPane')).toEqual({
        ok: false,
        reason: 'disabled.noSidePanel',
      });
      const asked: string[] = [];
      const stop = listenPaneRequests(editor, (pane) => asked.push(pane));
      expect(runToolbarCommand(editor, 'arrange.selectionPane')).toEqual({ ok: true });
      expect(runToolbarCommand(editor, 'shape.properties')).toEqual({ ok: true });
      expect(asked).toEqual(['selection', 'format']);
      stop();
      expect(requestPane(editor, 'format')).toBe(false);
      expect(asked).toEqual(['selection', 'format']);
      // A listener that throws does not stop the others.
      const heard: string[] = [];
      const offA = listenPaneRequests(editor, () => {
        throw new Error('host');
      });
      const offB = listenPaneRequests(editor, (pane) => heard.push(pane));
      expect(() => requestPane(editor, 'selection')).toThrow('host');
      expect(heard).toEqual(['selection']);
      offA();
      offB();
      // @ts-expect-error A typo in a typed slot value does not compile.
      runToolbarCommand(null, 'shape.effects', 'shadow.typo');
      // A host that bypasses the value type still gets a refusal and one warning.
      const bogus = 'shadow.bogus' as ShapeEffectsValue;
      expect(runToolbarCommand(editor, 'shape.effects', bogus)).toEqual({
        ok: false,
        reason: 'disabled.invalidValue',
      });
      runToolbarCommand(editor, 'shape.effects', bogus);
      expect(warn).toHaveBeenCalledTimes(1);
      expect(String(warn.mock.calls[0][0])).toContain(
        '"shadow.bogus" is not a choice of "shape.effects"'
      );
      // A cyclic value (an event a host passed as the value) refuses; it does not throw.
      const cyclic: { self?: unknown } = {};
      cyclic.self = cyclic;
      expect(runToolbarCommand(editor, 'arrange.rotate', cyclic as never)).toEqual({
        ok: false,
        reason: 'disabled.invalidValue',
      });
    } finally {
      warn.mockRestore();
      resetDevWarnings();
    }
  });
});

describe('Quick Styles preview', () => {
  const row = (map: Record<string, string>) => [
    ['bg1', 'tx1', 'bg2', 'tx2', 'accent1'].map((slot) => ({
      slot: slot as 'bg1',
      shift: 0,
      scheme: map[slot] ?? slot,
      color: { lt1: '#FFFFFF', dk1: '#000000', accent1: '#4472C4' }[map[slot] ?? slot] ?? '#808080',
    })),
  ];
  test('dk1 and lt1 name the theme colors, not the mapped slots', () => {
    const light = row({ bg1: 'lt1', tx1: 'dk1' });
    const dark = row({ bg1: 'dk1', tx1: 'lt1' });
    // Colored Outline: lt1 fill, dk1 text, on either color map.
    for (const colors of [light, dark]) {
      const p = shapeStylePreview({ row: 'coloredOutline', color: 'accent1' }, colors);
      expect(p.fill).toBe('#FFFFFF');
      expect(p.text).toBe('#000000');
    }
    expect(shapeStylePreview({ row: 'coloredFill', color: 'dk1' }, dark).fill).toBe('#000000');
    // a:shade val="15000" in linear RGB, as the slide draws it (not 15 % of each sRGB byte).
    expect(shapeStylePreview({ row: 'coloredFill', color: 'accent1' }, light).line).toBe('#172C51');
  });

  test('the preview text stays readable on its fill (Subtle Effect: Dark 1)', () => {
    const light = row({ bg1: 'lt1', tx1: 'dk1' });
    const dark = row({ bg1: 'dk1', tx1: 'lt1' });
    for (const colors of [light, dark])
      for (const row of ['subtleEffect', 'coloredFill', 'intenseEffect'] as const) {
        const p = shapeStylePreview({ row, color: 'dk1' }, colors);
        expect(p.text === p.fill).toBe(false);
      }
    // A readable pair keeps the style's own font color.
    expect(readableText('#000000', '#FFFFFF')).toBe('#000000');
    expect(readableText('#1B1A19', '#1B1A19')).toBe('#FFFFFF');
    expect(readableText('#F0F0F0', '#FFFFFF')).toBe('#000000');
  });
});

describe('Toolbar option state (slotOptionState)', () => {
  test('only a refusal disables an option; an option that changes nothing stays enabled', async () => {
    expect(slotOptionState(null, 'arrange.align', 'distributeHorizontal')).toEqual({
      enabled: false,
      reason: 'disabled.notWired',
    });
    const editor = createPptxEditor();
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: [0, 2, 4]
            .map((n) =>
              shape({ id: 2 + n, x: IN * (1 + n), y: IN, cx: IN, cy: IN, fill: solid('4472C4') })
            )
            .join(''),
        },
      ],
    });
    expect((await editor.load(bytes)).ok).toBe(true);
    expect(slotOptionState(editor, 'arrange.align', 'distributeHorizontal')).toEqual({
      enabled: false,
      reason: 'disabled.noSelection',
    });
    const ids = editor.session!.layout(0)!.items.map((i) => i.element.nodeId);
    editor.setSelection({ kind: 'shapes', nodeIds: ids.slice(0, 2) });
    expect(slotOptionState(editor, 'arrange.align', 'distributeHorizontal')).toEqual({
      enabled: false,
      reason: 'disabled.distributeNeedsThree',
    });
    // Three boxes that are already evenly spaced: Distribute changes nothing, and stays enabled.
    editor.setSelection({ kind: 'shapes', nodeIds: ids });
    expect(editor.can({ type: 'distributeObjects', axis: 'horizontal' }).ok).toBe(false);
    expect(slotOptionState(editor, 'arrange.align', 'distributeHorizontal')).toEqual({
      enabled: true,
      reason: null,
    });
  });
});
