// The Vue insert, shape, table, and slide controls render the same markup, ARIA, and keys as
// their React twins (packages/react/src/ShapeControls.tsx, ShapeFormatControls.tsx,
// TableControls.tsx, SlideControls.tsx, and the gallery rows of menu-rows.tsx). Each case renders
// both adapters on the same deck and compares a canonical form of the DOM (attributes sorted,
// generated ids replaced by their order of use), then runs the Vue part's behavior.

import { afterEach, describe, expect, test } from 'bun:test';
import { createElement, type ReactElement } from 'react';
import { act, cleanup, render } from '@testing-library/react';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import {
  chromeControlForSlot,
  createPptxEditor,
  type ChromeMenuEntry,
  type ChromeSlotId,
  type GradientFill,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../../react/src/context.ts';
import * as RS from '../../react/src/ShapeControls.tsx';
import * as RF from '../../react/src/ShapeFormatControls.tsx';
import * as RT from '../../react/src/TableControls.tsx';
import * as RL from '../../react/src/SlideControls.tsx';
import * as RP from '../../react/src/PictureControls.tsx';
import { MenuEntries as RMenuEntries } from '../../react/src/menu-rows.tsx';
import { providePptxEditor } from '../src/composables.ts';
import * as VS from '../src/ShapeControls.tsx';
import * as VF from '../src/ShapeFormatControls.tsx';
import * as VT from '../src/TableControls.tsx';
import * as VL from '../src/SlideControls.tsx';
import * as VP from '../src/PictureControls.tsx';
import { MenuEntries } from '../src/menu-rows.tsx';
import {
  FillControl,
  GradientDialog,
  LayoutGallery,
  MenuTableGrid,
  ShapeGrid,
  ShapeStyleGallery,
} from '../src/index.ts';
import { canonicalDom } from './canonical-dom.ts';

const IN = 914_400;

let app: App | null = null;
let host: HTMLElement | null = null;
function unmountVue(): void {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
}
afterEach(() => {
  cleanup();
  unmountVue();
  document.body.replaceChildren();
});

async function mountVue(
  renderFn: () => unknown,
  editor?: PptxEditorInstance | null
): Promise<HTMLElement> {
  unmountVue();
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        if (editor !== undefined) providePptxEditor(shallowRef(editor));
        return renderFn;
      },
    }) as Component
  );
  app.mount(host);
  await nextTick();
  return host;
}

function mountReact(element: ReactElement, editor?: PptxEditorInstance | null): HTMLElement {
  const wrapped =
    editor === undefined
      ? element
      : createElement(PptxEditorContext.Provider, { value: editor }, element);
  return render(wrapped).container;
}

/** The canonical form of each root's children, with ids shared across the roots. */
const canon = (...roots: Element[]): string =>
  roots.map((r) => canonicalDom(r, { dropSelectValue: true, splitIdReferences: true })).join('|');

/** The host plus every popover panel in the document. */
const shown = (root: HTMLElement): string =>
  canon(root, ...[...document.querySelectorAll('body > .pptx-menu')].map(wrap));
const wrap = (el: Element): Element => {
  const box = document.createElement('div');
  box.append(el.cloneNode(true));
  return box;
};

async function same(
  react: ReactElement,
  vue: () => unknown,
  editor?: PptxEditorInstance,
  open?: string
): Promise<string> {
  unmountVue();
  document.body.replaceChildren();
  const r = mountReact(react, editor);
  if (open) act(() => (r.querySelector(open) as HTMLElement).click());
  const rOut = shown(r);
  cleanup();
  document.body.replaceChildren();
  const v = await mountVue(vue, editor);
  if (open) {
    (v.querySelector(open) as HTMLElement).click();
    await nextTick();
    await nextTick();
  }
  const vOut = shown(v);
  expect(vOut).toEqual(rOut);
  return vOut;
}

async function deck(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Hi')) }),
      },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const control = (slot: ChromeSlotId) => chromeControlForSlot(slot)!;

describe('shape parts', () => {
  test('ShapeIcon, ShapeGrid, and LineGlyph match React', async () => {
    await same(createElement(RS.ShapeIcon, { preset: 'rightArrow' }), () =>
      h(VS.ShapeIcon, { preset: 'rightArrow' })
    );
    await same(
      createElement(RS.ShapeGrid, { presets: ['rect', 'ellipse'], onPick: () => {} }),
      () => h(ShapeGrid, { presets: ['rect', 'ellipse'], onPick: () => {} })
    );
    for (const value of [
      'line',
      'elbow',
      'curved',
      'arrow',
      'triangle',
      'stealth',
      'oval',
      'diamond',
      'none',
    ])
      for (const start of [undefined, true, false])
        await same(
          createElement(RS.LineGlyph, { value, ...(start === undefined ? {} : { start }) }),
          () => h(VS.LineGlyph, { value, ...(start === undefined ? {} : { start }) })
        );
  });

  test('a ShapeGrid pick calls onPick with the preset', async () => {
    const picks: string[] = [];
    const el = await mountVue(() =>
      h(ShapeGrid, { presets: ['rect', 'ellipse'], onPick: (p: string) => picks.push(p) })
    );
    (el.querySelectorAll('button')[1] as HTMLElement).click();
    expect(picks).toEqual(['ellipse']);
  });

  test('the toolbar shape, line, fill, style, and effect controls match React, closed and open', async () => {
    const editor = await deck();
    editor.setSelection({ kind: 'shapes', nodeIds: [shapeId(editor)] });
    const cases: [
      ChromeSlotId,
      Component,
      (p: { slot: ChromeSlotId; control: never }) => ReactElement,
    ][] = [
      ['insert.shape', VS.ShapeGalleryControl, (p) => createElement(RS.ShapeGalleryControl, p)],
      ['insert.line', VS.LineMenuControl, (p) => createElement(RS.LineMenuControl, p)],
      [
        'shape.styles',
        VF.ShapeStyleGalleryControl,
        (p) => createElement(RF.ShapeStyleGalleryControl, p),
      ],
      ['shape.effects', VF.EffectMenuControl, (p) => createElement(RF.EffectMenuControl, p)],
    ];
    for (const [slot, vue, react] of cases) {
      const props = { slot, control: control(slot) as never };
      await same(react(props), () => h(vue, props), editor);
      const open = await same(react(props), () => h(vue, props), editor, `[data-slot="${slot}"]`);
      expect(open).toContain('role="menu"');
    }
    const fill = { slot: 'shape.fill' as ChromeSlotId, control: control('shape.fill') as never };
    await same(createElement(RS.FillControl, fill), () => h(FillControl, fill), editor);
    await same(
      createElement(RS.FillControl, fill),
      () => h(FillControl, fill),
      editor,
      '.pptx-toolbar__step'
    );
  });

  test('the Corrections grid matches React, closed and open; a cell sets both values', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const image = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
    // A 1 x 1 PNG: the control needs a selected picture, not its pixels.
    const png = Uint8Array.from(
      atob(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
      ),
      (c) => c.charCodeAt(0)
    );
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          rels: [['rIdImg', image, '../media/image1.png']],
          shapes:
            '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
            '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
            `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
        },
      ],
      media: [{ name: 'image1.png', bytes: png, contentType: 'image/png' }],
    });
    expect((await editor.load(bytes)).ok).toBe(true);
    editor.setSelection({ kind: 'shapes', nodeIds: [shapeId(editor)] });
    const slot = 'image.corrections' as ChromeSlotId;
    const props = { slot, control: control(slot) as never };
    await same(
      createElement(RP.CorrectionsControl, props),
      () => h(VP.CorrectionsControl, props),
      editor
    );
    const open = await same(
      createElement(RP.CorrectionsControl, props),
      () => h(VP.CorrectionsControl, props),
      editor,
      `[data-slot="${slot}"]`
    );
    expect(open).toContain('pptx-corrections-grid');
    (document.querySelector('[data-value="-40,40"]') as HTMLElement).click();
    await nextTick();
    expect(editor.snapshot().image).toMatchObject({ brightness: -0.4, contrast: 0.4 });
    expect(document.querySelector('.pptx-corrections-grid') === null).toBe(true);
  });

  test('a disabled shape gallery says why and does not open', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const props = { slot: 'insert.shape' as ChromeSlotId, control: control('insert.shape') };
    const out = await same(
      createElement(RS.ShapeGalleryControl, props),
      () => h(VS.ShapeGalleryControl, props),
      editor
    );
    expect(out).toContain('aria-disabled="true"');
  });

  test('a shape pick in the toolbar gallery inserts the shape', async () => {
    const editor = await deck();
    const before = editor.snapshot().revision;
    const el = await mountVue(
      () => h(VS.ShapeGalleryControl, { slot: 'insert.shape', control: control('insert.shape') }),
      editor
    );
    (el.querySelector('[data-slot="insert.shape"]') as HTMLElement).click();
    await nextTick();
    const item = document.querySelector('.pptx-shape-gallery__item[aria-label]') as HTMLElement;
    expect(item !== null).toBe(true);
    item.click();
    await nextTick();
    expect(editor.snapshot().revision).not.toBe(before);
    expect(document.querySelector('.pptx-shape-gallery') === null).toBe(true);
  });
});

function shapeId(editor: PptxEditorInstance): string {
  return editor.session!.layout(0)!.items[0]!.element.nodeId;
}

describe('ShapeStyleGallery', () => {
  test('matches React with a checked cell and a host class, and calls onPick', async () => {
    const editor = await deck();
    await same(
      createElement(RF.ShapeStyleGallery, {
        value: 'subtleEffect.accent3',
        onPick: () => {},
        className: 'host',
      }),
      () =>
        h(ShapeStyleGallery, { value: 'subtleEffect.accent3', class: 'host', onPick: () => {} }),
      editor
    );
    const picks: unknown[] = [];
    const el = await mountVue(
      () => h(ShapeStyleGallery, { onPick: (c: unknown) => picks.push(c) }),
      editor
    );
    (el.querySelector('[data-style]') as HTMLElement).click();
    expect(picks).toHaveLength(1);
  });
});

describe('GradientDialog', () => {
  const start: GradientFill = {
    kind: 'linear',
    angle: 90,
    rotWithShape: true,
    stops: [
      { pos: 1, color: '#000000' },
      { pos: 0, color: '#FFFFFF' },
    ],
  };

  test('matches React closed and open', async () => {
    const noop = (): void => {};
    await same(
      createElement(RS.GradientDialog, {
        open: false,
        initial: null,
        onClose: noop,
        onApply: noop,
      }),
      () => h(GradientDialog, { open: false, onClose: noop, onApply: noop })
    );
    await same(
      createElement(RS.GradientDialog, {
        open: true,
        initial: start,
        onClose: noop,
        onApply: noop,
      }),
      () => h(GradientDialog, { open: true, initial: start, onClose: noop, onApply: noop })
    );
  });

  test('adds and removes stops, and applies the edited gradient', async () => {
    const applied: GradientFill[] = [];
    let closed = 0;
    const el = await mountVue(() =>
      h(GradientDialog, {
        open: true,
        initial: start,
        onApply: (g: GradientFill) => applied.push(g),
        onClose: () => (closed += 1),
      })
    );
    const buttons = () => [...el.querySelectorAll<HTMLButtonElement>('.pptx-button')];
    expect(el.querySelectorAll('.pptx-gradient__handle')).toHaveLength(2);
    buttons()[0]!.click();
    await nextTick();
    expect(el.querySelectorAll('.pptx-gradient__handle')).toHaveLength(3);
    const handle = el.querySelectorAll<HTMLElement>('.pptx-gradient__handle')[1]!;
    handle.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
    await nextTick();
    buttons().at(-1)!.click();
    expect(applied).toHaveLength(1);
    expect(applied[0]!.stops.map((s) => s.pos)).toEqual([0, 0.51, 1]);
    buttons().at(-2)!.click();
    expect(closed).toBe(1);
  });
});

describe('MenuTableGrid', () => {
  test('matches React with a class and a label', async () => {
    const editor = await deck();
    await same(
      createElement(RT.MenuTableGrid, { className: 'host', label: 'Grid' }),
      () => h(MenuTableGrid, { class: 'host', label: 'Grid' }),
      editor
    );
  });

  test('a hover names the size, the arrows move the tab stop, and a pick runs the listener', async () => {
    const editor = await deck();
    const picks: [number, number][] = [];
    const el = await mountVue(
      () => h(MenuTableGrid, { onPick: (r: number, c: number) => picks.push([r, c]) }),
      editor
    );
    const cell = (id: string) => el.querySelector<HTMLElement>(`[data-cell="${id}"]`)!;
    cell('2x3').dispatchEvent(new MouseEvent('mouseenter'));
    await nextTick();
    expect(el.querySelector('.pptx-table-grid__label')!.textContent).toBe('3 × 2 table');
    expect(el.querySelectorAll('[data-active]')).toHaveLength(6);
    el.querySelector('[role="grid"]')!.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true })
    );
    await nextTick();
    expect(cell('2x1').getAttribute('tabindex')).toBe('0');
    cell('3x4').click();
    expect(picks).toEqual([[3, 4]]);
  });

  test('without a pick listener it inserts a table and emits inserted', async () => {
    const editor = await deck();
    const before = editor.snapshot().revision;
    let inserted = 0;
    const el = await mountVue(
      () => h(MenuTableGrid, { onInserted: () => (inserted += 1) }),
      editor
    );
    el.querySelector<HTMLElement>('[data-cell="2x2"]')!.click();
    expect(inserted).toBe(1);
    expect(editor.snapshot().revision).not.toBe(before);
  });

  test('the table toolbar controls match React, closed and open', async () => {
    const editor = await deck();
    for (const [slot, vue, react] of [
      ['table.insert', VT.TableGridControl, RT.TableGridControl],
      ['table.style', VT.TableStyleControl, RT.TableStyleControl],
    ] as const) {
      const props = { slot: slot as ChromeSlotId, control: control(slot) };
      await same(createElement(react, props), () => h(vue, props), editor);
    }
    const props = { slot: 'table.insert' as ChromeSlotId, control: control('table.insert') };
    await same(
      createElement(RT.TableGridControl, props),
      () => h(VT.TableGridControl, props),
      editor,
      '[data-slot="table.insert"]'
    );
  });
});

describe('LayoutGallery and the slide controls', () => {
  test('matches React with and without a current layout, and emits change', async () => {
    const editor = await deck();
    const layouts = editor.snapshot().layouts ?? [];
    expect(layouts.length).toBeGreaterThan(0);
    await same(
      createElement(RL.LayoutGallery, { onChange: () => {}, className: 'host' }),
      () => h(LayoutGallery, { class: 'host' }),
      editor
    );
    await same(
      createElement(RL.LayoutGallery, { value: layouts[0]!.id, onChange: () => {} }),
      () => h(LayoutGallery, { value: layouts[0]!.id }),
      editor
    );
    await same(
      createElement(RL.LayoutGallery, { onChange: () => {}, disabled: true, disabledReason: 'No' }),
      () => h(LayoutGallery, { disabled: true, disabledReason: 'No' }),
      editor
    );
    await same(
      createElement(RL.LayoutGallery, { onChange: () => {}, layouts: [] }),
      () => h(LayoutGallery, { layouts: [] }),
      editor
    );
    const picked: string[] = [];
    const el = await mountVue(
      () => h(LayoutGallery, { onChange: (id: string) => picked.push(id) }),
      editor
    );
    el.querySelector<HTMLElement>('.pptx-layout-gallery__item')!.click();
    expect(picked).toEqual([layouts[0]!.id]);
  });

  test('New Slide, Layout, and the section buttons match React, closed and open', async () => {
    const editor = await deck();
    for (const [slot, vue, react] of [
      ['slide.add', VL.NewSlideControl, RL.NewSlideControl],
      ['slide.layout', VL.LayoutControl, RL.LayoutControl],
      ['section.add', VL.SectionAddControl, RL.SectionAddControl],
      ['section.rename', VL.SectionRenameControl, RL.SectionRenameControl],
    ] as const) {
      const props = { slot: slot as ChromeSlotId, control: control(slot) };
      await same(createElement(react, props), () => h(vue, props), editor);
    }
    const add = { slot: 'slide.add' as ChromeSlotId, control: control('slide.add') };
    await same(
      createElement(RL.NewSlideControl, add),
      () => h(VL.NewSlideControl, add),
      editor,
      '.pptx-toolbar__caret-button'
    );
    const layout = { slot: 'slide.layout' as ChromeSlotId, control: control('slide.layout') };
    await same(
      createElement(RL.LayoutControl, layout),
      () => h(VL.LayoutControl, layout),
      editor,
      '[data-slot="slide.layout"]'
    );
  });

  test('a New Slide gallery pick adds a slide on that layout', async () => {
    const editor = await deck();
    const el = await mountVue(
      () => h(VL.NewSlideControl, { slot: 'slide.add', control: control('slide.add') }),
      editor
    );
    el.querySelector<HTMLElement>('.pptx-toolbar__caret-button')!.click();
    await nextTick();
    document.querySelector<HTMLElement>('.pptx-layout-gallery__item')!.click();
    await nextTick();
    expect(editor.snapshot().slides).toHaveLength(2);
    expect(document.querySelector('.pptx-layout-gallery') === null).toBe(true);
  });
});

describe('menu rows', () => {
  const entries: ChromeMenuEntry[] = [
    { kind: 'item', slot: 'insert.shape' },
    { kind: 'item', slot: 'insert.line' },
    { kind: 'item', slot: 'table.insert' },
    { kind: 'item', slot: 'slide.add' },
    { kind: 'item', slot: 'slide.layout' },
  ];

  test('Insert > Shape, Line, Table, New Slide, and Layout match React', async () => {
    const editor = await deck();
    const out = await same(
      createElement(RMenuEntries, { entries }),
      () => h(MenuEntries, { entries }),
      editor
    );
    expect(out).not.toContain('aria-disabled');
    expect(out).toContain('pptx-line-glyph');
  });

  test('each gallery row opens its gallery as React does', async () => {
    const editor = await deck();
    for (const entry of entries) {
      const one = [entry];
      const open = await same(
        createElement(RMenuEntries, { entries: one }),
        () => h(MenuEntries, { entries: one }),
        editor,
        '[aria-haspopup="menu"]'
      );
      expect(open).toContain('pptx-menu--submenu');
    }
  });

  test('Crop to shape cascades the shape grid as Insert > Shape does, as in React', async () => {
    const editor = await deck();
    const crop: ChromeMenuEntry[] = [{ kind: 'item', slot: 'image.cropToShape' }];
    const out = await same(
      createElement(RMenuEntries, { entries: crop }),
      () => h(MenuEntries, { entries: crop }),
      editor
    );
    // A submenu row, not the disabled "pick in the toolbar" row of a picker with no menu form.
    expect(out).toContain('aria-haspopup="menu"');
  });

  test('a New Slide row pick adds a slide on the picked layout', async () => {
    const editor = await deck();
    const el = await mountVue(
      () => h(MenuEntries, { entries: [{ kind: 'item', slot: 'slide.add' }] }),
      editor
    );
    el.querySelector<HTMLElement>('[aria-haspopup="menu"]')!.click();
    await nextTick();
    await nextTick();
    document.querySelector<HTMLElement>('.pptx-layout-gallery__item')!.click();
    expect(editor.snapshot().slides).toHaveLength(2);
  });
});
