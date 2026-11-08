// The slide master view (`setPresentationView` with `slideMaster`): the rail lists each master
// and its layouts, the canvas edits a layout or a master as a slide, and every slide that uses
// the layout shows the edit. Commands that act on slides are off, undo stays in the view, and
// closing the view goes back to the slide it opened from.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { serializeXmlPart, TEMPLATE_BASE, xmlPartOf } from '../../store/index.ts';
import { createTranslator } from '@pptx-editor.dev/i18n';
import {
  chromeCommandState,
  createPptxEditor,
  canvasMenuRows,
  deckSlideCount,
  CANVAS_CONTEXT_MENUS,
  masterBarHint,
  masterBarTitle,
  masterRailLabel,
  masterRailUsage,
  masterStepTarget,
  slideStatus,
} from '../index.ts';
import { startSlideshowController } from '../slideshow-screen.ts';

const IN = 914_400;

/** Two slides on a layout with a header text, and one slide on the blank layout. */
function deck(): Uint8Array {
  return buildPptx({
    layouts: [
      {
        name: 'Branded',
        shapes: shape({
          id: 9,
          name: 'Header',
          x: IN / 4,
          y: IN / 8,
          cx: 3 * IN,
          cy: IN / 2,
          text: para(run('Brand')),
        }),
      },
    ],
    slides: [
      {
        layoutName: 'Branded',
        shapes: shape({ id: 2, name: 'A', x: IN, y: 2 * IN, cx: IN, cy: IN }),
      },
      { layout: 'blank', shapes: '' },
      { layoutName: 'Branded', shapes: '' },
    ],
  });
}

async function opened() {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(deck())).ok).toBe(true);
  return e;
}

const layoutPart = (e: Awaited<ReturnType<typeof opened>>) =>
  e.snapshot().slides.find((s) => s.name === 'Branded')!.id;

describe('the slide master view', () => {
  test('opens on the layout of the current slide and lists masters with their layouts', async () => {
    const e = await opened();
    e.goToSlide(2);
    expect(e.exec({ type: 'setPresentationView', view: 'slideMaster' }).ok).toBe(true);
    const snap = e.snapshot();
    expect(snap.presentationView).toBe('slideMaster');
    // A master first, then its layouts, each with a template index and a name.
    expect(snap.slides[0]!.template).toBe('master');
    expect(snap.slides.slice(1).every((s) => s.template === 'layout')).toBe(true);
    expect(snap.slides.every((s) => s.index >= TEMPLATE_BASE)).toBe(true);
    const branded = snap.slides.find((s) => s.name === 'Branded')!;
    expect(snap.currentSlide).toBe(branded.index);
    expect(branded.masterId).toBe(snap.slides[0]!.id);
  });

  test('a layout shape is selectable and editable, and its slides show the edit', async () => {
    const e = await opened();
    e.exec({ type: 'setPresentationView', view: 'slideMaster' });
    const shapes = e.snapshot().shapes!;
    const header = shapes.find((s) => s.name === 'Header')!;
    e.setSelection({ kind: 'shapes', nodeIds: [header.nodeId] });
    expect(e.exec({ type: 'nudge', dx: IN, dy: 0 }).ok).toBe(true);
    const part = layoutPart(e);
    const xml = serializeXmlPart(xmlPartOf(e.session!.store.package, part)!);
    expect(xml).toContain(`<a:off x="${IN / 4 + IN}"`);
    // Back on the slides: slide 1 and slide 3 use the layout and show the moved header.
    expect(e.exec({ type: 'setPresentationView', view: 'normal' }).ok).toBe(true);
    for (const index of [0, 2]) {
      const item = e.session!.layout(index)!.items.find((i) => i.element.name === 'Header')!;
      expect(item.element.layer).toBe('layout');
      expect(item.element.xfrm.x).toBe(IN / 4 + IN);
    }
  });

  test('typing edits the layout text', async () => {
    const e = await opened();
    e.exec({ type: 'setPresentationView', view: 'slideMaster' });
    const header = e.snapshot().shapes!.find((s) => s.name === 'Header')!;
    e.setSelection({
      kind: 'text',
      nodeId: header.nodeId,
      anchor: { paragraph: 0, offset: 5 },
      focus: { paragraph: 0, offset: 5 },
    });
    for (const ch of ' Co') expect(e.input('text', ch).ok).toBe(true);
    const xml = serializeXmlPart(xmlPartOf(e.session!.store.package, layoutPart(e))!);
    expect(xml).toContain('Brand Co');
  });

  test('commands that act on slides are off', async () => {
    const e = await opened();
    e.exec({ type: 'setPresentationView', view: 'slideMaster' });
    for (const command of [
      { type: 'addSlide' },
      { type: 'deleteSlide' },
      { type: 'duplicateSlides' },
      { type: 'moveSlides', to: 0 },
      { type: 'setNotes', text: 'x' },
      { type: 'addSection', name: 'x' },
    ] as const) {
      const r = e.can(command as never);
      expect(r.ok).toBe(false);
      expect(r.ok ? '' : r.reason).toBe('disabled.slideMaster');
    }
    // Shape and text edits still run.
    expect(e.can({ type: 'insertTextBox' } as never).ok).toBe(true);
  });

  test('undo stays in the view, and closing goes back to the slide it opened from', async () => {
    const e = await opened();
    e.goToSlide(1);
    e.exec({ type: 'setPresentationView', view: 'slideMaster' });
    const at = e.snapshot().currentSlide;
    const header = e.snapshot().shapes?.find((s) => s.name === 'Header');
    // The blank layout opened: go to the branded one, as a click in the rail does.
    const branded = e.snapshot().slides.find((s) => s.name === 'Branded')!;
    expect(at).not.toBe(branded.index);
    expect(header).toBeUndefined();
    e.goToSlide(branded.index);
    const id = e.snapshot().shapes!.find((s) => s.name === 'Header')!.nodeId;
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    expect(e.exec({ type: 'nudge', dx: IN, dy: 0 }).ok).toBe(true);
    e.goToSlide(TEMPLATE_BASE);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    // The undo shows the edited layout with the shape selected again.
    expect(e.snapshot().currentSlide).toBe(branded.index);
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [id] });
    // A slide index does not leave the view.
    e.goToSlide(0);
    expect(e.snapshot().currentSlide).toBe(branded.index);
    expect(e.exec({ type: 'setPresentationView', view: 'normal' }).ok).toBe(true);
    expect(e.snapshot().currentSlide).toBe(1);
    expect(e.snapshot().slides).toHaveLength(3);
  });

  test('the background goes on the template on the canvas, never on all slides', async () => {
    const e = await opened();
    e.exec({ type: 'setPresentationView', view: 'slideMaster' });
    const branded = e.snapshot().slides.find((s) => s.name === 'Branded')!;
    e.goToSlide(branded.index);
    const fill = { kind: 'solid', color: '#336699' } as const;
    expect(e.exec({ type: 'setSlideBackground', background: fill, allSlides: true }).ok).toBe(true);
    const pkg = e.session!.store.package;
    expect(serializeXmlPart(xmlPartOf(pkg, branded.id)!)).toContain('<p:bg>');
    // The slides keep their own (inherited) backgrounds: no slide part gains a `p:bg`.
    for (const part of ['/ppt/slides/slide1.xml', '/ppt/slides/slide3.xml'])
      expect(serializeXmlPart(xmlPartOf(pkg, part)!)).not.toContain('<p:bg>');
  });

  test('Find, the show, and the slide status are off in the view', async () => {
    const e = await opened();
    e.exec({ type: 'setPresentationView', view: 'slideMaster' });
    for (const slot of ['find.find', 'find.replace', 'view.playFromStart'] as const) {
      const state = chromeCommandState(e, slot);
      expect(state.isEnabled).toBe(false);
      expect(state.disabledReason).toBe('disabled.slideMaster');
    }
    expect(startSlideshowController(e, 'current')).toBeNull();
    expect(slideStatus(e.snapshot())).toBeNull();
  });

  test('the rail names each template, counts its slides, and steps with the keys', async () => {
    const e = await opened();
    e.exec({ type: 'setPresentationView', view: 'slideMaster' });
    const t = createTranslator();
    const [master, ...layouts] = e.snapshot().slides;
    expect(masterRailLabel(t, master!)).toMatch(/^Master/);
    expect(masterRailUsage(t, master!)).toBe('Used by 3 slides');
    const branded = layouts.find((s) => s.name === 'Branded')!;
    expect(masterRailLabel(t, branded)).toBe('Branded');
    expect(masterRailUsage(t, branded)).toBe('Used by 2 slides');
    const unused = layouts.find((s) => (s.usedBy ?? 0) === 0);
    if (unused) expect(masterRailUsage(t, unused)).toBe('Not used by any slide');
    const all = e.snapshot().slides;
    expect(masterStepTarget('ArrowDown', all, master!.index)).toBe(all[1]!.index);
    expect(masterStepTarget('ArrowUp', all, master!.index)).toBeNull();
    expect(masterStepTarget('End', all, master!.index)).toBe(all.at(-1)!.index);
    expect(masterStepTarget('a', all, master!.index)).toBeNull();
    expect(masterBarTitle(t, branded)).toBe('Layout: Branded');
    expect(masterBarHint(t, branded)).toBe('Changes apply to the 2 slides that use this layout.');
    expect(masterBarHint(t, master!)).toBe(
      'Changes apply to every layout of this master and to its 3 slides.'
    );
  });

  test('the slide menu offers Close master, and no New slide, in the view', async () => {
    const e = await opened();
    const keys = () =>
      canvasMenuRows(CANVAS_CONTEXT_MENUS.slide, e).flatMap((r) =>
        r.kind === 'item' ? [`${r.slot}:${r.value ?? ''}:${r.labelKey ?? ''}`] : []
      );
    expect(keys()).toContain('view.presentationView:slideMaster:contextMenu.editLayout');
    expect(keys()).toContain('slide.add::contextMenu.newSlide');
    e.exec({ type: 'setPresentationView', view: 'slideMaster' });
    expect(keys()).toContain('view.presentationView:normal:masterView.close');
    expect(keys().some((k) => k.startsWith('slide.add'))).toBe(false);
  });

  test('a store change made outside the editor keeps the template on the canvas', async () => {
    const e = await opened();
    e.exec({ type: 'setPresentationView', view: 'slideMaster' });
    const before = e.snapshot().currentSlide;
    const part = '/ppt/slides/slide2.xml';
    const root = xmlPartOf(e.session!.store.package, part)!.root;
    const op = {
      op: 'setAttribute',
      part,
      node: root.id,
      name: { ns: '', local: 'show' },
      value: '1',
    };
    expect(e.session!.store.transact([op as never]).ok).toBe(true);
    expect(e.snapshot().presentationView).toBe('slideMaster');
    expect(e.snapshot().currentSlide).toBe(before);
  });

  test('Copy carries a layout shape', async () => {
    const e = await opened();
    e.exec({ type: 'setPresentationView', view: 'slideMaster' });
    e.goToSlide(e.snapshot().slides.find((s) => s.name === 'Branded')!.index);
    const header = e.snapshot().shapes!.find((s) => s.name === 'Header')!;
    e.setSelection({ kind: 'shapes', nodeIds: [header.nodeId] });
    expect(e.copy()).not.toBeNull();
  });

  test('no slide is selected in the view, and the link picker counts the slides', async () => {
    const e = await opened();
    e.exec({ type: 'setPresentationView', view: 'slideMaster' });
    expect(e.snapshot().selectedSlides).toEqual([]);
    expect(deckSlideCount(e)).toBe(3);
    e.exec({ type: 'setPresentationView', view: 'normal' });
    expect(e.snapshot().selectedSlides).toEqual([0]);
  });
});
