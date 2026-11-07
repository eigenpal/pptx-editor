// The outline of each shape of a multiple selection (surface-handles.ts `memberOutline`): a thin
// line in screen pixels just outside the shape, with a narrow halo, so it never covers the
// shape's text. The selection marks take `--doc-selection-mark`, not the chrome's primary color.

import { afterAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { GlobalWindow } from 'happy-dom';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import type { LayoutItem } from '../../layout/index.ts';
import { createPptxEditor } from '../index.ts';
import { withHalos } from '../surface-halo.ts';
import { MEMBER_OUTSET, memberOutline } from '../surface-handles.ts';
import { itemById } from '../surface.ts';

const window = new GlobalWindow();
const doc = window.document as unknown as Document;
afterAll(() => window.happyDOM.close());

const box = (id: number, x: number): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Box ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/>` +
  `</p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="914400"/><a:ext cx="1828800" cy="457200"/>` +
  `</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:p><a:r><a:t>Hybrid</a:t></a:r></a:p></p:txBody></p:sp>`;

async function items(): Promise<LayoutItem[]> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const deck = buildPptx({
    slides: [{ layout: 'blank', shapes: box(2, 914400) + box(3, 2743200) }],
  });
  expect((await e.load(deck)).ok).toBe(true);
  const layout = e.session!.layout(0)!;
  return e.snapshot().shapes!.map((s) => itemById(layout, s.nodeId)!);
}

/** EMU per CSS pixel at 100% zoom, as the surface uses. */
const EMU_PER_PX = 9_525;

const num = (el: Element, name: string): number => Number(el.getAttribute(name));

describe('the outline of one shape of several', () => {
  for (const zoom of [0.25, 1, 4]) {
    test(`is 1 screen pixel wide and lies outside the shape at zoom ${zoom}`, async () => {
      const px = EMU_PER_PX / zoom;
      const [item] = await items();
      const g = memberOutline(doc, item!, px);
      const [halo, outline] = withHalos([g])[0]!.children as unknown as Element[];
      expect(outline!.getAttribute('class')).toBe(
        'pptx-surface__outline pptx-surface__outline--member'
      );
      // The stroke is in screen pixels: it does not grow with zoom.
      expect(num(outline!, 'stroke-width')).toBeCloseTo(px);
      // Outside the box: the inner edge of the line and of its halo never reach the content.
      const { cx, cy } = item!.element.xfrm;
      expect(num(outline!, 'x')).toBeCloseTo(-MEMBER_OUTSET * px);
      expect(num(outline!, 'width')).toBeCloseTo(cx + 2 * MEMBER_OUTSET * px);
      expect(num(outline!, 'height')).toBeCloseTo(cy + 2 * MEMBER_OUTSET * px);
      expect(halo!.getAttribute('class')).toBe('pptx-surface__halo pptx-surface__halo--member');
      const haloWidth = num(halo!, 'stroke-width');
      expect(haloWidth).toBeCloseTo(2 * px);
      expect(MEMBER_OUTSET * px - haloWidth / 2).toBeGreaterThanOrEqual(-1e-6);
    });
  }
});

describe('a multiple selection on the slide', () => {
  test('outlines each shape thinly, outside its box; the common box keeps the handles', async () => {
    const g = globalThis as { ResizeObserver?: unknown };
    g.ResizeObserver ??= (window as unknown as { ResizeObserver: unknown }).ResizeObserver;
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const deck = buildPptx({
      slides: [{ layout: 'blank', shapes: box(2, 914400) + box(3, 2743200) }],
    });
    expect((await e.load(deck)).ok).toBe(true);
    const host = doc.createElement('div') as unknown as HTMLElement;
    doc.body.append(host as unknown as Node);
    e.attach(host);
    e.setZoom(0.5);
    const ids = e.snapshot().shapes!.map((s) => s.nodeId);
    e.setSelection({ kind: 'shapes', nodeIds: ids });
    const overlay = host.querySelector('.pptx-surface__overlay')!;
    const outlines = [...overlay.querySelectorAll('.pptx-surface__outline')];
    expect(outlines.length).toBe(2);
    const px = EMU_PER_PX / 0.5;
    for (const outline of outlines) {
      expect(outline.classList.contains('pptx-surface__outline--member')).toBe(true);
      expect(num(outline, 'stroke-width')).toBeCloseTo(px);
      expect(num(outline, 'x')).toBeLessThan(0);
    }
    expect(overlay.querySelectorAll('.pptx-surface__group-outline').length).toBe(1);
    expect(overlay.querySelectorAll('.pptx-surface__handle').length).toBe(8);
    // One shape alone keeps its outline on its box, under its own handles.
    e.setSelection({ kind: 'shapes', nodeIds: [ids[0]!] });
    const single = overlay.querySelector('.pptx-surface__outline')!;
    expect(single.classList.contains('pptx-surface__outline--member')).toBe(false);
    e.destroy();
  });
});

describe('selection mark color', () => {
  const css = readFileSync(new URL('../../styles/editor.css', import.meta.url), 'utf8');
  const rule = (selector: string): string =>
    css.slice(css.indexOf(`${selector} {`), css.indexOf('}', css.indexOf(`${selector} {`)));

  test('the marks on the slide use their own token, not the chrome primary color', () => {
    // A host can set a dark or light `--doc-primary` for its buttons; the marks stay an accent.
    expect(css).toMatch(/--doc-selection-mark: #[0-9a-f]{6};/);
    for (const selector of [
      '.pptx-surface__outline',
      '.pptx-surface__marquee',
      '.pptx-surface__handle',
    ]) {
      expect(rule(selector)).toContain('var(--doc-selection-mark)');
      expect(rule(selector)).not.toContain('--doc-primary');
    }
  });
});
