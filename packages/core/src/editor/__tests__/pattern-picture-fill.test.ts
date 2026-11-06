// Pattern and picture fills: `setShapePattern`, `setShapePicture`, and the pattern and tiled
// picture slide backgrounds. The write, the state the chrome reads, the refusals, undo, and save
// and reopen. Also the pattern dialog helpers that both adapters share.

import { describe, expect, test } from 'bun:test';
import { PNG } from 'pngjs';
import { createTranslator } from '@pptx-editor.dev/i18n';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  child,
  elements,
  find,
  presentationInfo,
  readPptx,
  shapeTreeOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import { PRESET_PATTERNS } from '../../output/preset-patterns.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import {
  PATTERN_CHOICES,
  patternCss,
  patternDialogStart,
  patternLabel,
} from '../pattern-fields.ts';

const box = (id: number, name: string): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1371600"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`;

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const deck = buildPptx({ slides: [{ layout: 'blank', shapes: box(2, 'Box') + box(3, 'Two') }] });
  expect((await e.load(deck)).ok).toBe(true);
  return e;
}

function png(): Uint8Array {
  const img = new PNG({ width: 4, height: 2 });
  img.data.fill(90);
  return new Uint8Array(PNG.sync.write(img));
}

function select(e: PptxEditorInstance, ...names: string[]): void {
  const ids = names.map((n) => e.snapshot().shapes!.find((s) => s.name === n)!.nodeId);
  e.setSelection({ kind: 'shapes', nodeIds: ids });
}

const pkgOf = (e: PptxEditorInstance): PptxPackage => e.session!.store.package;
const reopen = (e: PptxEditorInstance): PptxPackage => {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('reopen');
  return r.pkg;
};
const slidePart = (p: PptxPackage): string => presentationInfo(p).slides[0]!.partName;

/** The fill element of shape `name`: the child of `p:spPr` after the geometry. */
function fillOf(p: PptxPackage, name: string): Element | undefined {
  const sp = elements(shapeTreeOf(p, slidePart(p))).find(
    (el) => attr(find(el, 'p:cNvPr'), 'name') === name
  )!;
  return elements(child(sp, 'p:spPr')).find((el) => el.local.endsWith('Fill'));
}

const PATTERN = { preset: 'dkDnDiag', foreground: '#C00000', background: '#FFF2CC' } as const;

describe('setShapePattern', () => {
  test('writes the pattern on each selected shape in one undo step, and the state shows it', async () => {
    const e = await open();
    select(e, 'Box', 'Two');
    expect(e.exec({ type: 'setShapePattern', pattern: PATTERN }).ok).toBe(true);
    const p = reopen(e);
    for (const name of ['Box', 'Two']) {
      const f = fillOf(p, name)!;
      expect([f.local, attr(f, 'prst')]).toEqual(['pattFill', 'dkDnDiag']);
      expect(attr(child(child(f, 'a:fgClr'), 'a:srgbClr'), 'val')).toBe('C00000');
      expect(attr(child(child(f, 'a:bgClr'), 'a:srgbClr'), 'val')).toBe('FFF2CC');
    }
    select(e, 'Box');
    expect(e.snapshot().geometry?.fill).toBe('pattern');
    expect(e.snapshot().geometry?.pattern).toEqual(PATTERN);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(fillOf(pkgOf(e), 'Box')).toBeUndefined();
  });

  test('refuses an unknown preset, a color that is not #RRGGBB, and no shape', async () => {
    const e = await open();
    select(e, 'Box');
    const bad = (pattern: unknown) =>
      e.exec({ type: 'setShapePattern', pattern: pattern as typeof PATTERN });
    expect(bad({ ...PATTERN, preset: 'pct5" x="1' })).toMatchObject({
      ok: false,
      reason: 'errors.unsupported',
    });
    expect(bad({ ...PATTERN, foreground: 'red' })).toMatchObject({
      ok: false,
      reason: 'errors.unsupported',
    });
    expect(bad(null)).toMatchObject({ ok: false, reason: 'errors.unsupported' });
    e.setSelection({ kind: 'none' });
    expect(bad(PATTERN).ok).toBe(false);
    expect(fillOf(pkgOf(e), 'Box')).toBeUndefined();
  });
});

describe('setShapePicture', () => {
  test('stretches or tiles one media part over the shapes, and keeps it after reopen', async () => {
    const e = await open();
    select(e, 'Box', 'Two');
    // The type comes from the bytes: a wrong `mime` from a file name does not matter.
    expect(e.exec({ type: 'setShapePicture', bytes: png(), mime: 'image/jpeg' }).ok).toBe(true);
    const p = reopen(e);
    const fill = fillOf(p, 'Box')!;
    expect(fill.local).toBe('blipFill');
    expect(child(fill, 'a:stretch')).toBeDefined();
    const rId = attr(child(fill, 'a:blip'), 'r:embed');
    expect(attr(child(fillOf(p, 'Two'), 'a:blip'), 'r:embed')).toBe(rId);
    expect([...p.parts.keys()].filter((n) => n.startsWith('/ppt/media/'))).toEqual([
      '/ppt/media/image1.png',
    ]);
    select(e, 'Box');
    expect(e.snapshot().geometry?.fill).toBe('picture');
    expect(
      e.exec({ type: 'setShapePicture', bytes: png(), mime: 'image/png', tile: true }).ok
    ).toBe(true);
    const tile = child(fillOf(reopen(e), 'Box'), 'a:tile');
    expect([attr(tile, 'sx'), attr(tile, 'algn')]).toEqual(['100000', 'tl']);
  });

  test('refuses bytes that are no PNG, JPEG, or GIF with a size', async () => {
    const e = await open();
    select(e, 'Box');
    const r = e.exec({
      type: 'setShapePicture',
      bytes: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
      mime: 'image/png',
    });
    expect(r).toMatchObject({ ok: false, reason: 'errors.unsupported' });
    expect([...pkgOf(e).parts.keys()].some((n) => n.startsWith('/ppt/media/'))).toBe(false);
  });
});

describe('pattern and tiled picture backgrounds', () => {
  test('setSlideBackground writes a pattern and a tiled picture, and the state shows them', async () => {
    const e = await open();
    const bgFill = (p: PptxPackage): Element | undefined =>
      elements(find(xmlPartOf(p, slidePart(p))?.root, 'p:bgPr'))[0];
    expect(
      e.exec({ type: 'setSlideBackground', background: { kind: 'pattern', pattern: PATTERN } }).ok
    ).toBe(true);
    expect(attr(bgFill(reopen(e)), 'prst')).toBe('dkDnDiag');
    expect(e.snapshot().slideBackground).toMatchObject({ kind: 'pattern', pattern: PATTERN });
    const picture = { kind: 'picture', bytes: png(), mime: 'image/png', tile: true } as const;
    expect(e.exec({ type: 'setSlideBackground', background: picture }).ok).toBe(true);
    expect(child(bgFill(reopen(e)), 'a:tile')).toBeDefined();
    expect(e.snapshot().slideBackground.kind).toBe('picture');
  });
});

describe('pattern dialog helpers', () => {
  test('the gallery lists each preset once, and each has a name', () => {
    expect([...PATTERN_CHOICES].sort() as string[]).toEqual([...PRESET_PATTERNS].sort());
    const t = createTranslator();
    const names = PATTERN_CHOICES.map((p) => patternLabel(t, p));
    expect(names.slice(0, 2)).toEqual(['5%', '10%']);
    expect(patternLabel(t, 'dkDnDiag')).toBe('Dark downward diagonal');
    expect(names.every((n) => n.length > 0 && !n.startsWith('pattern.'))).toBe(true);
    expect(new Set(names).size).toBe(names.length);
  });

  test('a swatch draws the bitmap in the two colors, and the start pattern is 50% on white', () => {
    const css = patternCss({ preset: 'pct50', foreground: '#112233', background: '#FFFFFF' });
    expect(css).toStartWith('url("data:image/svg+xml,');
    expect(decodeURIComponent(css)).toContain('fill="#112233"');
    expect(css).toEndWith('0 0 / 8px 8px repeat');
    expect(patternDialogStart(null, '#ABCDEF')).toEqual({
      preset: 'pct50',
      foreground: '#ABCDEF',
      background: '#FFFFFF',
    });
    expect(patternDialogStart(PATTERN)).toBe(PATTERN);
  });
});
