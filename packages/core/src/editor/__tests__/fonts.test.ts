import { describe, expect, test } from 'bun:test';
import { readPptx } from '../../store/index.ts';
import { uprightPieces } from '../../output/paint-text.ts';
import { createPptxEditor, namedFamilies, openPresentation, withFontOrigin } from '../index.ts';
// `face` makes a Liberation Sans face under another name: it stands in for a licensed face.
import { face, missingFontsDeck } from './font-test-helpers.ts';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';

const substituted = (editor: ReturnType<typeof createPptxEditor>): string[] =>
  editor.snapshot().substitutedFonts.map((f) => f.family);

describe('font coverage', () => {
  test('namedFamilies lists run, theme, and bullet fonts once', () => {
    const r = readPptx(missingFontsDeck());
    if (!r.ok) throw new Error('fixture');
    const names = namedFamilies(r.pkg);
    for (const f of ['Aptos', 'Segoe UI', 'Calibri', 'Arial']) expect(names).toContain(f);
    expect(new Set(names.map((n) => n.toLowerCase())).size).toBe(names.length);
  });

  test('substitutions are reported, with metric compatibility', async () => {
    const r = await openPresentation(missingFontsDeck());
    if (!r.ok) throw new Error('open');
    const subs = r.session.substitutedFamilies();
    expect(subs.find((s) => s.family === 'Aptos')?.metricCompatible).toBe(false);
    expect(subs.find((s) => s.family === 'Segoe UI')?.metricCompatible).toBe(false);
    expect(subs.find((s) => s.family === 'Calibri')).toEqual({
      family: 'Calibri',
      substitute: 'Carlito',
      metricCompatible: true,
    });
  });

  test('a deck that names a packaged open face draws with that face, not a substitute', async () => {
    const families = [
      'Carlito',
      'Caladea',
      'TeX Gyre Adventor',
      'Liberation Sans',
      'Liberation Serif',
      'Liberation Mono',
    ];
    const text = families
      .map((f) => para(run(`${f} text`, '', `<a:latin typeface="${f}"/>`)))
      .join('');
    const deck = buildPptx({
      slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 9_000_000, cy: 5_000_000, text }) }],
    });
    const r = await openPresentation(deck);
    if (!r.ok) throw new Error('open');
    const loaded = new Set(r.session.faces.map((f) => f.family));
    for (const f of families) expect(loaded).toContain(f);
    expect(r.session.substitutedFamilies().map((s) => s.family)).not.toEqual(
      expect.arrayContaining(families)
    );
    expect(
      r.session.layout(0)?.diagnostics.filter((d) => d.startsWith('font substituted'))
    ).toEqual([]);
  });

  test('the fonts resolver is asked for the deck families and its faces win', async () => {
    let asked: readonly string[] = [];
    const editor = createPptxEditor({
      fonts: (request) => {
        asked = request.families;
        return { sources: [face('Aptos')] };
      },
    });
    expect((await editor.load(missingFontsDeck())).ok).toBe(true);
    expect(asked).toContain('Aptos');
    expect(asked).toContain('Segoe UI');
    expect(substituted(editor)).not.toContain('Aptos');
    expect(substituted(editor)).toContain('Segoe UI');
    editor.destroy();
  });

  test('a failing resolver does not refuse the document', async () => {
    const errors: string[] = [];
    const editor = createPptxEditor({
      fonts: () => Promise.reject(new Error('offline')),
      onFontError: (error) => void errors.push(error.code),
    });
    expect((await editor.load(missingFontsDeck())).ok).toBe(true);
    expect(errors.length).toBeGreaterThan(0);
    expect(substituted(editor)).toContain('Aptos');
    editor.destroy();
  });

  test('setFonts re-measures and keeps the undo history', async () => {
    const editor = createPptxEditor();
    expect((await editor.load(missingFontsDeck())).ok).toBe(true);
    expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
    const before = editor.session!.layout(0);
    expect(substituted(editor)).toContain('Segoe UI');
    await editor.setFonts({ sources: [face('Segoe UI')] });
    expect(substituted(editor)).not.toContain('Segoe UI');
    expect(editor.session!.layout(0)).not.toBe(before);
    expect(editor.snapshot().slides).toHaveLength(2);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.snapshot().slides).toHaveLength(1);
    editor.destroy();
  });

  test('withFontOrigin adds faces after the host fonts and keeps them', async () => {
    const host = { sources: [face('Aptos')] };
    const editor = createPptxEditor({ fonts: host });
    expect((await editor.load(missingFontsDeck())).ok).toBe(true);
    expect(editor.fonts).toBe(host);
    expect(substituted(editor)).not.toContain('Aptos');
    expect(substituted(editor)).toContain('Segoe UI');
    // The local faces of the substitution notice take this path.
    await editor.setFonts(withFontOrigin(editor.fonts, { sources: [face('Segoe UI')] }));
    expect(substituted(editor)).not.toContain('Aptos');
    expect(substituted(editor)).not.toContain('Segoe UI');
    expect(withFontOrigin(undefined, host)).toBe(host);
    editor.destroy();
  });

  test('the substitution list is reference-stable while unchanged', async () => {
    const editor = createPptxEditor();
    await editor.load(missingFontsDeck());
    const a = editor.snapshot().substitutedFonts;
    editor.goToSlide(0);
    editor.setZoom(0.5);
    expect(editor.snapshot().substitutedFonts).toBe(a);
  });
});

describe('upright East Asian characters', () => {
  test('eaVert splits ideographs out, keeps Latin stretches together', () => {
    expect(uprightPieces('縦書きAB 1', 'eastAsian')).toEqual([
      { start: 0, end: 1, upright: true },
      { start: 1, end: 2, upright: true },
      { start: 2, end: 3, upright: true },
      { start: 3, end: 7, upright: false },
    ]);
  });

  test('wordArtVert stacks every character except spaces', () => {
    expect(uprightPieces('A B', 'all').map((p) => p.upright)).toEqual([true, false, true]);
  });

  test('horizontal text is one piece', () => {
    expect(uprightPieces('縦書き', 'none')).toEqual([{ start: 0, end: 3, upright: false }]);
  });

  test('a surrogate pair stays one character', () => {
    expect(uprightPieces('𠀋a', 'eastAsian')).toEqual([
      { start: 0, end: 2, upright: true },
      { start: 2, end: 3, upright: false },
    ]);
  });
});
