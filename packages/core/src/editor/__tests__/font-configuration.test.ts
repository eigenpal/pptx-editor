// The `fonts` option of `createPptxEditor`, DOCX Editor's way: what a deck asks the host's
// font origins, how their faces rank, and what stays the same without them.

import { afterEach, describe, expect, test } from 'bun:test';
import { FONT_ASSET_ROOT, packagedFonts } from '@docx-editor.dev/fonts';
import { cjkFonts } from '@docx-editor.dev/fonts-cjk/browser';
import { attr, elements, find, readPptx, xmlPartOf } from '../../store/index.ts';
import {
  createPptxEditor,
  defineFontResolver,
  type EditorFontError,
  type FontOrigin,
  type FontResolutionRequest,
  type PptxEditorOptions,
} from '../index.ts';
import { deckDefaultFamily, deckFontFamilies } from '../deck-font-request.ts';
import { resolveSlide } from '../../cascade/index.ts';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { face, fontBytes, settle, themedDeck, missingFontsDeck } from './font-test-helpers.ts';

const realFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = realFetch;
});

const substituted = (editor: ReturnType<typeof createPptxEditor>): string[] =>
  editor.snapshot().substitutedFonts.map((f) => f.family);

describe('the request a deck makes', () => {
  test('theme fonts come first, then every typeface the deck names, East Asian too', () => {
    const r = readPptx(
      themedDeck({ major: 'Aptos Display', minor: 'Aptos', latin: 'Segoe UI', ea: 'Yu Gothic' })
    );
    if (!r.ok) throw new Error('deck');
    const families = deckFontFamilies(r.pkg);
    expect(families.slice(0, 2)).toEqual(['Aptos Display', 'Aptos']);
    expect(families).toContain('Segoe UI');
    expect(families).toContain('Yu Gothic');
    expect(new Set(families.map((f) => f.toLowerCase())).size).toBe(families.length);
  });

  test("the theme script fonts of the runs' languages, not every script font", () => {
    const r = readPptx(
      buildPptx({
        scriptFonts: [
          ['Jpan', 'Yu Gothic'],
          ['Hang', 'Malgun Gothic'],
          ['Arab', 'Tahoma'],
        ],
        slides: [
          {
            shapes: shape({
              id: 2,
              x: 0,
              y: 0,
              cx: 914_400,
              cy: 914_400,
              text: para(run('日本語', 'altLang="ja-JP"')),
            }),
          },
        ],
      })
    );
    if (!r.ok) throw new Error('deck');
    const families = deckFontFamilies(r.pkg);
    expect(families).toContain('Yu Gothic');
    expect(families).not.toContain('Malgun Gothic');
    expect(families).not.toContain('Tahoma');
  });

  test("the default text style's language picks a script font for runs with no language", () => {
    const r = readPptx(
      buildPptx({
        defaultLanguage: 'ja-JP',
        scriptFonts: [['Jpan', 'Yu Gothic']],
        slides: [
          {
            shapes: shape({
              id: 2,
              x: 0,
              y: 0,
              cx: 914_400,
              cy: 914_400,
              text: para('<a:r><a:rPr dirty="0"/><a:t>日本語</a:t></a:r>'),
            }),
          },
        ],
      })
    );
    if (!r.ok) throw new Error('deck');
    // The cascade gives the run Yu Gothic, so the deck asks for it.
    const sp = resolveSlide(r.pkg, 0)!.elements.find((e) => e.kind === 'shape' && e.text);
    expect(sp?.kind === 'shape' ? sp.text?.paragraphs[0].runs[0].props.ea : undefined).toBe(
      'Yu Gothic'
    );
    expect(deckFontFamilies(r.pkg)).toContain('Yu Gothic');
  });

  test('the default family is the theme minor Latin font, not Calibri', () => {
    const r = readPptx(themedDeck({ minor: 'Aptos' }));
    if (!r.ok) throw new Error('deck');
    expect(deckDefaultFamily(r.pkg)).toBe('Aptos');
  });

  test('a resolver receives the families, the default family, and a signal', async () => {
    const requests: FontResolutionRequest[] = [];
    const editor = createPptxEditor({
      fonts: (request) => (requests.push(request), undefined),
    });
    await editor.load(themedDeck({ major: 'Aptos Display', minor: 'Aptos', latin: 'Segoe UI' }));
    expect(requests).toHaveLength(1);
    expect(requests[0]!.families.slice(0, 2)).toEqual(['Aptos Display', 'Aptos']);
    expect(requests[0]!.families).toContain('Segoe UI');
    expect(requests[0]!.defaultFamily).toBe('Aptos');
    expect(requests[0]!.signal?.aborted).toBe(false);
    editor.destroy();
  });
});

describe('host faces', () => {
  test('a host face removes its family from the substitution report', async () => {
    const editor = createPptxEditor({ fonts: { sources: [face('Segoe UI')] } });
    await editor.load(missingFontsDeck());
    expect(substituted(editor)).not.toContain('Segoe UI');
    expect(substituted(editor)).toContain('Aptos');
    editor.destroy();
  });

  test('a host substitution draws its target face under the requested name', async () => {
    const editor = createPptxEditor({
      fonts: {
        sources: [face('Acme Sans')],
        substitutions: [
          {
            from: { family: 'Segoe UI', weight: 400, style: 'normal' },
            to: { family: 'Acme Sans', weight: 400, style: 'normal' },
          },
        ],
      },
    });
    await editor.load(missingFontsDeck());
    const segoe = editor.snapshot().substitutedFonts.find((f) => f.family === 'Segoe UI');
    expect(segoe).toEqual({ family: 'Segoe UI', substitute: 'Acme Sans', metricCompatible: false });
    editor.destroy();
  });

  test('packagedFonts() is an origin: Calibri still measures with Carlito', async () => {
    const editor = createPptxEditor({ fonts: packagedFonts() });
    await editor.load(missingFontsDeck());
    expect(editor.snapshot().substitutedFonts.find((f) => f.family === 'Calibri')).toEqual({
      family: 'Calibri',
      substitute: 'Carlito',
      metricCompatible: true,
    });
    editor.destroy();
  });

  test('a face with bytes that do not match its hash is refused and reported', async () => {
    const errors: EditorFontError[] = [];
    const good = face('Segoe UI');
    const editor = createPptxEditor({
      fonts: { sources: [{ ...good, hash: 'sha256:0000' }] },
      onFontError: (e) => errors.push(e),
    });
    await editor.load(missingFontsDeck());
    expect(errors.map((e) => e.code)).toEqual(['hashMismatch']);
    expect(substituted(editor)).toContain('Segoe UI');
    editor.destroy();
  });

  test('a face over maxFontBytes is refused and its siblings still install', async () => {
    const errors: EditorFontError[] = [];
    const editor = createPptxEditor({
      fonts: {
        maxFontBytes: 200_000,
        sources: [
          face('Segoe UI', 'LiberationSans-Regular.ttf'),
          face('Aptos', 'Carlito-Regular.ttf'),
        ],
      },
      onFontError: (e) => errors.push(e),
    });
    await editor.load(missingFontsDeck());
    expect(fontBytes('Carlito-Regular.ttf').byteLength).toBeGreaterThan(200_000);
    expect(errors.length).toBeGreaterThan(0);
    expect(substituted(editor)).toContain('Aptos');
    editor.destroy();
  });

  test('a failing resolver reports through onFontError and the deck still opens', async () => {
    const errors: EditorFontError[] = [];
    const editor = createPptxEditor({
      fonts: () => Promise.reject(new Error('offline')),
      onFontError: (e) => errors.push(e),
    });
    expect((await editor.load(missingFontsDeck())).ok).toBe(true);
    expect(errors).toHaveLength(1);
    expect(errors[0]!.name).toBe('EditorFontError');
    expect(errors[0]!.diagnostic).toBe('offline');
    editor.destroy();
  });

  test('a family picked in the font picker is asked of the resolver', async () => {
    const asked: string[][] = [];
    const editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      fonts: defineFontResolver(({ families }) => (asked.push([...families]), undefined)),
    });
    await editor.load(themedDeck({ minor: 'Aptos' }));
    const slide = xmlPartOf(
      editor.session!.store.package,
      editor.session!.info.slides[0]!.partName
    )!;
    const sp = elements(find(slide.root, 'p:spTree')!).find(
      (c) => attr(find(c, 'p:cNvPr'), 'name') === 'Text'
    )!;
    editor.setSelection({
      kind: 'text',
      nodeId: sp.id,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 5 },
    });
    expect(editor.exec({ type: 'setFontFamily', family: 'Zapfino Extra' }).ok).toBe(true);
    await settle();
    expect(asked.at(-1)).toEqual(['Zapfino Extra']);
    editor.destroy();
  });
});

describe('the default stays the same', () => {
  test('no fonts: the substitutes load, and nothing outside the package is fetched', async () => {
    const urls: string[] = [];
    globalThis.fetch = ((input: RequestInfo | URL, init?: RequestInit) => {
      urls.push(String(input instanceof Request ? input.url : input));
      return realFetch(input, init);
    }) as typeof fetch;
    const editor = createPptxEditor();
    await editor.load(missingFontsDeck());
    expect(editor.snapshot().substitutedFonts.find((f) => f.family === 'Calibri')?.substitute).toBe(
      'Carlito'
    );
    for (const url of urls) expect(url.startsWith(FONT_ASSET_ROOT.href)).toBe(true);
    editor.destroy();
  });
});

describe('origin types', () => {
  test('packagedFonts() and cjkFonts() are accepted as fonts and as list origins', () => {
    // Compile-time checks: the `@docx-editor.dev/fonts` and `@docx-editor.dev/fonts-cjk`
    // resolvers carry the same mark key, so no adapter code is needed.
    const asOption: PptxEditorOptions['fonts'] = packagedFonts();
    const cjkOption: PptxEditorOptions['fonts'] = cjkFonts();
    const list: FontOrigin[] = [packagedFonts(), cjkFonts()];
    expect(typeof asOption).toBe('function');
    expect(typeof cjkOption).toBe('function');
    expect(list).toHaveLength(2);
  });
});
