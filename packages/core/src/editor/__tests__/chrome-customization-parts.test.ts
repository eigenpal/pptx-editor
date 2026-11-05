// The framework-free parts behind the adapters' customization API: the link popup store
// (`useHyperlinkPopup`), the named menu rows (`Menu.Cut`, `ContextMenu.Cut`), the source helpers
// of `usePptxSource`, and the host document of `useEditorHostDocument`.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  CANVAS_CONTEXT_MENUS,
  CHROME_MENUS,
  CONTEXT_MENU_NAMED_PARTS,
  MENU_NAMED_PARTS,
  TOOLBAR_NAMED_PARTS,
  chromeMenuRowKey,
  createHyperlinkPopupStore,
  createPptxEditor,
  editorHostDocument,
  defineFontResolver,
  onEditorHostChange,
  pptxFontsShape,
  pptxSourceBytes,
  resolveEagerPptxFonts,
  type ChromeMenuEntry,
  type PptxEditorInstance,
} from '../index.ts';
import { setSurfaceRoot } from '../surface-registry.ts';
import { attr, elements, find, xmlPartOf } from '../../store/index.ts';

const IN = 914400;

async function textEditor(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          name: 'T',
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: IN,
          text: para(run('one two')),
        }),
      },
    ],
  });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

/** Select characters `from` to `to` of the text box. */
function select(e: PptxEditorInstance, from: number, to: number): void {
  const root = xmlPartOf(e.session!.store.package, e.session!.info.slides[0]!.partName)!.root;
  const box = elements(find(root, 'p:spTree')!).find(
    (c) => attr(find(c, 'p:cNvPr'), 'name') === 'T'
  )!;
  e.setSelection({
    kind: 'text',
    nodeId: box.id,
    anchor: { paragraph: 0, offset: from },
    focus: { paragraph: 0, offset: to },
  });
}

describe('the link popup store', () => {
  test('opens for the selection, applies the draft, and closes', async () => {
    const e = await textEditor();
    select(e, 0, 3);
    const store = createHyperlinkPopupStore();
    store.setEditor(e);
    let changes = 0;
    const off = store.subscribe(() => void (changes += 1));
    store.openAtCaret();
    expect(store.getState()).toMatchObject({
      mode: 'editing',
      href: null,
      draft: '',
      canEdit: true,
    });
    store.setUrl('https://example.test/a');
    expect(store.commitEdit()).toBe(true);
    expect(store.getState().mode).toBe('closed');
    expect(e.snapshot().format?.link).toBe('https://example.test/a');
    // Opened again on the linked word, the field holds its target, and Remove takes it off.
    store.openAtCaret();
    expect(store.getState()).toMatchObject({
      url: 'https://example.test/a',
      href: 'https://example.test/a',
    });
    expect(store.getState().url).toBe('https://example.test/a');
    expect(store.unlink()).toBe(true);
    expect(e.snapshot().format?.link).toBeUndefined();
    expect(changes).toBeGreaterThan(3);
    off();
    store.setEditor(null);
    e.destroy();
  });

  test('a refused Apply stays open and says why', async () => {
    const e = await textEditor();
    select(e, 0, 3);
    const store = createHyperlinkPopupStore();
    store.setEditor(e);
    store.openAtCaret();
    store.setUrl('javascript:alert(1)');
    expect(store.commitEdit()).toBe(false);
    expect(store.getState()).toMatchObject({
      mode: 'editing',
      error: 'errors.unsafeLink',
    });
    // A new draft clears the refusal.
    store.setUrl('https://example.test');
    expect(store.getState()).toMatchObject({ error: null });
    store.setEditor(null);
    e.destroy();
  });

  test('open() reads the selected link and ignores caller-supplied targets', async () => {
    const e = await textEditor();
    select(e, 0, 3);
    expect(e.exec({ type: 'setLink', url: 'https://example.test/' }).ok).toBe(true);
    select(e, 1, 1);
    const store = createHyperlinkPopupStore();
    store.setEditor(e);
    store.open({ url: 'javascript:alert(1)', href: 'javascript:alert(1)' });
    expect(store.getState()).toMatchObject({
      mode: 'reading',
      url: 'https://example.test/',
      href: 'https://example.test/',
    });
    store.beginEdit();
    expect(store.getState().mode).toBe('editing');
    store.close();
    store.setUrl('x');
    expect(store.getState().draft).toBe('');
    store.setEditor(null);
    e.destroy();
  });

  test('without an editor, caller-supplied targets cannot open or copy', async () => {
    const store = createHyperlinkPopupStore();
    store.open({ url: 'javascript:alert(1)', href: 'https://example.test/' });
    expect(store.getState().mode).toBe('closed');
    expect(await store.copy()).toBe(false);
    expect(store.openTarget()).toBe(false);
  });
});

describe('the named menu rows', () => {
  const keys = (entries: readonly ChromeMenuEntry[]): string[] =>
    entries.flatMap((e) =>
      e.kind === 'item' ? [chromeMenuRowKey(e)] : e.kind === 'submenu' ? keys(e.items) : []
    );

  test('every menu bar row names a row of the bar', () => {
    const rows = new Set(CHROME_MENUS.flatMap((menu) => keys(menu.entries)));
    for (const [name, row] of Object.entries(MENU_NAMED_PARTS))
      expect([name, rows.has(chromeMenuRowKey(row))]).toEqual([name, true]);
  });

  test('every context menu row names a row of a canvas menu', () => {
    const rows = new Set(Object.values(CANVAS_CONTEXT_MENUS).flatMap((entries) => keys(entries)));
    for (const [name, row] of Object.entries(CONTEXT_MENU_NAMED_PARTS))
      expect([name, rows.has(chromeMenuRowKey(row))]).toEqual([name, true]);
  });

  test('a name that the toolbar also uses stands for the same slot', () => {
    const toolbar = TOOLBAR_NAMED_PARTS as Readonly<Record<string, string>>;
    for (const table of [MENU_NAMED_PARTS, CONTEXT_MENU_NAMED_PARTS])
      for (const [name, row] of Object.entries(table))
        if (toolbar[name] && !('value' in row))
          expect([name, row.slot]).toEqual([name, toolbar[name]]);
  });
});

describe('the source helpers', () => {
  test('bytes pass through, and a URL needs a fetch', () => {
    const bytes = new Uint8Array([1, 2]);
    expect(pptxSourceBytes(bytes)).toBe(bytes);
    expect(pptxSourceBytes(bytes.buffer)).toEqual(bytes);
    expect(pptxSourceBytes('/deck.pptx')).toBeNull();
  });

  test('a marked resolver makes the fonts on demand; values and loaders are eager', () => {
    expect(pptxFontsShape(undefined)).toBe('absent');
    expect(pptxFontsShape({ sources: [] })).toBe('eager');
    expect(pptxFontsShape(() => ({ sources: [] }))).toBe('eager');
    const resolver = defineFontResolver(() => undefined);
    expect(pptxFontsShape([{ sources: [] }, resolver])).toBe('on-demand');
  });

  test('one eager origin that fails does not take the others with it', async () => {
    const warn = console.warn;
    console.warn = () => undefined;
    try {
      const composed = await resolveEagerPptxFonts([
        () => Promise.reject(new Error('offline')),
        { supportedFamilies: ['Acme Sans'] },
      ]);
      expect(composed?.supportedFamilies).toContain('Acme Sans');
      expect(await resolveEagerPptxFonts([() => Promise.reject(new Error('x'))])).toBeUndefined();
    } finally {
      console.warn = warn;
    }
  });
});

describe('the host document', () => {
  test('follows the element the slide is mounted in', () => {
    const editor = {};
    const other = { nodeType: 1 } as unknown as Document;
    let calls = 0;
    const off = onEditorHostChange(editor, () => void (calls += 1));
    // With no DOM and no slide, there is no document.
    expect(editorHostDocument(editor)).toBe(
      typeof document === 'undefined' ? null : (document as Document)
    );
    setSurfaceRoot(editor, { ownerDocument: other } as unknown as HTMLElement);
    expect(editorHostDocument(editor)).toBe(other);
    expect(calls).toBe(1);
    // The same document again is no change.
    setSurfaceRoot(editor, { ownerDocument: other } as unknown as HTMLElement);
    expect(calls).toBe(1);
    off();
    setSurfaceRoot(editor, null);
    expect(calls).toBe(1);
  });
});
