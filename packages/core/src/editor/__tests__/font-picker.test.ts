// The font family picker's groups: the slide's theme fonts (Headings, Body), the deck's other
// families, then every family the editor offers; the search filter, the face value, and the
// scroll to the current family. After DOCX Editor's FontFamily catalog test.

import { afterEach, describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, elements, find, xmlPartOf, type Element } from '../../store/index.ts';
import {
  createPptxEditor,
  firstFontPickerFamily,
  focusFontOption,
  fontFaceValue,
  fontPickerGroups,
  fontPickerRow,
  isCurrentFontRow,
  resetDevWarnings,
  revealCurrentFont,
  runToolbarCommand,
  slideThemeFonts,
  type FontPickerRow,
  type PptxEditorInstance,
} from '../index.ts';
import { GlobalWindow } from 'happy-dom';

// A DOM of this file's own: the core run has no global DOM (`bun run test`).
const win = new GlobalWindow();
const document = win.document as unknown as Document;

const IN = 914_400;

async function open(latin = 'Courier New', minorFont = 'Verdana'): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck(latin, minorFont))).ok).toBe(true);
  return editor;
}

function deck(latin: string, minorFont = 'Verdana'): Uint8Array {
  return buildPptx({
    majorFont: 'Georgia',
    minorFont,
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 2 * IN,
          cy: IN,
          text: para(run('Hi', '', `<a:latin typeface="${latin}"/>`)),
        }),
      },
    ],
  });
}

async function reopen(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

describe('fontPickerGroups', () => {
  test('theme fonts first with their roles, then the deck families, then every family', async () => {
    const editor = await open();
    expect(slideThemeFonts(editor.session, 0)).toEqual({ headings: 'Georgia', body: 'Verdana' });
    const groups = fontPickerGroups(editor, 0, '');
    expect(groups.map((g) => [g.id, g.labelKey])).toEqual([
      ['theme', 'toolbar.fontGroups.theme'],
      ['document', 'toolbar.fontGroups.document'],
      ['all', 'toolbar.fontGroups.all'],
    ]);
    // A theme row applies the theme font reference, not the family name.
    expect(groups[0]!.rows).toEqual([
      {
        family: 'Georgia',
        value: '+mj-lt',
        theme: 'major',
        face: '"Georgia"',
        roleKey: 'toolbar.fontHeadings',
        disabled: false,
      },
      {
        family: 'Verdana',
        value: '+mn-lt',
        theme: 'minor',
        face: '"Verdana"',
        roleKey: 'toolbar.fontBody',
        disabled: false,
      },
    ]);
    expect(groups[1]!.rows.every((r) => r.value === r.family && r.theme === null)).toBe(true);
    // The theme fonts do not repeat in the deck group; the run's own family is there.
    const used = groups[1]!.rows.map((r) => r.family);
    expect(used).toContain('Courier New');
    expect(used).not.toContain('Georgia');
    expect(used.every((f) => groups[1]!.rows.find((r) => r.family === f)!.roleKey === null)).toBe(
      true
    );
    // The full list is the editor's, in its order.
    expect(groups[2]!.rows.map((r) => r.family)).toEqual([...editor.fontFamilies()]);
    expect(firstFontPickerFamily(groups)).toBe('+mj-lt');
  });

  test('the search filters every group and leaves out the empty ones', async () => {
    const editor = await open();
    const groups = fontPickerGroups(editor, 0, ' courier ');
    expect(groups.map((g) => g.id)).toEqual(['document', 'all']);
    expect(groups.every((g) => g.rows.every((r) => /courier/i.test(r.family)))).toBe(true);
    expect(firstFontPickerFamily(groups)).toBe('Courier New');
    expect(fontPickerGroups(editor, 0, 'no such font')).toEqual([]);
    expect(firstFontPickerFamily([])).toBe(null);
  });

  test('with no deck, only the editor families show', () => {
    const groups = fontPickerGroups({ session: null, fontFamilies: () => ['Arial'] }, 0, '');
    expect(groups.map((g) => g.id)).toEqual(['all']);
    expect(slideThemeFonts(null, 0)).toEqual({ headings: null, body: null });
  });
});

// The run's `a:latin` typeface in the slide part, as written.
function latinOf(editor: PptxEditorInstance): string | undefined {
  const root: Element = xmlPartOf(
    editor.session!.store.package,
    editor.session!.info.slides[0].partName
  )!.root;
  const r = elements(find(root, 'p:txBody')!)
    .filter((c) => c.local === 'p')
    .flatMap((p) => elements(p).filter((c) => c.local === 'r'))[0]!;
  return attr(child(child(r, 'a:rPr'), 'a:latin'), 'typeface');
}

function selectAll(editor: PptxEditorInstance): void {
  const root = xmlPartOf(
    editor.session!.store.package,
    editor.session!.info.slides[0].partName
  )!.root;
  const sp = elements(find(root, 'p:spTree')!).find((c) => c.local === 'sp')!;
  editor.setSelection({
    kind: 'text',
    nodeId: sp.id,
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 2 },
  });
}

// The rows the picker marks for the current selection.
function checked(editor: PptxEditorInstance): string[] {
  const f = editor.snapshot().format;
  return fontPickerGroups(editor, 0, '').flatMap((g) =>
    g.rows
      .filter((r: FontPickerRow) => isCurrentFontRow(r, f?.fontFamily, f?.fontFamilyTheme))
      .map((r) => `${g.id}:${r.family}${r.theme ? `:${r.theme}` : ''}`)
  );
}

describe('theme font rows', () => {
  test('a theme row writes the theme reference, and the check follows it', async () => {
    const editor = await open();
    selectAll(editor);
    expect(checked(editor)).toEqual(['document:Courier New', 'all:Courier New']);
    const headings = fontPickerGroups(editor, 0, '')[0]!.rows[0]!;
    expect(runToolbarCommand(editor, 'font.family', headings.value).ok).toBe(true);
    expect(latinOf(editor)).toBe('+mj-lt');
    // The label shows the theme's family; only the Headings row is marked.
    expect(editor.snapshot().format?.fontFamily).toBe('Georgia');
    expect(editor.snapshot().format?.fontFamilyTheme).toBe('major');
    expect(checked(editor)).toEqual(['theme:Georgia:major']);
    const body = fontPickerGroups(editor, 0, '')[0]!.rows[1]!;
    expect(runToolbarCommand(editor, 'font.family', body.value).ok).toBe(true);
    expect(latinOf(editor)).toBe('+mn-lt');
    expect(checked(editor)).toEqual(['theme:Verdana:minor']);
  });

  test('a literal family with a theme font name marks the literal rows only', async () => {
    const editor = await open('Georgia');
    selectAll(editor);
    expect(editor.snapshot().format?.fontFamily).toBe('Georgia');
    expect(editor.snapshot().format?.fontFamilyTheme).toBeUndefined();
    expect(checked(editor)).toEqual(['all:Georgia']);
  });

  test('a theme font with the same name as a literal row marks the theme row only', async () => {
    const editor = await open('+mn-lt');
    selectAll(editor);
    expect(editor.snapshot().format?.fontFamily).toBe('Verdana');
    expect(checked(editor)).toEqual(['theme:Verdana:minor']);
  });

  test('a reference other than the Latin theme fonts is refused', async () => {
    const editor = await open();
    selectAll(editor);
    expect(editor.exec({ type: 'setFontFamily', family: '+mj-ea' }).ok).toBe(false);
    expect(editor.exec({ type: 'setFontFamily', family: '+xx' }).ok).toBe(false);
    expect(latinOf(editor)).toBe('Courier New');
  });

  test('+mj-lt survives save and reopen, and follows a later theme change', async () => {
    const editor = await open();
    selectAll(editor);
    expect(editor.exec({ type: 'setFontFamily', family: '+mj-lt' }).ok).toBe(true);
    const again = await reopen(editor.save()!);
    expect(latinOf(again)).toBe('+mj-lt');
    selectAll(again);
    expect(again.snapshot().format?.fontFamilyTheme).toBe('major');
    expect(again.snapshot().format?.fontFamily).toBe('Georgia');
    // The theme's major font changes: the text follows it.
    const session = again.session!;
    const themePart = session.info.slides[0]!.themePart!;
    const theme = xmlPartOf(session.store.package, themePart)!.root;
    const latin = child(find(theme, 'a:majorFont')!, 'a:latin')!;
    expect(
      session.store.transact([
        {
          op: 'setAttribute',
          part: themePart,
          node: latin.id,
          name: { ns: '', local: 'typeface' },
          value: 'Tahoma',
        },
      ]).ok
    ).toBe(true);
    selectAll(again);
    expect(again.snapshot().format?.fontFamily).toBe('Tahoma');
    expect(again.snapshot().format?.fontFamilyTheme).toBe('major');
  });
});

describe('script faces', () => {
  // Runs with a Latin and an East Asian face, as CJK decks carry them.
  function cjkDeck(): Uint8Array {
    return buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: IN,
            text: para(run('曾經', '', '<a:latin typeface="Calibri"/><a:ea typeface="SimSun"/>')),
          }),
        },
      ],
    });
  }

  function facesOf(editor: PptxEditorInstance): Record<string, string | undefined> {
    const root = xmlPartOf(
      editor.session!.store.package,
      editor.session!.info.slides[0].partName
    )!.root;
    const rPr = child(find(root, 'a:r')!, 'a:rPr');
    return Object.fromEntries(
      ['latin', 'ea', 'cs'].map((s) => [s, attr(child(rPr, `a:${s}`), 'typeface')])
    );
  }

  test('a family applies to the Latin, East Asian, and complex script faces', async () => {
    const editor = await reopen(cjkDeck());
    selectAll(editor);
    expect(editor.exec({ type: 'setFontFamily', family: 'Noto Sans TC' }).ok).toBe(true);
    const again = await reopen(editor.save()!);
    expect(facesOf(again)).toEqual({
      latin: 'Noto Sans TC',
      ea: 'Noto Sans TC',
      cs: 'Noto Sans TC',
    });
    selectAll(again);
    expect(again.snapshot().format?.fontFamily).toBe('Noto Sans TC');
  });

  test('a theme font maps to the theme reference of each face', async () => {
    const editor = await reopen(cjkDeck());
    selectAll(editor);
    expect(editor.exec({ type: 'setFontFamily', family: '+mn-lt' }).ok).toBe(true);
    const again = await reopen(editor.save()!);
    expect(facesOf(again)).toEqual({ latin: '+mn-lt', ea: '+mn-ea', cs: '+mn-cs' });
  });

  test('a second pick keeps one face of each slot, in schema order', async () => {
    const editor = await reopen(cjkDeck());
    selectAll(editor);
    expect(editor.exec({ type: 'setFontFamily', family: 'Lato' }).ok).toBe(true);
    selectAll(editor);
    expect(editor.exec({ type: 'setFontFamily', family: 'Noto Sans TC' }).ok).toBe(true);
    const root = xmlPartOf(
      editor.session!.store.package,
      editor.session!.info.slides[0].partName
    )!.root;
    const rPr = child(find(root, 'a:r')!, 'a:rPr')!;
    expect(elements(rPr).map((c) => c.local)).toEqual(['latin', 'ea', 'cs']);
    expect(facesOf(editor).ea).toBe('Noto Sans TC');
  });
});

describe('fontPickerRow', () => {
  afterEach(() => resetDevWarnings());

  test('a theme reference gives the theme row; a family gives a plain row', async () => {
    const editor = await open();
    expect(fontPickerRow(editor, 0, '+mj-lt')).toMatchObject({
      family: 'Georgia',
      theme: 'major',
      disabled: false,
    });
    expect(fontPickerRow(editor, 0, 'Inter')).toMatchObject({
      family: 'Inter',
      value: 'Inter',
      disabled: false,
    });
  });

  test('a theme font the theme leaves empty gives a disabled row with its role', async () => {
    const editor = await open('Courier New', '');
    expect(slideThemeFonts(editor.session, 0).body).toBe(null);
    expect(fontPickerRow(editor, 0, '+mn-lt')).toEqual({
      family: '',
      value: '+mn-lt',
      theme: 'minor',
      face: null,
      roleKey: 'toolbar.fontBody',
      disabled: true,
    });
    // The default list leaves the empty theme font out.
    expect(fontPickerGroups(editor, 0, '')[0]!.rows.map((r) => r.theme)).toEqual(['major']);
  });

  test('a value that is not a font name gives no row', () => {
    expect(fontPickerRow(null, 0, 'Bad"; color: red')).toBe(null);
    expect(fontPickerRow(null, 0, '')).toBe(null);
  });
});

describe('fontFamilies and setFontFamilies', () => {
  test('a new list replaces the configured families and publishes one change', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, fontFamilies: ['Inter'] });
    const before = editor.fontFamilies();
    expect(before).toContain('Inter');
    let changes = 0;
    editor.subscribe(() => (changes += 1));
    // The same families: nothing changes, and the list keeps its identity.
    editor.setFontFamilies(['Inter']);
    expect(changes).toBe(0);
    expect(editor.fontFamilies()).toBe(before);
    editor.setFontFamilies(['Lato']);
    expect(changes).toBe(1);
    expect(editor.fontFamilies()).toContain('Lato');
    expect(editor.fontFamilies()).not.toContain('Inter');
    editor.setFontFamilies([]);
    expect(editor.fontFamilies()).not.toContain('Lato');
    editor.destroy();
  });
});

describe('focusFontOption', () => {
  function picker(search: boolean, selected: number | null, disabledFirst = false): HTMLElement {
    const root = document.createElement('div');
    root.className = 'pptx-font-picker';
    if (search) root.append(document.createElement('input'));
    const list = document.createElement('div');
    list.setAttribute('role', 'listbox');
    for (let i = 0; i < 3; i += 1) {
      const row = document.createElement('button');
      row.setAttribute('role', 'option');
      row.setAttribute('aria-selected', String(i === selected));
      row.dataset.index = String(i);
      if (disabledFirst && i === 0) row.disabled = true;
      list.append(row);
    }
    root.append(list);
    document.body.append(root);
    return list;
  }
  const focused = (): string | undefined =>
    (document.activeElement as HTMLElement | null)?.dataset?.index;

  test('a list with no search field focuses its current option, else its first enabled one', () => {
    focusFontOption(picker(false, 2));
    expect(focused()).toBe('2');
    focusFontOption(picker(false, null));
    expect(focused()).toBe('0');
    focusFontOption(picker(false, null, true));
    expect(focused()).toBe('1');
  });

  test('a list with a search field leaves focus alone', () => {
    const list = picker(true, 1);
    (document.activeElement as HTMLElement | null)?.blur();
    focusFontOption(list);
    expect(focused()).toBe(undefined);
    focusFontOption(null);
  });
});

describe('fontFaceValue', () => {
  test('quotes a valid family and refuses any other text', () => {
    expect(fontFaceValue('Open Sans')).toBe('"Open Sans"');
    expect(fontFaceValue('Bad"; color: red')).toBe(null);
    expect(fontFaceValue('a\\b')).toBe(null);
    expect(fontFaceValue('')).toBe(null);
  });
});

describe('revealCurrentFont', () => {
  const box = (top: number, height: number) =>
    ({ top, bottom: top + height, height, left: 0, right: 100, width: 100 }) as DOMRect;

  function list(rowTop: number): HTMLElement {
    const el = document.createElement('div');
    const row = document.createElement('button');
    row.setAttribute('role', 'option');
    row.setAttribute('aria-selected', 'true');
    el.append(row);
    el.getBoundingClientRect = () => box(0, 300);
    row.getBoundingClientRect = () => box(rowTop, 30);
    Object.defineProperty(el, 'clientHeight', { value: 300 });
    return el;
  }

  test('centers the current row when it is out of view', () => {
    const el = list(900);
    revealCurrentFont(el);
    expect(el.scrollTop).toBe(900 - 135);
  });

  test('leaves the list alone when the row is in view, or there is no current row', () => {
    const el = list(100);
    revealCurrentFont(el);
    expect(el.scrollTop).toBe(0);
    const empty = document.createElement('div');
    revealCurrentFont(empty);
    revealCurrentFont(null);
    expect(empty.scrollTop).toBe(0);
  });
});
