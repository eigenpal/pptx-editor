// Regression tests for the editing defects (E2–E10, E12,
// E14, E15, E18, E20), S7, S8, and case-insensitive part names. Each test failed before its fix.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  child,
  elements,
  find,
  path,
  presentationInfo,
  readPptx,
  relationshipsOf,
  textOf,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import { partNameKey } from '@eigenpal/ooxml-core';
import { createPptxEditor, notesText, type PptxEditorInstance } from '../index.ts';
import { CLIPBOARD_MIME, MAX_PAYLOAD_CHARS, readClipboard } from '../clipboard.ts';
import { CommandRefusal } from '../../automation/commands/op-batch.ts';
import { xmlAttr } from '../../automation/commands/package-edits.ts';
import { dispatchBeforeInput, type SurfaceInputHost } from '../surface-input.ts';
import { moveCaret } from '../text-input.ts';

const IN = 914_400;
const SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const partOf = (e: PptxEditorInstance, i = 0): string => e.session!.info.slides[i].partName;
const rootOf = (e: PptxEditorInstance, i = 0): Element =>
  xmlPartOf(e.session!.store.package, partOf(e, i))!.root;
const tree = (e: PptxEditorInstance, i = 0): Element => find(rootOf(e, i), 'p:spTree')!;
const nameOf = (el: Element): string | undefined => attr(find(el, 'p:cNvPr'), 'name');
const idByName = (e: PptxEditorInstance, name: string, i = 0): string =>
  elements(tree(e, i)).find((c) => nameOf(c) === name)!.id;
const bodyText = (e: PptxEditorInstance, i = 0): string => textOf(find(rootOf(e, i), 'p:txBody'));
const caret = (e: PptxEditorInstance, nodeId: string, paragraph: number, offset: number) =>
  e.setSelection({
    kind: 'text',
    nodeId,
    anchor: { paragraph, offset },
    focus: { paragraph, offset },
  });

/** Four slides, each with a text shape `S<n>-text` and a box `S<n>-box`. */
const fourSlides = (): Uint8Array => {
  const s = (n: string) => ({
    shapes:
      shape({ id: 2, name: `${n}-text`, x: IN, y: IN, cx: 4 * IN, cy: IN, text: para(run(n)) }) +
      shape({ id: 3, name: `${n}-box`, x: IN, y: 3 * IN, cx: IN, cy: IN }),
  });
  return buildPptx({ slides: [s('S0'), s('S1'), s('S2'), s('S3')] });
};

/** One slide with one text shape holding `text` (paragraph XML). */
const textDeck = (text: string): Uint8Array =>
  buildPptx({
    slides: [{ shapes: shape({ id: 2, name: 'T', x: IN, y: IN, cx: 4 * IN, cy: IN, text }) }],
  });

describe('E2: load generations', () => {
  test('a later load wins over an earlier, slower one', async () => {
    let first = true;
    const e = createPptxEditor({
      substitutes: false,
      fonts: async () => {
        if (first) {
          first = false;
          await new Promise((r) => setTimeout(r, 60));
        }
        return undefined;
      },
    });
    const a = buildPptx({ slides: [{ shapes: '' }, { shapes: '' }, { shapes: '' }] });
    const b = buildPptx({ slides: [{ shapes: '' }] });
    const [ra, rb] = await Promise.all([e.load(a), e.load(b)]);
    expect(rb.ok).toBe(true);
    expect(ra).toEqual({ ok: false, reason: 'errors.superseded' });
    expect(e.snapshot().slides).toHaveLength(1);
    e.destroy();
  });

  test('a load in flight during destroy installs nothing', async () => {
    const e = createPptxEditor({
      substitutes: false,
      fonts: async () => (await new Promise((r) => setTimeout(r, 30)), undefined),
    });
    const pending = e.load(buildPptx({ slides: [{ shapes: '' }] }));
    e.destroy();
    expect(await pending).toEqual({ ok: false, reason: 'disabled.destroyed' });
    expect(e.session).toBeNull();
    expect(e.snapshot().status).toBe('empty');
    expect(await e.load(buildPptx({ slides: [{ shapes: '' }] }))).toEqual({
      ok: false,
      reason: 'disabled.destroyed',
    });
  });
});

describe('E3, E20: undo and redo go to the changed slide', () => {
  test('undo from another slide restores the selection on the edited slide', async () => {
    const e = await open(fourSlides());
    e.goToSlide(1);
    const box = idByName(e, 'S1-box', 1);
    e.setSelection({ kind: 'shapes', nodeIds: [box] });
    e.exec({ type: 'nudge', dx: 9525 * 8, dy: 0 });
    e.goToSlide(0);
    e.exec({ type: 'undo' });
    expect(e.snapshot().currentSlide).toBe(1);
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [box] });
    // Delete now removes the shape the user edited, not one on slide 0.
    e.exec({ type: 'delete' });
    expect(elements(tree(e, 0)).some((c) => nameOf(c) === 'S0-box')).toBe(true);
    expect(elements(tree(e, 1)).some((c) => nameOf(c) === 'S1-box')).toBe(false);
  });

  test('redo restores the selection the edit left', async () => {
    const e = await open(fourSlides());
    const box = idByName(e, 'S0-box');
    e.setSelection({ kind: 'shapes', nodeIds: [box] });
    e.exec({ type: 'nudge', dx: 9525, dy: 0 });
    e.exec({ type: 'undo' });
    e.setSelection({ kind: 'none' });
    e.goToSlide(2);
    e.exec({ type: 'redo' });
    expect(e.snapshot().currentSlide).toBe(0);
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [box] });
  });

  test('undo of a deleted slide returns to it', async () => {
    const e = await open(fourSlides());
    e.goToSlide(2);
    e.exec({ type: 'deleteSlide', index: 2 });
    e.goToSlide(0);
    e.exec({ type: 'undo' });
    expect(e.snapshot().currentSlide).toBe(2);
    expect(nameOf(elements(tree(e, 2))[2]!)).toBe('S2-text');
  });
});

describe('E4: moveSlide', () => {
  test('clears the selection', async () => {
    const e = await open(fourSlides());
    e.setSelection({ kind: 'shapes', nodeIds: [idByName(e, 'S0-box')] });
    e.exec({ type: 'moveSlide', from: 3, to: 1 });
    expect(e.snapshot().selection).toEqual({ kind: 'none' });
    expect(e.exec({ type: 'delete' }).ok).toBe(false);
  });
});

describe('E5, E6: slide parts and ids', () => {
  const link =
    '<a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rId9" action="ppaction://hlinksldjump"/></a:rPr><a:t>Go to 1</a:t></a:r>';
  const linked = (): Uint8Array =>
    buildPptx({
      slides: [
        { shapes: shape({ id: 2, x: IN, y: IN, cx: 3 * IN, cy: IN, text: para(run('First')) }) },
        {
          shapes: shape({ id: 2, x: IN, y: IN, cx: 3 * IN, cy: IN, text: `<a:p>${link}</a:p>` }),
          rels: [['rId9', SLIDE_REL, 'slide1.xml']],
        },
      ],
    });

  test('deleting a slide removes slide jumps to it, with their relationships', async () => {
    const e = await open(linked());
    expect(e.exec({ type: 'deleteSlide', index: 0 }).ok).toBe(true);
    const saved = readPptx(e.save()!);
    expect(saved.ok).toBe(true);
    if (!saved.ok) return;
    const part = presentationInfo(saved.pkg).slides[0].partName;
    for (const rel of relationshipsOf(saved.pkg, part))
      if (!rel.external) expect(saved.pkg.parts.has(rel.target!)).toBe(true);
    expect(relationshipsOf(saved.pkg, part).some((r) => r.type === SLIDE_REL)).toBe(false);
    const root = xmlPartOf(saved.pkg, part)!.root;
    expect(find(root, 'a:hlinkClick')).toBeUndefined();
    expect(textOf(find(root, 'p:txBody'))).toBe('Go to 1');
  });

  test('a deleted slide part name is not given to a new slide', async () => {
    const e = await open(fourSlides());
    const deleted = partOf(e, 1);
    e.exec({ type: 'deleteSlide', index: 1 });
    e.exec({ type: 'addSlide' });
    const names = e.session!.info.slides.map((s) => partNameKey(s.partName));
    expect(names).not.toContain(partNameKey(deleted));
    // Undo brings the deleted slide back under its own name.
    e.exec({ type: 'undo' });
    e.exec({ type: 'undo' });
    expect(partOf(e, 1)).toBe(deleted);
    expect(readPptx(e.save()!).ok).toBe(true);
  });

  test('a new slide id is the smallest free id from 256', async () => {
    const e = await open(fourSlides());
    const pres = e.session!.info.presentationPart;
    const list = () => path(xmlPartOf(e.session!.store.package, pres)!.root, 'p:sldIdLst')!;
    const last = elements(list())[3];
    e.session!.store.transact([
      {
        op: 'setAttribute',
        part: pres,
        node: last.id,
        name: { ns: '', local: 'id' },
        value: '2147483647',
      },
    ]);
    e.exec({ type: 'addSlide' });
    const ids = elements(list()).map((s) => Number(attr(s, 'id')));
    expect(ids).toContain(259);
    for (const id of ids) expect(id >= 256 && id <= 2_147_483_647).toBe(true);
  });
});

describe('E7: connectors', () => {
  const cxn =
    '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="5" name="C"/><p:cNvCxnSpPr><a:stCxn id="2" idx="3"/><a:endCxn id="3" idx="1"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>' +
    `<p:spPr><a:xfrm><a:off x="${2 * IN}" y="${IN}"/><a:ext cx="${IN}" cy="0"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>`;
  const deck = (): Uint8Array =>
    buildPptx({
      slides: [
        {
          shapes:
            shape({ id: 2, name: 'A', x: IN, y: IN, cx: IN, cy: IN }) +
            shape({ id: 3, name: 'B', x: 3 * IN, y: IN, cx: IN, cy: IN }) +
            cxn,
        },
      ],
    });
  const ends = (e: PptxEditorInstance) =>
    elements(tree(e))
      .filter((c) => c.local === 'cxnSp')
      .map((c) => [
        attr(find(c, 'p:cNvPr'), 'id'),
        attr(find(c, 'a:stCxn'), 'id'),
        attr(find(c, 'a:endCxn'), 'id'),
      ]);
  const idOf = (e: PptxEditorInstance, nodeId: string) =>
    attr(
      find(
        elements(tree(e)).find((c) => c.id === nodeId)!,
        'p:cNvPr'
      ),
      'id'
    );

  test('duplicate glues the copied connector to the copied shapes', async () => {
    const e = await open(deck());
    e.setSelection({ kind: 'shapes', nodeIds: ['A', 'B', 'C'].map((n) => idByName(e, n)) });
    e.exec({ type: 'duplicate' });
    const sel = e.snapshot().selection;
    if (sel.kind !== 'shapes') throw new Error('no selection');
    const [a, b] = sel.nodeIds.map((id) => idOf(e, id));
    expect(ends(e)).toEqual([
      ['5', '2', '3'],
      [idOf(e, sel.nodeIds[2]), a, b],
    ]);
  });

  test('paste glues the pasted connector to the pasted shapes, or to nothing', async () => {
    const e = await open(deck());
    e.setSelection({ kind: 'shapes', nodeIds: ['A', 'B', 'C'].map((n) => idByName(e, n)) });
    e.exec({ type: 'paste', payload: e.copy()! });
    const sel = e.snapshot().selection;
    if (sel.kind !== 'shapes') throw new Error('no selection');
    const [a, b, c] = sel.nodeIds.map((id) => idOf(e, id));
    expect(ends(e)[1]).toEqual([c, a, b]);
    // The connector alone is not glued to the originals.
    e.setSelection({ kind: 'shapes', nodeIds: [idByName(e, 'C')] });
    e.exec({ type: 'paste', payload: e.copy()! });
    expect(ends(e)[2].slice(1)).toEqual([undefined, undefined]);
  });
});

describe('part names', () => {
  const PNG = Uint8Array.from(
    atob(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
    ),
    (c) => c.charCodeAt(0)
  );

  test('a new picture does not take a name that differs only in case', async () => {
    const e = await open(
      buildPptx({
        slides: [{ shapes: '' }],
        media: [{ name: 'Image1.png', bytes: PNG, contentType: 'image/png' }],
      })
    );
    const r = e.exec({ type: 'insertImage', bytes: PNG, mime: 'image/png', width: 1, height: 1 });
    expect(r.ok).toBe(true);
    const keys = [...e.session!.store.package.parts.keys()].map(partNameKey);
    expect(new Set(keys).size).toBe(keys.length);
    expect(readPptx(e.save()!).ok).toBe(true);
  });
});

describe('E8: fields', () => {
  test('typing after a slide number field lands at the end of the paragraph', async () => {
    const fld =
      '<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld>';
    const e = await open(textDeck(`<a:p>${run('Page ')}${fld}${run(' end')}</a:p>`));
    const id = idByName(e, 'T');
    const item = e.session!.layout(0)!.items.find((i) => i.element.nodeId === id)!;
    if (item.kind !== 'shape' || !item.text) throw new Error('no text');
    const end = item.text.paragraphLengths[0];
    expect(end).toBe('Page 1 end'.length);
    caret(e, id, 0, end);
    expect(e.input('text', 'Z').ok).toBe(true);
    expect(bodyText(e)).toBe('Page ‹#› endZ');
    // Backspace at the field's end removes the whole field.
    caret(e, id, 0, 'Page 1'.length);
    e.input('deleteBackward');
    expect(bodyText(e)).toBe('Page  endZ');
  });
});

describe('E9: characters outside the BMP', () => {
  test('backspace deletes a whole emoji', async () => {
    const e = await open(textDeck(para(run('a\u{1F600}'))));
    const id = idByName(e, 'T');
    caret(e, id, 0, 3);
    expect(e.input('deleteBackward').ok).toBe(true);
    expect(bodyText(e)).toBe('a');
  });

  test('arrows step over a whole emoji, and a caret inside one snaps out', async () => {
    const e = await open(textDeck(para(run('a\u{1F600}b'))));
    const id = idByName(e, 'T');
    const item = e.session!.layout(0)!.items.find((i) => i.element.nodeId === id)!;
    if (item.kind !== 'shape') throw new Error('no shape');
    expect(moveCaret(item, { paragraph: 0, offset: 1 }, 'ArrowRight').offset).toBe(3);
    expect(moveCaret(item, { paragraph: 0, offset: 3 }, 'ArrowLeft').offset).toBe(1);
    caret(e, id, 0, 2);
    const sel = e.snapshot().selection;
    expect(sel.kind === 'text' && sel.focus.offset).toBe(1);
    e.input('text', 'x');
    expect(bodyText(e)).toBe('ax\u{1F600}b');
  });

  test('a grapheme of several code points is one step', async () => {
    const family = '\u{1F468}‍\u{1F469}‍\u{1F467}';
    const e = await open(textDeck(para(run(`a${family}`))));
    const id = idByName(e, 'T');
    caret(e, id, 0, 1 + family.length);
    e.input('deleteBackward');
    expect(bodyText(e)).toBe('a');
  });
});

describe('E10: pasted text', () => {
  test('U+000B becomes a line break and control characters are removed', async () => {
    const e = await open(textDeck(para(run('Hello'))));
    const id = idByName(e, 'T');
    caret(e, id, 0, 5);
    expect(e.input('text', ' one\u000Btwo\u0001!').ok).toBe(true);
    const p = find(rootOf(e), 'a:p')!;
    expect(elements(p).some((c) => c.kind === 'break')).toBe(true);
    expect(bodyText(e)).toBe('Hello onetwo!');
  });

  test('text with nothing a slide can hold is refused and published', async () => {
    const e = await open(textDeck(para(run('Hello'))));
    caret(e, idByName(e, 'T'), 0, 5);
    const r = e.input('text', '\u0000\u0001');
    expect(r).toEqual({ ok: false, reason: 'errors.invalidText' });
    expect(e.snapshot().lastRejection).toEqual({ reason: 'errors.invalidText' });
  });
});

describe('E14: word deletion and other input types', () => {
  test('deleteWordBackward and deleteWordForward delete one word', async () => {
    const e = await open(textDeck(para(run('one two three'))));
    const id = idByName(e, 'T');
    caret(e, id, 0, 7);
    e.input('deleteWordBackward');
    expect(bodyText(e)).toBe('one  three');
    e.input('deleteWordForward');
    expect(bodyText(e)).toBe('one ');
  });

  test('beforeinput types map to edits, commands, or a typed refusal', () => {
    const calls: unknown[] = [];
    const host: SurfaceInputHost = {
      textInput: (kind, data) => calls.push(['text', kind, data]),
      shortcut: () => false,
      copy: () => false,
      paste: () => false,
      run: (command) => calls.push(['run', command.type]),
      beginHistoryGroup: () => {
        throw new Error('no gesture here');
      },
      refuse: (reason, detail) => calls.push(['refuse', reason, detail]),
    };
    const event = (inputType: string, data: string | null = null) =>
      ({ inputType, data, dataTransfer: null }) as unknown as InputEvent;
    dispatchBeforeInput(host, event('deleteWordBackward'));
    dispatchBeforeInput(host, event('deleteWordForward'));
    dispatchBeforeInput(host, event('historyUndo'));
    dispatchBeforeInput(host, event('formatBold'));
    dispatchBeforeInput(host, event('insertText', 'q'));
    dispatchBeforeInput(host, event('insertTranspose'));
    expect(calls).toEqual([
      ['text', 'deleteWordBackward', undefined],
      ['text', 'deleteWordForward', undefined],
      ['run', 'undo'],
      ['run', 'toggleBold'],
      ['text', 'text', 'q'],
      ['refuse', 'errors.inputUnsupported', 'insertTranspose'],
    ]);
  });
});

describe('E15, E18: speaker notes', () => {
  const notesDeck = (): Uint8Array =>
    buildPptx({ slides: [{ shapes: '', notes: 'Bold note' }, { shapes: '' }] });
  const notesRuns = (e: PptxEditorInstance) => {
    const part = e.session!.info.slides[0].notesPart!;
    const root = xmlPartOf(e.session!.store.package, part)!.root;
    const body = elements(find(root, 'p:spTree')!)
      .map((sp) => child(sp, 'p:txBody'))
      .find((b) => b);
    return { part, runs: elements(find(body!, 'a:p')!).filter((c) => c.kind === 'run') };
  };

  test('editing notes keeps the formatting of the text that stays', async () => {
    const e = await open(notesDeck());
    const { part, runs } = notesRuns(e);
    const rPr = child(runs[0], 'a:rPr')!;
    e.session!.store.transact([
      { op: 'setAttribute', part, node: rPr.id, name: { ns: '', local: 'b' }, value: '1' },
    ]);
    expect(e.exec({ type: 'setNotes', index: 0, text: 'Bold note, more' }).ok).toBe(true);
    expect(notesText(e.session!.store.package, 0)).toBe('Bold note, more');
    for (const r of notesRuns(e).runs) expect(attr(child(r, 'a:rPr'), 'b')).toBe('1');
    // Paragraphs and characters XML cannot hold.
    e.exec({ type: 'setNotes', index: 0, text: 'Bold\u0000 note\r\nsecond' });
    expect(notesText(e.session!.store.package, 0)).toBe('Bold note\nsecond');
  });

  test('an edit inside a notes field keeps the field text', async () => {
    const e = await open(notesDeck());
    const { part, runs } = notesRuns(e);
    const paragraph = find(xmlPartOf(e.session!.store.package, part)!.root, 'a:p')!;
    const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    e.session!.store.transact([
      {
        op: 'insertNode',
        part,
        parent: paragraph.id,
        index: paragraph.children.indexOf(runs[0]) + 1,
        spec: {
          ns: A,
          local: 'fld',
          prefix: 'a',
          attrs: [{ ns: '', local: 'type', value: 'slidenum' }],
          children: [{ ns: A, local: 't', prefix: 'a', children: [{ text: '17' }] }],
        },
      },
    ]);
    expect(notesText(e.session!.store.package, 0)).toBe('Bold note17');
    e.exec({ type: 'setNotes', index: 0, text: 'Bold note1x7' });
    expect(notesText(e.session!.store.package, 0)).toBe('Bold note1x7');
  });

  test('new notes are written as nodes, with any text', async () => {
    const e = await open(notesDeck());
    expect(e.exec({ type: 'setNotes', index: 1, text: '<a:t> & "x"' }).ok).toBe(true);
    expect(notesText(e.session!.store.package, 1)).toBe('<a:t> & "x"');
    expect(readPptx(e.save()!).ok).toBe(true);
  });

  test('two notes edits are two undo steps', async () => {
    const e = await open(notesDeck());
    e.exec({ type: 'setNotes', index: 0, text: 'first' });
    e.exec({ type: 'setNotes', index: 0, text: 'second' });
    e.exec({ type: 'undo' });
    expect(e.snapshot().notes).toBe('first');
  });
});

describe('E12, S7: clipboard', () => {
  const transfer = (data: Record<string, string>) => ({ getData: (t: string) => data[t] ?? '' });

  test('a clipboard without text never pastes the last copy', async () => {
    const e = await open(
      buildPptx({ slides: [{ shapes: shape({ id: 2, name: 'A', x: IN, y: IN, cx: IN, cy: IN }) }] })
    );
    e.setSelection({ kind: 'shapes', nodeIds: [idByName(e, 'A')] });
    const copied = e.copy()!;
    expect(readClipboard(transfer({ 'text/html': '<img src=x>' }), copied)).toBeNull();
    expect(readClipboard(transfer({}), copied)).toBeNull();
  });

  test('a payload above the size cap is refused before parsing', () => {
    let parsed = false;
    const big = 'x'.repeat(MAX_PAYLOAD_CHARS + 1);
    const r = readClipboard(transfer({ [CLIPBOARD_MIME]: big }), null, (raw) => {
      parsed = raw.length <= MAX_PAYLOAD_CHARS;
      return parsed ? {} : null;
    });
    expect(r).toEqual({ refused: true });
    expect(parsed).toBe(false);
  });
});

describe('S8: template values', () => {
  test('characters XML cannot hold are refused, not written', () => {
    expect(() => xmlAttr('a\u0001b', 'name')).toThrow(CommandRefusal);
    expect(xmlAttr('a"<b', 'name')).toBe('a&quot;&lt;b');
  });
});
