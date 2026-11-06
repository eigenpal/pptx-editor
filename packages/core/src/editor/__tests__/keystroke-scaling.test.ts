// A keystroke costs the same in a long deck as in a short one: the work it does per slide or
// per part is incremental, and each incremental answer equals the one a full pass gives.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { textDeck } from '../../../../../scripts/bench/text-deck.ts';
import {
  applyOps,
  find,
  presentationInfo,
  readPptx,
  writePptx,
  xmlPartOf,
  type PptxPackage,
} from '../../store/index.ts';
import { DerivedParts } from '../../store/package/parts-map.ts';
import { namedFamilies } from '../document-session.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { createSlideSummaries } from '../slide-selection.ts';

// A DOM for this file only: the engine test run registers no global DOM.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IN = 914_400;

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

function typeOn(editor: PptxEditorInstance, slide: number, text = 'x'): void {
  editor.goToSlide(slide);
  const lay = editor.session!.layout(slide)!;
  const item = lay.items.find((i) => i.kind === 'shape' && i.element.name === 'Body')!;
  const end = item.kind === 'shape' ? (item.text?.paragraphLengths[0] ?? 0) : 0;
  const at = { paragraph: 0, offset: end };
  editor.setSelection({ kind: 'text', nodeId: item.element.nodeId, anchor: at, focus: at });
  expect(editor.input('text', text).ok).toBe(true);
}

/** A package no cache has seen: the same deck, written and read again. */
function fresh(pkg: PptxPackage): PptxPackage {
  const r = readPptx(writePptx(pkg));
  if (!r.ok) throw new Error('reopen');
  return r.pkg;
}

describe('a keystroke in a long deck', () => {
  test('derives its package from the last one, and changes only its slide part', async () => {
    const editor = await open(textDeck(30));
    const before = editor.session!.store.package;
    typeOn(editor, 12);
    const after = editor.session!.store.package;
    expect(after.parts).toBeInstanceOf(DerivedParts);
    const slide = presentationInfo(after).slides[12]!.partName;
    for (const [name, part] of before.parts)
      if (name !== slide) expect(after.parts.get(name)).toBe(part);
    expect([...after.parts.keys()]).toEqual([...before.parts.keys()]);
    editor.destroy();
  });

  test('the slide summaries equal a full pass, after typing and after a master edit', async () => {
    const editor = await open(textDeck(12));
    const session = editor.session!;
    const full = () => createSlideSummaries()(session, session.store.package).map((s) => s.key);
    const keys = () => editor.snapshot().slides.map((s) => s.key);
    const first = editor.snapshot().slides;
    typeOn(editor, 4);
    typeOn(editor, 4);
    const typed = editor.snapshot().slides;
    expect(keys()).toEqual(full());
    typed.forEach((s, i) => (i === 4 ? expect(s).not.toBe(first[i]!) : expect(s).toBe(first[i]!)));
    // A master edit changes every slide that uses the master.
    const pkg = session.store.package;
    const master = presentationInfo(pkg).slides[0]!.masterPart!;
    const off = find(xmlPartOf(pkg, master)!.root, 'a:off')!;
    const r = session.store.transact([
      { op: 'setAttribute', part: master, node: off.id, name: { ns: '', local: 'x' }, value: '9' },
    ]);
    expect(r.ok).toBe(true);
    const moved = editor.snapshot().slides;
    expect(keys()).toEqual(full());
    moved.forEach((s, i) => expect(s).not.toBe(typed[i]!));
    // Undo goes back to the keys of the typed deck.
    editor.exec({ type: 'undo' });
    expect(keys()).toEqual(typed.map((s) => s.key));
    editor.destroy();
  });

  test('the named families equal a full pass as fonts come, go, and change spelling', () => {
    const slide = (face: string) => ({
      shapes: shape({
        id: 2,
        name: 'Body',
        x: IN,
        y: IN,
        cx: 4 * IN,
        cy: IN,
        text: para(run('Text', '', `<a:latin typeface="${face}"/>`)),
      }),
    });
    const r = readPptx(
      buildPptx({ slides: [slide('Liberation Serif'), slide('Liberation Mono'), slide('Carlito')] })
    );
    if (!r.ok) throw new Error('deck');
    let pkg = r.pkg;
    const latinOf = (p: PptxPackage, index: number) => {
      const part = presentationInfo(p).slides[index]!.partName;
      return { part, el: find(xmlPartOf(p, part)!.root, 'a:latin')! };
    };
    const setFace = (index: number, face: string) => {
      const { part, el } = latinOf(pkg, index);
      const out = applyOps(pkg, [
        { op: 'setAttribute', part, node: el.id, name: { ns: '', local: 'typeface' }, value: face },
      ]);
      if (!out.ok) throw new Error(JSON.stringify(out.rejection));
      pkg = out.pkg;
    };
    expect(namedFamilies(pkg)).toEqual(namedFamilies(fresh(pkg)));
    for (const [index, face] of [
      [1, 'Liberation Sans'],
      [0, 'Liberation Sans'],
      [1, 'liberation sans'],
      [0, 'Carlito'],
      [1, 'Carlito'],
    ] as const) {
      const prev = pkg;
      setFace(index, face);
      const expected = namedFamilies(fresh(pkg));
      // The last package read is the one before the edit: the count changes by one part.
      namedFamilies(prev);
      expect(namedFamilies(pkg)).toEqual(expected);
    }
  });

  test('typing reads no computed style for the fit zoom; a new slide size does', async () => {
    const g = globalThis as { ResizeObserver?: unknown };
    g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
    const editor = await open(textDeck(6));
    const host = doc.createElement('div') as unknown as HTMLElement;
    doc.body.append(host as unknown as Node);
    editor.attach(host);
    const view = doc.defaultView as unknown as { getComputedStyle: (e: Element) => unknown };
    const real = view.getComputedStyle.bind(view);
    let reads = 0;
    view.getComputedStyle = (e: Element) => ((reads += 1), real(e));
    typeOn(editor, 2);
    typeOn(editor, 2);
    expect(reads).toBe(0);
    const pkg = editor.session!.store.package;
    const info = presentationInfo(pkg);
    const size = find(xmlPartOf(pkg, info.presentationPart)!.root, 'p:sldSz')!;
    const r = editor.session!.store.transact([
      {
        op: 'setAttribute',
        part: info.presentationPart,
        node: size.id,
        name: { ns: '', local: 'cx' },
        value: String(info.width * 2),
      },
    ]);
    expect(r.ok).toBe(true);
    expect(reads).toBeGreaterThan(0);
    view.getComputedStyle = real;
    editor.destroy();
  });
});
