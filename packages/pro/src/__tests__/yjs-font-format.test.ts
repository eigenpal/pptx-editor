// The Font group's commands (`setFontFormat`, the toggles, and a format armed at a caret)
// through two editors on two linked Y.Docs, as AGENTS.md "Collaboration verification" asks:
// concurrent edits with the link paused, undo while a peer edits, and a deletion while a peer
// edits the same shape. The replicas converge, before and after save and reopen. An armed
// format is session state, so it never reaches a peer until its text is typed. Concurrent
// inserts of the same run child keep one. A run split while a peer types into that run is in
// `yjs-text-structure.test.ts`.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  attr,
  canonicalPackageFingerprint,
  child,
  elements,
  find,
  readPptx,
  textOf,
  writePptx,
  xmlPartOf,
  type Element,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

const IN = 914_400;
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          name: 'T',
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: 2 * IN,
          text: para(run('Hello') + run(' world')),
        }),
      },
    ],
  });

/** Two docs and a sync link that can pause, to make edits concurrent. */
function link(a: Y.Doc, b: Y.Doc) {
  const queue: [Y.Doc, Uint8Array][] = [];
  let paused = false;
  const send = (to: Y.Doc) => (update: Uint8Array, origin: unknown) => {
    if (origin === 'link') return;
    if (paused) queue.push([to, update]);
    else Y.applyUpdate(to, update, 'link');
  };
  a.on('update', send(b));
  b.on('update', send(a));
  return {
    pause: () => (paused = true),
    flush: (reverse = false) => {
      paused = false;
      const updates = queue.splice(0);
      if (reverse) updates.reverse();
      for (const [to, update] of updates) {
        Y.applyUpdate(to, update, 'link');
        if (reverse) Y.applyUpdate(to, update, 'link');
      }
    },
  };
}

async function replicas(swap = false) {
  const docs = [new Y.Doc(), new Y.Doc()] as const;
  docs[0].clientID = swap ? 202 : 101;
  docs[1].clientID = swap ? 101 : 202;
  const sync = link(docs[0], docs[1]);
  const editors: PptxEditorInstance[] = [];
  for (const [i, doc] of docs.entries()) {
    const session = createYjsCollaborationSession({
      doc,
      documentId: 'deck',
      identity: { actorId: `p${i}`, name: `p${i}` },
    });
    const editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [collaborationModule({ session })],
    });
    expect((await editor.load(deck())).ok).toBe(true);
    editors.push(editor);
  }
  return { a: editors[0]!, b: editors[1]!, sync, docs };
}

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
/** Both replicas match, live and after save and reopen. */
function converged(a: PptxEditorInstance, b: PptxEditorInstance): void {
  expect(fingerprint(a)).toBe(fingerprint(b));
  const saved = (e: PptxEditorInstance) => {
    const reopened = readPptx(writePptx(e.session!.store.package));
    if (!reopened.ok) throw new Error('reopen failed');
    return canonicalPackageFingerprint(reopened.pkg);
  };
  expect(saved(a)).toBe(saved(b));
}
const root = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
const shapeOf = (e: PptxEditorInstance): Element | undefined =>
  elements(find(root(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T');
const runs = (e: PptxEditorInstance): Element[] =>
  elements(elements(find(shapeOf(e)!, 'p:txBody')!).find((c) => c.local === 'p')!).filter(
    (c) => c.local === 'r'
  );
const texts = (e: PptxEditorInstance): string[] => runs(e).map((r) => textOf(child(r, 'a:t')));
const rPr = (r: Element | undefined, name: string) => attr(child(r!, 'a:rPr'), name);
const select = (e: PptxEditorInstance, from: number, to = from) =>
  e.setSelection({
    kind: 'text',
    nodeId: shapeOf(e)!.id,
    anchor: { paragraph: 0, offset: from },
    focus: { paragraph: 0, offset: to },
  });

describe('font formatting under collaboration', () => {
  test('two peers format different runs at once: both survive', async () => {
    const { a, b, sync } = await replicas();
    select(a, 0, 5);
    select(b, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'setFontFormat', bold: true, caps: 'all' }).ok).toBe(true);
    expect(b.exec({ type: 'setFontFormat', fontSizePt: 30, characterSpacingPt: 2 }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    const [hello] = runs(a);
    const world = runs(a).find((r) => textOf(child(r, 'a:t')) === 'world');
    expect([rPr(hello, 'b'), rPr(hello, 'cap')]).toEqual(['1', 'all']);
    expect([rPr(world, 'sz'), rPr(world, 'spc')]).toEqual(['3000', '200']);
  });

  test('two peers set one field of one run: one value', async () => {
    const { a, b, sync } = await replicas();
    select(a, 0, 5);
    select(b, 0, 5);
    sync.pause();
    expect(a.exec({ type: 'setFontFormat', baseline: 30 }).ok).toBe(true);
    expect(b.exec({ type: 'setFontFormat', baseline: -25 }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    expect(['30000', '-25000']).toContain(rPr(runs(a)[0], 'baseline') ?? '');
  });

  test('an armed format never reaches the peer; its typed text does, with the format', async () => {
    const { a, b } = await replicas();
    select(a, 11);
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(a.snapshot().format?.bold).toBe(true);
    // Nothing reached the room: the peer's caret there is not bold, and its text is not.
    converged(a, b);
    select(b, 11);
    expect(b.snapshot().format?.bold).toBe(false);
    expect(b.input('text', '!').ok).toBe(true);
    expect(texts(a).join('')).toBe('Hello world!');
    expect(runs(a).some((r) => rPr(r, 'b') === '1')).toBe(false);
    // The arm moves with the caret, past the peer's text, and the text typed next takes it.
    const caret = a.snapshot().selection;
    expect(caret.kind === 'text' && caret.focus.offset).toBe(12);
    expect(a.snapshot().format?.bold).toBe(true);
    e2eType(a, 'X');
    converged(a, b);
    expect(texts(b).join('')).toBe('Hello world!X');
    expect(
      rPr(
        runs(b).find((r) => textOf(child(r, 'a:t')) === 'X'),
        'b'
      )
    ).toBe('1');
  });

  test("undo of armed text reverts it and its format, and keeps the peer's edit", async () => {
    const { a, b, sync } = await replicas();
    select(a, 11);
    a.exec({ type: 'toggleItalic' });
    select(b, 0, 5);
    sync.pause();
    expect(a.input('text', 'Z').ok).toBe(true);
    expect(b.exec({ type: 'toggleUnderline' }).ok).toBe(true);
    sync.flush();
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    converged(a, b);
    expect(texts(a).join('')).toBe('Hello world');
    expect(rPr(runs(a)[0], 'u')).toBe('sng');
    expect(runs(a).some((r) => rPr(r, 'i') === '1')).toBe(false);
    // Redo brings the text and its format back on both replicas.
    expect(a.exec({ type: 'redo' }).ok).toBe(true);
    converged(a, b);
    expect(runs(b).find((r) => textOf(child(r, 'a:t')) === 'Z')).toBeDefined();
    expect(
      rPr(
        runs(b).find((r) => textOf(child(r, 'a:t')) === 'Z'),
        'i'
      )
    ).toBe('1');
  });

  test('a peer deletes the shape while the other formats it: both end without it', async () => {
    const { a, b, sync } = await replicas();
    select(a, 0, 5);
    b.setSelection({ kind: 'shapes', nodeIds: [shapeOf(b)!.id] });
    sync.pause();
    expect(a.exec({ type: 'setFontFormat', underline: 'double', kerningPt: 12 }).ok).toBe(true);
    expect(b.exec({ type: 'delete' }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    expect(shapeOf(a)).toBeUndefined();
  });

  test('one peer sets an underline color and the other a text color: both survive', async () => {
    const { a, b, sync } = await replicas();
    select(a, 0, 5);
    select(b, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'setFontFormat', underlineColor: { theme: 'accent2' } }).ok).toBe(true);
    expect(b.exec({ type: 'setFontFormat', color: { theme: 'accent1' } }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    const [hello] = runs(a);
    const world = runs(a).find((r) => textOf(child(r, 'a:t')) === 'world');
    expect(child(child(hello!, 'a:rPr')!, 'a:uFill')).toBeDefined();
    expect(child(child(world!, 'a:rPr')!, 'a:solidFill')).toBeDefined();
  });

  test("undo of an underline color keeps the peer's edit", async () => {
    const { a, b, sync } = await replicas();
    select(a, 0, 5);
    select(b, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'setFontFormat', underlineColor: { theme: 'accent2' } }).ok).toBe(true);
    expect(b.exec({ type: 'setFontFormat', bold: true }).ok).toBe(true);
    sync.flush();
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    converged(a, b);
    expect(child(child(runs(a)[0]!, 'a:rPr')!, 'a:uFill')).toBeUndefined();
    const world = runs(a).find((r) => textOf(child(r, 'a:t')) === 'world');
    expect(rPr(world, 'b')).toBe('1');
  });

  test('two peers set an underline color on one run at once: one a:uFill', async () => {
    const { a, b, sync } = await replicas();
    select(a, 0, 5);
    select(b, 0, 5);
    sync.pause();
    expect(a.exec({ type: 'setFontFormat', underlineColor: { theme: 'accent2' } }).ok).toBe(true);
    expect(b.exec({ type: 'setFontFormat', underlineColor: { theme: 'accent1' } }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    const hello = child(runs(a)[0]!, 'a:rPr')!;
    expect(
      elements(hello).filter((c) => c.local === 'uFill' || c.local === 'uFillTx')
    ).toHaveLength(1);
  });

  test('two peers set a font family on a run with no a:latin: one a:latin', async () => {
    const { a, b, sync } = await replicas();
    select(a, 0, 5);
    select(b, 0, 5);
    sync.pause();
    expect(a.exec({ type: 'setFontFormat', fontFamily: 'Arial' }).ok).toBe(true);
    expect(b.exec({ type: 'setFontFormat', fontFamily: 'Georgia' }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    const hello = child(runs(a)[0]!, 'a:rPr')!;
    expect(elements(hello).filter((c) => c.local === 'latin')).toHaveLength(1);
  });

  test('Change Case while a peer types into the run: the typed text stays in place', async () => {
    const { a, b, sync } = await replicas();
    a.setSelection({ kind: 'shapes', nodeIds: [shapeOf(a)!.id] });
    select(b, 3);
    sync.pause();
    expect(a.exec({ type: 'changeCase', case: 'upper' }).ok).toBe(true);
    e2eType(b, 'q');
    sync.flush();
    converged(a, b);
    // The typed text stays at its place, or one character after it: Yjs orders two inserts
    // in one gap (the new "L" and the "q") by client id.
    const text = texts(a).join('');
    expect(text.replace('q', '')).toBe('HELLO WORLD');
    expect([3, 4]).toContain(text.indexOf('q'));
  });

  test('a retyped emoji and a Change Case of a non-BMP letter reach the peer whole', async () => {
    const { a, b } = await replicas();
    select(a, 0);
    e2eType(a, '😀');
    select(a, 0, 2);
    e2eType(a, '😁');
    converged(a, b);
    expect(texts(b).join('')).toBe('😁Hello world');
    select(a, 0, 2);
    e2eType(a, '𐐨');
    a.setSelection({ kind: 'shapes', nodeIds: [shapeOf(a)!.id] });
    expect(a.exec({ type: 'changeCase', case: 'upper' }).ok).toBe(true);
    converged(a, b);
    expect(texts(b).join('')).toBe('𐐀HELLO WORLD');
  });

  test('two peers change different letters of one run at once: both changes stay', async () => {
    const { a, b, sync } = await replicas();
    a.setSelection({ kind: 'shapes', nodeIds: [shapeOf(a)!.id] });
    b.setSelection({ kind: 'shapes', nodeIds: [shapeOf(b)!.id] });
    sync.pause();
    expect(a.exec({ type: 'changeCase', case: 'upper' }).ok).toBe(true);
    expect(b.exec({ type: 'changeCase', case: 'lower' }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    const text = texts(a).join('');
    expect(text).toHaveLength('Hello world'.length);
    expect(text.toLowerCase()).toBe('hello world');
  });

  test('undo of Change Case keeps the text a peer typed', async () => {
    const { a, b } = await replicas();
    a.setSelection({ kind: 'shapes', nodeIds: [shapeOf(a)!.id] });
    expect(a.exec({ type: 'changeCase', case: 'upper' }).ok).toBe(true);
    select(b, 3);
    e2eType(b, 'q');
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    converged(a, b);
    // The typed text survives at its place: after the third letter.
    expect(texts(a).join('')).toBe('Helqlo world');
    expect(a.exec({ type: 'redo' }).ok).toBe(true);
    converged(a, b);
    expect(texts(a).join('')).toBe('HELqLO WORLD');
  });
  for (const swap of [false, true]) {
    test(`concurrent case changes preserve text through both undo histories (${swap})`, async () => {
      const { a, b, sync } = await replicas(swap);
      try {
        for (const e of [a, b]) e.setSelection({ kind: 'shapes', nodeIds: [shapeOf(e)!.id] });
        sync.pause();
        for (const e of [a, b]) expect(e.exec({ type: 'changeCase', case: 'upper' }).ok).toBe(true);
        sync.flush(swap);
        converged(a, b);
        expect(texts(a).join('')).toBe('HELLO WORLD');
        expect(a.exec({ type: 'undo' }).ok).toBe(true);
        converged(a, b);
        expect(texts(a).join('')).toBe('HELLO WORLD');
        expect(b.exec({ type: 'undo' }).ok).toBe(true);
        converged(a, b);
        expect(texts(a).join('')).toBe('Hello world');
        for (const e of [b, a]) {
          expect(e.exec({ type: 'redo' }).ok).toBe(true);
          converged(a, b);
          expect(texts(a).join('')).toBe('HELLO WORLD');
        }
      } finally {
        a.destroy();
        b.destroy();
      }
    });
    for (const concurrent of [false, true])
      test(`case changes follow formatted copies and undo (${swap}, ${concurrent})`, async () => {
        const { a, b, sync } = await replicas(swap);
        try {
          if (concurrent) sync.pause();
          a.setSelection({ kind: 'shapes', nodeIds: [shapeOf(a)!.id] });
          expect(a.exec({ type: 'changeCase', case: 'upper' }).ok).toBe(true);
          select(b, 2, 4);
          expect(b.exec({ type: 'toggleBold' }).ok).toBe(true);
          sync.flush();
          converged(a, b);
          expect(texts(a).join('')).toBe('HELLO WORLD');
          expect(a.exec({ type: 'undo' }).ok).toBe(true);
          converged(a, b);
          expect(texts(a).join('')).toBe('Hello world');
          expect(a.exec({ type: 'redo' }).ok).toBe(true);
          converged(a, b);
          expect(texts(a).join('')).toBe('HELLO WORLD');
          expect(b.exec({ type: 'undo' }).ok).toBe(true);
          converged(a, b);
          expect(texts(a).join('')).toBe('HELLO WORLD');
        } finally {
          a.destroy();
          b.destroy();
        }
      });
  }
  test('concurrent Unicode case changes preserve complete characters and undo', async () => {
    const { a, b, sync } = await replicas();
    try {
      select(a, 0, 11);
      e2eType(a, '𐐨 café ß 😀');
      sync.pause();
      for (const e of [a, b]) {
        e.setSelection({ kind: 'shapes', nodeIds: [shapeOf(e)!.id] });
        expect(e.exec({ type: 'changeCase', case: 'upper' }).ok).toBe(true);
      }
      sync.flush(true);
      converged(a, b);
      expect(texts(a).join('')).toBe('𐐀 CAFÉ ß 😀');
      for (const e of [a, b]) expect(e.exec({ type: 'undo' }).ok).toBe(true);
      converged(a, b);
      expect(texts(a).join('')).toBe('𐐨 café ß 😀');
    } finally {
      a.destroy();
      b.destroy();
    }
  });
  test('typing a differently cased replacement creates a new character identity', async () => {
    const { a, b } = await replicas();
    try {
      select(a, 0);
      e2eType(a, 'Č');
      select(b, 0, 1);
      e2eType(b, 'č');
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      converged(a, b);
      expect(texts(a).join('')).toBe('čHello world');
    } finally {
      a.destroy();
      b.destroy();
    }
  });
  for (const edit of ['paragraph', 'deleteBackward'] as const)
    test(`case changes and concurrent ${edit} retain character identities`, async () => {
      const { a, b, sync } = await replicas();
      try {
        sync.pause();
        a.setSelection({ kind: 'shapes', nodeIds: [shapeOf(a)!.id] });
        expect(a.exec({ type: 'changeCase', case: 'upper' }).ok).toBe(true);
        select(b, 3, edit === 'paragraph' ? 3 : 5);
        expect(b.input(edit).ok).toBe(true);
        sync.flush(true);
        converged(a, b);
        expect(textOf(find(shapeOf(a)!, 'p:txBody'))).toBe(
          edit === 'paragraph' ? 'HELLO WORLD' : 'HEL WORLD'
        );
        expect(a.exec({ type: 'undo' }).ok).toBe(true);
        converged(a, b);
        expect(textOf(find(shapeOf(a)!, 'p:txBody'))).toBe(
          edit === 'paragraph' ? 'Hello world' : 'Hel world'
        );
        expect(b.exec({ type: 'undo' }).ok).toBe(true);
        converged(a, b);
        expect(textOf(find(shapeOf(a)!, 'p:txBody'))).toBe('Hello world');
      } finally {
        a.destroy();
        b.destroy();
      }
    });
  test('saved room updates preserve case choices on a fresh replica', async () => {
    const { a, b, docs } = await replicas();
    const restored = new Y.Doc();
    const session = createYjsCollaborationSession({
      doc: restored,
      documentId: 'deck',
      identity: { actorId: 'restored', name: 'Restored' },
    });
    const c = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [collaborationModule({ session })],
    });
    try {
      a.setSelection({ kind: 'shapes', nodeIds: [shapeOf(a)!.id] });
      expect(a.exec({ type: 'changeCase', case: 'upper' }).ok).toBe(true);
      select(b, 2, 4);
      expect(b.exec({ type: 'toggleBold' }).ok).toBe(true);
      Y.applyUpdate(restored, Y.encodeStateAsUpdate(docs[0]));
      expect((await c.load(deck())).ok).toBe(true);
      converged(a, c);
      expect(texts(c).join('')).toBe('HELLO WORLD');
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      Y.applyUpdate(restored, Y.encodeStateAsUpdate(docs[0], Y.encodeStateVector(restored)));
      converged(a, c);
      expect(texts(c).join('')).toBe('Hello world');
    } finally {
      a.destroy();
      b.destroy();
      c.destroy();
      session.destroy();
      restored.destroy();
    }
  });
  for (const offset of [0, 1, 2, 3, 4, 5, 8, 11])
    test(`peer formatting preserves the caret at ${offset}`, async () => {
      const { a, b } = await replicas();
      try {
        select(b, offset);
        select(a, 2, 4);
        expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
        e2eType(b, '!');
        converged(a, b);
        expect(texts(a).join('')).toBe(
          'Hello world'.slice(0, offset) + '!' + 'Hello world'.slice(offset)
        );
      } finally {
        a.destroy();
        b.destroy();
      }
    });
  for (const [from, to] of [
    [1, 4],
    [4, 1],
    [2, 5],
    [5, 2],
  ])
    test(`peer formatting preserves selection ${from}..${to}`, async () => {
      const { a, b } = await replicas();
      try {
        select(b, from!, to!);
        select(a, 2, 4);
        expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
        e2eType(b, '!');
        converged(a, b);
        expect(texts(a).join('')).toBe(
          'Hello world'.slice(0, Math.min(from!, to!)) +
            '!' +
            'Hello world'.slice(Math.max(from!, to!))
        );
      } finally {
        a.destroy();
        b.destroy();
      }
    });
  test('peer formatting undo and redo preserve the caret', async () => {
    const { a, b } = await replicas();
    try {
      select(b, 3);
      select(a, 2, 4);
      for (const type of ['toggleBold', 'undo', 'redo'] as const)
        expect(a.exec({ type }).ok).toBe(true);
      e2eType(b, '!');
      converged(a, b);
      expect(texts(a).join('')).toBe('Hel!lo world');
    } finally {
      a.destroy();
      b.destroy();
    }
  });
});

/** Type at the caret as the surface does. */
function e2eType(e: PptxEditorInstance, text: string): void {
  expect(e.input('text', text).ok).toBe(true);
}
