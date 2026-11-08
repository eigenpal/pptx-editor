// An IME composition on one peer while another peer edits around it: the composition keeps
// its place by its paragraph, its text goes in only there, and both peers hold one deck.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import * as Y from 'yjs';
import type { PptxEditorInstance } from '@pptx-editor.dev/core';
import { xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, cell, para, run, shape, table } from '../../../../scripts/fixtures/builder.ts';
import { differingParts, link, replica } from './convergence-harness.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IN = 914_400;
const COMPOSING = 'pptx-surface__input--composing';

/** A deck with shape `T` (the paragraphs `text`) and a 2 x 1 table. */
const deck = (text: string) => () =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({ id: 2, name: 'T', x: IN, y: IN, cx: 6 * IN, cy: 2 * IN, text }) +
          table({
            id: 3,
            x: IN,
            y: 4 * IN,
            cx: 2 * IN,
            cy: 2 * 370_840,
            columns: [2 * IN],
            tblPr: '<a:tblPr/>',
            rows: ['abc', 'def'].map((t) => `<a:tr h="370840">${cell(t)}</a:tr>`).join(''),
          }),
      },
    ],
  });

type Node = { type: 'text'; value: string } | { type: 'element'; local: string; children: Node[] };
/** The text of each paragraph on the first slide, in document order. */
function paragraphs(e: PptxEditorInstance): string[] {
  const part = e.session!.info.slides[0]!.partName;
  const text = (n: Node): string =>
    n.type === 'text'
      ? n.value
      : n.local === 't'
        ? n.children.map(text).join('')
        : n.children.map(text).join('');
  const walk = (n: Node): string[] =>
    n.type === 'text' ? [] : n.local === 'p' ? [text(n)] : n.children.flatMap(walk);
  return walk(xmlPartOf(e.session!.store.package, part)!.root as unknown as Node);
}

/** Two linked peers; peer A has a surface, and composes in its input. */
async function peers(text: string) {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const bytes = deck(text)();
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  link(docA, docB);
  const a = await replica(docA, 'a', bytes);
  const b = await replica(docB, 'b', bytes);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node as never);
  a.attach(host);
  const input = host.querySelector('textarea') as HTMLTextAreaElement;
  const fire = (type: string, data = '') => {
    const event = new win.CompositionEvent(type, { bubbles: true }) as unknown as Event;
    Object.defineProperty(event, 'data', { value: data });
    input.dispatchEvent(event);
  };
  /** Peer A composes `data` at its caret, and the composition stays open. */
  const compose = (data: string) => {
    fire('compositionstart');
    input.value += data;
    input.setSelectionRange(input.value.length, input.value.length);
  };
  /** The composition ends with `data`, as the field shows it. */
  const commit = (data: string) => fire('compositionend', data);
  const composing = () => input.classList.contains(COMPOSING);
  const done = () => {
    expect(differingParts(a.session!.store.package, b.session!.store.package)).toEqual([]);
    a.destroy();
    b.destroy();
  };
  /** A composition starts over the input selection, as a keyboard that composes a word again. */
  const start = () => fire('compositionstart');
  return { a, b, input, start, compose, commit, composing, done };
}

const nodeId = (e: PptxEditorInstance, name: string) =>
  e.snapshot().shapes!.find((s) => s.name === name)!.nodeId;
const tableId = (e: PptxEditorInstance) => e.snapshot().shapes!.find((s) => s.name !== 'T')!.nodeId;
const caret = (
  e: PptxEditorInstance,
  id: string,
  paragraph: number,
  offset: number,
  cellAt?: { row: number; column: number }
) => {
  const at = { paragraph, offset };
  e.setSelection({
    kind: 'text',
    nodeId: id,
    ...(cellAt ? { cell: cellAt } : {}),
    anchor: at,
    focus: at,
  });
};

describe('a composition while a peer edits', () => {
  test('a peer Enter above the paragraph keeps the composition at its paragraph', async () => {
    const { a, b, compose, commit, composing, done } = await peers(
      para(run('one')) + para(run('two'))
    );
    caret(a, nodeId(a, 'T'), 1, 3);
    compose('日本');
    // Peer B presses Enter at the end of the first paragraph: a new paragraph above.
    caret(b, nodeId(b, 'T'), 0, 3);
    expect(b.input('paragraph').ok).toBe(true);
    expect(composing()).toBe(true);
    commit('日本');
    expect(paragraphs(a).slice(0, 3)).toEqual(['one', '', 'two日本']);
    expect(a.snapshot().lastRejection).toBeNull();
    done();
  });

  test('a peer row inserted above the cell keeps the composition in its cell', async () => {
    const { a, b, compose, commit, composing, done } = await peers(para(run('x')));
    caret(a, tableId(a), 0, 3, { row: 1, column: 0 });
    compose('日本');
    caret(b, tableId(b), 0, 0, { row: 0, column: 0 });
    expect(b.exec({ type: 'insertRow', where: 'above' }).ok).toBe(true);
    expect(composing()).toBe(true);
    commit('日本');
    expect(paragraphs(a)).toEqual(['x', '', 'abc', 'def日本']);
    done();
  });

  test('a peer delete before the caret in the paragraph moves the place with it', async () => {
    const { a, b, compose, commit, composing, done } = await peers(para(run('Hello world')));
    caret(a, nodeId(a, 'T'), 0, 11);
    compose('日本');
    const at = (offset: number) => ({ paragraph: 0, offset });
    b.setSelection({ kind: 'text', nodeId: nodeId(b, 'T'), anchor: at(0), focus: at(6) });
    expect(b.input('deleteBackward').ok).toBe(true);
    expect(composing()).toBe(true);
    commit('日本');
    expect(paragraphs(a)[0]).toBe('world日本');
    done();
  });

  test('a peer Enter before the caret takes the place to the new paragraph', async () => {
    const { a, b, compose, commit, composing, done } = await peers(para(run('Hello world')));
    caret(a, nodeId(a, 'T'), 0, 11);
    compose('日本');
    // Peer B splits the paragraph before "world": the place goes with "world".
    caret(b, nodeId(b, 'T'), 0, 5);
    expect(b.input('paragraph').ok).toBe(true);
    expect(composing()).toBe(true);
    commit('日本');
    expect(paragraphs(a).slice(0, 2)).toEqual(['Hello', ' world日本']);
    expect(a.snapshot().lastRejection).toBeNull();
    done();
  });

  test('a peer Enter after the caret keeps the place', async () => {
    const { a, b, compose, commit, done } = await peers(para(run('Hello world, bye')));
    caret(a, nodeId(a, 'T'), 0, 11);
    compose('日本');
    caret(b, nodeId(b, 'T'), 0, 11);
    expect(b.input('paragraph').ok).toBe(true);
    commit('日本');
    expect(paragraphs(a).slice(0, 2)).toEqual(['Hello world日本', ', bye']);
    done();
  });

  test('a peer Backspace that merges the paragraph keeps the place in the merged text', async () => {
    const { a, b, compose, commit, composing, done } = await peers(
      para(run('one')) + para(run('two'))
    );
    caret(a, nodeId(a, 'T'), 1, 3);
    compose('日本');
    // Peer B joins the paragraph to the one above: its element id is gone.
    caret(b, nodeId(b, 'T'), 1, 0);
    expect(b.input('deleteBackward').ok).toBe(true);
    expect(composing()).toBe(true);
    commit('日本');
    expect(paragraphs(a)[0]).toBe('onetwo日本');
    expect(a.snapshot().lastRejection).toBeNull();
    done();
  });

  test('a peer typing before the caret in the paragraph moves the place with it', async () => {
    const { a, b, compose, commit, done } = await peers(para(run('Hello world')));
    caret(a, nodeId(a, 'T'), 0, 11);
    compose('日本');
    caret(b, nodeId(b, 'T'), 0, 0);
    expect(b.input('text', 'X').ok).toBe(true);
    commit('日本');
    expect(paragraphs(a)[0]).toBe('XHello world日本');
    done();
  });

  test('a peer change inside a recomposed word refuses the word change', async () => {
    const { a, b, input, start, commit, composing, done } = await peers(para(run('hello wrold')));
    caret(a, nodeId(a, 'T'), 0, 11);
    // A keyboard composes "wrold" again.
    input.setSelectionRange(6, 11);
    start();
    const at = (offset: number) => ({ paragraph: 0, offset });
    b.setSelection({ kind: 'text', nodeId: nodeId(b, 'T'), anchor: at(8), focus: at(9) });
    expect(b.input('text', 'X').ok).toBe(true);
    expect(composing()).toBe(true);
    input.value = 'hello world';
    input.setSelectionRange(11, 11);
    commit('world');
    // The text before the place is no longer the word: the change goes in nowhere.
    expect(paragraphs(a)[0]).toBe('hello wrXld');
    expect(a.snapshot().lastRejection?.reason).toBe('errors.inputUnsupported');
    done();
  });

  test('a peer delete of the paragraph ends the composition, and its text goes nowhere', async () => {
    const { a, b, compose, composing, done } = await peers(para(run('one')) + para(run('two')));
    caret(a, nodeId(a, 'T'), 1, 3);
    compose('日本');
    const one = { paragraph: 0, offset: 3 };
    const two = { paragraph: 1, offset: 3 };
    b.setSelection({ kind: 'text', nodeId: nodeId(b, 'T'), anchor: one, focus: two });
    expect(b.input('deleteBackward').ok).toBe(true);
    // The place is gone: the composition ends, and the text is refused with a reason.
    expect(composing()).toBe(false);
    expect(paragraphs(a).slice(0, 1)).toEqual(['one']);
    expect(a.snapshot().lastRejection?.reason).toBe('errors.inputUnsupported');
    done();
  });
});
