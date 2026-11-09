// An IME composition on one peer while another peer edits around it. An edit after the place,
// or in another paragraph, keeps the composition at its place. Any other edit ends it, and its
// text is refused with a reason: never put anywhere else. Both peers hold one deck.

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

/** A deck with shape `T` (the paragraphs `text`) and a table with one column of `cells`. */
const deck = (text: string, cells = ['abc', 'def']) =>
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
            cy: cells.length * 370_840,
            columns: [2 * IN],
            tblPr: '<a:tblPr/>',
            rows: cells.map((t) => `<a:tr h="370840">${cell(t)}</a:tr>`).join(''),
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
/** The paragraphs of shape `T`: all but the two of the table. */
const shapeText = (e: PptxEditorInstance): string[] => paragraphs(e).slice(0, -2);

/** Two linked peers; peer A has a surface, and composes in its input. */
async function peers(text: string, cells?: string[]) {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const bytes = deck(text, cells);
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

  test('a peer Enter after the caret keeps the place', async () => {
    const { a, b, compose, commit, done } = await peers(para(run('Hello world, bye')));
    caret(a, nodeId(a, 'T'), 0, 11);
    compose('日本');
    caret(b, nodeId(b, 'T'), 0, 11);
    expect(b.input('paragraph').ok).toBe(true);
    commit('日本');
    expect(shapeText(a)).toEqual(['Hello world日本', ', bye']);
    done();
  });

  test('a peer typing after the place in its paragraph keeps the place', async () => {
    const { a, b, compose, commit, composing, done } = await peers(para(run('Hello world')));
    caret(a, nodeId(a, 'T'), 0, 5);
    compose('日本');
    caret(b, nodeId(b, 'T'), 0, 11);
    expect(b.input('text', '!').ok).toBe(true);
    expect(composing()).toBe(true);
    commit('日本');
    expect(shapeText(a)).toEqual(['Hello日本 world!']);
    expect(a.snapshot().lastRejection).toBeNull();
    done();
  });

  test('a peer edit in another paragraph keeps the place', async () => {
    const { a, b, compose, commit, done } = await peers(para(run('one')) + para(run('two')));
    caret(a, nodeId(a, 'T'), 1, 3);
    compose('日本');
    caret(b, nodeId(b, 'T'), 0, 0);
    expect(b.input('text', 'X').ok).toBe(true);
    commit('日本');
    expect(shapeText(a)).toEqual(['Xone', 'two日本']);
    done();
  });

  test('a peer Enter at the place in a cell keeps the text before it', async () => {
    const { a, b, compose, commit, done } = await peers(para(run('x')), ['ab', 'ab']);
    caret(a, tableId(a), 0, 2, { row: 1, column: 0 });
    compose('日本');
    caret(b, tableId(b), 0, 2, { row: 1, column: 0 });
    expect(b.input('paragraph').ok).toBe(true);
    commit('日本');
    expect(paragraphs(a)).toEqual(['x', 'ab', 'ab日本', '']);
    done();
  });
});

type Peer = (e: PptxEditorInstance) => void;
/** The caret at `offset` of paragraph `paragraph` in shape `T`. */
const at = (paragraph: number, offset: number) => (e: PptxEditorInstance) =>
  caret(e, nodeId(e, 'T'), paragraph, offset);
/** A range in shape `T` from `[paragraph, offset]` to `[paragraph, offset]`. */
const range =
  ([p0, o0]: [number, number], [p1, o1]: [number, number]) =>
  (e: PptxEditorInstance) =>
    e.setSelection({
      kind: 'text',
      nodeId: nodeId(e, 'T'),
      anchor: { paragraph: p0, offset: o0 },
      focus: { paragraph: p1, offset: o1 },
    });
const then =
  (...steps: Peer[]): Peer =>
  (e) => {
    for (const step of steps) step(e);
  };
const typed = (text: string) => (e: PptxEditorInstance) =>
  expect(e.input('text', text).ok).toBe(true);
const key = (kind: 'paragraph' | 'deleteBackward') => (e: PptxEditorInstance) =>
  expect(e.input(kind).ok).toBe(true);
const history = (type: 'undo' | 'redo') => (e: PptxEditorInstance) =>
  expect(e.exec({ type }).ok).toBe(true);

/** Peer edits before or over the place: each ends the composition and refuses its text. */
const refusals: {
  readonly name: string;
  readonly paragraphs: string[];
  readonly place: [number, number];
  readonly first?: Peer;
  readonly peer: Peer;
  readonly after: string[];
}[] = [
  {
    name: 'a peer Enter before the caret',
    paragraphs: ['Hello world'],
    place: [0, 11],
    peer: then(at(0, 5), key('paragraph')),
    after: ['Hello', ' world'],
  },
  {
    name: 'a peer Backspace that merges the paragraph',
    paragraphs: ['one', 'two'],
    place: [1, 3],
    peer: then(at(1, 0), key('deleteBackward')),
    after: ['onetwo'],
  },
  {
    name: 'a merge into a paragraph that starts like it',
    paragraphs: ['The cat', 'The dog'],
    place: [1, 3],
    peer: then(at(1, 0), key('deleteBackward')),
    after: ['The catThe dog'],
  },
  {
    name: 'a merge of two equal paragraphs',
    paragraphs: ['Item', 'Item'],
    place: [1, 4],
    peer: then(at(1, 0), key('deleteBackward')),
    after: ['ItemItem'],
  },
  {
    name: 'a merge with the same text in another paragraph',
    paragraphs: ['one', 'two', 'two'],
    place: [1, 3],
    peer: then(at(1, 0), key('deleteBackward')),
    after: ['onetwo', 'two'],
  },
  {
    name: 'a peer delete before the caret',
    paragraphs: ['Hello world'],
    place: [0, 11],
    peer: then(range([0, 0], [0, 6]), key('deleteBackward')),
    after: ['world'],
  },
  {
    name: 'a peer typing before the caret',
    paragraphs: ['Hello world'],
    place: [0, 11],
    peer: then(at(0, 0), typed('X')),
    after: ['XHello world'],
  },
  {
    name: 'a peer typing the same word before the caret',
    paragraphs: ['Hello'],
    place: [0, 5],
    peer: then(at(0, 0), typed('Hello ')),
    after: ['Hello Hello'],
  },
  {
    name: 'a peer edit over the place',
    paragraphs: ['Hello world'],
    place: [0, 8],
    peer: then(range([0, 6], [0, 10]), typed('Q')),
    after: ['Hello Qd'],
  },
  {
    name: 'a peer delete of the paragraph',
    paragraphs: ['one', 'two'],
    place: [1, 3],
    peer: then(range([0, 3], [1, 3]), key('deleteBackward')),
    after: ['one'],
  },
  {
    name: 'a peer delete of the paragraph when another paragraph has its start',
    paragraphs: ['Hello world', 'Hello'],
    place: [1, 5],
    peer: then(range([0, 11], [1, 5]), key('deleteBackward')),
    after: ['Hello world'],
  },
  {
    name: 'a peer undo of its typing before the place',
    paragraphs: ['Hello'],
    place: [0, 7],
    first: then(at(0, 0), typed('AB')),
    peer: history('undo'),
    after: ['Hello'],
  },
  {
    name: 'a peer undo of a merge',
    paragraphs: ['one', 'two'],
    place: [0, 6],
    first: then(at(1, 0), key('deleteBackward')),
    peer: history('undo'),
    after: ['one', 'two'],
  },
  {
    name: 'a peer redo of typing before the place',
    paragraphs: ['Hello'],
    place: [0, 5],
    first: then(at(0, 0), typed('AB'), history('undo')),
    peer: history('redo'),
    after: ['ABHello'],
  },
  {
    name: 'a peer split, then typing in the new paragraph',
    paragraphs: ['aa bb cc'],
    place: [0, 8],
    peer: then(at(0, 3), key('paragraph'), at(1, 0), typed('cc ')),
    after: ['aa ', 'cc bb cc'],
  },
];

describe('a peer edit before or over a composition', () => {
  for (const c of refusals) {
    test(`${c.name} ends the composition and refuses its text`, async () => {
      const { a, b, compose, commit, composing, done } = await peers(
        c.paragraphs.map((t) => para(run(t))).join('')
      );
      c.first?.(b);
      at(...c.place)(a);
      compose('日本');
      c.peer(b);
      expect(composing()).toBe(false);
      expect(a.snapshot().lastRejection?.reason).toBe('errors.compositionChanged');
      // The browser's late end adds nothing: the text goes in nowhere.
      commit('日本');
      expect(shapeText(a)).toEqual(c.after);
      done();
    });
  }

  test('a peer change inside a recomposed word refuses the word change', async () => {
    const { a, b, input, start, commit, composing, done } = await peers(para(run('hello wrold')));
    caret(a, nodeId(a, 'T'), 0, 11);
    // A keyboard composes "wrold" again.
    input.setSelectionRange(6, 11);
    start();
    then(range([0, 8], [0, 9]), typed('X'))(b);
    expect(composing()).toBe(false);
    input.value = 'hello world';
    input.setSelectionRange(11, 11);
    commit('world');
    expect(shapeText(a)).toEqual(['hello wrXld']);
    done();
  });

  test('a peer Enter before the place in a cell refuses its text', async () => {
    const { a, b, compose, commit, composing, done } = await peers(para(run('x')), [
      'abc',
      'hello world',
    ]);
    caret(a, tableId(a), 0, 11, { row: 1, column: 0 });
    compose('日本');
    caret(b, tableId(b), 0, 5, { row: 1, column: 0 });
    expect(b.input('paragraph').ok).toBe(true);
    expect(composing()).toBe(false);
    expect(a.snapshot().lastRejection?.reason).toBe('errors.compositionChanged');
    commit('日本');
    expect(paragraphs(a)).toEqual(['x', 'abc', 'hello', ' world']);
    done();
  });

  test('peer edits before the place in a long paragraph refuse its text', async () => {
    const long = 'abcdefghij'.repeat(1000);
    const { a, b, compose, commit, done } = await peers(para(run(long)));
    caret(a, nodeId(a, 'T'), 0, 10_000);
    compose('日本');
    for (let i = 0; i < 20; i++) then(at(0, i * 400), typed('Z'))(b);
    commit('日本');
    expect(shapeText(a)[0]).not.toContain('日本');
    expect(a.snapshot().lastRejection?.reason).toBe('errors.compositionChanged');
    done();
  });
});
