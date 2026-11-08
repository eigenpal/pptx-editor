// The composition state: its place, each transition, and where its text goes in.

import { describe, expect, test } from 'bun:test';
import type { EditorSelection } from '../../contracts/index.ts';
import {
  applyComposition,
  contextChange,
  ending,
  finalEdit,
  placeKept,
  placeLost,
  IDLE,
  partialEdit,
  placement,
  startComposition,
  type CompositionHost,
} from '../composition.ts';
import type { ReplacementTarget, TextPlace, TypingContext } from '../replacement-text.ts';

const caret = (offset: number, paragraph = 0): Extract<EditorSelection, { kind: 'text' }> => ({
  kind: 'text',
  nodeId: 'n1',
  anchor: { paragraph, offset },
  focus: { paragraph, offset },
});
const ctx = (text: string, revision = 1, key = 'p0', paragraphId = 'para-a'): TypingContext => ({
  key,
  text,
  selection: caret(text.length),
  revision,
  paragraphId,
  before: text,
});

describe('the composition state', () => {
  test('a start records the context and the input selection the system made', () => {
    const state = startComposition(ctx('hello world'), 'hello world', 6, 11, 80);
    expect(state).toEqual({
      phase: 'composing',
      place: { context: ctx('hello world'), value: 'hello world', from: 6, to: 11 },
      seen: ctx('hello world'),
      width: 80,
    });
  });

  test('a peer edit is checked again; a local move or edit ends the composition', () => {
    const state = startComposition(ctx('abc '), 'abc ', 4, 4, 0);
    expect(contextChange(state, ctx('abc '), false)).toBe('same');
    // A peer's edit: a new revision, so the place is checked again, wherever the caret went.
    expect(contextChange(state, ctx('xabc ', 2), true)).toBe('remote');
    expect(contextChange(state, ctx('abc ', 2, 'p1', 'para-b'), true)).toBe('remote');
    // A local or API edit: the composition ends, and the caret stays where the edit left it.
    expect(contextChange(state, ctx('abc !', 2), false)).toBe('left');
    // A click: the same revision, another caret, also in another paragraph.
    expect(contextChange(state, ctx('a'), false)).toBe('left');
    expect(contextChange(state, ctx('abc ', 1, 'p1', 'para-b'), false)).toBe('left');
    // No context, before or now.
    expect(contextChange(state, null, true)).toBe('left');
    expect(contextChange(startComposition(null, '', 0, 0, 0), null, false)).toBe('left');
    expect(contextChange(IDLE, null, false)).toBe('same');
  });

  test('a peer edit that keeps the place moves its context; an end waits for the text', () => {
    const state = startComposition(ctx('abc '), 'abc ', 4, 4, 0);
    const moved = placeKept(state, ctx('xabc ', 2));
    expect(moved.phase === 'composing' && moved.seen).toEqual(ctx('xabc ', 2));
    expect(contextChange(moved, ctx('xabc ', 2), false)).toBe('same');
    const lost = placeLost(state);
    expect(lost.phase !== 'idle' && lost.lost).toBe(true);
    expect(ending(state).phase).toBe('ending');
    expect(ending(IDLE)).toBe(IDLE);
  });

  test('the place is the caret, its paragraph by element id, and the exact text before it', () => {
    const state = startComposition(ctx('abc '), 'abc ', 4, 4, 0);
    expect(placement(state)).toEqual({ at: caret(4), paragraph: 'para-a', before: 'abc ' });
    // An edit that kept the place: its context there.
    expect(placement(placeKept(state, ctx('abc ', 2, 'p1')))).toEqual({
      at: caret(4),
      paragraph: 'para-a',
      before: 'abc ',
    });
    // No paragraph id or exact text: no place, so nothing is written.
    const loose = { key: 'p0', text: 'abc ', selection: caret(4), revision: 1 };
    expect(placement(startComposition(loose, 'abc ', 4, 4, 0))).toBeNull();
    expect(placement(startComposition(null, '', 0, 0, 0))).toBeNull();
    expect(placement(IDLE)).toBeNull();
  });
});

describe('the change a composition makes', () => {
  const place = (value: string, from: number, to: number) => ({
    context: ctx(value),
    value,
    from,
    to,
  });

  test('a recomposed word changes only its changed characters', () => {
    expect(finalEdit(place('a helo', 6, 6), 'a hello', 7, 'hello')).toEqual({
      target: { text: '', after: 1, between: 'o', paragraph: 'p0' },
      text: 'l',
    });
  });

  test('a commit the field does not hold replaces what the system selected', () => {
    // A reconversion of "world" whose field the browser left unchanged.
    expect(finalEdit(place('hello world', 6, 11), 'hello world', 11, 'WORLD')).toEqual({
      target: { text: 'world', after: 0, between: '', paragraph: 'p0' },
      text: 'WORLD',
    });
    expect(finalEdit({ context: null, value: '', from: 0, to: 0 }, 'x', 1, 'x')).toBeNull();
  });

  test('a composition that the browser ends with no text keeps what the field shows', () => {
    expect(partialEdit(place('abc ', 4, 4), 'abc にほ')).toEqual({
      target: { text: '', after: 0, between: '', paragraph: 'p0' },
      text: 'にほ',
    });
    expect(partialEdit(place('abc ', 4, 4), 'abc ')).toBeNull();
  });
});

describe('applyComposition', () => {
  type Call = [TextPlace, ReplacementTarget, string];
  const fake = (results: boolean[]) => {
    const calls: Call[] = [];
    const refused: string[] = [];
    const host: CompositionHost = {
      compose: (place, target, text) => (calls.push([place, target, text]), results.shift()!),
      refuse: (reason, detail) => void refused.push(`${reason}:${detail}`),
    };
    return { host, calls, refused };
  };
  const insert = { target: { text: '', after: 0, between: '', paragraph: 'p0' }, text: 'にほ' };
  const here: TextPlace = { at: caret(4), paragraph: 'para-a', before: 'abc ' };

  test('every change goes to the place, through the one checked write', () => {
    const { host, calls } = fake([true]);
    expect(applyComposition(host, here, insert)).toBe(true);
    expect(calls).toEqual([[here, insert.target, 'にほ']]);
  });

  test('a place that no longer holds writes nothing else', () => {
    const { host, calls } = fake([false]);
    expect(applyComposition(host, here, insert)).toBe(false);
    // No second write: never the paragraph start, never another paragraph.
    expect(calls).toHaveLength(1);
  });

  test('no place refuses with a reason; no change applies nothing', () => {
    const { host, refused, calls } = fake([]);
    expect(applyComposition(host, null, insert)).toBe(false);
    expect(refused).toEqual(['errors.inputUnsupported:insertCompositionText']);
    const none = { target: { ...insert.target }, text: '' };
    expect(applyComposition(host, null, none)).toBe(true);
    expect(calls).toEqual([]);
  });
});
