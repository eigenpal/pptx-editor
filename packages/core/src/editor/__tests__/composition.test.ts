// The composition state: its place, each transition, and where its text goes in.

import { describe, expect, test } from 'bun:test';
import type { EditorSelection } from '../../contracts/index.ts';
import {
  applyComposition,
  contextChange,
  continues,
  endedComposition,
  ending,
  finalEdit,
  followEdit,
  IDLE,
  partialEdit,
  placement,
  resumed,
  startComposition,
  type CompositionHost,
} from '../composition.ts';
import type { ReplacementPlace, ReplacementTarget, TypingContext } from '../replacement-text.ts';

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
});
const expectA = (text: string) => ({ paragraph: 'para-a', text });

describe('the composition state', () => {
  test('a start records the context and the input selection the system made', () => {
    const state = startComposition(ctx('hello world'), 'hello world', 6, 11, 80);
    expect(state).toEqual({
      phase: 'composing',
      place: { context: ctx('hello world'), value: 'hello world', from: 6, to: 11 },
      seen: ctx('hello world'),
      width: 80,
      resumes: null,
    });
  });

  test('a new context is the same place, an edit that moves it, or a caret that left', () => {
    const state = startComposition(ctx('abc '), 'abc ', 4, 4, 0);
    expect(contextChange(state, ctx('abc '))).toBe('same');
    // A peer's edit before the caret: a new revision, the same paragraph.
    expect(contextChange(state, ctx('xabc ', 2))).toBe('follow');
    // A click: the same revision, another caret.
    expect(contextChange(state, ctx('a'))).toBe('left');
    // A peer's Enter above: a new index (key), the same paragraph element.
    expect(contextChange(state, ctx('abc ', 2, 'p1'))).toBe('follow');
    // An edit that took the caret to another paragraph: the place is looked up again.
    expect(contextChange(state, ctx('abc ', 2, 'p1', 'para-b'))).toBe('moved');
    // A click into another paragraph, with no edit.
    expect(contextChange(state, ctx('abc ', 1, 'p1', 'para-b'))).toBe('left');
    expect(contextChange(state, null)).toBe('left');
    expect(contextChange(IDLE, null)).toBe('same');
  });

  test('an edit moves the place; an end waits for the final text', () => {
    const state = startComposition(ctx('abc '), 'abc ', 4, 4, 0);
    const moved = followEdit(state, ctx('xabc ', 2));
    expect(moved.phase === 'composing' && moved.seen).toEqual(ctx('xabc ', 2));
    expect(contextChange(moved, ctx('xabc ', 2))).toBe('same');
    expect(ending(state).phase).toBe('ending');
    expect(ending(IDLE)).toBe(IDLE);
  });

  test('the text goes in at the caret while it is the place, else at the place', () => {
    const state = startComposition(ctx('abc '), 'abc ', 4, 4, 0);
    const expect4 = expectA('abc ');
    expect(placement(state, ctx('abc '))).toEqual({ typed: true, expect: expect4 });
    // An edit that sync has not followed yet: the place, checked by its paragraph and text.
    expect(placement(state, ctx('xabc ', 2))).toEqual({ at: caret(4), expect: expect4 });
    expect(placement(state, ctx('a'))).toEqual({ at: caret(4), expect: expect4 });
    expect(placement(state, null)).toEqual({ at: caret(4), expect: expect4 });
    expect(placement(startComposition(null, '', 0, 0, 0), null)).toBeNull();
    expect(placement(IDLE, null)).toBeNull();
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
  type Call = [ReplacementTarget | null, string, ReplacementPlace | undefined];
  const fake = (results: boolean[]) => {
    const calls: Call[] = [];
    const typed: string[] = [];
    const refused: string[] = [];
    const host: CompositionHost = {
      textInput: (_kind, data) => void typed.push(data),
      replaceText: (target, text, where) => (calls.push([target, text, where]), results.shift()!),
      refuse: (reason, detail) => void refused.push(`${reason}:${detail}`),
    };
    return { host, calls, typed, refused };
  };
  const insert = { target: { text: '', after: 0, between: '', paragraph: 'p0' }, text: 'にほ' };

  test('plain typing at the caret is the editor text input', () => {
    const { host, typed, calls } = fake([]);
    expect(applyComposition(host, { typed: true }, insert)).toBe(true);
    expect(typed).toEqual(['にほ']);
    expect(calls).toEqual([]);
  });

  test('at a place the caret left, the text goes to that place', () => {
    const { host, calls } = fake([true]);
    expect(applyComposition(host, { at: caret(4) }, insert)).toBe(true);
    expect(calls).toEqual([[insert.target, 'にほ', { at: caret(4) }]]);
  });

  test('added text whose place is gone goes to the start of its own paragraph', () => {
    const { host, calls } = fake([false, true]);
    const where = { at: caret(9, 2), expect: expectA('abc ') };
    expect(applyComposition(host, where, insert)).toBe(true);
    // The paragraph by its element id: never the paragraph that now has its index.
    const start = { at: caret(0, 2), expect: expectA('') };
    expect(calls[1]).toEqual([{ ...insert.target, paragraph: null }, 'にほ', start]);
  });

  test('a place with no paragraph id gets no fallback', () => {
    const { host, calls } = fake([false]);
    expect(applyComposition(host, { at: caret(9, 2) }, insert)).toBe(false);
    expect(calls).toHaveLength(1);
  });

  test('a changed word whose place is gone is refused, and nothing else changes', () => {
    const { host, calls } = fake([false]);
    const word = {
      target: { text: 'wrold', after: 0, between: '', paragraph: 'p0' },
      text: 'world',
    };
    expect(applyComposition(host, { at: caret(11) }, word)).toBe(false);
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

describe('the rest of an ended composition', () => {
  test('input continues the committed text when it adds, takes, or corrects', () => {
    expect(continues('にほ', 'にほん。')).toBe(true); // adds
    expect(continues('world', 'worldx')).toBe(true);
    expect(continues('world', 'worl')).toBe(true); // a Backspace
    expect(continues('wrold', 'world')).toBe(true); // a swap
    expect(continues('teh', 'the')).toBe(true);
    expect(continues('wor', 'abcd')).toBe(false); // another word
    expect(continues('hello', 'help me')).toBe(false);
    expect(continues('', 'x')).toBe(false);
    expect(continues('x', '')).toBe(false);
  });

  test('an ended composition keeps its caret, its text, and its paragraph', () => {
    const start = startComposition(ctx('hello world'), 'hello world', 6, 11, 0);
    expect(endedComposition(start, 'hello world')).toEqual({
      at: caret(11),
      expect: expectA('hello world'),
      target: { text: 'world', after: 0, between: '', paragraph: 'p0' },
    });
    const typing = startComposition(ctx('abc '), 'abc ', 4, 4, 0);
    const ended = endedComposition(typing, 'abc にほ');
    expect(ended).toEqual({
      at: caret(6),
      expect: expectA('abc にほ'),
      target: { text: 'にほ', after: 0, between: '', paragraph: 'p0' },
    });
    // The rest of it replaces the committed text there.
    expect(resumed(ended!, 'にほん。')).toEqual({
      at: caret(8),
      expect: expectA('abc にほん。'),
      target: { text: 'にほん。', after: 0, between: '', paragraph: 'p0' },
    });
  });

  test('a place an edit moved, or a field changed outside it, keeps no ended composition', () => {
    const typing = startComposition(ctx('abc '), 'abc ', 4, 4, 0);
    expect(endedComposition(followEdit(typing, ctx('xabc ', 2)), 'abc にほ')).toBeNull();
    expect(endedComposition(typing, 'xbc にほ')).toBeNull();
    expect(endedComposition(IDLE, '')).toBeNull();
  });
});
