// The framework-neutral pane rules that the React and Vue panes share (`pane-support.ts`).

import { describe, expect, test } from 'bun:test';
import {
  dragOverSlide,
  notesCanFormat,
  notesFormatRange,
  notesFormatShortcut,
  notesHistoryShortcut,
  notesWordAt,
  slideClickMode,
  slideDropMarker,
} from '../pane-support.ts';

const keys = { metaKey: false, ctrlKey: false, shiftKey: false, altKey: false };

describe('pane support', () => {
  test('a click selects by its modifier', () => {
    expect(slideClickMode({ shiftKey: true, metaKey: true, ctrlKey: false })).toBe('range');
    expect(slideClickMode({ shiftKey: false, metaKey: false, ctrlKey: true })).toBe('toggle');
    expect(slideClickMode({ shiftKey: false, metaKey: false, ctrlKey: false })).toBe('replace');
  });

  test('a drag marks the side of the slide under the pointer, never a moving slide', () => {
    const drag = { moving: [0], over: 0, side: 'before' as const };
    const box = { left: 0, top: 100, width: 50, height: 40 };
    const over = dragOverSlide(drag, 2, 'y', box, { clientX: 0, clientY: 130 });
    expect(over).toEqual({ moving: [0], over: 2, side: 'after' });
    expect(dragOverSlide(drag, 2, 'x', box, { clientX: 10, clientY: 130 }).side).toBe('before');
    expect(slideDropMarker(over, 2)).toBe('after');
    expect(slideDropMarker(over, 1)).toBeUndefined();
    expect(slideDropMarker({ ...over, over: 0 }, 0)).toBeUndefined();
  });

  test('a notes format acts on the selection, or on the word around a caret', () => {
    expect(notesWordAt('one two', 5)).toEqual({ start: 4, end: 7 });
    // A caret at the edge of a word is not in it.
    expect(notesWordAt('one two', 4)).toBeNull();
    expect(notesFormatRange('one two', { start: 1, end: 1 })).toEqual({ start: 0, end: 3 });
    expect(notesFormatRange('one two', { start: 0, end: 7 })).toEqual({ start: 0, end: 7 });
    // A caret in no word still takes a format: it arms the next typed text.
    expect(notesCanFormat('one two', { start: 3, end: 3 })).toBe(true);
    expect(notesCanFormat('one two', null)).toBe(false);
  });

  test('the notes shortcuts', () => {
    expect(notesFormatShortcut({ ...keys, key: 'B', ctrlKey: true })).toBe('bold');
    expect(notesFormatShortcut({ ...keys, key: 'b', ctrlKey: true, shiftKey: true })).toBeNull();
    expect(notesHistoryShortcut({ ...keys, key: 'z', metaKey: true })).toBe('undo');
    expect(notesHistoryShortcut({ ...keys, key: 'z', metaKey: true, shiftKey: true })).toBe('redo');
    expect(notesHistoryShortcut({ ...keys, key: 'y', ctrlKey: true })).toBe('redo');
  });
});
