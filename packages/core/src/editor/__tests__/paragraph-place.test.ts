// The offset of a composition's place through a change of its paragraph's text.

import { describe, expect, test } from 'bun:test';
import { diffHunks } from '../../store/package/text-diff.ts';
import { mapOffset } from '../paragraph-place.ts';

const map = (before: string, after: string, offset: number) =>
  mapOffset(diffHunks(before, after), offset);

describe('mapOffset', () => {
  test('text changed before the offset moves it; text after it does not', () => {
    expect(map('Hello world', 'XHello world', 11)).toBe(12);
    expect(map('Hello world', 'world', 11)).toBe(5);
    expect(map('Hello world', 'Hello', 5)).toBe(5);
  });

  test('text inserted at the offset comes after it', () => {
    expect(map('ab', 'aXb', 1)).toBe(1);
    expect(map('Hello world', 'Hello world, bye', 11)).toBe(11);
  });

  test('an offset inside a changed region has no place', () => {
    expect(map('Hello world', 'Hello XY', 9)).toBeNull();
  });
});
