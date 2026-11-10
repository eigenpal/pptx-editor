import { expect, test } from 'bun:test';
import { shownPieces } from '../collaboration/yjs/shown-pieces.ts';

test('piece lookup keeps cut and document-order ties within the winning group', () => {
  const pieces = new Map([
    ['a.early', { group: 'first', at: 2 }],
    ['a.tie', { group: 'first', at: 2 }],
    ['a.equal', { group: 'first', at: 2 }],
    ['a.last', { group: 'first', at: 8 }],
    ['a.hidden', { group: 'first', at: 4 }],
    ['b.loser', { group: 'first', at: 3 }],
    ['b.other', { group: 'second', at: 5 }],
  ]);
  // Document order and cut order differ. A hidden piece has no shown position.
  const where = new Map([
    ['a.early', { order: 1 }],
    ['a.tie', { order: 4 }],
    ['a.equal', { order: 4 }],
    ['a.last', { order: 2 }],
    ['b.loser', { order: 5 }],
    ['b.other', { order: 6 }],
  ]);
  const index = shownPieces(pieces, where, (group) => (group === 'first' ? 'a' : 'b'));
  expect(index.preceding('first', 1)).toBeUndefined();
  for (const cut of [2, 3, 4, 7]) expect(index.preceding('first', cut)).toBe('a.tie');
  for (const cut of [8, 9]) expect(index.preceding('first', cut)).toBe('a.last');
  expect(index.last.get('first')).toBe('a.tie');
  expect(index.preceding('second', 4)).toBeUndefined();
  expect(index.preceding('second', 5)).toBe('b.other');
  expect(index.preceding('missing', 9)).toBeUndefined();
});
