// The comment chrome helpers: the review author roster, the avatar text, and the marker styles.

import { describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import type { CommentMarkerStyle, CommentThread } from '../../contracts/index.ts';
import type { SlideLayout } from '../../layout/index.ts';
import { commentInitials } from '../comment-authors.ts';
import {
  createStableReviewAuthorSlots,
  REVIEW_AUTHOR_SLOTS,
  reviewAuthorHooks,
  reviewAuthorSlotColor,
  reviewAuthorsOf,
} from '../review-authors.ts';
import { CommentMarkers } from '../comment-markers.ts';

describe('review author slots', () => {
  test('give each author the next slot by first appearance, and keep it', () => {
    const slots = createStableReviewAuthorSlots();
    const first = slots.resolve(new Map(), ['Grace', 'Ada', 'Grace', 'Alan']);
    expect([...first]).toEqual([
      ['Grace', 0],
      ['Ada', 1],
      ['Alan', 2],
    ]);
    // Ada's comments go, and a new author comes: nobody else changes color.
    const next = slots.resolve(new Map(), ['Grace', 'Alan', 'Edsger']);
    expect([...next]).toEqual([
      ['Grace', 0],
      ['Alan', 2],
      ['Edsger', 3],
    ]);
    // Ada comes back to her own slot.
    expect(slots.resolve(new Map(), ['Ada']).get('Ada')).toBe(1);
  });

  test('draw eight authors in eight colors, then wrap', () => {
    const names = ['Ada', 'Grace', 'Alan', 'Edsger', 'Barbara', 'Donald', 'Frances', 'John', 'Kay'];
    const roster = reviewAuthorsOf(createStableReviewAuthorSlots().resolve(new Map(), names));
    expect(new Set(roster.slice(0, REVIEW_AUTHOR_SLOTS).map((a) => a.color)).size).toBe(8);
    expect(roster[8]).toEqual({ author: 'Kay', slot: 8, color: 'var(--doc-review-author-0)' });
    expect(reviewAuthorSlotColor(3)).toBe('var(--doc-review-author-3)');
  });

  test('hooks name the author and the wrapped slot, and nothing for no author', () => {
    const info = { author: 'Kay', slot: 9, color: 'var(--doc-review-author-1)' };
    expect(reviewAuthorHooks('Kay', info)).toEqual({ author: 'Kay', slot: 1, color: info.color });
    expect(reviewAuthorHooks('', info)).toBeNull();
  });
});

describe('commentInitials', () => {
  test('uses own initials, then the first two words, then "?"', () => {
    expect(commentInitials({ initials: ' abcd ', author: 'Ada Lovelace' })).toBe('abc');
    expect(commentInitials({ initials: '', author: 'ada byron lovelace' })).toBe('AB');
    expect(commentInitials({ initials: '', author: '𝒜da' })).toBe('𝒜');
    expect(commentInitials({ initials: '  ', author: '  ' })).toBe('?');
  });
});

describe('CommentMarkers', () => {
  const thread = (over: Partial<CommentThread>): CommentThread =>
    ({
      id: 't1',
      author: 'Ada Lovelace',
      initials: '',
      text: '<b>file text</b>',
      resolved: false,
      replies: [],
      anchor: { kind: 'point', x: 100, y: 100 },
      ...over,
    }) as unknown as CommentThread;
  const layout = { items: [], slide: { width: 10_000, height: 10_000 } } as unknown as SlideLayout;

  function draw(
    style: CommentMarkerStyle,
    threads: readonly CommentThread[],
    active = 't1',
    px = 1
  ) {
    const doc = new Window().document as unknown as Document;
    const markers = new CommentMarkers({
      threads: () => threads,
      active: () => active,
      style: () => style,
      pick: () => {},
      selection: () => ({ kind: 'none' }),
      author: (name) =>
        name === 'Ada Lovelace'
          ? { author: name, slot: 5, color: 'var(--doc-review-author-5)' }
          : undefined,
    });
    return markers.elements(doc, layout, px);
  }

  test('initials and the reply count are sized in screen pixels at any zoom', () => {
    // One screen pixel at 100% zoom is 9,525 EMU: a font size in EMU passes the browser cap.
    const replies = [{ id: 'r1' }] as unknown as CommentThread['replies'];
    const [g] = draw('initials', [thread({ replies })], 't1', 9_525);
    for (const cls of ['comment-marker-text', 'comment-marker-count-text']) {
      const text = g!.querySelector(`.pptx-surface__${cls}`)!;
      expect(Number(text.getAttribute('font-size'))).toBeLessThan(12);
      expect(text.getAttribute('transform')).toContain('scale(9525)');
    }
  });

  test('icon: a box and a comment glyph, a check when resolved', () => {
    const [open, done] = draw('icon', [thread({}), thread({ id: 't2', resolved: true })]);
    expect(open!.getAttribute('data-marker')).toBe('icon');
    expect(open!.getAttribute('class')).toContain('is-active');
    expect(open!.querySelector('.pptx-surface__comment-marker-box')).not.toBeNull();
    expect(done!.getAttribute('class')).toContain('is-resolved');
    // Markers at one place sit side by side.
    expect(open!.getAttribute('transform')).not.toBe(done!.getAttribute('transform'));
  });

  test("initials: a badge in the roster's author color, with the reply count", () => {
    const replies = [{ id: 'r1' }, { id: 'r2' }] as unknown as CommentThread['replies'];
    const [g] = draw('initials', [thread({ replies })]);
    expect(g!.getAttribute('data-marker')).toBe('initials');
    expect(g!.getAttribute('style')).toBe(
      '--doc-review-author-current: var(--doc-review-author-5)'
    );
    expect(g!.getAttribute('data-review-author')).toBe('Ada Lovelace');
    expect(g!.getAttribute('data-review-author-slot')).toBe('5');
    expect(g!.querySelector('.pptx-surface__comment-marker-text')!.textContent).toBe('AL');
    expect(g!.querySelector('.pptx-surface__comment-marker-count-text')!.textContent).toBe('2');
  });

  test('file text goes in the title as text only', () => {
    const [g] = draw('icon', [thread({})]);
    const title = g!.querySelector('title')!;
    expect(title.textContent).toBe('Ada Lovelace: <b>file text</b>');
    expect(title.children).toHaveLength(0);
  });
});
