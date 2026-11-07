// The comment chrome helpers: the author color slot, the avatar text, and the marker styles.

import { describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import type { CommentMarkerStyle, CommentThread } from '../../contracts/index.ts';
import type { SlideLayout } from '../../layout/index.ts';
import { COMMENT_AUTHOR_SLOTS, commentAuthorSlot, commentInitials } from '../comment-authors.ts';
import { CommentMarkers } from '../comment-markers.ts';

describe('commentAuthorSlot', () => {
  test('gives one stable slot per name, in range', () => {
    for (const name of ['Ada Lovelace', 'Grace Hopper', 'Alan Turing', '名前', 'x']) {
      const slot = commentAuthorSlot(name);
      expect(slot).toBe(commentAuthorSlot(name));
      expect(slot).toBeGreaterThanOrEqual(0);
      expect(slot).toBeLessThan(COMMENT_AUTHOR_SLOTS);
    }
    expect(commentAuthorSlot('')).toBe(0);
  });

  test('spreads names over more than one slot', () => {
    const names = ['Ada', 'Grace', 'Alan', 'Edsger', 'Barbara', 'Donald', 'Frances', 'John'];
    expect(new Set(names.map(commentAuthorSlot)).size).toBeGreaterThan(2);
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

  function draw(style: CommentMarkerStyle, threads: readonly CommentThread[], active = 't1') {
    const doc = new Window().document as unknown as Document;
    const markers = new CommentMarkers({
      threads: () => threads,
      active: () => active,
      style: () => style,
      pick: () => {},
    });
    return markers.elements(doc, layout, 1);
  }

  test('icon: a box and a comment glyph, a check when resolved', () => {
    const [open, done] = draw('icon', [thread({}), thread({ id: 't2', resolved: true })]);
    expect(open!.getAttribute('data-marker')).toBe('icon');
    expect(open!.getAttribute('class')).toContain('is-active');
    expect(open!.querySelector('.pptx-surface__comment-marker-box')).not.toBeNull();
    expect(done!.getAttribute('class')).toContain('is-resolved');
    // Markers at one place sit side by side.
    expect(open!.getAttribute('transform')).not.toBe(done!.getAttribute('transform'));
  });

  test('initials: a badge in the author slot, with the reply count', () => {
    const replies = [{ id: 'r1' }, { id: 'r2' }] as unknown as CommentThread['replies'];
    const [g] = draw('initials', [thread({ replies })]);
    expect(g!.getAttribute('data-marker')).toBe('initials');
    expect(g!.getAttribute('style')).toBe(
      `--doc-review-author-current: var(--doc-review-author-${commentAuthorSlot('Ada Lovelace')})`
    );
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
