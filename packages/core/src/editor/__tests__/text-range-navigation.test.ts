import { expect, test } from 'bun:test';
import { collapseTextRange } from '../text-input.ts';

for (const backward of [false, true]) {
  test(`horizontal range collapse across paragraphs (${backward ? 'backward' : 'forward'})`, () => {
    const start = { paragraph: 0, offset: 3 };
    const end = { paragraph: 1, offset: 2 };
    const selection = {
      kind: 'text' as const,
      nodeId: 'text-shape',
      anchor: backward ? end : start,
      focus: backward ? start : end,
    };
    expect(collapseTextRange(selection, 'ArrowLeft')).toEqual(start);
    expect(collapseTextRange(selection, 'ArrowRight')).toEqual(end);
    expect(collapseTextRange(selection, 'ArrowUp')).toBeNull();
    expect(collapseTextRange(selection, 'Home')).toBeNull();
    expect(collapseTextRange({ ...selection, anchor: end, focus: end }, 'ArrowRight')).toBeNull();
  });
}
