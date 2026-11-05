// Plain text offsets and paragraph points of one text body.
//
// The automation protocol addresses text by UTF-16 offset over the paragraphs joined with `\r`
// (`plan-text.ts`). The editor addresses it by paragraph and offset (`TextPoint`). A browser host
// maps between the two when it reads or moves the reader's text selection.

import type { TextPoint } from '../contracts/neutral.ts';
import type { Element } from '../store/index.ts';
import { paragraphLength, paragraphsOf } from './commands/text-model.ts';

/** The plain offset of a paragraph point. Points past the end clamp to the end. */
export function textOffsetOf(
  txBody: Element,
  point: TextPoint,
  slideNumber: number | undefined
): number {
  const paragraphs = paragraphsOf(txBody);
  let base = 0;
  for (let i = 0; i < paragraphs.length; i += 1) {
    const len = paragraphLength(paragraphs[i]!, slideNumber);
    if (i === point.paragraph) return base + Math.max(0, Math.min(point.offset, len));
    base += len + 1;
  }
  return Math.max(0, base - 1);
}

/** The paragraph point of a plain offset, or undefined past the end of the text. */
export function textPointOf(
  txBody: Element,
  offset: number,
  slideNumber: number | undefined
): TextPoint | undefined {
  const paragraphs = paragraphsOf(txBody);
  let base = 0;
  for (let i = 0; i < paragraphs.length; i += 1) {
    const len = paragraphLength(paragraphs[i]!, slideNumber);
    if (offset <= base + len) return { paragraph: i, offset: offset - base };
    base += len + 1;
  }
  return undefined;
}
