// The surface's input while an IME composes: it shows the text being composed at the caret
// (the slide shows the text only once the composition ends). DOM only; the state of the
// composition is in `composition.ts`.

/** A strong right-to-left letter: Hebrew, Arabic, Syriac, Thaana, N'Ko, and their forms. */
const RTL = /[֐-ࣿיִ-﷿ﹰ-ﻼ]/u;

/**
 * The width of a text in the input's font, in CSS pixels, or 0 where no canvas measures it
 * (then the composing overlay keeps its plain placement).
 */
function measurer(input: HTMLTextAreaElement): (text: string) => number {
  let context: CanvasRenderingContext2D | null | undefined;
  return (text) => {
    try {
      context ??= input.ownerDocument.createElement('canvas').getContext('2d');
      const view = input.ownerDocument.defaultView;
      if (!context || !view) return 0;
      context.font = view.getComputedStyle(input).font;
      return context.measureText(text).width;
    } catch {
      return 0;
    }
  };
}

/**
 * Fit the input to a composition that started from the value `from` (text width
 * `fromWidth`), or reset it when `from` is null (no composition). The input grows with the
 * composed text, and the context before it scrolls out of view. A keyboard that composes the
 * word before the caret again (`compositionEdit`) moves the input left over that word, so the
 * composed text covers it; in right-to-left text that word is on the right of the caret, so
 * the input stays at the caret and grows right.
 */
export function createCompositionOverlay(input: HTMLTextAreaElement) {
  const width = measurer(input);
  return (from: string | null, fromWidth: number): void => {
    input.style.width = '';
    input.style.marginLeft = '';
    input.style.direction = '';
    const value = input.value;
    if (from === null || value === from) return;
    let shared = 0;
    while (shared < from.length && shared < value.length && from[shared] === value[shared])
      shared += 1;
    const back = shared < from.length ? width(from.slice(shared)) : 0;
    const shown = back > 0 ? width(value.slice(shared)) : 0;
    if (back > 0 && shown > 0) {
      input.style.width = `${Math.ceil(shown) + 1}px`;
      if (RTL.test(from.slice(shared))) {
        input.style.direction = 'rtl';
        input.scrollLeft = -input.scrollWidth; // The end of right-to-left text is on its left.
        return;
      }
      input.style.marginLeft = `${-back}px`;
      input.scrollLeft = input.scrollWidth;
      return;
    }
    input.style.width = `${Math.max(1, input.scrollWidth - fromWidth) + 1}px`;
    input.scrollLeft = input.scrollWidth;
  };
}
