// The surface's side of the composition preview (`composition-preview.ts`): the composed
// text the input reports, the slide laid out with it, and the underline that marks it, as a
// platform text field marks the text it composes. DOM only.

import type { EditorSelection } from '../contracts/index.ts';
import type { LayoutItem, SlideLayout } from '../layout/index.ts';
import { SVG_NS } from '../output/index.ts';
import type { CompositionPreview, CompositionView } from './composition-preview.ts';
import { textFrameMatrix } from './surface-geometry.ts';
import { bidiSelectionBands } from './visual-caret.ts';

/** What the composition layer asks of the editor. */
export interface CompositionLayerHost {
  readonly layout: () => SlideLayout | undefined;
  /** The slide with a composition's change in it (`compositionView`). */
  compositionView?(preview: CompositionPreview): CompositionView | undefined;
}

/** The composition that the slide shows, if any. */
export class CompositionLayer {
  #preview: CompositionPreview | null = null;
  #memo: { preview: CompositionPreview; layout: SlideLayout; view?: CompositionView } | null = null;

  constructor(private readonly host: CompositionLayerHost) {}

  /** Show `preview` (null: none); whether the slide can show it. */
  set(preview: CompositionPreview | null): boolean {
    this.#preview = preview;
    return !preview || this.view() !== undefined;
  }

  /** The slide with the composed text, for the current layout; undefined with none. */
  view(): CompositionView | undefined {
    const preview = this.#preview;
    const layout = this.host.layout();
    if (!preview || !layout || !this.host.compositionView) return undefined;
    const memo = this.#memo;
    if (memo?.preview === preview && memo.layout === layout) return memo.view;
    const view = this.host.compositionView(preview);
    this.#memo = { preview, layout, ...(view ? { view } : {}) };
    return view;
  }

  /** The selection the overlay draws: the caret in the composed text while there is one. */
  selection(selection: EditorSelection): EditorSelection {
    return this.view()?.selection ?? selection;
  }

  /**
   * The underline of the composed text in `item`, in slide units (`px`: one pixel), in the
   * text's `color` when it has one, as the caret is.
   */
  marks(doc: Document, item: LayoutItem, px: number, color?: string): Element[] {
    const view = this.view();
    if (!view || item.kind !== 'shape' || !item.text) return [];
    const m = textFrameMatrix(item);
    const { from, to } = view;
    if (!m || to.offset <= from.offset) return [];
    const g = doc.createElementNS(SVG_NS, 'g');
    g.setAttribute('transform', `matrix(${m.join(' ')})`);
    g.setAttribute('class', 'pptx-surface__composition');
    if (color) g.style.stroke = color;
    for (const line of item.text.lines) {
      if (line.paragraph !== from.paragraph) continue;
      const s = Math.max(from.offset, line.start);
      const e = Math.min(to.offset, line.end);
      if (e <= s) continue;
      const bands = line.bidi ? bidiSelectionBands(line, s, e) : [];
      if (bands.length === 0) {
        const x1 = line.caretX[s - line.start] ?? line.caretX[0] ?? 0;
        const x2 = line.caretX[Math.min(e - line.start, line.caretX.length - 1)] ?? x1;
        bands.push([Math.min(x1, x2), Math.max(x1, x2)]);
      }
      // Under the letters, clear of the baseline: a descender shows over it.
      const y = line.baseline + Math.max(line.fontHeight * 0.12, px * 2);
      for (const [x1, x2] of bands) {
        const l = doc.createElementNS(SVG_NS, 'line');
        l.setAttribute('x1', String(x1));
        l.setAttribute('x2', String(x2));
        l.setAttribute('y1', String(y));
        l.setAttribute('y2', String(y));
        l.setAttribute('stroke-width', String(px * 1.5));
        g.append(l);
      }
    }
    return [g];
  }
}

/**
 * Keep the hidden input on the caret, as tall as the caret line: the stage holds it, so this
 * follows zoom and scroll, and an IME candidate window opens under the caret.
 */
export function placeInput(
  input: HTMLTextAreaElement,
  caret: { x1: number; y1: number; x2: number; y2: number },
  scale: number
): void {
  const top = Math.min(caret.y1, caret.y2) * scale;
  const height = Math.max(1, Math.abs(caret.y2 - caret.y1) * scale);
  const style = input.style;
  style.left = `${Math.min(caret.x1, caret.x2) * scale}px`;
  style.top = `${top}px`;
  style.height = `${height}px`;
  style.fontSize = style.lineHeight = `${height}px`;
}
