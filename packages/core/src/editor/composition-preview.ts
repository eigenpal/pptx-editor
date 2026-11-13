// The text an IME composes, shown in the slide while the user composes it. The composition
// writes to the model only when it ends (`composition.ts`). Until then the slide lays out the
// change as typing would write it, over a copy of the package: the run's own font, size, and
// color, the line breaks it makes, and the autofit it needs. Nothing reaches the store, the
// undo history, or a peer. DOCX Editor `surface-composition-preview.ts` at 4644e6c5d keeps
// the composed text visible in its page for the same reason.

import type { EditorSelection, TextPoint } from '../contracts/index.ts';
import type { SlideLayout } from '../layout/index.ts';
import type { PptxPackage } from '../store/index.ts';
import { refit } from '../automation/commands/autofit.ts';
import { CommandRefusal } from '../automation/commands/op-batch.ts';
import { comparePoints } from '../automation/commands/text-model.ts';
import type { DocumentSession } from './document-session.ts';
import { caretParagraph, replacementRange, type ContextEdit } from './replacement-text.ts';
import { selectionBody, textItemOf } from './table-cells.ts';
import { editText } from './text-input.ts';

type TextSelection = Extract<EditorSelection, { kind: 'text' }>;

/** What a composition changed so far, and where the IME caret is in its text. */
export type CompositionPreview =
  /** At a caret: the change to the caret's context (`partialEdit`). */
  | { readonly kind: 'caret'; readonly edit: ContextEdit; readonly caret: number }
  /** Over a selected range: the text that replaces the range. */
  | { readonly kind: 'range'; readonly text: string; readonly caret: number };

/** The slide with the composed text in it. */
export interface CompositionView {
  readonly layout: SlideLayout;
  /** The composed text: the range the slide underlines. */
  readonly from: TextPoint;
  readonly to: TextPoint;
  /** The caret in the composed text, where the candidate window opens. */
  readonly selection: TextSelection;
}

/** What `compositionView` reads from the editor. */
export interface CompositionViewHost {
  session(): DocumentSession | null;
  part(): string | undefined;
  selection(): EditorSelection;
  layout(): SlideLayout | undefined;
  /** The package with the format armed at the caret, which typed text takes (`CaretArm`). */
  armed(): PptxPackage | undefined;
}

/**
 * The slide with a composition's change laid out in it, or undefined when the slide cannot
 * show it (a diagram node, a cell range, a place that no longer holds): the surface then shows
 * the composed text in its input.
 */
export function compositionView(
  host: CompositionViewHost,
  preview: CompositionPreview
): CompositionView | undefined {
  const session = host.session();
  const part = host.part();
  const layout = host.layout();
  const sel = host.selection();
  if (!session || !part || !layout || sel.kind !== 'text' || sel.diagramNode !== undefined)
    return undefined;
  const item = textItemOf(layout, sel);
  if (!item) return undefined;
  let from: TextPoint;
  let to: TextPoint;
  let text: string;
  if (preview.kind === 'caret') {
    const at = caretParagraph(session, part, sel);
    const target = { ...preview.edit.target, paragraph: null };
    const range = at && replacementRange(at.text, at.offset, at.key, target);
    if (!range) return undefined;
    from = { paragraph: sel.focus.paragraph, offset: range.from };
    to = { paragraph: sel.focus.paragraph, offset: range.to };
    text = preview.edit.text;
  } else {
    [from, to] =
      comparePoints(sel.anchor, sel.focus) <= 0 ? [sel.anchor, sel.focus] : [sel.focus, sel.anchor];
    text = preview.text;
  }
  const caret = (offset: number): TextSelection => {
    const point = { paragraph: from.paragraph, offset };
    return { ...sel, anchor: point, focus: point };
  };
  const end = from.offset + text.length;
  const shown = { from, to: { paragraph: from.paragraph, offset: end } };
  const at = caret(from.offset + Math.max(0, Math.min(preview.caret, text.length)));
  // Nothing composed yet: the slide as it is.
  if (!text && comparePoints(from, to) === 0) return { layout, ...shown, selection: at };
  try {
    const batch = session.batch(text ? host.armed() : undefined);
    const over: TextSelection = { ...sel, anchor: from, focus: to };
    const txBody = selectionBody(batch, part, over);
    if (!txBody) return undefined;
    editText(batch, part, txBody.id, item, over, text ? 'text' : 'deleteBackward', text);
    refit(batch, layout.slide.index, part, sel.nodeId, session.measurer);
    const laid = session.layoutOf(batch.pkg, layout.slide.index);
    return laid && { layout: laid, ...shown, selection: at };
  } catch (e) {
    if (e instanceof CommandRefusal) return undefined;
    throw e;
  }
}
