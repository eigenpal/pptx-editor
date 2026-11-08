// A caret's paragraph by its element id, which an edit never changes: a peer's Enter above
// the paragraph or a table row inserted above its cell moves its index, not its id. An IME
// composition keeps the id of its paragraph, so its text goes in only there (`composition.ts`).
// DOM-free.

import type { EditorSelection, TableCellAddress } from '../contracts/index.ts';
import type { OpBatch } from '../automation/commands/op-batch.ts';
import { elements, type Element } from '../store/index.ts';

type TextCaret = Extract<EditorSelection, { kind: 'text' }>;

/** Where a paragraph is in its shape or table: its cell, and its index in the text body. */
interface ParagraphAt {
  readonly cell?: TableCellAddress;
  readonly paragraph: number;
}

/** The paragraph with element id `id` under `el`, depth first; tables give its cell. */
function find(el: Element, id: string, cell?: TableCellAddress): ParagraphAt | null {
  const kids = elements(el);
  const index = kids.filter((k) => k.kind === 'paragraph').findIndex((p) => p.id === id);
  if (index >= 0) return { ...(cell ? { cell } : {}), paragraph: index };
  let row = 0;
  for (const kid of kids) {
    if (kid.kind === 'paragraph') continue;
    if (kid.local !== 'tr') {
      const found = find(kid, id, cell);
      if (found) return found;
      continue;
    }
    let column = 0;
    for (const tc of elements(kid)) {
      if (tc.local !== 'tc') continue;
      const found = find(tc, id, { row, column });
      if (found) return found;
      column += 1;
    }
    row += 1;
  }
  return null;
}

/**
 * The caret `sel` moved to the paragraph with element id `id` in the same shape or table, at
 * the same offset; null when that paragraph is gone. Reads only.
 */
export function caretInParagraph(
  batch: OpBatch,
  part: string,
  sel: TextCaret,
  id: string
): TextCaret | null {
  const root = batch.element(part, sel.nodeId);
  const at = root ? find(root, id) : null;
  if (!at) return null;
  const point = { paragraph: at.paragraph, offset: sel.focus.offset };
  const { cell: _cell, ...rest } = sel;
  return { ...rest, ...(at.cell ? { cell: at.cell } : {}), anchor: point, focus: point };
}
