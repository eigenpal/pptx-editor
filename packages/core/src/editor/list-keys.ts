// Enter, Tab, and Backspace while editing text, with lists.
//
// The hidden input is a <textarea>, and a textarea reports Enter as `insertLineBreak`: it has
// no paragraphs. So Enter is handled on keydown, before the browser makes it a line break:
// Enter starts a paragraph, Shift+Enter inserts a line break (`a:br`).
//
// The list keys:
// - Enter starts a paragraph with the same level and bullet, also on an empty list item. A list
//   does not end on Enter. Backspace or the Bullets button ends it.
// - Tab and Shift+Tab change the level of every paragraph a selection across paragraphs
//   touches, and of a list item whose start holds the caret or the selection start. Anywhere
//   else, both keys type a tab character, and a selection is replaced.
// - Backspace at the start of a list item removes its bullet or number and keeps its level.
//   The text moves to where the bullet was. The next Backspace joins the previous paragraph.

import type { EditorCommand, EditorSelection } from '../contracts/index.ts';
import type { LayoutItem } from '../layout/index.ts';
import type { TextInputKind } from './text-input.ts';

type ShapeItem = Extract<LayoutItem, { kind: 'shape' }>;
type TextSelection = Extract<EditorSelection, { kind: 'text' }>;

/** What a key does instead of its browser default, or null to leave the key alone. */
export type ListKeyAction =
  | { readonly kind: 'input'; readonly input: TextInputKind; readonly data?: string }
  | { readonly kind: 'command'; readonly command: EditorCommand };

export function listKeyAction(
  item: LayoutItem | undefined,
  sel: TextSelection,
  key: string,
  modifiers: { readonly shift: boolean; readonly other: boolean }
): ListKeyAction | null {
  if (modifiers.other) return null;
  const paragraphs = item?.kind === 'shape' ? (item as ShapeItem).element.text?.paragraphs : null;
  if (!paragraphs) return null;
  const collapsed =
    sel.anchor.paragraph === sel.focus.paragraph && sel.anchor.offset === sel.focus.offset;
  if (key === 'Enter') return { kind: 'input', input: modifiers.shift ? 'lineBreak' : 'paragraph' };
  if (key === 'Tab') {
    const several = sel.anchor.paragraph !== sel.focus.paragraph;
    const start =
      sel.anchor.paragraph < sel.focus.paragraph ||
      (sel.anchor.paragraph === sel.focus.paragraph && sel.anchor.offset <= sel.focus.offset)
        ? sel.anchor
        : sel.focus;
    const listStart = start.offset === 0 && !!paragraphs[start.paragraph]?.bullet;
    if (several || listStart)
      return { kind: 'command', command: { type: 'indent', direction: modifiers.shift ? -1 : 1 } };
    return { kind: 'input', input: 'text', data: '\t' };
  }
  const p = paragraphs[sel.focus.paragraph];
  if (key === 'Backspace' && !modifiers.shift && collapsed && p?.bullet && sel.focus.offset === 0)
    return { kind: 'command', command: { type: 'removeBullet' } };
  return null;
}
