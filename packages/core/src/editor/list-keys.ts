// Enter, Tab, and Backspace while editing text, with lists.
//
// The hidden input is a <textarea>, and a textarea reports Enter as `insertLineBreak`: it has
// no paragraphs. So Enter is handled on keydown, before the browser makes it a line break:
// Enter starts a paragraph, Shift+Enter inserts a line break (`a:br`).
//
// Lists follow Google Slides' editing feel:
// - Tab and Shift+Tab change the list level of the selected paragraphs, of a list item, or of
//   a paragraph whose start holds the caret. Elsewhere, Tab inserts a tab.
// - Backspace at the start of a list item leaves the list: one level up, or off at level 0.
// - Enter on an empty list item does the same, so a list ends the way it does in an editor.

import type { EditorCommand, EditorSelection } from '../contracts/index.ts';
import type { LayoutItem } from '../layout/index.ts';
import type { TextInputKind } from './text-input.ts';

type ShapeItem = Extract<LayoutItem, { kind: 'shape' }>;
type TextSelection = Extract<EditorSelection, { kind: 'text' }>;

/** What a key does instead of its browser default, or null to leave the key alone. */
export type ListKeyAction =
  | { readonly kind: 'input'; readonly input: TextInputKind; readonly data?: string }
  | { readonly kind: 'command'; readonly command: EditorCommand };

/** Leave the list one step: up one level, or off at level 0. */
function leaveList(bullet: 'char' | 'autonum', level: number): ListKeyAction {
  if (level > 0) return { kind: 'command', command: { type: 'indent', direction: -1 } };
  return {
    kind: 'command',
    command: { type: bullet === 'char' ? 'toggleBullets' : 'toggleNumbering' },
  };
}

export function listKeyAction(
  item: LayoutItem | undefined,
  sel: TextSelection,
  key: string,
  modifiers: { readonly shift: boolean; readonly other: boolean }
): ListKeyAction | null {
  if (modifiers.other) return null;
  const paragraphs = item?.kind === 'shape' ? (item as ShapeItem).element.text?.paragraphs : null;
  if (!paragraphs) return null;
  const p = paragraphs[sel.focus.paragraph];
  const collapsed =
    sel.anchor.paragraph === sel.focus.paragraph && sel.anchor.offset === sel.focus.offset;
  const bullet = p?.bullet?.type;
  const level = p?.level ?? 0;
  const empty = !p || p.runs.every((r) => r.kind !== 'break' && r.text.length === 0);
  if (key === 'Enter') {
    if (!modifiers.shift && collapsed && bullet && empty) return leaveList(bullet, level);
    return { kind: 'input', input: modifiers.shift ? 'lineBreak' : 'paragraph' };
  }
  if (key === 'Tab') {
    const several = sel.anchor.paragraph !== sel.focus.paragraph;
    const atStart = collapsed && sel.focus.offset === 0;
    if (several || bullet || atStart || modifiers.shift)
      return { kind: 'command', command: { type: 'indent', direction: modifiers.shift ? -1 : 1 } };
    return { kind: 'input', input: 'text', data: '\t' };
  }
  if (key === 'Backspace' && !modifiers.shift && collapsed && bullet && sel.focus.offset === 0)
    return leaveList(bullet, level);
  return null;
}
