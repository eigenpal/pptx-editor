// The keys on selected shapes: Tab through the shapes, nudge, delete, duplicate, Escape out of
// a group, Enter into a group or into text, and a typed character that starts text. Also what
// Select All selects. Split from `pptx-editor.ts`, which applies the action, so the facade
// stays under the line cap.
//
// Keyboard focus on the slide: the slide is ONE Tab stop (`.pptx-surface`, role `group`), as
// the ARIA practice for a composite widget asks. Inside it, Tab and Shift+Tab move the
// selection through the shapes in z-order, as slide editors do: with nothing selected, Tab
// selects the backmost shape and Shift+Tab the frontmost. Past the last shape (or before the
// first), the selection clears and Tab moves focus on out of the slide, so Tab never traps
// focus. Inside an entered group, Tab cycles through the group's children; Escape leaves the
// group. Escape clears a selection, and Enter on a selected text shape starts text editing.

import type { EditorCommand, EditorMode, EditorSelection } from '../contracts/index.ts';
import type { SlideLayout } from '../layout/index.ts';
import { nudgeCommand } from './format-keys.ts';
import { parentGroupOf, scopeItems } from '../automation/commands/placement.ts';
import { textItemOf } from './table-cells.ts';
import { itemById } from './surface.ts';

/** What a key on selected shapes does, or null when it does nothing. @internal */
export type ShapeKeyAction =
  | { readonly kind: 'exec'; readonly command: EditorCommand }
  | { readonly kind: 'select'; readonly selection: EditorSelection }
  /** Enter the shape's text; `type` is a character to type at its end, else all text is selected. */
  | { readonly kind: 'enterText'; readonly nodeId: string; readonly type?: string }
  /** The key is handled and does nothing more (Enter on a group with nothing inside). */
  | { readonly kind: 'handled' }
  /** Tab past the last shape: clear the selection, and let the key move focus off the slide. */
  | { readonly kind: 'leave' };

/**
 * The selection Tab (or Shift+Tab, `back`) moves to: the next slide shape of the current scope
 * in z-order, from the last selected one. Null when Tab leaves the slide: past the end of the
 * slide's shapes, or a slide with none. Inside an entered group it wraps. @internal
 */
export function tabSelection(
  layout: SlideLayout | undefined,
  selection: Extract<EditorSelection, { kind: 'shapes' | 'none' }>,
  back: boolean
): EditorSelection | null {
  const group = selection.kind === 'shapes' ? selection.group : undefined;
  const ids = scopeItems(layout?.items ?? [], group)
    .filter((i) => i.element.layer === 'slide')
    .map((i) => i.element.nodeId);
  if (ids.length === 0) return null;
  const from = selection.kind === 'shapes' ? ids.indexOf(selection.nodeIds.at(-1) ?? '') : -1;
  const step = back ? -1 : 1;
  let at = from < 0 ? (back ? ids.length - 1 : 0) : from + step;
  if (at < 0 || at >= ids.length) {
    if (!group) return null;
    at = (at + ids.length) % ids.length;
  }
  return { kind: 'shapes', nodeIds: [ids[at]], ...(group ? { group } : {}) };
}

/** The action of `key` on a shapes selection, or on no selection (Tab only). @internal */
export function shapeKeyAction(
  key: string,
  keys: { readonly shift: boolean; readonly mod: boolean; readonly alt?: boolean },
  selection: Extract<EditorSelection, { kind: 'shapes' | 'none' }>,
  layout: SlideLayout | undefined,
  mode: EditorMode
): ShapeKeyAction | null {
  const { shift, mod } = keys;
  if (key === 'Tab' && !mod && !keys.alt) {
    const next = tabSelection(layout, selection, shift);
    return next ? { kind: 'select', selection: next } : { kind: 'leave' };
  }
  if (selection.kind === 'none') return null;
  const nudge = nudgeCommand(key, shift, mod);
  if (nudge) return { kind: 'exec', command: nudge };
  if (key === 'Delete' || key === 'Backspace') return { kind: 'exec', command: { type: 'delete' } };
  if (mod && key.toLowerCase() === 'd') return { kind: 'exec', command: { type: 'duplicate' } };
  if (key === 'Escape') {
    // Leave an entered group, selecting it; otherwise clear the selection.
    const group = selection.group;
    const outer = group ? parentGroupOf(layout?.items ?? [], group) : undefined;
    return {
      kind: 'select',
      selection: group
        ? { kind: 'shapes', nodeIds: [group], ...(outer ? { group: outer } : {}) }
        : { kind: 'none' },
    };
  }
  const only = selection.nodeIds.length === 1 ? itemById(layout, selection.nodeIds[0]) : undefined;
  if (key === 'Enter' && only?.kind === 'group') {
    const kids = only.children.filter((c) => c.element.layer === 'slide');
    return kids.length
      ? {
          kind: 'select',
          selection: {
            kind: 'shapes',
            nodeIds: kids.map((c) => c.element.nodeId),
            group: only.element.nodeId,
          },
        }
      : { kind: 'handled' };
  }
  if ((key === 'Enter' || key === 'F2') && selection.nodeIds.length === 1)
    return { kind: 'enterText', nodeId: selection.nodeIds[0] };
  if (key.length === 1 && !mod && selection.nodeIds.length === 1 && mode === 'edit')
    return { kind: 'enterText', nodeId: selection.nodeIds[0], type: key };
  return null;
}

/**
 * What Select All selects on a slide: all the text of the edited text body, or every slide
 * shape in the current scope (the slide, or the entered group). @internal
 */
export function selectAllOf(layout: SlideLayout, selection: EditorSelection): EditorSelection {
  if (selection.kind === 'text') {
    const lens = textItemOf(layout, selection)?.text?.paragraphLengths ?? [0];
    return {
      ...selection,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: lens.length - 1, offset: lens[lens.length - 1] ?? 0 },
    };
  }
  const group = selection.kind === 'shapes' ? selection.group : undefined;
  const ids = scopeItems(layout.items, group)
    .filter((i) => i.element.layer === 'slide')
    .map((i) => i.element.nodeId);
  return ids.length
    ? { kind: 'shapes', nodeIds: ids, ...(group ? { group } : {}) }
    : { kind: 'none' };
}
