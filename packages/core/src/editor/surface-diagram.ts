// The caret in a diagram shape. A diagram frame draws its saved drawing; the drawing shapes
// that show one node's text take the caret (`EditorSelection.diagramNode`). A press on such a
// shape of the selected diagram puts the caret there; the first press selects the frame, so a
// drag moves it. Enter on a selected diagram puts the caret in its first node, and Tab and
// Shift+Tab move the caret from node to node, as from cell to cell in a table.

import { apply, invert, type LayoutItem, type SlideLayout } from '../layout/index.ts';
import type { EditorSelection, TextPoint } from '../contracts/index.ts';
import { findItem } from '../automation/commands/placement.ts';
import { pointAt } from './surface-geometry.ts';
import type { ShapeItem } from './table-cells.ts';

type Point = { readonly x: number; readonly y: number };

/** The drawing shapes of a diagram frame that show a node's text, in drawing order. */
export function diagramTextItems(layout: SlideLayout | undefined, frameId: string): ShapeItem[] {
  const frame = layout ? findItem(layout.items, frameId) : undefined;
  const out: ShapeItem[] = [];
  const walk = (items: readonly LayoutItem[] | undefined): void => {
    for (const item of items ?? []) {
      if (item.kind === 'shape' && item.element.kind === 'shape' && item.element.diagramNode)
        out.push(item);
      if (item.kind === 'group') walk(item.children);
    }
  };
  if (frame?.kind === 'frame') walk(frame.drawing);
  return out;
}

/** The end of a shape's text. */
function endOf(item: ShapeItem): TextPoint {
  const lens = item.text?.paragraphLengths ?? [0];
  return { paragraph: Math.max(0, lens.length - 1), offset: lens[lens.length - 1] ?? 0 };
}

function caretIn(frameId: string, item: ShapeItem, at: TextPoint, anchor?: TextPoint) {
  const node = item.element.kind === 'shape' ? (item.element.diagramNode ?? '') : '';
  return {
    kind: 'text',
    nodeId: frameId,
    diagramNode: node,
    anchor: anchor ?? at,
    focus: at,
  } as const;
}

/**
 * The caret a press at `p` on the diagram frame `frameId` places, or null: the frame must be
 * selected already (or hold the caret), and the press must be on a shape that shows a node.
 */
export function diagramPress(
  layout: SlideLayout,
  frameId: string,
  sel: EditorSelection,
  p: Point,
  extend: boolean
): EditorSelection | null {
  const selected =
    (sel.kind === 'shapes' && sel.nodeIds.length === 1 && sel.nodeIds[0] === frameId) ||
    (sel.kind === 'text' && sel.nodeId === frameId && sel.diagramNode !== undefined);
  if (!selected) return null;
  const items = diagramTextItems(layout, frameId);
  for (let i = items.length - 1; i >= 0; i -= 1) {
    const item = items[i]!;
    const local = apply(invert(item.matrix), p.x, p.y);
    const { cx, cy } = item.element.xfrm;
    if (local.x < 0 || local.y < 0 || local.x > cx || local.y > cy) continue;
    const at = pointAt(item, p.x, p.y) ?? endOf(item);
    const same =
      extend && sel.kind === 'text' && item.element.kind === 'shape'
        ? sel.diagramNode === item.element.diagramNode
        : false;
    return caretIn(frameId, item, at, same && sel.kind === 'text' ? sel.anchor : undefined);
  }
  return null;
}

/** The caret at the end of the first node of a diagram frame, or undefined. */
export function firstDiagramCaret(
  layout: SlideLayout | undefined,
  frameId: string
): EditorSelection | undefined {
  const first = diagramTextItems(layout, frameId)[0];
  return first ? caretIn(frameId, first, endOf(first)) : undefined;
}

/**
 * Tab in a diagram's text: the caret at the end of the next node (Shift+Tab: the previous one),
 * or null when the selection is not in a diagram. The last node wraps to the first.
 */
export function diagramTab(
  layout: SlideLayout | undefined,
  sel: EditorSelection,
  back: boolean
): EditorSelection | null {
  if (sel.kind !== 'text' || sel.diagramNode === undefined) return null;
  const items = diagramTextItems(layout, sel.nodeId);
  if (items.length === 0) return null;
  const at = items.findIndex(
    (i) => i.element.kind === 'shape' && i.element.diagramNode === sel.diagramNode
  );
  const next = items[(at + (back ? -1 : 1) + items.length) % items.length]!;
  return caretIn(sel.nodeId, next, endOf(next));
}
