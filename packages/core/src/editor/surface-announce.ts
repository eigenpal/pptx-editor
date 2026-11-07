// What the slide's live region says about the selection: the table cell or the diagram shape the
// caret is in, or the kind, name, and alt text of the one shape that Tab (or a click) selected.
// Split from `surface.ts` for the line cap.

import type { EditorSelection } from '../contracts/index.ts';
import type { LayoutItem, SlideLayout } from '../layout/index.ts';
import type { AltTextPort } from '../output/index.ts';
import { findItem } from '../automation/commands/placement.ts';
import { diagramTextItems } from './surface-diagram.ts';

/** What the announcement asks of the editor (a part of `SurfaceHost`). */
export interface AnnouncementHost {
  readonly selection: () => EditorSelection;
  readonly label: (key: string, params?: Record<string, string | number>) => string;
  readonly altText?: AltTextPort;
}

/** The i18n key of an item's kind: the Format panel's element kinds, and the frame kinds. */
function kindKey(item: LayoutItem): string {
  if (item.kind === 'frame') return `canvas.kinds.${item.element.frameKind}`;
  if (item.kind === 'shape') return `panel.format.elementKind.${item.element.kind}`;
  return `panel.format.elementKind.${item.kind}`;
}

/**
 * Whether a name starts with the kind as a whole word ("Picture 3", "Table 1"), so the kind is
 * said once. "Tablet photo" does not start with the word "Table".
 */
function startsWithWord(name: string, kind: string): boolean {
  if (!kind || !name.toLocaleLowerCase().startsWith(kind.toLocaleLowerCase())) return false;
  return !/^\p{L}/u.test(name.slice(kind.length));
}

/**
 * The live region text for a selection, or an empty string. A shape is said as its kind, its
 * name, and its description, unless it is decorative or has none. Spaces collapse, so an empty
 * kind leaves no gap.
 */
export function selectionAnnouncement(layout: SlideLayout, host: AnnouncementHost): string {
  const { label, altText } = host;
  const selection = host.selection();
  const cell =
    selection.kind === 'text'
      ? selection.cell
      : selection.kind === 'cells'
        ? selection.focus
        : undefined;
  if (cell) return label('canvas.tableCell', { row: cell.row + 1, column: cell.column + 1 });
  // The caret in a diagram: which shape of the diagram it is in.
  if (selection.kind === 'text' && selection.diagramNode !== undefined) {
    const shapes = diagramTextItems(layout, selection.nodeId);
    const at = shapes.findIndex(
      (i) => i.element.kind === 'shape' && i.element.diagramNode === selection.diagramNode
    );
    return at < 0 ? '' : label('canvas.diagramShape', { index: at + 1, count: shapes.length });
  }
  const one =
    selection.kind === 'shapes' && selection.nodeIds.length === 1
      ? selection.nodeIds[0]
      : undefined;
  const item = one ? findItem(layout.items, one) : undefined;
  // `p:cNvPr/@name` is required: an element without one is not said.
  if (!item?.element.name) return '';
  const alt = altText?.(item.element.partName, item.element.nodeId);
  const description = alt && !alt.decorative ? alt.description.trim() : '';
  const name = item.element.name;
  const kind = label(kindKey(item));
  const params = { kind: startsWithWord(name, kind) ? '' : kind, name };
  const said = description
    ? label('canvas.shapeSelectedWithAltText', { ...params, description })
    : label('canvas.shapeSelected', params);
  return said.replace(/\s+/g, ' ').trim();
}
