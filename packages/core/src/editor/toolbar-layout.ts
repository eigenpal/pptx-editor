// Toolbar layout rules both adapters share: the named parts (`Toolbar.Bold` …), each group's
// collapse priority, and which groups leave the bar for the More menu when it runs out of width.
// After DOCX Editor's toolbar (docs/site/content/customization/toolbar.mdx: named parts,
// `priority`, `pinned`, `overflow`) and core's `packages/core/src/editor/toolbar-item-fit.ts`. DOCX
// Editor commit adf0e178.
//
// Every export is adapter glue, not a stable API: `@internal`.

import type { ChromeGroup, ChromeSlotId } from './chrome-controls.ts';

/**
 * The named toolbar parts and the slot each one stands for. `Toolbar.Bold` is shorthand for
 * `Toolbar.Slot slotId="text.bold"`, with the same props. The names follow DOCX Editor's parts
 * where the idea is the same (`Bold`, `FontFamily`, `BulletList`, `ClearFormatting`,
 * `ImageInsert`). @internal
 */
export const TOOLBAR_NAMED_PARTS = {
  Undo: 'history.undo',
  Redo: 'history.redo',
  FontFamily: 'font.family',
  FontSize: 'font.size',
  Bold: 'text.bold',
  Italic: 'text.italic',
  Underline: 'text.underline',
  Strike: 'text.strike',
  FontColor: 'text.color',
  Highlight: 'text.highlight',
  Link: 'text.link',
  ClearFormatting: 'format.clear',
  CharacterSpacing: 'text.characterSpacing',
  ChangeCase: 'text.changeCase',
  Font: 'text.dialog',
  Superscript: 'script.super',
  Subscript: 'script.sub',
  Alignment: 'paragraph.align',
  VerticalAlignment: 'paragraph.valign',
  BulletList: 'list.bullet',
  BulletStyle: 'list.bulletStyle',
  NumberedList: 'list.numbered',
  NumberStyle: 'list.numberStyle',
  Outdent: 'list.outdent',
  Indent: 'list.indent',
  LineSpacing: 'list.lineSpacing',
  Paragraph: 'paragraph.dialog',
  TextDirection: 'paragraph.textDirection',
  Autofit: 'paragraph.autofit',
  Columns: 'paragraph.columns',
  TextBox: 'insert.textBox',
  Line: 'insert.line',
  Shapes: 'insert.shape',
  ImageInsert: 'insert.image',
  Equation: 'insert.equation',
  ShapeFill: 'shape.fill',
  LineColor: 'shape.lineColor',
  LineWidth: 'shape.lineWidth',
  LineStart: 'shape.lineStart',
  LineEnd: 'shape.lineEnd',
  LineDash: 'shape.lineDash',
  ShapeStyles: 'shape.styles',
  ShapeEffects: 'shape.effects',
  TextEffects: 'shape.textEffects',
  FormatShape: 'shape.properties',
  FormatPainter: 'format.painter',
  Order: 'arrange.order',
  AlignObjects: 'arrange.align',
  GroupObjects: 'arrange.group',
  UngroupObjects: 'arrange.ungroup',
  Rotate: 'arrange.rotate',
  SelectionPane: 'arrange.selectionPane',
  Duplicate: 'more.duplicate',
  Delete: 'more.delete',
  Find: 'find.find',
  Replace: 'find.replace',
  SmartGuides: 'view.smartGuides',
  PageSetup: 'file.pageSetup',
  Cut: 'edit.cut',
  Copy: 'edit.copy',
  Paste: 'edit.paste',
  SelectAll: 'edit.selectAll',
  FormatBackground: 'slide.background',
  NewSlide: 'slide.add',
  SlideLayout: 'slide.layout',
  ResetSlide: 'slide.reset',
  HideSlide: 'slide.hide',
  PresentationView: 'view.presentationView',
  AddSection: 'section.add',
  RenameSection: 'section.rename',
  RemoveSection: 'section.remove',
  RemoveAllSections: 'section.removeAll',
  CollapseAllSections: 'section.collapseAll',
  ExpandAllSections: 'section.expandAll',
  PlayFromStart: 'view.playFromStart',
  PlayFromCurrent: 'view.playFromCurrent',
  ImageCorrections: 'image.corrections',
  ImageColor: 'image.color',
  ImageTransparency: 'image.transparency',
  ImageChange: 'image.change',
  ImageReset: 'image.reset',
  ImageStyles: 'image.styles',
  ImageAltText: 'image.altText',
  ImageCrop: 'image.crop',
  ImageCropToShape: 'image.cropToShape',
  ImageCropFit: 'image.cropFit',
  ImageProperties: 'image.properties',
  ChartEditData: 'chart.editData',
  DiagramAddAfter: 'diagram.addAfter',
  DiagramAddBefore: 'diagram.addBefore',
  DiagramAddBelow: 'diagram.addBelow',
  DiagramDelete: 'diagram.delete',
  AddComment: 'review.addComment',
  // The table group: DOCX Editor's names where it has the part (`TableInsert`,
  // `TableCellFill`, `TableBorderTarget`), and the same `Table` prefix for PPTX-only ideas.
  TableInsert: 'table.insert',
  TableStyle: 'table.style',
  TableHeaderRow: 'table.headerRow',
  TableTotalRow: 'table.totalRow',
  TableBandedRows: 'table.bandedRows',
  TableFirstColumn: 'table.firstColumn',
  TableLastColumn: 'table.lastColumn',
  TableBandedColumns: 'table.bandedColumns',
  TableCellFill: 'table.cellFill',
  TableBorderTarget: 'table.borderTarget',
  TableBorderColor: 'table.borderColor',
  TableBorderStyle: 'table.borderStyle',
  TableBorderWidth: 'table.borderWidth',
  TableSelect: 'table.select',
  TableDelete: 'table.delete',
  TableInsertRowAbove: 'table.insertRowAbove',
  TableInsertRowBelow: 'table.insertRowBelow',
  TableInsertColumnLeft: 'table.insertColumnLeft',
  TableInsertColumnRight: 'table.insertColumnRight',
  TableMergeCells: 'table.mergeCells',
  TableSplitCell: 'table.splitCell',
  TableDistributeRows: 'table.distributeRows',
  TableDistributeColumns: 'table.distributeColumns',
  TableCellMargins: 'table.cellMargins',
} as const satisfies Record<string, ChromeSlotId>;

/**
 * The slots whose named part is a compound, as DOCX Editor's `FontFamily`: the part's children
 * compose its sub-parts (`Toolbar.FontFamily.Trigger`, `.Content`, `.Item`) around the shared
 * state, instead of replacing the control. @internal
 */
export const TOOLBAR_COMPOUND_SLOTS: ReadonlySet<ChromeSlotId> = new Set<ChromeSlotId>([
  'font.family',
]);

/** The name of a named toolbar part. @internal */
export type ToolbarPartName = keyof typeof TOOLBAR_NAMED_PARTS;

/** The contextual group of a selected picture, which stays on the bar longest. */
function isPictureGroup(group: Pick<ChromeGroup, 'id' | 'contextual'>): boolean {
  return group.contextual === true && group.id === 'image';
}

/** The priority of a host group that sets none: it leaves the bar before every built-in group. */
const HOST_PRIORITY = 0;

/**
 * The collapse order of the built-in groups: the first id leaves the bar first. After DOCX Editor's
 * `TOOLBAR_COLLAPSE_ORDER` (packages/core/src/editor/toolbar-fit.ts at adf0e178), where zoom goes
 * first and text formatting and history last: the view and file controls, Find, the clipboard, and
 * the slide controls leave before the formatting groups, and the insert controls leave after the
 * formatting details, so the most used controls (font, size, bold, italic, underline, the paragraph
 * controls) stay on the bar the longest, on the toolbar and on every ribbon tab. @internal
 */
export const TOOLBAR_COLLAPSE_ORDER: readonly ChromeGroup['id'][] = [
  'view',
  'file',
  'section',
  'find',
  'edit',
  'more',
  'slide',
  'arrange',
  'format',
  'shape',
  'script',
  // Text box, shape, and picture are the core actions of a slide editor: after the formatting
  // details, before the font and paragraph controls.
  'insert',
  'list',
  'paragraph',
  // The contextual table group shows only while the selection is in a table: the table is the
  // task then, so its controls stay on the bar after the list and paragraph controls.
  'table',
  // The font pair leaves before the text formatting, as in DOCX Editor's order.
  'font',
  'text',
  'history',
];

/**
 * The collapse priority of a built-in group of `groups`: its place in `TOOLBAR_COLLAPSE_ORDER`
 * in steps of 10, so `view` is 10 and `history` is 180, as DOCX Editor's built-in groups use
 * steps of 10. The `image` group of a selected picture stays longest: the selection asked for
 * it. A group that is not in `groups` is a host group. A lower priority leaves the bar first.
 * @internal
 */
export function builtInGroupPriority(
  groupId: string,
  groups: readonly Pick<ChromeGroup, 'id' | 'contextual'>[]
): number {
  // The rank does not depend on the other groups of the bar, so a group's priority (and a host
  // priority set between two of them) does not change when the contextual `image` group joins.
  // DOCX Editor's contextual `image` group leaves first; PPTX Editor keeps it longest: a
  // selected picture makes picture formatting the main task.
  const group = groups.find((g) => g.id === groupId);
  if (!group) return HOST_PRIORITY;
  if (isPictureGroup(group)) return (TOOLBAR_COLLAPSE_ORDER.length + 1) * 10;
  const rank = TOOLBAR_COLLAPSE_ORDER.indexOf(group.id);
  return rank === -1 ? HOST_PRIORITY : (rank + 1) * 10;
}

/**
 * The priority a toolbar group collapses with: the host's `priority`, else the built-in
 * priority of a registered group, else that of a host group (first to go). @internal
 */
export function toolbarGroupPriority(
  groupId: string,
  priority: number | undefined,
  groups: readonly Pick<ChromeGroup, 'id' | 'contextual'>[]
): number {
  // Only NaN counts as unset: `Infinity` keeps a group longest, as in DOCX Editor.
  if (priority !== undefined && !Number.isNaN(priority)) return priority;
  return builtInGroupPriority(groupId, groups);
}

/** One group, as the collapse rule sees it. @internal */
export interface ToolbarCollapseItem {
  /** The group's width in CSS pixels, separator included. */
  readonly width: number;
  readonly priority: number;
  /** A pinned group stays on the bar at every width. */
  readonly pinned?: boolean;
}

/**
 * Which groups leave the bar for the More menu, by index. Nothing leaves when every group
 * fits. Otherwise groups leave in priority order, lowest first (at equal priority, the later
 * group first), until the rest and the More button fit. Pinned groups never leave. The
 * adapters pass the measured More button cost; 45 px is its default width with a separator.
 * @internal
 */
export function toolbarCollapse(
  items: readonly ToolbarCollapseItem[],
  inner: number,
  moreWidth = 45
): ReadonlySet<number> {
  let used = 0;
  for (const item of items) used += item.width;
  const hidden = new Set<number>();
  if (used <= inner) return hidden;
  used += moreWidth;
  const order = items
    .map((item, index) => ({ item, index }))
    .filter(({ item }) => !item.pinned)
    .sort((a, b) => a.item.priority - b.item.priority || b.index - a.index);
  for (const { item, index } of order) {
    if (used <= inner) break;
    hidden.add(index);
    used -= item.width;
  }
  // Keep the bar full: a group that fits in the room a wider group left comes back, highest
  // priority first (a greedy pass alone can leave half the bar empty).
  for (const { item, index } of [...order].reverse()) {
    if (!hidden.has(index) || used + item.width > inner) continue;
    hidden.delete(index);
    used += item.width;
  }
  return hidden;
}
