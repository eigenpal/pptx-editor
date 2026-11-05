// The ribbon arrangement: tabs as named sets of chrome groups, and the contextual tabs that show while a selection
// needs them. The ribbon is a layer over the toolbar: each tab panel is the toolbar's preset bar
// with that tab's groups, so a tab keeps the whole customization ladder. DOCX Editor has no
// ribbon; its equivalent of a tab is a `Toolbar` with `preset={false}` and that tab's groups.
//
// A tab names registry groups, or some controls of one group: the `slide` group's New Slide,
// Layout, and Reset are on Home, its Format Background on Design, and its Hide Slide on Slide
// Show. The ribbon adds no command and no slot: every button is a `ChromeSlotId` of
// `CHROME_GROUPS`, so a tab's button runs the same command as the toolbar's and the menus'.

import type { TranslationKey } from '@pptx-editor.dev/i18n';
import type { EditorSnapshot } from '../contracts/index.ts';
import { CHROME_GROUPS, type ChromeGroup, type ChromeGroupId } from './chrome-controls.ts';

/**
 * Every ribbon tab id. Stable public API, as `ChromeGroupId`. `quickAccess` is the Quick
 * Access Toolbar beside the tab names. @public
 */
export type RibbonTabId =
  | 'quickAccess'
  | 'home'
  | 'insert'
  | 'design'
  | 'slideShow'
  | 'view'
  | 'shapeFormat'
  | 'pictureFormat'
  | 'tableDesign'
  | 'tableLayout';

/** What a contextual tab waits for: shapes, a picture, or a table in the selection. @public */
export type RibbonContext = 'shape' | 'picture' | 'table';

/** Some controls of one registry group, in the given order. @public */
export interface RibbonGroupPick {
  readonly id: ChromeGroupId;
  /** Control ids of the group (`add` for `slide.add`). */
  readonly controls: readonly string[];
}

/**
 * One ribbon tab: its title and its groups, in order. A tab names each group once: the bar keys
 * its groups by id. @public
 */
export interface RibbonTab {
  readonly id: RibbonTabId;
  readonly labelKey: TranslationKey;
  /** Whole groups by id, or some controls of a group. */
  readonly groups: readonly (ChromeGroupId | RibbonGroupPick)[];
  /** A contextual tab shows only while its context holds. */
  readonly context?: RibbonContext;
}

/** The Paragraph group in the reference order: columns, alignment, then direction. */
const PARAGRAPH: RibbonGroupPick = {
  id: 'paragraph',
  controls: ['columns', 'align', 'textDirection', 'valign', 'autofit', 'dialog'],
};

/** The Arrange group in the reference order: order, Selection Pane, align, group. */
const ARRANGE: RibbonGroupPick = {
  id: 'arrange',
  controls: ['order', 'selectionPane', 'align', 'group', 'ungroup', 'rotate'],
};

/** The Insert group in the reference order: pictures, shapes, then text box. */
const INSERT: RibbonGroupPick = { id: 'insert', controls: ['image', 'shape', 'line', 'textBox'] };

/**
 * The ribbon, in the reference order. A group may sit on more than one tab:
 * Shape and Arrange are on Home and on Shape Format. @public
 */
export const RIBBON_TABS: readonly RibbonTab[] = [
  { id: 'quickAccess', labelKey: 'ribbon.tabs.quickAccess', groups: ['history'] },
  {
    id: 'home',
    labelKey: 'ribbon.tabs.home',
    groups: [
      { id: 'edit', controls: ['paste', 'cut', 'copy', 'selectAll'] },
      'format',
      { id: 'slide', controls: ['add', 'layout', 'reset'] },
      'font',
      'text',
      'script',
      // Lists and indents first, as the reference Paragraph group.
      'list',
      PARAGRAPH,
      INSERT,
      'shape',
      ARRANGE,
      'more',
      'find',
    ],
  },
  {
    id: 'insert',
    labelKey: 'ribbon.tabs.insert',
    groups: [{ id: 'slide', controls: ['add'] }, { id: 'table', controls: ['insert'] }, INSERT],
  },
  {
    id: 'design',
    labelKey: 'ribbon.tabs.design',
    groups: ['file', { id: 'slide', controls: ['background'] }],
  },
  {
    id: 'slideShow',
    labelKey: 'ribbon.tabs.slideShow',
    groups: [
      { id: 'view', controls: ['playFromStart', 'playFromCurrent'] },
      { id: 'slide', controls: ['hide'] },
    ],
  },
  {
    id: 'view',
    labelKey: 'ribbon.tabs.view',
    groups: [{ id: 'view', controls: ['presentationView', 'smartGuides'] }],
  },
  {
    id: 'shapeFormat',
    labelKey: 'ribbon.tabs.shapeFormat',
    groups: [{ id: 'insert', controls: ['shape', 'line', 'textBox'] }, 'shape', ARRANGE],
    context: 'shape',
  },
  {
    id: 'pictureFormat',
    labelKey: 'ribbon.tabs.pictureFormat',
    groups: ['image', ARRANGE],
    context: 'picture',
  },
  {
    id: 'tableDesign',
    labelKey: 'ribbon.tabs.tableDesign',
    groups: [
      {
        id: 'table',
        controls: [
          'headerRow',
          'firstColumn',
          'totalRow',
          'lastColumn',
          'bandedRows',
          'bandedColumns',
          'style',
          'cellFill',
          'borderTarget',
        ],
      },
    ],
    context: 'table',
  },
  {
    id: 'tableLayout',
    labelKey: 'ribbon.tabs.tableLayout',
    groups: [
      {
        id: 'table',
        controls: [
          'select',
          'delete',
          'insertRowAbove',
          'insertRowBelow',
          'insertColumnLeft',
          'insertColumnRight',
          'mergeCells',
          'splitCell',
          'distributeRows',
          'distributeColumns',
          'cellMargins',
        ],
      },
    ],
    context: 'table',
  },
];

/** A tab as the ribbon shows it now: its entry and its registered groups. @public */
export interface RibbonTabLayout {
  readonly tab: RibbonTab;
  readonly groups: readonly ChromeGroup[];
}

/** What a contextual tab reads from the snapshot. @public */
export type RibbonSelection = Pick<EditorSnapshot, 'selection' | 'image' | 'table'>;

/** Whether a contextual tab's context holds for `snapshot`. @public */
export function ribbonContextActive(snapshot: RibbonSelection, context: RibbonContext): boolean {
  switch (context) {
    case 'shape':
      // Shapes are selected, or the text of a shape is being edited, as in the reference
      // application. Pictures alone show Picture Format, and a table shows its two tabs.
      return (
        (snapshot.selection.kind === 'shapes' || snapshot.selection.kind === 'text') &&
        snapshot.image == null &&
        !snapshot.table
      );
    case 'table':
      // The selection is in or on a table (`tableChromeVisible`).
      return !!snapshot.table;
    case 'picture':
      // Every selected element is a picture.
      return snapshot.image != null;
  }
}

/**
 * The groups of `tab` that the registry holds, in the tab's order. A picked group keeps only
 * the picked controls. An id or a control the registry does not hold is skipped. @public
 */
export function ribbonTabGroups(
  tab: RibbonTab,
  groups: readonly ChromeGroup[] = CHROME_GROUPS
): ChromeGroup[] {
  const byId = new Map(groups.map((g) => [g.id as string, g]));
  return tab.groups.flatMap((entry) => {
    const group = byId.get(typeof entry === 'string' ? entry : entry.id);
    if (!group) return [];
    if (typeof entry === 'string') return [group];
    const picked = entry.controls.flatMap((id) => group.controls.filter((c) => c.id === id));
    return picked.length ? [{ ...group, controls: picked }] : [];
  });
}

/**
 * The tabs to show for `snapshot`, with their groups: every tab with groups, and a contextual
 * tab only while its context holds. The Quick Access Toolbar comes first. @public
 */
export function ribbonTabs(
  snapshot: RibbonSelection,
  groups: readonly ChromeGroup[] = CHROME_GROUPS
): RibbonTabLayout[] {
  return RIBBON_TABS.flatMap((tab) => {
    if (tab.context && !ribbonContextActive(snapshot, tab.context)) return [];
    const tabGroups = ribbonTabGroups(tab, groups);
    return tabGroups.length ? [{ tab, groups: tabGroups }] : [];
  });
}
