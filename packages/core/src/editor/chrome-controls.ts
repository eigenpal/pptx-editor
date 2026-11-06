// The chrome registry: the toolbar taxonomy both the React adapter and any future adapter
// derive their default arrangement FROM (docx-editor/packages/core/src/editor/chrome-controls.ts).
//
// `ChromeSlotId` values are public API: renames are breaking. A control's descriptor says HOW
// it renders and dispatches, never whether it is enabled. Enabled state has exactly one source,
// `toolbarCommandState`, which asks `Editor.can`.

import type {
  Editor,
  EditorCommand,
  EditorExecOptions,
  SelectedImageState,
} from '../contracts/index.ts';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { warnChromeSlotRun, warnDialogSlotRun, warnInvalidSlotValue } from './dev-warnings.ts';
import { ICON_PATHS } from './icon-paths.ts';
import { IMAGE_CHROME_GROUP, imageSlotState } from './chrome-picture.ts';
import { adapterActionState } from './chrome-adapter-state.ts';
import type { ImageSlotId } from './chrome-picture.ts';
import { isNothingToDo } from '../automation/commands/op-batch.ts';
import type { ToolbarSlotValue } from './toolbar-values.ts';
import {
  DISTRIBUTE_OPTIONS,
  QUICK_STYLES_CONTROL,
  ROTATE_CONTROL,
  SELECTION_PANE_CONTROL,
  SHAPE_FORMAT_CONTROLS,
  SHAPE_FORMAT_PROBES,
  shapeFormatSlotValue,
} from './chrome-shapes.ts';
import {
  CHROME_SLOT_REASONS,
  SECTION_GROUP,
  SLIDE_STRUCTURE_CONTROLS,
  slideSlotActive,
  slideSlotProbe,
  slideSlotValue,
} from './chrome-slide-groups.ts';
import {
  type TableSlotId,
  TABLE_GROUPS,
  TABLE_PROBES,
  armTableBorderPen,
  rememberTableBorderPen,
  tableBorderPen,
  tableSlotState,
  tableSlotActive,
  tableSlotValue,
} from './chrome-table.ts';
import {
  LINE_DASHES,
  LINE_ENDS,
  LINE_KINDS,
  LINE_WIDTHS,
  commandForSlot,
  commandForSlotValue,
} from './chrome-slot-commands.ts';
import {
  BULLET_STYLE_CONTROL,
  NUMBER_STYLE_CONTROL,
  TEXT_DIRECTION_CONTROL,
  AUTOFIT_CONTROL,
  COLUMNS_CONTROL,
  LINE_SPACING_CONTROL,
  PARAGRAPH_DIALOG_CONTROL,
  TEXT_FORMAT_PROBES,
  fontSizeSlotValue,
  lineSpacingSlotValue,
  textSlotValue,
} from './chrome-paragraph-controls.ts';
import {
  CHANGE_CASE_CONTROL,
  CHARACTER_SPACING_CONTROL,
  FONT_DIALOG_CONTROL,
  SCRIPT_GROUP,
} from './chrome-font-controls.ts';

export { commandForSlot, commandForSlotValue } from './chrome-slot-commands.ts';

/** How a control dispatches. @public */
export type ChromeControlState =
  /** One fixed command: `commandForSlot(slot)`. */
  | { readonly kind: 'command' }
  /** A picked value: `commandForSlotValue(slot, value)`. */
  | { readonly kind: 'value' }
  /** Bytes from the host (an image file) to insert: `insertImage`. */
  | { readonly kind: 'load' }
  /**
   * Bytes from the host (an image file) that replace the selected picture's image: Change
   * Picture, `replaceImage`. A kind apart from `load` because the two dispatch differently, and
   * an adapter branching on `kind` must tell them apart (as DOCX Editor's `load` JSDoc asks).
   */
  | { readonly kind: 'replace' }
  /**
   * Opens a panel the adapter owns: the find bar. The adapter dispatches a `pptx-find-request`
   * event with `{ editor, replace }`, and the panel answers it.
   */
  | { readonly kind: 'panel' }
  /** Opens a task pane: the Format pane or the selection pane. Neither is built yet. */
  | { readonly kind: 'pane'; readonly pane: 'format' | 'selection' }
  /**
   * Opens a dialog of the `popups` registry by its key. With a probe command, the slot is on when
   * the engine takes it; the dialog runs the command.
   */
  | { readonly kind: 'dialog'; readonly dialog: ChromeDialogKey }
  /**
   * A clipboard action the adapter runs (cut, copy, paste, select all): not an editing
   * command, because a browser gives the system clipboard to a page only on a user gesture.
   */
  | { readonly kind: 'clipboard' }
  /** Starts the adapter's slide show, from the first shown slide or the current slide. */
  | { readonly kind: 'slideshow'; readonly from: 'start' | 'current' };

/** The registry dialogs a control, a dropdown row, or a menu row opens. @public */
export type ChromeDialogKey =
  | 'pageSetup'
  | 'background'
  | 'paragraph'
  | 'font'
  | 'bulletsAndNumbering'
  | 'columns'
  | 'imageAltText'
  | 'imageProperties';

/** How a control renders. @public */
export type ChromeControlShape =
  | 'icon'
  | 'stepper'
  | 'dropdown'
  | 'colorSplit'
  | 'gallery'
  /** A rows × columns grid to pick a table size (Insert > Table). */
  | 'tableGrid';

/**
 * How a control shows its state. `pressed`: a two-state button that carries `aria-pressed`
 * (bold, bullets). `choice`: a dropdown of exclusive options, where the option equal to the
 * slot's value is checked (`menuitemradio`). A control without it runs an action and shows no
 * state. @public
 */
export type ChromeControlToggle = 'pressed' | 'choice';

/** One toolbar control. @public */
export interface ChromeControl {
  readonly id: string;
  readonly shape?: ChromeControlShape;
  /** i18n key for the accessible name and tooltip. */
  readonly labelKey: TranslationKey;
  /** Material Symbols path data, or null for a picker. */
  readonly paths: readonly string[] | null;
  /** How the control shows its state; absent for an action. */
  readonly toggle?: ChromeControlToggle;
  /**
   * `false` keeps this public slot composable (`Toolbar.Slot` in a host group or with
   * `preset={false}`, menus, the ribbon) but leaves it out of the default formatting bar, as in
   * DOCX Editor.
   */
  readonly defaultToolbar?: false;
  /** Dropdown options: value plus i18n label key (and optional icon). */
  readonly options?: readonly {
    readonly value: string;
    readonly labelKey: TranslationKey;
    /** Parameters for the option's label (`{value} pt`). */
    readonly labelParams?: Readonly<Record<string, string | number>>;
    readonly paths?: readonly string[];
    /** A text preview drawn before the label, where an icon would go (a gallery bullet). */
    readonly glyph?: string;
    /** A row that opens this registry dialog ("Line Spacing Options…"); never checked. */
    readonly dialog?: ChromeDialogKey;
    /** The `optionGroups` id of the submenu the option is in. */
    readonly group?: string;
  }[];
  /**
   * Submenus of a dropdown whose options have a `group` (Shape Effects: Shadow, Reflection,
   * Glow, Soft Edges). The value names the chosen option of each group.
   */
  readonly optionGroups?: readonly { readonly id: string; readonly labelKey: TranslationKey }[];
  readonly state: ChromeControlState;
}

/**
 * The ids of the registered chrome groups. A host passes one to `Toolbar.Group` to change that
 * group, as with DOCX Editor's `ChromeGroupId`. @public
 */
export type ChromeGroupId =
  | 'history'
  | 'font'
  | 'text'
  | 'script'
  | 'paragraph'
  | 'list'
  | 'insert'
  | 'shape'
  | 'format'
  | 'arrange'
  | 'more'
  | 'find'
  | 'view'
  | 'file'
  | 'edit'
  | 'slide'
  | 'image'
  | 'section'
  | 'table';

/** One toolbar group. @public */
export interface ChromeGroup {
  readonly id: ChromeGroupId;
  readonly labelKey: TranslationKey;
  /**
   * A group the formatting bar does not show: a menu places its slots (View > Smart guides).
   * The same flag as DOCX Editor's registry.
   */
  readonly contextual?: boolean;
  readonly controls: readonly ChromeControl[];
}

const icon = (name: string): readonly string[] => ICON_PATHS[name] ?? [];

/** The complete chrome, in bar order. @public */
export const CHROME_GROUPS: readonly ChromeGroup[] = [
  {
    id: 'history',
    labelKey: 'toolbar.groups.history',
    controls: [
      { id: 'undo', labelKey: 'toolbar.undo', paths: icon('undo'), state: { kind: 'command' } },
      { id: 'redo', labelKey: 'toolbar.redo', paths: icon('redo'), state: { kind: 'command' } },
    ],
  },
  {
    // The font pair, as DOCX Editor's `font` group: `font.family` and `font.size`.
    id: 'font',
    labelKey: 'toolbar.groups.font',
    controls: [
      {
        id: 'family',
        shape: 'dropdown',
        labelKey: 'toolbar.fontFamily',
        toggle: 'choice',
        paths: null,
        state: { kind: 'value' },
      },
      {
        id: 'size',
        shape: 'stepper',
        labelKey: 'toolbar.fontSize',
        paths: null,
        state: { kind: 'value' },
      },
    ],
  },
  {
    id: 'text',
    labelKey: 'toolbar.groups.text',
    controls: [
      {
        id: 'bold',
        labelKey: 'toolbar.bold',
        toggle: 'pressed',
        paths: icon('format_bold'),
        state: { kind: 'command' },
      },
      {
        id: 'italic',
        labelKey: 'toolbar.italic',
        toggle: 'pressed',
        paths: icon('format_italic'),
        state: { kind: 'command' },
      },
      {
        id: 'underline',
        labelKey: 'toolbar.underline',
        toggle: 'pressed',
        paths: icon('format_underlined'),
        state: { kind: 'command' },
      },
      {
        id: 'strike',
        labelKey: 'toolbar.strikethrough',
        toggle: 'pressed',
        paths: icon('strikethrough_s'),
        state: { kind: 'command' },
      },
      CHARACTER_SPACING_CONTROL,
      CHANGE_CASE_CONTROL,
      {
        id: 'color',
        shape: 'colorSplit',
        labelKey: 'toolbar.color',
        paths: icon('format_color_text'),
        state: { kind: 'value' },
      },
      {
        id: 'highlight',
        shape: 'colorSplit',
        labelKey: 'toolbar.highlight',
        paths: icon('ink_highlighter'),
        state: { kind: 'value' },
      },
      {
        // A link popover: the URL, apply, remove, and open.
        id: 'link',
        labelKey: 'toolbar.link',
        paths: icon('link'),
        state: { kind: 'value' },
      },
      FONT_DIALOG_CONTROL,
    ],
  },
  SCRIPT_GROUP,
  {
    id: 'paragraph',
    labelKey: 'toolbar.groups.paragraph',
    controls: [
      {
        id: 'align',
        shape: 'dropdown',
        labelKey: 'toolbar.alignLeft',
        toggle: 'choice',
        paths: icon('format_align_left'),
        options: [
          { value: 'l', labelKey: 'toolbar.alignLeft', paths: icon('format_align_left') },
          { value: 'ctr', labelKey: 'toolbar.alignCenter', paths: icon('format_align_center') },
          { value: 'r', labelKey: 'toolbar.alignRight', paths: icon('format_align_right') },
          { value: 'just', labelKey: 'toolbar.alignJustify', paths: icon('format_align_justify') },
        ],
        state: { kind: 'value' },
      },
      TEXT_DIRECTION_CONTROL,
      {
        id: 'valign',
        shape: 'dropdown',
        labelKey: 'toolbar.verticalAlign',
        toggle: 'choice',
        paths: icon('vertical_align_top'),
        options: [
          { value: 't', labelKey: 'toolbar.valign.t', paths: icon('vertical_align_top') },
          { value: 'ctr', labelKey: 'toolbar.valign.ctr', paths: icon('vertical_align_center') },
          { value: 'b', labelKey: 'toolbar.valign.b', paths: icon('vertical_align_bottom') },
        ],
        state: { kind: 'value' },
      },
      AUTOFIT_CONTROL,
      COLUMNS_CONTROL,
      PARAGRAPH_DIALOG_CONTROL,
    ],
  },
  {
    // Lists, indents, and line spacing, as DOCX Editor's `list` group (`list.bullet`,
    // `list.numbered`, `list.outdent`, `list.indent`, `list.lineSpacing`). The bullet and number
    // galleries sit after their toggles.
    id: 'list',
    labelKey: 'toolbar.groups.list',
    controls: [
      {
        id: 'bullet',
        labelKey: 'toolbar.bullets',
        toggle: 'pressed',
        paths: icon('format_list_bulleted'),
        state: { kind: 'command' },
      },
      BULLET_STYLE_CONTROL,
      {
        id: 'numbered',
        labelKey: 'toolbar.numbering',
        toggle: 'pressed',
        paths: icon('format_list_numbered'),
        state: { kind: 'command' },
      },
      NUMBER_STYLE_CONTROL,
      {
        id: 'outdent',
        labelKey: 'toolbar.indentDecrease',
        paths: icon('format_indent_decrease'),
        state: { kind: 'command' },
      },
      {
        id: 'indent',
        labelKey: 'toolbar.indentIncrease',
        paths: icon('format_indent_increase'),
        state: { kind: 'command' },
      },
      LINE_SPACING_CONTROL,
    ],
  },
  {
    id: 'insert',
    labelKey: 'toolbar.groups.insert',
    controls: [
      {
        id: 'textBox',
        labelKey: 'toolbar.textBox',
        paths: icon('text_fields'),
        state: { kind: 'command' },
      },
      {
        // Lines and connectors; the chrome draws each option as the line it inserts.
        id: 'line',
        shape: 'dropdown',
        labelKey: 'toolbar.line',
        paths: null,
        options: LINE_KINDS.map((k) => ({ value: k, labelKey: `toolbar.lines.${k}` as const })),
        state: { kind: 'value' },
      },
      {
        // A gallery of every preset in `SHAPE_GALLERY`, drawn from the spec geometry.
        id: 'shape',
        shape: 'gallery',
        labelKey: 'toolbar.shape',
        paths: icon('shapes'),
        state: { kind: 'value' },
      },
      { id: 'image', labelKey: 'toolbar.image', paths: icon('image'), state: { kind: 'load' } },
    ],
  },
  {
    id: 'shape',
    labelKey: 'toolbar.groups.shape',
    controls: [
      QUICK_STYLES_CONTROL,
      {
        id: 'fill',
        shape: 'colorSplit',
        labelKey: 'toolbar.fill',
        paths: icon('format_color_fill'),
        state: { kind: 'value' },
      },
      {
        id: 'lineColor',
        shape: 'colorSplit',
        labelKey: 'toolbar.lineColor',
        paths: icon('border_color'),
        state: { kind: 'value' },
      },
      {
        id: 'lineWidth',
        shape: 'dropdown',
        labelKey: 'toolbar.lineWidth',
        toggle: 'choice',
        paths: icon('line_weight'),
        options: LINE_WIDTHS.map((w) => ({
          value: String(w),
          labelKey: 'toolbar.lineWidthValue',
          labelParams: { value: w },
        })),
        state: { kind: 'value' },
      },
      {
        id: 'lineStart',
        shape: 'dropdown',
        labelKey: 'toolbar.lineStart',
        toggle: 'choice',
        paths: null,
        options: LINE_ENDS.map((e) => ({ value: e, labelKey: `toolbar.lineEnds.${e}` as const })),
        state: { kind: 'value' },
      },
      {
        id: 'lineEnd',
        shape: 'dropdown',
        labelKey: 'toolbar.lineEnd',
        toggle: 'choice',
        paths: null,
        options: LINE_ENDS.map((e) => ({ value: e, labelKey: `toolbar.lineEnds.${e}` as const })),
        state: { kind: 'value' },
      },
      {
        id: 'lineDash',
        shape: 'dropdown',
        labelKey: 'toolbar.lineDash',
        toggle: 'choice',
        paths: icon('horizontal_rule'),
        options: LINE_DASHES.map((d) => ({ value: d, labelKey: `toolbar.dashes.${d}` as const })),
        state: { kind: 'value' },
      },
      ...SHAPE_FORMAT_CONTROLS,
    ],
  },
  {
    // The Format Painter copies text and shape formatting, so it sits after the formatting
    // groups. One press paints once; a double-click keeps it on. The engine tells the two
    // apart (format-painter.ts), so the control dispatches one command per click and reports
    // the mode as its value.
    id: 'format',
    labelKey: 'toolbar.groups.format',
    controls: [
      {
        id: 'painter',
        labelKey: 'toolbar.formatPainter',
        toggle: 'pressed',
        paths: icon('format_paint'),
        state: { kind: 'command' },
      },
      {
        // DOCX Editor's `format.clear`: the formatting pair, painter and clear.
        id: 'clear',
        labelKey: 'toolbar.clearFormatting',
        paths: icon('format_clear'),
        state: { kind: 'command' },
      },
    ],
  },
  {
    id: 'arrange',
    labelKey: 'toolbar.groups.arrange',
    controls: [
      {
        id: 'order',
        shape: 'dropdown',
        labelKey: 'toolbar.bringForward',
        paths: icon('flip_to_front'),
        options: [
          { value: 'forward', labelKey: 'toolbar.bringForward', paths: icon('flip_to_front') },
          { value: 'backward', labelKey: 'toolbar.sendBackward', paths: icon('flip_to_back') },
          { value: 'front', labelKey: 'toolbar.bringToFront', paths: icon('flip_to_front') },
          { value: 'back', labelKey: 'toolbar.sendToBack', paths: icon('flip_to_back') },
        ],
        state: { kind: 'value' },
      },
      {
        id: 'align',
        shape: 'dropdown',
        labelKey: 'toolbar.alignObjects',
        paths: icon('align_horizontal_center'),
        options: [
          ...(['left', 'center', 'right', 'top', 'middle', 'bottom'] as const).map((v) => ({
            value: v,
            labelKey: `toolbar.objectAlign.${v}` as const,
          })),
          ...DISTRIBUTE_OPTIONS,
        ],
        state: { kind: 'value' },
      },
      ROTATE_CONTROL,
      {
        id: 'group',
        labelKey: 'toolbar.group',
        paths: icon('call_merge'),
        state: { kind: 'command' },
      },
      {
        id: 'ungroup',
        labelKey: 'toolbar.ungroup',
        paths: icon('call_split'),
        state: { kind: 'command' },
      },
      SELECTION_PANE_CONTROL,
    ],
  },
  {
    id: 'more',
    labelKey: 'toolbar.groups.more',
    controls: [
      {
        id: 'duplicate',
        labelKey: 'toolbar.duplicate',
        paths: icon('content_copy'),
        state: { kind: 'command' },
      },
      {
        id: 'delete',
        labelKey: 'toolbar.delete',
        paths: icon('delete'),
        state: { kind: 'command' },
      },
    ],
  },
  {
    id: 'find',
    labelKey: 'toolbar.groups.find',
    controls: [
      { id: 'find', labelKey: 'toolbar.find', paths: icon('search'), state: { kind: 'panel' } },
      {
        id: 'replace',
        labelKey: 'toolbar.replace',
        paths: icon('swap_horiz'),
        state: { kind: 'panel' },
      },
    ],
  },
  {
    id: 'view',
    labelKey: 'toolbar.groups.view',
    contextual: true,
    controls: [
      {
        // View > Normal or Slide Sorter: the Presentation Views, first on the View tab.
        id: 'presentationView',
        shape: 'dropdown',
        labelKey: 'toolbar.presentationView',
        toggle: 'choice',
        paths: icon('grid_view'),
        options: [
          { value: 'normal', labelKey: 'toolbar.views.normal', paths: icon('view_sidebar') },
          { value: 'slideSorter', labelKey: 'toolbar.views.slideSorter', paths: icon('grid_view') },
        ],
        state: { kind: 'value' },
      },
      {
        id: 'smartGuides',
        labelKey: 'toolbar.smartGuides',
        paths: icon('align_horizontal_center'),
        toggle: 'pressed',
        state: { kind: 'command' },
      },
      // The slide show plays the presentation: a view, as in DOCX Editor's `view.*` group.
      {
        id: 'playFromStart',
        labelKey: 'menu.presentFromStart',
        paths: icon('slideshow'),
        state: { kind: 'slideshow', from: 'start' },
      },
      {
        id: 'playFromCurrent',
        labelKey: 'menu.presentFromCurrent',
        paths: icon('play_arrow'),
        state: { kind: 'slideshow', from: 'current' },
      },
    ],
  },
  // The groups below are menu and ribbon controls: the flat formatting bar does not show them.
  {
    id: 'file',
    labelKey: 'toolbar.groups.file',
    contextual: true,
    controls: [
      {
        id: 'pageSetup',
        labelKey: 'menu.pageSetup',
        paths: icon('description'),
        state: { kind: 'dialog', dialog: 'pageSetup' },
      },
    ],
  },
  {
    id: 'edit',
    labelKey: 'toolbar.groups.edit',
    contextual: true,
    controls: [
      { id: 'cut', labelKey: 'menu.cut', paths: icon('content_cut'), state: { kind: 'clipboard' } },
      {
        id: 'copy',
        labelKey: 'menu.copy',
        paths: icon('content_copy'),
        state: { kind: 'clipboard' },
      },
      {
        id: 'paste',
        labelKey: 'menu.paste',
        paths: icon('content_paste'),
        state: { kind: 'clipboard' },
      },
      {
        id: 'selectAll',
        labelKey: 'menu.selectAll',
        paths: icon('select_all'),
        state: { kind: 'clipboard' },
      },
    ],
  },
  {
    // The slide object (Office.js `Slide`): Format Background, new slide, layout, reset, and hide.
    // Themes and slide size come later.
    id: 'slide',
    labelKey: 'toolbar.groups.slide',
    contextual: true,
    controls: [
      {
        id: 'background',
        labelKey: 'menu.background',
        paths: icon('wallpaper'),
        state: { kind: 'dialog', dialog: 'background' },
      },
      ...SLIDE_STRUCTURE_CONTROLS,
    ],
  },
  SECTION_GROUP,
  ...TABLE_GROUPS,
  IMAGE_CHROME_GROUP,
];

// A group with no control left on the bar (`script`) is left out.
const BAR_GROUPS: readonly ChromeGroup[] = CHROME_GROUPS.filter((g) => !g.contextual)
  .map((g) => ({ ...g, controls: g.controls.filter((c) => c.defaultToolbar !== false) }))
  .filter((g) => g.controls.length > 0);
/** The formatting bar while pictures are selected. Not exported from the package. @internal */
export const PICTURE_BAR_GROUPS: readonly ChromeGroup[] = [...BAR_GROUPS, IMAGE_CHROME_GROUP];

/**
 * The groups the formatting bar shows: every group that is not `contextual`, without the
 * controls marked `defaultToolbar: false`, and the contextual `image` group while every selected
 * element is a picture (`EditorSnapshot.image`), as DOCX Editor's bar adds its `image` group for
 * a selected drawing. The same array comes back for the same answer. @public
 */
export function formattingBarChromeGroups(
  image?: SelectedImageState | null
): readonly ChromeGroup[] {
  return image ? PICTURE_BAR_GROUPS : BAR_GROUPS;
}

/** Every public slot id: `<group>.<control>`. @public */
export type ChromeSlotId =
  | 'history.undo'
  | 'history.redo'
  | 'font.family'
  | 'font.size'
  | 'text.bold'
  | 'text.italic'
  | 'text.underline'
  | 'text.strike'
  | 'text.color'
  | 'text.highlight'
  | 'text.link'
  | 'text.characterSpacing'
  | 'text.changeCase'
  | 'text.dialog'
  | 'script.super'
  | 'script.sub'
  | 'paragraph.align'
  | 'list.bullet'
  | 'list.numbered'
  | 'list.bulletStyle'
  | 'list.numberStyle'
  | 'list.outdent'
  | 'list.indent'
  | 'list.lineSpacing'
  | 'paragraph.valign'
  | 'paragraph.dialog'
  | 'paragraph.textDirection'
  | 'paragraph.autofit'
  | 'paragraph.columns'
  | 'insert.textBox'
  | 'insert.line'
  | 'insert.shape'
  | 'insert.image'
  | 'shape.fill'
  | 'shape.lineColor'
  | 'shape.lineWidth'
  | 'shape.lineStart'
  | 'shape.lineEnd'
  | 'shape.lineDash'
  | 'shape.styles'
  | 'shape.effects'
  | 'shape.properties'
  | 'format.painter'
  | 'format.clear'
  | 'arrange.order'
  | 'arrange.align'
  | 'arrange.group'
  | 'arrange.ungroup'
  | 'arrange.rotate'
  | 'arrange.selectionPane'
  | 'more.duplicate'
  | 'more.delete'
  | 'find.find'
  | 'find.replace'
  | 'view.smartGuides'
  | 'view.presentationView'
  | 'file.pageSetup'
  | 'edit.cut'
  | 'edit.copy'
  | 'edit.paste'
  | 'edit.selectAll'
  | 'slide.add'
  | 'slide.layout'
  | 'slide.reset'
  | 'slide.hide'
  | 'slide.background'
  | 'view.playFromStart'
  | 'view.playFromCurrent'
  | 'section.add'
  | 'section.rename'
  | 'section.remove'
  | 'section.removeAll'
  | 'section.collapseAll'
  | 'section.expandAll'
  | TableSlotId
  | ImageSlotId;

/** Every slot id, derived from the registry. @public */
export function chromeSlotIds(): ChromeSlotId[] {
  return CHROME_GROUPS.flatMap((g) => g.controls.map((c) => `${g.id}.${c.id}` as ChromeSlotId));
}

/** The registry entry behind a slot, or null for an id the registry does not hold. @public */
export function chromeControlForSlot(slot: ChromeSlotId): ChromeControl | null {
  const dot = slot.indexOf('.');
  const group = CHROME_GROUPS.find((g) => g.id === slot.slice(0, dot));
  return group?.controls.find((c) => c.id === slot.slice(dot + 1)) ?? null;
}

/**
 * Whether a slot's control is a two-state toggle that carries `aria-pressed`. It reads the
 * registry's `toggle` flag, so an adapter never hand-lists the toggle slots. @public
 */
export function chromeSlotIsToggle(slot: ChromeSlotId): boolean {
  return chromeControlForSlot(slot)?.toggle === 'pressed';
}

/** Probe values: well-formed picks used only to ask `Editor.can` (never executed). */
const PROBES: Partial<Record<ChromeSlotId, unknown>> = {
  'font.family': 'Arial',
  'font.size': 18,
  'text.color': '#000000',
  'text.highlight': '#FFFF00',
  'paragraph.align': 'l',
  'list.lineSpacing': 1,
  ...TEXT_FORMAT_PROBES,
  'insert.shape': 'rect',
  'shape.fill': '#000000',
  'shape.lineColor': '#000000',
  'shape.lineWidth': '1',
  'shape.lineDash': 'solid',
  'paragraph.valign': 't',
  'text.link': 'https://example.com',
  'insert.line': 'line',
  'shape.lineStart': 'none',
  'shape.lineEnd': 'none',
  'arrange.order': 'forward',
  'arrange.align': 'left',
  ...SHAPE_FORMAT_PROBES,
  'view.presentationView': 'normal',
  ...TABLE_PROBES,
};

/** The shared can-before-exec state of one control. @public */
export interface ToolbarCommandState {
  readonly id: ChromeSlotId;
  readonly enabled: boolean;
  /** i18n key of the reason when disabled. */
  readonly disabledReason: string | null;
  readonly active: boolean;
  readonly value?: string | number;
}

/** Enabled and active state of a slot: the ONE source of truth for both. @public */
export function toolbarCommandState(editor: Editor | null, id: ChromeSlotId): ToolbarCommandState {
  if (!editor) return { id, enabled: false, disabledReason: 'disabled.noEditor', active: false };
  const image = imageSlotState(editor, id);
  if (image) return image;
  const snap = editor.snapshot();
  const f = snap.format;
  if (id === 'insert.image') {
    const probe = editor.can({ type: 'insertTextBox' });
    return { id, enabled: probe.ok, disabledReason: probe.ok ? null : probe.reason, active: false };
  }
  const kind = chromeControlForSlot(id)?.state.kind;
  if (kind === 'pane')
    // The Format pane and the selection pane are not built yet: the slot says so.
    return { id, enabled: false, disabledReason: 'disabled.paneNotBuilt', active: false };
  // A dialog with a probe (`paragraph.dialog`) is enabled when the engine takes its command.
  const unprobed = kind === 'dialog' && PROBES[id] === undefined;
  if (unprobed || kind === 'clipboard' || kind === 'slideshow')
    return adapterActionState(editor, id);
  if (id === 'find.find' || id === 'find.replace') {
    // Find reads any open deck; replace edits it, so it needs editing mode.
    const reason =
      snap.status !== 'ready'
        ? 'disabled.noEditor'
        : id === 'find.replace' && snap.mode === 'view'
          ? 'disabled.viewing'
          : null;
    return { id, enabled: !reason, disabledReason: reason, active: false };
  }
  const table = tableSlotState(editor, id);
  if (table) return table;
  const command =
    slideSlotProbe(id, snap) ??
    (PROBES[id] !== undefined ? commandForSlotValue(id, PROBES[id]) : commandForSlot(id));
  if (!command) return { id, enabled: false, disabledReason: 'disabled.notWired', active: false };
  const can = editor.can(command);
  const active =
    (id === 'text.bold' && !!f?.bold) ||
    (id === 'text.italic' && !!f?.italic) ||
    (id === 'text.underline' && !!f?.underline) ||
    (id === 'text.strike' && !!f?.strike) ||
    (id === 'script.super' && !!f?.superscript) ||
    (id === 'script.sub' && !!f?.subscript) ||
    (id === 'list.bullet' && !!f?.bullets) ||
    (id === 'list.numbered' && !!f?.numbering) ||
    (id === 'view.smartGuides' && snap.smartGuides) ||
    slideSlotActive(id, snap) ||
    tableSlotActive(id, snap) ||
    (id === 'format.painter' && (snap.formatPainter ?? 'off') !== 'off');
  const value =
    id === 'format.painter'
      ? snap.formatPainter
      : id === 'font.size'
        ? fontSizeSlotValue(f)
        : id === 'font.family'
          ? f?.fontFamily
          : id === 'paragraph.align'
            ? f?.align
            : id === 'text.color'
              ? f?.color
              : id === 'list.lineSpacing'
                ? lineSpacingSlotValue(f)
                : (textSlotValue(id, f) ??
                  (id === 'shape.lineWidth' && snap.geometry?.lineWidth !== undefined
                    ? String(snap.geometry.lineWidth)
                    : id === 'shape.lineDash'
                      ? snap.geometry?.lineDash
                      : id === 'text.link'
                        ? f?.link
                        : id === 'paragraph.valign'
                          ? f?.verticalAlign
                          : id === 'paragraph.autofit'
                            ? snap.geometry?.textFitting
                            : id === 'shape.lineStart'
                              ? snap.geometry?.lineHead
                              : id === 'shape.lineEnd'
                                ? snap.geometry?.lineTail
                                : (slideSlotValue(id, snap) ??
                                  shapeFormatSlotValue(id, snap) ??
                                  tableSlotValue(id, snap, tableBorderPen(editor)))));
  return {
    id,
    enabled: can.ok,
    disabledReason: can.ok ? null : can.reason,
    active,
    ...(value !== undefined ? { value } : {}),
  };
}

/**
 * Why one option's command cannot run, or null. One rule for the toolbar's options and the menu
 * rows: an option that would change nothing (Distribute on boxes already evenly spaced, No
 * shadow on a shape without one) stays enabled; only a refusal disables it.
 * @internal
 */
export function optionRefusal(editor: Editor, command: EditorCommand): string | null {
  const can = editor.can(command);
  if (can.ok) return null;
  return isNothingToDo(can.reason) && can.reason !== 'disabled.noSelection' ? null : can.reason;
}

/**
 * Run a slot's command (or its value command). Options come last, as in DOCX Editor: pass
 * `{ historyGroup }` to make the steps of one gesture (a slider drag) one undo step. A slot
 * without a value takes `undefined` as its value.
 *
 * @example
 * ```ts
 * const drag = editor.beginHistoryGroup();
 * runToolbarCommand(editor, 'font.size', 20, { historyGroup: drag });
 * runToolbarCommand(editor, 'font.size', 24, { historyGroup: drag });
 * drag.end(); // one undo step
 * ```
 * @public
 */
export function runToolbarCommand<K extends ChromeSlotId>(
  editor: Editor | null,
  id: K,
  value?: ToolbarSlotValue<K>,
  options?: EditorExecOptions
) {
  if (!editor) return { ok: false as const, reason: 'disabled.noEditor' };
  // A dialog slot run with no value, or a dialog row's value ("Line spacing options…"): say how
  // to open the dialog instead of running nothing or a refused command.
  const control = chromeControlForSlot(id);
  const row = control?.options?.find((o) => o.dialog && o.value === value);
  const dialog =
    row?.dialog ??
    (value === undefined && control?.state.kind === 'dialog' ? control.state.dialog : undefined);
  if (dialog) {
    warnDialogSlotRun(id, dialog, row?.value);
    return { ok: false as const, reason: 'disabled.dialogSlot' };
  }
  const command = value === undefined ? commandForSlot(id) : commandForSlotValue(id, value, editor);
  if (!command) {
    const state = chromeControlForSlot(id)?.state;
    if (state?.kind === 'pane') return { ok: false as const, reason: 'disabled.paneNotBuilt' };
    const reason = CHROME_SLOT_REASONS[id];
    // A raw hook that runs such a slot must not fail in silence.
    if (reason) {
      warnChromeSlotRun(id, reason);
      return { ok: false as const, reason };
    }
    if (value === undefined) return { ok: false as const, reason: 'disabled.notWired' };
    warnInvalidSlotValue(id, value);
    return { ok: false as const, reason: 'disabled.invalidValue' };
  }
  // A pen pick before any border target pick arms the pen and writes nothing.
  const armed = armTableBorderPen(editor, id, value, command);
  if (armed) return armed;
  const result = editor.exec(command, options);
  // A refused draw (no table) leaves the pen as it was.
  if (result.ok) rememberTableBorderPen(editor, id, value);
  return result;
}
