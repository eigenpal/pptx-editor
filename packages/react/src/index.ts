/**
 * `@pptx-editor.dev/react`, the React adapter for PPTX Editor. `PptxEditor.Root` owns the engine
 * instance, `.Viewport` scopes the styles, and `.Content` attaches the slide surface. Every piece
 * of chrome is a hook consumer, and the toolbar derives from the engine's `CHROME_GROUPS`.
 *
 * @packageDocumentation
 * @public
 */

import { createElement, forwardRef, useMemo, type FunctionComponent, type ReactNode } from 'react';
import type { ToolbarProps } from './Toolbar.tsx';
import { withChromeTranslate } from './headless-hooks.tsx';
import { MenuTableGrid } from './TableControls.tsx';
import {
  warnOpenFailureActions,
  type ColorMode,
  type EditorFontError,
  type EditorFontsOption,
  type EditorMode,
  type EditorModule,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import type { LocaleCatalog } from '@pptx-editor.dev/i18n';
import {
  PptxEditorContent,
  PptxEditorRoot,
  PptxEditorViewport,
  type PptxEditorRef,
} from './PptxEditorRoot.tsx';
import { Toolbar as ToolbarRoot, ToolbarControl, ToolbarGroup, ToolbarSlot } from './Toolbar.tsx';
import { ToolbarAction, ToolbarSeparator } from './toolbar-action.tsx';
import { TOOLBAR_PARTS, type ToolbarNamedParts } from './toolbar-named-parts.tsx';
import {
  NotesPane,
  SelectionPane,
  SidePanel,
  SlideRail as SlideRailRoot,
  ZoomPill,
  type NotesPaneProps,
  type SlideRailProps,
} from './panes.tsx';
import {
  SlideRailAddSlide,
  SlideRailChooseLayout,
  SlideRailCollapse,
  SlideRailExpand,
  SlideRailDelete,
  SlideRailDuplicate,
  SlideRailMoveEarlier,
  SlideRailMoveLater,
} from './rail-parts.tsx';
import { FontNotice } from './FontNotice.tsx';
import { EditNotice } from './EditNotice.tsx';
import {
  OpenFailure,
  type OpenFailureContext,
  type PptxEditorOpenFailureProps,
} from './OpenFailure.tsx';
import { Loading } from './Loading.tsx';
import { PageSetupDialog } from './PageSetupDialog.tsx';
import { BackgroundDialog } from './BackgroundDialog.tsx';
import { ImageAltTextDialog } from './ImageAltTextDialog.tsx';
import { ImagePropertiesDialog } from './ImagePropertiesDialog.tsx';
import { ParagraphDialog } from './ParagraphDialog.tsx';
import { FontDialog } from './FontDialog.tsx';
import { BulletsAndNumberingDialog } from './BulletsAndNumberingDialog.tsx';
import { ColumnsDialog } from './ColumnsDialog.tsx';
import { Slideshow } from './Slideshow.tsx';
import { FindBar } from './FindBar.tsx';
import { DefaultWorkspace, SlideSorter } from './SlideSorter.tsx';
import type { PptxEditorPopups } from './popups.tsx';
import {
  MENU_PARTS,
  MENU_ROW_PARTS,
  Menu as MenuRoot,
  MenuEntry,
  MenuGroup,
  MenuSlotItem,
  type MenuPartComponent,
  type MenuProps,
  type MenuRowPartComponent,
} from './menu-parts.tsx';
import {
  CONTEXT_MENU_ROW_PARTS,
  PptxEditorContextMenu as ContextMenuRoot,
  type PptxEditorContextMenuProps,
} from './context-menu-parts.tsx';
import type { ContextMenuPartName, MenuRowPartName } from '@pptx-editor.dev/core/editor';
import { MenuRow, MenuSeparator, MenuSubmenu } from './ContextMenu.tsx';
import { Ribbon as RibbonRoot, RibbonTab, type RibbonProps } from './Ribbon.tsx';
import {
  StatusBar as StatusBarRoot,
  StatusBarNotes,
  StatusBarSlideNumber,
  StatusBarViews,
  StatusBarZoom,
  type StatusBarProps,
} from './StatusBar.tsx';

export {
  PptxEditorContext,
  usePptxEditor,
  useEditorState,
  useEditorCommand,
  useEditorValueCommand,
  useTranslation,
  useSlideshow,
  translatorFor,
  SlideshowContext,
  type UseEditorCommandReturn,
  type UseEditorValueCommandReturn,
  type UseSlideshowReturn,
} from './context.ts';
export {
  PptxEditorRoot,
  PptxEditorViewport,
  PptxEditorContent,
  type PptxEditorContentProps,
  type PptxEditorRef,
  type PptxEditorRootProps,
  type PptxEditorViewportProps,
} from './PptxEditorRoot.tsx';
export {
  ToolbarControl,
  ToolbarGroup,
  ToolbarSlot,
  type ToolbarControlProps,
  type ToolbarProps,
} from './Toolbar.tsx';
export type {
  ToolbarColorSplitComponent,
  ToolbarColorSplitProps,
  ToolbarGroupProps,
  ToolbarPartComponent,
  ToolbarPartProps,
  ToolbarSlotPartComponent,
  ToolbarSlotPartProps,
  ToolbarSlotBaseProps,
  ToolbarSlotProps,
} from './toolbar-parts.tsx';
export {
  ToolbarAction,
  ToolbarSeparator,
  type ToolbarActionProps,
  type ToolbarSeparatorProps,
} from './toolbar-action.tsx';
export type { ToolbarNamedParts } from './toolbar-named-parts.tsx';
export type {
  FontFamilyItemProps,
  FontFamilyNamespace,
  FontFamilyPartProps,
  FontFamilyProps,
} from './FontFamilyControl.tsx';
export {
  definePopup,
  type DefinePopupOptions,
  useEditorDialog,
  type UseEditorDialogReturn,
  type PptxEditorDialogKey,
  type PptxEditorPopup,
  type PptxEditorPopups,
  type FindPopupProps,
  type HyperlinkPopupProps,
} from './popups.tsx';
export { HyperlinkPopup } from './HyperlinkPopup.tsx';
import { HyperLink } from './HyperLink.tsx';
export {
  HyperLink,
  type HyperLinkActionProps,
  type HyperLinkNamespace,
  type HyperLinkPartProps,
  type HyperLinkProps,
} from './HyperLink.tsx';
export {
  RailContextMenu,
  type RailContextMenuProps,
  type RailMenuState,
} from './RailContextMenu.tsx';
export {
  SectionContextMenu,
  type SectionContextMenuProps,
  type SectionMenuState,
} from './RailSections.tsx';
export {
  MenuEntry,
  MenuGroup,
  MenuSlotItem,
  type MenuEntryProps,
  type MenuGroupProps,
  type MenuSlotItemProps,
  type MenuPartComponent,
  type MenuProps,
  type MenuRowPartComponent,
  type MenuRowPartProps,
} from './menu-parts.tsx';
export type {
  ContextMenuPartName,
  MenuNamedRow,
  MenuRowPartName,
} from '@pptx-editor.dev/core/editor';
export {
  useContextMenuRequest,
  useContextMenuTarget,
  type ContextMenuRequest,
  type PptxEditorContextMenuProps,
} from './context-menu-parts.tsx';

/**
 * `PptxEditor.Toolbar` with its statics, as DOCX Editor's `DocxEditorToolbarNamespace`. @public
 */
export interface PptxEditorToolbarNamespace extends ToolbarNamedParts {
  /** The toolbar, with every registered group. See `ToolbarProps`. */
  (props: ToolbarProps): ReactNode;
  /** Changes, replaces, or hides one registered control. */
  readonly Slot: typeof ToolbarSlot;
  /** Changes a built-in group or adds a host group. */
  readonly Group: typeof ToolbarGroup;
  /** A host-owned action the chrome registry does not describe. */
  readonly Action: typeof ToolbarAction;
  /** A vertical rule between controls. */
  readonly Separator: typeof ToolbarSeparator;
}

/**
 * The toolbar, with its parts: `Toolbar.Slot` changes or replaces one registered control,
 * `Toolbar.Group` changes a built-in group or adds a host group, `Toolbar.Action` is a host's
 * own button, and the named parts (`Toolbar.Bold`, `Toolbar.FontFamily`, …) stand for one slot
 * each. @public
 */
export const Toolbar: PptxEditorToolbarNamespace = Object.assign(
  withChromeTranslate<ToolbarProps>(ToolbarRoot, 'Toolbar'),
  {
    ...TOOLBAR_PARTS,
    Slot: ToolbarSlot,
    Group: ToolbarGroup,
    Action: ToolbarAction,
    Separator: ToolbarSeparator,
  }
);
/**
 * The named rows of the menu bar (`Menu.Cut`, `Menu.PageSetup`, …), one per row of core's
 * `MENU_NAMED_PARTS`. Put one in the menu that holds its row. @public
 */
export type MenuNamedRows = { readonly [K in MenuRowPartName]: MenuRowPartComponent };

/**
 * The named rows of the canvas context menu (`ContextMenu.Cut`, `ContextMenu.MergeCells`, …),
 * one per row of core's `CONTEXT_MENU_NAMED_PARTS`. @public
 */
export type ContextMenuNamedRows = {
  readonly [K in ContextMenuPartName]: MenuRowPartComponent;
};

/** The menu bar with its parts attached as statics. @public */
export interface MenuNamespace extends MenuNamedRows {
  /** The menu bar. See `MenuProps`. */
  (props: MenuProps): ReactNode;
  /** A menu of the bar: a built-in id changes it, another id adds a host menu. */
  readonly Entry: typeof MenuEntry;
  /** A row for a registered slot. */
  readonly Item: typeof MenuSlotItem;
  /** A host row: label, icon, shortcut, and an action. */
  readonly Row: typeof MenuRow;
  /** A labeled set of rows. */
  readonly Group: typeof MenuGroup;
  /** A line between rows. */
  readonly Separator: typeof MenuSeparator;
  /** A row that opens a nested panel. */
  readonly Submenu: typeof MenuSubmenu;
  /** The Insert > Table grid, for a host's own Table submenu (DOCX Editor's `Menu.TableGrid`). */
  readonly TableGrid: typeof MenuTableGrid;
  /** The File menu. */
  readonly File: MenuPartComponent;
  /** The Edit menu. */
  readonly Edit: MenuPartComponent;
  /** The View menu. */
  readonly View: MenuPartComponent;
  /** The Insert menu. */
  readonly Insert: MenuPartComponent;
  /** The Format menu. */
  readonly Format: MenuPartComponent;
  /** The Arrange menu. */
  readonly Arrange: MenuPartComponent;
  /** The Slide Show menu. */
  readonly SlideShow: MenuPartComponent;
}

/**
 * The menu bar (`PptxEditor.Menu`): File, Edit, View, Insert, Format, Arrange, and Slide Show,
 * with DOCX Editor's parts (`Entry`, `Item`, `Row`, `Group`, `Separator`, `Submenu`) and one
 * named part per menu (`Menu.File`, …). @public
 */
export const PptxEditorMenu: MenuNamespace = Object.assign(
  withChromeTranslate<MenuProps>(MenuRoot, 'Menu'),
  {
    ...MENU_ROW_PARTS,
    ...MENU_PARTS,
    Entry: MenuEntry,
    Item: MenuSlotItem,
    Row: MenuRow,
    Group: MenuGroup,
    Separator: MenuSeparator,
    Submenu: MenuSubmenu,
    TableGrid: MenuTableGrid,
  }
);

/** The canvas context menu with its parts attached as statics. @public */
export interface ContextMenuNamespace extends ContextMenuNamedRows {
  /** The canvas context menu, for the target that opened it. See `PptxEditorContextMenuProps`. */
  (props: PptxEditorContextMenuProps): ReactNode;
  /** A row for a registered slot; replaces or hides the built-in row of the same slot. */
  readonly Slot: typeof MenuSlotItem;
  /** A host row: label, icon, shortcut, and an action. */
  readonly Item: typeof MenuRow;
  /** A labeled set of rows. */
  readonly Group: typeof MenuGroup;
  /** A line between rows. */
  readonly Separator: typeof MenuSeparator;
  /** A row that opens a nested panel. */
  readonly Submenu: typeof MenuSubmenu;
}

/**
 * The slide canvas's context menu (`PptxEditor.ContextMenu`), with DOCX Editor's parts (`Slot`,
 * `Item`, `Group`, `Separator`, `Submenu`). Place it in the root or the viewport, or render it
 * in `popups.contextMenu`. @public
 */
export const PptxEditorContextMenu: ContextMenuNamespace = Object.assign(
  withChromeTranslate<PptxEditorContextMenuProps>(ContextMenuRoot, 'ContextMenu'),
  {
    ...CONTEXT_MENU_ROW_PARTS,
    Slot: MenuSlotItem,
    Item: MenuRow,
    Group: MenuGroup,
    Separator: MenuSeparator,
    Submenu: MenuSubmenu,
  }
);

export { RibbonTab, type RibbonProps, type RibbonTabProps } from './Ribbon.tsx';

/** The ribbon with its parts attached as statics. @public */
export interface RibbonNamespace {
  /** The ribbon. See `RibbonProps`. */
  (props: RibbonProps): ReactNode;
  /** A tab: a built-in id changes it, another id adds a host tab. */
  readonly Tab: typeof RibbonTab;
}

/**
 * The ribbon (`PptxEditor.Ribbon`): the reference PDF's tabs over the toolbar's
 * groups, with the Quick Access Toolbar and contextual tabs. Each tab panel takes toolbar
 * children (`Toolbar.Slot`, named parts, `Toolbar.Group`, `Toolbar.Action`) through
 * `Ribbon.Tab`. @public
 */
export const PptxEditorRibbon: RibbonNamespace = Object.assign(RibbonRoot, { Tab: RibbonTab });

export {
  StatusBarNotes,
  StatusBarSlideNumber,
  StatusBarViews,
  StatusBarZoom,
  type StatusBarNotesProps,
  type StatusBarPartProps,
  type StatusBarProps,
} from './StatusBar.tsx';

/** The status bar with its parts attached as statics. @public */
export interface StatusBarNamespace {
  /** The status bar. See `StatusBarProps`. */
  (props: StatusBarProps): ReactNode;
  /** "Slide x of y". */
  readonly SlideNumber: typeof StatusBarSlideNumber;
  /** The Notes button. */
  readonly Notes: typeof StatusBarNotes;
  /** The view switch. */
  readonly Views: typeof StatusBarViews;
  /** The zoom pill. */
  readonly Zoom: typeof StatusBarZoom;
  /** A host-owned button with the bar's styling: the same component as `Toolbar.Action`. */
  readonly Action: typeof ToolbarAction;
}

/**
 * The status bar (`PptxEditor.StatusBar`): slide x of y, the Notes button, the view switch,
 * and the zoom pill. @public
 */
export const PptxEditorStatusBar: StatusBarNamespace = Object.assign(
  withChromeTranslate<StatusBarProps>(StatusBarRoot, 'StatusBar'),
  {
    SlideNumber: StatusBarSlideNumber,
    Notes: StatusBarNotes,
    Views: StatusBarViews,
    Zoom: StatusBarZoom,
    Action: ToolbarAction,
  }
);

export {
  SlideRailAddSlide,
  SlideRailChooseLayout,
  SlideRailCollapse,
  SlideRailExpand,
  SlideRailDelete,
  SlideRailDuplicate,
  SlideRailMoveEarlier,
  SlideRailMoveLater,
  type SlideRailPartProps,
} from './rail-parts.tsx';

/** The slide rail with its action parts attached as statics, as `StatusBarNamespace`. @public */
export interface SlideRailNamespace {
  /** The slide rail. See `SlideRailProps`. */
  (props: SlideRailProps): ReactNode;
  /** Add slide. */
  readonly AddSlide: typeof SlideRailAddSlide;
  /** The layout gallery button. */
  readonly ChooseLayout: typeof SlideRailChooseLayout;
  /** Duplicate the selected slides. */
  readonly Duplicate: typeof SlideRailDuplicate;
  /** Move the selected slides earlier. */
  readonly MoveEarlier: typeof SlideRailMoveEarlier;
  /** Move the selected slides later. */
  readonly MoveLater: typeof SlideRailMoveLater;
  /** Delete the selected slides. */
  readonly Delete: typeof SlideRailDelete;
  /** Hide thumbnails. */
  readonly Collapse: typeof SlideRailCollapse;
  /** Show thumbnails: the button of the folded strip. */
  readonly Expand: typeof SlideRailExpand;
  /** A host-owned button with the rail's styling: the same component as `Toolbar.Action`. */
  readonly Action: typeof ToolbarAction;
}

/**
 * The slide rail (`PptxEditor.SlideRail`) with its action parts: a named part child replaces
 * its button, `hidden` removes it, other children go after Delete, and `preset={false}` renders
 * only the children. @public
 */
export const SlideRail: SlideRailNamespace = Object.assign(
  withChromeTranslate<SlideRailProps>(SlideRailRoot, 'SlideRail'),
  {
    AddSlide: SlideRailAddSlide,
    ChooseLayout: SlideRailChooseLayout,
    Duplicate: SlideRailDuplicate,
    MoveEarlier: SlideRailMoveEarlier,
    MoveLater: SlideRailMoveLater,
    Delete: SlideRailDelete,
    Collapse: SlideRailCollapse,
    Expand: SlideRailExpand,
    Action: ToolbarAction,
  }
);

export {
  NotesPane,
  ZoomPill,
  SidePanel,
  SelectionPane,
  listenPaneRequests,
  requestPane,
  type NotesPaneProps,
  type PaneProps,
  type SelectionPaneProps,
  type SidePanelProps,
  type SlideRailProps,
  type ZoomPillProps,
} from './panes.tsx';
export { useNotesPane, type UseNotesPaneReturn } from './pane-state.ts';
export { useEditorColorMode, type UseEditorColorModeReturn } from './color-mode.ts';
export {
  ToolbarContext,
  useChromeTranslate,
  useDocumentSearch,
  useEditorCaret,
  useHistoryGroup,
  useToolbarContext,
  useToolbarLabel,
  useToolbarLabelFor,
  type ToolbarContextValue,
  type UseDocumentSearchResult,
  type UseHistoryGroupReturn,
} from './headless-hooks.tsx';
export type {
  ChromeTranslate,
  DocumentSearchFindOptions,
  DocumentSearchState,
  EditorCaret,
  HistoryGroupBindingOptions,
} from '@pptx-editor.dev/core/editor';
export {
  useEditorEvent,
  useEditorSnapshot,
  useScopeClassName,
  useZoom,
  type UseZoomResult,
} from './customization-hooks.ts';
export type { PptxEditorEvents } from '@pptx-editor.dev/core/editor';
export type { ColorMode, ResolvedColorMode } from '@pptx-editor.dev/core/editor';
export { ShapeStyleGallery, type ShapeStyleGalleryProps } from './ShapeFormatControls.tsx';
export { FontNotice, type FontNoticeProps } from './FontNotice.tsx';
export { MenuTableGrid, type MenuTableGridProps } from './TableControls.tsx';
export { EditNotice, rejectionMessage, type EditNoticeProps } from './EditNotice.tsx';
export {
  OpenFailure,
  useOpenFailure,
  type OpenFailureContext,
  type PptxEditorOpenFailureProps,
  type UseOpenFailureOptions,
  type UseOpenFailureReturn,
} from './OpenFailure.tsx';
export {
  useHyperlinkPopup,
  useHyperlinkPopupInstance,
  type UseHyperlinkPopupResult,
} from './useHyperlinkPopup.ts';
export type {
  HyperlinkPopupActions,
  HyperlinkPopupLink,
  HyperlinkPopupMode,
  HyperlinkPopupState,
} from '@pptx-editor.dev/core/editor';
export {
  useEditorHostDocument,
  usePptxSource,
  type UseEditorHostDocumentReturn,
  type UsePptxSourceResult,
} from './usePptxSource.ts';
export type {
  PptxFontOrigin,
  PptxFontsInput,
  PptxFontsSource,
  PptxSource,
  UsePptxSourceOptions,
} from '@pptx-editor.dev/core/editor';
export {
  Loading,
  LoadingSpinner,
  type LoadingComponent,
  type LoadingProps,
  type LoadingSpinnerProps,
} from './Loading.tsx';
export {
  FillControl,
  GradientDialog,
  LineGlyph,
  LineMenuControl,
  ShapeGalleryControl,
  ShapeGrid,
  ShapeIcon,
  type FillControlProps,
  type GradientDialogProps,
  type LineGlyphProps,
  type LineMenuControlProps,
  type ShapeGalleryControlProps,
  type ShapeGridProps,
  type ShapeIconProps,
} from './ShapeControls.tsx';
export { PatternDialog, type PatternDialogProps } from './PatternDialog.tsx';
export { SlideSorter, type SlideSorterProps } from './SlideSorter.tsx';
export { requestSectionRename } from './RailSections.tsx';
export {
  LayoutControl,
  LayoutGallery,
  NewSlideControl,
  useSlideLayouts,
  type LayoutControlProps,
  type LayoutGalleryProps,
  type NewSlideControlProps,
} from './SlideControls.tsx';
export { gradientCss } from '@pptx-editor.dev/core/editor';
export * from './dialog-exports.ts';
export { useParagraphFormat, type UseParagraphFormatReturn } from './useParagraphFormat.ts';
export { useFontFormat, type UseFontFormatReturn } from './useFontFormat.ts';
export { useBulletFormat, type UseBulletFormatReturn } from './useBulletFormat.ts';
export { useFonts, type FontsInput, type UseFontsReturn } from './useFonts.ts';
export { useFontFamily, type UseFontFamilyReturn } from './useFontFamily.ts';
export { EditorFontError } from '@pptx-editor.dev/core/editor';
export type {
  EditorFontErrorCode,
  FontConfiguration,
  FontFaceRequest,
  FontSource,
  FontSourceSubstitution,
} from '@pptx-editor.dev/core/contracts';
// The font-composition surface, re-exported so a fonts package and an adapter need no core import.
export {
  MAX_RESOLVER_FAMILIES,
  DEFAULT_FONT,
  composeFontConfiguration,
  composeFontOrigins,
  createFontSource,
  deferFonts,
  defineFontResolver,
  isFontResolver,
  loadFonts,
  type FontConfigurationBase,
  type FontConfigurationFragment,
  type FontLoadFailure,
  type FontLoadFailureReason,
  type FontOrigin,
  type FontResolutionRequest,
  type FontResolver,
  type FontResolverMark,
  type MarkedFontResolver,
  type FontUrlSource,
  type LoadFontsRequest,
  type LoadFontsResult,
} from '@pptx-editor.dev/core/editor';
// The core types of the adapter's props, hooks, and parts, re-exported as DOCX Editor's adapters
// do, so a host types its handlers with an import from this package only.
export type {
  BulletFormat,
  BulletFormatMixed,
  BulletFormatRead,
  BulletsDialogFields,
  CanvasMenuKind,
  ChromeControl,
  ChromeGroupId,
  ChromeMenuId,
  ChromeSlotId,
  ColumnsDialogFields,
  CommandResult,
  DocumentOpenFailure,
  EditorCommand,
  EditorExecOptions,
  EditorFontsOption,
  EditorMode,
  EditorModule,
  EditorRejection,
  EditorRemoval,
  EditorSnapshot,
  FontDialogFields,
  FontFormatCommand,
  FontFormatRead,
  FontPickerGroup,
  FontPickerGroupId,
  FontPickerRow,
  GradientFill,
  PatternFill,
  PatternPreset,
  ParagraphDialogFields,
  ParagraphFormatCommand,
  ParagraphFormatRead,
  PptxEditorInstance,
  RibbonTabId,
  ShapeStyleChoice,
  ShapeStyleValue,
  SidePanelTab,
  SlideBackground,
  SlideLayoutSummary,
  SlideRailOrientation,
  SlideSection,
  SlideshowController,
  ToolbarShownValue,
  ToolbarValueMap,
} from '@pptx-editor.dev/core/editor';
export { Slideshow } from './Slideshow.tsx';
export { FindBar, requestFindBar, type FindBarProps } from './FindBar.tsx';
export { platformShortcuts } from './context.ts';
export {
  Icon,
  ToolButton,
  Popover,
  MenuItem,
  keepFocus,
  type IconProps,
  type MenuItemProps,
  type PopoverProps,
  type ToolButtonProps,
} from './chrome.tsx';
export {
  ContextMenu,
  MenuRow,
  MenuSeparator,
  MenuSubmenu,
  type ContextMenuPoint,
  type ContextMenuProps,
  type MenuRowProps,
  type MenuSubmenuProps,
} from './ContextMenu.tsx';

/** Props of the `<PptxEditor>` sugar component. @public */
export interface PptxEditorProps {
  /** `.pptx` bytes. Identity change loads the new document. */
  readonly document?: Uint8Array;
  /** `'edit'` (default) or `'view'`. Changes apply without a reload. */
  readonly mode?: EditorMode;
  /** UI locale catalog, as on `PptxEditor.Root`. English when omitted. */
  readonly locale?: LocaleCatalog;
  /** Fired once per instance after it is published, as on `PptxEditor.Root`. */
  readonly onReady?: (editor: PptxEditorInstance) => void;
  /** Fired after every change of editor state, as on `PptxEditor.Root`. */
  readonly onChange?: (editor: PptxEditorInstance) => void;
  /** Appended after the `pptx-default` class of the viewport. */
  readonly className?: string;
  /** Customize the dialogs the editor mounts, as on `PptxEditor.Root`. */
  readonly popups?: PptxEditorPopups;
  /** The chrome color mode, as on `PptxEditor.Root`. */
  readonly colorMode?: ColorMode;
  /** Font bytes for the host's faces, as on `PptxEditor.Root`. */
  readonly fonts?: EditorFontsOption;
  /** Called for each font failure, as on `PptxEditor.Root`. */
  readonly onFontError?: (error: EditorFontError) => void;
  /** Families the font picker offers, as on `PptxEditor.Root`. A change applies in place. */
  readonly fontFamilies?: readonly string[];
  /** Capability modules (`@pptx-editor.dev/pro`), as on `PptxEditor.Root`. Read at mount. */
  readonly modules?: readonly EditorModule[];
  /** AutoCorrect while typing, as on `PptxEditor.Root`. Read at mount. */
  readonly autoCorrect?: boolean;
  /**
   * The notice for a file that does not open (`PptxEditor.OpenFailure`). An object is
   * `PptxEditorOpenFailureProps`, passed straight through. `false` removes the notice, for a host that
   * reports the failure itself through `useOpenFailure`. Default `true`.
   */
  readonly openFailure?: boolean | PptxEditorOpenFailureProps;
  /**
   * Your own buttons in the open-failure notice, before **Close**. It replaces
   * `openFailure.actions`; pass only one. With `openFailure={false}` it renders nothing. Both
   * cases print a development warning.
   */
  readonly renderOpenFailureActions?: (context: OpenFailureContext) => ReactNode;
  /**
   * The toolbar (`PptxEditor.Toolbar`). An object is `ToolbarProps`, passed straight through,
   * so `toolbar={{ children: <PptxEditor.Toolbar.Bold hidden /> }}` changes one control.
   * `false` removes it. Default `true`.
   */
  readonly toolbar?: boolean | ToolbarProps;
  /**
   * The menu bar (`PptxEditor.Menu`), above the toolbar. An object is `MenuProps`, passed
   * straight through, as DOCX Editor's `menu`. Default `false`: the packaged editor shows the
   * toolbar only.
   */
  readonly menu?: boolean | MenuProps;
  /**
   * The slide canvas's context menu. An object is `PptxEditorContextMenuProps`, passed straight
   * through, so `contextMenu={{ children: <PptxEditor.ContextMenu.Cut hidden /> }}` changes one
   * row. `false` shows no menu. It sets `popups.contextMenu`, and a `popups.contextMenu` entry
   * wins over it. Default `true`.
   */
  readonly contextMenu?: boolean | PptxEditorContextMenuProps;
  /**
   * The packaged link popup. `false` removes it, as `popups.hyperlink: false` does: a host
   * panel that reads `useHyperlinkPopup()` then opens for the link controls. A
   * `popups.hyperlink` entry wins over it. Default `true`.
   */
  readonly hyperlinkPopup?: boolean;
  /**
   * The slide rail (`PptxEditor.SlideRail`), as DOCX Editor's `navigation`. An object is
   * `SlideRailProps`, passed straight through, so `slideRail={{ open, onOpenChange }}` drives
   * it from your own chrome. `false` removes it. Default `true`.
   */
  readonly slideRail?: boolean | SlideRailProps;
  /**
   * The speaker notes pane (`PptxEditor.NotesPane`). An object is `NotesPaneProps`, passed
   * straight through. `false` removes it. Default `true`.
   */
  readonly notesPane?: boolean | NotesPaneProps;
  /** Extra content inside the viewport, after the slide show. */
  readonly children?: ReactNode;
}

/** The props an object value gives, or none for `true`. */
const partProps = <P extends object>(value: boolean | P | undefined): P =>
  typeof value === 'object' ? value : ({} as P);

/**
 * The `popups` the sugar passes to the root: the host's entries, with `contextMenu` and
 * `hyperlinkPopup` mapped onto the same entries. A host entry wins.
 */
function sugarPopups(
  popups: PptxEditorPopups | undefined,
  hyperlinkPopup: boolean | undefined,
  contextMenu: boolean | PptxEditorContextMenuProps | undefined
): PptxEditorPopups | undefined {
  const linkOff = hyperlinkPopup === false && popups?.hyperlink === undefined;
  const menuEntry =
    popups?.contextMenu !== undefined || contextMenu === undefined || contextMenu === true
      ? undefined
      : contextMenu === false
        ? (false as const)
        : (props: PptxEditorContextMenuProps) =>
            createElement(PptxEditorContextMenu, { ...props, ...contextMenu });
  if (!linkOff && menuEntry === undefined) return popups;
  return {
    ...popups,
    ...(linkOff ? { hyperlink: false as const } : {}),
    ...(menuEntry !== undefined ? { contextMenu: menuEntry } : {}),
  };
}

const OPEN_FAILURE_ACTION_NAMES = {
  actions: 'renderOpenFailureActions',
  off: 'openFailure={false}',
} as const;

/**
 * The complete default editor: toolbar, slide rail, canvas, speaker notes, zoom pill, and the
 * slide show (F5).
 * Sugar over the same primitives a custom layout composes. Takes a `ref` to a `PptxEditorRef`.
 *
 * @public
 */
const PptxEditorDefault = forwardRef<PptxEditorRef, PptxEditorProps>(function PptxEditor(
  {
    document,
    mode = 'edit',
    locale,
    onReady,
    onChange,
    className,
    popups,
    colorMode,
    fonts,
    onFontError,
    fontFamilies,
    modules,
    autoCorrect,
    openFailure = true,
    renderOpenFailureActions,
    toolbar = true,
    menu = false,
    contextMenu,
    hyperlinkPopup,
    slideRail = true,
    notesPane = true,
    children,
  },
  ref
) {
  warnOpenFailureActions(
    openFailure,
    renderOpenFailureActions !== undefined,
    OPEN_FAILURE_ACTION_NAMES
  );
  // Memoized, so stable props give the root the same `popups` at each render.
  const rootPopups = useMemo(
    () => sugarPopups(popups, hyperlinkPopup, contextMenu),
    [popups, hyperlinkPopup, contextMenu]
  );
  const dock =
    menu === false && toolbar === false
      ? null
      : createElement(
          'div',
          { className: 'pptx-default__dock' },
          menu === false ? null : createElement(PptxEditorMenu, partProps(menu)),
          toolbar === false ? null : createElement(Toolbar, partProps(toolbar))
        );
  return createElement(
    PptxEditorRoot,
    {
      ref,
      ...(document ? { document } : {}),
      mode,
      ...(locale ? { locale } : {}),
      ...(onReady ? { onReady } : {}),
      ...(onChange ? { onChange } : {}),
      ...(rootPopups ? { popups: rootPopups } : {}),
      ...(colorMode ? { colorMode } : {}),
      ...(fonts ? { fonts } : {}),
      ...(onFontError ? { onFontError } : {}),
      ...(fontFamilies ? { fontFamilies } : {}),
      ...(modules ? { modules } : {}),
      ...(autoCorrect === false ? { autoCorrect } : {}),
    },
    createElement(
      PptxEditorViewport,
      { className: `pptx-default${className ? ` ${className}` : ''}` },
      dock,
      createElement(DefaultWorkspace, {
        overlay:
          openFailure === false
            ? null
            : createElement(OpenFailure, {
                ...(typeof openFailure === 'object' ? openFailure : {}),
                ...(renderOpenFailureActions ? { actions: renderOpenFailureActions } : {}),
              }),
        normal: createElement(
          'div',
          { className: 'pptx-default__workspace' },
          slideRail === false ? null : createElement(SlideRail, partProps(slideRail)),
          createElement(
            'div',
            { className: 'pptx-default__main' },
            createElement(PptxEditorContent),
            createElement(FindBar),
            createElement(EditNotice),
            createElement(Loading, { overlay: true }),
            createElement(FontNotice),
            notesPane === false
              ? null
              : createElement(NotesPane as FunctionComponent<NotesPaneProps>, partProps(notesPane))
          )
        ),
      }),
      createElement(Slideshow),
      children
    )
  );
});

/** The compound namespace plus the sugar component. @public */
export const PptxEditor = Object.assign(PptxEditorDefault, {
  Root: PptxEditorRoot,
  Viewport: PptxEditorViewport,
  Content: PptxEditorContent,
  Toolbar,
  ToolbarControl,
  SlideRail,
  SlideSorter,
  NotesPane,
  ZoomPill,
  SidePanel,
  SelectionPane,
  FontNotice,
  EditNotice,
  OpenFailure,
  Loading,
  PageSetupDialog,
  BackgroundDialog,
  ImageAltTextDialog,
  ImagePropertiesDialog,
  ParagraphDialog,
  FontDialog,
  BulletsAndNumberingDialog,
  ColumnsDialog,
  Slideshow,
  FindBar,
  HyperLink,
  Menu: PptxEditorMenu,
  ContextMenu: PptxEditorContextMenu,
  Ribbon: PptxEditorRibbon,
  StatusBar: PptxEditorStatusBar,
});
