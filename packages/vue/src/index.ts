// @pptx-editor.dev/vue: the Vue 3 adapter. Provider-first, like the React adapter and DOCX
// Editor's Vue adapter: `PptxEditorRoot` owns the instance, `PptxEditorViewport` is the
// styling scope, `PptxEditorContent` attaches the canvas, and all chrome reads the editor
// through composables. `scripts/parity/parity.contract.json` lists which React members Vue
// pairs and which it defers.

import {
  defineComponent,
  getCurrentInstance,
  h,
  mergeProps,
  shallowRef,
  type ExtractPublicPropTypes,
  type PropType,
  type SlotsType,
  type VNodeChild,
} from 'vue';
import {
  createEditorRefApi,
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
} from './components.ts';
import { PptxEditorToolbar } from './toolbar-parts.ts';
import { PptxEditorLoading } from './loading.ts';
import type { PptxEditorPopups } from './popups.ts';
import { DefaultWorkspace } from './SlideSorter.tsx';
import { PptxEditorSlideRail } from './panes.tsx';
import { PptxEditorNotesPane } from './NotesPane.tsx';
import { PptxEditorSelectionPane } from './SelectionPane.tsx';
import { PptxEditorFindBar } from './FindBar.tsx';
import { PptxEditorEditNotice } from './EditNotice.tsx';
import {
  PptxEditorOpenFailure,
  type OpenFailureContext,
  type PptxEditorOpenFailureProps,
} from './OpenFailure.tsx';
import { PptxEditorFontNotice } from './FontNotice.tsx';
import { PptxEditorSlideshow } from './Slideshow.tsx';
import { PptxEditorSlideSorter } from './SlideSorter.tsx';
import { PptxEditorSidePanel, PptxEditorZoomPill } from './panes.tsx';
import { PptxEditorMenu, type MenuProps } from './menu-parts.tsx';
import { PptxEditorContextMenu, type PptxEditorContextMenuProps } from './context-menu-parts.tsx';
import type { ToolbarProps } from './toolbar.ts';
import type { SlideRailProps } from './panes.tsx';
import type { NotesPaneProps } from './NotesPane.tsx';
import { PptxEditorRibbon } from './ribbon.ts';
import { PptxEditorStatusBar } from './status-bar.ts';
import { ToolbarControl } from './toolbar.ts';
import { PageSetupDialog } from './PageSetupDialog.ts';
import { ImageAltTextDialog } from './ImageAltTextDialog.ts';
import { ImagePropertiesDialog } from './ImagePropertiesDialog.ts';
import { BackgroundDialog } from './BackgroundDialog.tsx';
import { ParagraphDialog } from './ParagraphDialog.tsx';
import { FontDialog } from './FontDialog.tsx';
import { BulletsAndNumberingDialog } from './BulletsAndNumberingDialog.tsx';
import { ColumnsDialog } from './ColumnsDialog.tsx';

export {
  pptxEditorKey,
  pptxTranslationKey,
  useBulletFormat,
  useEditorCommand,
  useEditorValueCommand,
  useEditorState,
  useFontFormat,
  useParagraphFormat,
  usePptxEditor,
  useSlideLayouts,
  useTranslation,
  type UseBulletFormatReturn,
  type UseEditorCommandReturn,
  type UseEditorValueCommandReturn,
  type UseEditorStateReturn,
  type UseFontFormatReturn,
  type UseParagraphFormatReturn,
  type UsePptxEditorReturn,
  type UseSlideLayoutsReturn,
  type UseTranslationReturn,
} from './composables.ts';
// The composable's value types, next to it, as DOCX Editor exports them with `useParagraphFormat`.
export type {
  BulletFormat,
  BulletFormatRead,
  FontFormatCommand,
  FontFormatRead,
  ParagraphFormatCommand,
  ParagraphFormatRead,
} from '@pptx-editor.dev/core/editor';
export { useEditorColorMode, type UseEditorColorModeReturn } from './color-mode.ts';
export {
  useEditorEvent,
  useEditorSnapshot,
  useScopeClassName,
  useZoom,
  type UseScopeClassNameReturn,
  type UseEditorEventReturn,
  type UseEditorSnapshotReturn,
  type UseZoomReturn,
} from './customization-hooks.ts';
export {
  useChromeTranslate,
  useDocumentSearch,
  useEditorCaret,
  useHistoryGroup,
  useToolbarContext,
  useToolbarLabel,
  useToolbarLabelFor,
  type UseChromeTranslateReturn,
  type UseDocumentSearchReturn,
  type UseEditorCaretReturn,
  type UseHistoryGroupReturn,
  type UseToolbarContextReturn,
  type UseToolbarLabelForReturn,
  type UseToolbarLabelReturn,
} from './headless-hooks.ts';
export type {
  ChromeTranslate,
  DocumentSearchFindOptions,
  DocumentSearchState,
  EditorCaret,
  HistoryGroupBindingOptions,
} from '@pptx-editor.dev/core/editor';
export type { PptxEditorEvents } from '@pptx-editor.dev/core/editor';
export type { ColorMode, ResolvedColorMode } from '@pptx-editor.dev/core/editor';
export {
  PptxEditorContent,
  PptxEditorRoot,
  PptxEditorViewport,
  type PptxEditorContentProps,
  type PptxEditorRef,
  type PptxEditorRootProps,
  type PptxEditorViewportProps,
} from './components.ts';
export {
  definePopup,
  type DefinePopupOptions,
  useEditorDialog,
  type PptxEditorDialogKey,
  type PptxEditorDialogSessionProps,
  type PptxEditorPopup,
  type PptxEditorPopups,
  type UseEditorDialogReturn,
} from './popups.ts';
// The overlays: the find bar, the notices, the link popup, and the slide show.
export {
  PptxEditorFindBar,
  requestFindBar,
  type FindPopupProps,
  type FindBarProps,
} from './FindBar.tsx';
export { PptxEditorEditNotice, rejectionMessage, type EditNoticeProps } from './EditNotice.tsx';
export {
  PptxEditorOpenFailure,
  useOpenFailure,
  type OpenFailureContext,
  type PptxEditorOpenFailureProps,
  type UseOpenFailureOptions,
  type UseOpenFailureReturn,
} from './OpenFailure.tsx';
export { PptxEditorFontNotice, type FontNoticeProps } from './FontNotice.tsx';
export {
  useHyperlinkPopup,
  useHyperlinkPopupInstance,
  type UseHyperlinkPopupInstanceReturn,
  type UseHyperlinkPopupReturn,
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
  type UsePptxSourceReturn,
} from './usePptxSource.ts';
export type {
  PptxFontOrigin,
  PptxFontsInput,
  PptxFontsSource,
  PptxSource,
  UsePptxSourceOptions,
} from '@pptx-editor.dev/core/editor';
export { HyperlinkPopup, type HyperlinkPopupProps } from './HyperlinkPopup.tsx';
export {
  PptxEditorHyperLink,
  type HyperLinkActionProps,
  type HyperLinkPartProps,
  type HyperLinkProps,
} from './HyperLink.tsx';
export {
  PptxEditorSlideshow,
  pptxSlideshowKey,
  useSlideshow,
  type UseSlideshowReturn,
} from './Slideshow.tsx';
export {
  PageSetupDialog,
  usePageSetupDialog,
  type PageSetupDialogFieldName,
  type PageSetupDialogFields,
  type PageSetupDialogProps,
  type UsePageSetupDialogReturn,
} from './PageSetupDialog.ts';
export {
  ImageAltTextDialog,
  useImageAltTextDialog,
  type ImageAltTextDialogFieldName,
  type ImageAltTextDialogFields,
  type ImageAltTextDialogProps,
  type UseImageAltTextDialogReturn,
} from './ImageAltTextDialog.ts';
export {
  ImagePropertiesDialog,
  useImagePropertiesDialog,
  type ImagePropertiesDialogFieldName,
  type ImagePropertiesDialogFields,
  type ImagePropertiesDialogProps,
  type UseImagePropertiesDialogReturn,
} from './ImagePropertiesDialog.ts';
export {
  BackgroundDialog,
  useBackgroundDialog,
  type BackgroundDialogFieldName,
  type BackgroundDialogFields,
  type BackgroundDialogProps,
  type UseBackgroundDialogReturn,
} from './BackgroundDialog.tsx';
export {
  ParagraphDialog,
  useParagraphDialog,
  type ParagraphDialogFieldName,
  type ParagraphDialogProps,
  type UseParagraphDialogReturn,
} from './ParagraphDialog.tsx';
export {
  FontDialog,
  useFontDialog,
  type FontDialogFieldName,
  type FontDialogProps,
  type UseFontDialogReturn,
} from './FontDialog.tsx';
export {
  BulletsAndNumberingDialog,
  useBulletsAndNumberingDialog,
  type BulletsAndNumberingDialogFieldName,
  type BulletsAndNumberingDialogProps,
  type UseBulletsAndNumberingDialogReturn,
} from './BulletsAndNumberingDialog.tsx';
export {
  ColumnsDialog,
  useColumnsDialog,
  type ColumnsDialogFieldName,
  type ColumnsDialogProps,
  type UseColumnsDialogReturn,
} from './ColumnsDialog.tsx';
export type { DialogCustomizationProps, DialogPartProps, UseDialogReturn } from './dialog-parts.ts';
export {
  Icon,
  MenuItem,
  Popover,
  ToolButton,
  keepFocus,
  type IconProps,
  type MenuItemProps,
  type PopoverProps,
  type ToolButtonProps,
} from './chrome.tsx';
export {
  MenuRow,
  MenuSeparator,
  MenuSubmenu,
  type MenuRowProps,
  type MenuSubmenuProps,
} from './menu-row.tsx';
export { ContextMenu, type ContextMenuPoint, type ContextMenuProps } from './ContextMenu.tsx';
export {
  MenuEntry,
  MenuGroup,
  MenuSlotItem,
  PptxEditorMenu,
  type MenuEntryProps,
  type MenuGroupProps,
  type MenuNamespace,
  type MenuPartComponent,
  type MenuProps,
  type MenuSlotItemProps,
  type MenuNamedRows,
  type MenuRowPartComponent,
  type MenuRowPartProps,
} from './menu-parts.tsx';
export type {
  ContextMenuPartName,
  MenuNamedRow,
  MenuRowPartName,
} from '@pptx-editor.dev/core/editor';
export {
  PptxEditorContextMenu,
  useContextMenuRequest,
  useContextMenuTarget,
  type ContextMenuNamedRows,
  type ContextMenuNamespace,
  type ContextMenuRequest,
  type PptxEditorContextMenuProps,
  type UseContextMenuRequestReturn,
  type UseContextMenuTargetReturn,
} from './context-menu-parts.tsx';
export {
  RailContextMenu,
  SectionContextMenu,
  type RailContextMenuProps,
  type RailMenuState,
  type SectionContextMenuProps,
  type SectionMenuState,
} from './RailContextMenu.tsx';
export { useNotesPane, type UseNotesPaneReturn } from './pane-state.ts';
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
export {
  PptxEditorSlideRail,
  PptxEditorSidePanel,
  PptxEditorZoomPill,
  type PaneProps,
  type SidePanelProps,
  type SlideRailProps,
  type ZoomPillProps,
} from './panes.tsx';
export { PptxEditorNotesPane, type NotesPaneProps } from './NotesPane.tsx';
export { PptxEditorSelectionPane, type SelectionPaneProps } from './SelectionPane.tsx';
// The toolbar's pane slots ask for a side panel tab; a host listens to show a closed panel.
export { listenPaneRequests, requestPane } from '@pptx-editor.dev/core/editor';
export { PptxEditorSlideSorter, type SlideSorterProps } from './SlideSorter.tsx';
export { requestSectionRename } from './RailSections.tsx';
export {
  PptxEditorLoading,
  PptxEditorLoadingSpinner,
  type LoadingProps,
  type LoadingSpinnerProps,
} from './loading.ts';
export { PptxEditorRibbon, RibbonTab, type RibbonProps, type RibbonTabProps } from './ribbon.ts';
export {
  PptxEditorStatusBar,
  StatusBarNotes,
  StatusBarSlideNumber,
  StatusBarViews,
  StatusBarZoom,
  type StatusBarNotesProps,
  type StatusBarPartProps,
  type StatusBarProps,
} from './status-bar.ts';
export {
  ToolbarControl,
  ToolbarGroup,
  ToolbarSlot,
  type ToolbarGroupProps,
  type ToolbarProps,
  type ToolbarSlotProps,
} from './toolbar.ts';
export type { ToolbarControlProps } from './toolbar-controls.tsx';
export {
  PptxEditorToolbar,
  ToolbarAction,
  ToolbarSeparator,
  type PptxEditorToolbarNamespace,
  type ToolbarActionProps,
  type ToolbarPartComponent,
  type ToolbarPartProps,
  type ToolbarSeparatorProps,
  type ToolbarColorSplitComponent,
  type ToolbarColorSplitProps,
  type ToolbarSlotPartComponent,
  type ToolbarSlotPartProps,
  type FontFamilyNamespace,
  type FontFamilyProps,
} from './toolbar-parts.ts';
export type { FontFamilyItemProps, FontFamilyPartProps } from './FontFamilyControl.tsx';

export { useFonts, type FontsInput, type UseFontsReturn } from './useFonts.ts';
export { useFontFamily, type UseFontFamilyReturn } from './useFontFamily.ts';
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
export { ShapeStyleGallery, type ShapeStyleGalleryProps } from './ShapeFormatControls.tsx';
export { MenuTableGrid, type MenuTableGridProps } from './TableControls.tsx';
export {
  LayoutControl,
  LayoutGallery,
  NewSlideControl,
  type LayoutControlProps,
  type LayoutGalleryProps,
  type NewSlideControlProps,
} from './SlideControls.tsx';
export { EditorFontError } from '@pptx-editor.dev/core/editor';
export type {
  EditorFontErrorCode,
  FontConfiguration,
  FontFaceRequest,
  FontSource,
  FontSourceSubstitution,
} from '@pptx-editor.dev/core/contracts';
// The font-composition surface, re-exported as DOCX Editor's adapters do, so the common path
// (a fonts package and an adapter) never needs a core import.
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
// The core types that the adapter's own props, hooks, and parts use, re-exported as DOCX
// Editor's adapters re-export `Editor`, `EditorCommand`, and `EditorSnapshot`, so a host types
// its handlers with an import from this package only.
export type {
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
  FontPickerGroup,
  FontPickerGroupId,
  FontPickerRow,
  GradientFill,
  PatternFill,
  PatternPreset,
  SidePanelTab,
  ParagraphDialogFields,
  PptxEditorInstance,
  RibbonTabId,
  ShapeStyleChoice,
  ShapeStyleValue,
  SlideBackground,
  SlideLayoutSummary,
  SlideRailOrientation,
  SlideSection,
  SlideshowController,
  ToolbarShownValue,
  ToolbarValueMap,
} from '@pptx-editor.dev/core/editor';

const pptxEditorProps = {
  /** `.pptx` bytes. Identity change loads the new document. */
  document: { type: Uint8Array as PropType<Uint8Array>, default: undefined },
  /** `'edit'` (default) or `'view'`. Changes apply without a reload. */
  mode: { type: String as PropType<EditorMode>, default: 'edit' },
  /** UI locale catalog, as on `PptxEditorRoot`. English when omitted. */
  locale: { type: Object as PropType<LocaleCatalog>, default: undefined },
  /** Families the font picker offers, as on `PptxEditorRoot`. A change applies in place. */
  fontFamilies: { type: Array as PropType<readonly string[]>, default: undefined },
  /** Capability modules (`@pptx-editor.dev/pro`), as on `PptxEditorRoot`. Read at mount. */
  modules: { type: Array as PropType<readonly EditorModule[]>, default: undefined },
  /** AutoCorrect while typing, as on `PptxEditorRoot`. Read at mount. */
  autoCorrect: { type: Boolean, default: true },
  /** Customize the dialogs the editor mounts, as on `PptxEditorRoot`. */
  popups: { type: Object as PropType<PptxEditorPopups>, default: undefined },
  /** The chrome color mode, as on `PptxEditorRoot`. */
  colorMode: { type: String as PropType<ColorMode>, default: undefined },
  /** Font bytes for the host's faces, as on `PptxEditorRoot`. */
  fonts: { type: [Object, Function] as PropType<EditorFontsOption>, default: undefined },
  /**
   * The notice for a file that does not open (`PptxEditorOpenFailure`). An object is
   * `PptxEditorOpenFailureProps`, passed straight through. `false` removes the notice, for a
   * host that reports the failure itself through `useOpenFailure`. For your own buttons, use
   * the `openFailureActions` slot: it replaces `openFailure.actions`, and with `false` it
   * renders nothing. Both cases print a development warning. Default `true`.
   */
  openFailure: {
    type: [Boolean, Object] as PropType<boolean | PptxEditorOpenFailureProps>,
    default: true,
  },
  /**
   * The toolbar (`PptxEditorToolbar`). An object is `ToolbarProps`, passed straight through.
   * `false` removes it. Default `true`.
   */
  toolbar: { type: [Boolean, Object] as PropType<boolean | ToolbarProps>, default: true },
  /**
   * The menu bar (`PptxEditorMenu`), above the toolbar. An object is `MenuProps`, passed
   * straight through, as DOCX Editor's `menu`. Default `false`: the packaged editor shows the
   * toolbar only.
   */
  menu: { type: [Boolean, Object] as PropType<boolean | MenuProps>, default: false },
  /**
   * The slide canvas's context menu. An object is `PptxEditorContextMenuProps`, passed straight
   * through. `false` shows no menu. It sets `popups.contextMenu`, and a `popups.contextMenu`
   * entry wins over it. Default `true`.
   */
  contextMenu: {
    type: [Boolean, Object] as PropType<boolean | PptxEditorContextMenuProps>,
    default: true,
  },
  /**
   * The packaged link popup. `false` removes it, as `popups.hyperlink: false` does: a host
   * panel that reads `useHyperlinkPopup()` then opens for the link controls. Default `true`.
   */
  hyperlinkPopup: { type: Boolean, default: true },
  /**
   * The slide rail (`PptxEditorSlideRail`), as DOCX Editor's `navigation`. An object is
   * `SlideRailProps`, passed straight through. `false` removes it. Default `true`.
   */
  slideRail: { type: [Boolean, Object] as PropType<boolean | SlideRailProps>, default: true },
  /**
   * The speaker notes pane (`PptxEditorNotesPane`). An object is `NotesPaneProps`, passed
   * straight through. `false` removes it. Default `true`.
   */
  notesPane: { type: [Boolean, Object] as PropType<boolean | NotesPaneProps>, default: true },
} as const;

/** The props an object value gives, or none for `true`. */
const partProps = <P extends object>(value: boolean | P | undefined): P | null =>
  typeof value === 'object' ? value : null;

/**
 * The `popups` the sugar passes to the root: the host's entries, with `contextMenu` and
 * `hyperlinkPopup` mapped onto the same entries, as React's sugar. A host entry wins.
 */
function sugarPopups(
  popups: PptxEditorPopups | undefined,
  hyperlinkPopup: boolean,
  contextMenu: boolean | PptxEditorContextMenuProps
): PptxEditorPopups | undefined {
  const linkOff = !hyperlinkPopup && popups?.hyperlink === undefined;
  const menuEntry =
    popups?.contextMenu !== undefined || contextMenu === true
      ? undefined
      : contextMenu === false
        ? (false as const)
        : (props: PptxEditorContextMenuProps) =>
            h(PptxEditorContextMenu, { ...props, ...contextMenu });
  if (!linkOff && menuEntry === undefined) return popups;
  return {
    ...popups,
    ...(linkOff ? { hyperlink: false as const } : {}),
    ...(menuEntry !== undefined ? { contextMenu: menuEntry } : {}),
  };
}

/**
 * Props of the `PptxEditor` sugar component, the twin of React's `PptxEditorProps`. Derived
 * from the component's runtime props, so the two cannot drift. @public
 */
export type PptxEditorProps = ExtractPublicPropTypes<typeof pptxEditorProps>;

const OPEN_FAILURE_ACTION_NAMES = {
  actions: 'The #openFailureActions slot',
  off: ':open-failure="false"',
} as const;

/**
 * The complete default editor: toolbar, slide rail, canvas, speaker notes, zoom pill, and the
 * slide show (F5), the twin of React's `PptxEditor`. Sugar over the same primitives a custom
 * layout composes. Emits `ready` with the instance. A template ref reads a `PptxEditorRef`.
 * @public
 */
const PptxEditorDefault = defineComponent({
  name: 'PptxEditor',
  // The root renders no element, so the host's attributes (`class`, `style`, `id`, `data-*`)
  // go onto the viewport, as React's `className` does.
  inheritAttrs: false,
  props: pptxEditorProps,
  emits: {
    ready: (_editor: PptxEditorInstance) => true,
    /** After every change of editor state, as on `PptxEditorRoot`. */
    change: (_editor: PptxEditorInstance) => true,
    /** A font failure, as on `PptxEditorRoot`. */
    fontError: (_error: EditorFontError) => true,
  },
  slots: Object as SlotsType<{
    /** Extra content inside the viewport, after the slide show. */
    default?: () => VNodeChild;
    /** Your own buttons in the open-failure notice, before **Close**. */
    openFailureActions?: (context: OpenFailureContext) => VNodeChild;
  }>,
  setup(props, { attrs, emit, expose, slots }) {
    // Read at render, not once: a parent re-render can add or remove the listener.
    const instance = getCurrentInstance();
    // The same members as the root's handle, read through the root at each call.
    const root = shallowRef<PptxEditorRef | null>(null);
    expose(createEditorRefApi(() => root.value?.getEditor() ?? null) satisfies PptxEditorRef);
    return () => {
      warnOpenFailureActions(
        props.openFailure,
        slots.openFailureActions !== undefined,
        OPEN_FAILURE_ACTION_NAMES
      );
      const popups = sugarPopups(props.popups, props.hyperlinkPopup, props.contextMenu);
      const dock =
        props.menu === false && props.toolbar === false
          ? null
          : h('div', { class: 'pptx-default__dock' }, [
              props.menu === false ? null : h(PptxEditorMenu, partProps(props.menu)),
              props.toolbar === false ? null : h(PptxEditorToolbar, partProps(props.toolbar)),
            ]);
      return h(
        PptxEditorRoot,
        {
          ref: root,
          ...(props.document ? { document: props.document } : {}),
          ...(props.locale ? { locale: props.locale } : {}),
          ...(props.fontFamilies ? { fontFamilies: props.fontFamilies } : {}),
          ...(popups ? { popups } : {}),
          ...(props.colorMode ? { colorMode: props.colorMode } : {}),
          ...(props.fonts ? { fonts: props.fonts } : {}),
          ...(props.modules ? { modules: props.modules } : {}),
          ...(props.autoCorrect === false ? { autoCorrect: false } : {}),
          // Forwarded only with a listener, so the root still warns when nobody listens.
          ...(instance?.vnode.props?.onFontError
            ? { onFontError: (error: EditorFontError) => emit('fontError', error) }
            : {}),
          mode: props.mode,
          onReady: (editor: PptxEditorInstance) => emit('ready', editor),
          onChange: (editor: PptxEditorInstance) => emit('change', editor),
        },
        {
          default: () =>
            h(PptxEditorViewport, mergeProps(attrs, { class: 'pptx-default' }), () => [
              dock,
              h(DefaultWorkspace, null, {
                overlay: () =>
                  props.openFailure === false
                    ? null
                    : h(
                        PptxEditorOpenFailure,
                        typeof props.openFailure === 'object' ? props.openFailure : null,
                        slots.openFailureActions ? { actions: slots.openFailureActions } : undefined
                      ),
                normal: () =>
                  h('div', { class: 'pptx-default__workspace' }, [
                    props.slideRail === false
                      ? null
                      : h(PptxEditorSlideRail, partProps(props.slideRail)),
                    h('div', { class: 'pptx-default__main' }, [
                      h(PptxEditorContent),
                      h(PptxEditorFindBar),
                      h(PptxEditorEditNotice),
                      h(PptxEditorLoading, { overlay: true }),
                      h(PptxEditorFontNotice),
                      props.notesPane === false
                        ? null
                        : h(PptxEditorNotesPane, partProps(props.notesPane)),
                    ]),
                  ]),
              }),
              h(PptxEditorSlideshow),
              slots.default?.(),
            ]),
        }
      );
    };
  },
});

/**
 * The default editor, with every part on it, the twin of React's `PptxEditor` namespace and of
 * DOCX Editor's Vue `DocxEditor`: `PptxEditor.Root`, `PptxEditor.Toolbar`, `PptxEditor.Menu`,
 * and the rest name the same parts as their `PptxEditor<Part>` exports. @public
 */
export const PptxEditor = Object.assign(PptxEditorDefault, {
  Root: PptxEditorRoot,
  Viewport: PptxEditorViewport,
  Content: PptxEditorContent,
  Toolbar: PptxEditorToolbar,
  ToolbarControl,
  SlideRail: PptxEditorSlideRail,
  SlideSorter: PptxEditorSlideSorter,
  NotesPane: PptxEditorNotesPane,
  ZoomPill: PptxEditorZoomPill,
  SidePanel: PptxEditorSidePanel,
  SelectionPane: PptxEditorSelectionPane,
  FontNotice: PptxEditorFontNotice,
  EditNotice: PptxEditorEditNotice,
  OpenFailure: PptxEditorOpenFailure,
  Loading: PptxEditorLoading,
  PageSetupDialog,
  BackgroundDialog,
  ImageAltTextDialog,
  ImagePropertiesDialog,
  ParagraphDialog,
  FontDialog,
  BulletsAndNumberingDialog,
  ColumnsDialog,
  Slideshow: PptxEditorSlideshow,
  FindBar: PptxEditorFindBar,
  Menu: PptxEditorMenu,
  ContextMenu: PptxEditorContextMenu,
  Ribbon: PptxEditorRibbon,
  StatusBar: PptxEditorStatusBar,
});
