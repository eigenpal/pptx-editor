// The popups registry: the Vue twin of the React adapter's `popups.tsx`, after DOCX Editor's
// Vue adapter (docx-editor/packages/vue/src/editor/popup-config.ts, dialog-host.ts).
// `PptxEditorRoot` mounts the editor's dialogs; a host changes one through `popups`: a render
// callback replaces the packaged dialog, and `false` turns off automatic rendering. Every
// opener calls `useEditorDialog(key).open()`, focus returns to the opener on close, and a
// dialog closes when the editor opens another document.

import { markReadingPopup } from '@pptx-editor.dev/core/editor';
import {
  computed,
  defineComponent,
  h,
  inject,
  provide,
  shallowRef,
  watch,
  type AllowedComponentProps,
  type Component,
  type ComponentCustomProps,
  type ComputedRef,
  type FunctionalComponent,
  type InjectionKey,
  type PropType,
  type Ref,
  type ShallowRef,
  type VNodeChild,
  type VNodeProps,
} from 'vue';
import {
  SLOT_POPUP_KEYS,
  chromeControlForSlot,
  dispatchFindRequest,
  pickImageInto,
  requestLinkPopup,
  warnDialogNeedsRenderer,
  type ChromeActionHost,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { chromeActionHostKey } from './composables.ts';
import { PageSetupDialog, type PageSetupDialogProps } from './PageSetupDialog.ts';
import { ImageAltTextDialog, type ImageAltTextDialogProps } from './ImageAltTextDialog.ts';
import { ImagePropertiesDialog, type ImagePropertiesDialogProps } from './ImagePropertiesDialog.ts';
import { EquationDialog, type EquationDialogProps } from './EquationDialog.ts';
import { BackgroundDialog, type BackgroundDialogProps } from './BackgroundDialog.tsx';
import { ParagraphDialog, type ParagraphDialogProps } from './ParagraphDialog.tsx';
import { FontDialog, type FontDialogProps } from './FontDialog.tsx';
import {
  BulletsAndNumberingDialog,
  type BulletsAndNumberingDialogProps,
} from './BulletsAndNumberingDialog.tsx';
import { ColumnsDialog, type ColumnsDialogProps } from './ColumnsDialog.tsx';
import type { PptxEditorContextMenuProps } from './context-menu-parts.tsx';
import type { RailContextMenuProps, SectionContextMenuProps } from './RailContextMenu.tsx';
import type { HyperlinkPopupProps } from './HyperlinkPopup.tsx';
import type { FindPopupProps } from './FindBar.tsx';
import type { UseSlideshowReturn } from './Slideshow.tsx';

/** A render callback for an automatic popup. `false` disables automatic rendering. @public */
export type PptxEditorPopup<Props extends object> = false | ((props: Props) => VNodeChild);

/**
 * Register an ordinary component in the `popups` map while keeping its props and state, as
 * React's `definePopup` and DOCX Editor's Vue `definePopup`: the component's required and
 * narrowed props stay checked against the entry; Vue's own attributes do not constrain it.
 * @public
 */
export function definePopup<C extends Component | FunctionalComponent<never>>(
  component: C,
  options?: DefinePopupOptions
): (
  props: C extends new (...args: never[]) => { $props: infer P }
    ? Omit<
        P,
        {
          [K in keyof P]-?: K extends keyof (VNodeProps &
            AllowedComponentProps &
            ComponentCustomProps)
            ? {} extends Pick<P, K>
              ? (VNodeProps & AllowedComponentProps & ComponentCustomProps)[K] extends P[K]
                ? K
                : never
              : never
            : never;
        }[keyof P]
      >
    : C extends (props: infer P, ...args: never[]) => unknown
      ? P
      : never
) => VNodeChild {
  const render = (props: object): VNodeChild =>
    h(component as Component, props as Record<string, unknown>);
  return options?.reading ? markReadingPopup(render) : render;
}

/** Options of `definePopup`, the twin of React's `DefinePopupOptions`. @public */
export interface DefinePopupOptions {
  /**
   * For `popups.hyperlink`: the renderer shows reading mode too, so a click into linked text
   * opens it with `mode: 'reading'`. Without it a renderer opens for the URL field alone.
   */
  readonly reading?: boolean;
}

/** The props a registry dialog receives: it shows, and how to close it. @public */
export interface PptxEditorDialogSessionProps {
  /** Always true while the registry renders the dialog. */
  readonly open: boolean;
  /** Close the dialog: the registry stops rendering it and focus returns to the opener. */
  readonly onClose: () => void;
}

/**
 * Render overrides for the dialogs the editor mounts. A render callback replaces the packaged
 * dialog; `false` disables one. @public
 */
export interface PptxEditorPopups {
  /** File > Page setup: slide size, appearance, and fonts. */
  pageSetup?: PptxEditorPopup<PageSetupDialogProps>;
  /**
   * The link popup (Ctrl/Cmd+K, the link control, Insert > Link). `PptxEditorRoot` hosts it, so
   * it opens with no toolbar on the page. `false` turns links off in the chrome; `setLink` keeps
   * working from the API.
   */
  hyperlink?: PptxEditorPopup<HyperlinkPopupProps>;
  /**
   * The find and replace bar (Ctrl/Cmd+F, Ctrl/Cmd+H), rendered where `PptxEditorFindBar`
   * stands. `false` turns it off; `editor.find` keeps working from the API.
   */
  find?: PptxEditorPopup<FindPopupProps>;
  /**
   * The slide canvas's context menu. A callback usually renders `PptxEditorContextMenu` with
   * host rows; `false` shows no menu. A `PptxEditorContextMenu` placed in the root or the
   * viewport gives the same entry, and this entry wins over it.
   */
  contextMenu?: PptxEditorPopup<PptxEditorContextMenuProps>;
  /**
   * The context menu of a slide in the slide rail and of a tile in the slide sorter. A callback
   * renders the host's menu, or `RailContextMenu` with its own options; `false` shows no menu.
   */
  railContextMenu?: PptxEditorPopup<RailContextMenuProps>;
  /**
   * The menu of a section header in the slide rail and in the slide sorter. A callback renders
   * the host's menu, or `SectionContextMenu` with its own options; `false` shows no menu.
   */
  sectionContextMenu?: PptxEditorPopup<SectionContextMenuProps>;
  /** Format background: slide background fill. */
  background?: PptxEditorPopup<BackgroundDialogProps>;
  /** The Alt Text dialog of the selected element, as DOCX Editor's `imageAltText`. */
  imageAltText?: PptxEditorPopup<ImageAltTextDialogProps>;
  /** The Format Picture dialog of the selected picture, as DOCX Editor's `imageProperties`. */
  imageProperties?: PptxEditorPopup<ImagePropertiesDialogProps>;
  /**
   * Insert > Equation: a gallery and the linear form of an equation, which inserts an equation
   * or edits the selected one. `false` turns it off; `insertEquation` and `setEquation` keep
   * working from the API.
   */
  equation?: PptxEditorPopup<EquationDialogProps>;
  /** Format > Paragraph: alignment, indents, spacing before and after, and line spacing. */
  paragraph?: PptxEditorPopup<ParagraphDialogProps>;
  /** Format > Font: font, effects, offset, character spacing, and kerning. */
  font?: PptxEditorPopup<FontDialogProps>;
  /** Format > Bullets and Numbering: bullet or number style, size, color, and start number. */
  bulletsAndNumbering?: PptxEditorPopup<BulletsAndNumberingDialogProps>;
  /** Home > Columns > More Columns: the number of text columns and the spacing between them. */
  columns?: PptxEditorPopup<ColumnsDialogProps>;
}

/** The `popups` entries that are menus, not dialogs: their parts read them (`usePopupEntry`). */
type MenuPopupKey = 'contextMenu' | 'railContextMenu' | 'sectionContextMenu';

/** The `popups` entries of the overlays: their parts read them (`usePopupEntry`). */
type OverlayPopupKey = 'hyperlink' | 'find';

/** The dialogs the registry mounts. @public */
export type PptxEditorDialogKey = Exclude<keyof PptxEditorPopups, MenuPopupKey | OverlayPopupKey>;

/** What `useEditorDialog` returns, the twin of React's `UseEditorDialogReturn`. @public */
export interface UseEditorDialogReturn {
  /** Show the dialog. Focus returns to `returnFocusTo`, or the element with focus now. */
  open(returnFocusTo?: HTMLElement | null): void;
  /** Hide the dialog. */
  close(): void;
  /** Whether the dialog shows now. */
  readonly isOpen: ComputedRef<boolean>;
  /** False when the dialog cannot show: hide or disable the control that opens it. */
  readonly isAvailable: ComputedRef<boolean>;
}

/** The canvas menu entry that a `PptxEditorContextMenu` placed in the root gives. */
type CanvasMenuEntry = PptxEditorPopup<PptxEditorContextMenuProps>;

/**
 * The `PptxEditorContextMenu` placed in the root or the viewport: it gives the canvas its menu
 * entry, as a `popups.contextMenu` entry does. The last one placed counts. React's twin is
 * `CanvasMenuRegistry`. @internal
 */
export interface CanvasMenuRegistry {
  get(): CanvasMenuEntry | undefined;
  place(token: object, entry: () => CanvasMenuEntry): void;
  remove(token: object): void;
}

function createCanvasMenuRegistry(): CanvasMenuRegistry {
  const current = shallowRef<{ readonly token: object; readonly entry: () => CanvasMenuEntry }>();
  return {
    get: () => current.value?.entry(),
    place(token, entry) {
      current.value = { token, entry };
    },
    remove(token) {
      if (current.value?.token === token) current.value = undefined;
    },
  };
}

interface DialogHost {
  readonly active: ShallowRef<PptxEditorDialogKey | null>;
  readonly popups: () => PptxEditorPopups | undefined;
  readonly canvasMenus: CanvasMenuRegistry;
  open(key: PptxEditorDialogKey, returnFocusTo?: HTMLElement | null): void;
  close(): void;
}

const dialogHostKey: InjectionKey<DialogHost> = Symbol('pptxDialogHost');

/** The packaged dialog of each registry key, which `ActiveDialog` renders without a host renderer. */
const PACKAGED: Readonly<Partial<Record<PptxEditorDialogKey, Component>>> = {
  pageSetup: PageSetupDialog,
  background: BackgroundDialog,
  imageAltText: ImageAltTextDialog,
  imageProperties: ImagePropertiesDialog,
  equation: EquationDialog,
  paragraph: ParagraphDialog,
  font: FontDialog,
  bulletsAndNumbering: BulletsAndNumberingDialog,
  columns: ColumnsDialog,
};

/** Whether the registry can show `key`: no `false` entry, and a renderer to show it. */
function available(popups: PptxEditorPopups | undefined, key: PptxEditorDialogKey): boolean {
  const entry = popups?.[key];
  if (entry === false) return false;
  // Every registry dialog is packaged in Vue, as in React; a host renderer replaces it.
  return key in PACKAGED || typeof entry === 'function';
}

/**
 * The overlays of one root: its slide show, and how many find bars it has mounted. The root
 * passes them, so the action host starts the show, opens find, and opens the link popup.
 * @internal
 */
export interface OverlayHost {
  readonly slideshow: UseSlideshowReturn;
  readonly findBars: Readonly<Ref<number>>;
  /** How many host panels show the link popup state (`useHyperlinkPopup`). */
  readonly linkViews?: Readonly<Ref<number>>;
}

/** Why an overlay slot (find, the link popup) cannot run now, as React's root says, or null. */
function overlayReason(
  popups: PptxEditorPopups | undefined,
  editor: PptxEditorInstance | null,
  overlays: OverlayHost,
  slot: Parameters<NonNullable<ChromeActionHost['unavailable']>>[0]
): string | null {
  const key = SLOT_POPUP_KEYS[slot];
  // With `hyperlink: false`, a host panel that reads `useHyperlinkPopup()` still opens.
  if (key === 'hyperlink')
    return popups?.hyperlink === false && !overlays.linkViews?.value ? 'link.unavailable' : null;
  if (key !== 'find') return null;
  if (popups?.find === false) return 'find.unavailable';
  // Find shows its matches on the slide canvas, which the slide sorter does not show.
  if (editor?.snapshot().presentationView === 'slideSorter') return 'disabled.findSorter';
  // Find needs a find bar to open: nothing answers the request without one.
  return overlays.findBars.value === 0 ? 'find.noBar' : null;
}

/**
 * The root's `ChromeActionHost`: it opens the dialogs the root renders, starts the slide show,
 * opens the find bar and the link popup, picks pictures, and says why a slot cannot run (a
 * `false` popups entry, a dialog with no renderer, no find bar). Core's `runChromeRow` and
 * `chromeRowState` use it.
 */
function actionHost(
  dialogs: DialogHost,
  editorRef: Readonly<Ref<PptxEditorInstance | null>>,
  overlays: OverlayHost
): ChromeActionHost {
  return {
    openDialog: (key) => dialogs.open(key as PptxEditorDialogKey),
    startSlideshow: (from) => overlays.slideshow.start(from),
    openFind: (replace) => {
      dispatchFindRequest(editorRef.value, replace);
    },
    openLink: () => {
      requestLinkPopup(editorRef.value);
    },
    pickImage: (target) => {
      if (editorRef.value && typeof document !== 'undefined')
        pickImageInto(editorRef.value, document, target);
    },
    unavailable: (slot, dialog) => {
      const state = chromeControlForSlot(slot)?.state;
      const reason = overlayReason(dialogs.popups(), editorRef.value, overlays, slot);
      if (reason) return reason;
      const key = (dialog ?? (state?.kind === 'dialog' ? state.dialog : undefined)) as
        PptxEditorDialogKey | undefined;
      if (key === undefined || available(dialogs.popups(), key)) return null;
      return 'dialog.unavailable';
    },
  };
}

/** The element that should take focus back, or null when it is gone or is the page itself. */
function focusTarget(el: Element | null | undefined): HTMLElement | null {
  if (!el || el.nodeType !== 1 || el === el.ownerDocument.body) return null;
  return el as HTMLElement;
}

/**
 * Provide the registry for one root and return the host. The root passes its editor: an
 * `inject` in the providing component would read the parent's value, not its own. @internal
 */
export function provideDialogHost(
  editorRef: Readonly<Ref<PptxEditorInstance | null>>,
  popups: () => PptxEditorPopups | undefined,
  overlays: OverlayHost
): DialogHost {
  const active = shallowRef<PptxEditorDialogKey | null>(null);
  let opener: HTMLElement | null = null;
  const host: DialogHost = {
    active,
    popups,
    canvasMenus: createCanvasMenuRegistry(),
    open(key, returnFocusTo) {
      if (!available(popups(), key)) {
        // A `false` entry turns the dialog off on purpose; no renderer is a host mistake.
        if (popups()?.[key] !== false) warnDialogNeedsRenderer('Vue', key);
        return;
      }
      opener =
        returnFocusTo === undefined
          ? focusTarget(typeof document === 'undefined' ? null : document.activeElement)
          : returnFocusTo;
      active.value = key;
    },
    close() {
      active.value = null;
    },
  };
  // Focus returns to the opener after the dialog closes. `post`: after the dialog unmounts and
  // leaves its modal state, so the page is no longer inert and the opener can take focus.
  watch(
    active,
    (key) => {
      if (key !== null) return;
      const target = opener;
      opener = null;
      if (target?.isConnected) target.focus({ preventScroll: true });
    },
    { flush: 'post' }
  );
  // A dialog belongs to the document it opened on: close it when another document opens.
  watch(
    editorRef,
    (editor, _old, onCleanup) => {
      if (!editor) return;
      let session = editor.session;
      const off = editor.subscribe(() => {
        if (editor.session === session) return;
        session = editor.session;
        active.value = null;
      });
      onCleanup(off);
    },
    { immediate: true }
  );
  // A popup entry that becomes `false` closes its open dialog.
  watch(popups, (next) => {
    if (active.value !== null && !available(next, active.value)) active.value = null;
  });
  provide(dialogHostKey, host);
  provide(chromeActionHostKey, actionHost(host, editorRef, overlays));
  return host;
}

/** Renders the dialog that is open. Placed by `PptxEditorRoot`. @internal */
export const ActiveDialog = defineComponent({
  name: 'PptxActiveDialog',
  props: {
    host: { type: Object as PropType<DialogHost>, required: true },
  },
  setup(props) {
    return () => {
      const key = props.host.active.value;
      if (key === null) return null;
      const session = { open: true, onClose: () => props.host.close() };
      const entry = props.host.popups()?.[key];
      if (entry === false) return null;
      if (typeof entry === 'function') return (entry as (p: typeof session) => VNodeChild)(session);
      const packaged = PACKAGED[key];
      return packaged ? h(packaged, session) : null;
    };
  },
});

/**
 * Open, close, and read one of the editor's dialogs, as React's `useEditorDialog`. Outside
 * `PptxEditorRoot` it reports the dialog as unavailable. @public
 */
export function useEditorDialog(key: PptxEditorDialogKey): UseEditorDialogReturn {
  const host = inject(dialogHostKey, null);
  return {
    open: (returnFocusTo?: HTMLElement | null) => host?.open(key, returnFocusTo),
    close: () => host?.close(),
    isOpen: computed(() => host?.active.value === key),
    isAvailable: computed(() => host !== null && available(host.popups(), key)),
  };
}

/**
 * The registry entry for `key` in the enclosing `PptxEditorRoot`, read at call time: a renderer,
 * `false`, or undefined for the packaged part. The parts that own a menu popup (the canvas, the
 * rail, a section header) read it, as React's `usePopupEntry`. @internal
 */
export function usePopupEntry<K extends keyof PptxEditorPopups>(
  key: K
): () => PptxEditorPopups[K] | undefined {
  const host = inject(dialogHostKey, null);
  return () => host?.popups()?.[key];
}

/**
 * The canvas menu entry of the enclosing `PptxEditorRoot`, read at call time: its
 * `popups.contextMenu` entry, else the entry of a `PptxEditorContextMenu` placed in the root,
 * else undefined for the packaged menu. An explicit `popups.contextMenu` entry wins, as in
 * React's `useCanvasMenuEntry`. @internal
 */
export function useCanvasMenuEntry(): () => CanvasMenuEntry | undefined {
  const host = inject(dialogHostKey, null);
  return () => {
    const explicit = host?.popups()?.contextMenu;
    return explicit !== undefined ? explicit : host?.canvasMenus.get();
  };
}

/** Where a `PptxEditorContextMenu` placed in the root registers, or null outside a root. @internal */
export function useCanvasMenuRegistry(): CanvasMenuRegistry | null {
  return inject(dialogHostKey, null)?.canvasMenus ?? null;
}
