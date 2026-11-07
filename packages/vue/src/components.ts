// Provider-first host for the editor facade: the Vue twin of the React adapter's
// `PptxEditor.Root` / `.Viewport` / `.Content` (packages/react/src/PptxEditorRoot.tsx), after
// DOCX Editor's Vue adapter (docx-editor/packages/vue/src/editor/DocxEditorRoot.ts).
//
// `PptxEditorRoot` renders no DOM of its own: it creates the instance on mount, provides it,
// and lets `PptxEditorContent` attach the surface wherever the host's template puts it. Props
// are declared once as a const; the public props type derives from it, so the two cannot
// drift (the parity check reads the runtime props).

import {
  defineComponent,
  getCurrentInstance,
  h,
  mergeProps,
  onMounted,
  onUnmounted,
  provide,
  shallowRef,
  watch,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
import {
  churnWarning,
  createEditorRefApi,
  createPptxEditor,
  localeContentKey,
  repaintSurfaceOf,
  reportFontError,
  translateKeyOrText,
  warnModulesChanged,
  type ColorMode,
  type CommandResult,
  type LoadOptions,
  type EditorCommand,
  type EditorExecOptions,
  type EditorSnapshot,
  type EditorFontError,
  type EditorFontsOption,
  type EditorMode,
  type EditorModule,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { createTranslator, type LocaleCatalog, type Translate } from '@pptx-editor.dev/i18n';
import {
  providePptxEditor,
  pptxTranslationKey,
  usePptxEditor,
  useTranslation,
} from './composables.ts';
import { ActiveDialog, provideDialogHost, type PptxEditorPopups } from './popups.ts';
import { provideChromeDark, provideScopedByAncestor, useChromeDark } from './color-mode.ts';
import { providePaneState } from './pane-state.ts';
import { provideSlideshow } from './Slideshow.tsx';
import { provideFindBars } from './FindBar.tsx';
import { LinkPopupHost } from './HyperlinkPopup.tsx';
import { provideHyperlinkPopup } from './useHyperlinkPopup.ts';
import { useCanvasContextMenu } from './CanvasContextMenu.tsx';
import { SlideSorterWatch } from './SlideSorter.tsx';
import { ChromeTooltip } from './ChromeTooltip.tsx';

const pptxEditorRootProps = {
  /** `.pptx` bytes. Identity change loads the new document. */
  document: { type: Uint8Array as PropType<Uint8Array>, default: undefined },
  /** `'edit'` (default) or `'view'`. Changes apply without a reload. */
  mode: { type: String as PropType<EditorMode>, default: 'edit' },
  /** UI locale catalog; English when omitted. */
  locale: { type: Object as PropType<LocaleCatalog>, default: undefined },
  /**
   * Font bytes for the host's faces, as DOCX Editor's `fonts` prop: a `FontConfiguration`, a
   * fragment, or a `FontResolver` such as `packagedFonts()`. Use `useFonts(...)` for a stable
   * value. A change replaces the fonts in place (`setFonts`).
   */
  fonts: { type: [Object, Function] as PropType<EditorFontsOption>, default: undefined },
  /**
   * Families the font picker offers besides the standard ones and the deck's own. A change
   * updates the picker in place (`setFontFamilies`); the document and the undo history stay.
   */
  fontFamilies: { type: Array as PropType<readonly string[]>, default: undefined },
  /**
   * Capability modules, as DOCX Editor's `modules` prop: the review and collaboration modules
   * of `@pptx-editor.dev/pro`. Read once, when the editor is created: a later change has no
   * effect and prints a development warning. Create the list once, outside the component.
   */
  modules: { type: Array as PropType<readonly EditorModule[]>, default: undefined },
  /**
   * AutoCorrect while typing (default `true`): curly quotes, symbols, fractions, en dashes,
   * sentence capitals, and lists started by typing `1.`, `1)`, `*`, or `-` and a space. Read
   * once, when the editor is created.
   */
  autoCorrect: { type: Boolean, default: true },
  /**
   * The author of later comments and replies, as DOCX Editor's `author` prop. A change applies
   * without a remount (`setAuthor`); existing comments keep their authors.
   */
  author: { type: String, default: undefined },
  /** Customize the dialogs the editor mounts. Set an entry to `false` to turn one off. */
  popups: { type: Object as PropType<PptxEditorPopups>, default: undefined },
  /** The chrome color mode: `'light'` (default), `'dark'`, or `'system'`. */
  colorMode: { type: String as PropType<ColorMode>, default: undefined },
} as const;

/** Props of `PptxEditorRoot`, the twin of React's `PptxEditorRootProps`. @public */
export type PptxEditorRootProps = ExtractPublicPropTypes<typeof pptxEditorRootProps>;

/**
 * The imperative handle of `PptxEditor` and `PptxEditorRoot`, the twin of React's
 * `PptxEditorRef` and of DOCX Editor's `DocxEditorRef`: a template ref on either component
 * reads these members (`expose`). Every member forwards to the editor instance and is safe to
 * call before the editor mounts: `load` and `exec` refuse with `disabled.noEditor`, `save` and
 * `getEditor` return `null`, `focus` does nothing, and `snapshot` returns an empty snapshot.
 * Everything else is on the instance that `getEditor` returns. @public
 */
export interface PptxEditorRef {
  /**
   * Open `.pptx` bytes. Resolves when the presentation is open, or with the refusal. A long
   * deck resolves when `options.startSlide` (or the first slide) is ready.
   */
  load(bytes: Uint8Array, options?: LoadOptions): Promise<CommandResult>;
  /**
   * The presentation as `.pptx` bytes. The bytes own an `ArrayBuffer`, so `new Blob([bytes])`
   * takes them as they are. `null` when no presentation is open, and when a slide of it is
   * damaged: `snapshot().readOnlyReason` is then `'damaged-slide'` (after edits), or
   * `snapshot().openFailure` reports the refused file.
   */
  save(): Uint8Array<ArrayBuffer> | null;
  /** The editor instance, or `null` before it mounts. */
  getEditor(): PptxEditorInstance | null;
  /** Give the keyboard back to the slide: the text caret, else the slide surface. */
  focus(): void;
  /** Run a command, as `editor.exec`. */
  exec(command: EditorCommand, options?: EditorExecOptions): CommandResult;
  /** The current editor state, as `editor.snapshot()`. */
  snapshot(): EditorSnapshot;
}

/** True for an element node from any window (an iframe or a pop-out has its own `HTMLElement`). */
function isElement(value: unknown): value is HTMLElement {
  return typeof value === 'object' && value !== null && (value as Node).nodeType === 1;
}

/**
 * Owns the editor instance, as React's `PptxEditor.Root`. Only `document` identity reloads
 * the document; `mode` and `locale` apply without a reload. Emits `ready` once per instance
 * and `change` after every change of editor state. @public
 */
export const PptxEditorRoot = defineComponent({
  name: 'PptxEditorRoot',
  props: pptxEditorRootProps,
  emits: {
    ready: (_editor: PptxEditorInstance) => true,
    change: (_editor: PptxEditorInstance) => true,
    /** A font failure: a failed origin, a refused face, or a failed resolution. */
    fontError: (_error: EditorFontError) => true,
  },
  setup(props, { emit, expose, slots }) {
    // With no `fontError` listener, the failure prints one warning instead of going nowhere.
    // The listener is read when the error occurs: a parent re-render can add or remove it.
    const instance = getCurrentInstance();
    const onFontError = (error: EditorFontError): void =>
      reportFontError(
        error,
        instance?.vnode.props?.onFontError
          ? (e: EditorFontError) => emit('fontError', e)
          : undefined
      );
    // The fonts the instance was created with: only a later change reaches `setFonts`.
    let createdFonts: EditorFontsOption | undefined;
    // The same for `fontFamilies` and `setFontFamilies`.
    let createdFamilies: readonly string[] | undefined;
    // The modules the instance was created with: modules register at creation only.
    let createdModules: readonly EditorModule[] | undefined;
    const editor = shallowRef<PptxEditorInstance | null>(null);
    const t = shallowRef<Translate>(createTranslator(props.locale));
    providePptxEditor(editor);
    provide(pptxTranslationKey, t);
    // The ref members read the live instance at each call, as React's `useImperativeHandle`.
    expose(createEditorRefApi(() => editor.value) satisfies PptxEditorRef);
    // The catalog's content, not its identity: an inline locale object makes a new value on
    // each parent render and must not retranslate or repaint, as React's root.
    watch(
      () => localeContentKey(props.locale),
      () => {
        t.value = createTranslator(props.locale);
        // The canvas repaints once with the new placeholder prompts and frame labels, and
        // emits no `change`, so a host whose `change` handler sets state does not loop.
        if (editor.value) repaintSurfaceOf(editor.value);
      }
    );

    let off: (() => void) | null = null;
    onMounted(() => {
      const instance = createPptxEditor({
        mode: props.mode,
        ...(props.fontFamilies ? { fontFamilies: props.fontFamilies } : {}),
        ...(props.fonts ? { fonts: props.fonts } : {}),
        ...(props.modules ? { modules: props.modules } : {}),
        ...(props.autoCorrect === false ? { autoCorrect: false } : {}),
        ...(props.author ? { author: props.author } : {}),
        onFontError,
        translate: (key, params) => translateKeyOrText(t.value, key, params),
      });
      createdFonts = props.fonts;
      createdFamilies = props.fontFamilies;
      createdModules = props.modules;
      editor.value = instance;
      emit('ready', instance);
      off = instance.subscribe(() => emit('change', instance));
      if (props.document) void instance.load(props.document);
    });
    // After the children unmounted, as DOCX Editor's root: Content detaches first.
    onUnmounted(() => {
      off?.();
      off = null;
      const instance = editor.value;
      editor.value = null;
      instance?.destroy();
    });
    // In development, a `document` or `fonts` made anew on each render warns.
    const documentChurn = churnWarning('PptxEditorRoot', 'document');
    const fontsChurn = churnWarning('PptxEditorRoot', 'fonts');
    watch(
      () => props.document,
      (bytes) => {
        if (!editor.value || !bytes) return;
        documentChurn();
        void editor.value.load(bytes);
      }
    );
    watch(
      () => props.mode,
      (mode) => editor.value?.setMode(mode)
    );
    // A new author applies to later comments; the instance ignores the same name.
    watch(
      () => props.author,
      (author) => editor.value?.setAuthor(author)
    );
    // A `modules` change after creation has no effect: say so in development, as React's root.
    // The check is by the identity of each module, so a new array of the same modules does not
    // warn.
    watch(
      () => props.modules,
      (modules) => {
        if (editor.value) warnModulesChanged(createdModules, modules, 'PptxEditorRoot');
      }
    );
    // A new `fonts` value replaces the instance's fonts in place, as DOCX Editor's root.
    watch(
      () => [editor.value, props.fonts] as const,
      ([instance, fonts]) => {
        if (!instance || createdFonts === fonts) return;
        createdFonts = fonts;
        fontsChurn();
        // Rejects only after destroy; font failures report through `fontError`.
        instance.setFonts(fonts).catch(() => {});
      },
      { flush: 'post' }
    );
    // A new `fontFamilies` list updates the font picker in place. The instance compares by
    // content, so an inline array made on each render changes nothing.
    watch(
      () => [editor.value, props.fontFamilies] as const,
      ([instance, families]) => {
        if (!instance || createdFamilies === families) return;
        createdFamilies = families;
        instance.setFontFamilies(families ?? []);
      },
      { flush: 'post' }
    );
    // The slide show and the find bar count, as React's root publishes them; the action host
    // starts the show and opens find and the link popup through them.
    const slideshow = provideSlideshow(editor);
    const findBars = provideFindBars();
    // The root's one link popup state, which the packaged popup and `useHyperlinkPopup` share.
    const link = provideHyperlinkPopup();
    const host = provideDialogHost(editor, () => props.popups, {
      slideshow,
      findBars,
      linkViews: link.views,
    });
    provideChromeDark(() => props.colorMode);
    // The notes pane's open state, which the pane and host chrome share (`useNotesPane`).
    providePaneState();
    // The root hosts the link popup, so Ctrl/Cmd+K and Insert > Link work with no toolbar.
    return () => [
      slots.default?.(),
      h(SlideSorterWatch),
      h(ActiveDialog, { host }),
      h(LinkPopupHost, { store: link.store, views: link.views.value }),
      h(ChromeTooltip),
    ];
  },
});

const pptxEditorViewportProps = {
  /** A class name added to the viewport, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
  /**
   * The name of the editor's landmark region, or a translation key. Default: "Presentation
   * editor" (`viewport.ariaLabel`). Give each editor its own name when a page shows more than one.
   * `false` renders no region, for a host whose own landmarks (`header`, `main`) are inside.
   */
  label: { type: [String, Boolean] as PropType<string | false>, default: undefined },
} as const;

/** The props of `PptxEditorViewport`, the twin of React's `PptxEditorViewportProps`. @public */
export type PptxEditorViewportProps = ExtractPublicPropTypes<typeof pptxEditorViewportProps>;

/**
 * The editor's scroll and styling scope: carries the `.pptx-editor` class, and `dark` with it
 * when the root's `colorMode` resolves to dark. It is a named landmark region ("Presentation
 * editor"), so a screen reader user can go to the editor from anywhere on the page. @public
 */
export const PptxEditorViewport = defineComponent({
  name: 'PptxEditorViewport',
  props: pptxEditorViewportProps,
  setup(props, { slots }) {
    const dark = useChromeDark();
    const { t } = useTranslation();
    provideScopedByAncestor();
    // The host's `class` attribute merges onto this element through attribute fallthrough.
    return () =>
      h(
        'div',
        {
          class: ['pptx-editor', dark.value ? 'dark' : null, props.className],
          ...(props.label === false
            ? {}
            : {
                role: 'region',
                'aria-label': props.label
                  ? translateKeyOrText(t, props.label)
                  : t('viewport.ariaLabel'),
              }),
        },
        slots.default?.()
      );
  },
});

const pptxEditorContentProps = {
  /** A class name added to the canvas, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** The props of `PptxEditorContent`, the twin of React's `PptxEditorContentProps`. @public */
export type PptxEditorContentProps = ExtractPublicPropTypes<typeof pptxEditorContentProps>;

/**
 * The slide canvas: attaches the engine surface while it is mounted. Right-click, Shift+F10, and
 * the Menu key open its context menu. @public
 */
export const PptxEditorContent = defineComponent({
  name: 'PptxEditorContent',
  // The canvas renders next to its context menu layer, so the host's attributes go on the canvas.
  inheritAttrs: false,
  props: pptxEditorContentProps,
  setup(props, { attrs }) {
    const editorRef = usePptxEditor();
    // Right-click, Shift+F10, and the Menu key open the canvas context menu, as React's content.
    const menu = useCanvasContextMenu();
    const host = shallowRef<HTMLElement | null>(null);
    let attached: PptxEditorInstance | null = null;
    const sync = (): void => {
      const editor = editorRef.value;
      const el = host.value;
      if (attached && (attached !== editor || !el)) {
        attached.detach();
        attached = null;
      }
      if (editor && el && attached !== editor) {
        editor.attach(el);
        attached = editor;
      }
    };
    onMounted(sync);
    watch(editorRef, sync);
    onUnmounted(() => {
      attached?.detach();
      attached = null;
    });
    return () => [
      // `mergeProps` keeps a host's own `@contextmenu` and `@keydown` next to the menu's.
      h(
        'div',
        mergeProps(attrs, {
          ref: (el: unknown) => {
            host.value = isElement(el) ? el : null;
          },
          class: ['pptx-canvas', props.className],
          'data-pptx-canvas': '',
          onContextmenu: menu.onContextmenu,
          onKeydown: menu.onKeydown,
        })
      ),
      menu.layer(),
    ];
  },
});
