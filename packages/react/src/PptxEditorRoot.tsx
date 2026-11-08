/** @jsxImportSource react */
// Provider-first host for the editor facade
// (docx-editor/packages/react/src/editor/DocxEditorRoot.tsx). DOCX Editor commit adf0e178.
//
// `Root` renders no DOM of its own: it creates the instance in an effect (StrictMode-safe: a
// fresh instance per effect run, destroyed on cleanup), publishes it through context, and lets
// `Content` attach the surface wherever the host's tree puts it.

import {
  forwardRef,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from 'react';
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
  type ClipboardDrawingReader,
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
import type { LocaleCatalog, Translate } from '@pptx-editor.dev/i18n';
import {
  PptxEditorContext,
  SlideshowContext,
  TranslationContext,
  translatorFor,
  usePptxEditor,
  useTranslation,
} from './context.ts';
import { useSlideshowHandle } from './Slideshow.tsx';
import { useCanvasContextMenu } from './CanvasContextMenu.tsx';
import { DialogProvider, type PptxEditorPopups } from './popups.tsx';
import { ChromeActionHostProvider } from './chrome-actions.tsx';
import { SlideSorterWatch } from './SlideSorter.tsx';
import { ChromeTooltip } from './ChromeTooltip.tsx';
import {
  ChromeDarkContext,
  ScopedByAncestorContext,
  useChromeDark,
  useResolvedDark,
} from './color-mode.ts';
import { PaneStateContext, type PaneState } from './pane-state.ts';

/** Props for `PptxEditor.Root`. Only `document` identity reloads the document. @public */
export interface PptxEditorRootProps {
  /** `.pptx` bytes. Identity change loads the new document. */
  readonly document?: Uint8Array;
  /** `'edit'` (default) or `'view'`. Changes apply without a reload. */
  readonly mode?: EditorMode;
  /** UI locale catalog; English when omitted. */
  readonly locale?: LocaleCatalog;
  /** Fired once per instance after it is published. */
  readonly onReady?: (editor: PptxEditorInstance) => void;
  /** Fired after every change of editor state. */
  readonly onChange?: (editor: PptxEditorInstance) => void;
  /**
   * Font bytes for the host's faces, as DOCX Editor's `fonts` prop: a `FontConfiguration`, a
   * fragment, or a `FontResolver` such as `packagedFonts()`. Use `useFonts(...)` for a stable
   * value. A change replaces the fonts in place (`setFonts`); the document, the selection, and
   * the undo history stay. Omitted, the packaged substitutes load and nothing is fetched.
   */
  readonly fonts?: EditorFontsOption;
  /** Called for each font failure: a failed origin, a refused face, or a failed resolution. */
  readonly onFontError?: (error: EditorFontError) => void;
  /**
   * Families the font picker offers besides the standard ones and the deck's own. A change
   * updates the picker in place (`setFontFamilies`); the document and the undo history stay.
   */
  readonly fontFamilies?: readonly string[];
  /**
   * Capability modules, as DOCX Editor's `modules` prop: the review and collaboration modules
   * of `@pptx-editor.dev/pro`. Read once, when the editor is created: a later change has no
   * effect and prints a development warning. Create the list once, outside the component.
   */
  readonly modules?: readonly EditorModule[];
  /**
   * AutoCorrect while typing (default `true`): curly quotes, symbols, fractions, en dashes,
   * sentence capitals, and lists started by typing `1.`, `1)`, `*`, or `-` and a space. Read
   * once, when the editor is created.
   */
  readonly autoCorrect?: boolean;
  /**
   * Reads the DrawingML clipboard package from the system clipboard, for a desktop host that
   * can. A paste of copied shapes from another application then pastes editable shapes, not
   * their picture. See `PptxEditorOptions.readClipboardDrawing`. Read once, when the editor is
   * created.
   */
  readonly readClipboardDrawing?: ClipboardDrawingReader;
  /**
   * The author of later comments and replies, as DOCX Editor's `author` prop. A change applies
   * without a remount (`setAuthor`); existing comments keep their authors.
   */
  readonly author?: string;
  /** Customize the dialogs the editor mounts. Set an entry to `false` to turn one off. */
  readonly popups?: PptxEditorPopups;
  /**
   * The chrome color mode: `'light'` (default), `'dark'`, or `'system'`, which follows the
   * operating system. Only the chrome is themed; the slide keeps the deck's own colors.
   */
  readonly colorMode?: ColorMode;
  /** The editor's layout: a `PptxEditor.Viewport` with its parts. */
  readonly children?: ReactNode;
}

/**
 * The imperative handle of `PptxEditor` and `PptxEditor.Root`, as DOCX Editor's
 * `DocxEditorRef`, and the same members as the Vue adapter's. Every member forwards to the
 * editor instance and is safe to call before the editor mounts: `load` and `exec` refuse with
 * `disabled.noEditor`, `save` and `getEditor` return `null`, `focus` does nothing, and
 * `snapshot` returns an empty snapshot. Everything else is on the instance that `getEditor`
 * returns. @public
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

/** Owns the editor instance. Takes a `ref` to a `PptxEditorRef`. @public */
export const PptxEditorRoot = forwardRef<PptxEditorRef, PptxEditorRootProps>(
  function PptxEditorRoot(props, ref) {
    const { document: bytes, mode = 'edit', locale, onReady, onChange, children } = props;
    const [editor, setEditor] = useState<PptxEditorInstance | null>(null);
    // The live instance for the ref: read at each call, so the handle never goes stale.
    const live = useRef<PptxEditorInstance | null>(null);
    useImperativeHandle(ref, () => createEditorRefApi(() => live.current), []);
    const callbacks = useRef({ onReady, onChange, onFontError: props.onFontError });
    callbacks.current = { onReady, onChange, onFontError: props.onFontError };
    // The fonts the instance was created with: only a later change reaches `setFonts`.
    const createdFonts = useRef(props.fonts);
    // The same for `fontFamilies` and `setFontFamilies`.
    const createdFamilies = useRef(props.fontFamilies);
    // The modules the instance was created with: modules register at creation only.
    const createdModules = useRef(props.modules);

    // The translator follows the catalog's content, not its identity: an inline `locale={{...}}`
    // makes a new object on each render and must not retranslate or repaint.
    const localeKey = localeContentKey(locale);
    const translator = useRef<{ key: string; t: ReturnType<typeof translatorFor> } | null>(null);
    if (translator.current?.key !== localeKey)
      translator.current = { key: localeKey, t: translatorFor(locale) };
    const t = translator.current.t;
    const tRef = useRef(t);
    tRef.current = t;
    useEffect(() => {
      const instance = createPptxEditor({
        mode,
        ...(props.fontFamilies ? { fontFamilies: props.fontFamilies } : {}),
        ...(props.fonts ? { fonts: props.fonts } : {}),
        ...(props.modules ? { modules: props.modules } : {}),
        ...(props.autoCorrect === false ? { autoCorrect: false } : {}),
        ...(props.readClipboardDrawing ? { readClipboardDrawing: props.readClipboardDrawing } : {}),
        ...(props.author ? { author: props.author } : {}),
        onFontError: (error) => reportFontError(error, callbacks.current.onFontError),
        translate: (key, params) => translateKeyOrText(tRef.current, key, params),
      });
      createdFonts.current = props.fonts;
      createdFamilies.current = props.fontFamilies;
      createdModules.current = props.modules;
      live.current = instance;
      setEditor(instance);
      callbacks.current.onReady?.(instance);
      const off = instance.subscribe(() => callbacks.current.onChange?.(instance));
      return () => {
        off();
        if (live.current === instance) live.current = null;
        instance.destroy();
      };
      // The instance lives for the component; mode and document apply through setters below.
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    // In development, a `document` or `fonts` made anew on each render warns.
    const [documentChurn] = useState(() => churnWarning('PptxEditor.Root', 'document'));
    const [fontsChurn] = useState(() => churnWarning('PptxEditor.Root', 'fonts'));
    useEffect(() => {
      if (!editor || !bytes) return;
      documentChurn();
      void editor.load(bytes);
    }, [editor, bytes, documentChurn]);

    useEffect(() => {
      editor?.setMode(mode);
    }, [editor, mode]);

    // A new author applies to later comments; the instance ignores the same name.
    const author = props.author;
    useEffect(() => {
      editor?.setAuthor(author);
    }, [editor, author]);

    // A `modules` change after creation has no effect: say so in development. The check is by
    // the identity of each module, so an inline array of the same modules does not warn.
    const modules = props.modules;
    useEffect(() => {
      if (editor) warnModulesChanged(createdModules.current, modules, 'PptxEditor.Root');
    }, [editor, modules]);

    // A new `fonts` value replaces the instance's fonts in place. The instance a remount just
    // created already has them, so only a change after creation reaches `setFonts`.
    const fonts = props.fonts;
    useEffect(() => {
      if (!editor || createdFonts.current === fonts) return;
      createdFonts.current = fonts;
      fontsChurn();
      // Rejects only after destroy; font failures report through `onFontError`.
      editor.setFonts(fonts).catch(() => {});
    }, [editor, fonts, fontsChurn]);

    // A new `fontFamilies` list updates the font picker in place. The instance compares by
    // content, so an inline array made on each render changes nothing.
    const fontFamilies = props.fontFamilies;
    useEffect(() => {
      if (!editor || createdFamilies.current === fontFamilies) return;
      createdFamilies.current = fontFamilies;
      editor.setFontFamilies(fontFamilies ?? []);
    }, [editor, fontFamilies]);

    // A locale change repaints the canvas once, for its placeholder prompts and frame labels,
    // which the engine translates through `tRef`. The repaint emits no `change`, so a host whose
    // `onChange` sets state does not loop.
    const painted = useRef(localeKey);
    useEffect(() => {
      if (!editor || painted.current === localeKey) return;
      painted.current = localeKey;
      repaintSurfaceOf(editor);
    }, [editor, localeKey]);

    const slideshow = useSlideshowHandle(editor);
    const dark = useResolvedDark(props.colorMode);
    const [notesOpen, setNotesOpen] = useState<boolean | undefined>(undefined);
    const notesRequest = useRef<((open: boolean) => void) | null>(null);
    const panes = useMemo<PaneState>(
      () => ({ notesOpen, setNotesOpen, notesRequest }),
      [notesOpen]
    );

    return (
      <PptxEditorContext.Provider value={editor}>
        <TranslationContext.Provider value={t}>
          <SlideshowContext.Provider value={slideshow}>
            <ChromeDarkContext.Provider value={dark}>
              <PaneStateContext.Provider value={panes}>
                <DialogProvider {...(props.popups ? { popups: props.popups } : {})}>
                  <ChromeActionHostProvider>
                    {children}
                    <SlideSorterWatch />
                    <ChromeTooltip />
                  </ChromeActionHostProvider>
                </DialogProvider>
              </PaneStateContext.Provider>
            </ChromeDarkContext.Provider>
          </SlideshowContext.Provider>
        </TranslationContext.Provider>
      </PptxEditorContext.Provider>
    );
  }
);

/** The props of `PptxEditorViewport`. @public */
export interface PptxEditorViewportProps {
  /** Appended after the `pptx-editor` class. */
  className?: string;
  /**
   * The name of the editor's landmark region, or a translation key. Default: "Presentation
   * editor" (`viewport.ariaLabel`). Give each editor its own name when a page shows more than one.
   * `false` renders no region, for a host whose own landmarks (`header`, `main`) are inside.
   */
  label?: string | false;
  /** The chrome and the content of the editor. */
  children?: ReactNode;
}

/** The region role and name of the viewport, or nothing for `label={false}`. */
function viewportLandmark(
  t: Translate,
  label: string | false | undefined
): { role?: 'region'; 'aria-label'?: string } {
  if (label === false) return {};
  return {
    role: 'region',
    'aria-label': label ? translateKeyOrText(t, label) : t('viewport.ariaLabel'),
  };
}

/**
 * The editor's scroll and styling scope: carries the `.pptx-editor` class, and `dark` with it
 * when the root's `colorMode` resolves to dark. It is a named landmark region ("Presentation
 * editor"), so a screen reader user can go to the editor from anywhere on the page. @public
 */
export function PptxEditorViewport({ className, label, children }: PptxEditorViewportProps) {
  const dark = useChromeDark();
  const { t } = useTranslation();
  return (
    <div
      className={['pptx-editor', dark ? 'dark' : '', className].filter(Boolean).join(' ')}
      {...viewportLandmark(t, label)}
    >
      <ScopedByAncestorContext.Provider value>{children}</ScopedByAncestorContext.Provider>
    </div>
  );
}

/** The props of `PptxEditorContent`. @public */
export interface PptxEditorContentProps {
  /** Appended after the `pptx-canvas` class. */
  className?: string;
}

/**
 * The slide canvas: attaches the engine surface in a layout effect. Right-click, Shift+F10, and
 * the Menu key open its context menu. @public
 */
export function PptxEditorContent({ className }: PptxEditorContentProps) {
  const editor = usePptxEditor();
  const host = useRef<HTMLDivElement>(null);
  const menu = useCanvasContextMenu();
  useLayoutEffect(() => {
    if (!editor || !host.current) return;
    editor.attach(host.current);
    return () => editor.detach();
  }, [editor]);
  return (
    <>
      <div
        ref={host}
        className={`pptx-canvas${className ? ` ${className}` : ''}`}
        data-pptx-canvas=""
        onContextMenu={menu.onContextMenu}
        onKeyDown={menu.onKeyDown}
      />
      {menu.layer}
    </>
  );
}
