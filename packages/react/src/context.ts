// Provider context and hooks. One multiplexed subscription per hook with slice memoization:
// a selector that reads the slide count does not re-render on a bold toggle (docx-editor's
// `useEditorState` contract).

import { createContext, useCallback, useContext, useMemo, useSyncExternalStore } from 'react';
import { createTranslator, type LocaleCatalog, type Translate } from '@pptx-editor.dev/i18n';
import { drawingRenderersVersion, subscribeDrawingRenderers } from '@pptx-editor.dev/core/layout';
import { NOTES_SLOTS, useNotesTarget } from './notes-target.ts';
import {
  canRenameSectionInPlace,
  chromeControlForSlot,
  chromeOfferedControl,
  subscribeSectionRenamers,
  subscribePaneListeners,
  paneListenersVersion,
  EMPTY_EDITOR_SNAPSHOT,
  editorCommandFor,
  platformShortcuts,
  runToolbarCommand,
  sameChromeCommandState,
  valueSlotOptions,
  valueSlotShown,
  type ChromeActionHost,
  type ChromeControl,
  type ChromeSlotId,
  type CommandResult,
  type EditorExecOptions,
  type ToolbarShownValue,
  type ToolbarSlotValue,
  type ToolbarValueMap,
  type EditorSnapshot,
  type PptxEditorInstance,
  type SlideshowController,
  warnOutsideRoot,
} from '@pptx-editor.dev/core/editor';

export { platformShortcuts };

/** The editor instance published by `PptxEditor.Root`. @public */
export const PptxEditorContext = createContext<PptxEditorInstance | null>(null);

/** The translator published by `PptxEditor.Root` from its `locale` prop (English by default). @public */
export const TranslationContext = createContext<Translate>(createTranslator());

/** The slide show of the editor in `PptxEditor.Root`. @public */
export interface UseSlideshowReturn {
  /** The running show, or null. `PptxEditor.Slideshow` paints it. */
  readonly controller: SlideshowController | null;
  /**
   * Start the show from the first shown slide or from the current slide. Does nothing without
   * an open document with slides. Viewing mode can present.
   */
  start(from: 'start' | 'current'): void;
  /** Close the show. */
  stop(): void;
}

/** The slide show outside `PptxEditor.Root`: also how `usePptxEditor` knows it has no root. */
const NO_SLIDESHOW: UseSlideshowReturn = {
  controller: null,
  start: () => undefined,
  stop: () => undefined,
};

/** The slide show published by `PptxEditor.Root`. @public */
export const SlideshowContext = createContext<UseSlideshowReturn>(NO_SLIDESHOW);

/** The slide show of the editor in `PptxEditor.Root`. @public */
export function useSlideshow(): UseSlideshowReturn {
  return useContext(SlideshowContext);
}

/**
 * The editor instance, or null before it mounts. Outside `PptxEditor.Root` it is always null,
 * and a development warning says so, as for the other parts. @public
 */
export function usePptxEditor(): PptxEditorInstance | null {
  const editor = useContext(PptxEditorContext);
  // The root publishes its slide show with the editor; a host that provides only an editor
  // (`PptxEditorContext.Provider`) passes one, so neither case warns.
  const outside = useContext(SlideshowContext) === NO_SLIDESHOW;
  if (!editor && outside) warnOutsideRoot('usePptxEditor');
  return editor;
}

const platformCache = new WeakMap<Translate, Translate>();

/** The translator, with shortcut text for this platform. @public */
export function useTranslation(): { t: Translate } {
  const base = useContext(TranslationContext);
  let t = platformCache.get(base);
  if (!t) {
    t = (key, params) => platformShortcuts(base(key, params));
    platformCache.set(base, t);
  }
  return { t };
}

/** Build a translator for a locale catalog. @public */
export function translatorFor(locale?: LocaleCatalog): Translate {
  return createTranslator(locale);
}

/**
 * Subscribe to a slice of editor state. Re-renders only when `isEqual` says the slice changed
 * (default `Object.is`).
 *
 * @public
 */
export function useEditorState<T>(
  selector: (snapshot: EditorSnapshot) => T,
  isEqual: (a: T, b: T) => boolean = Object.is
): T {
  const editor = usePptxEditor();
  const subscribe = useCallback(
    (notify: () => void) => (editor ? editor.subscribe(notify) : () => undefined),
    [editor]
  );
  // DOCX Editor's memo: `snapshot()` returns the same object until the state changes, so the
  // last slice is returned while the snapshot is unchanged. A selector that builds a new array
  // or object each call (`ribbonTabs(s)`) then gives `useSyncExternalStore` a stable value and
  // does not loop. A new snapshot runs the selector, and `isEqual` keeps an equal slice. The memo
  // is keyed on the selector, so a selector that reads new props never returns a stale slice.
  const select = useMemo(() => {
    let memo: { snapshot: EditorSnapshot; value: T } | null = null;
    return (snapshot: EditorSnapshot): T => {
      if (memo && memo.snapshot === snapshot) return memo.value;
      const next = selector(snapshot);
      if (memo && isEqual(memo.value, next)) {
        memo = { snapshot, value: memo.value };
        return memo.value;
      }
      memo = { snapshot, value: next };
      return next;
    };
  }, [selector, isEqual]);
  const get = useCallback(
    (): T => select(editor ? editor.snapshot() : EMPTY_EDITOR_SNAPSHOT),
    [editor, select]
  );
  return useSyncExternalStore(subscribe, get, get);
}

/** A chrome slot's state and its runner: one source of enabled state. @public */
export interface UseEditorCommandReturn {
  /**
   * Run the slot's command, with a value for value controls: the picked option's value, a
   * number of points for `font.size`, or `{ step: 1 }` / `{ step: -1 }` for its Increase and
   * Decrease Font Size buttons (`commandForSlotValue`). Returns true when the engine ran it and
   * false on a refusal, as in DOCX Editor.
   */
  readonly execute: (value?: unknown) => boolean;
  /** Whether the command can run now. */
  readonly isEnabled: boolean;
  /** Whether a toggle control shows as pressed. */
  readonly isActive: boolean;
  /**
   * Why the command cannot run: a catalog key, or the engine's own text. Show it with
   * `translateKeyOrText(t, reason)`.
   */
  readonly disabledReason: string | null;
  /** The value a value control shows (a font size, a family), when it has one. */
  readonly value?: string | number;
}

const shallow = (a: UseEditorCommandReturn | null, b: UseEditorCommandReturn | null): boolean =>
  sameChromeCommandState(a, b);

/**
 * The adapter actions of the enclosing `PptxEditor.Root` (dialogs, the link popup, find, the
 * show, a picture, the clipboard). @internal
 */
export const ChromeActionHostContext = createContext<ChromeActionHost | null>(null);

/** A subscription that never calls back, for a slot that does not follow the rename registry. */
const NO_SUBSCRIPTION = (): (() => void) => () => {};

/**
 * State and execution for one `ChromeSlotId`. A slot the adapter runs (a dialog, the link
 * popup, find, the slide show) runs through `PptxEditor.Root`, as a menu row does; outside a
 * root it is disabled with the reason. `useEditorCommand('text.link').execute(url)` sets the
 * link (`null` removes it) with or without the popup. While the speaker notes field has the
 * cursor, `text.bold`, `text.italic`, `text.underline`, and `text.strike` format the notes
 * text. `execute` takes any value, so a host can pass a
 * registry option's `value`; `useEditorValueCommand` types it. @public
 */
export function useEditorCommand(slot: ChromeSlotId): UseEditorCommandReturn {
  const editor = usePptxEditor();
  const host = useContext(ChromeActionHostContext);
  // Rename Section is on while a rail or a sorter can show the name field: follow them. Only
  // the two section slots that open the field listen.
  const section = slot === 'section.rename' || slot === 'section.add';
  const renamable = useSyncExternalStore(
    section ? subscribeSectionRenamers : NO_SUBSCRIPTION,
    () => section && canRenameSectionInPlace(editor),
    () => false
  );
  // The pane slots are on while a side panel or the host listens for pane requests.
  const panes = useSyncExternalStore(
    subscribePaneListeners,
    paneListenersVersion,
    paneListenersVersion
  );
  const select = useCallback((): UseEditorCommandReturn => {
    void panes;
    const { state, execute } = editorCommandFor(editor, slot, host, 'disabled.noRoot');
    return {
      execute: (value?: unknown) => execute(value),
      ...state,
    };
    // `renamable` changes the answer of `editorCommandFor` for the section slots.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editor, slot, host, renamable, panes]);
  const engine = useEditorState(
    select as unknown as (s: EditorSnapshot) => UseEditorCommandReturn,
    shallow as (a: UseEditorCommandReturn, b: UseEditorCommandReturn) => boolean
  );
  const notes = useNotesTarget(editor);
  const format = NOTES_SLOTS[slot];
  return useMemo(
    () =>
      notes && format
        ? {
            execute: () => notes.apply(format),
            isEnabled: !notes.readOnly && notes.canFormat,
            isActive: notes.formats[format],
            disabledReason: notes.readOnly
              ? 'disabled.viewing'
              : notes.canFormat
                ? null
                : 'disabled.selectNotesText',
          }
        : engine,
    [notes, format, engine]
  );
}

/** A typed value slot's state and runner, after DOCX Editor's `useEditorValueCommand`. @public */
export interface UseEditorValueCommandReturn<V, S = V> {
  /**
   * Run the slot's command with a value of its `ToolbarValueMap` type. The result says why the
   * engine refused it (`disabled.invalidValue`, `disabled.noEffectChange`). Pass
   * `{ historyGroup }` to make the values of one gesture (a slider drag) one undo step.
   */
  readonly execute: (value: V, options?: EditorExecOptions) => CommandResult;
  /** The value the slot shows (`ToolbarShownValue`), or null when it is mixed or unavailable. */
  readonly value: S | null;
  /**
   * The slot's fixed choices, from the registry (the Rotate, Align, line width, or Shape Effects
   * rows); none for a slot with a free value (`font.size`, `shape.fill`, `shape.styles`).
   * Without the `@pptx-editor.dev/pro/drawing` module, they leave out the effects it paints
   * (`chromeOfferedControl`), so `value` can be an effect from the file that is not in the list.
   * They keep the rows that remove an effect the selection has (`reflection.none`).
   */
  readonly options: readonly V[];
  /** Whether the command can run now. */
  readonly isEnabled: boolean;
  /** Why the command cannot run: a catalog key, or the engine's own text. */
  readonly disabledReason: string | null;
}

/**
 * The control a slot shows in this build (`chromeOfferedControl`), read again on each editor
 * change and after each drawing renderer install. Without the drawing module, a selection
 * change adds or drops the rows that remove an effect the selection has. @internal
 */
export function useOfferedControl(
  slot: ChromeSlotId,
  control: ChromeControl | null | undefined
): ChromeControl | undefined {
  const editor = usePptxEditor();
  const renderers = useSyncExternalStore(
    subscribeDrawingRenderers,
    drawingRenderersVersion,
    drawingRenderersVersion
  );
  // A new selector after an install, so the slice is read again with an unchanged snapshot.
  const select = useCallback(
    () => (void renderers, control ? chromeOfferedControl(slot, control, editor) : undefined),
    [slot, control, editor, renderers]
  );
  // The same choices give the same object, so the default `Object.is` keeps the slice.
  return useEditorState(select as unknown as (s: EditorSnapshot) => ChromeControl | undefined);
}

/**
 * State and execution for one typed value slot: `execute` takes the slot's `ToolbarValueMap`
 * type, so a typo in a value does not compile, and returns the engine's result. @public
 */
export function useEditorValueCommand<K extends keyof ToolbarValueMap>(
  slot: K
): UseEditorValueCommandReturn<ToolbarValueMap[K], ToolbarShownValue<K>> {
  const editor = usePptxEditor();
  const cmd = useEditorCommand(slot);
  // The options follow the installed drawing renderers and, without them, the effects of the
  // selection: `offered` changes with either, so the options run again.
  const offered = useOfferedControl(slot, chromeControlForSlot(slot));
  return useMemo(
    () => ({
      execute: (value: ToolbarValueMap[K], options?: EditorExecOptions): CommandResult =>
        runToolbarCommand(editor, slot, value as ToolbarSlotValue<K>, options),
      value: valueSlotShown(slot, cmd.value),
      options: (void offered, valueSlotOptions(slot, editor)),
      isEnabled: cmd.isEnabled,
      disabledReason: cmd.disabledReason,
    }),
    [editor, slot, cmd, offered]
  );
}
