// Injection keys and composables: the Vue twin of the React adapter's context and hooks
// (packages/react/src/context.ts). The state each composable derives comes from the shared
// adapter support in core (`editorCommandFor`, `EMPTY_EDITOR_SNAPSHOT`,
// `platformShortcuts`), so React and Vue report the same thing. After DOCX Editor's Vue
// adapter (docx-editor/packages/vue/src/editor/context.ts, useDocxEditorRoot.ts,
// useEditorState.ts, useEditorCommand.ts): the root holds ONE editor subscription and
// publishes a change tick, and every composable reads the tick, so a toolbar of fifty controls
// costs one subscription, not fifty. Every composable annotates a named `Use<Name>Return`
// type, so core's internal types do not leak into the API snapshot.

import {
  computed,
  getCurrentScope,
  inject,
  onScopeDispose,
  provide,
  shallowRef,
  toValue,
  watch,
  type ComputedRef,
  type InjectionKey,
  type MaybeRefOrGetter,
  type ShallowRef,
} from 'vue';
import { createTranslator, type Translate } from '@pptx-editor.dev/i18n';
import { drawingRenderersVersion, subscribeDrawingRenderers } from '@pptx-editor.dev/core/layout';
import {
  EMPTY_EDITOR_SNAPSHOT,
  chromeCommandState,
  editorCommandFor,
  platformShortcuts,
  runToolbarCommand,
  sameBulletFormat,
  sameChromeCommandState,
  sameFontFormat,
  sameParagraphFormat,
  valueSlotOptions,
  valueSlotShown,
  type BulletFormat,
  type BulletFormatRead,
  type ChromeActionHost,
  type ChromeCommandState,
  type ChromeSlotId,
  type CommandResult,
  type EditorExecOptions,
  type ToolbarShownValue,
  type ToolbarSlotValue,
  type ToolbarValueMap,
  type EditorSnapshot,
  type FontFormatCommand,
  type FontFormatRead,
  type ParagraphFormatCommand,
  type ParagraphFormatRead,
  type SlideLayoutSummary,
  type PptxEditorInstance,
  warnOutsideRoot,
} from '@pptx-editor.dev/core/editor';
import { NOTES_SLOTS, useNotesTarget } from './notes-target.ts';

/** The editor instance published by `PptxEditorRoot`. @public */
export const pptxEditorKey: InjectionKey<ShallowRef<PptxEditorInstance | null>> =
  Symbol('pptxEditor');

/** The translator published by `PptxEditorRoot` from its `locale` prop. @public */
export const pptxTranslationKey: InjectionKey<ShallowRef<Translate>> = Symbol('pptxTranslation');

/** A counter the root bumps on every editor change: one subscription for all composables. */
const pptxEditorTickKey: InjectionKey<ShallowRef<number>> = Symbol('pptxEditorTick');

const nullEditor = shallowRef<PptxEditorInstance | null>(null);
const englishTranslator = shallowRef<Translate>(createTranslator());

let activeSubscriptions = 0;

/** Live editor subscriptions the adapter holds. Tests check cleanup with it. @internal */
export function editorStateActiveSubscriptionCount(): number {
  return activeSubscriptions;
}

/** Subscribe to `editor`, counting the subscription. Returns the unsubscribe function. */
function subscribe(editor: PptxEditorInstance, listener: () => void): () => void {
  activeSubscriptions += 1;
  const off = editor.subscribe(listener);
  let live = true;
  return () => {
    if (!live) return;
    live = false;
    off();
    activeSubscriptions -= 1;
  };
}

/**
 * Publish `editor` and one shared change tick to the components below. `PptxEditorRoot`
 * calls it; a host that owns the instance can call it in its own `setup`. @internal
 */
export function providePptxEditor(editor: ShallowRef<PptxEditorInstance | null>): void {
  provide(pptxEditorKey, editor);
  const tick = shallowRef(0);
  provide(pptxEditorTickKey, tick);
  let off: (() => void) | null = null;
  watch(
    editor,
    (next) => {
      off?.();
      off = next ? subscribe(next, () => (tick.value += 1)) : null;
      tick.value += 1;
    },
    { immediate: true }
  );
  onScopeDispose(() => {
    off?.();
    off = null;
  });
}

/** What `usePptxEditor` returns: the editor, or null before `PptxEditorRoot` mounts it. @public */
export type UsePptxEditorReturn = Readonly<ShallowRef<PptxEditorInstance | null>>;

/**
 * The editor instance, or null before it mounts. Outside `PptxEditorRoot` it is always null,
 * and a development warning says so, as for the other parts. @public
 */
export function usePptxEditor(): UsePptxEditorReturn {
  const editor = inject(pptxEditorKey, nullEditor);
  if (editor === nullEditor) warnOutsideRoot('usePptxEditor');
  return editor;
}

/** What `useTranslation` returns. @public */
export interface UseTranslationReturn {
  /** Translate a key, with shortcut text for this platform. Reactive to the `locale` prop. */
  readonly t: Translate;
}

/** The translator, with shortcut text for this platform. @public */
export function useTranslation(): UseTranslationReturn {
  const base = inject(pptxTranslationKey, englishTranslator);
  return { t: (key, params) => platformShortcuts(base.value(key, params)) };
}

/** What `useEditorState` returns: the selected slice, updated when it changes. @public */
export type UseEditorStateReturn<T> = Readonly<ShallowRef<T>>;

/** `useEditorState`, plus a source whose change re-reads the slice (the command's slot). */
function useSlice<T>(
  selector: (snapshot: EditorSnapshot) => T,
  isEqual: (a: T, b: T) => boolean,
  deps?: () => unknown
): UseEditorStateReturn<T> {
  const editorRef = usePptxEditor();
  const tick = inject(pptxEditorTickKey, null);
  // DOCX Editor's memo: `snapshot()` returns the same object until the state changes, so a
  // change with an unchanged snapshot skips the selector. A selector that builds a new array on
  // each call (`ribbonTabs(s)`) then updates the ref only when the state changed. A new editor
  // or a changed `deps` source always re-reads the slice.
  let seen: EditorSnapshot | null = null;
  const current = (): EditorSnapshot => {
    const editor = editorRef.value;
    return editor ? editor.snapshot() : EMPTY_EDITOR_SNAPSHOT;
  };
  const read = (): T => {
    seen = current();
    return selector(seen);
  };
  const out = shallowRef(read()) as ShallowRef<T>;
  const update = (): void => {
    const next = read();
    if (!isEqual(out.value, next)) out.value = next;
  };
  const onChange = (): void => {
    if (current() !== seen) update();
  };
  if (!getCurrentScope()) return out;
  if (deps) watch(deps, update);
  if (tick) {
    watch(editorRef, update);
    watch(tick, onChange);
    return out;
  }
  // No root above (a host that provides only the editor): subscribe directly.
  let off: (() => void) | null = null;
  watch(
    editorRef,
    (editor) => {
      off?.();
      off = editor ? subscribe(editor, onChange) : null;
      update();
    },
    { immediate: true }
  );
  onScopeDispose(() => off?.());
  return out;
}

/**
 * Subscribe to a slice of editor state. The ref updates only when `isEqual` says the slice
 * changed (default `Object.is`), as React's `useEditorState`. @public
 */
export function useEditorState<T>(
  selector: (snapshot: EditorSnapshot) => T,
  isEqual: (a: T, b: T) => boolean = Object.is
): UseEditorStateReturn<T> {
  return useSlice(selector, isEqual);
}

/** What `useSlideLayouts` returns: the deck's slide layouts. @public */
export type UseSlideLayoutsReturn = UseEditorStateReturn<readonly SlideLayoutSummary[]>;

const NO_LAYOUTS: readonly SlideLayoutSummary[] = [];

/**
 * The deck's slide layouts (`EditorSnapshot.layouts`), in the order the galleries list them, as
 * React's `useSlideLayouts`. For a host's own layout picker. @public
 */
export function useSlideLayouts(): UseSlideLayoutsReturn {
  return useSlice((s) => s.layouts ?? NO_LAYOUTS, Object.is);
}

/** What `useEditorCommand` returns: one source of enabled state, and the runner. @public */
export interface UseEditorCommandReturn {
  /**
   * Run the slot's command, with a value for value controls: the picked option's value, a
   * number of points for `font.size`, or `{ step: 1 }` / `{ step: -1 }` for its Increase and
   * Decrease Font Size buttons (`commandForSlotValue`). Returns true when the engine ran it and
   * false on a refusal, as in DOCX Editor.
   */
  readonly execute: (value?: unknown) => boolean;
  /** Whether the command can run now. */
  readonly isEnabled: ComputedRef<boolean>;
  /** Whether a toggle control shows as pressed. */
  readonly isActive: ComputedRef<boolean>;
  /** A translation key, or the engine's own text, when the control is disabled. */
  readonly disabledReason: ComputedRef<string | null>;
  /** The value a value control shows (a font size, a family), when it has one. */
  readonly value: ComputedRef<string | number | undefined>;
}

/** The root's `ChromeActionHost`, provided with its dialog host. @internal */
export const chromeActionHostKey: InjectionKey<ChromeActionHost> = Symbol('pptxChromeActionHost');

/**
 * State and execution for one `ChromeSlotId`, as React's `useEditorCommand`. The slot may be
 * a ref or a getter, as in DOCX Editor's Vue adapter: the state follows it. @public
 */
export function useEditorCommand(slot: MaybeRefOrGetter<ChromeSlotId>): UseEditorCommandReturn {
  const editorRef = usePptxEditor();
  // The root's action host opens its dialogs, the find bar, the link popup, and the slide
  // show, and picks pictures, as React's. Outside a root there is no host: those slots show as
  // disabled with the reason.
  const host = inject(chromeActionHostKey, null);
  const read = (): ChromeCommandState =>
    editorCommandFor(editorRef.value, toValue(slot), host, 'disabled.noRoot').state;
  // The deps read the slot and the root's `popups` (through the host), so a host that changes
  // `popups` at runtime turns the control on or off at once, not at the next editor change.
  const state = useSlice<ChromeCommandState>(read, sameChromeCommandState, () => [
    toValue(slot),
    read(),
  ]);
  // While a speaker notes field has focus, bold, italic, underline, and strikethrough format
  // the notes text, as in React (`NOTES_SLOTS`).
  const notes = useNotesTarget(editorRef);
  const notesFormat = computed(() => {
    const target = notes.value;
    const format = NOTES_SLOTS[toValue(slot)];
    return target && format ? { target, format } : null;
  });
  return {
    execute: (value?: unknown) => {
      const n = notesFormat.value;
      if (n) return n.target.apply(n.format);
      return editorCommandFor(editorRef.value, toValue(slot), host, 'disabled.noRoot').execute(
        value
      );
    },
    isEnabled: computed(() => {
      const n = notesFormat.value;
      return n ? !n.target.readOnly && n.target.canFormat : state.value.isEnabled;
    }),
    isActive: computed(() => {
      const n = notesFormat.value;
      return n ? n.target.formats[n.format] : state.value.isActive;
    }),
    disabledReason: computed(() => {
      const n = notesFormat.value;
      if (!n) return state.value.disabledReason;
      if (n.target.readOnly) return 'disabled.viewing';
      return n.target.canFormat ? null : 'disabled.selectNotesText';
    }),
    value: computed(() => state.value.value),
  };
}

/** What `useParagraphFormat` returns. @public */
export interface UseParagraphFormatReturn {
  /** The selected paragraphs' formatting, or null when no text is selected. */
  readonly format: UseEditorStateReturn<ParagraphFormatRead | null>;
  /** Whether the engine can write paragraph formatting now (text or a shape with text selected, editable). */
  readonly isEnabled: ComputedRef<boolean>;
  /** Write the given fields as ONE undo step. Returns whether the engine accepted them. */
  readonly apply: (update: ParagraphFormatCommand) => boolean;
}

/**
 * The selection's paragraph formatting, and the command to change it, as React's
 * `useParagraphFormat` and DOCX Editor's Vue composable of the same name. @public
 */
export function useParagraphFormat(): UseParagraphFormatReturn {
  const editorRef = usePptxEditor();
  const format = useEditorState((s) => s.format?.paragraph ?? null, sameParagraphFormat);
  // The engine's answer, not the dialog slot's: Vue has no dialog host, and hiding a dialog
  // never turns off the operation.
  const enabled = useSlice<boolean>(
    () => chromeCommandState(editorRef.value, 'paragraph.dialog').isEnabled,
    Object.is
  );
  return {
    format,
    isEnabled: computed(() => enabled.value),
    apply: (update) => {
      const editor = editorRef.value;
      return editor !== null && editor.exec({ ...update, type: 'setParagraphFormat' }).ok;
    },
  };
}

/** What `useFontFormat` returns. @public */
export interface UseFontFormatReturn {
  /** The font of the selected text, or of every run of the selected shapes; null with none. */
  readonly format: UseEditorStateReturn<FontFormatRead | null>;
  /** Whether the engine can write font formatting now (text or a shape with text selected, editable). */
  readonly isEnabled: ComputedRef<boolean>;
  /** Write the given fields as ONE undo step. Returns whether the engine accepted them. */
  readonly apply: (update: FontFormatCommand) => boolean;
}

/**
 * The font of the selected text (or of every run of the selected shapes), and the command to
 * change it, as React's `useFontFormat`. @public
 */
export function useFontFormat(): UseFontFormatReturn {
  const editorRef = usePptxEditor();
  const format = useEditorState((s) => s.format?.font ?? null, sameFontFormat);
  // The engine's answer, not the dialog slot's: Vue has no dialog host, and hiding a dialog
  // never turns off the operation.
  const enabled = useSlice<boolean>(
    () => chromeCommandState(editorRef.value, 'text.dialog').isEnabled,
    Object.is
  );
  return {
    format,
    isEnabled: computed(() => enabled.value),
    apply: (update) => {
      const editor = editorRef.value;
      return editor !== null && editor.exec({ ...update, type: 'setFontFormat' }).ok;
    },
  };
}

/** What `useBulletFormat` returns. @public */
export interface UseBulletFormatReturn {
  /**
   * The bullet format of the selected paragraphs: null when they have no list, undefined when
   * they differ or nothing with text is selected.
   */
  readonly format: UseEditorStateReturn<BulletFormatRead | null | undefined>;
  /** Whether the engine can write a bullet format now (text or a shape with text selected, editable). */
  readonly isEnabled: ComputedRef<boolean>;
  /**
   * Write a bullet format (null turns the list off) as ONE undo step. An omitted field keeps
   * each paragraph's value; with no `kind`, the size and color change and each paragraph keeps
   * its own bullet or number. Returns whether the engine accepted it.
   */
  readonly apply: (format: BulletFormat | null) => boolean;
}

/** The bullet format of the selected paragraphs, and the command to change it, as React's `useBulletFormat`. @public */
export function useBulletFormat(): UseBulletFormatReturn {
  const editorRef = usePptxEditor();
  const format = useEditorState((s) => s.format?.bulletFormat, sameBulletFormat);
  const { isEnabled } = useEditorCommand('list.bulletStyle');
  return {
    format,
    isEnabled,
    apply: (next) => {
      const editor = editorRef.value;
      return editor !== null && editor.exec({ type: 'setBulletFormat', format: next }).ok;
    },
  };
}

/** What `useEditorValueCommand` returns, as React's `UseEditorValueCommandReturn`. @public */
export interface UseEditorValueCommandReturn<V, S = V> {
  /**
   * Run the slot's command with a value of its `ToolbarValueMap` type. The result says why the
   * engine refused it (`disabled.invalidValue`, `disabled.noEffectChange`). Pass
   * `{ historyGroup }` to make the values of one gesture (a slider drag) one undo step.
   */
  readonly execute: (value: V, options?: EditorExecOptions) => CommandResult;
  /** The value the slot shows (`ToolbarShownValue`), or null when it is mixed or unavailable. */
  readonly value: ComputedRef<S | null>;
  /**
   * The slot's fixed choices, from the registry (the Rotate, Align, line width, or Shape Effects
   * rows); none for a slot with a free value (`font.size`, `shape.fill`, `shape.styles`).
   */
  readonly options: ComputedRef<readonly V[]>;
  /** Whether the command can run now. */
  readonly isEnabled: ComputedRef<boolean>;
  /** Why the command cannot run: a catalog key, or the engine's own text. */
  readonly disabledReason: ComputedRef<string | null>;
}

/**
 * State and execution for one typed value slot, as React's `useEditorValueCommand`: `execute`
 * takes the slot's `ToolbarValueMap` type and returns the engine's result. @public
 */
export function useEditorValueCommand<K extends keyof ToolbarValueMap>(
  slot: MaybeRefOrGetter<K>
): UseEditorValueCommandReturn<ToolbarValueMap[K], ToolbarShownValue<K>> {
  const editorRef = usePptxEditor();
  const cmd = useEditorCommand(slot);
  // The options follow the installed drawing renderers, as in React: an install offers the Pro
  // choices.
  const renderers = shallowRef(drawingRenderersVersion());
  const stop = subscribeDrawingRenderers(() => (renderers.value = drawingRenderersVersion()));
  if (getCurrentScope()) onScopeDispose(stop);
  return {
    execute: (value: ToolbarValueMap[K], options?: EditorExecOptions): CommandResult =>
      runToolbarCommand(editorRef.value, toValue(slot), value as ToolbarSlotValue<K>, options),
    value: computed(() => valueSlotShown(toValue(slot), cmd.value.value)),
    options: computed(() => (void renderers.value, valueSlotOptions(toValue(slot)))),
    isEnabled: cmd.isEnabled,
    disabledReason: cmd.disabledReason,
  };
}
