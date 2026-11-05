// The Vue twins of the React adapter's `headless-hooks.tsx`: `useChromeTranslate` and the
// compounds' `t` prop, `useToolbarContext`, `useToolbarLabel`, `useToolbarLabelFor`,
// `useHistoryGroup`, `useEditorCaret`, and `useDocumentSearch`, with DOCX Editor's names. Each
// returns refs where React returns values, and works outside a root: no matches, no caret,
// and actions that do nothing.

import {
  computed,
  getCurrentScope,
  inject,
  onScopeDispose,
  provide,
  shallowRef,
  watch,
  type ComputedRef,
  type InjectionKey,
  type ShallowRef,
} from 'vue';
import {
  IDLE_DOCUMENT_SEARCH,
  bindHistoryGroup,
  documentSearchOf,
  editorCaretOf,
  type ChromeTranslate,
  type DocumentSearch,
  type DocumentSearchFindOptions,
  type DocumentSearchState,
  type EditorCaret,
  type EditorExecOptions,
  type FindMatch,
  type HistoryGroupBinding,
  type HistoryGroupBindingOptions,
} from '@pptx-editor.dev/core/editor';
import type { Translate, TranslationKey } from '@pptx-editor.dev/i18n';
import { pptxTranslationKey, usePptxEditor, useTranslation } from './composables.ts';

// --- Translation ----------------------------------------------------------------------------

/** What {@link useChromeTranslate} returns: the translation function. @public */
export type UseChromeTranslateReturn = ChromeTranslate;

/**
 * The chrome's translation function, as React's `useChromeTranslate`: the `t` of the nearest
 * compound that takes one, else the root's `locale`. `overrides` wins for the keys it holds.
 * @public
 */
export function useChromeTranslate(
  overrides?: ReadonlyMap<string, string>
): UseChromeTranslateReturn {
  const { t } = useTranslation();
  return (key, params) => overrides?.get(key) ?? t(key as TranslationKey, params);
}

/** What {@link useToolbarContext} returns, as React's `ToolbarContextValue`. @public */
export interface UseToolbarContextReturn {
  /** The host's `t` on the nearest compound, or undefined. */
  readonly t: Readonly<ShallowRef<ChromeTranslate | undefined>>;
}

const noHostTranslate = shallowRef<ChromeTranslate | undefined>(undefined);

/** The nearest compound's `t`. @internal */
export const toolbarContextKey: InjectionKey<Readonly<ShallowRef<ChromeTranslate | undefined>>> =
  Symbol('pptxToolbarContext');

/** The nearest compound's published value, as React's `useToolbarContext`. @public */
export function useToolbarContext(): UseToolbarContextReturn {
  return { t: inject(toolbarContextKey, noHostTranslate) };
}

/** What {@link useToolbarLabelFor} and {@link useToolbarLabel} return. @public */
export type UseToolbarLabelForReturn = (key: string) => string;

/** What {@link useToolbarLabel} returns. @public */
export type UseToolbarLabelReturn = UseToolbarLabelForReturn;

/** The label for a key: the host's translation, else the locale catalog. @public */
export function useToolbarLabelFor(t: ChromeTranslate | undefined): UseToolbarLabelForReturn {
  const { t: catalog } = useTranslation();
  return (key) => t?.(key) ?? catalog(key as TranslationKey);
}

/** The label for a key: the nearest compound's `t`, else the locale catalog. @public */
export function useToolbarLabel(): UseToolbarLabelReturn {
  const host = useToolbarContext().t;
  const { t: catalog } = useTranslation();
  return (key) => host.value?.(key) ?? catalog(key as TranslationKey);
}

/** The `t` prop of a compound, for its runtime props. @internal */
export const chromeTranslateProp = {
  /**
   * Your translation function for this compound and its parts, as DOCX Editor's `t`. Without
   * it, the root's `locale` gives the labels.
   */
  t: { type: Function as unknown as () => ChromeTranslate, default: undefined },
} as const;

/**
 * Give a compound DOCX Editor's `t` prop: its parts read the host's resolver, and the
 * returned `t` is the compound's own (a component does not see what it provides). @internal
 */
export function useChromeTranslateScope(host: () => ChromeTranslate | undefined): {
  readonly t: Translate;
} {
  const parent = inject(pptxTranslationKey, null);
  const { t: inherited } = useTranslation();
  const scoped = computed<Translate>(
    () => (host() as Translate | undefined) ?? parent?.value ?? inherited
  );
  provide(pptxTranslationKey, scoped as unknown as ShallowRef<Translate>);
  const parentHost = inject(toolbarContextKey, noHostTranslate);
  provide(
    toolbarContextKey,
    computed(() => host() ?? parentHost.value) as unknown as ShallowRef<ChromeTranslate | undefined>
  );
  return {
    t: (key, params) => {
      const own = host();
      return own ? own(key, params) : inherited(key, params);
    },
  };
}

// --- History groups -------------------------------------------------------------------------

/** What {@link useHistoryGroup} returns, as React's `UseHistoryGroupReturn`. @public */
export interface UseHistoryGroupReturn {
  /** Put it on the control as a function ref: the slider, the stepper, or the color input. */
  readonly ref: (element: unknown) => void;
  /** The options for each command of the gesture: `{ historyGroup }`. */
  readonly options: () => EditorExecOptions;
}

/**
 * One undo step for one gesture on your own control, as React's `useHistoryGroup`. Bind the
 * control with `:ref="group.ref"`, and pass `group.options()` to each command of the gesture.
 * @public
 */
export function useHistoryGroup({ kind }: HistoryGroupBindingOptions): UseHistoryGroupReturn {
  const editor = usePptxEditor();
  const element = shallowRef<HTMLElement | null>(null);
  let binding: HistoryGroupBinding | null = null;
  if (getCurrentScope()) {
    const stop = watch(
      [editor, element],
      ([instance, el], _old, onCleanup) => {
        if (!instance || !el) return;
        const current = bindHistoryGroup(instance, el, { kind });
        binding = current;
        onCleanup(() => {
          current.dispose();
          binding = null;
        });
      },
      { immediate: true, flush: 'sync' }
    );
    onScopeDispose(stop);
  }
  return {
    ref: (node) => {
      const el = (node as { $el?: unknown } | null)?.$el ?? node;
      // `nodeType`, not `instanceof HTMLElement`: an element in an iframe has another realm.
      element.value = (el as Node | null)?.nodeType === 1 ? (el as HTMLElement) : null;
    },
    options() {
      if (!binding) throw new Error('history group control is not mounted');
      return binding.options();
    },
  };
}

// --- Caret ----------------------------------------------------------------------------------

/** What {@link useEditorCaret} returns: the caret, or null. @public */
export type UseEditorCaretReturn = Readonly<ShallowRef<EditorCaret | null>>;

/**
 * The caret, as React's `useEditorCaret`: the text box, paragraph, and offset, or null when no
 * text is being edited. The ref changes only when the caret moves. @public
 */
export function useEditorCaret(): UseEditorCaretReturn {
  const editor = usePptxEditor();
  const caret = shallowRef<EditorCaret | null>(null);
  if (!getCurrentScope()) return caret;
  const stop = watch(
    editor,
    (instance, _old, onCleanup) => {
      const read = (): void => {
        caret.value = instance ? editorCaretOf(instance.snapshot(), caret.value) : null;
      };
      read();
      if (instance) onCleanup(instance.subscribe(read));
    },
    { immediate: true }
  );
  onScopeDispose(stop);
  return caret;
}

// --- Search ---------------------------------------------------------------------------------

/** What {@link useDocumentSearch} returns: the session's state as refs, and its actions. @public */
export interface UseDocumentSearchReturn {
  readonly query: ComputedRef<string>;
  readonly matchCase: ComputedRef<boolean>;
  readonly wholeWord: ComputedRef<boolean>;
  readonly matches: ComputedRef<readonly FindMatch[]>;
  readonly truncated: ComputedRef<boolean>;
  readonly activeIndex: ComputedRef<number>;
  readonly activeMatch: ComputedRef<FindMatch | null>;
  readonly isPending: ComputedRef<boolean>;
  readonly setQuery: (query: string) => void;
  readonly find: (query: string, options?: DocumentSearchFindOptions) => readonly FindMatch[];
  readonly setMatchCase: (value: boolean) => void;
  readonly setWholeWord: (value: boolean) => void;
  readonly goTo: (index: number) => boolean;
  readonly next: () => boolean;
  readonly previous: () => boolean;
  readonly clear: () => void;
}

/**
 * The editor's find session with no UI, for your own find bar, as React's
 * `useDocumentSearch`. Every composable on one editor shares the session. @public
 */
export function useDocumentSearch(): UseDocumentSearchReturn {
  const editor = usePptxEditor();
  const state = shallowRef<DocumentSearchState>(IDLE_DOCUMENT_SEARCH);
  const search = computed<DocumentSearch | null>(() =>
    editor.value ? documentSearchOf(editor.value) : null
  );
  if (getCurrentScope()) {
    const stop = watch(
      search,
      (session, _old, onCleanup) => {
        state.value = session?.getState() ?? IDLE_DOCUMENT_SEARCH;
        if (session) onCleanup(session.subscribe(() => void (state.value = session.getState())));
      },
      { immediate: true }
    );
    onScopeDispose(stop);
  }
  const field = <K extends keyof DocumentSearchState>(key: K) => computed(() => state.value[key]);
  return {
    query: field('query'),
    matchCase: field('matchCase'),
    wholeWord: field('wholeWord'),
    matches: field('matches'),
    truncated: field('truncated'),
    activeIndex: field('activeIndex'),
    activeMatch: field('activeMatch'),
    isPending: field('isPending'),
    setQuery: (query) => search.value?.setQuery(query),
    find: (query, options) => search.value?.find(query, options) ?? IDLE_DOCUMENT_SEARCH.matches,
    setMatchCase: (value) => search.value?.setMatchCase(value),
    setWholeWord: (value) => search.value?.setWholeWord(value),
    goTo: (index) => search.value?.goTo(index) ?? false,
    next: () => search.value?.next() ?? false,
    previous: () => search.value?.previous() ?? false,
    clear: () => search.value?.clear(),
  };
}
