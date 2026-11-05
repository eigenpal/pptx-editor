/** @jsxImportSource react */
// The hooks for a host's own chrome that keep DOCX Editor's names and shapes:
// `useChromeTranslate` and the compounds' `t` prop, `useToolbarContext`, `useToolbarLabel`,
// `useToolbarLabelFor`, `useHistoryGroup`, `useEditorCaret`, and `useDocumentSearch`. The
// behavior is in core (`bind-history-group.ts`, `document-search.ts`, `editorCaretOf`); each
// hook is a subscription and a call. The Vue twins are in `headless-hooks.ts` of the Vue
// adapter.

import {
  createContext,
  createElement,
  useCallback,
  useContext,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
  type ComponentType,
  type ReactNode,
} from 'react';
import {
  IDLE_DOCUMENT_SEARCH,
  bindHistoryGroup,
  documentSearchOf,
  editorCaretOf,
  type ChromeTranslate,
  type DocumentSearchFindOptions,
  type DocumentSearchState,
  type EditorCaret,
  type EditorExecOptions,
  type FindMatch,
  type HistoryGroupBinding,
  type HistoryGroupBindingOptions,
} from '@pptx-editor.dev/core/editor';
import type { Translate, TranslationKey } from '@pptx-editor.dev/i18n';
import { TranslationContext, useTranslation, usePptxEditor } from './context.ts';

// --- Translation ----------------------------------------------------------------------------

/**
 * The chrome's translation function, as DOCX Editor's `useChromeTranslate`: the `t` of the
 * nearest compound that takes one, else the root's `locale`. `overrides` wins for the keys it
 * holds. Use it for the labels of your own parts, so they follow the same language.
 *
 * @example
 * ```tsx
 * const t = useChromeTranslate();
 * <PptxEditor.Toolbar.Action label={t('toolbar.save')} onSelect={save} />
 * ```
 *
 * @public
 */
export function useChromeTranslate(overrides?: ReadonlyMap<string, string>): ChromeTranslate {
  const { t } = useTranslation();
  return useCallback(
    (key: string, params?: Readonly<Record<string, string | number>>) =>
      overrides?.get(key) ?? t(key as TranslationKey, params),
    [overrides, t]
  );
}

/** What a compound's `t` prop publishes to its parts, as DOCX Editor's toolbar context. @public */
export interface ToolbarContextValue {
  /** The host's `t` on the nearest compound, or undefined. */
  readonly t: ChromeTranslate | undefined;
}

/** The nearest compound's `t`. @public */
export const ToolbarContext = createContext<ToolbarContextValue>({ t: undefined });

/** The nearest compound's published value, as DOCX Editor's `useToolbarContext`. @public */
export function useToolbarContext(): ToolbarContextValue {
  return useContext(ToolbarContext);
}

/**
 * The label for a key, given a host resolver: the host's translation, else the locale
 * catalog, as DOCX Editor's `useToolbarLabelFor`. @public
 */
export function useToolbarLabelFor(t: ChromeTranslate | undefined): (key: string) => string {
  const { t: catalog } = useTranslation();
  return useCallback((key: string) => t?.(key) ?? catalog(key as TranslationKey), [t, catalog]);
}

/** The label for a key: the nearest compound's `t`, else the locale catalog. @public */
export function useToolbarLabel(): (key: string) => string {
  return useToolbarLabelFor(useContext(ToolbarContext).t);
}

/**
 * Give a compound DOCX Editor's `t` prop: with `t`, every label of the compound and its parts
 * comes from the host's resolver. @internal
 */
export function withChromeTranslate<P extends { readonly t?: ChromeTranslate | undefined }>(
  Component: ComponentType<Omit<P, 't'>>,
  name: string
): (props: P) => ReactNode {
  function Translated(props: P): ReactNode {
    const { t, ...rest } = props;
    // The providers always render, so a change of `t` between set and unset keeps the
    // compound mounted, with its state and focus. Without `t`, they pass on what is inherited.
    const inheritedHost = useContext(ToolbarContext).t;
    const inheritedTranslate = useContext(TranslationContext);
    const host = t ?? inheritedHost;
    const value = useMemo(() => ({ t: host }), [host]);
    const translate = (t as Translate | undefined) ?? inheritedTranslate;
    return (
      <ToolbarContext.Provider value={value}>
        <TranslationContext.Provider value={translate}>
          {createElement(Component, rest as Omit<P, 't'>)}
        </TranslationContext.Provider>
      </ToolbarContext.Provider>
    );
  }
  Object.defineProperty(Translated, 'name', { value: name });
  return Translated;
}

// --- History groups -------------------------------------------------------------------------

/** What {@link useHistoryGroup} returns. @public */
export interface UseHistoryGroupReturn {
  /** Put it on the control: the slider, the stepper, or the color input. */
  readonly ref: (element: HTMLElement | null) => void;
  /** The options for each command of the gesture: `{ historyGroup }`. */
  readonly options: () => EditorExecOptions;
}

/**
 * One undo step for one gesture on your own control, as DOCX Editor's `useHistoryGroup`. Bind
 * the control with `ref`, and pass `options()` to each command that the gesture runs. A drag,
 * a held key, or a color pick is then one step of `editor.exec({ type: 'undo' })`.
 *
 * @example
 * ```tsx
 * const group = useHistoryGroup({ kind: 'range' });
 * const size = useEditorValueCommand('font.size');
 * <input type="range" ref={group.ref} onChange={(e) => size.execute(+e.target.value, group.options())} />
 * ```
 *
 * @public
 */
export function useHistoryGroup({ kind }: HistoryGroupBindingOptions): UseHistoryGroupReturn {
  const editor = usePptxEditor();
  const [element, ref] = useState<HTMLElement | null>(null);
  const binding = useRef<HistoryGroupBinding | null>(null);
  useLayoutEffect(() => {
    if (!editor || !element) return;
    const current = bindHistoryGroup(editor, element, { kind });
    binding.current = current;
    return () => {
      current.dispose();
      binding.current = null;
    };
  }, [editor, element, kind]);
  const options = useCallback((): EditorExecOptions => {
    if (!binding.current) throw new Error('history group control is not mounted');
    return binding.current.options();
  }, []);
  return useMemo(() => ({ ref, options }), [options]);
}

// --- Caret ----------------------------------------------------------------------------------

/**
 * The caret, as DOCX Editor's `useEditorCaret`: the text box, paragraph, and offset, or null
 * when no text is being edited. Compared by value, so a consumer re-renders only when the
 * caret moves.
 *
 * @public
 */
export function useEditorCaret(): EditorCaret | null {
  const editor = usePptxEditor();
  const cached = useRef<EditorCaret | null>(null);
  const subscribe = useCallback(
    (notify: () => void) => (editor ? editor.subscribe(notify) : () => undefined),
    [editor]
  );
  const read = useCallback((): EditorCaret | null => {
    cached.current = editor ? editorCaretOf(editor.snapshot(), cached.current) : null;
    return cached.current;
  }, [editor]);
  return useSyncExternalStore(subscribe, read, () => null);
}

// --- Search ---------------------------------------------------------------------------------

/** What {@link useDocumentSearch} returns, as DOCX Editor's `UseDocumentSearchResult`. @public */
export interface UseDocumentSearchResult extends DocumentSearchState {
  /** Change the text to find. The matches update at once. */
  readonly setQuery: (query: string) => void;
  /** Search now and return the matches. `selectFirst` selects the first match. */
  readonly find: (query: string, options?: DocumentSearchFindOptions) => readonly FindMatch[];
  readonly setMatchCase: (value: boolean) => void;
  readonly setWholeWord: (value: boolean) => void;
  /** Select the match at `index` and show its slide. False for an index with no match. */
  readonly goTo: (index: number) => boolean;
  /** Select the next match, after the last one the first. */
  readonly next: () => boolean;
  /** Select the previous match, before the first one the last. */
  readonly previous: () => boolean;
  /** Empty the query and drop the matches. The selection stays. */
  readonly clear: () => void;
}

const noSubscription = () => () => undefined;
const idleState = () => IDLE_DOCUMENT_SEARCH;

/**
 * The editor's find session with no UI, for your own find bar, as DOCX Editor's
 * `useDocumentSearch`. Every hook on one editor shares the session. Outside a root, it has no
 * matches and every action does nothing.
 *
 * @public
 */
export function useDocumentSearch(): UseDocumentSearchResult {
  const editor = usePptxEditor();
  const search = useMemo(() => (editor ? documentSearchOf(editor) : null), [editor]);
  const state = useSyncExternalStore(
    search ? search.subscribe : noSubscription,
    search ? search.getState : idleState,
    idleState
  );
  const actions = useMemo(
    () => ({
      setQuery: (query: string) => search?.setQuery(query),
      find: (query: string, options?: DocumentSearchFindOptions) =>
        search?.find(query, options) ?? IDLE_DOCUMENT_SEARCH.matches,
      setMatchCase: (value: boolean) => search?.setMatchCase(value),
      setWholeWord: (value: boolean) => search?.setWholeWord(value),
      // No arguments: `onClick={search.next}` passes its event, which must not count.
      goTo: (index: number) => search?.goTo(index) ?? false,
      next: () => search?.next() ?? false,
      previous: () => search?.previous() ?? false,
      clear: () => search?.clear(),
    }),
    [search]
  );
  return useMemo(() => ({ ...state, ...actions }), [state, actions]);
}
