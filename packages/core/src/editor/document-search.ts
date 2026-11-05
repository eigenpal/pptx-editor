// A find session with no UI, for a host's own find bar: the query, the match options, the
// matches, and the current match, over the engine's `editor.find`. After DOCX Editor's
// `createDocumentSearch` and `useDocumentSearch`. Each editor has one session, so every hook
// on it shows the same search.

import type { FindMatch, FindOptions } from '../contracts/index.ts';
import type { PptxEditorInstance } from './pptx-editor.ts';

/** The state of a find session. @public */
export interface DocumentSearchState {
  /** The text to find. */
  readonly query: string;
  readonly matchCase: boolean;
  readonly wholeWord: boolean;
  /** Every match of `query`, in deck order: slides, then their notes. */
  readonly matches: readonly FindMatch[];
  /** Always false: the engine returns every match. Kept for DOCX Editor's shape. */
  readonly truncated: boolean;
  /** The index of the selected match, or `-1` when the selection is on no match. */
  readonly activeIndex: number;
  /** The selected match, or null. */
  readonly activeMatch: FindMatch | null;
  /** Always false: the search runs at once. Kept for DOCX Editor's shape. */
  readonly isPending: boolean;
}

/** Options of `find`. @public */
export interface DocumentSearchFindOptions extends FindOptions {
  /** Select the first match. */
  readonly selectFirst?: boolean;
}

/** A find session. @internal */
export interface DocumentSearch {
  getState(): DocumentSearchState;
  subscribe(listener: () => void): () => void;
  setQuery(query: string): void;
  setMatchCase(value: boolean): void;
  setWholeWord(value: boolean): void;
  find(query: string, options?: DocumentSearchFindOptions): readonly FindMatch[];
  goTo(index: number): boolean;
  next(): boolean;
  previous(): boolean;
  clear(): void;
}

const NO_MATCHES: readonly FindMatch[] = Object.freeze([]);

/** The state of an editor with no session. @internal */
export const IDLE_DOCUMENT_SEARCH: DocumentSearchState = Object.freeze({
  query: '',
  matchCase: false,
  wholeWord: false,
  matches: NO_MATCHES,
  truncated: false,
  activeIndex: -1,
  activeMatch: null,
  isPending: false,
});

const sessions = new WeakMap<PptxEditorInstance, DocumentSearch>();

/** The find session of `editor`, made on first use. @internal */
export function documentSearchOf(editor: PptxEditorInstance): DocumentSearch {
  let search = sessions.get(editor);
  if (!search) sessions.set(editor, (search = createDocumentSearch(editor)));
  return search;
}

const sameMatches = (a: readonly FindMatch[], b: readonly FindMatch[]): boolean =>
  a.length === b.length &&
  a.every((m, i) => {
    const n = b[i]!;
    return m.txBodyId === n.txBodyId && m.paragraph === n.paragraph && m.start === n.start;
  });

function createDocumentSearch(editor: PptxEditorInstance): DocumentSearch {
  let state = IDLE_DOCUMENT_SEARCH;
  const listeners = new Set<() => void>();
  let unsubscribe: (() => void) | null = null;
  const options = (): FindOptions => ({ matchCase: state.matchCase, wholeWord: state.wholeWord });
  /** Read the matches and the current one again; tell the listeners when they changed. */
  const refresh = (next: Partial<DocumentSearchState> = {}): void => {
    const draft = { ...state, ...next };
    const opts = { matchCase: draft.matchCase, wholeWord: draft.wholeWord };
    const matches = draft.query ? editor.find.matches(draft.query, opts) : NO_MATCHES;
    const activeIndex = matches.length ? editor.find.position(draft.query, opts) : -1;
    const changed =
      draft.query !== state.query ||
      draft.matchCase !== state.matchCase ||
      draft.wholeWord !== state.wholeWord ||
      activeIndex !== state.activeIndex ||
      !sameMatches(matches, state.matches);
    if (!changed) return;
    state = Object.freeze({
      ...draft,
      matches: sameMatches(matches, state.matches) ? state.matches : matches,
      activeIndex,
      activeMatch: matches[activeIndex] ?? null,
    });
    for (const listener of [...listeners]) listener();
  };
  const step = (move: 'next' | 'previous'): boolean => {
    if (!state.query) return false;
    const result = editor.find[move](state.query, options());
    refresh();
    return result.ok && result.match !== null;
  };
  return {
    getState: () => state,
    subscribe(listener) {
      listeners.add(listener);
      // Edits and selection changes move the matches and the current one.
      if (!unsubscribe) {
        unsubscribe = editor.subscribe(() => refresh());
        // With no listener, nothing followed the edits or a new deck: read the matches again.
        refresh();
      }
      return () => {
        listeners.delete(listener);
        if (listeners.size === 0) {
          unsubscribe?.();
          unsubscribe = null;
        }
      };
    },
    setQuery: (query) => refresh({ query }),
    setMatchCase: (matchCase) => refresh({ matchCase }),
    setWholeWord: (wholeWord) => refresh({ wholeWord }),
    find(query, findOptions = {}) {
      const { selectFirst, ...rest } = findOptions;
      refresh({
        query,
        matchCase: rest.matchCase ?? state.matchCase,
        wholeWord: rest.wholeWord ?? state.wholeWord,
      });
      if (selectFirst && state.matches.length) this.goTo(0);
      return state.matches;
    },
    goTo(index) {
      const total = state.matches.length;
      if (!state.query || index < 0 || index >= total) return false;
      // `next` wraps at the end, so at most `total` steps reach any match.
      for (let i = 0; i <= total && editor.find.position(state.query, options()) !== index; i++)
        editor.find.next(state.query, options());
      refresh();
      return state.activeIndex === index;
    },
    next: () => step('next'),
    previous: () => step('previous'),
    clear: () => refresh({ query: '' }),
  };
}
