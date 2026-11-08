// Adapted from DOCX Editor `packages/react/src/editor/useDocxSource.ts` and
// `packages/react/src/editor/useEditorHostDocument.ts` at commit 5993bd223.
//
// Opening a presentation, as a hook: fetch the bytes, resolve the fonts, keep the two from
// racing, cancel both when the component unmounts or the source changes, and turn a failure
// into a value to render. The rules (the two font paths, the fetch) are core's
// (`pptx-source.ts`), shared with the Vue twin; this file owns only the React lifetime.

import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import {
  editorHostDocument,
  fetchPptxSource,
  isAbortError,
  onEditorHostChange,
  pptxFontsShape,
  pptxSourceBytes,
  pptxSourceFontsResolver,
  resolveEagerPptxFonts,
  type FontResolver,
  type MarkedFontResolver,
  type PptxSource,
  type UsePptxSourceOptions,
} from '@pptx-editor.dev/core/editor';
import type { FontConfiguration } from '@pptx-editor.dev/core/contracts';
import { usePptxEditor } from './context.ts';

/** What `usePptxSource` returns, as DOCX Editor's `UseDocxSourceResult`. @public */
export interface UsePptxSourceResult {
  /** Bytes for the `document` prop; undefined until they arrive. */
  readonly document: Uint8Array | undefined;
  /**
   * What to give the `fonts` prop: a composed configuration on the eager path, a stable
   * resolver on the on-demand path, or undefined until there is something to give.
   */
  readonly fonts: FontConfiguration | FontResolver | undefined;
  /** Why the presentation could not be opened. A font failure never shows here. */
  readonly error: Error | null;
  /** True until the presentation arrives or fails, and while eager fonts load. */
  readonly isLoading: boolean;
}

/**
 * Load a presentation, and optionally its fonts, for `PptxEditor`, as DOCX Editor's
 * `useDocxSource`.
 *
 * ```tsx
 * const { document, fonts, error } = usePptxSource(url, { fonts: packagedFonts() });
 * if (error) return <p>{error.message}</p>;
 * return <PptxEditor document={document} fonts={fonts} />;
 * ```
 *
 * Eager fonts (a value, a promise, or a zero-argument loader) hold the deck back until they
 * settle, so the slides lay out once. An on-demand resolver (`packagedFonts()`,
 * `googleFonts()`) needs the parsed deck, so the bytes go through at once and the slides lay
 * out again when the faces arrive. A font failure never fails the deck.
 *
 * A URL is fetched with the browser's `fetch`, as written. Check a URL from user input first.
 *
 * @public
 */
export function usePptxSource(
  source: PptxSource | null | undefined,
  options: UsePptxSourceOptions = {}
): UsePptxSourceResult {
  const [bytes, setBytes] = useState<Uint8Array | undefined>(undefined);
  const [fonts, setFonts] = useState<FontConfiguration | undefined>(undefined);
  const [error, setError] = useState<Error | null>(null);
  const [documentLoading, setDocumentLoading] = useState(source != null);
  // The path, read again at each render: a host that builds its origins later changes it.
  const shape = pptxFontsShape(options.fonts);
  const [fontsSettled, setFontsSettled] = useState(options.fonts === undefined);

  // Options are new at each render when written inline: read them from a ref, so a new object
  // does not start another fetch.
  const latest = useRef(options);
  latest.current = options;
  // One resolver for the hook's life, so the `fonts` prop keeps its identity.
  const resolver = useMemo<MarkedFontResolver>(
    () => pptxSourceFontsResolver(() => latest.current.fonts),
    []
  );

  // Fonts load once per path, not per deck: they belong to the app, not to the file.
  useEffect(() => {
    const fontsSource = latest.current.fonts;
    if (fontsSource === undefined) {
      setFonts(undefined);
      setFontsSettled(true);
      return undefined;
    }
    if (pptxFontsShape(fontsSource) === 'on-demand') {
      setFontsSettled(true);
      return undefined;
    }
    setFontsSettled(false);
    let live = true;
    void resolveEagerPptxFonts(fontsSource).then((composed) => {
      if (!live) return;
      if (composed) setFonts(composed);
      setFontsSettled(true);
    });
    return () => {
      live = false;
    };
  }, [shape]);

  useEffect(() => {
    if (source == null) {
      setBytes(undefined);
      setError(null);
      setDocumentLoading(false);
      return undefined;
    }
    const immediate = pptxSourceBytes(source);
    if (immediate) {
      setBytes(immediate);
      setError(null);
      setDocumentLoading(false);
      return undefined;
    }
    // `live` as well as the abort: a response already read can resolve after unmount.
    let live = true;
    const controller = new AbortController();
    setDocumentLoading(true);
    setError(null);
    fetchPptxSource(source as string | URL, latest.current.fetchOptions, controller.signal).then(
      (loaded) => {
        if (!live) return;
        setBytes(loaded);
        setDocumentLoading(false);
      },
      (cause: unknown) => {
        if (!live || isAbortError(cause)) return;
        setError(cause instanceof Error ? cause : new Error('could not open the presentation'));
        setDocumentLoading(false);
      }
    );
    return () => {
      live = false;
      controller.abort();
    };
  }, [source]);

  return {
    // Held back until eager fonts settle, so the first layout is the only layout.
    document: fontsSettled ? bytes : undefined,
    fonts: shape === 'on-demand' ? resolver : fonts,
    error,
    isLoading: documentLoading || !fontsSettled,
  };
}

/** What `useEditorHostDocument` returns: the host document, or null with no DOM. @public */
export type UseEditorHostDocumentReturn = Document | null;

const NOOP_UNSUBSCRIBE = (): void => {};

/**
 * The document that the editor's slide is mounted in, as DOCX Editor's
 * `useEditorHostDocument`: an iframe's when the host mounted the editor there, else the page's
 * own. Null during server rendering. Your own chrome that adds document listeners, reads
 * focus, or adds a `<style>` uses it in place of the global `document`. Its `defaultView` is
 * the editor's window.
 *
 * @example
 * ```tsx
 * const hostDocument = useEditorHostDocument();
 * useEffect(() => {
 *   if (!hostDocument) return;
 *   hostDocument.addEventListener('keydown', onShortcut);
 *   return () => hostDocument.removeEventListener('keydown', onShortcut);
 * }, [hostDocument]);
 * ```
 *
 * @public
 */
export function useEditorHostDocument(): UseEditorHostDocumentReturn {
  const editor = usePptxEditor();
  const subscribe = useCallback(
    (onChange: () => void) => (editor ? onEditorHostChange(editor, onChange) : NOOP_UNSUBSCRIBE),
    [editor]
  );
  const read = useCallback(
    (): Document | null =>
      editor
        ? editorHostDocument(editor)
        : typeof globalThis.document === 'undefined'
          ? null
          : globalThis.document,
    [editor]
  );
  return useSyncExternalStore(subscribe, read, () => null);
}
