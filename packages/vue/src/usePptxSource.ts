// Adapted from DOCX Editor `packages/vue/src/editor/useDocxSource.ts` and
// `useEditorHostDocument.ts` at commit 5993bd223.
//
// Opening a presentation, as a composable: the Vue twin of React's `usePptxSource.ts`. The
// rules (the two font paths, the fetch) are core's (`pptx-source.ts`); this file owns only the
// Vue lifetime.

import {
  computed,
  shallowRef,
  toValue,
  watch,
  type ComputedRef,
  type MaybeRefOrGetter,
  type ShallowRef,
} from 'vue';
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
  type PptxSource,
  type UsePptxSourceOptions,
} from '@pptx-editor.dev/core/editor';
import type { FontConfiguration } from '@pptx-editor.dev/core/contracts';
import { usePptxEditor } from './composables.ts';

/** What `usePptxSource` returns, the twin of React's, with refs. @public */
export interface UsePptxSourceReturn {
  /** Bytes for the `document` prop; undefined until they arrive. */
  readonly document: ComputedRef<Uint8Array | undefined>;
  /** What to give the `fonts` prop: a configuration, a stable resolver, or undefined. */
  readonly fonts: ComputedRef<FontConfiguration | FontResolver | undefined>;
  /** Why the presentation could not be opened. A font failure never shows here. */
  readonly error: ComputedRef<Error | null>;
  /** True until the presentation arrives or fails, and while eager fonts load. */
  readonly isLoading: ComputedRef<boolean>;
}

/**
 * Load a presentation, and optionally its fonts, for `PptxEditor`, the twin of React's
 * `usePptxSource` and of DOCX Editor's `useDocxSource`. Eager fonts hold the deck back until
 * they settle; an on-demand resolver does not. A font failure never fails the deck. @public
 */
export function usePptxSource(
  source: MaybeRefOrGetter<PptxSource | null | undefined>,
  options: MaybeRefOrGetter<UsePptxSourceOptions> = {}
): UsePptxSourceReturn {
  const bytes = shallowRef<Uint8Array | undefined>(undefined);
  const fonts = shallowRef<FontConfiguration | FontResolver | undefined>(undefined);
  const error = shallowRef<Error | null>(null);
  const documentLoading = shallowRef(toValue(source) != null);
  const fontsSettled = shallowRef(toValue(options).fonts === undefined);
  // One resolver for the composable's life, which reads the options at each request.
  const resolver = pptxSourceFontsResolver(() => toValue(options).fonts);

  // Fonts load once per path, not per options object, as React's: a getter that builds a new
  // `fonts` value at each read must not resolve the fonts again and reload the deck.
  watch(
    () => pptxFontsShape(toValue(options).fonts),
    (_shape, _previous, onCleanup) => {
      const fontsSource = toValue(options).fonts;
      if (fontsSource === undefined) {
        fonts.value = undefined;
        fontsSettled.value = true;
        return;
      }
      // On demand: the answer needs the parsed deck, so there is nothing to wait for.
      if (pptxFontsShape(fontsSource) === 'on-demand') {
        fonts.value = resolver;
        fontsSettled.value = true;
        return;
      }
      let live = true;
      fonts.value = undefined;
      fontsSettled.value = false;
      void resolveEagerPptxFonts(fontsSource).then((composed) => {
        if (!live) return;
        if (composed) fonts.value = composed;
        fontsSettled.value = true;
      });
      onCleanup(() => {
        live = false;
      });
    },
    { immediate: true }
  );

  watch(
    () => toValue(source),
    (next, _previous, onCleanup) => {
      if (next == null) {
        bytes.value = undefined;
        error.value = null;
        documentLoading.value = false;
        return;
      }
      const immediate = pptxSourceBytes(next);
      if (immediate) {
        bytes.value = immediate;
        error.value = null;
        documentLoading.value = false;
        return;
      }
      // Keep the old bytes while the new URL loads, as React does, so the editor keeps its deck.
      let live = true;
      const controller = new AbortController();
      documentLoading.value = true;
      error.value = null;
      fetchPptxSource(next as string | URL, toValue(options).fetchOptions, controller.signal).then(
        (loaded) => {
          if (!live) return;
          bytes.value = loaded;
          documentLoading.value = false;
        },
        (cause: unknown) => {
          if (!live || isAbortError(cause)) return;
          error.value =
            cause instanceof Error ? cause : new Error('could not open the presentation');
          documentLoading.value = false;
        }
      );
      onCleanup(() => {
        live = false;
        controller.abort();
      });
    },
    { immediate: true }
  );

  return {
    document: computed(() => (fontsSettled.value ? bytes.value : undefined)),
    fonts: computed(() => fonts.value),
    error: computed(() => error.value),
    isLoading: computed(() => documentLoading.value || !fontsSettled.value),
  };
}

/** What `useEditorHostDocument` returns: a read-only ref to the host document. @public */
export type UseEditorHostDocumentReturn = Readonly<ShallowRef<Document | null>>;

/**
 * The document that the editor's slide is mounted in, the twin of React's
 * `useEditorHostDocument`: an iframe's when the host mounted the editor there, else the page's
 * own. Null during server rendering. @public
 */
export function useEditorHostDocument(): UseEditorHostDocumentReturn {
  const editorRef = usePptxEditor();
  const current = shallowRef<Document | null>(null);
  const read = (): void => {
    const editor = editorRef.value;
    current.value = editor
      ? editorHostDocument(editor)
      : typeof globalThis.document === 'undefined'
        ? null
        : globalThis.document;
  };
  watch(
    () => editorRef.value,
    (editor, _previous, onCleanup) => {
      read();
      if (editor) onCleanup(onEditorHostChange(editor, read));
    },
    { immediate: true, flush: 'post' }
  );
  return current;
}
