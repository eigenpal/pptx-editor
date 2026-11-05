// The composables a host uses to build its own chrome: the Vue twins of the React adapter's
// `customization-hooks.ts`, after DOCX Editor's Vue `useZoom`, `useEditorEvent`,
// `useEditorSnapshot`, and `useScopeClassName`. Each reads the editor through its one
// subscription and calls back into the instance: no composable keeps a second copy of editor
// state. Call them in `setup`.

import {
  computed,
  getCurrentScope,
  onScopeDispose,
  shallowRef,
  toValue,
  watch,
  type ComputedRef,
  type MaybeRefOrGetter,
  type Ref,
} from 'vue';
import {
  ZOOM_LEVELS,
  onEditorEvent,
  stepEditorZoom,
  zoomStep,
  type EditorSnapshot,
  type PptxEditorEvents,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor } from './composables.ts';
import { useChromeDark, useScopedByAncestor } from './color-mode.ts';

/** What {@link useZoom} returns: React's `UseZoomReturn`, with refs for the state. @public */
export interface UseZoomReturn {
  /** The zoom in force. 1 is 100 %. In fit mode it is the zoom that fits the slide. */
  readonly zoom: ComputedRef<number>;
  /** Whether the editor fits the slide to the window and follows its size. */
  readonly isFit: ComputedRef<boolean>;
  /** Set a fixed zoom, kept in the range 10 % to 400 %. Leaves fit mode. */
  readonly setZoom: (zoom: number) => void;
  /** Fit the slide to the window, and keep it fitted when the window changes size. */
  readonly fit: () => void;
  /** A fixed 100 %. */
  readonly reset: () => void;
  /** One step of `levels` in from the zoom in force. Does nothing at 400 %. */
  readonly zoomIn: () => void;
  /** One step of `levels` out from the zoom in force. Does nothing at 10 %. */
  readonly zoomOut: () => void;
  /** Whether `zoomIn` can step. */
  readonly canZoomIn: ComputedRef<boolean>;
  /** Whether `zoomOut` can step. */
  readonly canZoomOut: ComputedRef<boolean>;
  /** The steps the zoom buttons move through, so a custom control shows the same levels. */
  readonly levels: readonly number[];
}

interface ZoomSlice {
  readonly zoom: number;
  readonly fit: boolean;
}

const selectZoom = (s: EditorSnapshot): ZoomSlice => ({ zoom: s.zoom, fit: s.fitZoom });
const sameZoom = (a: ZoomSlice, b: ZoomSlice): boolean => a.zoom === b.zoom && a.fit === b.fit;

/**
 * Read and set the slide zoom, as React's `useZoom` and DOCX Editor's. Outside
 * `PptxEditorRoot`, and before the editor mounts, it reports 100 % and every action does
 * nothing. @public
 */
export function useZoom(): UseZoomReturn {
  const editor = usePptxEditor();
  const slice = useEditorState(selectZoom, sameZoom);
  const setZoom = (next: number): void => editor.value?.setZoom(next);
  const next = computed(() => zoomStep(slice.value.zoom, 1));
  const previous = computed(() => zoomStep(slice.value.zoom, -1));
  return {
    zoom: computed(() => slice.value.zoom),
    isFit: computed(() => slice.value.fit),
    setZoom,
    fit: () => editor.value?.setZoom('fit'),
    reset: () => setZoom(1),
    zoomIn: () => stepEditorZoom(editor.value, 1),
    zoomOut: () => stepEditorZoom(editor.value, -1),
    canZoomIn: computed(() => !!editor.value && next.value !== null),
    canZoomOut: computed(() => !!editor.value && previous.value !== null),
    levels: ZOOM_LEVELS,
  };
}

/** What {@link useEditorEvent} returns: nothing; the subscription ends with the scope. @public */
export type UseEditorEventReturn = void;

/** What {@link useEditorSnapshot} returns: the change counter. @public */
export type UseEditorSnapshotReturn = Readonly<Ref<number>>;

/**
 * Subscribe to an editor event (`'change'`, `'selectionChange'`, `'rejection'`,
 * `'openFailure'`) for the lifetime of the component, as React's `useEditorEvent` and DOCX
 * Editor's. It does nothing until `PptxEditorRoot` creates the editor, and subscribes again
 * when the instance changes. @public
 */
export function useEditorEvent<E extends keyof PptxEditorEvents>(
  event: E,
  handler: PptxEditorEvents[E]
): UseEditorEventReturn {
  const editor = usePptxEditor();
  if (!getCurrentScope()) return;
  watch(
    editor,
    (instance, _prev, onCleanup) => {
      if (instance) onCleanup(onEditorEvent(instance, event, handler));
    },
    { immediate: true }
  );
}

/**
 * A counter that changes each time `editor` publishes a change, as React's
 * `useEditorSnapshot` and DOCX Editor's. Read it in a render function or a `computed` to
 * follow the editor. Prefer `useEditorState` with a selector. @public
 */
export function useEditorSnapshot(
  editor: MaybeRefOrGetter<PptxEditorInstance | null>
): UseEditorSnapshotReturn {
  const revision = shallowRef(0);
  if (!getCurrentScope()) return revision;
  const stop = watch(
    () => toValue(editor),
    (instance, _prev, onCleanup) => {
      if (!instance) return;
      const bump = (): void => void (revision.value += 1);
      onCleanup(instance.subscribe(bump));
      bump();
    },
    { immediate: true }
  );
  onScopeDispose(stop);
  return revision;
}

/** What {@link useScopeClassName} returns. @public */
export type UseScopeClassNameReturn = ComputedRef<'' | 'pptx-editor ' | 'pptx-editor dark '>;

/**
 * The scope classes a part you build yourself adds to its own root element, so the editor's
 * `--doc-*` tokens and chrome styles apply to it, as React's `useScopeClassName`. Inside
 * `PptxEditorViewport` it is `''`. Outside it, it is `'pptx-editor '`, or `'pptx-editor dark '`
 * when the root's `colorMode` resolved dark. It is a ref, so it follows a `'system'` color
 * mode. The value ends with a space, so it concatenates cleanly. @public
 */
export function useScopeClassName(): UseScopeClassNameReturn {
  const scoped = useScopedByAncestor();
  const dark = useChromeDark();
  return computed(() => (scoped ? '' : dark.value ? 'pptx-editor dark ' : 'pptx-editor '));
}
