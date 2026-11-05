// The link popup's behavior with no UI: the Vue twin of React's `useHyperlinkPopup.ts`, after
// DOCX Editor's composable of the same name. The state and the actions are core's
// (`createHyperlinkPopupStore` in `overlays.ts`); these composables subscribe to one store and
// bind it to the editor's link requests.

import {
  computed,
  getCurrentScope,
  inject,
  onScopeDispose,
  provide,
  ref,
  shallowRef,
  watch,
  useId,
  type ComputedRef,
  type InjectionKey,
  type Ref,
  type ShallowRef,
} from 'vue';
import {
  createHyperlinkPopupStore,
  hyperlinkPopupActions,
  listenHyperlinkPopup,
  openHyperlinkPopup,
  type HyperlinkPopupController,
  type HyperlinkPopupActions,
  type HyperlinkPopupState,
  type HyperlinkPopupStore,
} from '@pptx-editor.dev/core/editor';
import { useEditorCommand, usePptxEditor } from './composables.ts';

/** What `useHyperlinkPopup` returns, the twin of React's, with the state as a ref. @public */
export interface UseHyperlinkPopupReturn extends HyperlinkPopupActions {
  /** Whether the popup shows, what it shows, and the draft. */
  readonly state: ComputedRef<HyperlinkPopupState>;
}

/** What `useHyperlinkPopupInstance` returns: the same shape as `useHyperlinkPopup`. @public */
export type UseHyperlinkPopupInstanceReturn = UseHyperlinkPopupReturn;

/** The link popup store of the enclosing root. @internal */
const hyperlinkPopupStoreKey: InjectionKey<HyperlinkPopupStore> = Symbol('pptxHyperlinkPopupStore');

/** The state and controller shared by the named link parts. @internal */
export interface HyperlinkPopupProvision {
  readonly controller: HyperlinkPopupController;
  readonly state: Readonly<ShallowRef<HyperlinkPopupState>>;
  readonly errorId: string;
}

/** @internal */
export const hyperlinkPopupKey: InjectionKey<HyperlinkPopupProvision> =
  Symbol('pptxHyperlinkPopup');

/** The link popup store of the enclosing root, or null, without counting a view. @internal */
export function injectHyperlinkPopupStore(): HyperlinkPopupStore | null {
  return inject(hyperlinkPopupStoreKey, null);
}
/** Counts the host panels that show the root's link popup state. @internal */
const hyperlinkPopupViewsKey: InjectionKey<() => () => void> = Symbol('pptxHyperlinkPopupViews');

/** Bind `store` to the editor, and while `active`, to its link requests. @internal */
export function useHyperlinkPopupBinding(store: HyperlinkPopupStore, active: () => boolean): void {
  const editorRef = usePptxEditor();
  const cmd = useEditorCommand('text.link');
  onScopeDispose(() => store.setEditor(null));
  watch(
    () => [editorRef.value, active()] as const,
    ([editor, on], _old, onCleanup) => {
      store.setEditor(editor);
      if (!editor || !on) return;
      // Requests for this editor only: another editor on the page opens its own popup.
      onCleanup(listenHyperlinkPopup(store, editor, () => cmd.isEnabled.value));
    },
    { immediate: true }
  );
}

/** The state and actions of `store`, with the state as a ref. */
function storeResult(store: HyperlinkPopupStore): UseHyperlinkPopupReturn {
  const state = shallowRef(store.getState());
  const stop = store.subscribe(() => {
    state.value = store.getState();
  });
  if (getCurrentScope()) onScopeDispose(stop);
  return { state: computed(() => state.value), ...hyperlinkPopupActions(store) };
}

/**
 * Provide the root's link popup store and return it, with the count of host panels that show
 * it. @internal
 */
export function provideHyperlinkPopup(): {
  readonly store: HyperlinkPopupStore;
  readonly views: Readonly<Ref<number>>;
} {
  const store = createHyperlinkPopupStore();
  const views = ref(0);
  provide(hyperlinkPopupStoreKey, store);
  const state = shallowRef(store.getState());
  onScopeDispose(
    store.subscribe(() => {
      state.value = store.getState();
    })
  );
  const relay: HyperlinkPopupController = {
    getState: store.getState,
    subscribe: store.subscribe,
    openReading: (anchor) => store.controller()?.openReading(anchor),
    openEditing: (anchor) => store.controller()?.openEditing(anchor),
    beginEdit: store.beginEdit,
    setUrl: store.setUrl,
    commitEdit: store.commitEdit,
    unlink: store.unlink,
    copy: store.copy,
    openTarget: store.openTarget,
    close: store.close,
    dispose: () => {},
  };
  provide(hyperlinkPopupKey, { controller: relay, state, errorId: useId() });
  provide(hyperlinkPopupViewsKey, () => {
    views.value += 1;
    let attached = true;
    return () => {
      if (!attached) return;
      attached = false;
      views.value -= 1;
    };
  });
  return { store, views };
}

/**
 * The link popup's behavior, for your own link UI, the twin of React's `useHyperlinkPopup`.
 * Inside `PptxEditorRoot` it is the shared state that the packaged popup shows; with
 * `popups.hyperlink: false`, the link controls open your panel while it is mounted. Outside a
 * root it is a standalone instance. @public
 */
export function useHyperlinkPopup(): UseHyperlinkPopupReturn {
  const provided = inject(hyperlinkPopupStoreKey, null);
  const attach = inject(hyperlinkPopupViewsKey, null);
  // Counted only where a scope can count it off again.
  if (provided && attach && getCurrentScope()) onScopeDispose(attach());
  if (provided) return storeResult(provided);
  // A standalone packaged popup provides its props-backed controller to its named parts.
  const own = inject(hyperlinkPopupKey, null);
  if (own) {
    const editorRef = usePptxEditor();
    const c = own.controller;
    return {
      state: computed(() => own.state.value),
      open: (link, anchor) => {
        if (editorRef.value) openHyperlinkPopup(c, editorRef.value, link, anchor);
      },
      openAtCaret: (anchor) => c.openEditing(anchor ?? null),
      close: () => c.close(),
      copy: () => c.copy(),
      beginEdit: () => c.beginEdit(),
      setUrl: (url) => c.setUrl(url),
      commitEdit: () => c.commitEdit(),
      unlink: () => c.unlink(),
      openTarget: () => c.openTarget(),
    };
  }
  return useHyperlinkPopupInstance();
}

/**
 * A link popup of its own, the twin of React's `useHyperlinkPopupInstance`. While `active`, it
 * opens for the link requests of the editor. @public
 */
export function useHyperlinkPopupInstance(active = true): UseHyperlinkPopupInstanceReturn {
  const store = createHyperlinkPopupStore();
  useHyperlinkPopupBinding(store, () => active);
  return storeResult(store);
}
