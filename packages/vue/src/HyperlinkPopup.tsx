/** @jsxImportSource vue */
// The link popup: the Vue twin of the React adapter's `HyperlinkPopup.tsx` and of its root host.
// `HyperlinkPopup` is the packaged panel, `PptxEditorHyperLink` (`HyperLink.tsx`) with its default
// parts, over the root's one popup state, else over a state built from its props.
// `PptxEditorRoot` hosts it (`LinkPopupHost`): the link control, Insert > Link, and Ctrl/Cmd+K
// open it in edit mode, a click into linked text in reading mode; the `popups.hyperlink`
// registry entry can replace it or turn it off. The request, anchor, and click rules are core's
// (`overlays.ts`, `link-popup.ts`).

import {
  defineComponent,
  inject,
  onUnmounted,
  provide,
  shallowRef,
  useId,
  watch,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
import {
  focusSurfaceOf,
  hyperlinkEntryReads,
  hyperlinkPopupFromRequest,
  linkPopupAnnouncement,
  linkPopupReturnsFocus,
  listenLinkClicks,
  listenLinkPanelKey,
  listenLinkRequests,
  type HyperlinkPopupController,
  type HyperlinkPopupStore,
  type HyperlinkPopupState,
} from '@pptx-editor.dev/core/editor';
import { useEditorCommand, usePptxEditor, useTranslation } from './composables.ts';
import { PptxEditorHyperLink } from './HyperLink.tsx';
import { hyperlinkPopupKey } from './useHyperlinkPopup.ts';
import { usePopupEntry } from './popups.ts';

const hyperlinkPopupProps = {
  /**
   * The element the popup opens under: the caret after a click into a link, the toolbar's link
   * button, else the element the request came from (the slide for Ctrl/Cmd+K).
   */
  // No runtime type check: an element from another window (an iframe) is not `HTMLElement`.
  anchor: { type: null as unknown as PropType<HTMLElement | null>, required: true as const },
  /** The link of the selection now, or an empty string. */
  url: { type: String, required: true as const },
  /**
   * `reading` after a click into linked text (the target, Copy, Edit, Unlink), `editing` for the
   * URL field. Absent means `editing`. `useHyperlinkPopup()` gives the state and the actions.
   */
  mode: { type: String as PropType<'reading' | 'editing'>, default: undefined },
  /** Close the popup. Apply the link with `editor.exec({ type: 'setLink', url })`. */
  onClose: { type: Function as PropType<() => void>, required: true as const },
  /** A class name added to the popup panel, for a host renderer that forwards these props. */
  className: { type: String, default: undefined },
} as const;

/** What a replacement link popup receives, the twin of React's `HyperlinkPopupProps`. @public */
export type HyperlinkPopupProps = ExtractPublicPropTypes<typeof hyperlinkPopupProps>;

/** Provide `controller` to the parts, with its state as a ref. Stops with the component. */
function provideController(controller: HyperlinkPopupController, errorId: string): void {
  const state = shallowRef<HyperlinkPopupState>(controller.getState());
  const stop = controller.subscribe(() => {
    state.value = controller.getState();
  });
  onUnmounted(stop);
  provide(hyperlinkPopupKey, { controller, state, errorId });
}

/**
 * The packaged link popup: `PptxEditorHyperLink` with its default parts. A `popups.hyperlink`
 * renderer can forward its props to it. A host `class` goes on the popup panel, as `className`
 * does. @public
 */
export const HyperlinkPopup = defineComponent({
  name: 'HyperlinkPopup',
  inheritAttrs: false,
  props: hyperlinkPopupProps,
  setup(props, { attrs }) {
    // Outside the root that owns the shared state, a state of its own from the props.
    if (!inject(hyperlinkPopupKey, null)) {
      const editorRef = usePptxEditor();
      const own = hyperlinkPopupFromRequest(editorRef.value, {
        anchor: props.anchor,
        url: props.url,
        onClose: () => props.onClose(),
        ...(props.mode ? { mode: props.mode } : {}),
      });
      onUnmounted(() => own.dispose());
      provideController(own, useId());
    }
    return () => {
      const extra = [props.className, attrs.class as string | undefined].filter(Boolean).join(' ');
      return <PptxEditorHyperLink {...(extra ? { className: extra } : {})} />;
    };
  },
});

/**
 * The link popup of one root: `popups.hyperlink`, or the packaged `PptxEditorHyperLink`. It owns
 * the root's one popup state. Placed by `PptxEditorRoot`, so every opener works with no toolbar
 * on the page, as React's root. @internal
 */
export const LinkPopupHost = defineComponent({
  name: 'PptxLinkPopupHost',
  props: {
    store: { type: Object as PropType<HyperlinkPopupStore>, required: true as const },
    views: { type: Number, required: true as const },
  },
  setup(props) {
    const editorRef = usePptxEditor();
    const entry = usePopupEntry('hyperlink');
    const cmd = useEditorCommand('text.link');
    const { t } = useTranslation();
    const state = shallowRef(props.store.getState());
    let open = false;
    const stopState = props.store.subscribe(() => {
      const now = props.store.getState();
      const editor = editorRef.value;
      const isOpen = now.mode !== 'closed';
      if (open && !isOpen && editor && linkPopupReturnsFocus(editor)) focusSurfaceOf(editor);
      open = isOpen;
      state.value = now;
    });
    onUnmounted(() => {
      props.store.setEditor(null);
      stopState();
    });
    watch(editorRef, (editor) => props.store.setEditor(editor), { immediate: true });
    // A host hook keeps link controls active when the packaged popup is disabled.
    watch(
      () => entry() === false && props.views === 0,
      (off) => {
        if (off) props.store.close();
      },
      { immediate: true }
    );
    watch(
      () => [editorRef.value, entry(), props.views] as const,
      ([editor, current, views], _old, onCleanup) => {
        const controller = props.store.controller();
        if (!editor || !controller || (current === false && views === 0)) return;
        const stopRequests = listenLinkRequests(
          editor,
          () => cmd.isEnabled.value,
          (at) => controller.openEditing(at)
        );
        // Preserve reading clicks and F6. Custom renderers still opt into reading mode.
        const reads = current === false ? views > 0 : hyperlinkEntryReads(current);
        const stopClicks = reads
          ? listenLinkClicks(editor, (at) => controller.openReading(at))
          : () => {};
        const stopKey = reads
          ? listenLinkPanelKey(editor, () => controller.getState().mode === 'reading')
          : () => {};
        onCleanup(() => {
          stopRequests();
          stopClicks();
          stopKey();
        });
      },
      { immediate: true }
    );
    return () => {
      const current = entry();
      const s = state.value;
      if (current === false) return null;
      // Keep the live region mounted so reading and copy state changes are announced.
      const live = (
        <span class="pptx-sr-only" role="status" aria-live="polite">
          {linkPopupAnnouncement(s, t)}
        </span>
      );
      if (s.mode === 'closed') return live;
      const panel =
        typeof current === 'function' ? (
          current({ anchor: s.anchor, url: s.url, mode: s.mode, onClose: props.store.close })
        ) : (
          <PptxEditorHyperLink />
        );
      return [panel, live];
    };
  },
});
