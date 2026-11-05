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
  createHyperlinkPopupController,
  focusSurfaceOf,
  hyperlinkEntryReads,
  hyperlinkPopupFromRequest,
  linkPopupAnnouncement,
  linkPopupReturnsFocus,
  listenLinkClicks,
  listenLinkPanelKey,
  listenLinkRequests,
  type HyperlinkPopupController,
  type HyperlinkPopupState,
} from '@pptx-editor.dev/core/editor';
import { useEditorCommand, usePptxEditor, useTranslation } from './composables.ts';
import { hyperlinkPopupKey, PptxEditorHyperLink } from './HyperLink.tsx';
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
  setup() {
    const editorRef = usePptxEditor();
    const entry = usePopupEntry('hyperlink');
    const cmd = useEditorCommand('text.link');
    const { t } = useTranslation();
    const errorId = useId();
    const state = shallowRef<HyperlinkPopupState | null>(null);
    const controllerRef = shallowRef<HyperlinkPopupController | null>(null);
    const closed: HyperlinkPopupState = {
      mode: 'closed',
      url: '',
      href: null,
      draft: '',
      error: null,
      copied: false,
      canEdit: false,
      anchor: null,
    };
    // The parts read the provision; it follows the current controller.
    const provided = shallowRef<HyperlinkPopupState>(closed);
    const relay: HyperlinkPopupController = {
      getState: () => provided.value,
      subscribe: (l) => controllerRef.value?.subscribe(l) ?? (() => {}),
      openReading: (a) => controllerRef.value?.openReading(a),
      openEditing: (a) => controllerRef.value?.openEditing(a),
      beginEdit: () => controllerRef.value?.beginEdit(),
      setUrl: (u) => controllerRef.value?.setUrl(u),
      commitEdit: () => controllerRef.value?.commitEdit() ?? false,
      unlink: () => controllerRef.value?.unlink() ?? false,
      copy: () => controllerRef.value?.copy() ?? Promise.resolve(false),
      openTarget: () => controllerRef.value?.openTarget() ?? false,
      close: () => controllerRef.value?.close(),
      dispose: () => {},
    };
    provide(hyperlinkPopupKey, { controller: relay, state: provided, errorId });
    // One controller per editor, as React's root: turning the entry off or on keeps it, so an
    // open popup is closed (with focus back on the slide), never dropped.
    watch(
      editorRef,
      (editor, _old, onCleanup) => {
        if (!editor) {
          state.value = null;
          provided.value = closed;
          return;
        }
        const controller = createHyperlinkPopupController(editor);
        controllerRef.value = controller;
        let open = false;
        const stopState = controller.subscribe(() => {
          const now = controller.getState();
          // On close, focus goes back to the slide (the text input while text is selected), as
          // DOCX Editor's HyperLink does; focus that moved elsewhere on purpose stays there.
          const isOpen = now.mode !== 'closed';
          if (open && !isOpen && linkPopupReturnsFocus(editor)) focusSurfaceOf(editor);
          open = isOpen;
          state.value = now;
          provided.value = now;
        });
        state.value = controller.getState();
        provided.value = state.value;
        onCleanup(() => {
          controller.close();
          provided.value = closed;
          stopState();
          controller.dispose();
          if (controllerRef.value === controller) controllerRef.value = null;
        });
      },
      { immediate: true }
    );
    // Turning links off closes an open popup; a host panel that reads `useHyperlinkPopup()` sees
    // it closed, as React's.
    watch(
      () => [controllerRef.value, entry() === false] as const,
      ([controller, off]) => {
        if (off) controller?.close();
      },
      { immediate: true }
    );
    watch(
      () =>
        [
          editorRef.value,
          controllerRef.value,
          entry() === false,
          hyperlinkEntryReads(entry()),
        ] as const,
      ([editor, controller, off, reads], _old, onCleanup) => {
        if (!editor || !controller || off) return;
        // Requests for this root's editor only: another editor on the page opens its own.
        const stopRequests = listenLinkRequests(
          editor,
          () => cmd.isEnabled.value,
          (at) => controller.openEditing(at)
        );
        // A click into linked text opens the reading panel under the caret, also for a viewer:
        // Edit and Unlink show only when the link command can run. A host renderer gets reading
        // mode only when it asks for it. F6 moves focus from the slide text into the panel.
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
      const controller = controllerRef.value;
      if (!s || !controller || current === false) return null;
      // A polite live region, always mounted so a screen reader hears the panel open and the copy.
      const live = (
        <span class="pptx-sr-only" role="status" aria-live="polite">
          {linkPopupAnnouncement(s, t)}
        </span>
      );
      if (s.mode === 'closed') return live;
      const panel =
        typeof current === 'function' ? (
          current({
            anchor: s.anchor,
            url: s.url,
            mode: s.mode,
            onClose: () => controller.close(),
          })
        ) : (
          <PptxEditorHyperLink />
        );
      return [panel, live];
    };
  },
});
