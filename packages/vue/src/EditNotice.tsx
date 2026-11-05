/** @jsxImportSource vue */
// The edit notice: the Vue twin of the React adapter's `EditNotice.tsx`. It says why the last
// edit was refused (a keystroke, a paste, a command), so a refused edit never looks like an
// editor that stopped responding. After a Delete of shapes or slides it says what went and
// offers Undo, for ten seconds; a pointer over it or the focus in it holds it. It reads the
// engine's `lastRejection` and `lastRemoval` and holds no state of its own besides the
// dismissal and the hold. The messages and the rules for when they show are core's
// (`overlays.ts`).

import {
  computed,
  defineComponent,
  onBeforeUnmount,
  onScopeDispose,
  shallowRef,
  watch,
  watchEffect,
  type ExtractPublicPropTypes,
} from 'vue';
import {
  editNoticeShows,
  editNoticeTone,
  editRejectionMessage,
  editRemovalMessage,
  focusSurfaceOf,
  linkPopupShowsRejection,
  REMOVAL_NOTICE_MS,
  removalNoticeShows,
  type EditorRejection,
} from '@pptx-editor.dev/core/editor';
import type { Translate } from '@pptx-editor.dev/i18n';
import { ToolButton, keepFocus } from './chrome.tsx';
import { useEditorState, usePptxEditor } from './composables.ts';
import { chromeTranslateProp, useChromeTranslateScope } from './headless-hooks.ts';
import { usePopupEntry } from './popups.ts';
import { injectHyperlinkPopupStore } from './useHyperlinkPopup.ts';

/** The message for a refusal: its reason, with the detail where the reason names one. @public */
export function rejectionMessage(t: Translate, rejection: EditorRejection): string {
  return editRejectionMessage(t, rejection);
}

const pptxEditorEditNoticeProps = {
  ...chromeTranslateProp,
  /** A class name added to the notice, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** The props of `PptxEditorEditNotice`, the same name as React's `EditNoticeProps`. @public */
export type EditNoticeProps = ExtractPublicPropTypes<typeof pptxEditorEditNoticeProps>;

/**
 * Why the last edit was refused, or what the last Delete removed with an Undo action; the twin
 * of React's `PptxEditor.EditNotice`. Renders nothing while edits apply, and nothing for a
 * command that had nothing to do (`EditorRejection.noop`). A host `class` goes on the notice.
 * @public
 */
export const PptxEditorEditNotice = defineComponent({
  name: 'PptxEditorEditNotice',
  inheritAttrs: false,
  props: pptxEditorEditNoticeProps,
  setup(props, { attrs }) {
    const editorRef = usePptxEditor();
    const { t } = useChromeTranslateScope(() => props.t);
    const rejection = useEditorState((s) => s.lastRejection ?? null);
    const removal = useEditorState((s) => s.lastRemoval ?? null);
    const dismissed = shallowRef<EditorRejection | null>(null);
    // A refused link shows its reason in the open link popup: the notice does not repeat it,
    // and stays closed when the popup closes. As React's `EditNotice`.
    const linkStore = injectHyperlinkPopupStore();
    const linkPopup = shallowRef(linkStore?.getState() ?? null);
    if (linkStore)
      onScopeDispose(
        linkStore.subscribe(() => {
          linkPopup.value = linkStore.getState();
        })
      );
    const linkEntry = usePopupEntry('hyperlink');
    const covered = computed(
      () => linkEntry() !== false && linkPopupShowsRejection(linkPopup.value, rejection.value)
    );
    watch(
      covered,
      (now) => {
        if (now) dismissed.value = rejection.value;
      },
      { immediate: true }
    );
    // The serial of the removal the user closed or that timed out.
    const closed = shallowRef<number | null>(null);
    const held = shallowRef(false);
    // A click on Undo or Dismiss removes the notice under the pointer, and no mouseleave or
    // focusout follows: each new removal starts unheld.
    watch(
      () => removal.value?.serial,
      () => {
        held.value = false;
      }
    );
    let timer: ReturnType<typeof setTimeout> | undefined;
    watchEffect((onCleanup) => {
      const current = removal.value;
      if (!removalNoticeShows(current, closed.value) || held.value) return;
      const serial = current.serial;
      timer = setTimeout(() => {
        closed.value = serial;
      }, REMOVAL_NOTICE_MS);
      onCleanup(() => clearTimeout(timer));
    });
    onBeforeUnmount(() => clearTimeout(timer));
    return () => {
      const classes = ['pptx-edit-notice', attrs.class, props.className];
      const refused = rejection.value;
      if (!covered.value && editNoticeShows(refused, dismissed.value))
        return (
          <div
            {...attrs}
            class={classes}
            // A state the edit cannot run in (view mode) is information, not an error.
            data-tone={editNoticeTone(refused)}
            role="status"
            onMousedown={keepFocus}
          >
            <span class="pptx-edit-notice__text">{rejectionMessage(t, refused)}</span>
            <ToolButton
              label={t('notice.dismiss')}
              icon="close"
              onClick={() => {
                dismissed.value = refused;
              }}
            />
          </div>
        );
      const removed = removal.value;
      if (!removalNoticeShows(removed, closed.value)) return null;
      return (
        <div
          {...attrs}
          class={classes}
          data-tone="info"
          role="status"
          onMousedown={keepFocus}
          onMouseenter={() => (held.value = true)}
          onMouseleave={() => (held.value = false)}
          onFocusin={() => (held.value = true)}
          onFocusout={() => (held.value = false)}
        >
          <span class="pptx-edit-notice__text">{editRemovalMessage(t, removed)}</span>
          <button
            type="button"
            class="pptx-edit-notice__action"
            onClick={() => {
              const editor = editorRef.value;
              if (!editor) return;
              editor.exec({ type: 'undo' });
              // The button goes with the notice; the focus goes back to the slide.
              focusSurfaceOf(editor);
            }}
          >
            {t('notice.undo')}
          </button>
          <ToolButton
            label={t('notice.dismiss')}
            icon="close"
            onClick={() => {
              closed.value = removed.serial;
              if (editorRef.value) focusSurfaceOf(editorRef.value);
            }}
          />
        </div>
      );
    };
  },
});
