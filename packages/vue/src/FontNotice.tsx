/** @jsxImportSource vue */
// The font substitution notice: the Vue twin of the React adapter's `FontNotice.tsx`, matching
// DOCX Editor's. It lists the families that render with a substitute that can change line
// breaks, and offers the fonts installed on this computer where the browser supports it.
// Metric-compatible substitutes are not listed: their line breaks match the original font. The
// list and the local font request are core's (`overlays.ts`).

import {
  computed,
  defineComponent,
  ref,
  shallowRef,
  watch,
  type ExtractPublicPropTypes,
} from 'vue';
import {
  applyLocalFonts,
  canQueryLocalFonts,
  fontNoticeFamilies,
  fontNoticeList,
} from '@pptx-editor.dev/core/editor';
import { ToolButton, keepFocus } from './chrome.tsx';
import { useEditorState, usePptxEditor } from './composables.ts';
import { chromeTranslateProp, useChromeTranslateScope } from './headless-hooks.ts';

const pptxEditorFontNoticeProps = {
  ...chromeTranslateProp,
  /** A class name added to the notice, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** The props of `PptxEditorFontNotice`, the same name as React's `FontNoticeProps`. @public */
export type FontNoticeProps = ExtractPublicPropTypes<typeof pptxEditorFontNoticeProps>;

/**
 * The substitution notice, the twin of React's `PptxEditor.FontNotice`. Renders nothing when
 * every font is available. A host `class` goes on the notice. @public
 */
export const PptxEditorFontNotice = defineComponent({
  name: 'PptxEditorFontNotice',
  inheritAttrs: false,
  props: pptxEditorFontNoticeProps,
  setup(props, { attrs }) {
    const editorRef = usePptxEditor();
    const { t } = useChromeTranslateScope(() => props.t);
    const substituted = useEditorState((s) => s.substitutedFonts);
    // A newly opened document starts with no outcome of an earlier local font request.
    const loadState = useEditorState((s) => s.status);
    const missing = computed(() => fontNoticeFamilies(substituted.value));
    const key = computed(() => missing.value.join('|'));
    const dismissed = shallowRef<string | null>(null);
    // The outcome of a local font request, for the substituted families when it finished.
    const outcome = shallowRef<{ text: string; key: string } | null>(null);
    // A later change to the substituted families makes the outcome stale.
    const status = computed(() => (outcome.value?.key === key.value ? outcome.value.text : null));
    const busy = ref(false);
    // Each opened document starts a new round; a request from an earlier round reports nothing.
    // A request that adds fonts changes `key` itself, so only a load starts a round.
    let round = 0;
    watch(loadState, () => {
      round += 1;
      outcome.value = null;
    });
    const addLocalFonts = async (): Promise<void> => {
      const editor = editorRef.value;
      if (!editor) return;
      busy.value = true;
      const asked = round;
      try {
        const result = await applyLocalFonts(editor, missing.value);
        const now = fontNoticeFamilies(editor.snapshot().substitutedFonts).join('|');
        if (asked === round) outcome.value = { text: t(result), key: now };
      } catch {
        // `setFonts` rejects only after the editor is destroyed; nothing is left to tell.
      } finally {
        busy.value = false;
      }
    };
    return () => {
      const families = missing.value;
      if ((families.length === 0 && !status.value) || dismissed.value === key.value) return null;
      return (
        <div
          {...attrs}
          class={['pptx-font-notice', attrs.class, props.className]}
          role="status"
          onMousedown={keepFocus}
        >
          <span class="pptx-font-notice__text">
            {families.length
              ? t('fonts.notice', { families: fontNoticeList(families, t) })
              : status.value}
            {families.length && status.value ? ` ${status.value}` : ''}
          </span>
          {families.length && canQueryLocalFonts() ? (
            <button
              type="button"
              class="pptx-font-notice__action"
              disabled={busy.value || loadState.value === 'loading'}
              onClick={() => void addLocalFonts()}
            >
              {t('fonts.useLocal')}
            </button>
          ) : null}
          <ToolButton
            label={t('fonts.dismiss')}
            icon="close"
            onClick={() => {
              dismissed.value = key.value;
            }}
          />
        </div>
      );
    };
  },
});
