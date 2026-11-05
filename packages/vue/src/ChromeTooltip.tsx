/** @jsxImportSource vue */
// The chrome tooltip bubble: the Vue twin of the React adapter's `ChromeTooltip.tsx`, after DOCX
// Editor's `Tooltip`. The editor root mounts one; it shows the `data-pptx-tooltip` label of the
// control under the pointer, or of the control a keyboard user moved to, after core's delay, with
// a disabled control's reason on its own line. Core's `chrome-tooltip.ts` owns the timing and the
// position.

import { Teleport, defineComponent, onMounted, onUnmounted, ref, shallowRef, watch } from 'vue';
import {
  listenChromeTooltips,
  positionChromeTooltip,
  type ChromeTooltipState,
} from '@pptx-editor.dev/core/editor';

/** The tooltip bubble for every chrome control on the page. @internal */
export const ChromeTooltip = defineComponent({
  name: 'ChromeTooltip',
  setup() {
    const state = shallowRef<ChromeTooltipState | null>(null);
    const bubble = ref<HTMLElement | null>(null);
    let stop: (() => void) | null = null;
    onMounted(() => {
      stop = listenChromeTooltips(document, (next) => (state.value = next));
    });
    onUnmounted(() => stop?.());
    // After the DOM update and before paint: the bubble measures itself, then core places it.
    watch(
      [state, bubble],
      ([next, el]) => {
        if (next && el) positionChromeTooltip(el, next.target);
      },
      { flush: 'post' }
    );
    return () => {
      const s = state.value;
      if (!s) return null;
      return (
        <Teleport to={s.scope}>
          {/* Visual only: the control's `aria-label` and `aria-describedby` carry the same text. */}
          <div ref={bubble} class="pptx-tooltip" role="tooltip" aria-hidden="true">
            <span class="pptx-tooltip__label">{s.label}</span>
            {s.reason ? <span class="pptx-tooltip__reason">{s.reason}</span> : null}
          </div>
        </Teleport>
      );
    };
  },
});
