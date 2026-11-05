/** @jsxImportSource vue */
// The toolbar's Font size control: Decrease, the size box, and Increase. The Vue twin of the
// React adapter's `toolbar-size.tsx`; the stepper shape of `CHROME_GROUPS` renders it, and core
// gives the box text and the parse (`toolbarSizeText`, `parseToolbarSize`).

import { defineComponent, onBeforeUnmount, ref, shallowRef, useId, type PropType } from 'vue';
import {
  SIZE_ERROR_MS,
  parseToolbarSize,
  sizeErrorPlace,
  toolbarSizeRefusal,
  toolbarSizeText,
  type ChromeControl,
  type ChromeSlotId,
} from '@pptx-editor.dev/core/editor';
import { ToolButton, focusText, useAriaDisabled } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { useSlot } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';

/** The `font.size` stepper. @internal */
export const SizeControl = defineComponent({
  name: 'PptxSizeControl',
  props: {
    slot: { type: String as PropType<ChromeSlotId>, required: true as const },
    control: { type: Object as PropType<ChromeControl>, required: true as const },
  },
  setup(props) {
    const { cmd, label, reason, tip } = useSlot(
      () => props.slot,
      () => props.control
    );
    const editor = usePptxEditor();
    const { t } = useTranslation();
    // The steps are off over a mixed size, as DOCX Editor turns its steppers off.
    const mixed = useEditorState((s) => s.format?.font?.disagrees.fontSizePt === true);
    const draft = ref<string | null>(null);
    const box = ref<HTMLInputElement | null>(null);
    // Set when Enter or Escape has already handled the box, so the blur that follows does nothing.
    let handled = false;
    const off = useAriaDisabled(() => !cmd.isEnabled.value, reason);
    const trigger = useTriggerAttributes(props.slot, 'pptx-toolbar__stepper', () => ({
      disabled: !cmd.isEnabled.value,
    }));
    // A size out of range is refused with its range, under the box, and announced.
    const errorId = useId();
    const error = shallowRef<{ text: string; left: number; top: number } | null>(null);
    let timer: ReturnType<typeof setTimeout> | undefined;
    const showError = (next: { text: string; left: number; top: number } | null): void => {
      clearTimeout(timer);
      error.value = next;
      if (next) timer = setTimeout(() => (error.value = null), SIZE_ERROR_MS);
    };
    onBeforeUnmount(() => clearTimeout(timer));
    const commit = (): void => {
      const refusal = toolbarSizeRefusal(draft.value);
      const n = parseToolbarSize(draft.value);
      if (!handled && refusal)
        showError({ text: t(refusal.key, refusal.params), ...sizeErrorPlace(box.value) });
      else if (!handled && n !== null) cmd.execute(n);
      handled = false;
      draft.value = null;
    };
    // Enter and Escape hand the caret back to the text, as DOCX Editor's size box does, so a
    // size armed at a caret applies to the text typed next.
    const leave = (apply: boolean): void => {
      if (apply) commit();
      draft.value = null;
      handled = true;
      focusText(editor.value, box.value);
      if (box.value?.ownerDocument.activeElement === box.value) handled = false;
    };
    return () => {
      const stepOff = !cmd.isEnabled.value || mixed.value;
      const stepReason = mixed.value ? t('disabled.mixedSize') : reason.value;
      return (
        <span {...trigger()} data-slot={props.slot}>
          <ToolButton
            label={t('toolbar.fontSizeDecrease')}
            icon="remove"
            className="pptx-toolbar__step"
            disabled={stepOff}
            disabledReason={stepReason}
            onClick={() => cmd.execute({ step: -1 })}
          />
          <input
            ref={box}
            class="pptx-toolbar__size"
            aria-label={label.value}
            {...tip.value}
            value={draft.value ?? toolbarSizeText(cmd.value.value, mixed.value)}
            {...(error.value ? { 'aria-invalid': true, 'aria-describedby': errorId } : {})}
            {...off.attrs.value}
            readonly={!cmd.isEnabled.value}
            onInput={(e: Event) => {
              draft.value = (e.target as HTMLInputElement).value;
              showError(null);
            }}
            onBlur={commit}
            onKeydown={(e: KeyboardEvent) => {
              // The key stops here: Enter must not reach the slide text, and Escape must not
              // close a More card around the box (focus goes to the text instead).
              if (e.key !== 'Enter' && e.key !== 'Escape') return;
              e.preventDefault();
              e.stopPropagation();
              leave(e.key === 'Enter');
            }}
          />
          <ToolButton
            label={t('toolbar.fontSizeIncrease')}
            icon="add"
            className="pptx-toolbar__step"
            disabled={stepOff}
            disabledReason={stepReason}
            onClick={() => cmd.execute({ step: 1 })}
          />
          {off.note()}
          {error.value ? (
            <span
              id={errorId}
              role="alert"
              class="pptx-toolbar__size-error"
              style={{ left: `${error.value.left}px`, top: `${error.value.top}px` }}
            >
              {error.value.text}
            </span>
          ) : null}
        </span>
      );
    };
  },
});
