/** @jsxImportSource vue */
import { defineComponent, useId, type PropType } from 'vue';
import {
  parseFormatNumber,
  sameFormatNumber,
  type EditorExecOptions,
} from '@pptx-editor.dev/core/editor';
import { usePptxEditor, useTranslation } from './composables.ts';
import { useLiveField } from './live-field.ts';

/**
 * Each valid keystroke writes at once, so the slide shows the value while the user types.
 * Arrow Up and Arrow Down step the value. Enter or blur ends the edit, and Escape restores the
 * start value. One edit is one undo step. Invalid values remain available for correction.
 * @internal
 */
export const FormatNumberField = defineComponent({
  name: 'FormatNumberField',
  props: {
    label: { type: String, required: true },
    value: { type: String as PropType<string | null>, default: null },
    unit: { type: String, required: true },
    property: String,
    disabled: Boolean,
    min: Number,
    exclusiveMin: Boolean,
    mixed: { type: Boolean, default: true },
    max: Number,
    /** The Arrow Up and Arrow Down increment. Shift multiplies it by 10. */
    step: { type: Number, default: 1 },
    /** Show the label beside the field, as the other rows of a format section do. */
    inline: Boolean,
    onCommit: {
      type: Function as PropType<(value: number, options: EditorExecOptions) => string | null>,
      required: true,
    },
  },
  setup(props) {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const errorId = useId();
    const field = useLiveField(
      () => ({
        editor: editor.value,
        value: () => props.value,
        same: sameFormatNumber,
        write(text, options, final) {
          const parsed = parseFormatNumber(text, {
            min: props.min,
            max: props.max,
            exclusiveMin: props.exclusiveMin,
          });
          if (!parsed.ok)
            return final || !parsed.incomplete ? t(parsed.key, parsed.params) : undefined;
          if (props.value !== null && parsed.value === Number(props.value)) return null;
          return props.onCommit(parsed.value, options);
        },
      }),
      () => props.value
    );
    const nudge = (direction: 1 | -1, large: boolean) => {
      const parsed = parseFormatNumber(field.draft ?? props.value ?? '', {});
      if (!parsed.ok) return;
      const delta = direction * props.step * (large ? 10 : 1);
      let next = Math.round((parsed.value + delta) * 1000) / 1000;
      if (props.max !== undefined) next = Math.min(props.max, next);
      if (props.min !== undefined) next = Math.max(props.min, next);
      if (props.exclusiveMin && props.min !== undefined && next <= props.min) return;
      field.input(String(next));
    };
    return () => {
      const error = field.error;
      return (
        <label class={props.inline ? 'pptx-field pptx-field--inline' : 'pptx-field'}>
          {props.inline ? <span class="pptx-field__label">{props.label}</span> : props.label}
          <span class="pptx-field__control">
            <input
              inputmode="decimal"
              aria-label={props.label}
              data-format-property={props.property}
              value={field.draft ?? props.value ?? ''}
              placeholder={props.value === null && props.mixed ? t('inspector.mixed') : undefined}
              disabled={props.disabled}
              aria-invalid={!!error}
              aria-describedby={error ? errorId : undefined}
              data-draft={field.draft !== null}
              onInput={(e: Event) =>
                field.input(
                  (e.target as HTMLInputElement).value,
                  (e as InputEvent).isComposing === true
                )
              }
              onCompositionend={(e: CompositionEvent) =>
                field.input((e.target as HTMLInputElement).value)
              }
              onBlur={() => field.commit()}
              onKeydown={(e: KeyboardEvent) => {
                if (e.isComposing) return;
                if (e.key === 'Enter') {
                  e.preventDefault();
                  field.commit();
                } else if (e.key === 'Escape') {
                  e.preventDefault();
                  e.stopPropagation();
                  field.cancel();
                } else if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
                  e.preventDefault();
                  nudge(e.key === 'ArrowUp' ? 1 : -1, e.shiftKey);
                }
              }}
            />
            <span class="pptx-field__unit">{props.unit}</span>
          </span>
          {error ? (
            <span id={errorId} class="pptx-chart-pane__error" role="alert">
              {error}
            </span>
          ) : null}
        </label>
      );
    };
  },
});
