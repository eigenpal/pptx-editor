/** @jsxImportSource vue */
import { defineComponent, ref, watch, useId, type PropType } from 'vue';
import { useTranslation } from './composables.ts';
/** Enter or blur commits once; invalid values remain available for correction. @internal */
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
    onCommit: { type: Function as PropType<(value: number) => string | null>, required: true },
  },
  setup(props) {
    const { t } = useTranslation();
    const errorId = useId();
    const draft = ref<string | null>(null);
    const error = ref<string | null>(null);
    const reset = () => {
      draft.value = null;
      error.value = null;
    };
    watch(() => props.value, reset, { flush: 'sync' });
    const commit = () => {
      if (draft.value === null) return;
      const number = draft.value.trim() === '' ? NaN : Number(draft.value.replace(',', '.'));
      if (
        !Number.isFinite(number) ||
        (props.min !== undefined &&
          (number < props.min || (props.exclusiveMin && number === props.min))) ||
        (props.max !== undefined && number > props.max)
      ) {
        error.value =
          props.min !== undefined && props.max !== undefined
            ? t(props.exclusiveMin ? 'inspector.numberRangeExclusive' : 'inspector.numberRange', {
                min: props.min,
                max: props.max,
              })
            : props.min !== undefined
              ? t('inspector.numberMinimum', { min: props.min })
              : t('disabled.invalidValue');
        return;
      }
      if (props.value !== null && number === Number(props.value)) return reset();
      const refused = props.onCommit(number);
      if (refused) error.value = refused;
      else reset();
    };
    return () => (
      <label class="pptx-field">
        {props.label}
        <span class="pptx-field__control">
          <input
            inputmode="decimal"
            aria-label={props.label}
            data-format-property={props.property}
            value={draft.value ?? props.value ?? ''}
            placeholder={props.value === null && props.mixed ? t('inspector.mixed') : undefined}
            disabled={props.disabled}
            aria-invalid={!!error.value}
            aria-describedby={error.value ? errorId : undefined}
            data-draft={draft.value !== null}
            onInput={(e: Event) => {
              draft.value = (e.target as HTMLInputElement).value;
            }}
            onBlur={commit}
            onKeydown={(e: KeyboardEvent) => {
              if (e.key === 'Enter' && !e.isComposing) {
                e.preventDefault();
                commit();
              }
              if (e.key === 'Escape') {
                e.preventDefault();
                e.stopPropagation();
                reset();
              }
            }}
          />
          <span class="pptx-field__unit">{props.unit}</span>
        </span>
        {error.value ? (
          <span id={errorId} class="pptx-chart-pane__error" role="alert">
            {error.value}
          </span>
        ) : null}
      </label>
    );
  },
});
