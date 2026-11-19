/** @jsxImportSource react */
import { useId } from 'react';
import {
  parseFormatNumber,
  sameFormatNumber,
  type EditorExecOptions,
} from '@pptx-editor.dev/core/editor';
import { usePptxEditor, useTranslation } from './context.ts';
import { useLiveField } from './live-field.ts';

/** Shared inspector numeric lifecycle. Null displays Mixed without inventing a value. @internal */
export interface FormatNumberFieldProps {
  readonly label: string;
  readonly value: string | null;
  readonly unit: string;
  readonly property?: string;
  readonly disabled?: boolean;
  readonly min?: number;
  readonly exclusiveMin?: boolean;
  readonly mixed?: boolean;
  readonly max?: number;
  /** The Arrow Up and Arrow Down increment. Shift multiplies it by 10. */
  readonly step?: number;
  /** Show the label beside the field, as the other rows of a format section do. */
  readonly inline?: boolean;
  /** The hint for a field without a value that is not a mixed selection. */
  readonly placeholder?: string;
  /** Write one value. Pass `options` to `exec` so that one edit is one undo step. */
  readonly onCommit: (value: number, options: EditorExecOptions) => string | null;
}

/**
 * Each valid keystroke writes at once, so the slide shows the value while the user types.
 * Arrow Up and Arrow Down step the value. Enter or blur ends the edit, and Escape restores the
 * start value. One edit is one undo step. Invalid values remain available for correction.
 * @internal
 */
export function FormatNumberField({
  label,
  value,
  unit,
  property,
  disabled,
  min,
  exclusiveMin,
  mixed = true,
  max,
  step = 1,
  inline,
  placeholder,
  onCommit,
}: FormatNumberFieldProps) {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const errorId = useId();
  const field = useLiveField(
    {
      editor,
      value: () => value,
      same: sameFormatNumber,
      write(text, options, final) {
        const parsed = parseFormatNumber(text, { min, max, exclusiveMin });
        if (!parsed.ok)
          return final || !parsed.incomplete ? t(parsed.key, parsed.params) : undefined;
        if (value !== null && parsed.value === Number(value)) return null;
        return onCommit(parsed.value, options);
      },
    },
    value
  );
  const nudge = (direction: 1 | -1, large: boolean) => {
    const parsed = parseFormatNumber(field.draft ?? value ?? '', {});
    if (!parsed.ok) return;
    let next = Math.round((parsed.value + direction * step * (large ? 10 : 1)) * 1000) / 1000;
    if (max !== undefined) next = Math.min(max, next);
    if (min !== undefined) next = Math.max(min, next);
    if (exclusiveMin && min !== undefined && next <= min) return;
    field.input(String(next));
  };
  const error = field.error;
  return (
    <label className={inline ? 'pptx-field pptx-field--inline' : 'pptx-field'}>
      {inline ? <span className="pptx-field__label">{label}</span> : label}
      <span className="pptx-field__control">
        <input
          inputMode="decimal"
          aria-label={label}
          data-format-property={property}
          value={field.draft ?? value ?? ''}
          placeholder={value === null ? (mixed ? t('inspector.mixed') : placeholder) : undefined}
          disabled={disabled}
          aria-invalid={!!error}
          aria-describedby={error ? errorId : undefined}
          data-draft={field.draft !== null}
          onChange={(e) =>
            field.input(e.target.value, (e.nativeEvent as InputEvent).isComposing === true)
          }
          onCompositionEnd={(e) => field.input(e.currentTarget.value)}
          onBlur={() => field.commit()}
          onKeyDown={(e) => {
            if (e.nativeEvent.isComposing) return;
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
        <span className="pptx-field__unit">{unit}</span>
      </span>
      {error ? (
        <span id={errorId} className="pptx-chart-pane__error" role="alert">
          {error}
        </span>
      ) : null}
    </label>
  );
}
