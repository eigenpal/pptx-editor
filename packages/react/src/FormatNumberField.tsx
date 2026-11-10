/** @jsxImportSource react */
import { useId, useRef, useState } from 'react';
import { useTranslation } from './context.ts';

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
  readonly onCommit: (value: number) => string | null;
}
/** Enter or blur commits once; invalid values remain available for correction. @internal */
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
  onCommit,
}: FormatNumberFieldProps) {
  const { t } = useTranslation();
  const errorId = useId();
  const [draft, setDraft] = useState<string | null>(null);
  const pending = useRef<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [source, setSource] = useState(value);
  const reset = () => {
    pending.current = null;
    setDraft(null);
    setError(null);
  };
  if (source !== value) {
    setSource(value);
    reset();
  }
  const commit = () => {
    if (pending.current === null) return;
    const text = pending.current;
    const number = text.trim() === '' ? NaN : Number(text.replace(',', '.'));
    if (
      !Number.isFinite(number) ||
      (min !== undefined && (number < min || (exclusiveMin && number === min))) ||
      (max !== undefined && number > max)
    ) {
      setError(
        min !== undefined && max !== undefined
          ? t(exclusiveMin ? 'inspector.numberRangeExclusive' : 'inspector.numberRange', {
              min,
              max,
            })
          : min !== undefined
            ? t('inspector.numberMinimum', { min })
            : t('disabled.invalidValue')
      );
      return;
    }
    if (value !== null && number === Number(value)) return reset();
    const refused = onCommit(number);
    if (refused) setError(refused);
    else reset();
  };
  return (
    <label className="pptx-field">
      {label}
      <span className="pptx-field__control">
        <input
          inputMode="decimal"
          aria-label={label}
          data-format-property={property}
          value={draft ?? value ?? ''}
          placeholder={value === null && mixed ? t('inspector.mixed') : undefined}
          disabled={disabled}
          aria-invalid={!!error}
          aria-describedby={error ? errorId : undefined}
          data-draft={draft !== null}
          onChange={(e) => {
            pending.current = e.target.value;
            setDraft(e.target.value);
          }}
          onBlur={commit}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.nativeEvent.isComposing) {
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
