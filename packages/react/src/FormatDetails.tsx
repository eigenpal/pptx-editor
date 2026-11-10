/** @jsxImportSource react */
import { useState, useRef, useLayoutEffect } from 'react';
import {
  chromeControlForSlot,
  colorChoiceOf,
  colorPickOf,
  formatPropertyGroups,
  formatTargetKey,
  translateKeyOrText,
  type ShapeFormatPatch,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation, useOfferedControl } from './context.ts';
import { FormatGradientPicker } from './FormatGradientPicker.tsx';
import { FormatNumberField } from './FormatNumberField.tsx';
import { ColorField } from './dialog-color.tsx';

/** Inline numeric fields for existing gradients, shadows, and reflections. @internal */
export function FormatDetails({ section }: { section: 'fill' | 'effects' }) {
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const snapshot = useEditorState((s) => s);
  const offered = useOfferedControl('shape.effects', chromeControlForSlot('shape.effects'));
  const effectsAvailable =
    offered?.options?.some((option) => option.value === 'reflection.tightTouching') === true;
  const [error, setError] = useState<string | null>(null);
  const root = useRef<HTMLDivElement>(null);
  const [selectedStop, setSelectedStop] = useState<string | null>(null);
  const [focusStop, setFocusStop] = useState<number | null>(null);
  useLayoutEffect(() => {
    if (focusStop === null) return;
    setFocusStop(null);
    root.current?.querySelector<HTMLElement>('[data-format-stop-select]')?.focus();
  }, [focusStop]);
  const read = editor?.shapeFormat();
  if (snapshot.selection.kind === 'shapes' && snapshot.selection.nodeIds.length > 1) return null;
  if (!read) return null;
  if (!read.ok)
    return (
      <p className="pptx-empty__body" role="status">
        {translateKeyOrText(t, read.reason)}
      </p>
    );
  const detail = read.format;
  if (section === 'effects' && detail.effectsReason)
    return (
      <p className="pptx-empty__body" role="status">
        {translateKeyOrText(t, detail.effectsReason)}
      </p>
    );
  const activeStop =
    detail.gradient?.stops.find((stop) => stop.id === selectedStop) ?? detail.gradient?.stops[0];
  const key = `${editor?.session?.id}:${formatTargetKey(snapshot)}`;
  const apply = (patch: ShapeFormatPatch): string | null => {
    if (!editor || `${editor.session?.id}:${formatTargetKey(editor.snapshot())}` !== key)
      return t('disabled.invalidValue');
    const result = editor.exec({ type: 'patchShapeFormat', target: detail.target, patch });
    return result.ok ? null : translateKeyOrText(t, result.reason);
  };
  const disabled = snapshot.mode === 'view';
  const unavailable = (id: string) =>
    !effectsAvailable &&
    (id === 'reflection' || (id === 'shadow' && detail.shadow?.kind === 'inner'));
  return (
    <div className="pptx-format-details" key={key} ref={root}>
      {section === 'fill' && detail.gradient && activeStop ? (
        <FormatGradientPicker
          gradient={detail.gradient}
          selected={activeStop.id}
          onSelect={setSelectedStop}
        />
      ) : null}
      {formatPropertyGroups(detail, section)
        .filter((group) => group.index === undefined || group.id === activeStop?.id)
        .map((group) => (
          <fieldset
            className="pptx-format-property-group"
            key={group.id}
            data-format-stop={group.index}
          >
            <legend>{t(group.labelKey, group.index ? { index: group.index } : undefined)}</legend>
            <div className="pptx-format-property-fields">
              {unavailable(group.id) ? (
                <p className="pptx-empty__body" role="status">
                  {t('inspector.effectUnavailable')}
                </p>
              ) : null}
              {group.colors.map((field) => {
                const pick = field.color ? colorPickOf(field.color) : 'keep';
                return (
                  <ColorField
                    key="color"
                    name={field.id}
                    data-format-property={field.id}
                    label={t(field.labelKey)}
                    value={pick}
                    seed={pick}
                    disabled={disabled || unavailable(group.id)}
                    onChange={(next) => {
                      const color = colorChoiceOf(next);
                      if (color) setError(apply(field.patch(color)));
                    }}
                  />
                );
              })}
              {group.numbers.map((field) => (
                <FormatNumberField
                  key={field.id.split('.').at(-1)}
                  value={String(field.value)}
                  unit={field.unit}
                  property={field.id}
                  min={field.min}
                  max={field.max}
                  label={t(field.labelKey)}
                  disabled={disabled || unavailable(group.id)}
                  onCommit={(value) => apply(field.patch(value))}
                />
              ))}
            </div>
            {group.index !== undefined ? (
              <button
                type="button"
                className="pptx-button"
                disabled={disabled || !group.remove}
                title={!group.remove ? t('inspector.stopLimit') : undefined}
                onClick={() => {
                  const refusal = apply(group.remove!);
                  setError(refusal);
                  if (!refusal) {
                    const remaining = detail.gradient!.stops.filter((stop) => stop.id !== group.id);
                    setSelectedStop(
                      remaining[Math.min((group.index ?? 1) - 1, remaining.length - 1)]!.id
                    );
                    setFocusStop((group.index ?? 1) - 1);
                  }
                }}
              >
                {t('gradient.remove')}
              </button>
            ) : null}
          </fieldset>
        ))}
      {section === 'fill' && detail.gradient ? (
        <div className="pptx-format-detail-actions">
          <button
            type="button"
            className="pptx-button"
            disabled={disabled || detail.gradient.stops.length >= 10}
            title={detail.gradient.stops.length >= 10 ? t('inspector.addStopLimit') : undefined}
            onClick={() => {
              const index = detail.gradient!.stops.findIndex((stop) => stop.position > 0.5);
              const refusal = apply({
                kind: 'gradient',
                addStop: { position: 0.5, color: { theme: 'accent1' }, opacity: 1 },
              });
              setError(refusal);
              if (!refusal) {
                const next = editor!.shapeFormat();
                if (next.ok)
                  setSelectedStop(
                    next.format.gradient?.stops[index < 0 ? detail.gradient!.stops.length : index]
                      ?.id ?? null
                  );
                setFocusStop(0);
              }
            }}
          >
            {t('gradient.add')}
          </button>
        </div>
      ) : null}
      {error ? (
        <p role="alert" className="pptx-chart-pane__error">
          {error}
        </p>
      ) : null}
    </div>
  );
}
