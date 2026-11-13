/** @jsxImportSource react */
import { useCallback, useRef, useState } from 'react';
import {
  formatFillTransparency,
  formatTargetKey,
  translateKeyOrText,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { FormatNumberField } from './FormatNumberField.tsx';

/**
 * The Transparency of one shape's solid fill: a slider and a percent field. The slider shows its
 * value while it moves and writes once on release, so one drag is one undo step. @internal
 */
export function FormatFillTransparency() {
  const editor = usePptxEditor();
  const snapshot = useEditorState((s) => s);
  const { t } = useTranslation();
  const [preview, setPreview] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);
  const field = formatFillTransparency(editor);
  const key = `${editor?.session?.id}:${formatTargetKey(snapshot)}`;
  const commit = (percent: number): string | null => {
    if (!editor || !field || `${editor.session?.id}:${formatTargetKey(editor.snapshot())}` !== key)
      return t('disabled.invalidValue');
    if (percent === field.value) return null;
    const result = editor.exec({
      type: 'patchShapeFormat',
      target: field.target,
      patch: field.patch(percent),
    });
    return result.ok ? null : translateKeyOrText(t, result.reason);
  };
  const latest = useRef(commit);
  latest.current = commit;
  // React's `onChange` fires on every `input` event; the native `change` event marks the release.
  const attached = useRef<(() => void) | null>(null);
  const slider = useCallback((el: HTMLInputElement | null) => {
    attached.current?.();
    attached.current = null;
    if (!el) return;
    const release = () => {
      setPreview(null);
      setError(latest.current(Number(el.value)));
    };
    el.addEventListener('change', release);
    attached.current = () => el.removeEventListener('change', release);
  }, []);
  if (!field) return null;
  const label = t(field.labelKey);
  const value = preview ?? field.value;
  const disabled = snapshot.mode === 'view';
  return (
    <div className="pptx-format-transparency" data-format-fill-transparency="">
      <input
        ref={slider}
        type="range"
        className="pptx-format-transparency__slider"
        data-format-slider={field.id}
        aria-label={label}
        aria-valuetext={`${value}${field.unit}`}
        min={field.min}
        max={field.max}
        step={1}
        value={value}
        disabled={disabled}
        onChange={(e) => setPreview(Number(e.target.value))}
        onBlur={() => setPreview(null)}
      />
      <FormatNumberField
        key={key}
        label={label}
        property={field.id}
        value={String(value)}
        unit={field.unit}
        min={field.min}
        max={field.max}
        disabled={disabled}
        onCommit={(percent) => {
          setError(null);
          return commit(percent);
        }}
      />
      {error ? (
        <span className="pptx-chart-pane__error" role="alert">
          {error}
        </span>
      ) : null}
    </div>
  );
}
