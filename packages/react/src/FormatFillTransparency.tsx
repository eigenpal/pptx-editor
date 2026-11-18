/** @jsxImportSource react */
import { useState } from 'react';
import {
  formatFillTransparency,
  formatTargetKey,
  translateKeyOrText,
  type EditorExecOptions,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { FormatNumberField } from './FormatNumberField.tsx';
import { useHistoryGroup } from './headless-hooks.tsx';

/**
 * The Transparency of one shape's solid fill: a slider and a percent field. The slider writes
 * while it moves, so the slide shows the value at once, and one drag is one undo step. @internal
 */
export function FormatFillTransparency() {
  const editor = usePptxEditor();
  const snapshot = useEditorState((s) => s);
  const { t } = useTranslation();
  const [error, setError] = useState<string | null>(null);
  const group = useHistoryGroup({ kind: 'range' });
  const field = formatFillTransparency(editor);
  const key = `${editor?.session?.id}:${formatTargetKey(snapshot)}`;
  const commit = (percent: number, options?: EditorExecOptions): string | null => {
    if (!editor || !field || `${editor.session?.id}:${formatTargetKey(editor.snapshot())}` !== key)
      return t('disabled.invalidValue');
    if (percent === field.value) return null;
    const result = editor.exec(
      { type: 'patchShapeFormat', target: field.target, patch: field.patch(percent) },
      options
    );
    return result.ok ? null : translateKeyOrText(t, result.reason);
  };
  if (!field) return null;
  const label = t(field.labelKey);
  const disabled = snapshot.mode === 'view';
  return (
    <div className="pptx-format-transparency" data-format-fill-transparency="">
      <input
        ref={group.ref}
        type="range"
        className="pptx-format-transparency__slider"
        data-format-slider={field.id}
        aria-label={label}
        aria-valuetext={`${field.value}${field.unit}`}
        min={field.min}
        max={field.max}
        step={1}
        value={field.value}
        disabled={disabled}
        onChange={(e) => setError(commit(Number(e.target.value), group.options()))}
      />
      <FormatNumberField
        key={key}
        label={label}
        property={field.id}
        value={String(field.value)}
        unit={field.unit}
        min={field.min}
        max={field.max}
        disabled={disabled}
        onCommit={(percent, options) => {
          setError(null);
          return commit(percent, options);
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
