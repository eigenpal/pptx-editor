/** @jsxImportSource react */
import { formatLineState, formatTargetKey, translateKeyOrText } from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { FormatNumberField } from './FormatNumberField.tsx';
/** Exact width in points, including a mixed selection. @internal */
export function FormatLineWidth() {
  const editor = usePptxEditor();
  const snapshot = useEditorState((s) => s);
  const { t } = useTranslation();
  const line = formatLineState(editor);
  if (!line) return null;
  const key = `${editor?.session?.id}:${formatTargetKey(snapshot)}`;
  return (
    <FormatNumberField
      key={key}
      label={t('inspector.lineWidth')}
      property="line.width"
      value={line.width === null ? null : String(Math.round(line.width * 100) / 100)}
      unit="pt"
      min={0}
      exclusiveMin
      mixed={line.widthMixed}
      max={1584}
      disabled={snapshot.mode === 'view'}
      onCommit={(points) => {
        if (!editor || `${editor.session?.id}:${formatTargetKey(editor.snapshot())}` !== key)
          return t('disabled.invalidValue');
        if (points === line.width) return null;
        const result = editor.exec({ type: 'setLineWidth', points });
        return result.ok ? null : translateKeyOrText(t, result.reason);
      }}
    />
  );
}
