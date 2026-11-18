/** @jsxImportSource react */
import { formatLineState, formatTargetKey, translateKeyOrText } from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { FormatNumberField } from './FormatNumberField.tsx';
import { ToolbarSlot } from './Toolbar.tsx';
/**
 * The outline width: a field for an exact value in points, including a mixed selection, and
 * the preset weights of the toolbar's Border weight menu beside it. @internal
 */
export function FormatLineWidth() {
  const editor = usePptxEditor();
  const snapshot = useEditorState((s) => s);
  const { t } = useTranslation();
  const line = formatLineState(editor);
  if (!line) return null;
  const key = `${editor?.session?.id}:${formatTargetKey(snapshot)}`;
  return (
    <div className="pptx-format-line-width">
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
        step={0.25}
        inline
        placeholder={t('inspector.noLine')}
        disabled={snapshot.mode === 'view'}
        onCommit={(points, options) => {
          if (!editor || `${editor.session?.id}:${formatTargetKey(editor.snapshot())}` !== key)
            return t('disabled.invalidValue');
          if (points === line.width) return null;
          const result = editor.exec({ type: 'setLineWidth', points }, options);
          return result.ok ? null : translateKeyOrText(t, result.reason);
        }}
      />
      <ToolbarSlot slotId="shape.lineWidth" />
    </div>
  );
}
