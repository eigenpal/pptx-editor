/** @jsxImportSource vue */
import { defineComponent } from 'vue';
import { formatLineState, formatTargetKey, translateKeyOrText } from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { FormatNumberField } from './FormatNumberField.tsx';
/** Exact width in points, including a mixed selection. @internal */
export const FormatLineWidth = defineComponent({
  name: 'FormatLineWidth',
  setup() {
    const editor = usePptxEditor();
    const snapshot = useEditorState((s) => s);
    const { t } = useTranslation();
    return () => {
      const line = formatLineState(editor.value);
      if (!line) return null;
      const key = `${editor.value?.session?.id}:${formatTargetKey(snapshot.value)}`;
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
          disabled={snapshot.value.mode === 'view'}
          onCommit={(points) => {
            if (
              !editor.value ||
              `${editor.value.session?.id}:${formatTargetKey(editor.value.snapshot())}` !== key
            )
              return t('disabled.invalidValue');
            if (points === line.width) return null;
            const result = editor.value.exec({ type: 'setLineWidth', points });
            return result.ok ? null : translateKeyOrText(t, result.reason);
          }}
        />
      );
    };
  },
});
