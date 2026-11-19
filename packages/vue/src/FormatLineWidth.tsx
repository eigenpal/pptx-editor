/** @jsxImportSource vue */
import { defineComponent } from 'vue';
import {
  formatLineState,
  formatTargetKey,
  translateKeyOrText,
  type EditorExecOptions,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { FormatNumberField } from './FormatNumberField.tsx';
import { ToolbarSlot } from './toolbar.ts';
/**
 * The outline width: a field for an exact value in points, including a mixed selection, and
 * the preset weights of the toolbar's Border weight menu beside it. @internal
 */
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
        <div class="pptx-format-line-width">
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
            disabled={snapshot.value.mode === 'view'}
            onCommit={(points: number, options: EditorExecOptions) => {
              if (
                !editor.value ||
                `${editor.value.session?.id}:${formatTargetKey(editor.value.snapshot())}` !== key
              )
                return t('disabled.invalidValue');
              if (points === line.width) return null;
              const result = editor.value.exec({ type: 'setLineWidth', points }, options);
              return result.ok ? null : translateKeyOrText(t, result.reason);
            }}
          />
          <ToolbarSlot slotId="shape.lineWidth" />
        </div>
      );
    };
  },
});
