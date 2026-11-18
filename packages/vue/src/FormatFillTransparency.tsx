/** @jsxImportSource vue */
import { defineComponent, ref } from 'vue';
import {
  formatFillTransparency,
  formatTargetKey,
  translateKeyOrText,
  type EditorExecOptions,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { FormatNumberField } from './FormatNumberField.tsx';
import { useHistoryGroup } from './headless-hooks.ts';

/**
 * The Transparency of one shape's solid fill: a slider and a percent field. The slider writes
 * while it moves, so the slide shows the value at once, and one drag is one undo step. @internal
 */
export const FormatFillTransparency = defineComponent({
  name: 'FormatFillTransparency',
  setup() {
    const editor = usePptxEditor();
    const snapshot = useEditorState((s) => s);
    const { t } = useTranslation();
    const group = useHistoryGroup({ kind: 'range' });
    const error = ref<string | null>(null);
    return () => {
      const field = formatFillTransparency(editor.value);
      if (!field) return null;
      const key = `${editor.value?.session?.id}:${formatTargetKey(snapshot.value)}`;
      const commit = (percent: number, options?: EditorExecOptions): string | null => {
        const instance = editor.value;
        if (!instance || `${instance.session?.id}:${formatTargetKey(instance.snapshot())}` !== key)
          return t('disabled.invalidValue');
        if (percent === field.value) return null;
        const result = instance.exec(
          { type: 'patchShapeFormat', target: field.target, patch: field.patch(percent) },
          options
        );
        return result.ok ? null : translateKeyOrText(t, result.reason);
      };
      const label = t(field.labelKey);
      const value = field.value;
      const disabled = snapshot.value.mode === 'view';
      return (
        <div class="pptx-format-transparency" data-format-fill-transparency="">
          <input
            type="range"
            class="pptx-format-transparency__slider"
            data-format-slider={field.id}
            aria-label={label}
            aria-valuetext={`${value}${field.unit}`}
            min={field.min}
            max={field.max}
            step={1}
            value={value}
            disabled={disabled}
            ref={group.ref}
            onInput={(e: Event) => {
              error.value = commit(Number((e.target as HTMLInputElement).value), group.options());
            }}
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
            onCommit={(percent: number, options: EditorExecOptions) => {
              error.value = null;
              return commit(percent, options);
            }}
          />
          {error.value ? (
            <span class="pptx-chart-pane__error" role="alert">
              {error.value}
            </span>
          ) : null}
        </div>
      );
    };
  },
});
