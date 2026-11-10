/** @jsxImportSource vue */
import { defineComponent, ref, nextTick, type PropType } from 'vue';
import {
  chromeControlForSlot,
  colorChoiceOf,
  colorPickOf,
  formatPropertyGroups,
  formatTargetKey,
  translateKeyOrText,
  type ShapeFormatPatch,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation, useOfferedControl } from './composables.ts';
import { FormatGradientPicker } from './FormatGradientPicker.tsx';
import { FormatNumberField } from './FormatNumberField.tsx';
import { ColorField } from './dialog-color.tsx';

/** Inline numeric fields for existing gradients, shadows, and reflections. @internal */
export const FormatDetails = defineComponent({
  props: { section: { type: String as PropType<'fill' | 'effects'>, required: true } },
  setup(props) {
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const snapshot = useEditorState((s) => s);
    const offered = useOfferedControl('shape.effects', chromeControlForSlot('shape.effects'));
    const root = ref<HTMLDivElement | null>(null);
    const selectedStop = ref<string | null>(null);
    const error = ref<string | null>(null);
    return () => {
      const read = editor.value?.shapeFormat();
      if (snapshot.value.selection.kind === 'shapes' && snapshot.value.selection.nodeIds.length > 1)
        return null;
      if (!read) return null;
      if (!read.ok)
        return (
          <p class="pptx-empty__body" role="status">
            {translateKeyOrText(t, read.reason)}
          </p>
        );
      const detail = read.format;
      if (props.section === 'effects' && detail.effectsReason)
        return (
          <p class="pptx-empty__body" role="status">
            {translateKeyOrText(t, detail.effectsReason)}
          </p>
        );
      const activeStop =
        detail.gradient?.stops.find((stop) => stop.id === selectedStop.value) ??
        detail.gradient?.stops[0];
      const key = `${editor.value?.session?.id}:${formatTargetKey(snapshot.value)}`;
      const apply = (patch: ShapeFormatPatch): string | null => {
        if (
          !editor.value ||
          `${editor.value.session?.id}:${formatTargetKey(editor.value.snapshot())}` !== key
        )
          return t('disabled.invalidValue');
        const result = editor.value.exec({
          type: 'patchShapeFormat',
          target: detail.target,
          patch,
        });
        return result.ok ? null : translateKeyOrText(t, result.reason);
      };
      const disabled = snapshot.value.mode === 'view';
      const effectsAvailable =
        offered.value?.options?.some((option) => option.value === 'reflection.tightTouching') ===
        true;
      const unavailable = (id: string) =>
        !effectsAvailable &&
        (id === 'reflection' || (id === 'shadow' && detail.shadow?.kind === 'inner'));
      return (
        <div class="pptx-format-details" key={key} ref={root}>
          {props.section === 'fill' && detail.gradient && activeStop ? (
            <FormatGradientPicker
              gradient={detail.gradient}
              selected={activeStop.id}
              onSelect={(id) => {
                selectedStop.value = id;
              }}
            />
          ) : null}
          {formatPropertyGroups(detail, props.section)
            .filter((group) => group.index === undefined || group.id === activeStop?.id)
            .map((group) => (
              <fieldset
                class="pptx-format-property-group"
                key={group.id}
                data-format-stop={group.index}
              >
                <legend>
                  {t(group.labelKey, group.index ? { index: group.index } : undefined)}
                </legend>
                <div class="pptx-format-property-fields">
                  {unavailable(group.id) ? (
                    <p class="pptx-empty__body" role="status">
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
                          if (color) error.value = apply(field.patch(color));
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
                    class="pptx-button"
                    disabled={disabled || !group.remove}
                    title={!group.remove ? t('inspector.stopLimit') : undefined}
                    onClick={() => {
                      error.value = apply(group.remove!);
                      if (!error.value) {
                        const remaining = detail.gradient!.stops.filter(
                          (stop) => stop.id !== group.id
                        );
                        selectedStop.value =
                          remaining[Math.min((group.index ?? 1) - 1, remaining.length - 1)]!.id;
                        void nextTick(() =>
                          root.value
                            ?.querySelector<HTMLElement>('[data-format-stop-select]')
                            ?.focus()
                        );
                      }
                    }}
                  >
                    {t('gradient.remove')}
                  </button>
                ) : null}
              </fieldset>
            ))}
          {props.section === 'fill' && detail.gradient ? (
            <div class="pptx-format-detail-actions">
              <button
                type="button"
                class="pptx-button"
                disabled={disabled || detail.gradient.stops.length >= 10}
                title={detail.gradient.stops.length >= 10 ? t('inspector.addStopLimit') : undefined}
                onClick={() => {
                  const index = detail.gradient!.stops.findIndex((stop) => stop.position > 0.5);
                  error.value = apply({
                    kind: 'gradient',
                    addStop: { position: 0.5, color: { theme: 'accent1' }, opacity: 1 },
                  });
                  if (!error.value) {
                    const next = editor.value!.shapeFormat();
                    if (next.ok)
                      selectedStop.value =
                        next.format.gradient?.stops[
                          index < 0 ? detail.gradient!.stops.length : index
                        ]?.id ?? null;
                    void nextTick(() =>
                      root.value?.querySelector<HTMLElement>('[data-format-stop-select]')?.focus()
                    );
                  }
                }}
              >
                {t('gradient.add')}
              </button>
            </div>
          ) : null}
          {error.value ? (
            <p role="alert" class="pptx-chart-pane__error">
              {error.value}
            </p>
          ) : null}
        </div>
      );
    };
  },
});
