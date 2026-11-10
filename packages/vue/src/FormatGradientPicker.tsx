/** @jsxImportSource vue */
import { defineComponent, ref, type PropType } from 'vue';
import {
  gradientPreviewCss,
  gradientPreviewStops,
  type GradientDetail,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, useTranslation } from './composables.ts';
/** Gradient overview with native stop selection and keyboard-accessible markers. @internal */
export const FormatGradientPicker = defineComponent({
  name: 'FormatGradientPicker',
  props: {
    gradient: { type: Object as PropType<GradientDetail>, required: true },
    selected: { type: String, required: true },
    onSelect: { type: Function as PropType<(id: string) => void>, required: true },
  },
  setup(props) {
    const { t } = useTranslation();
    const rows = useEditorState((s) => s.themeColors);
    const root = ref<HTMLDivElement | null>(null);
    return () => {
      const stops = gradientPreviewStops(props.gradient, rows.value);
      const background = gradientPreviewCss(stops);
      const label = (index: number) =>
        `${t('inspector.gradientStop', { index: index + 1 })} · ${Math.round(stops[index]!.position * 100)}%`;
      const move = (index: number) => {
        const stop = stops[index];
        if (!stop) return;
        props.onSelect(stop.id);
        const buttons = root.value?.querySelectorAll<HTMLButtonElement>(
          '[data-format-stop-marker]'
        );
        buttons?.[index]?.focus();
      };
      return (
        <div class="pptx-gradient-picker" ref={root}>
          <div
            class="pptx-gradient-picker__ramp"
            style={background ? { backgroundImage: background } : undefined}
            role="img"
            aria-label={t(background ? 'gradient.preview' : 'inspector.previewUnavailable')}
          />
          <div class="pptx-gradient-picker__markers" role="group" aria-label={t('gradient.stops')}>
            {stops.map((stop, index) => (
              <button
                type="button"
                key={stop.id}
                data-format-stop-marker={stop.id}
                class="pptx-gradient-picker__marker"
                style={{ left: `${stop.position * 100}%` }}
                aria-label={label(index)}
                aria-pressed={props.selected === stop.id}
                tabindex={props.selected === stop.id ? 0 : -1}
                onClick={() => props.onSelect(stop.id)}
                onKeydown={(e: KeyboardEvent) => {
                  const next =
                    e.key === 'ArrowRight'
                      ? Math.min(stops.length - 1, index + 1)
                      : e.key === 'ArrowLeft'
                        ? Math.max(0, index - 1)
                        : e.key === 'Home'
                          ? 0
                          : e.key === 'End'
                            ? stops.length - 1
                            : null;
                  if (next !== null) {
                    e.preventDefault();
                    move(next);
                  }
                }}
              >
                <span
                  style={
                    stop.color ? { backgroundColor: stop.color, opacity: stop.opacity } : undefined
                  }
                />
              </button>
            ))}
          </div>
          <label class="pptx-field">
            {t('inspector.selectStop')}
            <select
              data-format-stop-select=""
              aria-label={t('inspector.selectStop')}
              value={props.selected}
              onChange={(e: Event) => props.onSelect((e.target as HTMLSelectElement).value)}
            >
              {stops.map((stop, index) => (
                <option key={stop.id} value={stop.id}>
                  {label(index)}
                </option>
              ))}
            </select>
          </label>
        </div>
      );
    };
  },
});
