/** @jsxImportSource react */
import { useRef } from 'react';
import {
  gradientPreviewCss,
  gradientPreviewStops,
  type GradientDetail,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, useTranslation } from './context.ts';
/** Gradient overview with native stop selection and keyboard-accessible markers. @internal */
export function FormatGradientPicker({
  gradient,
  selected,
  onSelect,
}: {
  gradient: GradientDetail;
  selected: string;
  onSelect: (id: string) => void;
}) {
  const { t } = useTranslation();
  const rows = useEditorState((s) => s.themeColors);
  const root = useRef<HTMLDivElement>(null);
  const stops = gradientPreviewStops(gradient, rows);
  const background = gradientPreviewCss(stops);
  const label = (index: number) =>
    `${t('inspector.gradientStop', { index: index + 1 })} · ${Math.round(stops[index]!.position * 100)}%`;
  const move = (index: number) => {
    const stop = stops[index];
    if (!stop) return;
    onSelect(stop.id);
    const buttons = root.current?.querySelectorAll<HTMLButtonElement>('[data-format-stop-marker]');
    buttons?.[index]?.focus();
  };
  return (
    <div className="pptx-gradient-picker" ref={root}>
      <div
        className="pptx-gradient-picker__ramp"
        style={background ? { backgroundImage: background } : undefined}
        role="img"
        aria-label={t(background ? 'gradient.preview' : 'inspector.previewUnavailable')}
      />
      <div className="pptx-gradient-picker__markers" role="group" aria-label={t('gradient.stops')}>
        {stops.map((stop, index) => (
          <button
            type="button"
            key={stop.id}
            data-format-stop-marker={stop.id}
            className="pptx-gradient-picker__marker"
            style={{ left: `${stop.position * 100}%` }}
            aria-label={label(index)}
            aria-pressed={selected === stop.id}
            tabIndex={selected === stop.id ? 0 : -1}
            onClick={() => onSelect(stop.id)}
            onKeyDown={(e) => {
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
      <label className="pptx-field">
        {t('inspector.selectStop')}
        <select
          data-format-stop-select=""
          aria-label={t('inspector.selectStop')}
          value={selected}
          onChange={(e) => onSelect(e.target.value)}
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
}
