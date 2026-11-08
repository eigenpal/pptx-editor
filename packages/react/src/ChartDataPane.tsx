/** @jsxImportSource react */
// The Chart Data tab of the side panel: the categories (or X values) of the selected chart, then
// one card per series with its name and a value for each category. A field applies when the
// user presses Enter or leaves it, as one undo step, so the slide updates beside the panel. The
// field checks and the commands come from core (`chart-pane.ts`), so the Vue twin
// (`packages/vue/src/ChartDataPane.tsx`) checks the same input with the same keys.

import { useId, useState } from 'react';
import {
  chartPaneCommand,
  chartValueText,
  nextChartCategoryLabel,
  type ChartData,
  type ChartPaneChange,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { ToolButton } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';

/**
 * A text field that applies on Enter or blur and shows the chart's value otherwise. Text that
 * the tab or the editor refuses stays in the field with its reason; Escape restores the value.
 */
function ChartField({
  label,
  visibleLabel,
  value,
  numeric,
  disabled,
  onCommit,
}: {
  label: string;
  /** The label over the field; without it, the field has only its accessible name. */
  visibleLabel?: string;
  value: string;
  numeric?: boolean;
  disabled: boolean;
  /** Applies the text: null when it applied, else the reason to show. */
  onCommit: (text: string) => string | null;
}) {
  const errorId = useId();
  const [draft, setDraft] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const reset = (): void => {
    setDraft(null);
    setError(null);
  };
  const commit = (): void => {
    if (draft === null) return;
    if (draft === value) return reset();
    const refused = onCommit(draft);
    if (refused === null) reset();
    else setError(refused);
  };
  return (
    <label className="pptx-field pptx-chart-pane__field">
      {visibleLabel ? <span className="pptx-chart-pane__label">{visibleLabel}</span> : null}
      <input
        aria-label={label}
        inputMode={numeric ? 'decimal' : undefined}
        aria-invalid={error !== null}
        aria-describedby={error !== null ? errorId : undefined}
        value={draft ?? value}
        disabled={disabled}
        // Uncommitted typing keeps Ctrl/Cmd+Z for the field's own undo (`panelHistoryKey`).
        data-draft={draft !== null}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === 'Enter') commit();
          if (e.key === 'Escape') reset();
        }}
      />
      {error !== null ? (
        <span id={errorId} className="pptx-chart-pane__error" role="alert">
          {error}
        </span>
      ) : null}
    </label>
  );
}

/** One series: its name, its values by category, and its removal. */
function SeriesCard({
  data,
  index,
  open,
  disabled,
  onToggle,
  apply,
  run,
}: {
  data: ChartData;
  index: number;
  open: boolean;
  disabled: boolean;
  onToggle: (open: boolean) => void;
  apply: (change: ChartPaneChange) => string | null;
  run: (change: ChartPaneChange) => boolean;
}) {
  const { t } = useTranslation();
  const series = data.series[index]!;
  const name = series.name || t('chartData.unnamed');
  return (
    <details
      className="pptx-chart-pane__card"
      open={open}
      onToggle={(e) => onToggle(e.currentTarget.open)}
    >
      <summary className="pptx-chart-pane__summary">{name}</summary>
      <div className="pptx-chart-pane__card-body">
        <ChartField
          label={t('chartData.seriesName', { number: index + 1 })}
          visibleLabel={t('chartData.name')}
          value={series.name}
          disabled={disabled}
          onCommit={(text) => apply({ type: 'seriesName', series: index, text })}
        />
        <div className="pptx-panel__grid">
          {data.categories.map((c, i) => (
            <ChartField
              key={i}
              label={t('chartData.cellLabel', { series: name, category: c.label })}
              visibleLabel={c.label}
              value={chartValueText(series.values[i])}
              numeric
              disabled={disabled}
              onCommit={(text) => apply({ type: 'value', series: index, category: i, text })}
            />
          ))}
        </div>
        <button
          type="button"
          className="pptx-button pptx-chart-pane__remove"
          disabled={disabled || data.series.length <= 1}
          aria-label={t('chartData.removeSeriesNamed', { name })}
          onClick={() => run({ type: 'removeSeries', series: index })}
        >
          {t('chartData.removeSeries')}
        </button>
      </div>
    </details>
  );
}

/** The Chart Data tab: the data of the selected chart, or why it shows none. */
export function ChartDataPane() {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const chart = useEditorState((s) => s.chart ?? null);
  // Every change to the document renders the tab again, which reads the chart's data again.
  useEditorState((s) => s.revision);
  const readOnly = useEditorState((s) => s.mode === 'view');
  const read = editor && chart ? editor.chartData(chart.nodeId) : null;
  const [refusal, setRefusal] = useState<string | null>(null);
  const [open, setOpen] = useState<ReadonlySet<number>>(() => new Set([0]));
  if (!chart || !read)
    return (
      <div className="pptx-empty">
        <p className="pptx-empty__title">{t('chartData.emptyTitle')}</p>
        <p className="pptx-empty__body">{t('chartData.emptyBody')}</p>
      </div>
    );
  if (!read.ok)
    return (
      <div className="pptx-empty">
        <p className="pptx-empty__title">{t('chartData.notEditableTitle')}</p>
        <p className="pptx-empty__body">{t(read.reason as TranslationKey)}</p>
      </div>
    );
  const data = read.data;
  const disabled = readOnly || !chart.editable;
  /** Applies one change: null when it applied, else the reason. */
  const apply = (change: ChartPaneChange): string | null => {
    const command = chartPaneCommand(data, change);
    if (!command.ok) return t(command.error);
    const result = editor?.exec(command.command);
    setRefusal(null);
    return !result || result.ok ? null : t(result.reason as TranslationKey);
  };
  /** Applies an add or a remove; the tab announces a refusal, as no field shows it. */
  const run = (change: ChartPaneChange): boolean => {
    const reason = apply(change);
    setRefusal(reason);
    return reason === null;
  };
  const toggle = (index: number, isOpen: boolean): void =>
    setOpen((prev) => {
      if (prev.has(index) === isOpen) return prev;
      const next = new Set(prev);
      if (isOpen) next.add(index);
      else next.delete(index);
      return next;
    });
  const xy = data.kind === 'xy';
  const note =
    data.workbook === 'kept'
      ? t('chartData.workbookKept')
      : data.workbook === 'external'
        ? t('chartData.workbookExternal')
        : null;
  return (
    <div className="pptx-chart-pane">
      <p className="pptx-chart-pane__note">{t('chartData.help')}</p>
      {note ? <p className="pptx-chart-pane__note">{note}</p> : null}
      <p className="pptx-chart-pane__error" role="alert">
        {refusal ?? ''}
      </p>
      <section className="pptx-panel__section">
        <h3 className="pptx-panel__title">
          {t(xy ? 'chartData.xValues' : 'chartData.categories')}
        </h3>
        <ul className="pptx-chart-pane__list">
          {data.categories.map((c, i) => (
            <li key={i} className="pptx-chart-pane__row">
              <ChartField
                label={t(xy ? 'chartData.xValueLabel' : 'chartData.categoryLabel', {
                  number: i + 1,
                })}
                value={c.label}
                numeric={xy}
                disabled={disabled}
                onCommit={(text) => apply({ type: 'label', category: i, text })}
              />
              <ToolButton
                label={t('chartData.removeCategoryNamed', { name: c.label })}
                icon="delete"
                disabled={disabled || data.categories.length <= 1}
                onClick={() => run({ type: 'removeCategory', category: i })}
              />
            </li>
          ))}
        </ul>
        <button
          type="button"
          className="pptx-button"
          disabled={disabled}
          onClick={() =>
            run({
              type: 'addCategory',
              label: nextChartCategoryLabel(
                data,
                t('chartData.newCategory', { number: data.categories.length + 1 })
              ),
            })
          }
        >
          {t('chartData.addCategory')}
        </button>
      </section>
      <section className="pptx-panel__section">
        <h3 className="pptx-panel__title">{t('chartData.series')}</h3>
        {data.series.map((_, j) => (
          <SeriesCard
            key={j}
            data={data}
            index={j}
            open={open.has(j)}
            disabled={disabled}
            onToggle={(isOpen) => toggle(j, isOpen)}
            apply={apply}
            run={run}
          />
        ))}
        <button
          type="button"
          className="pptx-button"
          disabled={disabled}
          onClick={() => {
            const at = data.series.length;
            if (run({ type: 'addSeries', name: t('chartData.newSeries', { number: at + 1 }) }))
              toggle(at, true);
          }}
        >
          {t('chartData.addSeries')}
        </button>
      </section>
    </div>
  );
}
