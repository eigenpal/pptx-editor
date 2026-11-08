/** @jsxImportSource react */
// The Chart tab of the side panel: the type of the selected chart, its elements (title, legend,
// data labels, axes, gridlines, axis titles), and its colors; then its categories (or X values)
// and one card per series with its name and a value for each category. A field applies when the
// user presses Enter or leaves it, and a check box, list, or color when it changes, each as one
// undo step, so the slide updates beside the panel. The field checks and the commands come from
// core (`chart-pane.ts`), so the Vue twin (`packages/vue/src/ChartDataPane.tsx`) checks the same
// input with the same keys.

import { useId, useRef, useState } from 'react';
import {
  CHART_LEGEND_CHOICES,
  CHART_TYPE_CHOICES,
  chartAxisTitleLabels,
  chartSeriesNote,
  changeChartTypeKeeping,
  chartTypeHint,
  chartColorCommand,
  chartColorPick,
  chartPaneCommand,
  chartPaneDraftKey,
  chartPaneDraftContinues,
  chartTitleCommand,
  chartValueText,
  focusChartRow,
  colorPickOf,
  nextChartCategoryLabel,
  renderedColorOf,
  type ChartData,
  type ChartFormatChange,
  type ChartType,
  type ChartPaneChange,
  type EditorCommand,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { ToolButton } from './chrome.tsx';
import { ColorField } from './dialog-color.tsx';
import { SelectField } from './dialog-form.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';

/**
 * A text field that applies on Enter or blur and shows the chart's value otherwise. Text that
 * the tab or the editor refuses stays in the field with its reason; Escape restores the value.
 */
function ChartField({
  draftKey,
  label,
  visibleLabel,
  value,
  numeric,
  disabled,
  onCommit,
}: {
  draftKey: string;
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
  const [scope, setScope] = useState(draftKey);
  if (scope !== draftKey) {
    setScope(draftKey);
    setDraft(null);
    setError(null);
  }
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
  draftKey,
  index,
  open,
  disabled,
  onToggle,
  apply,
  run,
}: {
  data: ChartData;
  draftKey: string;
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
      data-chart-row="series"
      className="pptx-chart-pane__card"
      open={open}
      onToggle={(e) => onToggle(e.currentTarget.open)}
    >
      <summary className="pptx-chart-pane__summary">
        <span className="pptx-chart-pane__label">{name}</span>
      </summary>
      <div className="pptx-chart-pane__card-body">
        <ChartField
          draftKey={draftKey}
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
              draftKey={draftKey}
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
          onClick={(e) => {
            const pane = e.currentTarget.closest('.pptx-chart-pane');
            if (run({ type: 'removeSeries', series: index })) focusChartRow(pane, 'series', index);
          }}
        >
          {t('chartData.removeSeries')}
        </button>
      </div>
    </details>
  );
}

/** The most color fields the tab shows: a pie chart can have thousands of points. */
const MAX_COLOR_FIELDS = 100;

/** A check box row of the Chart tab. */
function ChartCheck({
  label,
  checked,
  disabled,
  onChange,
}: {
  label: string;
  checked: boolean;
  disabled: boolean;
  onChange: (checked: boolean) => void;
}) {
  return (
    <label className="pptx-radio">
      <input
        type="checkbox"
        checked={checked}
        disabled={disabled}
        onChange={(e) => onChange(e.currentTarget.checked)}
      />
      {label}
    </label>
  );
}

/**
 * The chart's type, its elements (title, legend, data labels, axes, gridlines, and axis
 * titles), and the color of each series, or of each point of a pie chart.
 */
function ChartFormatSection({
  nodeId,
  draftKey,
  disabled,
  seriesCount,
  run,
}: {
  nodeId: string;
  draftKey: string;
  disabled: boolean;
  /** The number of series in the chart's data. */
  seriesCount: number;
  /** Runs a command: null when it ran (or there was none), else the reason. */
  run: (command: EditorCommand | null) => string | null;
}) {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const rows = useEditorState((s) => s.themeColors);
  // What the last type change dropped, for the chart it changed.
  const [typeNote, setTypeNote] = useState<{ nodeId: string; key: TranslationKey } | null>(null);
  const read = editor ? editor.chartFormat(nodeId) : null;
  if (!read?.ok) return null;
  const format = read.format;
  const changeType = (chartType: ChartType) => {
    if (!editor) return;
    const { reason, note } = changeChartTypeKeeping(editor, nodeId, chartType);
    // A refusal runs the command again through the tab, which announces it.
    if (reason) run({ type: 'setChartType', nodeId, chartType });
    else run(null);
    setTypeNote(note ? { nodeId, key: note } : null);
  };
  const notes = [
    typeNote?.nodeId === nodeId ? typeNote.key : chartTypeHint(format),
    chartSeriesNote(format, seriesCount),
  ].filter((k) => k !== null);
  const axisLabels = chartAxisTitleLabels(format);
  const set = (change: ChartFormatChange) => run({ type: 'setChartFormat', format: change });
  return (
    <>
      <section className="pptx-panel__section">
        <SelectField
          name="chartType"
          label={t('chartFormat.type')}
          value={format.type ?? ''}
          mixed={t('chartFormat.types.other')}
          options={CHART_TYPE_CHOICES.map((v) => [v, t(`chartFormat.types.${v}`)] as const)}
          disabled={disabled}
          onChange={changeType}
        />
        <div role="status">
          {notes.map((key) => (
            <p key={key} className="pptx-chart-pane__note">
              {t(key)}
            </p>
          ))}
        </div>
      </section>
      <section className="pptx-panel__section">
        <h3 className="pptx-panel__title">{t('chartFormat.elements')}</h3>
        <ChartCheck
          label={t('chartFormat.showTitle')}
          checked={format.title !== null}
          disabled={disabled}
          onChange={(on) => set({ title: on ? t('chartFormat.defaultTitle') : null })}
        />
        {format.title !== null ? (
          <div className="pptx-chart-pane__sub">
            <ChartField
              draftKey={draftKey}
              label={t('chartFormat.titleText')}
              visibleLabel={t('chartFormat.titleText')}
              value={format.title}
              disabled={disabled}
              onCommit={(text) => run(chartTitleCommand(format, 'title', text))}
            />
          </div>
        ) : null}
        <SelectField
          name="chartLegend"
          label={t('chartFormat.legend')}
          value={format.legend}
          options={CHART_LEGEND_CHOICES.map(
            (v) => [v, t(`chartFormat.legendPositions.${v}`)] as const
          )}
          disabled={disabled}
          onChange={(legend) => set({ legend })}
        />
        <ChartCheck
          label={t('chartFormat.dataLabels')}
          checked={format.dataLabels}
          disabled={disabled}
          onChange={(dataLabels) => set({ dataLabels })}
        />
        {format.axes !== null ? (
          <>
            <ChartCheck
              label={t('chartFormat.axes')}
              checked={format.axes}
              disabled={disabled}
              onChange={(axes) => set({ axes })}
            />
            <ChartCheck
              label={t('chartFormat.gridlines')}
              checked={format.gridlines === true}
              disabled={disabled}
              onChange={(gridlines) => set({ gridlines })}
            />
            <div className="pptx-panel__grid">
              <ChartField
                draftKey={draftKey}
                label={t(axisLabels.category)}
                visibleLabel={t(axisLabels.category)}
                value={format.axisTitles?.category ?? ''}
                disabled={disabled || !format.axes}
                onCommit={(text) => run(chartTitleCommand(format, 'categoryAxisTitle', text))}
              />
              <ChartField
                draftKey={draftKey}
                label={t(axisLabels.value)}
                visibleLabel={t(axisLabels.value)}
                value={format.axisTitles?.value ?? ''}
                disabled={disabled || !format.axes}
                onCommit={(text) => run(chartTitleCommand(format, 'valueAxisTitle', text))}
              />
            </div>
            <p className="pptx-chart-pane__note">
              {t(format.axes ? 'chartFormat.axisTitleHelp' : 'chartFormat.axesOffHelp')}
            </p>
          </>
        ) : null}
      </section>
      <section className="pptx-panel__section">
        <h3 className="pptx-panel__title">{t('chartFormat.colors')}</h3>
        {format.colors.slice(0, MAX_COLOR_FIELDS).map((entry, i) => (
          <ColorField
            key={i}
            name={`chartColor${i}`}
            label={t('chartFormat.seriesColor', { name: entry.name || t('chartData.unnamed') })}
            value={chartColorPick(entry)}
            seed={undefined}
            automatic={renderedColorOf(colorPickOf(entry.automatic), rows)}
            disabled={disabled}
            onChange={(pick) => run(chartColorCommand(i, pick))}
          />
        ))}
      </section>
    </>
  );
}

/** The Chart tab: the type, elements, colors, and data of the selected chart, or why it shows none. */
export function ChartDataPane() {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const chart = useEditorState((s) => s.chart ?? null);
  // Every change to the document renders the tab again, which reads the chart's data again.
  useEditorState((s) => s.revision);
  const readOnly = useEditorState((s) => s.mode === 'view');
  const read = editor && chart ? editor.chartData(chart.nodeId) : null;
  const [refusal, setRefusal] = useState<string | null>(null);
  const pane = useRef<HTMLDivElement>(null);
  const [open, setOpen] = useState<ReadonlySet<number>>(() => new Set([0]));
  const [draftScope, setDraftScope] = useState({ key: '', reset: 0 });
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
  const scopeKey = chartPaneDraftKey(editor!.snapshot(), data);
  const nextScope =
    scopeKey === draftScope.key
      ? draftScope
      : {
          key: scopeKey,
          reset:
            draftScope.reset +
            (draftScope.key && chartPaneDraftContinues(draftScope.key, scopeKey) ? 0 : 1),
        };
  if (nextScope !== draftScope) setDraftScope(nextScope);
  const draftKey = String(nextScope.reset);
  const disabled = readOnly || !chart.editable;
  /** Applies one change: null when it applied, else the reason. */
  const apply = (change: ChartPaneChange): string | null => {
    const command = chartPaneCommand(data, change);
    if (!command.ok) return t(command.error);
    const result = editor?.exec(command.command);
    setRefusal(null);
    return !result || result.ok ? null : t(result.reason as TranslationKey);
  };
  /** Runs a command of the type, elements, or colors; the tab announces a refusal. */
  const runFormat = (command: EditorCommand | null): string | null => {
    if (!command) {
      setRefusal(null);
      return null;
    }
    const result = editor?.exec(command);
    const reason = !result || result.ok ? null : t(result.reason as TranslationKey);
    setRefusal(reason);
    return reason;
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
    <div ref={pane} className="pptx-chart-pane">
      <p className="pptx-chart-pane__error" role="alert">
        {refusal ?? ''}
      </p>
      <ChartFormatSection
        nodeId={chart.nodeId}
        draftKey={draftKey}
        disabled={readOnly}
        seriesCount={data.series.length}
        run={runFormat}
      />
      <section className="pptx-panel__section">
        <h3 className="pptx-panel__title">
          {t(xy ? 'chartData.xValues' : 'chartData.categories')}
        </h3>
        <p className="pptx-chart-pane__note">{t('chartData.help')}</p>
        {note ? <p className="pptx-chart-pane__note">{note}</p> : null}
        <ul className="pptx-chart-pane__list">
          {data.categories.map((c, i) => (
            <li key={i} data-chart-row="category" className="pptx-chart-pane__row">
              <ChartField
                draftKey={draftKey}
                label={t(xy ? 'chartData.xValueLabel' : 'chartData.categoryLabel', {
                  number: i + 1,
                })}
                value={c.label}
                numeric={xy}
                disabled={disabled}
                onCommit={(text) => apply({ type: 'label', category: i, text })}
              />
              <ToolButton
                label={t(xy ? 'chartData.removeXValueNamed' : 'chartData.removeCategoryNamed', {
                  name: c.label,
                })}
                icon="delete"
                disabled={disabled || data.categories.length <= 1}
                onClick={() => {
                  if (run({ type: 'removeCategory', category: i }))
                    focusChartRow(pane.current, 'category', i);
                }}
              />
            </li>
          ))}
        </ul>
        <button
          type="button"
          className="pptx-button"
          disabled={disabled}
          onClick={() => {
            const label = nextChartCategoryLabel(
              data,
              t('chartData.newCategory', { number: data.categories.length + 1 })
            );
            if (run({ type: 'addCategory', label }))
              focusChartRow(pane.current, 'category', data.categories.length, true);
          }}
        >
          {t(xy ? 'chartData.addXValue' : 'chartData.addCategory')}
        </button>
      </section>
      <section className="pptx-panel__section">
        <h3 className="pptx-panel__title">{t('chartData.series')}</h3>
        {data.series.map((_, j) => (
          <SeriesCard
            key={j}
            data={data}
            draftKey={draftKey}
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
            if (run({ type: 'addSeries', name: t('chartData.newSeries', { number: at + 1 }) })) {
              toggle(at, true);
              focusChartRow(pane.current, 'series', at, true);
            }
          }}
        >
          {t('chartData.addSeries')}
        </button>
      </section>
    </div>
  );
}
