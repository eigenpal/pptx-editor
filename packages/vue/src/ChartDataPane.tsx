/** @jsxImportSource vue */
// The Chart tab of the side panel: the Vue twin of the React adapter's `ChartDataPane.tsx`, with
// the same markup, classes, and ARIA, on the same core helpers (`chart-pane.ts`): the chart's
// type, elements, and colors, then its data. A field applies when the user presses Enter or
// leaves it, and a check box, list, or color when it changes, each as one undo step.

import { defineComponent, ref, useId, watch, type PropType } from 'vue';
import {
  CHART_LEGEND_CHOICES,
  CHART_TYPE_CHOICES,
  chartAxisTitleLabels,
  chartSeriesNote,
  chartTypeChangeNote,
  chartTypeHint,
  chartColorCommand,
  chartColorPick,
  chartPaneCommand,
  chartPaneDraftKey,
  chartPaneDraftContinues,
  chartTitleCommand,
  chartValueText,
  colorPickOf,
  nextChartCategoryLabel,
  renderedColorOf,
  type ChartData,
  type ChartFormatChange,
  type ChartLegendPosition,
  type ChartPaneChange,
  type ChartType,
  type ColorPick,
  type EditorCommand,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { ToolButton } from './chrome.tsx';
import { ColorField } from './dialog-color.tsx';
import { SelectField } from './dialog-form.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';

/**
 * A text field that applies on Enter or blur and shows the chart's value otherwise. Text that
 * the tab or the editor refuses stays in the field with its reason; Escape restores the value.
 */
const ChartField = defineComponent({
  name: 'PptxChartDataField',
  props: {
    draftKey: { type: String, required: true },
    label: { type: String, required: true },
    visibleLabel: { type: String, default: undefined },
    value: { type: String, required: true },
    numeric: { type: Boolean, default: false },
    disabled: { type: Boolean, required: true },
    onCommit: { type: Function as PropType<(text: string) => string | null>, required: true },
  },
  setup(props) {
    const errorId = useId();
    const draft = ref<string | null>(null);
    const error = ref<string | null>(null);
    const reset = (): void => {
      draft.value = null;
      error.value = null;
    };
    watch(() => props.draftKey, reset, { flush: 'sync' });
    const commit = (): void => {
      const d = draft.value;
      if (d === null) return;
      if (d === props.value) return reset();
      const refused = props.onCommit(d);
      if (refused === null) reset();
      else error.value = refused;
    };
    return () => (
      <label class="pptx-field pptx-chart-pane__field">
        {props.visibleLabel ? (
          <span class="pptx-chart-pane__label">{props.visibleLabel}</span>
        ) : null}
        <input
          aria-label={props.label}
          inputmode={props.numeric ? 'decimal' : undefined}
          aria-invalid={error.value !== null}
          aria-describedby={error.value !== null ? errorId : undefined}
          value={draft.value ?? props.value}
          disabled={props.disabled}
          data-draft={draft.value !== null}
          onInput={(e: Event) => {
            draft.value = (e.target as HTMLInputElement).value;
          }}
          onBlur={commit}
          onKeydown={(e: KeyboardEvent) => {
            if (e.key === 'Enter') commit();
            if (e.key === 'Escape') reset();
          }}
        />
        {error.value !== null ? (
          <span id={errorId} class="pptx-chart-pane__error" role="alert">
            {error.value}
          </span>
        ) : null}
      </label>
    );
  },
});

/** One series: its name, its values by category, and its removal. */
const SeriesCard = defineComponent({
  name: 'PptxChartSeriesCard',
  props: {
    data: { type: Object as PropType<ChartData>, required: true },
    draftKey: { type: String, required: true },
    index: { type: Number, required: true },
    open: { type: Boolean, required: true },
    disabled: { type: Boolean, required: true },
    onToggle: { type: Function as PropType<(open: boolean) => void>, required: true },
    apply: {
      type: Function as PropType<(change: ChartPaneChange) => string | null>,
      required: true,
    },
    run: { type: Function as PropType<(change: ChartPaneChange) => boolean>, required: true },
  },
  setup(props) {
    const { t } = useTranslation();
    return () => {
      const { data, index, disabled } = props;
      const series = data.series[index]!;
      const name = series.name || t('chartData.unnamed');
      return (
        <details
          class="pptx-chart-pane__card"
          open={props.open}
          onToggle={(e: Event) => props.onToggle((e.currentTarget as HTMLDetailsElement).open)}
        >
          <summary class="pptx-chart-pane__summary">{name}</summary>
          <div class="pptx-chart-pane__card-body">
            <ChartField
              draftKey={props.draftKey}
              label={t('chartData.seriesName', { number: index + 1 })}
              visibleLabel={t('chartData.name')}
              value={series.name}
              disabled={disabled}
              onCommit={(text: string) => props.apply({ type: 'seriesName', series: index, text })}
            />
            <div class="pptx-panel__grid">
              {data.categories.map((c, i) => (
                <ChartField
                  key={i}
                  draftKey={props.draftKey}
                  label={t('chartData.cellLabel', { series: name, category: c.label })}
                  visibleLabel={c.label}
                  value={chartValueText(series.values[i])}
                  numeric
                  disabled={disabled}
                  onCommit={(text: string) =>
                    props.apply({ type: 'value', series: index, category: i, text })
                  }
                />
              ))}
            </div>
            <button
              type="button"
              class="pptx-button pptx-chart-pane__remove"
              disabled={disabled || data.series.length <= 1}
              aria-label={t('chartData.removeSeriesNamed', { name })}
              onClick={() => props.run({ type: 'removeSeries', series: index })}
            >
              {t('chartData.removeSeries')}
            </button>
          </div>
        </details>
      );
    };
  },
});

/** The most color fields the tab shows: a pie chart can have thousands of points. */
const MAX_COLOR_FIELDS = 100;

/** A check box row of the Chart tab. */
const ChartCheck = defineComponent({
  name: 'PptxChartCheck',
  props: {
    label: { type: String, required: true },
    checked: { type: Boolean, required: true },
    disabled: { type: Boolean, required: true },
    onChange: { type: Function as PropType<(checked: boolean) => void>, required: true },
  },
  setup(props) {
    return () => (
      <label class="pptx-radio">
        <input
          type="checkbox"
          checked={props.checked}
          disabled={props.disabled}
          onChange={(e: Event) => props.onChange((e.currentTarget as HTMLInputElement).checked)}
        />
        {props.label}
      </label>
    );
  },
});

/**
 * The chart's type, its elements (title, legend, data labels, axes, gridlines, and axis
 * titles), and the color of each series, or of each point of a pie chart.
 */
const ChartFormatSection = defineComponent({
  name: 'PptxChartFormatSection',
  props: {
    nodeId: { type: String, required: true },
    draftKey: { type: String, required: true },
    disabled: { type: Boolean, required: true },
    /** The number of series in the chart's data. */
    seriesCount: { type: Number, required: true },
    /** Runs a command: null when it ran (or there was none), else the reason. */
    run: {
      type: Function as PropType<(command: EditorCommand | null) => string | null>,
      required: true,
    },
  },
  setup(props) {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const rows = useEditorState((s) => s.themeColors);
    // Every change to the document renders the section again, which reads the chart again.
    const revision = useEditorState((s) => s.revision);
    // What the last type change dropped, for the chart it changed.
    const typeNote = ref<{ nodeId: string; key: TranslationKey } | null>(null);
    return () => {
      void revision.value;
      const read = editor.value ? editor.value.chartFormat(props.nodeId) : null;
      if (!read?.ok) return null;
      const format = read.format;
      const { draftKey, disabled, run } = props;
      const axisLabels = chartAxisTitleLabels(format);
      const set = (change: ChartFormatChange) => run({ type: 'setChartFormat', format: change });
      const nodeId = props.nodeId;
      const changeType = (chartType: ChartType) => {
        const reason = run({ type: 'setChartType', chartType });
        const after = reason === null && editor.value ? editor.value.chartFormat(nodeId) : null;
        const key = after?.ok ? chartTypeChangeNote(format, after.format) : null;
        typeNote.value = key ? { nodeId, key } : null;
      };
      const notes = [
        typeNote.value?.nodeId === nodeId ? typeNote.value.key : chartTypeHint(format),
        chartSeriesNote(format, props.seriesCount),
      ].filter((k) => k !== null);
      return (
        <>
          <section class="pptx-panel__section">
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
                <p key={key} class="pptx-chart-pane__note">
                  {t(key)}
                </p>
              ))}
            </div>
          </section>
          <section class="pptx-panel__section">
            <h3 class="pptx-panel__title">{t('chartFormat.elements')}</h3>
            <ChartCheck
              label={t('chartFormat.showTitle')}
              checked={format.title !== null}
              disabled={disabled}
              onChange={(on: boolean) => set({ title: on ? t('chartFormat.defaultTitle') : null })}
            />
            {format.title !== null ? (
              <div class="pptx-chart-pane__sub">
                <ChartField
                  draftKey={draftKey}
                  label={t('chartFormat.titleText')}
                  visibleLabel={t('chartFormat.titleText')}
                  value={format.title}
                  disabled={disabled}
                  onCommit={(text: string) => run(chartTitleCommand(format, 'title', text))}
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
              onChange={(legend: ChartLegendPosition) => set({ legend })}
            />
            <ChartCheck
              label={t('chartFormat.dataLabels')}
              checked={format.dataLabels}
              disabled={disabled}
              onChange={(dataLabels: boolean) => set({ dataLabels })}
            />
            {format.axes !== null ? (
              <>
                <ChartCheck
                  label={t('chartFormat.axes')}
                  checked={format.axes}
                  disabled={disabled}
                  onChange={(axes: boolean) => set({ axes })}
                />
                <ChartCheck
                  label={t('chartFormat.gridlines')}
                  checked={format.gridlines === true}
                  disabled={disabled || !format.axes}
                  onChange={(gridlines: boolean) => set({ gridlines })}
                />
                <div class="pptx-panel__grid">
                  <ChartField
                    draftKey={draftKey}
                    label={t(axisLabels.category)}
                    visibleLabel={t(axisLabels.category)}
                    value={format.axisTitles?.category ?? ''}
                    disabled={disabled || !format.axes}
                    onCommit={(text: string) =>
                      run(chartTitleCommand(format, 'categoryAxisTitle', text))
                    }
                  />
                  <ChartField
                    draftKey={draftKey}
                    label={t(axisLabels.value)}
                    visibleLabel={t(axisLabels.value)}
                    value={format.axisTitles?.value ?? ''}
                    disabled={disabled || !format.axes}
                    onCommit={(text: string) =>
                      run(chartTitleCommand(format, 'valueAxisTitle', text))
                    }
                  />
                </div>
                <p class="pptx-chart-pane__note">{t('chartFormat.axisTitleHelp')}</p>
              </>
            ) : null}
          </section>
          <section class="pptx-panel__section">
            <h3 class="pptx-panel__title">{t('chartFormat.colors')}</h3>
            {format.colors.slice(0, MAX_COLOR_FIELDS).map((entry, i) => (
              <ColorField
                key={i}
                name={`chartColor${i}`}
                label={t('chartFormat.seriesColor', { name: entry.name || t('chartData.unnamed') })}
                value={chartColorPick(entry)}
                automatic={renderedColorOf(colorPickOf(entry.automatic), rows.value)}
                disabled={disabled}
                onChange={(pick: ColorPick) => run(chartColorCommand(i, pick))}
              />
            ))}
          </section>
        </>
      );
    };
  },
});

/** The Chart tab: the type, elements, colors, and data of the selected chart, or why it shows none. */
export const ChartDataPane = defineComponent({
  name: 'PptxChartDataPane',
  setup() {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const chart = useEditorState((s) => s.chart ?? null);
    // Every change to the document renders the tab again, which reads the chart's data again.
    const revision = useEditorState((s) => s.revision);
    const readOnly = useEditorState((s) => s.mode === 'view');
    const refusal = ref<string | null>(null);
    const open = ref<ReadonlySet<number>>(new Set([0]));
    let draftScope = { key: '', reset: 0 };
    const toggle = (index: number, isOpen: boolean): void => {
      if (open.value.has(index) === isOpen) return;
      const next = new Set(open.value);
      if (isOpen) next.add(index);
      else next.delete(index);
      open.value = next;
    };
    return () => {
      void revision.value;
      const selected = chart.value;
      const e = editor.value;
      const read = e && selected ? e.chartData(selected.nodeId) : null;
      if (!selected || !read)
        return (
          <div class="pptx-empty">
            <p class="pptx-empty__title">{t('chartData.emptyTitle')}</p>
            <p class="pptx-empty__body">{t('chartData.emptyBody')}</p>
          </div>
        );
      if (!read.ok)
        return (
          <div class="pptx-empty">
            <p class="pptx-empty__title">{t('chartData.notEditableTitle')}</p>
            <p class="pptx-empty__body">{t(read.reason as TranslationKey)}</p>
          </div>
        );
      const data = read.data;
      const scopeKey = chartPaneDraftKey(e!.snapshot(), data);
      if (scopeKey !== draftScope.key) {
        draftScope = {
          key: scopeKey,
          reset:
            draftScope.reset +
            (draftScope.key && chartPaneDraftContinues(draftScope.key, scopeKey) ? 0 : 1),
        };
      }
      const draftKey = String(draftScope.reset);
      const disabled = readOnly.value || !selected.editable;
      /** Applies one change: null when it applied, else the reason. */
      const apply = (change: ChartPaneChange): string | null => {
        const command = chartPaneCommand(data, change);
        if (!command.ok) return t(command.error);
        const result = editor.value?.exec(command.command);
        refusal.value = null;
        return !result || result.ok ? null : t(result.reason as TranslationKey);
      };
      /** Runs a command of the type, elements, or colors; the tab announces a refusal. */
      const runFormat = (command: EditorCommand | null): string | null => {
        if (!command) return null;
        const result = editor.value?.exec(command);
        const reason = !result || result.ok ? null : t(result.reason as TranslationKey);
        refusal.value = reason;
        return reason;
      };
      /** Applies an add or a remove; the tab announces a refusal, as no field shows it. */
      const run = (change: ChartPaneChange): boolean => {
        const reason = apply(change);
        refusal.value = reason;
        return reason === null;
      };
      const xy = data.kind === 'xy';
      const note =
        data.workbook === 'kept'
          ? t('chartData.workbookKept')
          : data.workbook === 'external'
            ? t('chartData.workbookExternal')
            : null;
      return (
        <div class="pptx-chart-pane">
          <p class="pptx-chart-pane__error" role="alert">
            {refusal.value ?? ''}
          </p>
          <ChartFormatSection
            nodeId={selected.nodeId}
            draftKey={draftKey}
            disabled={readOnly.value}
            seriesCount={data.series.length}
            run={runFormat}
          />
          <section class="pptx-panel__section">
            <h3 class="pptx-panel__title">
              {t(xy ? 'chartData.xValues' : 'chartData.categories')}
            </h3>
            <p class="pptx-chart-pane__note">{t('chartData.help')}</p>
            {note ? <p class="pptx-chart-pane__note">{note}</p> : null}
            <ul class="pptx-chart-pane__list">
              {data.categories.map((c, i) => (
                <li key={i} class="pptx-chart-pane__row">
                  <ChartField
                    draftKey={draftKey}
                    label={t(xy ? 'chartData.xValueLabel' : 'chartData.categoryLabel', {
                      number: i + 1,
                    })}
                    value={c.label}
                    numeric={xy}
                    disabled={disabled}
                    onCommit={(text: string) => apply({ type: 'label', category: i, text })}
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
              class="pptx-button"
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
          <section class="pptx-panel__section">
            <h3 class="pptx-panel__title">{t('chartData.series')}</h3>
            {data.series.map((_, j) => (
              <SeriesCard
                key={j}
                data={data}
                draftKey={draftKey}
                index={j}
                open={open.value.has(j)}
                disabled={disabled}
                onToggle={(isOpen: boolean) => toggle(j, isOpen)}
                apply={apply}
                run={run}
              />
            ))}
            <button
              type="button"
              class="pptx-button"
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
    };
  },
});
