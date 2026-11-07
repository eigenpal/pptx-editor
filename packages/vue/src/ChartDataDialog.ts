// The Edit Data dialog: the Vue twin of the React adapter's `ChartDataDialog`, on the same
// dialog parts and the same core rules (`chart-data-fields.ts`), so both show the same grid of
// categories (rows) by series (columns) and apply the same `setChartData` command. The host
// owns visibility (`open`, `onClose`).

import {
  Fragment,
  computed,
  defineComponent,
  h,
  nextTick,
  shallowRef,
  useId,
  watch,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
  type VNode,
} from 'vue';
import {
  chartDraftCell,
  chartDraftCellErrors,
  chartDraftCommand,
  chartDraftFrom,
  chartGridMove,
  commitDialogCommand,
  insertChartDraftColumn,
  insertChartDraftRow,
  removeChartDraftColumn,
  removeChartDraftRow,
  setChartDraftCell,
  translateDialogErrors,
  translateKeyOrText,
  validateChartDraft,
  type ChartDataDraft,
  type ChartDataDraftFields,
  type ChartGridPosition,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { ToolButton } from './chrome.tsx';
import {
  DialogFrame,
  applyOnEnter,
  createDialogParts,
  useCloseOnNewDocument,
  useDialogDraft,
  type UseDialogReturn,
} from './dialog-parts.ts';

/** The draft of the Edit Data dialog: the grid of the chart's data, as text. @public */
export type ChartDataDialogFields = ChartDataDraftFields;

/** The Edit Data draft and its actions. @public */
export type UseChartDataDialogReturn = UseDialogReturn<ChartDataDialogFields>;

/** The field names of the Edit Data dialog, for `ChartDataDialog.Field`. @public */
export type ChartDataDialogFieldName = 'grid';

const parts = createDialogParts<UseChartDataDialogReturn, ChartDataDialogFieldName>(
  'ChartDataDialog'
);

const chartDataDialogProps = {
  /** Whether the dialog is shown. The host owns this state. */
  open: { type: Boolean, required: true as const },
  /** Called on Cancel, Escape, a click on the backdrop, and after OK. */
  onClose: { type: Function as PropType<() => void>, required: true as const },
  /** `false`: only the default slot, with no default arrangement. */
  preset: { type: Boolean, default: true },
  /** A class name added to the dialog surface, as React's `className` (`class` works too). */
  className: { type: String, default: undefined },
  /** Styles added to the dialog surface. */
  style: { type: Object as PropType<CSSProperties>, default: undefined },
} as const;

/** Props of `ChartDataDialog`, the twin of React's `ChartDataDialogProps`. @public */
export type ChartDataDialogProps = ExtractPublicPropTypes<typeof chartDataDialogProps>;

/** The draft of a chart that cannot be edited: the dialog shows why, with no grid. */
const EMPTY: ChartDataDialogFields = {
  grid: { kind: 'category', rows: [], columns: [], workbook: 'none', numericX: false },
};

const toPosition = (key: string): ChartGridPosition => {
  const [row, column] = key.split(',').map(Number) as [number, number];
  return { row, column };
};

const ChartDataDialogOpen = defineComponent({
  name: 'ChartDataDialogOpen',
  inheritAttrs: false,
  props: chartDataDialogProps,
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    const editorRef = usePptxEditor();
    const id = useId();
    const read = editorRef.value?.chartData() ?? {
      ok: false as const,
      reason: 'errors.chartData.noChart',
    };
    const enabled = useEditorState(
      (s) => s.status === 'ready' && s.mode === 'edit' && !!s.chart?.editable
    );
    // As the React twin: the draft belongs to the chart the dialog opened on.
    const nodeId = useEditorState((s) => s.chart?.nodeId ?? null);
    const openedOn = nodeId.value;
    watch(nodeId, (now) => {
      if (now !== openedOn) props.onClose();
    });
    useCloseOnNewDocument(() => props.onClose());
    const draft = useDialogDraft<ChartDataDialogFields>({
      initial: read.ok ? chartDraftFrom(read.data) : EMPTY,
      validate: validateChartDraft,
      commit: (values) => commitDialogCommand(editorRef.value, chartDraftCommand(values)),
    });
    const focus = shallowRef<ChartGridPosition>({ row: 1, column: 1 });
    const cellId = (at: ChartGridPosition): string => `${id}-r${at.row}-c${at.column}`;
    const focusCell = (at: ChartGridPosition): void => {
      focus.value = at;
      document.getElementById(cellId(at))?.focus();
    };
    const edit = (next: ChartDataDraft, at?: ChartGridPosition): void => {
      draft.setValue('grid', next);
      if (at) void nextTick(() => focusCell(at));
    };
    const cellErrors = computed(() => chartDraftCellErrors(draft.values.value.grid));
    const state: UseChartDataDialogReturn = {
      values: draft.values,
      errors: computed(() => translateDialogErrors(t, draft.errors.value)),
      isEnabled: computed(() => enabled.value && read.ok),
      setValue: draft.setValue,
      apply: () => {
        if (draft.apply()) return props.onClose();
        const bad = cellErrors.value.keys().next();
        if (!bad.done) focusCell(toPosition(bad.value));
      },
      cancel: () => props.onClose(),
    };
    const render = parts.provideParts(state);
    return () => {
      const grid = draft.values.value.grid;
      const errors = cellErrors.value;
      const size = { rows: grid.rows.length + 1, columns: grid.columns.length + 1 };
      const cellName = (at: ChartGridPosition): string => {
        const label = grid.kind === 'xy' ? 'chartData.xValueLabel' : 'chartData.categoryLabel';
        if (at.column === 0) return t(label, { number: at.row });
        const category = grid.rows[at.row - 1]?.label || t(label, { number: at.row });
        const series =
          grid.columns[at.column - 1]?.name || t('chartData.newSeries', { number: at.column });
        return at.row === 0
          ? t('chartData.seriesName', { number: at.column })
          : t('chartData.cellLabel', { series, category });
      };
      const firstBad = errors.entries().next();
      const formError = draft.errors.value.form;
      const error = !firstBad.done
        ? t('chartData.errors.cell', {
            cell: cellName(toPosition(firstBad.value[0])),
            message: translateKeyOrText(t, firstBad.value[1]),
          })
        : formError
          ? translateKeyOrText(t, formError)
          : '';
      const field = (at: ChartGridPosition): VNode => {
        const bad = errors.has(`${at.row},${at.column}`);
        const active = focus.value.row === at.row && focus.value.column === at.column;
        return h('input', {
          id: cellId(at),
          type: 'text',
          class: 'pptx-chart-data__cell',
          inputmode: at.row > 0 && (at.column > 0 || grid.kind === 'xy') ? 'decimal' : undefined,
          value: chartDraftCell(grid, at.row, at.column),
          'aria-label': cellName(at),
          'aria-invalid': bad ? 'true' : 'false',
          'aria-describedby': bad && error ? `${id}-error` : undefined,
          tabindex: active ? 0 : -1,
          disabled: !enabled.value,
          onFocus: () => (focus.value = at),
          onInput: (e: Event) =>
            edit(
              setChartDraftCell(
                draft.values.value.grid,
                at.row,
                at.column,
                (e.target as HTMLInputElement).value
              )
            ),
          onKeydown: (e: KeyboardEvent) => {
            if (e.isComposing) return;
            const input = e.currentTarget as HTMLInputElement;
            const next = chartGridMove(
              { key: e.key, shiftKey: e.shiftKey, ctrlKey: e.ctrlKey || e.metaKey },
              at,
              size,
              {
                atStart: input.selectionStart === 0 && input.selectionEnd === 0,
                atEnd: input.selectionStart === input.value.length,
              }
            );
            // Return in the grid moves down, as in a sheet; it never applies the dialog.
            if (next || e.key === 'Enter') e.preventDefault();
            if (next) focusCell(next);
          },
        });
      };
      const f = focus.value;
      const rowsAt = f.row > 0 ? f.row : grid.rows.length;
      const columnAt = f.column > 0 ? f.column : grid.columns.length;
      const note =
        read.ok && read.data.workbook === 'kept'
          ? t('chartData.workbookKept')
          : read.ok && read.data.workbook === 'external'
            ? t('chartData.workbookExternal')
            : null;
      const button = (label: string, disabled: boolean, onClick: () => void) =>
        h('button', { type: 'button', class: 'pptx-button', disabled, onClick }, label);
      const body = !read.ok
        ? h(
            'p',
            { class: 'pptx-dialog__muted', role: 'status' },
            t('chartData.notEditable', { reason: translateKeyOrText(t, read.reason) })
          )
        : h(
            'section',
            {
              'data-pptx-part': 'field',
              'data-pptx-field': 'grid',
              class: 'pptx-dialog__section pptx-chart-data',
            },
            [
              h(
                'div',
                {
                  class: 'pptx-chart-data__actions',
                  role: 'toolbar',
                  'aria-label': t('chartData.grid'),
                },
                [
                  button(t('chartData.addCategory'), !enabled.value, () =>
                    edit(
                      insertChartDraftRow(
                        grid,
                        rowsAt + 1,
                        t('chartData.newCategory', { number: grid.rows.length + 1 })
                      ),
                      { row: rowsAt + 1, column: 0 }
                    )
                  ),
                  button(
                    t('chartData.removeCategory'),
                    !enabled.value || grid.rows.length <= 1 || f.row === 0,
                    () =>
                      edit(removeChartDraftRow(grid, f.row), {
                        row: Math.min(f.row, grid.rows.length - 1),
                        column: f.column,
                      })
                  ),
                  button(t('chartData.addSeries'), !enabled.value, () =>
                    edit(
                      insertChartDraftColumn(
                        grid,
                        columnAt + 1,
                        t('chartData.newSeries', { number: grid.columns.length + 1 })
                      ),
                      { row: 0, column: columnAt + 1 }
                    )
                  ),
                  button(
                    t('chartData.removeSeries'),
                    !enabled.value || grid.columns.length <= 1 || f.column === 0,
                    () =>
                      edit(removeChartDraftColumn(grid, f.column), {
                        row: f.row,
                        column: Math.min(f.column, grid.columns.length - 1),
                      })
                  ),
                ]
              ),
              h('p', { class: 'pptx-dialog__muted', id: `${id}-help` }, t('chartData.help')),
              h('div', { class: 'pptx-chart-data__scroll' }, [
                h(
                  'table',
                  {
                    role: 'grid',
                    class: 'pptx-chart-data__grid',
                    'aria-label': t('chartData.grid'),
                    'aria-describedby': `${id}-help`,
                    'aria-rowcount': size.rows,
                    'aria-colcount': size.columns,
                  },
                  [
                    h('thead', [
                      h('tr', { role: 'row', 'aria-rowindex': 1 }, [
                        h(
                          'th',
                          { role: 'columnheader', 'aria-colindex': 1, scope: 'col' },
                          t(grid.kind === 'xy' ? 'chartData.xValues' : 'chartData.categories')
                        ),
                        ...grid.columns.map((_, j) =>
                          h(
                            'th',
                            { key: j, role: 'columnheader', 'aria-colindex': j + 2, scope: 'col' },
                            [field({ row: 0, column: j + 1 })]
                          )
                        ),
                      ]),
                    ]),
                    h(
                      'tbody',
                      grid.rows.map((_, i) =>
                        h('tr', { key: i, role: 'row', 'aria-rowindex': i + 2 }, [
                          h('th', { role: 'rowheader', 'aria-colindex': 1, scope: 'row' }, [
                            field({ row: i + 1, column: 0 }),
                          ]),
                          ...grid.columns.map((__, j) =>
                            h('td', { key: j, role: 'gridcell', 'aria-colindex': j + 2 }, [
                              field({ row: i + 1, column: j + 1 }),
                            ])
                          ),
                        ])
                      )
                    ),
                  ]
                ),
              ]),
              note ? h('p', { class: 'pptx-dialog__muted', 'data-pptx-part': 'note' }, note) : null,
            ]
          );
      const defaults = h(Fragment, [
        h('div', { 'data-pptx-part': 'header', class: 'pptx-dialog__header' }, [
          h('h2', { 'data-pptx-part': 'title', class: 'pptx-dialog__title' }, t('chartData.title')),
          h(ToolButton, { label: t('chartData.close'), icon: 'close', onClick: props.onClose }),
        ]),
        h('div', { 'data-pptx-part': 'body', class: 'pptx-dialog__body' }, [body]),
        h('div', { 'data-pptx-part': 'footer', class: 'pptx-dialog__footer' }, [
          h(
            'span',
            {
              id: `${id}-error`,
              'data-pptx-part': 'error',
              role: 'alert',
              class: 'pptx-dialog__error',
            },
            error
          ),
          h(
            'button',
            {
              type: 'button',
              'data-pptx-part': 'cancel',
              class: 'pptx-button',
              onClick: state.cancel,
            },
            t('dialog.cancel')
          ),
          h(
            'button',
            {
              type: 'button',
              'data-pptx-part': 'apply',
              class: 'pptx-button pptx-button--primary',
              disabled: !state.isEnabled.value,
              onClick: state.apply,
            },
            t('dialog.ok')
          ),
        ]),
      ]);
      return h(
        DialogFrame,
        {
          ...attrs,
          class: [attrs.class, props.className],
          style: [attrs.style, props.style],
          kind: 'chartData',
          label: t('chartData.title'),
          onClose: props.onClose,
          onKeydown: (event: KeyboardEvent) =>
            applyOnEnter(event, state.isEnabled.value, state.apply),
        },
        () => render(defaults, slots.default?.(), props.preset)
      );
    };
  },
});

/**
 * The data of the selected chart as a grid of categories by series, with rows and columns to
 * add and remove, and replaceable parts (`Header`, `Title`, `Body`, `Footer`, `Field`,
 * `Error`, `Apply`, `Cancel`), as React's `ChartDataDialog`. Renders nothing while `open` is
 * false. @public
 */
export const ChartDataDialog = Object.assign(
  defineComponent({
    name: 'ChartDataDialog',
    inheritAttrs: false,
    props: chartDataDialogProps,
    setup(props, { attrs, slots }) {
      return () => (props.open ? h(ChartDataDialogOpen, { ...attrs, ...props }, slots) : null);
    },
  }),
  {
    Header: parts.Header,
    Title: parts.Title,
    Body: parts.Body,
    Footer: parts.Footer,
    Field: parts.Field,
    Error: parts.Error,
    Apply: parts.Apply,
    Cancel: parts.Cancel,
  }
);

/**
 * Read the draft of the enclosing Edit Data dialog, for a host's own part inside it. It does
 * not open the dialog: `useEditorDialog('chartData').open()` does. @public
 */
export function useChartDataDialog(): UseChartDataDialogReturn {
  return parts.useState();
}
