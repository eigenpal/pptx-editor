/** @jsxImportSource react */
// The Edit Data dialog (the chart toolbar group, and Edit Data in the chart context menu): a
// grid of the selected chart's categories (rows) by series (columns), with rows and columns to
// add and remove. Built on the dialog parts (`dialog-parts.tsx`): the cells write a draft, OK
// applies it as one undo step (`setChartData`), and Cancel discards it. The grid rules and the
// arrow key moves live in core (`chart-data-fields.ts`), so the Vue twin agrees.

import { useEffect, useId, useMemo, useRef, useState, type ReactNode } from 'react';
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
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { ToolButton } from './chrome.tsx';
import {
  DialogFrame,
  applyOnEnter,
  createDialogParts,
  useCloseOnNewDocument,
  useDialogDraft,
  type DialogCustomizationProps,
  type UseDialogReturn,
} from './dialog-parts.tsx';

/** The draft of the Edit Data dialog: the grid of the chart's data, as text. @public */
export type ChartDataDialogFields = ChartDataDraftFields;

/** The Edit Data draft and its actions. @public */
export type UseChartDataDialogReturn = UseDialogReturn<ChartDataDialogFields>;

/** The field names of the Edit Data dialog, for `ChartDataDialog.Field`. @public */
export type ChartDataDialogFieldName = 'grid';

/** Props for `PptxEditor.ChartDataDialog`. @public */
export interface ChartDataDialogProps extends DialogCustomizationProps {
  /** Whether the dialog is shown. The host owns this state. */
  readonly open: boolean;
  /** Called on Cancel, Escape, a click on the backdrop, and after OK. */
  readonly onClose: () => void;
}

const parts = createDialogParts<ChartDataDialogFieldName, UseChartDataDialogReturn>(
  'ChartDataDialog'
);

function ChartDataDialogRoot(props: ChartDataDialogProps): ReactNode {
  return props.open ? <ChartDataDialogOpen {...props} /> : null;
}

/** The id of the field of one grid cell. */
const cellId = (id: string, at: ChartGridPosition): string => `${id}-r${at.row}-c${at.column}`;

function ChartDataDialogOpen({
  onClose,
  className,
  style,
  preset = true,
  children,
}: ChartDataDialogProps): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const id = useId();
  const [read] = useState(
    () => editor?.chartData() ?? { ok: false as const, reason: 'errors.chartData.noChart' }
  );
  const enabled = useEditorState(
    (s) => s.status === 'ready' && s.mode === 'edit' && !!s.chart?.editable
  );
  // The draft belongs to the chart the dialog opened on: another selection closes the dialog.
  const nodeId = useEditorState((s) => s.chart?.nodeId ?? null);
  const [openedOn] = useState(nodeId);
  const closeRef = useRef(onClose);
  closeRef.current = onClose;
  useEffect(() => {
    if (nodeId !== openedOn) closeRef.current();
  }, [nodeId, openedOn]);
  useCloseOnNewDocument(onClose);
  const [initial] = useState(() => (read.ok ? chartDraftFrom(read.data) : null));
  const draft = useDialogDraft<ChartDataDialogFields>(() => ({
    initial: initial ?? EMPTY,
    validate: validateChartDraft,
    commit: (values) => commitDialogCommand(editor, chartDraftCommand(values)),
  }));
  const [focus, setFocus] = useState<ChartGridPosition>({ row: 1, column: 1 });
  const translatedErrors = useMemo(() => translateDialogErrors(t, draft.errors), [t, draft.errors]);
  const grid = draft.values.grid;
  const cellErrors = useMemo(() => chartDraftCellErrors(grid), [grid]);
  const size = { rows: grid.rows.length + 1, columns: grid.columns.length + 1 };
  const focusCell = (at: ChartGridPosition): void => {
    setFocus(at);
    document.getElementById(cellId(id, at))?.focus();
  };
  const edit = (next: ChartDataDraft, at?: ChartGridPosition): void => {
    draft.setValue('grid', next);
    if (at) requestAnimationFrame(() => focusCell(at));
  };
  const state: UseChartDataDialogReturn = {
    values: draft.values,
    setValue: draft.setValue,
    errors: translatedErrors,
    isEnabled: enabled && read.ok,
    apply: () => {
      if (draft.apply()) return onClose();
      const bad = cellErrors.keys().next();
      if (!bad.done) {
        const [row, column] = bad.value.split(',').map(Number) as [number, number];
        focusCell({ row, column });
      }
    },
    cancel: onClose,
  };
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
  const firstBad = cellErrors.entries().next();
  const error = !firstBad.done
    ? t('chartData.errors.cell', {
        cell: cellName(toPosition(firstBad.value[0])),
        message: translateKeyOrText(t, firstBad.value[1]),
      })
    : draft.errors.form
      ? translateKeyOrText(t, draft.errors.form)
      : null;
  const field = (at: ChartGridPosition): ReactNode => {
    const key = `${at.row},${at.column}`;
    const bad = cellErrors.has(key);
    const active = focus.row === at.row && focus.column === at.column;
    return (
      <input
        id={cellId(id, at)}
        type="text"
        className="pptx-chart-data__cell"
        inputMode={at.row > 0 && (at.column > 0 || grid.kind === 'xy') ? 'decimal' : undefined}
        value={chartDraftCell(grid, at.row, at.column)}
        aria-label={cellName(at)}
        aria-invalid={bad}
        aria-describedby={bad && error ? `${id}-error` : undefined}
        tabIndex={active ? 0 : -1}
        disabled={!enabled}
        onFocus={() => setFocus(at)}
        onChange={(e) => edit(setChartDraftCell(grid, at.row, at.column, e.target.value))}
        onKeyDown={(e) => {
          if (e.nativeEvent.isComposing) return;
          const input = e.currentTarget;
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
        }}
      />
    );
  };
  const rowsAt = focus.row > 0 ? focus.row : grid.rows.length;
  const columnAt = focus.column > 0 ? focus.column : grid.columns.length;
  const note =
    read.ok && read.data.workbook === 'kept'
      ? t('chartData.workbookKept')
      : read.ok && read.data.workbook === 'external'
        ? t('chartData.workbookExternal')
        : null;
  return (
    <DialogFrame
      kind="chartData"
      label={t('chartData.title')}
      className={className}
      style={style}
      onClose={onClose}
      onKeyDown={(event) => applyOnEnter(event, state.isEnabled, state.apply)}
    >
      <parts.Composition
        state={state}
        preset={preset}
        defaults={
          <>
            <div data-pptx-part="header" className="pptx-dialog__header">
              <h2 data-pptx-part="title" className="pptx-dialog__title">
                {t('chartData.title')}
              </h2>
              <ToolButton label={t('chartData.close')} icon="close" onClick={onClose} />
            </div>
            <div data-pptx-part="body" className="pptx-dialog__body">
              {!read.ok ? (
                <p className="pptx-dialog__muted" role="status">
                  {t('chartData.notEditable', { reason: translateKeyOrText(t, read.reason) })}
                </p>
              ) : (
                <section
                  data-pptx-part="field"
                  data-pptx-field="grid"
                  className="pptx-dialog__section pptx-chart-data"
                >
                  <div
                    className="pptx-chart-data__actions"
                    role="toolbar"
                    aria-label={t('chartData.grid')}
                  >
                    <button
                      type="button"
                      className="pptx-button"
                      disabled={!enabled}
                      onClick={() =>
                        edit(
                          insertChartDraftRow(
                            grid,
                            rowsAt + 1,
                            t('chartData.newCategory', { number: grid.rows.length + 1 })
                          ),
                          { row: rowsAt + 1, column: 0 }
                        )
                      }
                    >
                      {t('chartData.addCategory')}
                    </button>
                    <button
                      type="button"
                      className="pptx-button"
                      disabled={!enabled || grid.rows.length <= 1 || focus.row === 0}
                      onClick={() =>
                        edit(removeChartDraftRow(grid, focus.row), {
                          row: Math.min(focus.row, grid.rows.length - 1),
                          column: focus.column,
                        })
                      }
                    >
                      {t('chartData.removeCategory')}
                    </button>
                    <button
                      type="button"
                      className="pptx-button"
                      disabled={!enabled}
                      onClick={() =>
                        edit(
                          insertChartDraftColumn(
                            grid,
                            columnAt + 1,
                            t('chartData.newSeries', { number: grid.columns.length + 1 })
                          ),
                          { row: 0, column: columnAt + 1 }
                        )
                      }
                    >
                      {t('chartData.addSeries')}
                    </button>
                    <button
                      type="button"
                      className="pptx-button"
                      disabled={!enabled || grid.columns.length <= 1 || focus.column === 0}
                      onClick={() =>
                        edit(removeChartDraftColumn(grid, focus.column), {
                          row: focus.row,
                          column: Math.min(focus.column, grid.columns.length - 1),
                        })
                      }
                    >
                      {t('chartData.removeSeries')}
                    </button>
                  </div>
                  <p className="pptx-dialog__muted" id={`${id}-help`}>
                    {t('chartData.help')}
                  </p>
                  <div className="pptx-chart-data__scroll">
                    <table
                      role="grid"
                      className="pptx-chart-data__grid"
                      aria-label={t('chartData.grid')}
                      aria-describedby={`${id}-help`}
                      aria-rowcount={size.rows}
                      aria-colcount={size.columns}
                    >
                      <thead>
                        <tr role="row" aria-rowindex={1}>
                          <th role="columnheader" aria-colindex={1} scope="col">
                            {t(grid.kind === 'xy' ? 'chartData.xValues' : 'chartData.categories')}
                          </th>
                          {grid.columns.map((_, j) => (
                            <th key={j} role="columnheader" aria-colindex={j + 2} scope="col">
                              {field({ row: 0, column: j + 1 })}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {grid.rows.map((_, i) => (
                          <tr key={i} role="row" aria-rowindex={i + 2}>
                            <th role="rowheader" aria-colindex={1} scope="row">
                              {field({ row: i + 1, column: 0 })}
                            </th>
                            {grid.columns.map((__, j) => (
                              <td key={j} role="gridcell" aria-colindex={j + 2}>
                                {field({ row: i + 1, column: j + 1 })}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {note ? (
                    <p className="pptx-dialog__muted" data-pptx-part="note">
                      {note}
                    </p>
                  ) : null}
                </section>
              )}
            </div>
            <div data-pptx-part="footer" className="pptx-dialog__footer">
              <span
                id={`${id}-error`}
                data-pptx-part="error"
                role="alert"
                className="pptx-dialog__error"
              >
                {error}
              </span>
              <button
                type="button"
                data-pptx-part="cancel"
                className="pptx-button"
                onClick={state.cancel}
              >
                {t('dialog.cancel')}
              </button>
              <button
                type="button"
                data-pptx-part="apply"
                className="pptx-button pptx-button--primary"
                disabled={!state.isEnabled}
                onClick={state.apply}
              >
                {t('dialog.ok')}
              </button>
            </div>
          </>
        }
      >
        {children}
      </parts.Composition>
    </DialogFrame>
  );
}

/** The draft of a chart that cannot be edited: the dialog shows why, with no grid. */
const EMPTY: ChartDataDialogFields = {
  grid: { kind: 'category', rows: [], columns: [], workbook: 'none', numericX: false },
};

const toPosition = (key: string): ChartGridPosition => {
  const [row, column] = key.split(',').map(Number) as [number, number];
  return { row, column };
};

/**
 * Read the draft of the enclosing Edit Data dialog, for a host's own part inside it. It does
 * not open the dialog: `useEditorDialog('chartData').open()` does. @public
 */
export function useChartDataDialog(): UseChartDataDialogReturn {
  return parts.useState();
}

/**
 * The data of the selected chart as a grid of categories by series, with rows and columns to
 * add and remove, and replaceable parts (`Header`, `Title`, `Body`, `Footer`, `Field`,
 * `Error`, `Apply`, `Cancel`).
 *
 * @example
 * ```tsx
 * <PptxEditor.Root popups={{ chartData: (props) => <ChartDataDialog {...props} /> }}>
 *   …
 * </PptxEditor.Root>
 * ```
 *
 * @public
 */
export const ChartDataDialog = Object.assign(ChartDataDialogRoot, {
  Header: parts.Header,
  Title: parts.Title,
  Body: parts.Body,
  Footer: parts.Footer,
  Field: parts.Field,
  Error: parts.Error,
  Apply: parts.Apply,
  Cancel: parts.Cancel,
});
