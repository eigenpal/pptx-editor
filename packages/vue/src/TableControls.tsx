/** @jsxImportSource vue */
// Table chrome: the Insert > Table grid, where a hover picks the size, and the Table Design
// style list, which lists the deck's styles (`tableStyleChoices`). The Vue twin of the React
// adapter's `TableControls.tsx`, on the same core rules (`insert-controls.ts`). A menu places the
// grid with `MenuTableGrid`, a toolbar with the `table.insert` slot.

import {
  defineComponent,
  inject,
  nextTick,
  onBeforeUnmount,
  onMounted,
  ref,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
import {
  TABLE_GRID_COLUMNS,
  TABLE_GRID_ROWS,
  tableGridKeyStep,
  tableGridMove,
  tableStyleChoices,
  type ChromeControl,
  type ChromeSlotId,
  type TableGridSize,
} from '@pptx-editor.dev/core/editor';
import { Icon, MenuItem, Popover, ToolButton, keepFocus, useAriaDisabled } from './chrome.tsx';
import { useEditorCommand, useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { menuCloseKey } from './menu-row.tsx';
import { useSlot } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';

const menuTableGridProps = {
  /** The slot whose enabled state gates the insert. Default `table.insert`. */
  slot: { type: String as PropType<ChromeSlotId>, default: 'table.insert' },
  /** Accessible name of the grid. Default: `table.insert` ("Table"). */
  label: { type: String, default: undefined },
  /** Move the focus into the grid when it mounts, for a popup opened from the keyboard. */
  autoFocus: { type: Boolean, default: false },
  // The callbacks have no `default`: an absent function prop is `undefined` anyway, and a
  // default puts them in the declaration's list of defaulted props, whose order the compiler
  // takes from the order it first meets each name. That order differed between builds.
  /** Runs instead of the insert, with the rows and columns picked, as React's `onPick`. */
  onPick: { type: Function as PropType<(rows: number, cols: number) => void> },
  /**
   * Runs after the grid inserts a table, as React's `onInserted`. Without it, the insert closes
   * the surrounding menu.
   */
  onInserted: { type: Function as PropType<() => void> },
  /** A class name added to the grid, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `MenuTableGrid`, the twin of React's `MenuTableGridProps`. @public */
export type MenuTableGridProps = ExtractPublicPropTypes<typeof menuTableGridProps>;

/**
 * The Insert > Table grid, 10 columns by 8 rows, as DOCX Editor's `TableSizeGrid`: the cell
 * under the pointer (or focus) picks the size and each cell names it. A `grid` with one tab
 * stop; the arrows, Home, and End move inside it. `onPick` runs instead of the insert, with the
 * rows and columns; after an insert it calls `onInserted`, and without one it closes the
 * surrounding menu. `className` (or a host `class`) goes on the grid. @public
 */
export const MenuTableGrid = defineComponent({
  name: 'MenuTableGrid',
  props: menuTableGridProps,
  setup(props) {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const cmd = useEditorCommand(() => props.slot);
    const size = ref<TableGridSize>({ rows: 0, cols: 0 });
    // The cell that holds the grid's single tab stop.
    const cursor = ref<TableGridSize>({ rows: 1, cols: 1 });
    const grid = ref<HTMLDivElement | null>(null);
    const closeMenu = inject(menuCloseKey, null);
    // As DOCX Editor's `TableInsert`: an opened grid takes the focus on its tab-stop cell. The
    // popup places itself (and becomes visible) after this mounts, so the focus waits for the
    // next frame: a hidden element takes no focus.
    let frame = 0;
    onMounted(() => {
      if (!props.autoFocus) return;
      frame = requestAnimationFrame(() =>
        grid.value?.querySelector<HTMLElement>('[tabindex="0"]')?.focus()
      );
    });
    onBeforeUnmount(() => cancelAnimationFrame(frame));
    const pick = (rows: number, cols: number): void => {
      if (props.onPick) return void props.onPick(rows, cols);
      // The slot gates the grid; the grid always inserts a table, as DOCX Editor's.
      const command = { type: 'insertTable' as const, rows, cols };
      const ed = editor.value;
      if (!ed || !cmd.isEnabled.value || !ed.can(command).ok) return;
      if (!ed.exec(command).ok) return;
      // In a menu, the insert closes it, as DOCX Editor's grid does through the menu context.
      if (props.onInserted) props.onInserted();
      else closeMenu?.();
    };
    const move = (step: Parameters<typeof tableGridMove>[1]): void => {
      const next = tableGridMove(cursor.value, step);
      cursor.value = next;
      size.value = next;
      void nextTick(() =>
        grid.value?.querySelector<HTMLElement>(`[data-cell="${next.rows}x${next.cols}"]`)?.focus()
      );
    };
    return () => (
      <div
        ref={grid}
        role="grid"
        aria-label={props.label ?? t('table.insert')}
        class={['pptx-table-grid', props.className]}
        data-slot={props.slot}
        data-disabled={!props.onPick && !cmd.isEnabled.value ? '' : undefined}
        onMousedown={keepFocus}
        onMouseleave={() => (size.value = { rows: 0, cols: 0 })}
        onKeydown={(e: KeyboardEvent) => {
          const step = tableGridKeyStep(e.key);
          if (!step) return;
          move(step);
          // The grid's arrows do not also walk the menu rows behind it.
          e.preventDefault();
          e.stopPropagation();
        }}
      >
        <div class="pptx-table-grid__cells" data-columns={TABLE_GRID_COLUMNS}>
          {Array.from({ length: TABLE_GRID_ROWS }, (_, r) => (
            <div key={r} role="row" class="pptx-table-grid__row">
              {Array.from({ length: TABLE_GRID_COLUMNS }, (_, c) => (
                <button
                  key={c}
                  type="button"
                  role="gridcell"
                  data-cell={`${r + 1}x${c + 1}`}
                  tabindex={cursor.value.rows === r + 1 && cursor.value.cols === c + 1 ? 0 : -1}
                  class="pptx-table-grid__cell"
                  aria-label={t('table.gridSize', { cols: c + 1, rows: r + 1 })}
                  data-active={r < size.value.rows && c < size.value.cols ? '' : undefined}
                  onMouseenter={() => (size.value = { rows: r + 1, cols: c + 1 })}
                  onFocus={() => (size.value = { rows: r + 1, cols: c + 1 })}
                  onClick={() => pick(r + 1, c + 1)}
                />
              ))}
            </div>
          ))}
        </div>
        {/* Not a live region: it would speak on every cell a pointer sweep crosses, and each
            cell already names its size. */}
        <p class="pptx-table-grid__label" aria-hidden="true">
          {size.value.rows
            ? t('table.gridSize', { cols: size.value.cols, rows: size.value.rows })
            : t('table.insert')}
        </p>
      </div>
    );
  },
});

const slotPartProps = {
  slot: { type: String as PropType<ChromeSlotId>, required: true as const },
  control: { type: Object as PropType<ChromeControl>, required: true as const },
} as const;

/** Insert > Table as a control: a button that opens the grid. @internal */
export const TableGridControl = defineComponent({
  name: 'PptxTableGridControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason } = useSlot(
      () => props.slot,
      () => props.control
    );
    const anchor = ref<HTMLSpanElement | null>(null);
    const open = ref(false);
    const close = (): void => {
      open.value = false;
    };
    return () => [
      <span ref={anchor} class="pptx-toolbar__group">
        <ToolButton
          slot={props.slot}
          label={label.value}
          paths={props.control.paths}
          disabled={!cmd.isEnabled.value}
          disabledReason={reason.value}
          hasPopup="dialog"
          expanded={open.value}
          onClick={() => (open.value = !open.value)}
        />
      </span>,
      open.value ? (
        // The popup anchors on the button, so Escape gives the focus back to it, as DOCX
        // Editor's `TableInsert`; it closes when the focus leaves the grid (Tab).
        <Popover
          anchor={anchor.value?.querySelector('button') ?? anchor.value}
          label={label.value}
          role="dialog"
          onClose={close}
        >
          {{
            default: () => (
              <>
                <div
                  onFocusout={(e: FocusEvent) => {
                    const box = e.currentTarget as HTMLElement;
                    if (!box.contains(e.relatedTarget as Node | null)) close();
                  }}
                >
                  <MenuTableGrid
                    slot={props.slot}
                    label={label.value}
                    autoFocus
                    onInserted={close}
                  />
                </div>
              </>
            ),
          }}
        </Popover>
      ) : null,
    ];
  },
});

/** Table Design > Table styles: the deck's styles by name; the table's style is checked. @internal */
export const TableStyleControl = defineComponent({
  name: 'PptxTableStyleControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason, tip } = useSlot(
      () => props.slot,
      () => props.control
    );
    // A disabled trigger stays focusable and says why, as every other picker does.
    const off = useAriaDisabled(() => !cmd.isEnabled.value, reason);
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const anchor = ref<HTMLButtonElement | null>(null);
    const open = ref(false);
    // The list comes from the open deck: read it again when the document changes.
    const revision = useEditorState((s) => s.revision);
    const trigger = useTriggerAttributes(
      props.slot,
      'pptx-toolbar__button pptx-toolbar__button--wide',
      () => ({ disabled: !cmd.isEnabled.value })
    );
    return () => {
      void revision.value;
      const styles = tableStyleChoices(editor.value);
      // Styles without a name are numbered among themselves, so the user can tell them apart.
      const unnamed = styles.filter((s) => !s.builtIn && !s.name);
      return [
        <button
          ref={anchor}
          type="button"
          {...trigger()}
          data-slot={props.slot}
          aria-label={label.value}
          aria-haspopup="menu"
          aria-expanded={open.value}
          {...tip.value}
          {...off.attrs.value}
          onMousedown={keepFocus}
          onClick={() => (open.value = !open.value)}
        >
          <Icon paths={props.control.paths} />
          <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
        </button>,
        off.note(),
        open.value ? (
          <Popover anchor={anchor.value} onClose={() => (open.value = false)}>
            {{
              default: () => (
                <>
                  {styles.map((s) => (
                    <MenuItem
                      key={s.id}
                      checked={s.id === cmd.value.value}
                      onSelect={() => {
                        cmd.execute(s.id);
                        open.value = false;
                      }}
                    >
                      {{
                        default: () =>
                          s.builtIn
                            ? t('table.defaultStyle')
                            : s.name || t('table.unnamedStyle', { n: unnamed.indexOf(s) + 1 }),
                      }}
                    </MenuItem>
                  ))}
                </>
              ),
            }}
          </Popover>
        ) : null,
      ];
    };
  },
});
