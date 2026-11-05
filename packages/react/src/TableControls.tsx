/** @jsxImportSource react */
// Table chrome: the Insert > Table grid, where a hover picks the size, and
// the Table Design style list, which lists the deck's styles (`tableStyleChoices`). A menu
// places the grid with `MenuTableGrid` (the demo's Insert > Table), a toolbar with the
// `table.insert` slot.

import { useContext, useEffect, useRef, useState } from 'react';
import {
  TABLE_GRID_COLUMNS as GRID_COLUMNS,
  TABLE_GRID_ROWS as GRID_ROWS,
  tableGridKeyStep,
  tableGridMove,
  tableStyleChoices,
  type ChromeSlotId,
} from '@pptx-editor.dev/core/editor';
import { Icon, MenuItem, Popover, ToolButton, keepFocus, useAriaDisabled } from './chrome.tsx';
import { useEditorCommand, useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { MenuCloseContext } from './ContextMenu.tsx';
import { useSlot, type SlotPartProps } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';

/**
 * Props of `MenuTableGrid` (`PptxEditor.Menu.TableGrid`), DOCX Editor's `Menu.TableGrid` part
 * (`slot`, `className`), with the props a host menu or toolbar popup needs. @public
 */
export interface MenuTableGridProps {
  /** The slot whose enabled state gates the insert. Default `table.insert`. */
  readonly slot?: ChromeSlotId;
  /** Accessible name of the grid. Default: `table.insert` ("Table"). */
  readonly label?: string;
  /** Runs instead of the insert, with the picked size. */
  readonly onPick?: (rows: number, cols: number) => void;
  /** Runs after a table is inserted. Default: close the surrounding `PptxEditor.Menu`. */
  readonly onInserted?: () => void;
  /** Move the focus into the grid when it mounts, for a popup opened from the keyboard. */
  readonly autoFocus?: boolean;
  /** A class name added to the grid. */
  readonly className?: string;
}

/**
 * The Insert > Table grid, 10 columns by 8 rows, as DOCX Editor's `TableSizeGrid`: the cell
 * under the pointer (or focus) picks the size and each cell names it. A `grid` with one tab
 * stop; the arrows, Home, and End move inside it. The insert puts the caret in the new
 * table's first cell, and the editor takes the focus. @public
 */
export function MenuTableGrid({
  slot = 'table.insert',
  label,
  onPick,
  onInserted,
  autoFocus,
  className,
}: MenuTableGridProps) {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const cmd = useEditorCommand(slot);
  const [size, setSize] = useState({ rows: 0, cols: 0 });
  // The cell that holds the grid's single tab stop.
  const [cursor, setCursor] = useState({ rows: 1, cols: 1 });
  const grid = useRef<HTMLDivElement>(null);
  const closeMenu = useContext(MenuCloseContext);
  // As DOCX Editor's `TableInsert`: an opened grid takes the focus on its tab-stop cell. The
  // popup places itself (and becomes visible) in its own effect after this one, so the focus
  // waits for the next frame: a hidden element takes no focus.
  useEffect(() => {
    if (!autoFocus) return;
    const frame = requestAnimationFrame(() =>
      grid.current?.querySelector<HTMLElement>('[tabindex="0"]')?.focus()
    );
    return () => cancelAnimationFrame(frame);
  }, [autoFocus]);
  const pick = (rows: number, cols: number): void => {
    if (onPick) return onPick(rows, cols);
    // The slot gates the grid; the grid always inserts a table, as DOCX Editor's.
    const command = { type: 'insertTable' as const, rows, cols };
    if (!editor || !cmd.isEnabled || !editor.can(command).ok) return;
    // In a menu, the insert closes it, as DOCX Editor's grid does through the menu context.
    if (editor.exec(command).ok) (onInserted ?? closeMenu)?.();
  };
  const move = (step: { rows?: number; cols?: number; toCol?: number }): void =>
    setCursor((current) => {
      const next = tableGridMove(current, step);
      setSize(next);
      queueMicrotask(() =>
        grid.current?.querySelector<HTMLElement>(`[data-cell="${next.rows}x${next.cols}"]`)?.focus()
      );
      return next;
    });
  return (
    <div
      ref={grid}
      role="grid"
      aria-label={label ?? t('table.insert')}
      className={`pptx-table-grid${className ? ` ${className}` : ''}`}
      data-slot={slot}
      data-disabled={!onPick && !cmd.isEnabled ? '' : undefined}
      onMouseDown={keepFocus}
      onMouseLeave={() => setSize({ rows: 0, cols: 0 })}
      onKeyDown={(e) => {
        const step = tableGridKeyStep(e.key);
        if (!step) return;
        move(step);
        // The grid's arrows do not also walk the menu rows behind it.
        e.preventDefault();
        e.stopPropagation();
      }}
    >
      <div className="pptx-table-grid__cells" data-columns={GRID_COLUMNS}>
        {Array.from({ length: GRID_ROWS }, (_, r) => (
          <div key={r} role="row" className="pptx-table-grid__row">
            {Array.from({ length: GRID_COLUMNS }, (_, c) => (
              <button
                key={c}
                type="button"
                role="gridcell"
                data-cell={`${r + 1}x${c + 1}`}
                tabIndex={cursor.rows === r + 1 && cursor.cols === c + 1 ? 0 : -1}
                className="pptx-table-grid__cell"
                aria-label={t('table.gridSize', { cols: c + 1, rows: r + 1 })}
                data-active={r < size.rows && c < size.cols ? '' : undefined}
                onMouseEnter={() => setSize({ rows: r + 1, cols: c + 1 })}
                onFocus={() => setSize({ rows: r + 1, cols: c + 1 })}
                onClick={() => pick(r + 1, c + 1)}
              />
            ))}
          </div>
        ))}
      </div>
      {/* Not a live region: it would speak on every cell a pointer sweep crosses, and each
          cell already names its size. */}
      <p className="pptx-table-grid__label" aria-hidden="true">
        {size.rows ? t('table.gridSize', { cols: size.cols, rows: size.rows }) : t('table.insert')}
      </p>
    </div>
  );
}

/** Insert > Table as a control: a button that opens the grid. */
export function TableGridControl({ slot, control }: SlotPartProps) {
  const { cmd, label, reason } = useSlot(slot, control);
  const anchor = useRef<HTMLSpanElement>(null);
  const [open, setOpen] = useState(false);
  return (
    <>
      <span ref={anchor} className="pptx-toolbar__group">
        <ToolButton
          slot={slot}
          label={label}
          paths={control.paths}
          disabled={!cmd.isEnabled}
          disabledReason={reason}
          hasPopup="dialog"
          expanded={open}
          onClick={() => setOpen((o) => !o)}
        />
      </span>
      {open ? (
        // The popup anchors on the button, so Escape gives the focus back to it, as DOCX
        // Editor's `TableInsert`; it closes when the focus leaves the grid (Tab).
        <Popover
          anchor={anchor.current?.querySelector('button') ?? anchor.current}
          label={label}
          role="dialog"
          onClose={() => setOpen(false)}
        >
          <div
            onBlur={(e) => {
              if (!e.currentTarget.contains(e.relatedTarget as Node | null)) setOpen(false);
            }}
          >
            <MenuTableGrid slot={slot} label={label} autoFocus onInserted={() => setOpen(false)} />
          </div>
        </Popover>
      ) : null}
    </>
  );
}

/** Table Design > Table styles: the deck's styles by name; the table's style is checked. */
export function TableStyleControl({ slot, control }: SlotPartProps) {
  const { cmd, label, reason, tip } = useSlot(slot, control);
  // A disabled trigger stays focusable and says why, as every other picker does.
  const off = useAriaDisabled(!cmd.isEnabled, reason);
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const anchor = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  // The list comes from the open deck: read it again when the document changes.
  useEditorState((s) => s.revision);
  const styles = tableStyleChoices(editor);
  // Styles without a name are numbered among themselves, so the user can tell them apart.
  const unnamed = styles.filter((s) => !s.builtIn && !s.name);
  const trigger = useTriggerAttributes(slot, 'pptx-toolbar__button pptx-toolbar__button--wide', {
    disabled: !cmd.isEnabled,
  });
  return (
    <>
      <button
        ref={anchor}
        type="button"
        {...trigger}
        data-slot={slot}
        aria-label={label}
        aria-haspopup="menu"
        aria-expanded={open}
        {...tip}
        {...off.attrs}
        onMouseDown={keepFocus}
        onClick={() => setOpen((o) => !o)}
      >
        <Icon paths={control.paths} />
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>
      {off.note}
      {open ? (
        <Popover anchor={anchor.current} onClose={() => setOpen(false)}>
          {styles.map((s) => (
            <MenuItem
              key={s.id}
              checked={s.id === cmd.value}
              onSelect={() => (cmd.execute(s.id), setOpen(false))}
            >
              {s.builtIn
                ? t('table.defaultStyle')
                : s.name || t('table.unnamedStyle', { n: unnamed.indexOf(s) + 1 })}
            </MenuItem>
          ))}
        </Popover>
      ) : null}
    </>
  );
}
