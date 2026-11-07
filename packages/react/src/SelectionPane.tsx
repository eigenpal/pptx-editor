/** @jsxImportSource react */
// The selection pane: the current slide's elements, top first, as a tree. A click selects an
// element, even a hidden one; the eye shows or hides it; a double-click or F2 renames it. The
// rows, keys, and selection rules come from core (`selection-pane.ts`), so the Vue twin behaves
// the same.

import {
  useEffect,
  useId,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent,
  type MouseEvent,
} from 'react';
import {
  SELECTION_PANE_NAME_MAX,
  selectFromPane,
  selectionPaneAncestors,
  selectionPaneKey,
  selectionPaneRows,
  selectionPaneSelected,
  selectionPaneAllIds,
  type SelectionPaneRow,
} from '@pptx-editor.dev/core/editor';
import { Icon, ToolButton, keepFocus } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';

/** Props of `SelectionPane`. @public */
export interface SelectionPaneProps {
  /** A class name added to the pane. */
  readonly className?: string;
}

/**
 * The selection pane: select, show or hide, rename, and reorder the elements of the current
 * slide. Render it inside the editor root, or open the side panel's **Selection** tab. @public
 */
export function SelectionPane(props: SelectionPaneProps = {}) {
  // Node ids repeat across slides: a new slide starts a new pane, so no rename, focus, or open
  // group carries over to an element of another slide.
  const slideId = useEditorState((s) => s.slides[s.currentSlide]?.id ?? '');
  return <SlidePane key={slideId} {...props} />;
}

function SlidePane({ className }: SelectionPaneProps) {
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const shapes = useEditorState((s) => s.shapes);
  const selection = useEditorState((s) => s.selection);
  const editable = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
  const [expanded, setExpanded] = useState<ReadonlySet<string>>(() => new Set());
  const [focused, setFocused] = useState<string | null>(null);
  const [renaming, setRenaming] = useState<string | null>(null);
  const tree = useRef<HTMLDivElement>(null);
  const hintId = useId();
  const selected = useMemo(() => selectionPaneSelected(selection), [selection]);
  // A selected child of a group shows: its groups open when the selection reaches it.
  useEffect(() => {
    const open = selectionPaneAncestors(shapes, selected);
    if (open.some((id) => !expanded.has(id))) setExpanded(new Set([...expanded, ...open]));
    // `expanded` is left out on purpose: a group the user closes stays closed.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shapes, selected]);
  const rows = useMemo(() => selectionPaneRows(shapes, expanded), [shapes, expanded]);
  // The one Tab stop of the tree: the focused row, else the first selected row, else the top.
  const stop =
    rows.find((r) => r.nodeId === focused)?.nodeId ??
    rows.find((r) => selected.has(r.nodeId))?.nodeId ??
    rows[0]?.nodeId ??
    null;
  const anySelected = rows.some((r) => selected.has(r.nodeId));
  const order = (move: 'forward' | 'backward') => editor?.exec({ type: 'order', move });

  const focusRow = (id: string): void => {
    setFocused(id);
    tree.current
      ?.querySelector<HTMLElement>(`[data-node-id="${CSS.escape(id)}"]`)
      ?.focus({ preventScroll: false });
  };
  const toggle = (id: string, open: boolean): void => {
    setExpanded((prev) => {
      const next = new Set(prev);
      if (open) next.add(id);
      else next.delete(id);
      return next;
    });
  };
  const setVisible = (ids: readonly string[], visible: boolean): void => {
    if (ids.length) editor?.exec({ type: 'setShapesVisible', nodeIds: ids, visible });
  };
  const onKeyDown = (e: KeyboardEvent<HTMLDivElement>, row: SelectionPaneRow): void => {
    if (e.target !== e.currentTarget || e.altKey) return;
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      if (editor) selectFromPane(editor, row, e.shiftKey || e.ctrlKey || e.metaKey);
      return;
    }
    if (e.ctrlKey || e.metaKey) return;
    const step = selectionPaneKey(rows, row.nodeId, e.key);
    if (!step) return;
    e.preventDefault();
    if ('focus' in step) focusRow(step.focus);
    else if ('expand' in step) toggle(step.expand, step.open);
    else if (!editable) return;
    else if ('rename' in step) setRenaming(step.rename);
    else setVisible([row.nodeId], row.hidden);
  };

  if (rows.length === 0) {
    return (
      <div className={['pptx-empty', className].filter(Boolean).join(' ')}>
        <p className="pptx-empty__title">{t('selectionPane.emptyTitle')}</p>
        <p className="pptx-empty__body">{t('selectionPane.emptyBody')}</p>
      </div>
    );
  }
  return (
    <div className={['pptx-selection-pane', className].filter(Boolean).join(' ')}>
      <div className="pptx-selection-pane__bar" onMouseDown={keepFocus}>
        <button
          type="button"
          className="pptx-button"
          disabled={!editable}
          onClick={() => setVisible(selectionPaneAllIds(shapes), true)}
        >
          {t('selectionPane.showAll')}
        </button>
        <button
          type="button"
          className="pptx-button"
          disabled={!editable}
          onClick={() => setVisible(selectionPaneAllIds(shapes), false)}
        >
          {t('selectionPane.hideAll')}
        </button>
        <span className="pptx-selection-pane__order">
          <ToolButton
            label={t('toolbar.bringForward')}
            icon="keyboard_arrow_up"
            disabled={!editable || !anySelected}
            disabledReason={editable ? t('selectionPane.selectFirst') : undefined}
            onClick={() => order('forward')}
          />
          <ToolButton
            label={t('toolbar.sendBackward')}
            icon="keyboard_arrow_down"
            disabled={!editable || !anySelected}
            disabledReason={editable ? t('selectionPane.selectFirst') : undefined}
            onClick={() => order('backward')}
          />
        </span>
      </div>
      {editable ? (
        <p id={hintId} className="pptx-selection-pane__hint">
          {t('selectionPane.keyHint')}
        </p>
      ) : null}
      <div
        ref={tree}
        role="tree"
        aria-label={t('selectionPane.label')}
        aria-describedby={editable ? hintId : undefined}
        aria-multiselectable="true"
        className="pptx-selection-pane__tree"
      >
        {rows.map((row) => (
          <PaneRow
            key={row.nodeId}
            row={row}
            selected={selected.has(row.nodeId)}
            tabStop={row.nodeId === stop}
            editable={editable}
            renaming={renaming === row.nodeId}
            onFocus={() => setFocused(row.nodeId)}
            onKeyDown={(e) => onKeyDown(e, row)}
            onSelect={(e) =>
              editor && selectFromPane(editor, row, e.shiftKey || e.ctrlKey || e.metaKey)
            }
            onToggle={() => toggle(row.nodeId, !row.expanded)}
            onVisible={() => setVisible([row.nodeId], row.hidden)}
            onRename={() => editable && setRenaming(row.nodeId)}
            onRenamed={(name, refocus) => {
              setRenaming(null);
              const next = name?.trim();
              // The element must still be on the slide: a peer can delete it while you type.
              const here = selectionPaneAllIds(editor?.snapshot().shapes).includes(row.nodeId);
              if (here && next && next !== row.name)
                editor?.exec({ type: 'renameShape', nodeId: row.nodeId, name: next });
              // Enter and Escape give the row its focus back; a click elsewhere keeps its own.
              if (refocus && here) focusRow(row.nodeId);
            }}
          />
        ))}
      </div>
    </div>
  );
}

interface PaneRowProps {
  readonly row: SelectionPaneRow;
  readonly selected: boolean;
  readonly tabStop: boolean;
  readonly editable: boolean;
  readonly renaming: boolean;
  readonly onFocus: () => void;
  readonly onKeyDown: (e: KeyboardEvent<HTMLDivElement>) => void;
  readonly onSelect: (e: MouseEvent) => void;
  readonly onToggle: () => void;
  readonly onVisible: () => void;
  readonly onRename: () => void;
  /** The new name, or null to keep the old one; `refocus` after Enter or Escape. */
  readonly onRenamed: (name: string | null, refocus: boolean) => void;
}

function PaneRow(props: PaneRowProps) {
  const { row } = props;
  const { t } = useTranslation();
  const kind = t(`panel.format.elementKind.${row.kind}`);
  return (
    <div
      role="treeitem"
      className="pptx-selection-pane__row"
      data-node-id={row.nodeId}
      data-hidden={row.hidden ? '' : undefined}
      aria-level={row.level}
      aria-posinset={row.position}
      aria-setsize={row.siblings}
      aria-selected={props.selected}
      aria-expanded={row.branch ? row.expanded : undefined}
      aria-label={row.hidden ? t('selectionPane.hiddenName', { name: row.name }) : row.name}
      aria-description={kind}
      aria-keyshortcuts={props.editable ? 'H F2' : undefined}
      tabIndex={props.tabStop ? 0 : -1}
      style={{ paddingInlineStart: `${(row.level - 1) * 16 + 4}px` }}
      onFocus={props.onFocus}
      onKeyDown={props.onKeyDown}
      onMouseDown={(e) => {
        // A press on the row selects; the row keeps the focus, not the slide.
        if ((e.target as HTMLElement).closest('input, button')) return;
        e.preventDefault();
        e.currentTarget.focus();
        props.onSelect(e);
      }}
      onDoubleClick={(e) => {
        if (!(e.target as HTMLElement).closest('input, button')) props.onRename();
      }}
    >
      {row.branch ? (
        <span
          className="pptx-selection-pane__twisty"
          aria-hidden="true"
          onMouseDown={(e) => {
            e.preventDefault();
            e.stopPropagation();
            props.onToggle();
          }}
        >
          <Icon name={row.expanded ? 'expand_more' : 'chevron_right'} />
        </span>
      ) : (
        <span className="pptx-selection-pane__twisty" aria-hidden="true" />
      )}
      {props.renaming ? (
        <RenameField name={row.name} label={t('selectionPane.rename')} onDone={props.onRenamed} />
      ) : (
        <span className="pptx-selection-pane__name">{row.name}</span>
      )}
      <button
        type="button"
        className="pptx-selection-pane__eye"
        tabIndex={-1}
        aria-label={t(row.hidden ? 'selectionPane.show' : 'selectionPane.hide', {
          name: row.name,
        })}
        title={t(row.hidden ? 'selectionPane.show' : 'selectionPane.hide', { name: row.name })}
        disabled={!props.editable}
        onMouseDown={keepFocus}
        onClick={props.onVisible}
      >
        <Icon name={row.hidden ? 'visibility_off' : 'visibility'} />
      </button>
    </div>
  );
}

/** The inline name field: Enter or a click away applies, Escape keeps the old name. */
function RenameField(props: {
  name: string;
  label: string;
  onDone: (name: string | null, refocus: boolean) => void;
}) {
  const field = useRef<HTMLInputElement>(null);
  const done = useRef(false);
  useEffect(() => {
    field.current?.focus();
    field.current?.select();
  }, []);
  const finish = (name: string | null, refocus: boolean): void => {
    if (done.current) return;
    done.current = true;
    props.onDone(name, refocus);
  };
  return (
    <input
      ref={field}
      className="pptx-selection-pane__field"
      aria-label={props.label}
      defaultValue={props.name}
      maxLength={SELECTION_PANE_NAME_MAX}
      onKeyDown={(e) => {
        e.stopPropagation();
        if (e.key === 'Enter') {
          e.preventDefault();
          finish(e.currentTarget.value, true);
        } else if (e.key === 'Escape') {
          e.preventDefault();
          finish(null, true);
        }
      }}
      onBlur={(e) => finish(e.currentTarget.value, false)}
    />
  );
}
