/** @jsxImportSource vue */
// The selection pane: the Vue twin of the React adapter's `SelectionPane.tsx`, with the same
// markup, classes, and ARIA, on the same core rules (`selection-pane.ts`).

import {
  computed,
  defineComponent,
  nextTick,
  onMounted,
  ref,
  shallowRef,
  useId,
  watch,
  h,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
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
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';

const selectionPaneProps = {
  /** A class name added to the pane, as React's `className`. */
  className: { type: String, default: undefined },
} as const;

/** Props of `PptxEditorSelectionPane`, the twin of React's `SelectionPaneProps`. @public */
export type SelectionPaneProps = ExtractPublicPropTypes<typeof selectionPaneProps>;

/** The inline name field: Enter or a click away applies, Escape keeps the old name. */
const RenameField = defineComponent({
  name: 'PptxSelectionPaneRename',
  props: {
    name: { type: String, required: true },
    label: { type: String, required: true },
    onDone: {
      type: Function as PropType<(name: string | null, refocus: boolean) => void>,
      required: true,
    },
  },
  setup(props) {
    const field = ref<HTMLInputElement | null>(null);
    let done = false;
    onMounted(() => {
      if (!field.value) return;
      field.value.value = props.name;
      field.value.focus();
      field.value.select();
    });
    const finish = (name: string | null, refocus: boolean): void => {
      if (done) return;
      done = true;
      props.onDone(name, refocus);
    };
    return () => (
      <input
        ref={field}
        class="pptx-selection-pane__field"
        aria-label={props.label}
        maxlength={SELECTION_PANE_NAME_MAX}
        onKeydown={(e: KeyboardEvent) => {
          e.stopPropagation();
          const value = (e.currentTarget as HTMLInputElement).value;
          if (e.key === 'Enter') {
            e.preventDefault();
            finish(value, true);
          } else if (e.key === 'Escape') {
            e.preventDefault();
            finish(null, true);
          }
        }}
        onBlur={(e: FocusEvent) => finish((e.currentTarget as HTMLInputElement).value, false)}
      />
    );
  },
});

/** The selection pane of one slide. */
const SlidePane = defineComponent({
  name: 'PptxSelectionPaneSlide',
  props: selectionPaneProps,
  setup(props) {
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const shapes = useEditorState((s) => s.shapes);
    const selection = useEditorState((s) => s.selection);
    const editable = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
    const expanded = shallowRef<ReadonlySet<string>>(new Set());
    const focused = ref<string | null>(null);
    const renaming = ref<string | null>(null);
    const tree = ref<HTMLDivElement | null>(null);
    const hintId = useId();
    const selected = computed(() => selectionPaneSelected(selection.value));
    // A selected child of a group shows: its groups open when the selection reaches it. A group
    // the user closes stays closed until the selection moves again.
    watch(
      [shapes, selected],
      () => {
        const open = selectionPaneAncestors(shapes.value, selected.value);
        if (open.some((id) => !expanded.value.has(id)))
          expanded.value = new Set([...expanded.value, ...open]);
      },
      { immediate: true }
    );
    const rows = computed(() => selectionPaneRows(shapes.value, expanded.value));

    const focusRow = (id: string): void => {
      focused.value = id;
      void nextTick(() =>
        tree.value
          ?.querySelector<HTMLElement>(`[data-node-id="${CSS.escape(id)}"]`)
          ?.focus({ preventScroll: false })
      );
    };
    const toggle = (id: string, open: boolean): void => {
      const next = new Set(expanded.value);
      if (open) next.add(id);
      else next.delete(id);
      expanded.value = next;
    };
    const setVisible = (ids: readonly string[], visible: boolean): void => {
      if (ids.length) editor.value?.exec({ type: 'setShapesVisible', nodeIds: ids, visible });
    };
    const select = (row: SelectionPaneRow, e: MouseEvent | KeyboardEvent): void => {
      if (editor.value) selectFromPane(editor.value, row, e.shiftKey || e.ctrlKey || e.metaKey);
    };
    const onKeydown = (e: KeyboardEvent, row: SelectionPaneRow): void => {
      if (e.target !== e.currentTarget || e.altKey) return;
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        select(row, e);
        return;
      }
      if (e.ctrlKey || e.metaKey) return;
      const step = selectionPaneKey(rows.value, row.nodeId, e.key);
      if (!step) return;
      e.preventDefault();
      if ('focus' in step) focusRow(step.focus);
      else if ('expand' in step) toggle(step.expand, step.open);
      else if (!editable.value) return;
      else if ('rename' in step) renaming.value = step.rename;
      else setVisible([row.nodeId], row.hidden);
    };
    const renamed = (row: SelectionPaneRow, name: string | null, refocus: boolean): void => {
      renaming.value = null;
      const next = name?.trim();
      // The element must still be on the slide: a peer can delete it while you type.
      const here = selectionPaneAllIds(editor.value?.snapshot().shapes).includes(row.nodeId);
      if (here && next && next !== row.name)
        editor.value?.exec({ type: 'renameShape', nodeId: row.nodeId, name: next });
      // Enter and Escape give the row its focus back; a click elsewhere keeps its own.
      if (refocus && here) focusRow(row.nodeId);
    };
    const order = (move: 'forward' | 'backward') => editor.value?.exec({ type: 'order', move });

    return () => {
      const list = rows.value;
      if (list.length === 0) {
        return (
          <div class={['pptx-empty', props.className]}>
            <p class="pptx-empty__title">{t('selectionPane.emptyTitle')}</p>
            <p class="pptx-empty__body">{t('selectionPane.emptyBody')}</p>
          </div>
        );
      }
      const sel = selected.value;
      // The one Tab stop of the tree: the focused row, else the first selected row, else the top.
      const stop =
        list.find((r) => r.nodeId === focused.value)?.nodeId ??
        list.find((r) => sel.has(r.nodeId))?.nodeId ??
        list[0]?.nodeId ??
        null;
      const anySelected = list.some((r) => sel.has(r.nodeId));
      const canEdit = editable.value;
      return (
        <div class={['pptx-selection-pane', props.className]}>
          <div class="pptx-selection-pane__bar" onMousedown={keepFocus}>
            <button
              type="button"
              class="pptx-button"
              disabled={!canEdit}
              onClick={() => setVisible(selectionPaneAllIds(shapes.value), true)}
            >
              {t('selectionPane.showAll')}
            </button>
            <button
              type="button"
              class="pptx-button"
              disabled={!canEdit}
              onClick={() => setVisible(selectionPaneAllIds(shapes.value), false)}
            >
              {t('selectionPane.hideAll')}
            </button>
            <span class="pptx-selection-pane__order">
              <ToolButton
                label={t('toolbar.bringForward')}
                icon="keyboard_arrow_up"
                disabled={!canEdit || !anySelected}
                disabledReason={canEdit ? t('selectionPane.selectFirst') : undefined}
                onClick={() => order('forward')}
              />
              <ToolButton
                label={t('toolbar.sendBackward')}
                icon="keyboard_arrow_down"
                disabled={!canEdit || !anySelected}
                disabledReason={canEdit ? t('selectionPane.selectFirst') : undefined}
                onClick={() => order('backward')}
              />
            </span>
          </div>
          {canEdit ? (
            <p id={hintId} class="pptx-selection-pane__hint">
              {t('selectionPane.keyHint')}
            </p>
          ) : null}
          <div
            ref={tree}
            role="tree"
            aria-label={t('selectionPane.label')}
            aria-describedby={canEdit ? hintId : undefined}
            aria-multiselectable="true"
            class="pptx-selection-pane__tree"
          >
            {list.map((row) => {
              const eye = t(row.hidden ? 'selectionPane.show' : 'selectionPane.hide', {
                name: row.name,
              });
              return (
                <div
                  key={row.nodeId}
                  role="treeitem"
                  class="pptx-selection-pane__row"
                  data-node-id={row.nodeId}
                  data-hidden={row.hidden ? '' : undefined}
                  aria-level={row.level}
                  aria-posinset={row.position}
                  aria-setsize={row.siblings}
                  aria-selected={sel.has(row.nodeId)}
                  aria-expanded={row.branch ? row.expanded : undefined}
                  aria-label={
                    row.hidden ? t('selectionPane.hiddenName', { name: row.name }) : row.name
                  }
                  aria-description={t(`panel.format.elementKind.${row.kind}`)}
                  aria-keyshortcuts={canEdit ? 'H F2' : undefined}
                  tabindex={row.nodeId === stop ? 0 : -1}
                  style={{ paddingInlineStart: `${(row.level - 1) * 16 + 4}px` }}
                  onFocus={() => (focused.value = row.nodeId)}
                  onKeydown={(e: KeyboardEvent) => onKeydown(e, row)}
                  onMousedown={(e: MouseEvent) => {
                    // A press on the row selects; the row keeps the focus, not the slide.
                    if ((e.target as HTMLElement).closest('input, button')) return;
                    e.preventDefault();
                    (e.currentTarget as HTMLElement).focus();
                    select(row, e);
                  }}
                  onDblclick={(e: MouseEvent) => {
                    if (!(e.target as HTMLElement).closest('input, button') && canEdit)
                      renaming.value = row.nodeId;
                  }}
                >
                  {row.branch ? (
                    <span
                      class="pptx-selection-pane__twisty"
                      aria-hidden="true"
                      onMousedown={(e: MouseEvent) => {
                        e.preventDefault();
                        e.stopPropagation();
                        toggle(row.nodeId, !row.expanded);
                      }}
                    >
                      <Icon name={row.expanded ? 'expand_more' : 'chevron_right'} />
                    </span>
                  ) : (
                    <span class="pptx-selection-pane__twisty" aria-hidden="true" />
                  )}
                  {renaming.value === row.nodeId ? (
                    <RenameField
                      name={row.name}
                      label={t('selectionPane.rename')}
                      onDone={(name: string | null, refocus: boolean) =>
                        renamed(row, name, refocus)
                      }
                    />
                  ) : (
                    <span class="pptx-selection-pane__name">{row.name}</span>
                  )}
                  <button
                    type="button"
                    class="pptx-selection-pane__eye"
                    tabindex={-1}
                    aria-label={eye}
                    title={eye}
                    disabled={!canEdit}
                    onMousedown={keepFocus}
                    onClick={() => setVisible([row.nodeId], row.hidden)}
                  >
                    <Icon name={row.hidden ? 'visibility_off' : 'visibility'} />
                  </button>
                </div>
              );
            })}
          </div>
        </div>
      );
    };
  },
});

/**
 * The selection pane: select, show or hide, rename, and reorder the elements of the current
 * slide, the twin of React's `PptxEditor.SelectionPane`. Render it inside the editor root, or
 * open the side panel's **Selection** tab. @public
 */
export const PptxEditorSelectionPane = defineComponent({
  name: 'PptxEditorSelectionPane',
  props: selectionPaneProps,
  setup(props) {
    // Node ids repeat across slides: a new slide starts a new pane, so no rename, focus, or
    // open group carries over to an element of another slide.
    const slideId = useEditorState((s) => s.slides[s.currentSlide]?.id ?? '');
    return () => h(SlidePane, { key: slideId.value, className: props.className });
  },
});
