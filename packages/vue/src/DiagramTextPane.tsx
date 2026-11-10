/** @jsxImportSource vue */
// The Diagram tab of the side panel: the Vue twin of the React adapter's `DiagramTextPane.tsx`,
// with the same markup, classes, and ARIA. A field applies when the user presses Enter or leaves
// it, as one undo step (`setDiagramText`).

import { defineComponent, nextTick, ref, useId, type PropType } from 'vue';
import type { DiagramNodeState, EditorCommand } from '@pptx-editor.dev/core/contracts';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { ToolButton } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';

/**
 * The text field of one node: it applies on Enter (Shift+Enter starts a paragraph) or blur and
 * shows the node's text otherwise. Text that the editor refuses stays with its reason; Escape
 * restores the text.
 */
const NodeField = defineComponent({
  name: 'PptxDiagramNodeField',
  props: {
    label: { type: String, required: true },
    value: { type: String, required: true },
    /** The node's data model id, so the tab can focus the field after an add or a delete. */
    node: { type: String, required: true },
    /** Shown but not editable: the text stays readable, focusable, and copyable. */
    readOnly: { type: Boolean, required: true },
    /** The id of a note that describes the field, if any. */
    note: { type: String, default: undefined },
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
    const commit = (): void => {
      const d = draft.value;
      if (d === null) return;
      if (d === props.value) return reset();
      const refused = props.onCommit(d);
      if (refused === null) reset();
      else error.value = refused;
    };
    return () => {
      const shown = draft.value ?? props.value;
      return (
        <label class="pptx-field pptx-diagram-pane__field">
          <textarea
            aria-label={props.label}
            aria-invalid={error.value !== null}
            aria-describedby={
              [error.value !== null ? errorId : null, props.note].filter(Boolean).join(' ') ||
              undefined
            }
            rows={Math.min(6, shown.split('\n').length)}
            value={shown}
            readonly={props.readOnly}
            data-node={props.node}
            data-draft={draft.value !== null}
            onInput={(e: Event) => {
              draft.value = (e.target as HTMLTextAreaElement).value;
            }}
            onBlur={commit}
            onKeydown={(e: KeyboardEvent) => {
              if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) {
                e.preventDefault();
                e.stopPropagation();
                commit();
              }
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
    };
  },
});

/** The Diagram tab: the text of the selected diagram's nodes, or why it shows none. */
export const DiagramTextPane = defineComponent({
  name: 'PptxDiagramTextPane',
  setup() {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const diagram = useEditorState((s) => s.diagram ?? null);
    const readOnly = useEditorState((s) => s.mode === 'view');
    // The last add or delete refusal, for the diagram it came from. Each one gets a new count,
    // so the alert remounts and a repeated refusal is announced again.
    const refusal = ref<{ text: string; frame: string; count: number } | null>(null);
    const noteId = useId();
    const root = ref<HTMLElement | null>(null);
    /** Focus the new node's field after an add, or the field at `index` after a delete. */
    const focusField = (node: string | undefined, index: number): void => {
      void nextTick(() => {
        const fields = [...(root.value?.querySelectorAll('textarea') ?? [])];
        const field =
          fields.find((f) => node !== undefined && f.dataset.node === node) ??
          fields[Math.min(index, fields.length - 1)];
        field?.focus();
      });
    };
    return () => {
      const selected = diagram.value;
      if (!selected)
        return (
          <div class="pptx-empty">
            <p class="pptx-empty__title">{t('diagram.pane.emptyTitle')}</p>
            <p class="pptx-empty__body">{t('diagram.pane.emptyBody')}</p>
          </div>
        );
      const frame = selected.nodeId;
      /** Runs one command: null when it applied, else the reason. */
      const apply = (command: EditorCommand): string | null => {
        const result = editor.value?.exec(command);
        return !result || result.ok ? null : t(result.reason as TranslationKey);
      };
      /**
       * Runs an add or a delete; the tab announces a refusal, as no field shows it. The focus
       * goes to the new node's field after an add, and to the field at `index` after a delete.
       */
      const run = (command: EditorCommand, index: number): void => {
        const reason = apply(command);
        refusal.value =
          reason === null ? null : { text: reason, frame, count: (refusal.value?.count ?? 0) + 1 };
        if (reason !== null) return;
        const sel = editor.value?.snapshot().selection;
        const added = sel?.kind === 'text' ? sel.diagramNode : undefined;
        focusField(command.type === 'addDiagramShape' ? added : undefined, index);
      };
      const canBelow = (node: DiagramNodeState): boolean =>
        !!editor.value?.can({
          type: 'addDiagramShape',
          position: 'below',
          nodeId: frame,
          node: node.node,
        }).ok;
      const structure = selected.structure && !readOnly.value;
      const topLevel = selected.nodes.filter((n) => n.level === 0).length;
      // A button names its node by the start of its text, so a long text keeps the name short.
      const nameOf = (node: DiagramNodeState): string => {
        const text = node.text.replace(/\s+/g, ' ').trim();
        return text.length > 40 ? `${text.slice(0, 40)}…` : text || t('diagram.pane.emptyShape');
      };
      return (
        <div ref={root} class="pptx-diagram-pane">
          <p class="pptx-chart-pane__note">{t('diagram.pane.help')}</p>
          {!selected.structure && selected.reason ? (
            <p class="pptx-chart-pane__note">{t(selected.reason as TranslationKey)}</p>
          ) : null}
          <p key={refusal.value?.count ?? 0} class="pptx-chart-pane__error" role="alert">
            {refusal.value?.frame === frame ? refusal.value.text : ''}
          </p>
          {selected.nodes.length === 0 ? (
            <p class="pptx-chart-pane__note">{t('diagram.pane.noNodes')}</p>
          ) : (
            <ul class="pptx-diagram-pane__list">
              {selected.nodes.map((node, i) => (
                <li
                  key={`${frame}:${node.node}`}
                  class="pptx-diagram-pane__row"
                  style={{ '--pptx-diagram-level': String(Math.min(node.level, 4)) }}
                >
                  <NodeField
                    label={
                      node.level > 0
                        ? t('diagram.pane.nodeLabelLevel', {
                            number: i + 1,
                            level: node.level + 1,
                          })
                        : t('diagram.pane.nodeLabel', { number: i + 1 })
                    }
                    value={node.text}
                    node={node.node}
                    readOnly={readOnly.value || !node.editable}
                    note={node.editable ? undefined : `${noteId}-${i}`}
                    onCommit={(text: string) =>
                      apply({ type: 'setDiagramText', nodeId: frame, node: node.node, text })
                    }
                  />
                  <div class="pptx-diagram-pane__actions">
                    <ToolButton
                      label={t('diagram.pane.addAfterNamed', { name: nameOf(node) })}
                      icon="add"
                      disabled={!structure}
                      onClick={() =>
                        run(
                          {
                            type: 'addDiagramShape',
                            position: 'after',
                            nodeId: frame,
                            node: node.node,
                          },
                          i + 1
                        )
                      }
                    />
                    <ToolButton
                      label={t('diagram.pane.addBelowNamed', { name: nameOf(node) })}
                      icon="add_row_below"
                      disabled={!structure || !canBelow(node)}
                      onClick={() =>
                        run(
                          {
                            type: 'addDiagramShape',
                            position: 'below',
                            nodeId: frame,
                            node: node.node,
                          },
                          i + 1
                        )
                      }
                    />
                    <ToolButton
                      label={t('diagram.pane.deleteNamed', { name: nameOf(node) })}
                      icon="delete"
                      disabled={!structure || (node.level === 0 && topLevel <= 1)}
                      onClick={() =>
                        run({ type: 'deleteDiagramShape', nodeId: frame, node: node.node }, i)
                      }
                    />
                  </div>
                  {!node.editable ? (
                    <p id={`${noteId}-${i}`} class="pptx-chart-pane__note">
                      {t('diagram.pane.readOnly')}
                    </p>
                  ) : null}
                </li>
              ))}
            </ul>
          )}
          <button
            type="button"
            class="pptx-button"
            disabled={!structure}
            onClick={() =>
              run(
                { type: 'addDiagramShape', position: 'after', nodeId: frame },
                selected.nodes.length
              )
            }
          >
            {t('diagram.pane.addShape')}
          </button>
        </div>
      );
    };
  },
});
