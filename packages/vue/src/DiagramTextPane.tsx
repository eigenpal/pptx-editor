/** @jsxImportSource vue */
// The Diagram tab of the side panel: the Vue twin of the React adapter's `DiagramTextPane.tsx`,
// with the same markup, classes, and ARIA. A field applies when the user presses Enter or leaves
// it, as one undo step (`setDiagramText`).

import { defineComponent, ref, useId, type PropType } from 'vue';
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
    disabled: { type: Boolean, required: true },
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
            disabled={props.disabled}
            data-draft={draft.value !== null}
            onInput={(e: Event) => {
              draft.value = (e.target as HTMLTextAreaElement).value;
            }}
            onBlur={commit}
            onKeydown={(e: KeyboardEvent) => {
              if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) {
                e.preventDefault();
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
    const refusal = ref<string | null>(null);
    const noteId = useId();
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
      /** Runs an add or a delete; the tab announces a refusal, as no field shows it. */
      const run = (command: EditorCommand): void => {
        refusal.value = apply(command);
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
        <div class="pptx-diagram-pane">
          <p class="pptx-chart-pane__note">{t('diagram.pane.help')}</p>
          {!selected.structure && selected.reason ? (
            <p class="pptx-chart-pane__note">{t(selected.reason as TranslationKey)}</p>
          ) : null}
          <p class="pptx-chart-pane__error" role="alert">
            {refusal.value ?? ''}
          </p>
          {selected.nodes.length === 0 ? (
            <p class="pptx-chart-pane__note">{t('diagram.pane.noNodes')}</p>
          ) : (
            <ul class="pptx-diagram-pane__list">
              {selected.nodes.map((node, i) => (
                <li
                  key={node.node}
                  class="pptx-diagram-pane__row"
                  style={{ '--pptx-diagram-level': String(Math.min(node.level, 4)) }}
                >
                  <NodeField
                    label={t('diagram.pane.nodeLabel', { number: i + 1 })}
                    value={node.text}
                    disabled={readOnly.value || !node.editable}
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
                        run({
                          type: 'addDiagramShape',
                          position: 'after',
                          nodeId: frame,
                          node: node.node,
                        })
                      }
                    />
                    <ToolButton
                      label={t('diagram.pane.addBelowNamed', { name: nameOf(node) })}
                      icon="add_row_below"
                      disabled={!structure || !canBelow(node)}
                      onClick={() =>
                        run({
                          type: 'addDiagramShape',
                          position: 'below',
                          nodeId: frame,
                          node: node.node,
                        })
                      }
                    />
                    <ToolButton
                      label={t('diagram.pane.deleteNamed', { name: nameOf(node) })}
                      icon="delete"
                      disabled={!structure || (node.level === 0 && topLevel <= 1)}
                      onClick={() =>
                        run({ type: 'deleteDiagramShape', nodeId: frame, node: node.node })
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
            onClick={() => run({ type: 'addDiagramShape', position: 'after', nodeId: frame })}
          >
            {t('diagram.pane.addShape')}
          </button>
        </div>
      );
    };
  },
});
