/** @jsxImportSource react */
// The Diagram tab of the side panel: the nodes of the selected diagram in outline order, each
// with its text field and its add and delete buttons. A field applies when the user presses
// Enter or leaves it, as one undo step (`setDiagramText`), so the slide updates beside the
// panel. The Vue twin (`packages/vue/src/DiagramTextPane.tsx`) shows the same rows.

import { useId, useState, type CSSProperties } from 'react';
import type { DiagramNodeState, EditorCommand } from '@pptx-editor.dev/core/contracts';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { ToolButton } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';

/**
 * The text field of one node: it applies on Enter (Shift+Enter starts a paragraph) or blur and
 * shows the node's text otherwise. Text that the editor refuses stays with its reason; Escape
 * restores the text.
 */
function NodeField({
  label,
  value,
  disabled,
  note,
  onCommit,
}: {
  label: string;
  value: string;
  /** The id of a note that describes the field, if any. */
  note?: string;
  disabled: boolean;
  /** Applies the text: null when it applied, else the reason to show. */
  onCommit: (text: string) => string | null;
}) {
  const errorId = useId();
  const [draft, setDraft] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
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
  const shown = draft ?? value;
  return (
    <label className="pptx-field pptx-diagram-pane__field">
      <textarea
        aria-label={label}
        aria-invalid={error !== null}
        aria-describedby={
          [error !== null ? errorId : null, note].filter(Boolean).join(' ') || undefined
        }
        rows={Math.min(6, shown.split('\n').length)}
        value={shown}
        disabled={disabled}
        // Ctrl/Cmd+Z in a text area is the field's own undo (`panelHistoryKey`).
        data-draft={draft !== null}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) {
            e.preventDefault();
            commit();
          }
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

/** The Diagram tab: the text of the selected diagram's nodes, or why it shows none. */
export function DiagramTextPane() {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const diagram = useEditorState((s) => s.diagram ?? null);
  const readOnly = useEditorState((s) => s.mode === 'view');
  const [refusal, setRefusal] = useState<string | null>(null);
  const noteId = useId();
  if (!diagram)
    return (
      <div className="pptx-empty">
        <p className="pptx-empty__title">{t('diagram.pane.emptyTitle')}</p>
        <p className="pptx-empty__body">{t('diagram.pane.emptyBody')}</p>
      </div>
    );
  const frame = diagram.nodeId;
  /** Runs one command: null when it applied, else the reason. */
  const apply = (command: EditorCommand): string | null => {
    const result = editor?.exec(command);
    return !result || result.ok ? null : t(result.reason as TranslationKey);
  };
  /** Runs an add or a delete; the tab announces a refusal, as no field shows it. */
  const run = (command: EditorCommand): void => setRefusal(apply(command));
  const canBelow = (node: DiagramNodeState): boolean =>
    !!editor?.can({ type: 'addDiagramShape', position: 'below', nodeId: frame, node: node.node })
      .ok;
  const structure = diagram.structure && !readOnly;
  const topLevel = diagram.nodes.filter((n) => n.level === 0).length;
  // A button names its node by the start of its text, so a long text keeps the name short.
  const nameOf = (node: DiagramNodeState): string => {
    const text = node.text.replace(/\s+/g, ' ').trim();
    return text.length > 40 ? `${text.slice(0, 40)}…` : text || t('diagram.pane.emptyShape');
  };
  return (
    <div className="pptx-diagram-pane">
      <p className="pptx-chart-pane__note">{t('diagram.pane.help')}</p>
      {!diagram.structure && diagram.reason ? (
        <p className="pptx-chart-pane__note">{t(diagram.reason as TranslationKey)}</p>
      ) : null}
      <p className="pptx-chart-pane__error" role="alert">
        {refusal ?? ''}
      </p>
      {diagram.nodes.length === 0 ? (
        <p className="pptx-chart-pane__note">{t('diagram.pane.noNodes')}</p>
      ) : (
        <ul className="pptx-diagram-pane__list">
          {diagram.nodes.map((node, i) => (
            <li
              key={node.node}
              className="pptx-diagram-pane__row"
              style={{ '--pptx-diagram-level': Math.min(node.level, 4) } as CSSProperties}
            >
              <NodeField
                label={t('diagram.pane.nodeLabel', { number: i + 1 })}
                value={node.text}
                disabled={readOnly || !node.editable}
                {...(node.editable ? {} : { note: `${noteId}-${i}` })}
                onCommit={(text) =>
                  apply({ type: 'setDiagramText', nodeId: frame, node: node.node, text })
                }
              />
              <div className="pptx-diagram-pane__actions">
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
                <p id={`${noteId}-${i}`} className="pptx-chart-pane__note">
                  {t('diagram.pane.readOnly')}
                </p>
              ) : null}
            </li>
          ))}
        </ul>
      )}
      <button
        type="button"
        className="pptx-button"
        disabled={!structure}
        onClick={() => run({ type: 'addDiagramShape', position: 'after', nodeId: frame })}
      >
        {t('diagram.pane.addShape')}
      </button>
    </div>
  );
}
