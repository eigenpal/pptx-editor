// The `diagram` chrome group: Edit Text, Add Shape After, Add Shape Before, Add Shape Below, and
// Delete Shape for the selected diagram. The toolbar shows it while a diagram frame is selected
// or a diagram shape holds the caret (`EditorSnapshot.diagram`), as the Diagram Design tab of
// the ribbon. Edit Text shows the Diagram tab of the side panel (`requestPane`). Each other
// control runs one fixed command, so the engine's probe of that command (`Editor.can`) turns it
// on and off. Kept apart from chrome-controls.ts, which spreads it into `CHROME_GROUPS`.

import type { Editor, EditorCommand } from '../contracts/index.ts';
import type { ChromeGroup, ChromeSlotId, ToolbarCommandState } from './chrome-controls.ts';
import { ICON_PATHS } from './icon-paths.ts';
import { hasPaneListener } from './pane-requests.ts';

const icon = (name: string): readonly string[] => ICON_PATHS[name] ?? [];

/** The `diagram` group. @internal */
export const DIAGRAM_CHROME_GROUP: ChromeGroup = {
  id: 'diagram',
  labelKey: 'toolbar.groups.diagram',
  contextual: true,
  controls: [
    {
      id: 'editText',
      labelKey: 'diagram.editText',
      paths: icon('edit_note'),
      state: { kind: 'pane', pane: 'diagram' },
    },
    {
      id: 'addAfter',
      labelKey: 'diagram.addAfter',
      paths: icon('add_column_right'),
      state: { kind: 'command' },
    },
    {
      id: 'addBefore',
      labelKey: 'diagram.addBefore',
      paths: icon('add_column_left'),
      state: { kind: 'command' },
    },
    {
      id: 'addBelow',
      labelKey: 'diagram.addBelow',
      paths: icon('add_row_below'),
      state: { kind: 'command' },
    },
    {
      id: 'delete',
      labelKey: 'diagram.delete',
      paths: icon('delete'),
      state: { kind: 'command' },
    },
  ],
};

/** The `diagram` slot ids. @public */
export type DiagramSlotId =
  | 'diagram.editText'
  | 'diagram.addAfter'
  | 'diagram.addBefore'
  | 'diagram.addBelow'
  | 'diagram.delete';

/** The command of each `diagram` command slot. @internal */
export const DIAGRAM_SLOT_COMMANDS: Readonly<
  Record<Exclude<DiagramSlotId, 'diagram.editText'>, EditorCommand>
> = {
  'diagram.addAfter': { type: 'addDiagramShape', position: 'after' },
  'diagram.addBefore': { type: 'addDiagramShape', position: 'before' },
  'diagram.addBelow': { type: 'addDiagramShape', position: 'below' },
  'diagram.delete': { type: 'deleteDiagramShape' },
};

/**
 * The state of a `diagram` pane slot, or undefined for another slot. Edit Text is on while a
 * diagram is selected and a side panel or the host listens for pane requests. The tab shows the
 * text read-only in view mode. @internal
 */
export function diagramSlotState(
  editor: Editor,
  id: ChromeSlotId
): ToolbarCommandState | undefined {
  if (id !== 'diagram.editText') return undefined;
  const snap = editor.snapshot();
  const reason =
    snap.status !== 'ready'
      ? 'disabled.noEditor'
      : !snap.diagram
        ? 'disabled.noDiagram'
        : hasPaneListener(editor)
          ? null
          : 'disabled.noSidePanel';
  return { id, enabled: reason === null, disabledReason: reason, active: false };
}
