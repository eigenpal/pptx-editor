// The `diagram` chrome group: Add Shape After, Add Shape Before, Add Shape Below, and Delete
// Shape for the selected diagram. The toolbar shows it while a diagram frame is selected or a
// diagram shape holds the caret (`EditorSnapshot.diagram`), as the Diagram Design tab of the
// ribbon. Each control runs one fixed command, so the engine's probe of that command
// (`Editor.can`) turns it on and off. Kept apart from chrome-controls.ts, which spreads it into
// `CHROME_GROUPS`.

import type { EditorCommand } from '../contracts/index.ts';
import type { ChromeGroup } from './chrome-controls.ts';
import { ICON_PATHS } from './icon-paths.ts';

const icon = (name: string): readonly string[] => ICON_PATHS[name] ?? [];

/** The `diagram` group. @internal */
export const DIAGRAM_CHROME_GROUP: ChromeGroup = {
  id: 'diagram',
  labelKey: 'toolbar.groups.diagram',
  contextual: true,
  controls: [
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
  'diagram.addAfter' | 'diagram.addBefore' | 'diagram.addBelow' | 'diagram.delete';

/** The command of each `diagram` slot. @internal */
export const DIAGRAM_SLOT_COMMANDS: Readonly<Record<DiagramSlotId, EditorCommand>> = {
  'diagram.addAfter': { type: 'addDiagramShape', position: 'after' },
  'diagram.addBefore': { type: 'addDiagramShape', position: 'before' },
  'diagram.addBelow': { type: 'addDiagramShape', position: 'below' },
  'diagram.delete': { type: 'deleteDiagramShape' },
};
