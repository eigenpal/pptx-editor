// The diagram types of the public contract: the selected diagram as the snapshot reports it, and
// the commands that add and delete diagram shapes. DOM-free.

/**
 * Where `addDiagramShape` puts the new shape: after or before the target node at its level, or
 * below it as its last child. @public
 */
export type DiagramShapePosition = 'after' | 'before' | 'below';

/** The selected diagram, while one diagram frame is selected or holds the caret. @public */
export interface SelectedDiagramState {
  /** The node id of the diagram's graphic frame. */
  readonly nodeId: string;
  /** The diagram node whose text holds the caret (`dgm:pt/@modelId`), or null. */
  readonly node: string | null;
  /**
   * Whether `addDiagramShape` and `deleteDiagramShape` can change this diagram: its saved
   * drawing reads as a line of blocks that the editor can lay out again.
   */
  readonly structure: boolean;
  /** Why the shapes cannot be added or deleted, as a translation key, when `structure` is false. */
  readonly reason?: string;
}

/** The diagram commands. @public */
export type DiagramCommand =
  | {
      /**
       * Add an empty shape to the selected diagram, or to the diagram frame `nodeId` on the
       * current slide: after or before the target node at its level, or below it as its last
       * child. The target is `node`, else the node that holds the caret; with only the frame
       * selected, `after` adds after the last top-level node and `before` before the first.
       * The data model gets the new node with its connection, and the saved drawing is laid
       * out again. The caret goes into the new shape. Refused with
       * `errors.diagram.layoutUnsupported` for a diagram whose drawing the editor cannot lay out.
       */
      readonly type: 'addDiagramShape';
      readonly position: DiagramShapePosition;
      readonly nodeId?: string;
      readonly node?: string;
    }
  | {
      /**
       * Delete the target node of the selected diagram (`node`, else the node that holds the
       * caret), with its children, from the data model and the saved drawing. A diagram keeps
       * at least one top-level node.
       */
      readonly type: 'deleteDiagramShape';
      readonly nodeId?: string;
      readonly node?: string;
    };
