// Connector routes and glued ends, as `setTransforms` carries them (`TransformItem.connector`).
// DOM-free, declaration-only.

/**
 * A connector's route and the shapes its ends glue to. Each field is optional: an absent field
 * keeps the connector's current value. @public
 */
export interface ConnectorRoute {
  /**
   * The connector preset (`a:prstGeom/@prst`): `straightConnector1`, `bentConnector2` to
   * `bentConnector5`, or `curvedConnector2` to `curvedConnector5`.
   */
  readonly preset?: string;
  /** The adjust values of the preset by guide name, in guide units (see `a:gd`). */
  readonly adjust?: Readonly<Record<string, number>>;
  /** The connection site the start glues to (`a:stCxn`); null unglues it. */
  readonly start?: ConnectionTarget | null;
  /** The connection site the end glues to (`a:endCxn`); null unglues it. */
  readonly end?: ConnectionTarget | null;
}

/** A connection site of a shape: the shape's node id and the site's index. @public */
export interface ConnectionTarget {
  /** The node id of the shape, as `EditorSelection` names shapes. */
  readonly nodeId: string;
  /** The index of the site in the shape's geometry (`a:cxnLst`), from 0. */
  readonly index: number;
}
