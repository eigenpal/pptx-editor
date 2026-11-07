// The clipboard contract: the format that copy writes and paste reads for shapes.

/** A serialized element, structurally the store's `NodeSpec`. @public */
export interface ClipboardNode {
  readonly ns: string;
  readonly local: string;
  readonly prefix?: string;
  readonly attrs?: readonly { ns: string; local: string; prefix?: string; value: string }[];
  readonly nsDecls?: readonly { prefix: string; uri: string }[];
  readonly children?: readonly (ClipboardNode | { readonly text: string })[];
}

/**
 * A relationship a copied shape uses, with the bytes of its target. A chart, a diagram, or an
 * embedded object carries its own relationships in `relationships`. A web link carries its
 * address in `target` and no bytes. A link to a slide names the slide part and carries no
 * bytes; paste keeps it only in the presentation it came from. @public
 */
export interface ClipboardRelationship {
  readonly id: string;
  readonly type: string;
  readonly part: string;
  readonly contentType: string;
  readonly base64: string;
  /** The address of an external link. */
  readonly target?: string;
  /** The relationships of the carried part. */
  readonly relationships?: readonly ClipboardRelationship[];
}

/** The clipboard format for shapes (`application/x-pptx-editor+json`). @public */
export interface ClipboardPayload {
  readonly version: 1;
  /** The session the shapes came from; pasting into it reuses its parts. */
  readonly document: string;
  /** The slide part the shapes came from; pasting onto it offsets them. */
  readonly slidePart: string;
  readonly shapes: readonly ClipboardNode[];
  readonly relationships: readonly ClipboardRelationship[];
}
