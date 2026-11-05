// The selection pane's list: every element of the current slide's shape tree, top first.
// It reads the store, not the layout, because hidden
// elements do not lay out but stay in the list. A table that concurrent deletes left with no
// column or row (`isEmptiedTable`) is not listed: the canvas does not draw it and the save
// leaves it out, as a local delete of its last column or row deletes it.

import {
  attr,
  child,
  elements,
  find,
  isEmptiedTable,
  nodeIndex,
  NS,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../store/index.ts';
import type { ShapeSummary } from '../contracts/index.ts';
import { contentElements } from '../cascade/index.ts';
import { aspectLocked } from '../automation/commands/arrange-commands.ts';

const KINDS: Readonly<Record<string, ShapeSummary['kind']>> = {
  sp: 'shape',
  pic: 'picture',
  grpSp: 'group',
  cxnSp: 'connector',
  graphicFrame: 'graphicFrame',
};

/** The deepest group nesting the list follows (the reader's cap, AGENTS.md "Security"). */
const MAX_DEPTH = 32;

function summaries(tree: Element, root: Element, depth: number): ShapeSummary[] {
  const out: ShapeSummary[] = [];
  // Markup-compatibility branches as the cascade reads them.
  for (const el of contentElements(tree, root)) {
    const kind = Object.prototype.hasOwnProperty.call(KINDS, el.kind) ? KINDS[el.kind] : undefined;
    if (!kind || (kind === 'graphicFrame' && isEmptiedTable(el))) continue;
    const nv = elements(el).find((c) => c.ns === NS.p && c.local.startsWith('nv'));
    const cNvPr = child(nv, 'p:cNvPr');
    const hidden = attr(cNvPr, 'hidden');
    const kids = kind === 'group' && depth < MAX_DEPTH ? summaries(el, root, depth + 1) : undefined;
    out.push({
      nodeId: el.id,
      name: attr(cNvPr, 'name') ?? '',
      kind,
      hidden: hidden === '1' || hidden === 'true',
      ...(kids ? { children: kids } : {}),
    });
  }
  // Top first: the last element in the tree draws on top.
  return out.reverse();
}

const EMPTY: readonly ShapeSummary[] = [];

/**
 * The elements of a slide part, top first. The function it returns keeps the list it returned
 * last while the part does not change (the snapshot contract).
 */
export function createShapeListSource(): (
  pkg: PptxPackage | undefined,
  part: string | undefined
) => readonly ShapeSummary[] {
  let lastPart: object | undefined;
  let last: readonly ShapeSummary[] = EMPTY;
  return (pkg, part) => {
    const xml = pkg && part ? xmlPartOf(pkg, part) : undefined;
    if (!xml) return EMPTY;
    if (xml === lastPart) return last;
    const tree = find(xml.root, 'p:spTree');
    const next = tree ? summaries(tree, xml.root, 0) : EMPTY;
    lastPart = xml;
    if (JSON.stringify(next) !== JSON.stringify(last)) last = next;
    return last;
  };
}

/** The element `id` of a part, or undefined. */
export function elementIn(
  pkg: PptxPackage | undefined,
  part: string | undefined,
  id: string
): Element | undefined {
  const xml = pkg && part ? xmlPartOf(pkg, part) : undefined;
  const node = xml ? nodeIndex(xml).get(id)?.node : undefined;
  return node?.type === 'element' ? node : undefined;
}

/** Whether the element `id` of a part keeps its aspect ratio (`noChangeAspect`). */
export function aspectLockedIn(
  pkg: PptxPackage | undefined,
  part: string | undefined,
  id: string
): boolean {
  return aspectLocked(elementIn(pkg, part, id));
}
