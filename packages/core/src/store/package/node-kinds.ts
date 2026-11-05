// How a typed node kind is derived. One rule for the reader, for nodes an op creates, and for
// nodes an op moves, so a node has the same kind in the store as after save and reopen.

import { NS } from './namespaces.ts';
import type { NodeKind } from './tree.ts';

const P = NS.p;
const A = NS.a;

/** Assign a typed kind from the expanded name and the parent's kind. */
export function kindOf(ns: string, local: string, parent: NodeKind | null): NodeKind {
  if (ns === P) {
    switch (local) {
      case 'spTree':
        return 'spTree';
      case 'grpSp':
        return parent === 'spTree' || parent === 'grpSp' ? 'grpSp' : 'generic';
      case 'sp':
        return parent === 'spTree' || parent === 'grpSp' ? 'sp' : 'generic';
      case 'pic':
        return parent === 'spTree' || parent === 'grpSp' ? 'pic' : 'generic';
      case 'cxnSp':
        return parent === 'spTree' || parent === 'grpSp' ? 'cxnSp' : 'generic';
      case 'graphicFrame':
        return parent === 'spTree' || parent === 'grpSp' ? 'graphicFrame' : 'generic';
      case 'txBody':
        return 'txBody';
    }
    return 'generic';
  }
  if (ns === A) {
    switch (local) {
      case 'txBody':
        return 'txBody';
      case 'p':
        return parent === 'txBody' ? 'paragraph' : 'generic';
      case 'r':
        return parent === 'paragraph' ? 'run' : 'generic';
      case 'br':
        return parent === 'paragraph' ? 'break' : 'generic';
      case 'fld':
        return parent === 'paragraph' ? 'field' : 'generic';
      case 't':
        return parent === 'run' || parent === 'field' ? 'text' : 'generic';
    }
  }
  return 'generic';
}

/**
 * Markup-compatibility wrappers are transparent for typing: a `p:sp` inside `mc:Fallback` is
 * still a shape-tree shape. The wrapper itself is `generic`.
 */
export function isTransparentWrapper(ns: string, local: string): boolean {
  return (
    ns === NS.mc && (local === 'AlternateContent' || local === 'Choice' || local === 'Fallback')
  );
}

/** The kind of a node, and the kind context its children are typed under. */
export function deriveKind(
  ns: string,
  local: string,
  context: NodeKind | null
): { readonly kind: NodeKind; readonly childContext: NodeKind | null } {
  if (isTransparentWrapper(ns, local)) return { kind: 'generic', childContext: context };
  const kind = kindOf(ns, local, context);
  return { kind, childContext: kind };
}
