// Remove unused image data from the export view only. The live package retains undo state.
import {
  asciiFold,
  normalizePartName,
  partNameKey,
  resolveInternalTarget,
} from '@eigenpal/ooxml-core';
import {
  CONTENT_TYPES_PART,
  contentTypeOf,
  equivalentPartName,
  relsPartName,
  type PptxPackage,
} from './package.ts';
import { NS, REL } from './namespaces.ts';
import { attr, elements } from './query.ts';
import type { Element, Node } from './tree.ts';

const OWNERS = new Set(['sld', 'sldLayout', 'sldMaster', 'notes', 'notesMaster', 'handoutMaster']);

/** Keep references from every namespace, including extension content and text values. */
function referencedIds(root: Element): Set<string> {
  const values = new Set<string>();
  const pending: Node[] = [root];
  while (pending.length) {
    const node = pending.pop()!;
    if (node.type === 'element') {
      for (const a of node.attrs) for (const token of a.value.split(/\s+/)) values.add(token);
      for (const child of node.children) pending.push(child);
    } else if (node.type === 'text') {
      for (const token of node.value.split(/\s+/)) values.add(token);
    }
  }
  return values;
}

/**
 * Project save-only relationship and image cleanup without changing the store or room.
 * Unknown relationship types and opaque owners retain their targets. Unreadable relationship
 * lists or internal targets disable cleanup, because they can name any image in the package.
 */
export function savedPictureParts(pkg: PptxPackage): PptxPackage['parts'] {
  const parts = new Map(pkg.parts);
  const referenced = new Set<string>();
  for (const [name, part] of pkg.parts) {
    if (!asciiFold(name).endsWith('.rels')) continue;
    const match = /^(.*)\/_rels\/([^/]+)\.rels$/i.exec(name);
    const source =
      asciiFold(name) === '/_rels/.rels' ? '/' : match ? `${match[1]}/${match[2]}` : null;
    if (source === null || part.kind !== 'xml' || part.xml.root.ns !== NS.rel) return pkg.parts;
    const owner = pkg.parts.get(equivalentPartName(pkg, source) ?? source);
    const used =
      owner?.kind === 'xml' && owner.xml.root.ns === NS.p && OWNERS.has(owner.xml.root.local)
        ? referencedIds(owner.xml.root)
        : null;
    const dropped = new Set<string>();
    for (const rel of elements(part.xml.root)) {
      if (rel.ns !== NS.rel || rel.local !== 'Relationship') return pkg.parts;
      if (attr(rel, 'TargetMode') === 'External') continue;
      const target = resolveInternalTarget(
        source === '/' ? '/x' : source,
        attr(rel, 'Target') ?? ''
      );
      if (!target.ok) return pkg.parts;
      const id = attr(rel, 'Id');
      if (used && id && attr(rel, 'Type') === REL.image && !used.has(id)) dropped.add(rel.id);
      else referenced.add(partNameKey(target.partName));
    }
    if (dropped.size)
      parts.set(name, {
        ...part,
        xml: {
          ...part.xml,
          root: {
            ...part.xml.root,
            children: part.xml.root.children.filter((n) => !dropped.has(n.id)),
          },
        },
      });
  }
  const removed = new Set<string>();
  for (const name of pkg.parts.keys()) {
    if (!contentTypeOf(pkg, name)?.toLowerCase().startsWith('image/')) continue;
    if (referenced.has(partNameKey(name)) || pkg.parts.has(relsPartName(name))) continue;
    parts.delete(name);
    removed.add(partNameKey(name));
  }
  const types = parts.get(CONTENT_TYPES_PART);
  if (removed.size && types?.kind === 'xml')
    parts.set(CONTENT_TYPES_PART, {
      ...types,
      xml: {
        ...types.xml,
        root: {
          ...types.xml.root,
          children: types.xml.root.children.filter(
            (n) =>
              n.type !== 'element' ||
              n.ns !== NS.ct ||
              n.local !== 'Override' ||
              !removed.has(overrideKey(n))
          ),
        },
      },
    });
  return parts;
}

/** Content type override names use the same URI decoding as the package reader. */
function overrideKey(node: Element): string {
  const normalized = normalizePartName(attr(node, 'PartName') ?? '');
  return normalized.ok ? partNameKey(normalized.partName) : '';
}
