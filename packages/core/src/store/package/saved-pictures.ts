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

/** Cap reference collection and searches. Exhausting the budget retains image relationships. */
const MAX_REFERENCE_WORK = 4 * 1024 * 1024;
const REL_ATTRIBUTES = new Set(['Id', 'Type', 'Target', 'TargetMode']);
const NAME_CHARACTER = /[\p{L}\p{N}\p{M}_.-]/u;

/** Recognize a complete id inside extension syntax without confusing rId1 with rId10. */
function containsId(value: string, id: string): boolean {
  for (let at = value.indexOf(id); at >= 0; at = value.indexOf(id, at + id.length)) {
    const before = value[at - 1];
    const after = value[at + id.length];
    if ((!before || !NAME_CHARACTER.test(before)) && (!after || !NAME_CHARACTER.test(after)))
      return true;
  }
  return false;
}

/** Keep ids embedded in unknown attribute or text syntax, such as extension JSON. */
function imageReferences(root: Element, references: number): ((id: string) => boolean) | null {
  const values = new Set<string>();
  const pending: Node[] = [root];
  let remaining = MAX_REFERENCE_WORK;
  const collect = (value: string): boolean => {
    remaining -= value.length + 1;
    if (remaining < 0) return false;
    values.add(value);
    return true;
  };
  while (pending.length) {
    if (--remaining < 0) return null;
    const node = pending.pop()!;
    if (node.type === 'element') {
      for (const a of node.attrs) if (!collect(a.value)) return null;
      for (const child of node.children) pending.push(child);
    } else if (!collect(node.value)) return null;
  }
  // Decide the budget before dropping anything. Attribute order cannot change which
  // relationships survive on replicas with equivalent source XML.
  if ((MAX_REFERENCE_WORK - remaining) * references > MAX_REFERENCE_WORK) return null;
  return (id) => {
    if (values.has(id)) return true;
    for (const value of values) if (containsId(value, id)) return true;
    return false;
  };
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
    if (
      source === null ||
      part.kind !== 'xml' ||
      part.xml.root.ns !== NS.rel ||
      part.xml.root.local !== 'Relationships' ||
      part.xml.root.attrs.length > 0
    )
      return pkg.parts;
    const owner = pkg.parts.get(equivalentPartName(pkg, source) ?? source);
    const relationships = elements(part.xml.root);
    const used =
      owner?.kind === 'xml' && owner.xml.root.ns === NS.p && OWNERS.has(owner.xml.root.local)
        ? imageReferences(owner.xml.root, relationships.length)
        : null;
    const dropped = new Set<string>();
    for (const rel of relationships) {
      if (rel.ns !== NS.rel || rel.local !== 'Relationship') return pkg.parts;
      // Extension metadata can also name another relationship. Preserve the whole package
      // when its meaning is unknown, including metadata on external relationships.
      if (
        rel.children.length > 0 ||
        rel.attrs.some((a) => a.ns !== '' || !REL_ATTRIBUTES.has(a.local))
      )
        return pkg.parts;
      if (attr(rel, 'TargetMode') === 'External') continue;
      const target = resolveInternalTarget(
        source === '/' ? '/x' : source,
        attr(rel, 'Target') ?? ''
      );
      if (!target.ok) return pkg.parts;
      const id = attr(rel, 'Id');
      if (used && id && attr(rel, 'Type') === REL.image && !used(id)) dropped.add(rel.id);
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
  for (const [name, part] of pkg.parts) {
    // Parsed package metadata is never image data, even under an incorrect content type.
    if (part.kind !== 'binary') continue;
    if (!contentTypeOf(pkg, name)?.toLowerCase().startsWith('image/')) continue;
    if (
      referenced.has(partNameKey(name)) ||
      equivalentPartName(pkg, relsPartName(name)) !== undefined
    )
      continue;
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
