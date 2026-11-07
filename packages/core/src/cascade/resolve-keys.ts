// Cache keys of the slide resolve. The package is persistent: an edit replaces only the
// elements on the path it changes. A resolved element stays valid while its XML element keeps
// its object and everything else it reads is the same object too. These keys hold that
// "everything else", compared by identity.

import { relationshipsOf, relsPartName, type Element, type PptxPackage } from '../store/index.ts';
import type { Layer } from './types.ts';

/** The content of a part by identity: the root of an XML part, the bytes of a binary one. */
export function partContent(pkg: PptxPackage, name: string | undefined): unknown {
  const part = name === undefined ? undefined : pkg.parts.get(name);
  return part && (part.kind === 'xml' ? part.xml.root : part.bytes);
}

/** A part's relationships and their internal targets (pictures, charts, diagrams). */
export function relsKey(pkg: PptxPackage, partName: string | undefined): unknown[] {
  if (partName === undefined) return [];
  const internal = relationshipsOf(pkg, partName).filter((rel) => !rel.external);
  return [
    partContent(pkg, relsPartName(partName)),
    ...internal.map((r) => partContent(pkg, r.target)),
  ];
}

/**
 * What an element of `partName` reads besides its own XML: the slide's shared entries, the
 * layer, the part's namespace declarations (alternate content), and its relationships.
 */
export function layerKey(
  pkg: PptxPackage,
  layer: Layer,
  partName: string,
  root: Element | undefined,
  shared: readonly unknown[]
): unknown[] {
  const declared =
    root &&
    [
      ...root.nsDecls.map((d) => `xmlns:${d.prefix}=${d.uri}`),
      ...root.attrs.map((a) => `${a.ns} ${a.local}=${a.value}`),
    ].join(' ');
  return [...shared, layer, partName, declared, ...relsKey(pkg, partName)];
}

/** Two keys hold the same objects in the same order. */
export const sameKey = (a: readonly unknown[], b: readonly unknown[]): boolean =>
  a.length === b.length && a.every((v, i) => v === b[i]);
