/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Shared helpers of the computed refresh changes (refresh-changes.ts): element access, part
// content and equality, and the collector that keeps every change id unique. DOM-free.

import {
  canonicalForm,
  elements,
  type Element,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import type { ComputedRefreshChange } from '@pptx-editor.dev/core/editor';

/** The most changes one comparison reports. */
export const MAX_CHANGES = 10000;

export const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

/** Elements never change once read, so the canonical form of each one is found once. */
const canonCache = new WeakMap<Element, string>();

/** The canonical form of an element, or '' for none. */
export function canon(el: Element | undefined): string {
  if (!el) return '';
  let hit = canonCache.get(el);
  if (hit === undefined) canonCache.set(el, (hit = canonicalForm(el)));
  return hit;
}

export const child = (el: Element | undefined, local: string): Element | undefined =>
  el ? elements(el).find((c) => c.local === local) : undefined;

export const attr = (el: Element | undefined, local: string): string | undefined =>
  el?.attrs.find((a) => a.ns === '' && a.local === local)?.value;

/** The root element of an XML part, or undefined. */
export function rootOf(pkg: PptxPackage, part: string | undefined): Element | undefined {
  const p = part ? pkg.parts.get(part) : undefined;
  return p?.kind === 'xml' ? p.xml.root : undefined;
}

/** The bytes or canonical XML of a part, for equality. */
export function partContent(
  pkg: PptxPackage,
  part: string | undefined
): string | Uint8Array | undefined {
  const p = part ? pkg.parts.get(part) : undefined;
  if (!p) return undefined;
  return p.kind === 'xml' ? canon(p.xml.root) : p.bytes;
}

/** Whether two part contents are equal. */
export function sameContent(
  a: string | Uint8Array | undefined,
  b: string | Uint8Array | undefined
): boolean {
  if (typeof a === 'string' || typeof b === 'string') return a === b;
  if (!a || !b) return a === b;
  if (a === b) return true;
  if (a.length !== b.length) return false;
  // Compare four bytes at a time when both views are aligned.
  let i = 0;
  if (a.byteOffset % 4 === 0 && b.byteOffset % 4 === 0) {
    const words = a.length >> 2;
    const wa = new Uint32Array(a.buffer, a.byteOffset, words);
    const wb = new Uint32Array(b.buffer, b.byteOffset, words);
    for (let w = 0; w < words; w += 1) if (wa[w] !== wb[w]) return false;
    i = words << 2;
  }
  for (; i < a.length; i += 1) if (a[i] !== b[i]) return false;
  return true;
}

/** Whether two lists of part contents are equal, item by item. */
export function sameList(
  a: readonly (string | Uint8Array | undefined)[],
  b: readonly (string | Uint8Array | undefined)[]
): boolean {
  return a.length === b.length && a.every((x, i) => sameContent(x, b[i]));
}

/** The changes of one comparison. Each id is unique: a repeated key gets a `#n` suffix. */
export class Collector {
  readonly changes: ComputedRefreshChange[] = [];
  truncated = false;
  private readonly ids = new Set<string>();

  add(change: ComputedRefreshChange): void {
    if (this.changes.length >= MAX_CHANGES) {
      this.truncated = true;
      return;
    }
    let id = change.id;
    for (let n = 2; this.ids.has(id); n += 1) id = `${change.id}#${n}`;
    this.ids.add(id);
    this.changes.push(id === change.id ? change : { ...change, id });
  }
}
