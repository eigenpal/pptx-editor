// Handle minting and resolution, after DOCX Editor's `packages/core/src/automation/handles.ts`.
// DOCX Editor commit adf0e178.
//
// INTERNAL. A handle is a name this table invented, mapped privately to the presentation
// identity it stands for. Consumers get the name; nothing else leaves.
//
// STABLE: the same object asked for twice yields the same ref, so an object model can hold a
// reference across batches and compare two references for identity.
//
// IDENTITY THAT SURVIVES EDITS: a slide is its `p:sldId/@id`, a layout or master its part, and a
// shape its owner and `p:cNvPr/@id`. None of them changes when other slides or shapes move, so a
// ref names the same object after a reorder. A tree node id would not survive a remote edit.
//
// HOST-SCOPED: every ref carries a token drawn from the platform CSPRNG when the table is
// created, so a ref obtained from one host is `invalid-handle` on every other host.
//
// Lookup is by `Map`, never by object key: a forged ref is untrusted input.

import type { AutomationHandle, AutomationHandleRef, AutomationObjectKind } from './protocol.ts';

/** The owner of a shape tree. */
export type AutomationOwner =
  | { readonly kind: 'slide'; readonly sldId: string }
  | { readonly kind: 'slideLayout' | 'slideMaster'; readonly part: string };

/** What a handle stands for. */
export type AutomationTarget =
  | { readonly kind: 'presentation' }
  | { readonly kind: 'slide'; readonly sldId: string }
  | { readonly kind: 'slideLayout'; readonly part: string }
  | { readonly kind: 'slideMaster'; readonly part: string }
  | { readonly kind: 'shape'; readonly owner: AutomationOwner; readonly shapeId: string };

/**
 * 128 bits of hex from the platform CSPRNG. Fails closed: a runtime without
 * `globalThis.crypto.getRandomValues` throws rather than falling back to a guessable token.
 */
function hostToken(): string {
  const source = (globalThis as { crypto?: { getRandomValues?: unknown } }).crypto;
  const fill = source?.getRandomValues;
  if (typeof fill !== 'function')
    throw new Error(
      'automation: no secure random source. globalThis.crypto.getRandomValues is required to ' +
        'scope presentation handles to one host.'
    );
  const bytes = (fill as (array: Uint8Array) => Uint8Array).call(source, new Uint8Array(16));
  let hex = '';
  for (const byte of bytes) hex += byte.toString(16).padStart(2, '0');
  return hex;
}

function ownerKey(owner: AutomationOwner): string {
  return owner.kind === 'slide' ? `slide:${owner.sldId}` : `${owner.kind}:${owner.part}`;
}

function targetKey(target: AutomationTarget): string {
  switch (target.kind) {
    case 'presentation':
      return 'presentation';
    case 'slide':
      return `slide:${target.sldId}`;
    case 'slideLayout':
    case 'slideMaster':
      return `${target.kind}:${target.part}`;
    case 'shape':
      return `shape:${ownerKey(target.owner)}#${target.shapeId}`;
  }
}

/** The handle table of one host. */
export interface HandleTable {
  /** The handle for a target: the same ref every time for the same object. */
  mint<T extends AutomationTarget>(target: T): AutomationHandle<T['kind']>;
  /** The target of a ref this table minted, or null for anything else. */
  resolve(handle: unknown): AutomationTarget | null;
  /** Forget every ref (dispose). */
  clear(): void;
}

/** A handle table with its own host token. */
export function createHandleTable(): HandleTable {
  const token = hostToken();
  const byKey = new Map<string, AutomationHandle>();
  const byRef = new Map<string, AutomationTarget>();
  let next = 1;
  return {
    mint(target) {
      const key = targetKey(target);
      const known = byKey.get(key);
      if (known) return known as never;
      const ref = `${token}:${target.kind}:${next}` as AutomationHandleRef;
      next += 1;
      const handle = Object.freeze({ kind: target.kind as AutomationObjectKind, ref });
      byKey.set(key, handle);
      byRef.set(ref, Object.freeze({ ...target }) as AutomationTarget);
      return handle as never;
    },
    resolve(handle) {
      if (typeof handle !== 'object' || handle === null) return null;
      const { ref, kind } = handle as { ref?: unknown; kind?: unknown };
      if (typeof ref !== 'string') return null;
      const target = byRef.get(ref);
      if (!target || target.kind !== kind) return null;
      return target;
    },
    clear() {
      byKey.clear();
      byRef.clear();
    },
  };
}
