// The headless automation host, after DOCX Editor's `packages/core/src/automation/server-host.ts`.
// DOCX Editor commit adf0e178.
//
// It owns what an editor session owns and nothing more: a package opened through the bounded
// OPC/XML reader, one `PptxDocumentStore` over it, and the same `transact` write path the editor
// uses. There is no second model: the store is the document.
//
// UNTRUSTED BYTES. A `.pptx` is a zip of XML an attacker controls, so opening one goes through
// `readPptx` (size and ratio caps, part name checks, XML with no DTD) and a refusal is a typed
// reason. This host fetches nothing, and has no HTML, DOM, or CSS sink: only `.pptx` bytes from
// the normalizing writer, plain values, and opaque handles leave it.

import { partNameKey } from '@eigenpal/ooxml-core';
import type { TextMeasurer } from '../layout/index.ts';
import {
  PptxDocumentStore,
  readPptx,
  writePptx,
  type PackageRejection,
  type ZipLimits,
} from '../store/index.ts';
import { OpBatch } from './commands/op-batch.ts';
import { openStore } from './commands/open-store.ts';
import { packageForSave } from './commands/save-repair.ts';
import { createAutomationHost, type AutomationDocumentPort } from './host.ts';
import type { AutomationImageTarget } from './plan-images.ts';
import type { AutomationHost } from './protocol.ts';

/** How a store host commits and measures. */
export interface StoreAutomationOptions {
  /**
   * The measurer text writes re-fit with (ECMA-376 Part 1, 21.1.2.1). Without one, a text write
   * keeps the shape's size and font scale as they are.
   */
  readonly measurer?: TextMeasurer;
  /** The label of each committed transaction (undo history). Default `automation`. */
  readonly label?: string;
  /**
   * The store is shared with a collaboration room. A table structure edit then marks the
   * table, as the editor's table commands do in a room. Default `false`.
   */
  readonly collaborating?: boolean;
  /**
   * Draws a PNG picture of a slide, or of one shape of it, from the committed package. With it,
   * the host claims the `images` capability and `getImageAsBase64` works; without it, a picture
   * request refuses with `unsupported-capability`. This lane has no layout or paint, so the
   * caller supplies the drawing: `@pptx-editor.dev/editor-api` builds it from the editor's own
   * paint and a rasterizer (`createServer({ rasterize })`). Rejects with a `CommandRefusal` when
   * it cannot draw.
   */
  readonly renderImage?: (target: AutomationImageTarget) => Promise<Uint8Array>;
}

/**
 * A port over a store this host owns or shares. New part names never reuse a name the store has
 * held, so a name in the history always means one part.
 */
export function storeAutomationPort(
  store: PptxDocumentStore,
  options: StoreAutomationOptions = {}
): AutomationDocumentPort {
  const used = new Set<string>();
  for (const name of store.package.parts.keys()) used.add(partNameKey(name));
  const unsubscribe = store.subscribe((change) => {
    for (const name of change.parts) used.add(partNameKey(name));
  });
  const label = options.label ?? 'automation';
  return {
    revision: () => store.revision,
    currentPackage: () => store.package,
    batch: (pkg) => new OpBatch(pkg, { reservedPartNames: used }),
    apply(ops) {
      const before = store.revision;
      const r = store.transact(ops, { label });
      if (!r.ok) return { ok: false, reason: r.rejection.code };
      return { ok: true, changed: store.revision !== before };
    },
    save: () => writePptx(packageForSave(store.package)),
    subscribe: (listener) => store.subscribe(() => listener()),
    measurer: () => options.measurer,
    collaborating: () => options.collaborating === true,
    ...(options.renderImage ? { renderImage: options.renderImage } : {}),
    dispose: unsubscribe,
  };
}

/** A host over a store: the server host's composition, and the base of the editor's. */
export function createStoreAutomationHost(
  store: PptxDocumentStore,
  options: StoreAutomationOptions = {}
): AutomationHost {
  return createAutomationHost({
    port: storeAutomationPort(store, options),
    capabilities: {
      document: true,
      save: true,
      events: true,
      selection: false,
      images: options.renderImage !== undefined,
      autofit: options.measurer !== undefined,
    },
  });
}

/** Why bytes could not be opened. */
export type ServerAutomationRejection = PackageRejection | { readonly reason: 'empty' };

/** Options for a host over bytes. */
export interface ServerAutomationOptions extends StoreAutomationOptions {
  /** Zip limits for the bounded reader. */
  readonly limits?: ZipLimits;
}

/** Open `.pptx` bytes as a headless host. Never throws for a malformed file. */
export function createServerAutomationHost(
  bytes: Uint8Array,
  options: ServerAutomationOptions = {}
):
  | { readonly ok: true; readonly host: AutomationHost }
  | { readonly ok: false; readonly error: ServerAutomationRejection } {
  if (!(bytes instanceof Uint8Array) || bytes.length === 0)
    return { ok: false, error: { reason: 'empty' } };
  const read = options.limits ? readPptx(bytes, options.limits) : readPptx(bytes);
  if (!read.ok) return { ok: false, error: read.error };
  const store = openStore(read.pkg);
  return { ok: true, host: createStoreAutomationHost(store, options) };
}
