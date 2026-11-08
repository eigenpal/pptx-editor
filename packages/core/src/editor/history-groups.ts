// The history groups of one editor: the handles a host opens with `beginHistoryGroup()`, and the
// gestures the editor groups itself (a typing burst, a drag, repeated nudges). Each one stands for
// a store token (`store/store/history-group.ts`); the store makes the commands that carry one token
// a single undo step. The handle follows DOCX Editor's
// `packages/core/src/contracts/editor-scope.ts` and `editor/editor-history-group-registry.ts`. DOCX
// Editor commit adf0e178.

import type { HistoryGroup } from '../contracts/index.ts';
import type { PptxPackage, HistoryGroup as StoreHistoryGroup } from '../store/index.ts';

/** One pointer drag: its store token, and the package when it started. */
interface Drag {
  readonly surface: unknown;
  readonly gesture: string;
  readonly token: StoreHistoryGroup;
  readonly base: PptxPackage | null;
}

const NO_DRAG: Drag = { surface: null, gesture: '', token: Symbol('drag'), base: null };

/** The handles and gestures of one editor, and the store token of each. */
export class EditorHistoryGroups {
  // Handles this editor minted, by handle. A handle from another editor is not here.
  readonly #tokens = new WeakMap<HistoryGroup, StoreHistoryGroup>();
  readonly #closers = new Set<() => void>();
  #typing: { readonly nodeId: string; readonly token: StoreHistoryGroup } | null = null;
  #nudge: StoreHistoryGroup | null = null;
  #drag: Drag = NO_DRAG;

  /** Open a host handle. It creates no undo entry. */
  begin(): HistoryGroup {
    const token = Symbol('history-group');
    let state: 'open' | 'closed' = 'open';
    const close = (): void => {
      state = 'closed';
      this.#closers.delete(close);
    };
    const handle = Object.freeze({
      get state() {
        return state;
      },
      end: close,
    }) as unknown as HistoryGroup;
    this.#tokens.set(handle, token);
    this.#closers.add(close);
    return handle;
  }

  /**
   * The store token of a host handle: `undefined` for no handle, `null` for a handle that is
   * closed or that another editor minted.
   */
  token(handle: HistoryGroup | undefined): StoreHistoryGroup | null | undefined {
    if (handle === undefined) return undefined;
    const token = this.#tokens.get(handle);
    return token !== undefined && handle.state === 'open' ? token : null;
  }

  /** Close every host handle: a new document loaded, or the editor was destroyed. */
  closeAll(): void {
    for (const close of [...this.#closers]) close();
    this.#drag = NO_DRAG;
    this.boundary();
  }

  /** The token of the typing burst in `nodeId`; a burst in another text body starts anew. */
  typing(nodeId: string): StoreHistoryGroup {
    if (this.#typing?.nodeId !== nodeId) this.#typing = { nodeId, token: Symbol('typing') };
    return this.#typing.token;
  }

  /** The token of consecutive nudges of one selection. */
  nudge(): StoreHistoryGroup {
    this.#nudge ??= Symbol('nudge');
    return this.#nudge;
  }

  /**
   * One pointer drag of a surface, by its gesture id: a new drag gets a new token and keeps
   * `pkg` as its base (a table resize scales from it). A new surface (an attach) numbers its
   * gestures from 1 again, so the surface is part of the key.
   */
  drag(surface: unknown, gesture: string, pkg: PptxPackage): Drag {
    const d = this.#drag;
    if (d.surface !== surface || d.gesture !== gesture)
      this.#drag = { surface, gesture, token: Symbol('drag'), base: pkg };
    return this.#drag;
  }

  /**
   * End the typing burst and the nudge series: the selection moved, or another command ran.
   * The store closes its open group on its own when a command with another token commits.
   */
  boundary(nudge = true): void {
    this.#typing = null;
    if (nudge) this.#nudge = null;
  }
}
