// The rail's cache of item vnodes: core's `createItemMemo`, with each vnode marked as `v-memo`
// marks the vnodes it caches. Vue copies a mounted vnode that it is given again, unless the
// vnode carries `memo`; with the mark, it finds the same vnode and skips the item at once.

import type { VNode } from 'vue';
import { createItemMemo } from '@pptx-editor.dev/core/editor';

/** A cache of rail item vnodes, keyed on the slide summary. @internal */
export function createRailItemMemo(): (
  key: object,
  inputs: readonly unknown[],
  make: () => VNode
) => VNode {
  const memo = createItemMemo<VNode>();
  return (key, inputs, make) =>
    memo(key, inputs, () => {
      const vnode = make();
      (vnode as VNode & { memo?: unknown[] }).memo = [];
      return vnode;
    });
}
