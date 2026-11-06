// The rail's cache of item vnodes: core's `createItemMemo`, with each vnode marked as `v-memo`
// marks the vnodes it caches. Vue copies a mounted vnode that it is given again, unless the
// vnode carries `memo`; with the mark, it finds the same vnode and skips the item at once.
//
// `vnode.memo` is not public API. This relies on Vue 3.5.44 (the version the workspace tests),
// where `cloneIfMounted` keeps a mounted vnode that has `memo`, and `patch` returns at once
// when the old and the new vnode are the same object. The field is there since `v-memo` came
// (Vue 3.2). `test/rail-keystroke.test.ts` fails if Vue patches the items again.

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
