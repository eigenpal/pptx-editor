// Compiles each DOM-free lane against its own tsconfig (lib without DOM, no ambient types), so a
// `document`, `window`, or DOM type such as `HTMLElement` in store, cascade, layout, automation,
// or ooxml-core fails, also when it arrives through an import. Same approach as
// docx-editor/scripts/check-lane-boundaries.mjs, with `bun check` in place of `tsc`.
//
// `bun check` reports the same errors as `tsc`. It also loads only what the lane imports: `tsc`
// once pulled `@types/react` into the automation lane, whose empty global `HTMLElement`
// declaration hid a DOM type in the contracts that the lane imported.
import { spawnSync } from 'node:child_process';

// The shared OOXML package layer is DOM-free too: it sits below the store lane.
const NEUTRAL = {
  'ooxml-core': 'packages/ooxml-core/src/tsconfig.json',
  store: 'packages/core/src/store/tsconfig.json',
  cascade: 'packages/core/src/cascade/tsconfig.json',
  layout: 'packages/core/src/layout/tsconfig.json',
  automation: 'packages/core/src/automation/tsconfig.json',
};
let failed = false;
for (const [lane, project] of Object.entries(NEUTRAL)) {
  const result = spawnSync('bun', ['check', '--no-pretty', '-p', project], { stdio: 'inherit' });
  if (result.status !== 0) {
    console.error(`Lane ${lane} is not DOM-free.`);
    failed = true;
  }
}
process.exit(failed ? 1 : 0);
