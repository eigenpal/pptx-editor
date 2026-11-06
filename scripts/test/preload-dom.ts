// Registers happy-dom once, before any adapter test module is evaluated (DOCX Editor's
// scripts/test/preload-dom.ts).
//
// Only the adapter tests preload it: `bun run test` passes `--preload` for the files under
// `packages/react/` and `packages/vue/` (`pptx-suite.mjs`), and `bun run test:dom` runs them
// alone. The engine tests stay without a DOM, so they take the same code paths as a server or
// a worker, and a DOM-only dependency in a DOM-free lane fails there.
//
// `bun run test:dom` runs every adapter file in one process, so all files share the DOM
// classes. A test that patches a DOM prototype must leave it as it found it. The guards below
// fail a test that leaves a copy of an inherited property on a prototype, and fail the run
// when a prototype differs at the end (`dom-prototypes.ts`).
import { afterAll, afterEach } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { prototypeChanges, shadowingCopies, snapshotPrototypes } from './dom-prototypes.ts';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const baseline = snapshotPrototypes(globalThis);
const RESTORE =
  'Restore the saved property descriptor, or delete an own property that was not there.';

// A preloaded `afterEach` runs after the test file's own `afterEach` hooks.
afterEach(() => {
  const copies = shadowingCopies(globalThis, baseline);
  if (copies.length)
    throw new Error(
      `${Bun.main} left a copy of an inherited DOM property: ${copies.join(', ')}. ${RESTORE}`
    );
});

// A preloaded `afterAll` runs once, after the last file.
afterAll(() => {
  const changes = prototypeChanges(globalThis, baseline);
  if (changes.length)
    throw new Error(
      `The adapter tests left DOM prototype changes: ${changes.join(', ')}. ${RESTORE}`
    );
});
