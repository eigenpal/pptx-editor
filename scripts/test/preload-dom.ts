// Registers happy-dom once, before any adapter test module is evaluated (DOCX Editor's
// scripts/test/preload-dom.ts).
//
// Only the adapter tests preload it: `bun run test` passes `--preload` for the files under
// `packages/react/` and `packages/vue/` (`pptx-suite.mjs`), and `bun run test:dom` runs them
// alone. The engine tests stay without a DOM, so they take the same code paths as a server or
// a worker, and a DOM-only dependency in a DOM-free lane fails there.
import { GlobalRegistrator } from '@happy-dom/global-registrator';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
