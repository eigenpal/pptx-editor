// PPTX Editor only: the facts about this repository that the runner copied from DOCX Editor
// (`run-parallel.mjs`, `changed.mjs`) needs. Kept apart, so the copied files stay equal to
// their source apart from the lines that read these values.

/**
 * The npm scopes of the workspace packages. Each maps `<scope><name>` to `packages/<name>`.
 * `@eigenpal/ooxml-core` lives in `packages/ooxml-core`, so a change there selects the core
 * tests that import it.
 */
export const WORKSPACE_SCOPES = ['@pptx-editor.dev/', '@eigenpal/'];

/** The adapter tests: the only files that run with a global DOM. */
export const DOM_TEST_ROOTS = ['packages/react/', 'packages/vue/'];

/** Registers happy-dom before an adapter test module is evaluated. */
export const DOM_PRELOAD = './scripts/test/preload-dom.ts';

/**
 * The extra `bun test` arguments for one file.
 *
 * DOCX Editor preloads happy-dom for the whole suite from `bunfig.toml`. PPTX Editor does
 * not: the engine tests (`store`, `cascade`, `layout`, and the rest of `core`, `pro`,
 * `pptx-to-pdf`, `editor-api`) run with no global `document`, so they take the same code
 * paths as a server or a worker, and a DOM-only dependency in a DOM-free lane fails there.
 * A test that needs a DOM makes its own happy-dom window. Only the adapter tests preload it.
 */
export function bunArgumentsFor(file) {
  return DOM_TEST_ROOTS.some((root) => file.startsWith(root)) ? ['--preload', DOM_PRELOAD] : [];
}
