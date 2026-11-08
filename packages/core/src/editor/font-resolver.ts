// Copied from DOCX Editor `packages/core/src/editor/font-resolver.ts` at commit adf0e178.
// Keep the exported API equal to the source.

// Moved to the layout lane so the headless export composition root can call it without a
// grandfathered lane edge; this shim keeps the editor-internal import paths stable.
export * from '../layout/font-resolver.ts';
