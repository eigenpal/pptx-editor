// Copied from DOCX Editor packages/editor-api/scripts/lib/typecheck-compat.d.mts at commit 5993bd223945cabdaf3900557d9fb59bfd8c2d8c.
// Changes: this comment names the PPTX Editor test that imports the helper. The declarations
// are equal.
// Declarations for the `.mjs` helper that `scripts/__tests__/conformance.test.ts` imports.
//
// The compat build steps are plain Node modules on purpose: they run before any TypeScript
// exists. A TypeScript caller that a typecheck covers needs these declarations, or the import
// is an implicit `any`.

/**
 * Compile a `tsconfig.json` in-process and answer its diagnostics, one formatted string each.
 * An empty array means the project compiled clean.
 *
 * `scope: 'program'` (the default) reports every file the program reaches. `scope: 'project'`
 * reports the files the tsconfig includes (syntactic, semantic, and declaration diagnostics,
 * plus option and global ones); the bodies of other files are not checked. `forbidHost` reports,
 * in every file the program reaches whose path has no `/node_modules/`, Node or Bun module
 * imports and type references, host `import.meta` members, and `HOST_GLOBALS` names that resolve
 * to no declaration, only to Node or Bun type packages, or only to ambient declarations.
 */
export function typecheckProject(
  tsconfigPath: string,
  options?: { readonly scope?: 'program' | 'project'; readonly forbidHost?: boolean }
): string[];

/** Globals only a server runtime declares; see `typecheckProject`'s `forbidHost`. */
export const HOST_GLOBALS: readonly string[];

/** `import.meta` members only Node or Bun provide. */
export const HOST_IMPORT_META: readonly string[];

/**
 * The per-test timeout a `typecheckProject` call needs. A compile of these projects is seconds,
 * not milliseconds, and `bun test`'s default budget is five seconds.
 */
export const TYPECHECK_TIMEOUT_MS: number;
