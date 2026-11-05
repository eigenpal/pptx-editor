// Copied from DOCX Editor `packages/vue/src/editor/useFonts.ts` at commit 5993bd223.
// Keep the exported API equal to the source.

import { toValue } from 'vue';
import {
  composeFontOrigins,
  defineFontResolver,
  type FontConfigurationFragment,
  type FontOrigin,
  type FontResolutionRequest,
  type FontResolver,
  type MarkedFontResolver,
} from '@pptx-editor.dev/core/editor';
import type { FontConfiguration } from '@pptx-editor.dev/core/contracts';
import type { MaybeRefOrGetter } from 'vue';

/**
 * What `useFonts` returns: one marked resolver for the component's life. Named for the
 * adapters' composable parity check (PPTX Editor addition).
 *
 * @public
 */
export type UseFontsReturn = MarkedFontResolver;

/**
 * The first argument of `useFonts`: a `FontConfiguration`, a fragment, a `FontResolver` (such
 * as `packagedFonts()`), a promise of a configuration or a fragment, or `undefined`. Each may
 * be a `ref` or a getter, which `useFonts` reads again at each resolve.
 *
 * @public
 */
export type FontsInput =
  | FontConfiguration
  | FontConfigurationFragment
  | FontResolver
  | Promise<FontConfiguration | FontConfigurationFragment | undefined>
  | undefined;

/**
 * Read one origin, resolving a `ref` but NEVER calling a function.
 *
 * `toValue` invokes a plain function to read it as a getter, and a resolver invoked with
 * no argument reads `request.defaultFamily` off `undefined` and throws — which is how
 * `useFonts(googleFonts())` came to leave the document on fixed measurement with no
 * diagnostic at all.
 *
 * Arity looked like a way to tell the two apart and is not one. A getter never declares a
 * parameter, so "declares one ⇒ not a getter" is sound — but the branch it falls THROUGH
 * to is the mirror-image unsound move, reading a zero-arity function as a getter and
 * calling it. `(request = FALLBACK) => …` and `(...args) => …` are both length 0 and both
 * resolvers, and this file's own contract names that trap by name.
 *
 * So: every function is a value here, exactly as in the React twin, and the two adapters
 * have no behavioural difference left to diverge on. A getter still works, because
 * `composeFontOrigins` calls every function origin once per resolve and a getter ignores
 * the argument it is handed — the re-read the getter form exists for is unaffected. The
 * one shape that changes is a getter returning another ORIGIN (`() => packagedFonts()`),
 * which cannot be told from a resolver by anything; `composeFontOrigins` reports it rather
 * than composing an empty fragment in silence. Wrap that in a `computed` if you need it.
 */
const readOrigin = (origin: MaybeRefOrGetter<FontsInput | FontOrigin>): FontOrigin =>
  (typeof origin === 'function' ? origin : toValue(origin)) as FontOrigin;

/**
 * Merge font origins into one stable value for the `fonts` prop of `PptxEditor` or
 * `PptxEditorRoot`, as React's `useFonts`: `useFonts(packagedFonts())`, or
 * `useFonts(brandFragment, packagedFonts())` for brand faces first. Origins compose
 * first-wins in argument order and resolve one after another. The value never changes
 * identity, so the fonts are not replaced on a re-render; `ref` and getter arguments are read
 * again when a deck opens or a family is requested.
 *
 * @public
 */
export function useFonts(
  source: MaybeRefOrGetter<FontsInput>,
  ...fragments: readonly MaybeRefOrGetter<FontConfigurationFragment | undefined>[]
): UseFontsReturn;
/**
 * The uniform form: every position takes the same `FontOrigin`, so composing two resolvers is
 * one extra argument. A resolver after the first position must carry the
 * `defineFontResolver` mark.
 *
 * @public
 */
export function useFonts(...origins: readonly MaybeRefOrGetter<FontOrigin>[]): UseFontsReturn;
export function useFonts(
  ...origins: readonly MaybeRefOrGetter<FontsInput | FontOrigin>[]
): UseFontsReturn {
  const reactiveOrigins = origins;
  const readOrigins = (): readonly FontOrigin[] => reactiveOrigins.map(readOrigin);

  // Marked, so this composition can itself be an origin of another list or
  // the `fonts` option of `createPptxEditor` without being read as a zero-argument loader.
  return defineFontResolver((request: FontResolutionRequest) =>
    composeFontOrigins(readOrigins(), request)
  );
}
