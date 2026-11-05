// Types for check-adapter-mirror.mjs, for the TypeScript tests.
export const VUE_ONLY: Readonly<Record<string, string>>;
export const REACT_ONLY: Readonly<Record<string, string>>;

export function mirrorErrors(input: {
  reactFiles: readonly string[];
  vueFiles: readonly string[];
  vueOnly?: Readonly<Record<string, string>>;
  reactOnly?: Readonly<Record<string, string>>;
}): string[];
