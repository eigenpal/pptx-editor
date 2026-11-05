import manifest from '../packages/core/licenses/harfbuzz-manifest.json';

export const HARFBUZZ_MANIFEST: typeof manifest;
export const HARFBUZZ_NOTICE: string;
export function validateHarfBuzzNotices(options?: {
  coreRoot?: string;
  dependencyRoot?: string;
}): void;
