/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The seeded generator and the shrink of the collaboration fuzzers, adapted from DOCX Editor
// `packages/pro/src/collaboration/__tests__/document-undo-groups-fuzz.test.ts` at commit
// 5993bd223. This is DOCX Editor pro code, adapted into this pro package under the same licensor
// and license. `shrink` takes `replay` as an argument, and `kindOf` also drops the seed prefix
// and keeps 60 characters. `seededRandom` is also the `mulberry32` of DOCX Editor's Apache-2.0
// `packages/core/src/store/__tests__/random-edit-test-support.ts` at 5993bd223.

/** A small deterministic generator (mulberry32), so a failure names its seed. */
export function seededRandom(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let value = state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4_294_967_296;
  };
}

/** The kind of a problem, without its trail and values, to keep a shrink on one defect. */
export function kindOf(problem: string): string {
  return problem
    .replace(/^seed \d+: /, '')
    .replace(/\n[\s\S]*$/, '')
    .replace(/:.*$/, '')
    .replace(/\d+/g, '#')
    .slice(0, 60);
}

/** Remove actions while the same kind of problem persists (delta debugging, as DOCX Editor). */
export async function shrink<A>(
  actions: readonly A[],
  kind: string,
  replay: (actions: readonly A[]) => Promise<string[]>
): Promise<A[]> {
  let current = [...actions];
  for (let size = Math.ceil(current.length / 2); size >= 1; size = Math.floor(size / 2)) {
    let removed = true;
    while (removed) {
      removed = false;
      for (let start = 0; start < current.length; start += size) {
        const candidate = [...current.slice(0, start), ...current.slice(start + size)];
        const problems = await replay(candidate);
        if (problems.some((problem) => kindOf(problem) === kind)) {
          current = candidate;
          removed = true;
          break;
        }
      }
    }
  }
  return current;
}
