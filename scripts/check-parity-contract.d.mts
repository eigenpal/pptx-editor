// Types for check-parity-contract.mjs, for the TypeScript tests.
import type { VueAdapterShape } from '../packages/vue/test/parity-shape.ts';

export function parityErrors(input: {
  reactApi: string;
  vueApi: string;
  vue: VueAdapterShape;
  // The parsed scripts/parity/parity.contract.json.
  contract: any;
}): string[];
