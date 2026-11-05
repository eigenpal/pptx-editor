// Types for check-composable-parity.mjs, for the TypeScript tests.
export function composableErrors(input: {
  reactApi: string;
  vueApi: string;
  // The parsed scripts/parity/parity.contract.json.
  contract: any;
}): string[];
