export interface PublicPackage {
  name: string;
  version: string;
}
export type DistTags = Record<string, string>;
export type DistTagLookup = (name: string) => Promise<DistTags | null>;
export interface ModeAnswer {
  mode: string;
  reason: string;
}
export interface GateAnswer {
  clear: boolean;
  reason: string;
}

export declare function pendingChangesets(root: string): string[];
export declare function publicPackages(root: string): PublicPackage[];
export declare function releaseMode(root: string): { mode: 'pending' | 'publish' };
export declare const CANARY_TAG: 'canary';
export declare function canarySuffix(commit: string): string;
export declare function pendingReleases(root: string): string[];
export declare function distTags(
  name: string,
  options?: { fetch?: typeof fetch }
): Promise<DistTags | null>;
export declare function canaryMode(
  root: string,
  commit: string,
  tags?: DistTagLookup
): Promise<ModeAnswer>;
export declare function stableReleaseClear(options: {
  repository: string;
  runId: string;
  token: string;
  fetch?: typeof fetch;
}): Promise<GateAnswer>;
