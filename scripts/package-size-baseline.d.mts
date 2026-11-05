export interface BaselineTarget {
  repository?: string;
  branch?: string;
  sha?: string;
  directory?: string;
}
export interface WorkflowRun {
  id: number;
  head_sha?: string;
  head_branch?: string;
  event?: string;
  path?: string;
  conclusion?: string;
  status?: string;
  head_repository?: { full_name?: string };
}
export declare function selectBaselineRun(
  runs: WorkflowRun[],
  target: Required<Pick<BaselineTarget, 'repository' | 'branch' | 'sha'>>
): WorkflowRun | undefined;
export declare function downloadBaseline(
  target: BaselineTarget,
  run?: (args: string[]) => string
): string | undefined;
