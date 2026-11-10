import { expect, test } from 'bun:test';
import { repositoryPath, sourceFormat, validateDecision } from './policy.mjs';
const before = { format: 'pptx-collaboration:1.8.1.1', values: [1, 8, 1, 1] };
const after = { format: 'pptx-collaboration:1.8.2.1', values: [1, 8, 2, 1] };
const decision = {
  impact: 'migration-required',
  before: before.format,
  after: after.format,
  fields: ['repairVersion'],
  reason: 'New repair interpretation',
  tests: ['scripts/collaboration/test.mjs'],
};
test('compatibility decisions must describe the actual source transition', () => {
  expect(() => validateDecision(decision, before, after)).not.toThrow();
  for (const change of [
    { before: 'unrelated' },
    { after: before.format },
    { impact: 'compatible' },
    { fields: [] },
    { fields: ['protocolVersion'] },
    { fields: ['repairVersion', 'repairVersion'] },
  ])
    expect(() => validateDecision({ ...decision, ...change }, before, after)).toThrow();
  expect(() =>
    validateDecision({ ...decision, before: after.format, after: before.format }, after, before)
  ).toThrow();
});
test('unchanged formats need no changed fields or migration claim', () => {
  const compatible = { ...decision, impact: 'compatible', after: before.format, fields: [] };
  expect(() => validateDecision(compatible, before, before)).not.toThrow();
  expect(() =>
    validateDecision({ ...compatible, impact: 'migration-required' }, before, before)
  ).toThrow();
});
test('source format parsing fails on partial tuples', () => {
  expect(sourceFormat('legacy', 'SCHEMA_VERSION = 7').format).toBe('pptx-collaboration-7');
  expect(() =>
    sourceFormat('COLLABORATION_VERSIONS = { protocolVersion: 1 }', 'SCHEMA_VERSION = 8')
  ).toThrow();
});
test('evidence paths cannot leave the repository', () => {
  expect(repositoryPath('scripts/collaboration/test.mjs')).toBe('scripts/collaboration/test.mjs');
  for (const path of ['/etc/passwd', '../test.ts', 'docs/../../test.ts', 'scripts\\test.ts', ''])
    expect(() => repositoryPath(path)).toThrow();
});
