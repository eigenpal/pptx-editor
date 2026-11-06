import { describe, expect, test } from 'bun:test';
import { cpus } from 'node:os';
import { laneCount } from './lanes.ts';

describe('laneCount', () => {
  test('takes PPTX_FIDELITY_JOBS, never more lanes than decks', () => {
    expect(laneCount(38, '3')).toBe(3);
    expect(laneCount(2, '8')).toBe(2);
    expect(laneCount(0, '8')).toBe(1);
  });

  test('without a valid PPTX_FIDELITY_JOBS, one lane for each core, at most 4', () => {
    const want = Math.min(4, cpus().length);
    expect(laneCount(38, undefined)).toBe(want);
    expect(laneCount(38, '0')).toBe(want);
    expect(laneCount(38, 'many')).toBe(want);
  });
});
