// Tests for the resource limits (`limits.ts`, copied from DOCX Editor `store/runtime/limits.ts`
// at 26095c63): a host may lower a limit but never raise or disable one.

import { describe, expect, test } from 'bun:test';
import {
  DEFAULT_IMAGE_RESOURCE_LIMITS,
  DEFAULT_LIMITS,
  HARD_CEILINGS,
  IMAGE_RESOURCE_HARD_CEILINGS,
  resolveImageResourceLimits,
  resolveLimits,
} from '../limits.ts';

describe('resolveLimits', () => {
  test('the defaults are finite and at or below their ceilings', () => {
    for (const key of Object.keys(HARD_CEILINGS) as (keyof typeof HARD_CEILINGS)[]) {
      expect(Number.isFinite(DEFAULT_LIMITS[key])).toBe(true);
      expect(DEFAULT_LIMITS[key]).toBeLessThanOrEqual(HARD_CEILINGS[key]);
    }
  });

  test('Infinity, zero, a negative, and NaN never turn a limit off', () => {
    for (const bad of [Infinity, 0, -1, Number.NaN]) {
      const resolved = resolveLimits({ maxElementCount: bad, maxRecursionDepth: bad });
      expect(resolved.maxElementCount).toBe(DEFAULT_LIMITS.maxElementCount);
      expect(resolved.maxRecursionDepth).toBe(DEFAULT_LIMITS.maxRecursionDepth);
    }
  });

  test('a host may lower a limit but not raise it past the ceiling', () => {
    expect(resolveLimits({ maxElementCount: 10 }).maxElementCount).toBe(10);
    expect(resolveLimits({ maxElementCount: 1e12 }).maxElementCount).toBe(
      HARD_CEILINGS.maxElementCount
    );
    expect(resolveLimits({ maxRecursionDepth: 2.9 }).maxRecursionDepth).toBe(2);
  });

  test('every limit set is frozen', () => {
    expect(Object.isFrozen(HARD_CEILINGS)).toBe(true);
    expect(Object.isFrozen(DEFAULT_LIMITS)).toBe(true);
    expect(Object.isFrozen(resolveLimits())).toBe(true);
  });
});

describe('resolveImageResourceLimits', () => {
  test('clamps into the image ceilings and ignores bad values', () => {
    const resolved = resolveImageResourceLimits({ maxPixels: 1e15, maxDimension: Number.NaN });
    expect(resolved.maxPixels).toBe(IMAGE_RESOURCE_HARD_CEILINGS.maxPixels);
    expect(resolved.maxDimension).toBe(DEFAULT_IMAGE_RESOURCE_LIMITS.maxDimension);
    expect(Object.isFrozen(resolved)).toBe(true);
  });
});
