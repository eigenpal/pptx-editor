// The blur rule of the SVG paint (ECMA-376 Part 1, 20.1.8: CT_OuterShadowEffect). The constant is
// measured on the `effects` and `shape-effects-presets` reference PDFs. The glow, soft edge,
// inner shadow, and reflection rules are tested with the Pro drawing renderers
// (`packages/pro/src/drawing/__tests__/effects-kernels.test.ts`).

import { describe, expect, test } from 'bun:test';
import { BLUR_SIGMA_PER_RADIUS, blurSigma } from '../effect-transforms.ts';

describe('blur kernel', () => {
  test('a blur radius is three standard deviations', () => {
    expect(BLUR_SIGMA_PER_RADIUS).toBeCloseTo(1 / 3, 12);
    // Offset: center (shape-effects-presets slide 1): blurRad 63500 is 10 px, σ about 3.3 px.
    expect(blurSigma(63_500)).toBeCloseTo(21_166.67, 1);
    expect(blurSigma(-5)).toBe(0);
  });
});
