// A deterministic, photo-like landscape for the showcase deck: a dusk sky, a low sun, three
// mountain ridges, a lake with the reflection, and a dark shore with pines. Generated here, so
// the deck holds no third-party image.

import { PNG } from 'pngjs';

type Rgb = readonly [number, number, number];

const mix = (a: Rgb, b: Rgb, t: number): Rgb => [
  a[0] + (b[0] - a[0]) * t,
  a[1] + (b[1] - a[1]) * t,
  a[2] + (b[2] - a[2]) * t,
];

/** A small seeded generator (mulberry32), so every build writes the same bytes. */
function random(seed: number): () => number {
  let s = seed >>> 0;
  return () => {
    s = (s + 0x6d2b79f5) >>> 0;
    let t = s;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** A ridge line: the sum of a few sines with seeded phases, in pixels above `base`. */
function ridge(width: number, base: number, amp: number, seed: number): Float64Array {
  const rnd = random(seed);
  const waves = [1, 2.3, 4.7, 9.1, 17.3].map((f, i) => ({
    f,
    phase: rnd() * Math.PI * 2,
    a: amp / (i + 1) ** 1.15,
  }));
  const out = new Float64Array(width);
  for (let x = 0; x < width; x += 1) {
    let h = 0;
    for (const w of waves) h += w.a * Math.sin((x / width) * Math.PI * w.f + w.phase);
    out[x] = base - h;
  }
  return out;
}

/** The landscape as PNG bytes, `width` × `height` pixels. */
export function landscapePng(width = 1000, height = 625): Uint8Array {
  const png = new PNG({ width, height });
  const horizon = Math.round(height * 0.64);
  const sun = { x: width * 0.66, y: horizon - height * 0.09, r: height * 0.075 };
  const top: Rgb = [38, 44, 92];
  const middle: Rgb = [176, 92, 96];
  const low: Rgb = [246, 190, 132];
  const ridges = [
    { line: ridge(width, horizon - height * 0.1, height * 0.07, 7), color: [150, 96, 122] as Rgb },
    { line: ridge(width, horizon - height * 0.03, height * 0.06, 19), color: [96, 58, 88] as Rgb },
    { line: ridge(width, horizon + 2, height * 0.035, 31), color: [52, 34, 54] as Rgb },
  ];
  const sky = (x: number, y: number): Rgb => {
    const t = Math.min(1, Math.max(0, y / horizon));
    let c = t < 0.62 ? mix(top, middle, t / 0.62) : mix(middle, low, (t - 0.62) / 0.38);
    const d = Math.hypot(x - sun.x, y - sun.y);
    if (d < sun.r) return [255, 238, 204];
    const glow = Math.exp(-(d - sun.r) / (height * 0.12)) * 0.65;
    c = mix(c, [255, 222, 170], glow);
    return c;
  };
  const above = (x: number, y: number): Rgb => {
    let c = sky(x, y);
    for (const [i, r] of ridges.entries()) {
      if (y < r.line[x]) continue;
      // Haze: the far ridges take some of the sky's color near their foot.
      const haze = (1 - i / ridges.length) * 0.35 * Math.min(1, (y - r.line[x]) / (height * 0.08));
      c = mix(r.color, low, haze * 0.6);
    }
    return c;
  };
  const shore = ridge(width, height * 0.93, height * 0.025, 43);
  const rnd = random(101);
  const pines = Array.from({ length: 6 }, (_, i) => ({
    x: width * (0.035 + i * 0.05 + rnd() * 0.018),
    h: height * (0.17 + rnd() * 0.1),
    w: width * (0.05 + rnd() * 0.015),
  }));
  const inPine = (x: number, y: number): boolean =>
    pines.some((p) => {
      const baseY = shore[Math.min(width - 1, Math.max(0, Math.round(p.x)))];
      const tipY = baseY - p.h;
      if (y < tipY || y > baseY) return false;
      const t = (y - tipY) / p.h;
      // A short trunk under three overlapping tiers, each a triangle widening downward.
      if (t > 0.88) return Math.abs(x - p.x) < p.w * 0.07;
      const tier = Math.min(1, ((t / 0.88) * 3) % 1.0001);
      const half = (p.w / 2) * (0.3 + 0.7 * (t / 0.88)) * (0.35 + 0.65 * tier);
      return Math.abs(x - p.x) <= half;
    });
  for (let y = 0; y < height; y += 1)
    for (let x = 0; x < width; x += 1) {
      let c: Rgb;
      if (y < horizon) c = above(x, y);
      else {
        // The lake mirrors the scene above the horizon, darker, with soft ripples.
        const depth = (y - horizon) / (height - horizon);
        const ripple = Math.sin(y * 0.45 + Math.sin(x * 0.012) * 1.5) * (0.4 + depth * 2.5);
        const my = Math.max(0, 2 * horizon - y);
        const mx = Math.min(width - 1, Math.max(0, Math.round(x + ripple)));
        c = mix(above(mx, my), [30, 30, 58], 0.28 + depth * 0.3);
        // The sun's path on the water: a soft column that shimmers with the ripples.
        const dx = (x - sun.x) / (sun.r * (1.1 - depth * 0.4));
        const glint = Math.exp(-dx * dx) * (0.55 + 0.45 * Math.sin(y * 0.7 + x * 0.05));
        c = mix(c, low, glint * 0.4 * (1 - depth * 0.6));
      }
      if (y >= shore[x] || inPine(x, y)) c = [24, 22, 30];
      const i = (y * width + x) * 4;
      png.data[i] = Math.round(c[0]);
      png.data[i + 1] = Math.round(c[1]);
      png.data[i + 2] = Math.round(c[2]);
      png.data[i + 3] = 255;
    }
  return new Uint8Array(PNG.sync.write(png));
}
