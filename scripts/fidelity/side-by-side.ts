// Join two screenshots side by side (left | right) for family-parity review.
//   bun scripts/fidelity/side-by-side.ts <left.png> <right.png> <out.png>
import { readFileSync, writeFileSync } from 'node:fs';
import { PNG } from 'pngjs';

const [l, r, out] = process.argv.slice(2);
const a = PNG.sync.read(readFileSync(l));
const b = PNG.sync.read(readFileSync(r));
const gap = 16;
const img = new PNG({ width: a.width + gap + b.width, height: Math.max(a.height, b.height) });
img.data.fill(255);
const blit = (src: PNG, ox: number): void => {
  for (let y = 0; y < src.height; y += 1)
    src.data.copy(img.data, (y * img.width + ox) * 4, y * src.width * 4, (y + 1) * src.width * 4);
};
blit(a, 0);
blit(b, a.width + gap);
writeFileSync(out, PNG.sync.write(img));
