// Build the demo's favicons from `examples/vite/public/icon.svg`, as DOCX Editor's demo ships
// them: `favicon.ico` (16, 32, and 48 px PNG images), `icon.png` (180 px, Apple touch icon),
// and `web-app-manifest-192x192.png` / `-512x512.png`. Chromium rasterizes the SVG.
// Usage: bun scripts/brand/build-icons.ts
import { chromium } from '@playwright/test';
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const dir = join(import.meta.dir, '../../examples/vite/public');
const svg = readFileSync(join(dir, 'icon.svg'), 'utf8');

const browser = await chromium.launch();
const page = await browser.newPage();
async function png(size: number): Promise<Uint8Array> {
  await page.setViewportSize({ width: size, height: size });
  await page.setContent(
    `<html><body style="margin:0;background:transparent">${svg.replace('<svg ', `<svg width="${size}" height="${size}" `)}</body></html>`
  );
  return new Uint8Array(await page.screenshot({ omitBackground: true, type: 'png' }));
}

/** An ICO file whose images are PNG data (supported by every current browser). */
function ico(images: readonly { size: number; data: Uint8Array }[]): Uint8Array {
  const header = 6 + 16 * images.length;
  const total = header + images.reduce((n, i) => n + i.data.length, 0);
  const out = new Uint8Array(total);
  const view = new DataView(out.buffer);
  view.setUint16(2, 1, true);
  view.setUint16(4, images.length, true);
  let offset = header;
  images.forEach((image, k) => {
    const e = 6 + 16 * k;
    out[e] = image.size >= 256 ? 0 : image.size;
    out[e + 1] = image.size >= 256 ? 0 : image.size;
    view.setUint16(e + 4, 1, true);
    view.setUint16(e + 6, 32, true);
    view.setUint32(e + 8, image.data.length, true);
    view.setUint32(e + 12, offset, true);
    out.set(image.data, offset);
    offset += image.data.length;
  });
  return out;
}

const small = [];
for (const size of [16, 32, 48]) small.push({ size, data: await png(size) });
writeFileSync(join(dir, 'favicon.ico'), ico(small));
writeFileSync(join(dir, 'icon.png'), await png(180));
writeFileSync(join(dir, 'web-app-manifest-192x192.png'), await png(192));
writeFileSync(join(dir, 'web-app-manifest-512x512.png'), await png(512));
await browser.close();
console.log(`Wrote favicons to ${dir}`);
