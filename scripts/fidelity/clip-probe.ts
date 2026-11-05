// Dev probe: does Chromium clip a
// shape in EMU coordinates correctly under a translation, with a clipPath and with a mask?
//   bun scripts/fidelity/clip-probe.ts
// Each case renders a gradient rectangle with no clip (the truth), then with each clip kind, and
// counts the cases where the red slide background shows in a different place.
import { chromium } from '@playwright/test';
import { PNG } from 'pngjs';

const W = 12_137_843;
const H = 6_858_000;
type Kind = 'none' | 'clipPath' | 'mask';

function defsFor(kind: Kind): string {
  if (kind === 'clipPath')
    return `<clipPath id="c"><path d="M0 0L${W} 0L${W} ${H}L0 ${H}Z"/></clipPath>`;
  if (kind === 'mask')
    return `<mask id="c" maskUnits="userSpaceOnUse" x="0" y="0" width="${W}" height="${H}"><path fill="white" d="M0 0L${W} 0L${W} ${H}L0 ${H}Z"/></mask>`;
  return '';
}

function svgFor(tx: number, ty: number, rot: number, kind: Kind): string {
  const use = kind === 'clipPath' ? 'clip-path="url(#c)"' : kind === 'mask' ? 'mask="url(#c)"' : '';
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="1440" viewBox="0 0 9144000 6858000"><defs>` +
    `<linearGradient id="g" gradientUnits="userSpaceOnUse" x1="0" x2="${W}" y2="${H}"><stop offset="0" stop-color="blue"/><stop offset="1" stop-color="lime"/></linearGradient>${defsFor(kind)}</defs>` +
    `<rect width="9144000" height="6858000" fill="red"/><g transform="translate(${tx} ${ty}) rotate(${rot} ${W / 2} ${H / 2})">` +
    `<g ${use}><rect width="${W}" height="${H}" fill="url(#g)"/></g></g></svg>`
  );
}

/** Pixels that are slide red in one image and not in the other. */
function redDifference(a: Buffer, b: Buffer): number {
  const x = PNG.sync.read(a);
  const y = PNG.sync.read(b);
  const red = (d: Buffer, i: number): boolean => d[i] > 200 && d[i + 1] < 60 && d[i + 2] < 60;
  let n = 0;
  for (let i = 0; i < x.data.length; i += 4) if (red(x.data, i) !== red(y.data, i)) n += 1;
  return n;
}

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1920, height: 1440 } });
const shot = async (markup: string): Promise<Buffer> => {
  await page.setContent(`<body style="margin:0">${markup}</body>`);
  return page.screenshot();
};
const wrong: Record<Kind, number> = { none: 0, clipPath: 0, mask: 0 };
let cases = 0;
for (const tx of [-6e6, -4.2e6, -2_340_768, -1e6, -123_457, 0, 777_777, 3e6])
  for (const ty of [-3e6, -500_000, 0, 1e6])
    for (const rot of [0, 7, 33, 90]) {
      cases += 1;
      const truth = await shot(svgFor(tx, ty, rot, 'none'));
      for (const kind of ['clipPath', 'mask'] as const)
        if (redDifference(truth, await shot(svgFor(tx, ty, rot, kind))) > 20_000) wrong[kind] += 1;
    }
await browser.close();
console.log(
  `${cases} placements: clipPath wrong in ${wrong.clipPath}, mask wrong in ${wrong.mask}.`
);
