// The checks that need a real browser, on the React demo. Every other demo flow runs in
// happy-dom as a unit test (`packages/react/src/__tests__/demo-*.test.tsx`,
// `packages/vue/test/demo-shell.test.ts`), which has no layout engine and no canvas. These
// keep only what those cannot prove: pixels drawn on a canvas, and boxes from CSS layout.
import { readFileSync } from 'node:fs';
import { unzipSync } from 'fflate';
import { expect, test, type Page } from '@playwright/test';

/** Open the demo on the basic deck and wait until the slide is painted. */
async function ready(page: Page): Promise<void> {
  await page.goto('/?deck=basic');
  await expect(page.locator('.pptx-rail__item')).toHaveCount(5, { timeout: 30_000 });
  await page.waitForSelector('.pptx-surface__slide');
}

const PNG = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

test('slide pictures: File > Current slide as PNG and All slides as PNG draw the slides', async ({
  page,
}, info) => {
  // The editor's paint drawn on a real canvas: happy-dom has none.
  await ready(page);
  await page.getByRole('menuitem', { name: 'File', exact: true }).click();
  const [one] = await Promise.all([
    page.waitForEvent('download'),
    page.getByRole('menuitem', { name: 'Current slide as PNG' }).click(),
  ]);
  expect(one.suggestedFilename()).toMatch(/\.png$/);
  const file = info.outputPath('slide.png');
  await one.saveAs(file);
  const png = readFileSync(file);
  expect([...png.subarray(0, 8)]).toEqual(PNG);
  expect(png.readUInt32BE(16)).toBe(1920);
  expect(png.readUInt32BE(20)).toBe(1080);
  // A painted slide, not an empty canvas.
  expect(png.length).toBeGreaterThan(20_000);

  await page.getByRole('menuitem', { name: 'File', exact: true }).click();
  const [all] = await Promise.all([
    page.waitForEvent('download'),
    page.getByRole('menuitem', { name: 'All slides as PNG (.zip)' }).click(),
  ]);
  const zip = info.outputPath('slides.zip');
  await all.saveAs(zip);
  const entries = unzipSync(new Uint8Array(readFileSync(zip)));
  expect(Object.keys(entries).sort()).toEqual([
    'slide-01.png',
    'slide-02.png',
    'slide-03.png',
    'slide-04.png',
    'slide-05.png',
  ]);
  for (const entry of Object.values(entries)) expect([...entry.subarray(0, 8)]).toEqual(PNG);
});

test('layout: the find bar keeps the slide clear, and a zoomed slide scrolls to every edge', async ({
  page,
}) => {
  // Boxes from CSS layout: happy-dom lays nothing out.
  await ready(page);
  const stage = page.locator('.pptx-surface__stage');
  await stage.click({ position: { x: 5, y: 5 } });
  const before = (await stage.boundingBox())!;
  await page.keyboard.press('ControlOrMeta+f');
  const bar = page.getByRole('search', { name: 'Find and replace' });
  await expect(bar).toBeVisible();
  // A bar in the canvas margin reserves no room: the slide does not move.
  const find = (await bar.boundingBox())!;
  if (before.y > find.y + find.height) {
    await page.waitForTimeout(100);
    expect(Math.abs((await stage.boundingBox())!.y - before.y)).toBeLessThan(1);
  }
  await expect
    .poll(async () => (await stage.boundingBox())!.y - ((await bar.boundingBox())!.y + find.height))
    .toBeGreaterThan(0);
  await page.keyboard.press('Escape');
  await expect(bar).toBeHidden();

  // Zoomed in, every edge of the slide scrolls into view.
  await page.evaluate(() =>
    (window as unknown as { __pptxEditor: { setZoom(z: number): void } }).__pptxEditor.setZoom(2)
  );
  const edges = await page.evaluate(() => {
    const c = document.querySelector('.pptx-canvas')!;
    c.scrollLeft = 0;
    c.scrollTop = 0;
    const box = document.querySelector('.pptx-surface__stage')!.getBoundingClientRect();
    const canvas = c.getBoundingClientRect();
    return { left: box.left >= canvas.left, top: box.top >= canvas.top };
  });
  expect(edges).toEqual({ left: true, top: true });
});

test('layout: a 390 px window keeps the font pair on the bar and the Text controls on screen', async ({
  page,
}) => {
  // The toolbar measures its groups to decide what moves to More: real widths only.
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await ready(page);
  await page.getByRole('button', { name: 'Text box' }).click();
  await page.keyboard.type('alpha beta');
  await page.keyboard.press('ControlOrMeta+a');
  // The font and size of the unit test's steps (`demo-menus.test.tsx`), so the bar measures the
  // same names.
  await page.locator('[data-slot="font.family"]').click();
  await page.keyboard.type('Georgia');
  await page.keyboard.press('Enter');
  await page.locator('.pptx-toolbar__size').fill('40');
  await page.locator('.pptx-toolbar__size').press('Enter');
  await page.setViewportSize({ width: 390, height: 800 });
  await expect(
    page.locator('.pptx-toolbar > [data-group="font"] .pptx-toolbar__size')
  ).toBeVisible();
  await page.getByRole('button', { name: 'More' }).click();
  // A pointer open keeps the focus where it was: never on a control inside the card.
  await expect(page.locator('.pptx-toolbar__overflow :focus')).toHaveCount(0);
  // The last Text control in the card (Change Case and Link are hidden in this example).
  const highlight = page.locator('.pptx-toolbar__overflow [data-slot="text.highlight"]').last();
  await expect(highlight).toBeVisible();
  const box = (await highlight.boundingBox())!;
  expect(box.x + box.width).toBeLessThanOrEqual(390);
  expect(errors).toEqual([]);
});
