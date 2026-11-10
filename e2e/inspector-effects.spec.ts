import { expect, test, type Page } from '@playwright/test';
import { unzipSync, strFromU8 } from 'fflate';
import type { PptxEditorInstance } from '../packages/core/src/editor/index.ts';

const browserErrors = new WeakMap<Page, string[]>();
test.beforeEach(({ page }) => {
  const errors: string[] = [];
  browserErrors.set(page, errors);
  page.on('pageerror', (error) => errors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });
});
test.afterEach(({ page }) => {
  expect(browserErrors.get(page)).toEqual([]);
});

async function ready(page: Page) {
  await page.goto('/');
  await expect(page.locator('.pptx-rail__item')).toHaveCount(20, { timeout: 30_000 });
  await page.locator('.pptx-panel').evaluate((el) =>
    Object.assign((el as HTMLElement).style, {
      width: '320px',
      minWidth: '320px',
      maxWidth: '320px',
    })
  );
}
async function select(page: Page, slide: number, name: string) {
  await page.locator(`.pptx-rail__item[data-index="${slide}"]`).click();
  const id = await page.evaluate(
    (name) =>
      (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor
        .snapshot()
        .shapes!.find((s) => s.name === name)!.nodeId,
    name
  );
  const shape = page.locator('.pptx-surface__slide').locator(`[data-node-id="${id}"]`);
  await expect(shape).toBeVisible();
  const bounds = (await shape.boundingBox())!;
  await page.mouse.click(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2);
  await expect(page.locator('.pptx-panel .pptx-panel__title').first()).toHaveText(name);
}
async function xml(page: Page, slide: number) {
  const bytes = await page.evaluate(() =>
    Array.from((window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor.save()!)
  );
  return strFromU8(unzipSync(new Uint8Array(bytes))[`ppt/slides/slide${slide}.xml`]!);
}
async function section(page: Page, id: string) {
  const toggle = page.locator(`[data-format-toggle="${id}"]`);
  if (!(await toggle.evaluate((el) => el.closest('details')!.open))) await toggle.click();
}
async function pick(page: Page, slot: string, name: string) {
  await page.locator('.pptx-panel').locator(`[data-slot="${slot}"]`).click();
  await page
    .getByRole('menuitem', { name, exact: true })
    .or(page.getByRole('menuitemradio', { name, exact: true }))
    .click();
}
async function undo(page: Page) {
  await page
    .getByRole('toolbar', { name: 'Formatting' })
    .getByRole('button', { name: /^Undo/ })
    .click();
}

test('shadow panel: all numeric properties, theme color, invalid drafts, Escape and toolbar undo', async ({
  page,
}) => {
  await ready(page);
  await select(page, 8, 'Outer shadow');
  await section(page, 'effects');
  const original = await xml(page, 9);
  const shadow = original.match(/<a:outerShdw[^>]*>[\s\S]*?<\/a:outerShdw>/)![0];
  for (const [property, value, from, to] of [
    ['blur', '6.25', 'rotWithShape="0"', 'rotWithShape="0" blurRad="79375"'],
    ['distance', '12.5', 'dist="101600"', 'dist="158750"'],
    ['angle', '125', 'dir="2700000"', 'dir="7500000"'],
    [
      'opacity',
      '35',
      '<a:schemeClr val="accent1"/>',
      '<a:schemeClr val="accent1"><a:alpha val="35000"/></a:schemeClr>',
    ],
  ]) {
    const input = page.locator(`[data-format-property="shadow.${property}"]`);
    const previous = await input.inputValue();
    await input.click();
    await input.fill(value!);
    await input.press('Enter');
    await expect(input).toBeFocused();
    await expect(input).toHaveValue(value!);
    expect(await xml(page, 9)).toBe(original.replace(shadow, shadow.replace(from!, to!)));
    await undo(page);
    await expect(input).toHaveValue(previous);
    expect(await xml(page, 9)).toBe(original);
    for (const bad of [
      '',
      '-1',
      property === 'angle' ? '361' : property === 'opacity' ? '101' : '151',
    ]) {
      await input.fill(bad);
      await input.press('Enter');
      await expect(input).toHaveValue(bad);
      await expect(input).toHaveAttribute('aria-invalid', 'true');
      const error = await input.getAttribute('aria-describedby');
      await expect(page.locator(`[id="${error}"]`)).toBeVisible();
      expect(await xml(page, 9)).toBe(original);
      if (property === 'blur' && bad === '-1')
        await page.screenshot({ path: test.info().outputPath('shadow-invalid-negative.png') });
      await input.press('Escape');
      await expect(input).toHaveValue(previous);
    }
  }
  const color = page.locator('.pptx-panel select[name="shadow.color"]');
  await color.selectOption('accent2');
  expect(await xml(page, 9)).toBe(
    original.replace(
      shadow,
      shadow.replace('<a:schemeClr val="accent1"/>', '<a:schemeClr val="accent2"/>')
    )
  );
  await undo(page);
  expect(await xml(page, 9)).toBe(original);
  await page.screenshot({ path: test.info().outputPath('shadow-click-edit.png') });
});

test('picture panel: add a shadow and edit its numeric properties using visible controls', async ({
  page,
}) => {
  await ready(page);
  await select(page, 9, 'Grayscale');
  await section(page, 'effects');
  await expect(page.locator('[data-format-property="shadow.blur"]')).toHaveCount(0);
  const original = await xml(page, 10);
  await page.screenshot({ path: test.info().outputPath('picture-shadow-before.png') });
  await page.locator('.pptx-panel [data-slot="shape.effects"]').click();
  await page.getByRole('menuitem', { name: 'Shadow', exact: true }).click();
  await page.getByRole('menuitemradio', { name: 'Offset: bottom right', exact: true }).click();
  const distance = page.locator('[data-format-property="shadow.distance"]');
  await expect(distance).toBeVisible();
  const withShadow = await xml(page, 10);
  expect(withShadow).not.toBe(original);
  await distance.fill('17');
  await distance.press('Enter');
  await expect(distance).toHaveValue('17');
  const changed = await xml(page, 10);
  expect(changed).toContain('dist="215900"');
  await page.screenshot({ path: test.info().outputPath('picture-shadow-after.png') });
  await undo(page);
  expect(await xml(page, 10)).toBe(withShadow);
  await undo(page);
  expect(await xml(page, 10)).toBe(original);
  await expect(distance).toHaveCount(0);
  await page.screenshot({ path: test.info().outputPath('picture-shadow.png') });
});

test('picture panel: corrections, recolor, transparency, crop, outline and alt text undo cleanly', async ({
  page,
}) => {
  await ready(page);
  await select(page, 9, 'Landscape');
  const original = await xml(page, 10);
  for (const [slot, option, source] of [
    ['image.corrections', 'Brightness: +20% Contrast: -20%', '<a:lum'],
    ['image.color', 'Grayscale', '<a:grayscl'],
    ['image.transparency', 'Transparency: 50%', '<a:alphaModFix'],
    ['image.cropFit', 'Aspect ratio 1:1', '<a:srcRect'],
  ]) {
    await pick(page, slot!, option!);
    const changed = await xml(page, 10);
    expect(changed).not.toBe(original);
    expect(changed).toContain(source!);
    await undo(page);
    expect(await xml(page, 10)).toBe(original);
  }
  await section(page, 'line');
  const width = page.locator('[data-format-property="line.width"]');
  await width.fill('3.75');
  await width.press('Enter');
  expect(await xml(page, 10)).toBe(original.replace('w="19050"', 'w="47625"'));
  await undo(page);
  expect(await xml(page, 10)).toBe(original);
  await page.locator('.pptx-panel').getByRole('button', { name: 'Alt text', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Alt text', exact: true });
  await dialog
    .getByRole('textbox', { name: 'Description', exact: true })
    .fill('Purple mountains reflected in a lake.');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  expect(await xml(page, 10)).toContain('descr="Purple mountains reflected in a lake."');
  await undo(page);
  expect(await xml(page, 10)).toBe(original);
});
