import { expect, test, type Locator } from '@playwright/test';
import type { PptxEditorInstance } from '../packages/core/src/editor/index.ts';

/** Measure computed colors against the surfaces on both sides of the boundary. */
async function contrast(control: Locator, focus = false): Promise<number[]> {
  return control.evaluate((element, focus) => {
    type RGB = [number, number, number];
    const rgba = (color: string) => color.match(/[\d.]+/g)!.map(Number);
    const over = (color: string, bg: RGB): RGB => {
      const [r, g, b, alpha = 1] = rgba(color);
      return [r!, g!, b!].map((c, i) => c * alpha + bg[i]! * (1 - alpha)) as RGB;
    };
    const background = (node: Element | null): RGB =>
      node
        ? over(getComputedStyle(node).backgroundColor, background(node.parentElement))
        : [255, 255, 255];
    const luminance = (rgb: RGB) =>
      rgb.reduce((sum, c, i) => {
        const s = c / 255;
        return (
          sum +
          (s <= 0.04045 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4) * [0.2126, 0.7152, 0.0722][i]!
        );
      }, 0);
    const indicator = focus ? (element.querySelector('.pptx-rail__thumb') ?? element) : element;
    const style = getComputedStyle(indicator);
    const color = focus ? style.outlineColor : style.borderTopColor;
    const visible = focus
      ? style.outlineStyle !== 'none' &&
        parseFloat(style.outlineWidth) >= 2 &&
        element.matches(':focus-visible')
      : style.borderTopStyle !== 'none' && parseFloat(style.borderTopWidth) >= 1;
    if (!visible) return [0];
    const outside = background(indicator.parentElement);
    let opacity = 1;
    for (let node: Element | null = indicator; node; node = node.parentElement)
      opacity *= Number(getComputedStyle(node).opacity);
    return (focus ? [outside] : [background(element), outside]).map((bg) => {
      const ink = over(color, bg).map(
        (channel, i) => channel * opacity + bg[i]! * (1 - opacity)
      ) as RGB;
      const a = luminance(ink),
        b = luminance(bg);
      return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
    });
  }, focus);
}
async function check(control: Locator, focus = false) {
  await expect(control).toBeVisible();
  if (focus) {
    // Establish keyboard modality, then move focus to each representative control.
    await control.page().keyboard.press('Tab');
    await control.focus();
    await expect(control).toBeFocused();
  }
  for (const ratio of await contrast(control, focus)) expect(ratio).toBeGreaterThanOrEqual(3);
}

for (const theme of ['light', 'dark']) {
  test(`contrast: focus and dialog boundaries in ${theme}`, async ({ page }) => {
    await page.goto(`/?theme=${theme}`);
    await expect(page.locator('.pptx-rail__item')).toHaveCount(20, { timeout: 30_000 });
    await expect(page.locator('.pptx-surface__slide')).toBeVisible();
    await check(page.locator('.pptx-menubar__trigger').first(), true);
    await check(page.locator('.pptx-toolbar__button[aria-disabled="true"]').first(), true);
    await check(page.locator('.pptx-rail__item').first(), true);
    if (!(await page.locator('.pptx-notes__text').isVisible()))
      await page.getByRole('button', { name: 'Show speaker notes', exact: true }).click();
    await check(page.locator('.pptx-notes__text'), true);

    await page.locator('.pptx-menubar__trigger').first().click();
    await page.getByRole('menuitem', { name: 'Page setup…' }).click();
    const setup = page.getByRole('dialog', { name: 'Page setup' });
    await expect(setup).toBeVisible();
    for (const field of await setup
      .locator('.pptx-field input:not([type="radio"]), .pptx-field select')
      .all()) {
      await check(field);
      await check(field, true);
    }
    await page.screenshot({ path: test.info().outputPath(`${theme}-page-setup.png`) });
    await page.keyboard.press('Escape');

    await page.locator('.pptx-rail__item[data-index="9"]').click();
    await page.evaluate(() => {
      const editor = (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor;
      const picture = editor.snapshot().shapes!.find((s) => s.name === 'Landscape')!;
      editor.setSelection({ kind: 'shapes', nodeIds: [picture.nodeId] });
    });
    await page.getByRole('button', { name: 'Alt text', exact: true }).first().click();
    const dialog = page.getByRole('dialog', { name: 'Alt text', exact: true });
    await expect(dialog).toBeVisible();
    const description = dialog.getByRole('textbox', { name: 'Description', exact: true });
    const title = dialog.getByRole('textbox', { name: 'Title', exact: true });
    for (const field of [description, title]) {
      await check(field);
      await check(field, true);
    }
    await description.fill('x'.repeat(4001));
    await dialog.getByRole('button', { name: 'OK', exact: true }).click();
    await expect(description).toHaveAttribute('aria-invalid', 'true');
    await check(description);
    await check(description, true);
    await page.screenshot({ path: test.info().outputPath(`${theme}-alt-text-error.png`) });
  });
}

for (const theme of ['light', 'dark']) {
  test(`contrast: packaged palette on all default surfaces in ${theme}`, async ({ page }) => {
    const surfaces = ['surface', 'card', 'bg', 'bg-subtle', 'bg-hover', 'bg-input'];
    await page.setContent(
      `<div class="pptx-editor ${theme === 'dark' ? 'dark' : ''}">${surfaces
        .map(
          (surface) =>
            `<section style="padding:12px;background:var(--doc-${surface})"><button class="pptx-toolbar__button">${surface}</button><button class="pptx-toolbar__button" aria-disabled="true">Disabled ${surface}</button><label class="pptx-field">Text<input value="Text"></label><label class="pptx-field">Number<input type="number" value="1"></label><label class="pptx-field">Choice<select><option>One</option></select></label><textarea class="pptx-dialog__textarea" aria-label="Description">Description</textarea></section>`
        )
        .join('')}</div>`
    );
    await page.addStyleTag({ path: 'packages/core/src/styles/editor.css' });
    for (const button of await page.getByRole('button').all()) await check(button, true);
    for (const field of await page.locator('input, select, textarea').all()) {
      await check(field);
      await check(field, true);
      await field.evaluate((element) => element.setAttribute('aria-invalid', 'true'));
      await check(field);
    }
  });
}
