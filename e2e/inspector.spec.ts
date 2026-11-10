// Browser checks for the selection inspector: actual pointer selection, field focus, and
// panel layout. Document assertions use the same public editor that the demo exposes.
import { expect, test, type Page } from '@playwright/test';
import { unzipSync } from 'fflate';
import type { PptxEditorInstance } from '../packages/core/src/editor/index.ts';

async function selectShape(page: Page, slide: number, name: string): Promise<void> {
  await page.locator(`.pptx-rail__item[data-index="${slide}"]`).click();
  const id = await page.evaluate((name) => {
    const editor = (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor;
    return editor.snapshot().shapes!.find((shape) => shape.name === name)!.nodeId;
  }, name);
  const box = await page
    .locator('.pptx-surface__slide')
    .locator(`[data-node-id="${id}"]`)
    .boundingBox();
  expect(box).not.toBeNull();
  // The surface owns hit testing; painted SVG objects do not receive pointer events.
  await page.mouse.click(box!.x + box!.width / 2, box!.y + box!.height / 2);
  await expect
    .poll(() =>
      page.evaluate(() => {
        const editor = (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor;
        return editor.snapshot().geometry?.name;
      })
    )
    .toBe(name);
}

async function ready(page: Page): Promise<void> {
  await page.goto('/');
  await expect(page.locator('.pptx-rail__item')).toHaveCount(20, { timeout: 30_000 });
  await expect(page.locator('.pptx-surface__slide')).toBeVisible();
  await expect(page.getByText(/Some fonts in this presentation/)).toHaveCount(0);
}

test('inspector: shape sections follow pointer selection and stay inside the panel', async ({
  page,
}) => {
  await ready(page);
  await selectShape(page, 6, 'Three stops');
  const panel = page.locator('.pptx-panel');
  await expect(panel.locator('[data-format-section="fill"]').first()).toBeVisible();
  await expect(panel.locator('[data-format-section="effects"]').first()).toBeVisible();
  await expect(panel.locator('[data-format-section="picture"]')).toHaveCount(0);
  const bounds = (await panel.boundingBox())!;
  const overflowing = await panel.locator('input, select, button').evaluateAll(
    (elements, right) =>
      elements
        .filter((element) => {
          const box = element.getBoundingClientRect();
          return box.width > 0 && box.right > right + 1;
        })
        .map((element) => element.getAttribute('aria-label') ?? element.textContent),
    bounds.x + bounds.width
  );
  expect(overflowing).toEqual([]);
  await selectShape(page, 8, 'Outer shadow');
  await expect(panel.locator('[data-format-section="effects"]').first()).toBeVisible();
  await page.screenshot({ path: test.info().outputPath('shadow-inspector.png') });
});

async function slideXml(page: Page, number: number): Promise<string> {
  const bytes = await page.evaluate(() => {
    const editor = (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor;
    return Array.from(editor.save()!);
  });
  return new TextDecoder().decode(
    unzipSync(new Uint8Array(bytes))[`ppt/slides/slide${number}.xml`]
  );
}

test('inspector: numeric gradient and shadow edits preserve source formatting and undo once', async ({
  page,
}) => {
  await ready(page);
  await selectShape(page, 6, 'Three stops');
  const initialGradient = await slideXml(page, 7);
  const angle = page.locator('input[data-format-property="gradient.angle"]');
  await expect(angle).toHaveValue('45');
  await page.screenshot({ path: test.info().outputPath('gradient-inspector.png') });
  await angle.fill('120');
  await angle.press('Enter');
  await expect(angle).toBeFocused();
  await expect(angle).toHaveValue('120');
  // No resolved-color rewrite: only the angle attribute changes in the entire slide.
  expect(await slideXml(page, 7)).toBe(initialGradient.replace('ang="2700000"', 'ang="7200000"'));
  await angle.press('ControlOrMeta+z');
  await expect(angle).toHaveValue('45');
  expect(await slideXml(page, 7)).toBe(initialGradient);

  await selectShape(page, 8, 'Outer shadow');
  const initialShadow = await slideXml(page, 9);
  const distance = page.locator('input[data-format-property="shadow.distance"]');
  await expect(distance).toHaveValue('8');
  await distance.fill('16');
  await distance.press('Enter');
  await expect(distance).toBeFocused();
  expect(await slideXml(page, 9)).toBe(initialShadow.replace('dist="101600"', 'dist="203200"'));
  await distance.press('ControlOrMeta+z');
  await expect(distance).toHaveValue('8');
  expect(await slideXml(page, 9)).toBe(initialShadow);
});

test('inspector: a selection change discards a draft before another object receives it', async ({
  page,
}) => {
  await ready(page);
  await selectShape(page, 6, 'Three stops');
  const initial = await slideXml(page, 7);
  const angle = page.locator('input[data-format-property="gradient.angle"]');
  await angle.fill('175');
  // A programmatic selection change represents a host or collaborator changing the target.
  await page.evaluate(() => {
    const editor = (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor;
    const nodeId = editor.snapshot().shapes!.find((shape) => shape.name === 'Linear')!.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
  });
  await expect(angle).toHaveValue('0');
  await angle.press('Enter');
  expect(await slideXml(page, 7)).toBe(initial);
});

test('inspector: chart, diagram, and equation selections expose their specialized controls', async ({
  page,
}) => {
  await ready(page);
  const panel = page.locator('.pptx-panel');
  for (const [slide, name, section] of [
    [17, 'Chart 1', 'chart'],
    [18, 'Diagram 1', 'diagram'],
    [19, 'Equation 3', 'equation'],
  ] as const) {
    await selectShape(page, slide, name);
    await expect(panel.locator(`[data-format-section="${section}"]`).first()).toBeVisible();
    for (const other of ['chart', 'diagram', 'equation'].filter((id) => id !== section))
      await expect(panel.locator(`[data-format-section="${other}"]`)).toHaveCount(0);
  }
});

test('inspector: removing a gradient stop keeps keyboard focus in the panel', async ({ page }) => {
  await ready(page);
  await selectShape(page, 6, 'Three stops');
  const panel = page.locator('.pptx-panel');
  const stops = panel.locator('select[data-format-stop-select]');
  await stops.selectOption({ index: 1 });
  await panel
    .getByRole('group', { name: 'Gradient stop 2', exact: true })
    .getByRole('button', { name: 'Remove', exact: true })
    .click();
  await expect(stops.locator('option')).toHaveCount(2);
  await expect(stops).toBeFocused();
  await panel.getByRole('button', { name: 'Add', exact: true }).click();
  await expect(stops.locator('option')).toHaveCount(3);
  await expect(stops).toBeFocused();
  const selectedId = await stops.inputValue();
  await expect(panel.locator(`[data-format-stop-marker="${selectedId}"]`)).toHaveAttribute(
    'aria-pressed',
    'true'
  );
  await expect
    .poll(() => page.evaluate(() => !!document.activeElement?.closest('.pptx-panel')))
    .toBe(true);
});
