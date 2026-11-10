import { expect, test, type Page } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import type { PptxEditorInstance } from '../packages/core/src/editor/index.ts';

async function ready(page: Page): Promise<void> {
  await page.goto('/');
  await expect(page.locator('.pptx-rail__item')).toHaveCount(20, { timeout: 30_000 });
  await expect(page.locator('.pptx-surface__slide')).toBeVisible();
}

/** Select public targets to isolate inspector interaction from canvas gesture tests. */
async function choose(page: Page, slide: number, names: string[]): Promise<void> {
  await page.evaluate(
    ({ slide, names }) => {
      const e = (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor;
      e.goToSlide(slide);
      e.setSelection({
        kind: 'shapes',
        nodeIds: names.map((name) => e.snapshot().shapes!.find((s) => s.name === name)!.nodeId),
      });
    },
    { slide, names }
  );
}

async function xml(page: Page, slide: number): Promise<string> {
  const bytes = await page.evaluate(() =>
    Array.from((window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor.save()!)
  );
  return strFromU8(unzipSync(new Uint8Array(bytes))[`ppt/slides/slide${slide}.xml`]!);
}
async function undo(page: Page): Promise<void> {
  await page.evaluate(() => {
    const e = (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor;
    if (!e.exec({ type: 'undo' }).ok) throw new Error('Undo refused');
  });
}
async function reopen(page: Page): Promise<void> {
  await page.evaluate(async () => {
    const e = (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor;
    if (!(await e.load(e.save()!)).ok) throw new Error('Saved deck did not reopen');
  });
}
async function openSection(page: Page, section: string): Promise<void> {
  const toggle = page.locator(`[data-format-toggle="${section}"]`);
  if (
    (await toggle.count()) &&
    !(await toggle.evaluate((element) => element.closest('details')!.open))
  )
    await toggle.click();
}
async function menuPick(page: Page, slot: string, label: string): Promise<void> {
  const control = page.locator('.pptx-panel').locator(`[data-slot="${slot}"]`);
  await control.click();
  await page
    .getByRole('menuitem', { name: label, exact: true })
    .or(page.getByRole('menuitemradio', { name: label, exact: true }))
    .click();
}
const connector = (source: string, name: string): string =>
  source.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g)!.find((s) => s.includes(`name="${name}"`))!;

test('inspector task: imported line width and independent arrow ends retain unrelated XML', async ({
  page,
}) => {
  const parts = unzipSync(new Uint8Array(readFileSync('examples/vite/public/showcase.pptx')));
  const original = strFromU8(parts['ppt/slides/slide8.xml']!);
  const line = connector(original, 'Straight connector');
  parts['ppt/slides/slide8.xml'] = strToU8(
    original.replace(
      line,
      line
        .replace('w="25400"', 'w="28575"')
        .replace('<a:headEnd ', '<a:headEnd xmlns:x="urn:inspector:test" x:keep="yes" ')
    )
  );
  const bytes = zipSync(parts);
  await page.route('**/showcase.pptx', (route) =>
    route.fulfill({
      body: Buffer.from(bytes),
      contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    })
  );
  await ready(page);
  await choose(page, 7, ['Straight connector']);
  const panel = page.locator('.pptx-panel');
  const width = panel.locator('[data-format-property="line.width"]');
  await expect(width).toHaveValue('2.25');
  await page.screenshot({ path: test.info().outputPath('line-inspector.png') });
  const before = connector(await xml(page, 8), 'Straight connector');
  await width.fill('3.125');
  await width.press('Enter');
  await expect(width).toBeFocused();
  await expect(width).toHaveValue('3.13');
  expect(connector(await xml(page, 8), 'Straight connector')).toBe(
    before.replace('w="28575"', 'w="39688"')
  );
  await width.press('ControlOrMeta+z');
  await expect(width).toHaveValue('2.25');
  expect(connector(await xml(page, 8), 'Straight connector')).toBe(before);
  await menuPick(page, 'shape.lineStart', 'Diamond');
  const changed = connector(await xml(page, 8), 'Straight connector');
  expect(changed).toBe(before.replace('type="oval"', 'type="diamond"'));
  await reopen(page);
  await choose(page, 7, ['Straight connector']);
  expect(connector(await xml(page, 8), 'Straight connector')).toBe(changed);
  await menuPick(page, 'shape.lineEnd', 'Triangle');
  expect(connector(await xml(page, 8), 'Straight connector')).toContain(
    'type="triangle" w="lg" len="lg"'
  );
  await menuPick(page, 'shape.lineDash', 'Dash dot');
  expect(connector(await xml(page, 8), 'Straight connector')).toContain('val="dashDot"');
});

test('inspector task: multiple connectors show mixed ends and apply one undoable edit', async ({
  page,
}) => {
  await ready(page);
  await choose(page, 7, ['Straight connector', 'Elbow connector']);
  const panel = page.locator('.pptx-panel');
  await expect(panel.locator('[data-slot="shape.lineStart"]')).toBeVisible();
  await expect(panel.locator('[data-slot="shape.lineEnd"]')).toBeVisible();
  await expect(panel.getByText('Mixed', { exact: true }).first()).toBeVisible();
  const before = await xml(page, 8);
  await menuPick(page, 'shape.lineEnd', 'Triangle');
  const after = await xml(page, 8);
  for (const name of ['Straight connector', 'Elbow connector'])
    expect(connector(after, name)).toContain('<a:tailEnd type="triangle"');
  await undo(page);
  expect(await xml(page, 8)).toBe(before);
  await choose(page, 7, ['Straight connector', 'Node']);
  await expect(panel.locator('[data-slot="shape.lineStart"]')).toHaveCount(0);
  await expect(panel.locator('[data-slot="shape.lineEnd"]')).toHaveCount(0);
});

test('inspector task: text-range formatting keeps the range and survives reopen', async ({
  page,
}) => {
  await ready(page);
  await page.evaluate(() => {
    const e = (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor;
    e.goToSlide(14);
    e.exec({ type: 'insertTextBox' });
    e.input('text', 'Alpha beta');
    const s = e.snapshot().selection;
    if (s.kind !== 'text') throw new Error('Missing text caret');
    e.setSelection({
      ...s,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 5 },
    });
  });
  const panel = page.locator('.pptx-panel');
  const before = await xml(page, 15);
  await panel.getByRole('button', { name: /^Bold/ }).click();
  const bold = await xml(page, 15);
  expect(bold).not.toBe(before);
  const selected = await page.evaluate(
    () =>
      (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor.snapshot().selection
  );
  expect(selected).toMatchObject({ kind: 'text', anchor: { offset: 0 }, focus: { offset: 5 } });
  await undo(page);
  expect(await xml(page, 15)).toBe(before);
  const size = panel.getByRole('textbox', { name: 'Font size', exact: true });
  await size.fill('30');
  await size.press('Enter');
  const changed = await xml(page, 15);
  expect(changed).toMatch(/sz="3000"[^>]*>[\s\S]*?<a:t>Alpha<\/a:t>/);
  expect(changed).toContain(' beta</a:t>');
  await reopen(page);
  expect(await xml(page, 15)).toBe(changed);
});

test('inspector task: selected table cells receive only the requested border', async ({ page }) => {
  await ready(page);
  await page.evaluate(() => {
    const e = (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor;
    e.goToSlide(10);
    const nodeId = e.snapshot().shapes!.find((s) => s.kind === 'graphicFrame')!.nodeId;
    e.setSelection({
      kind: 'cells',
      nodeId,
      anchor: { row: 2, column: 1 },
      focus: { row: 2, column: 2 },
    });
  });
  const before = await xml(page, 11);
  await menuPick(page, 'table.borderTarget', 'Bottom border');
  const changed = await xml(page, 11);
  const rows = (s: string) => s.match(/<a:tr\b[^>]*>[\s\S]*?<\/a:tr>/g)!;
  const cells = (s: string) => s.match(/<a:tc\b[^>]*>[\s\S]*?<\/a:tc>/g)!;
  const prior = rows(before).map(cells);
  const next = rows(changed).map(cells);
  // Shared boundaries must agree: selected bottoms and adjacent tops change together.
  for (let row = 0; row < prior.length; row++) {
    for (let column = 0; column < prior[row]!.length; column++) {
      const side =
        column === 1 || column === 2 ? (row === 2 ? 'lnB' : row === 3 ? 'lnT' : null) : null;
      if (side) {
        expect(next[row]![column]).toContain(`<a:${side}`);
        const edge = new RegExp(`<a:${side}\\b[^>]*>[\\s\\S]*?<\\/a:${side}>`, 'g');
        const withoutEdge = (cell: string) =>
          cell.replace(edge, '').replace(/<a:tcPr([^>]*)><\/a:tcPr>/g, '<a:tcPr$1/>');
        expect(withoutEdge(next[row]![column]!)).toBe(withoutEdge(prior[row]![column]!));
      } else expect(next[row]![column]).toBe(prior[row]![column]);
    }
  }
  await undo(page);
  expect(await xml(page, 11)).toBe(before);
  await menuPick(page, 'table.borderTarget', 'Bottom border');
  await reopen(page);
  expect(await xml(page, 11)).toBe(changed);
});

test('inspector task: picture crop and transparency preserve the picture and undo', async ({
  page,
}) => {
  await ready(page);
  await choose(page, 9, ['Landscape']);
  const before = await xml(page, 10);
  await menuPick(page, 'image.cropFit', 'Aspect ratio 1:1');
  const cropped = await xml(page, 10);
  expect(cropped).not.toBe(before);
  await undo(page);
  expect(await xml(page, 10)).toBe(before);
  await menuPick(page, 'image.transparency', 'Transparency: 50%');
  const changed = await xml(page, 10);
  expect(changed).toContain('alphaModFix');
  await reopen(page);
  await choose(page, 9, ['Landscape']);
  expect(
    await page.evaluate(
      () =>
        (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor.snapshot().image
          ?.transparency
    )
  ).toBe(0.5);
  expect(await xml(page, 10)).toBe(changed);
});

test('inspector task: invalid geometry keeps its draft, explains the error, and Escape restores it', async ({
  page,
}) => {
  await ready(page);
  await choose(page, 6, ['Three stops']);
  await openSection(page, 'geometry');
  const field = page.locator('[data-format-property="geometry.width"]');
  const before = await xml(page, 7);
  const value = await field.inputValue();
  await field.fill('-2');
  await field.press('Enter');
  await expect(field).toHaveValue('-2');
  await expect(field).toHaveAttribute('aria-invalid', 'true');
  const error = await field.getAttribute('aria-describedby');
  expect(error).toBeTruthy();
  await expect(page.locator(`[id="${error}"]`)).toBeVisible();
  expect(await xml(page, 7)).toBe(before);
  await field.press('Escape');
  await expect(field).toHaveValue(value);
  await expect(field).not.toHaveAttribute('aria-invalid', 'true');
});

for (const width of [280, 320, 400]) {
  test(`inspector density: ${width}px and 200% equivalent zoom retain usable fields`, async ({
    page,
  }) => {
    await ready(page);
    await choose(page, 6, ['Three stops']);
    for (const zoom of [1, 2]) {
      await page.evaluate(
        ({ width, zoom }) => {
          const panel = document.querySelector<HTMLElement>('.pptx-panel')!;
          panel.style.width = `${width}px`;
          panel.style.minWidth = `${width}px`;
          panel.style.maxWidth = `${width}px`;
          // CSS zoom exercises doubled control geometry; it is not a native browser-zoom claim.
          panel.style.zoom = String(zoom);
        },
        { width, zoom }
      );
      const panel = page.locator('.pptx-panel');
      await expect
        .poll(() => panel.evaluate((el) => el.scrollWidth <= el.clientWidth + 1))
        .toBe(true);
      const angle = panel.locator('[data-format-property="gradient.angle"]');
      await angle.fill('90');
      await angle.press('Enter');
      await expect(angle).toHaveValue('90');
      await angle.fill('45');
      await angle.press('Enter');
      const stop = panel.locator('[data-format-stop-select]');
      await stop.selectOption({ index: 1 });
      const opacity = panel.locator('[data-format-property="gradient-stop.1.opacity"]');
      await opacity.fill('75');
      await opacity.press('Enter');
      await expect(opacity).toHaveValue('75');
      await opacity.fill('100');
      await opacity.press('Enter');
      await stop.selectOption({ index: 0 });
      await angle.scrollIntoViewIfNeeded();
      const clipped = await panel.locator('input,select,button').evaluateAll((fields) =>
        fields
          .filter((el) => {
            const rect = el.getBoundingClientRect();
            const owner = el.closest('.pptx-panel')!.getBoundingClientRect();
            return rect.width > 0 && (rect.left < owner.left - 1 || rect.right > owner.right + 1);
          })
          .map((el) => el.getAttribute('aria-label') ?? el.textContent)
      );
      expect(clipped).toEqual([]);
      await page.screenshot({ path: test.info().outputPath(`gradient-${width}-${zoom}x.png`) });
    }
  });
}

test('inspector task: specialized chart, diagram, and equation edits undo and reopen', async ({
  page,
}) => {
  await ready(page);
  const panel = page.locator('.pptx-panel');
  await choose(page, 17, ['Chart 1']);
  const titleToggle = panel.getByRole('checkbox', { name: 'Title', exact: true });
  if (await titleToggle.count()) await titleToggle.check();
  const title = panel.getByRole('textbox', { name: 'Title text', exact: true });
  await title.fill('Inspector chart title');
  await title.press('Enter');
  await expect(title).toHaveValue('Inspector chart title');
  await undo(page);
  await expect(title).not.toHaveValue('Inspector chart title');
  await title.fill('Inspector chart title');
  await title.press('Enter');
  await reopen(page);
  await choose(page, 17, ['Chart 1']);
  await expect(title).toHaveValue('Inspector chart title');

  await choose(page, 18, ['Diagram 1']);
  const node = panel.getByRole('textbox', { name: /^Text of shape 1(?:,|$)/ });
  const oldText = await node.inputValue();
  await node.fill('Inspector process node');
  await node.press('Enter');
  await expect(node).toBeFocused();
  await undo(page);
  await expect(node).toHaveValue(oldText);
  await node.fill('Inspector process node');
  await node.press('Enter');
  await reopen(page);
  await choose(page, 18, ['Diagram 1']);
  await expect(node).toHaveValue('Inspector process node');

  await choose(page, 19, ['Equation 3']);
  const shapeCount = () =>
    page.evaluate(
      () =>
        (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor.snapshot().shapes!
          .length
    );
  const originalCount = await shapeCount();
  const before = await xml(page, 20);
  await panel.getByRole('button', { name: 'Equation', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Edit equation', exact: true });
  const expression = dialog.getByRole('textbox', { name: 'Linear form', exact: true });
  await expect(expression).not.toHaveValue('');
  await dialog.getByRole('textbox', { name: 'Linear form', exact: true }).fill('x^2+5');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  const changed = await xml(page, 20);
  expect(changed).not.toBe(before);
  expect(changed).toContain('<m:t>5</m:t>');
  expect(await shapeCount()).toBe(originalCount);
  await undo(page);
  expect(await xml(page, 20)).toBe(before);
  await panel.getByRole('button', { name: 'Equation', exact: true }).click();
  await dialog.getByRole('textbox', { name: 'Linear form', exact: true }).fill('x^2+5');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await reopen(page);
  expect(await xml(page, 20)).toBe(changed);
});

test('inspector task: picture replacement fits the original frame and supports undo', async ({
  page,
}) => {
  await ready(page);
  await choose(page, 9, ['Landscape']);
  const before = await xml(page, 10);
  const png = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jvS8AAAAASUVORK5CYII=',
    'base64'
  );
  const picking = page.waitForEvent('filechooser');
  await page
    .locator('.pptx-panel')
    .getByRole('button', { name: /Change picture/i })
    .click();
  await (await picking).setFiles({ name: 'replacement.png', mimeType: 'image/png', buffer: png });
  await expect.poll(() => xml(page, 10)).not.toBe(before);
  const after = await xml(page, 10);
  const picture = (s: string) =>
    s.match(/<p:pic>[\s\S]*?<\/p:pic>/g)!.find((pic) => pic.includes('name="Landscape"'))!;
  const xfrm = (s: string) => picture(s).match(/<a:xfrm[^>]*>[\s\S]*?<\/a:xfrm>/)![0];
  const dimensions = (s: string) =>
    Object.fromEntries(
      [...xfrm(s).matchAll(/(x|y|cx|cy)="(\d+)"/g)].map((m) => [m[1], Number(m[2])])
    );
  const original = dimensions(before);
  const fitted = dimensions(after);
  expect(fitted.cx).toBe(original.cy);
  expect(fitted.cy).toBe(original.cy);
  expect(fitted.x + fitted.cx / 2).toBe(original.x + original.cx / 2);
  expect(fitted.y + fitted.cy / 2).toBe(original.y + original.cy / 2);
  expect(picture(after).match(/<p:cNvPr[^>]*\/>/)![0]).toBe(
    picture(before).match(/<p:cNvPr[^>]*\/>/)![0]
  );
  await undo(page);
  expect(await xml(page, 10)).toBe(before);
});

test('inspector task: disclosure preferences and gradient markers support the keyboard', async ({
  page,
}) => {
  await ready(page);
  await choose(page, 6, ['Three stops']);
  const panel = page.locator('.pptx-panel');
  const primary = panel.locator('[data-format-toggle]').first();
  await expect(primary).toHaveAttribute('data-format-toggle', 'fill');
  const geometry = panel.locator('[data-format-toggle="geometry"]');
  await geometry.focus();
  const initial = await geometry.evaluate((el) => el.closest('details')!.open);
  await geometry.press('Enter');
  await expect.poll(() => geometry.evaluate((el) => el.closest('details')!.open)).toBe(!initial);
  await choose(page, 6, ['Linear']);
  await expect.poll(() => geometry.evaluate((el) => el.closest('details')!.open)).toBe(!initial);
  await choose(page, 6, ['Three stops']);
  const markers = panel.locator('[data-format-stop-marker]');
  await markers.first().focus();
  await markers.first().press('ArrowRight');
  await expect(markers.nth(1)).toBeFocused();
  await expect(markers.nth(1)).toHaveAttribute('aria-pressed', 'true');
  const selector = panel.locator('[data-format-stop-select]');
  await expect(selector).toHaveValue(
    (await markers.nth(1).getAttribute('data-format-stop-marker'))!
  );
  await markers.nth(1).press('End');
  await expect(markers.last()).toBeFocused();
  await markers.last().press('Home');
  await expect(markers.first()).toBeFocused();
});

test('inspector task: a new selection category reveals its primary controls', async ({ page }) => {
  await ready(page);
  await choose(page, 6, ['Three stops']);
  const panel = page.locator('.pptx-panel');
  await openSection(page, 'geometry');
  const width = panel.locator('[data-format-property="geometry.width"]');
  await width.scrollIntoViewIfNeeded();
  await choose(page, 9, ['Landscape']);
  const primary = panel.locator('[data-format-toggle]').first();
  await expect(primary).toHaveAttribute('data-format-toggle', 'picture');
  await expect(primary).toBeInViewport();
  await expect(panel.getByRole('button', { name: /Change picture/i })).toBeInViewport();
  await page.screenshot({ path: test.info().outputPath('picture-inspector.png') });
});
