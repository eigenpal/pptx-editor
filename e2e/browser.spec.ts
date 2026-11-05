// The checks that need a real browser, on the React demo. Every other demo flow runs in
// happy-dom as a unit test (`packages/react/src/__tests__/demo-*.test.tsx`,
// `packages/vue/test/demo-shell.test.ts`), which has no layout engine and no canvas. These
// keep only what those cannot prove: pixels drawn on a canvas, and boxes from CSS layout.
import { readFileSync } from 'node:fs';
import { unzipSync } from 'fflate';
import { expect, test, type Locator, type Page } from '@playwright/test';

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

// A host reset in the shape of Tailwind's preflight (its element rules): no margins or padding,
// headings and lists without their defaults, controls that inherit the font, and pictures as
// blocks. The editor chrome must look the same under it.
const HOST_RESET = `
  *, ::after, ::before { box-sizing: border-box; border: 0 solid; margin: 0; padding: 0; }
  html { line-height: 1.5; font-family: ui-sans-serif, system-ui, sans-serif; }
  body { line-height: inherit; }
  hr { height: 0; color: inherit; border-top-width: 1px; }
  h1, h2, h3, h4, h5, h6 { font-size: inherit; font-weight: inherit; }
  a { color: inherit; text-decoration: inherit; }
  b, strong { font-weight: bolder; }
  small { font-size: 80%; }
  table { text-indent: 0; border-color: inherit; border-collapse: collapse; }
  ol, ul, menu { list-style: none; }
  img, svg, video, canvas, audio, iframe, embed, object { display: block; vertical-align: middle; }
  img, video { max-width: 100%; height: auto; }
  button, input, select, optgroup, textarea, ::file-selector-button {
    font: inherit; letter-spacing: inherit; color: inherit; border-radius: 0;
    background-color: transparent; opacity: 1;
  }
  textarea { resize: vertical; }
`;

/** The box and text styles of every chrome element, for a before-and-after comparison. */
async function chromeStyles(page: Page): Promise<Record<string, string>> {
  return page.evaluate(() => {
    const props = [
      'margin-top',
      'margin-right',
      'margin-bottom',
      'margin-left',
      'padding-top',
      'padding-right',
      'padding-bottom',
      'padding-left',
      'border-top-width',
      'border-right-width',
      'border-bottom-width',
      'border-left-width',
      'display',
      'width',
      'height',
      'max-width',
      'font-size',
      'font-weight',
      'font-family',
      'line-height',
      'letter-spacing',
      'list-style-type',
      'color',
      'background-color',
      'border-radius',
      'text-decoration-line',
      'opacity',
      'vertical-align',
    ];
    const out: Record<string, string> = {};
    // The editor's parts and what they hold. Demo shell elements are host content, but a part
    // inside a shell wrapper (the toolbar in its dock) is chrome.
    const PART =
      "[class^='pptx-']:not([class^='pptx-editor']), [class*=' pptx-']:not([class^='pptx-editor'])";
    const chrome = [...document.querySelectorAll('.pptx-editor *')].filter((el) => {
      const part = el.closest(PART);
      const shell = el.closest('[class*="demo-"]');
      return !!part && (!shell || shell.contains(part));
    });
    chrome.forEach((el, i) => {
      const s = getComputedStyle(el);
      const parent = el.parentElement ? getComputedStyle(el.parentElement).display : '';
      const tag = `${i} ${el.tagName.toLowerCase()}.${el.getAttribute('class') ?? ''}`;
      for (const p of props) {
        // A flex or grid item ignores vertical-align.
        if (p === 'vertical-align' && /flex|grid/.test(parent)) continue;
        out[`${tag} ${p}`] = s.getPropertyValue(p);
      }
    });
    return out;
  });
}

test('embedding: a host CSS reset changes no chrome style', async ({ page }) => {
  const states: [string, () => Promise<void>][] = [
    ['the chrome', async () => {}],
    ['an open menu', () => page.locator('.pptx-menubar__trigger').first().click()],
    [
      'the find bar',
      async () => {
        await page.locator('.pptx-surface__stage').click({ position: { x: 5, y: 5 } });
        await page.keyboard.press('ControlOrMeta+f');
      },
    ],
    [
      'a dialog',
      async () => {
        await page.locator('.pptx-menubar__trigger').first().click();
        await page.getByRole('menuitem', { name: 'Page setup…' }).click();
        await expect(page.getByRole('dialog', { name: 'Page setup' })).toBeVisible();
      },
    ],
  ];
  for (const [name, open] of states) {
    await ready(page);
    // The demo shell sets its own font and line height on the editor scope. Take them off, so
    // what the host page passes down reaches the editor as it does on a plain page.
    await page.evaluate(() => document.querySelector('.demo-root')?.classList.remove('demo-root'));
    await open();
    // No tooltip under a resting pointer, and every open animation done.
    await page.mouse.move(0, 0);
    await page.waitForFunction(() =>
      document.getAnimations().every((a) => a.playState !== 'running')
    );
    const before = await chromeStyles(page);
    await page.addStyleTag({ content: HOST_RESET });
    const after = await chromeStyles(page);
    const changed = Object.keys(before).filter((k) => before[k] !== after[k]);
    expect([name, changed.map((k) => `${k}: ${before[k]} -> ${after[k]}`)]).toEqual([name, []]);
  }
});

test('embedding: the chrome stylesheet changes nothing on the host page', async ({ page }) => {
  // Every computed property of the host's own elements, with and without the stylesheet.
  const css = readFileSync(
    new URL('../packages/core/src/styles/editor.css', import.meta.url),
    'utf8'
  );
  const host =
    '<h1 class="dark">Title</h1><p>Text <a href="#x">link</a></p><button type="button">Go</button>' +
    '<input aria-label="Name"><ul><li>Item</li></ul>' +
    '<svg width="20" height="20" aria-hidden="true"><rect width="20" height="20"/></svg>' +
    '<dialog open>Note</dialog>';
  const styles = (): Promise<Record<string, string>[]> =>
    page.evaluate(() =>
      [document.documentElement, document.body, ...document.body.querySelectorAll('*')].map(
        (el) => {
          const s = getComputedStyle(el);
          const out: Record<string, string> = {};
          for (const name of s) out[name] = s.getPropertyValue(name);
          for (const name of ['--doc-primary', '--background', '--doc-z-overlay'])
            out[name] = s.getPropertyValue(name);
          return out;
        }
      )
    );
  await page.setContent(`<!doctype html><html><head></head><body>${host}</body></html>`);
  const before = await styles();
  await page.setContent(
    `<!doctype html><html><head><style>${css}</style></head><body>${host}</body></html>`
  );
  expect(await styles()).toEqual(before);
});

test('embedding: host content beside the chrome keeps the host styles', async ({ page }) => {
  // Host classes that load before the editor stylesheet, at the specificity of one class: the
  // editor's chrome defaults must not win over them inside the `.pptx-editor` scope.
  const css = readFileSync(
    new URL('../packages/core/src/styles/editor.css', import.meta.url),
    'utf8'
  );
  const host = `
    .btn { padding: 8px 12px; font-size: 20px; }
    .list { list-style: none; margin: 0; padding: 0; }
    .note { margin: 0; }
    .title { font-size: 30px; margin: 0; }
    .icon { display: none; }`;
  await page.setContent(
    `<!doctype html><html><head><style>${host}</style><style>${css}</style></head><body>` +
      '<div class="pptx-editor"><button class="btn">Go</button><ul class="list"><li>A</li></ul>' +
      '<p class="note">Note</p><h2 class="title">Title</h2>' +
      '<svg class="icon" width="10" height="10"></svg></div></body></html>'
  );
  const seen = await page.evaluate(() => {
    const get = (sel: string, prop: string): string =>
      getComputedStyle(document.querySelector(sel)!).getPropertyValue(prop);
    return {
      button: [get('.btn', 'padding'), get('.btn', 'font-size')],
      list: [get('.list', 'list-style-type'), get('.list', 'margin'), get('.list', 'padding')],
      note: get('.note', 'margin'),
      title: [get('.title', 'font-size'), get('.title', 'margin')],
      icon: get('.icon', 'display'),
    };
  });
  expect(seen).toEqual({
    button: ['8px 12px', '20px'],
    list: ['none', '0px', '0px'],
    note: '0px',
    title: ['30px', '0px'],
    icon: 'none',
  });
});

test('embedding: in a transformed or scaled host with a CSS reset, menus and tooltips open at their controls', async ({
  page,
}) => {
  // A host drawer or sheet that slides in with a transform becomes the containing block of
  // every `position: fixed` element inside it.
  await ready(page);
  const toolbar = page.locator('.pptx-toolbar').first();
  const plain = (await toolbar.boundingBox())!;
  await page.addStyleTag({
    content: `${HOST_RESET} #app { transform: translate(37px, 23px); filter: saturate(1); }`,
  });
  // The reset changes no editor box: the toolbar keeps its height.
  await expect.poll(async () => (await toolbar.boundingBox())!.height).toBe(plain.height);

  // A menu bar menu opens under its trigger.
  const trigger = page.locator('.pptx-menubar__trigger').first();
  await trigger.click();
  const menu = page.locator('.pptx-menu[role="menu"]').first();
  await expect(menu).toBeVisible();
  // Polled: the panel scales in, so its box is final only after the open animation.
  const offset = async (a: Locator, b: Locator): Promise<{ x: number; y: number }> => {
    const [p, q] = [(await a.boundingBox())!, (await b.boundingBox())!];
    return { x: Math.round(p.x - q.x), y: Math.round(p.y - (q.y + q.height)) };
  };
  await expect.poll(() => offset(menu, trigger)).toEqual({ x: 0, y: 6 });
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();

  // A host scaled up in a short window: the menu still opens under its trigger and its height
  // limit keeps it in the window.
  await page.setViewportSize({ width: 1440, height: 360 });
  await page.addStyleTag({ content: '#app { transform: translate(37px, 23px) scale(1.25); }' });
  await trigger.click();
  await expect(menu).toBeVisible();
  await expect
    .poll(async () => {
      const [m, t] = [(await menu.boundingBox())!, (await trigger.boundingBox())!];
      const under = Math.abs(m.x - t.x) < 2 && m.y >= t.y + t.height;
      return under && m.y + m.height <= 360;
    })
    .toBe(true);
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await page.setViewportSize({ width: 1440, height: 900 });

  // A drawer that clips its overflow: the menu stays in the drawer's visible box.
  await page.addStyleTag({
    content: '#app { transform: translate(37px, 23px); overflow: hidden; height: 300px; }',
  });
  await trigger.click();
  await expect(menu).toBeVisible();
  await expect
    .poll(async () => {
      const [m, d] = [(await menu.boundingBox())!, (await page.locator('#app').boundingBox())!];
      return m.y >= d.y && m.y + m.height <= d.y + d.height + 0.5;
    })
    .toBe(true);
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await page.addStyleTag({
    content: '#app { transform: translate(37px, 23px); overflow: visible; height: auto; }',
  });

  // A tooltip opens centered under its control.
  const button = page.getByRole('button', { name: 'Text box' });
  await button.hover();
  const tip = page.locator('.pptx-tooltip');
  await expect(tip).toBeVisible();
  await expect
    .poll(async () => {
      const [b, tb] = [(await button.boundingBox())!, (await tip.boundingBox())!];
      const centered = Math.abs(tb.x + tb.width / 2 - (b.x + b.width / 2)) < 2;
      const gap = tb.y - (b.y + b.height);
      return centered && gap >= 0 && gap < 16;
    })
    .toBe(true);
  await page.mouse.move(0, 0);

  // A context menu opens at the pointer.
  const stage = (await page.locator('.pptx-surface__stage').boundingBox())!;
  const at = { x: stage.x + 120, y: stage.y + 80 };
  await page.mouse.click(at.x, at.y, { button: 'right' });
  const context = page.locator('.pptx-contextmenu');
  await expect(context).toBeVisible();
  await expect
    .poll(async () => {
      // Within a pixel: the stage sits at a fractional offset.
      const c = (await context.boundingBox())!;
      return Math.abs(c.x - at.x) <= 1.5 && Math.abs(c.y - at.y) <= 1.5;
    })
    .toBe(true);
});

test('embedding: Ctrl/Cmd+F with the focus outside the editor stays with the browser', async ({
  page,
}) => {
  await ready(page);
  const prevented = await page.evaluate(() => {
    (document.activeElement as HTMLElement | null)?.blur();
    const e = new KeyboardEvent('keydown', {
      key: 'f',
      ctrlKey: true,
      bubbles: true,
      cancelable: true,
    });
    document.body.dispatchEvent(e);
    return e.defaultPrevented;
  });
  expect(prevented).toBe(false);
  await expect(page.getByRole('search', { name: 'Find and replace' })).toBeHidden();
});
