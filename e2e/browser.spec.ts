// The checks that need a real browser, on the React demo. Every other demo flow runs in
// happy-dom as a unit test (`packages/react/src/__tests__/demo-*.test.tsx`,
// `packages/vue/test/demo-shell.test.ts`), which has no layout engine and no canvas. These
// keep only what those cannot prove: pixels drawn on a canvas, and boxes from CSS layout.
import { readFileSync } from 'node:fs';
import { unzipSync } from 'fflate';
import { expect, test, type Locator, type Page } from '@playwright/test';
import { textDeck } from '../scripts/bench/text-deck.ts';

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

test('PDF export: File > Presentation as PDF downloads one page per slide', async ({
  page,
}, info) => {
  // The round trip to the development server's `/api/convert` (examples/shared/pdf-export-plugin.ts).
  await ready(page);
  await page.getByRole('menuitem', { name: 'File', exact: true }).click();
  const [pdf] = await Promise.all([
    page.waitForEvent('download', { timeout: 60_000 }),
    page.getByRole('menuitem', { name: 'Presentation as PDF' }).click(),
  ]);
  expect(pdf.suggestedFilename()).toMatch(/\.pdf$/);
  const file = info.outputPath('deck.pdf');
  await pdf.saveAs(file);
  const bytes = readFileSync(file);
  expect(bytes.subarray(0, 5).toString('latin1')).toBe('%PDF-');
  expect(bytes.toString('latin1').match(/\/Type \/Page\b(?!s)/g)).toHaveLength(5);
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

  // At the fit zoom, a wider surface padding (a host style that comes late) fits the slide again.
  const fitWidth = (await stage.boundingBox())!.width;
  await page.evaluate(() =>
    document
      .querySelector<HTMLElement>('.pptx-surface')!
      .style.setProperty('--doc-slide-padding', '120px')
  );
  await expect.poll(async () => (await stage.boundingBox())!.width).toBeLessThan(fitWidth - 100);
  await page.evaluate(() =>
    document
      .querySelector<HTMLElement>('.pptx-surface')!
      .style.removeProperty('--doc-slide-padding')
  );
  await expect.poll(async () => (await stage.boundingBox())!.width).toBeCloseTo(fitWidth, 0);

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
  const toolbar = page.getByRole('toolbar', { name: 'Formatting' });
  await toolbar.locator('[data-slot="font.family"]').click();
  await page.keyboard.type('Georgia');
  await page.keyboard.press('Enter');
  await toolbar.locator('.pptx-toolbar__size').fill('40');
  await toolbar.locator('.pptx-toolbar__size').press('Enter');
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

/** Hold `requestIdleCallback` until `window.releaseIdle()`, so no idle paint runs before. */
const HOLD_IDLE = (): void => {
  const w = window as unknown as {
    requestIdleCallback: (cb: IdleRequestCallback, o?: IdleRequestOptions) => number;
    cancelIdleCallback: (id: number) => void;
    releaseIdle: () => void;
  };
  const real = {
    request: w.requestIdleCallback.bind(window),
    cancel: w.cancelIdleCallback.bind(window),
  };
  const held = new Map<number, IdleRequestCallback>();
  let next = 1;
  w.requestIdleCallback = (cb) => {
    const id = next++;
    held.set(id, cb);
    return id;
  };
  w.cancelIdleCallback = (id) => void held.delete(id);
  w.releaseIdle = () => {
    w.requestIdleCallback = real.request;
    w.cancelIdleCallback = real.cancel;
    for (const cb of held.values()) real.request(cb);
    held.clear();
  };
};

/** Which rail items have a painted thumbnail, by slide index, and which are in the list's view. */
async function railState(page: Page) {
  return page.evaluate(() => {
    const list = document.querySelector('.pptx-rail__list')!;
    const view = list.getBoundingClientRect();
    const painted: number[] = [];
    const inView: number[] = [];
    for (const item of document.querySelectorAll<HTMLElement>('.pptx-rail__item')) {
      const index = Number(item.dataset.index);
      if (item.querySelector('.pptx-rail__thumb > svg')) painted.push(index);
      const r = item.getBoundingClientRect();
      if (r.bottom > view.top && r.top < view.bottom && r.height > 0) inView.push(index);
    }
    return { painted, inView, mounted: document.querySelectorAll('.pptx-rail__item').length };
  });
}

test('long deck: the rail paints the thumbnails in view first, then more up to its limit', async ({
  page,
}) => {
  // Visibility from CSS layout and the browser's IntersectionObserver: happy-dom has neither.
  // With idle time held, only the first-paint burst, the items in view (or near it), and the
  // ones that scroll into view can paint: the observer path, alone.
  // A slow CPU (Chromium's CPU throttling) spreads the rail's renders over many tasks: a
  // container the rail releases and asks for again must not paint out of view.
  const SLIDES = 1000;
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Emulation.setCPUThrottlingRate', { rate: 6 });
  await page.addInitScript(HOLD_IDLE);
  await ready(page);
  const base64 = Buffer.from(textDeck(SLIDES)).toString('base64');
  const opened = await page.evaluate(async (data) => {
    const bytes = Uint8Array.from(atob(data), (c) => c.charCodeAt(0));
    const editor = (
      window as unknown as {
        __pptxEditor: {
          load(b: Uint8Array, o: { startSlide: number }): Promise<unknown>;
          save(): Uint8Array | null;
          snapshot(): { pendingSlides?: number };
        };
      }
    ).__pptxEditor;
    await editor.load(bytes, { startSlide: 600 });
    // Finish parsing before testing paint priority. Background progress updates each allow
    // a first-paint burst, so their timing must not decide which offscreen items paint here.
    // The next test separately checks input while a long presentation opens progressively.
    const saved = editor.save();
    return { saved: saved !== null, pending: editor.snapshot().pendingSlides ?? 0 };
  }, base64);
  expect(opened).toEqual({ saved: true, pending: 0 });
  // The current slide's item is mounted and scrolled into view, and the items in view paint.
  await expect
    .poll(async () => {
      const { painted, inView } = await railState(page);
      return inView.length > 0 && inView.includes(600) && inView.every((i) => painted.includes(i));
    })
    .toBe(true);
  // Far from the view, past the overscan and the burst at the top: nothing paints yet.
  const band = (painted: number[]) => painted.filter((i) => i >= 100 && i <= 500);
  expect(band((await railState(page)).painted)).toEqual([]);
  // Every item mounts, the end scrolls into view, and the last item paints while the band waits.
  await expect.poll(async () => (await railState(page)).mounted, { timeout: 30_000 }).toBe(SLIDES);
  await page.locator('.pptx-rail__list').evaluate((el) => (el.scrollTop = el.scrollHeight));
  const last = page.locator('.pptx-rail__item').last().locator('.pptx-rail__thumb > svg');
  await expect(last).toHaveCount(1);
  expect(band((await railState(page)).painted)).toEqual([]);
  // In idle time, thumbnails paint up to the rail's limit (`MAX_PAINTED` in `thumbnails.ts`),
  // and no further: each painted thumbnail holds memory. The items in view stay painted.
  const MAX_PAINTED = 200;
  await page.evaluate(() => (window as unknown as { releaseIdle(): void }).releaseIdle());
  await expect
    .poll(async () => (await railState(page)).painted.length, { timeout: 30_000 })
    .toBe(MAX_PAINTED);
  await page.waitForTimeout(1_000);
  const idle = await railState(page);
  expect(idle.painted.length).toBeLessThanOrEqual(MAX_PAINTED);
  expect(idle.inView.every((i) => idle.painted.includes(i))).toBe(true);
});

test('long deck: the first slide takes input while the rest opens, and every slide is there after', async ({
  page,
}) => {
  // Frames and tasks of a real browser: the rail mounts over frames, and the slides parse in
  // tasks between them.
  await ready(page);
  const n = 3000;
  await page.locator('input[type=file][accept*=".pptx"]').setInputFiles({
    name: 'long.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: Buffer.from(textDeck(n)),
  });
  await page.waitForFunction(
    () => document.querySelector('.pptx-surface__slide')?.textContent?.includes('Slide 1 title'),
    undefined,
    { timeout: 60_000 }
  );
  type Ed = {
    snapshot(): { pendingSlides?: number; slides: { pending?: true }[] };
    session: { layout(i: number): { items: { element: { name?: string; nodeId: string } }[] } };
    setSelection(s: unknown): void;
    focus(): void;
  };
  // Type into the first slide at once.
  await page.evaluate(() => {
    const ed = (window as unknown as { __pptxEditor: Ed }).__pptxEditor;
    const body = ed.session.layout(0).items.find((x) => x.element.name === 'Body')!;
    const at = { paragraph: 0, offset: 0 };
    ed.setSelection({ kind: 'text', nodeId: body.element.nodeId, anchor: at, focus: at });
    ed.focus();
  });
  await page.keyboard.type('Hello ');
  await expect(page.locator('.pptx-surface__slide')).toContainText('Hello');
  // The rest opens: every rail item mounts, and no slide is pending.
  await expect(page.locator('.pptx-rail__item')).toHaveCount(n, { timeout: 60_000 });
  await expect
    .poll(
      () =>
        page.evaluate(
          () =>
            (window as unknown as { __pptxEditor: Ed }).__pptxEditor.snapshot().pendingSlides ?? 0
        ),
      { timeout: 60_000 }
    )
    .toBe(0);
  // The thumbnails in view paint.
  await expect(
    page.locator('.pptx-rail__item').first().locator('.pptx-rail__thumb > svg')
  ).toHaveCount(1);
});

test('shape drag: the preview puts the shape where the release does', async ({ page }) => {
  // Boxes of the painted SVG: happy-dom lays nothing out.
  await ready(page);
  await page.locator('.pptx-rail__item').nth(1).click();
  const shape = page.locator('.pptx-surface__slide [data-kind="shape"]').first();
  const box = (await shape.boundingBox())!;
  const id = await shape.getAttribute('data-node-id');
  // During a drag, the shape is hidden and its preview painted after it.
  const at = () =>
    page.locator(
      `.pptx-surface__slide :is([data-preview-of="${id}"], [data-node-id="${id}"]:not([display="none"]))`
    );
  // Press on the shape's border band (its top edge) and drag it.
  const x = box.x + box.width / 2;
  const y = box.y + 2;
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(x + 20, y + 10);
  await page.mouse.move(x + 60, y + 30, { steps: 4 });
  const during = (await at().boundingBox())!;
  expect(during.x).toBeGreaterThan(box.x + 30);
  await page.mouse.up();
  const after = (await at().boundingBox())!;
  for (const k of ['x', 'y', 'width', 'height'] as const)
    expect(Math.abs(after[k] - during[k])).toBeLessThan(1.5);
  // Escape during a drag puts the shape back.
  await page.mouse.move(after.x + after.width / 2, after.y + 2);
  await page.mouse.down();
  await page.mouse.move(after.x + after.width / 2 + 80, after.y + 40, { steps: 4 });
  await page.keyboard.press('Escape');
  await page.mouse.up();
  const back = (await at().boundingBox())!;
  expect(Math.abs(back.x - after.x)).toBeLessThan(1.5);
});

test('shape resize: the preview matches the release, and Escape on a handle cancels', async ({
  page,
}) => {
  // Boxes of the painted SVG: happy-dom lays nothing out.
  await ready(page);
  await page.locator('.pptx-rail__item').nth(1).click();
  const shape = page.locator('.pptx-surface__slide [data-kind="shape"]').first();
  const id = await shape.getAttribute('data-node-id');
  const at = () =>
    page.locator(
      `.pptx-surface__slide :is([data-preview-of="${id}"], [data-node-id="${id}"]:not([display="none"]))`
    );
  const box = (await shape.boundingBox())!;
  await page.mouse.click(box.x + box.width / 2, box.y + 2);
  const handle = (await page.locator('[data-handle="se"]').boundingBox())!;
  const hx = handle.x + handle.width / 2;
  const hy = handle.y + handle.height / 2;
  await page.mouse.move(hx, hy);
  await page.mouse.down();
  await page.mouse.move(hx - 60, hy + 40, { steps: 5 });
  const during = (await at().boundingBox())!;
  const text = await at().textContent();
  await page.mouse.up();
  const after = (await at().boundingBox())!;
  for (const k of ['x', 'y', 'width', 'height'] as const)
    expect(Math.abs(after[k] - during[k])).toBeLessThan(1.5);
  expect(await at().textContent()).toBe(text);
  // Escape during a resize puts the shape back, and the release writes nothing.
  const corner = (await page.locator('[data-handle="se"]').boundingBox())!;
  await page.mouse.move(corner.x + corner.width / 2, corner.y + corner.height / 2);
  await page.mouse.down();
  await page.mouse.move(corner.x + 120, corner.y + 90, { steps: 5 });
  await page.keyboard.press('Escape');
  await page.mouse.up();
  const back = (await at().boundingBox())!;
  for (const k of ['x', 'y', 'width', 'height'] as const)
    expect(Math.abs(back[k] - after[k])).toBeLessThan(1.5);
});

test('long deck: slide edits leave no thumbnail in view empty', async ({ page }) => {
  // Layout boxes and the browser's IntersectionObserver: happy-dom has neither.
  await ready(page);
  await page.locator('input[type=file][accept*=".pptx"]').setInputFiles({
    name: 'long.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: Buffer.from(textDeck(200)),
  });
  await expect(page.locator('.pptx-rail__item')).toHaveCount(200, { timeout: 60_000 });
  const run = (command: unknown) =>
    page.evaluate((c) => {
      const ed = (window as unknown as Record<string, unknown>).__pptxEditor as {
        exec(c: unknown): { ok: boolean };
      };
      return ed.exec(c).ok;
    }, command);
  const goTo = (index: number) =>
    page.evaluate((i) => {
      const ed = (window as unknown as Record<string, unknown>).__pptxEditor as {
        goToSlide(i: number): void;
      };
      ed.goToSlide(i);
    }, index);
  const painted = () => page.locator('.pptx-rail__thumb > svg').count();
  const scrolled = () => page.locator('.pptx-rail__list').evaluate((el) => el.scrollTop);
  // On slide 101, delete slide 3, then go back to slide 101: the rail follows it down.
  await goTo(100);
  expect(await run({ type: 'deleteSlides', indices: [2] })).toBe(true);
  await goTo(100);
  await expect.poll(painted, { timeout: 30_000 }).toBe(199);
  await page.waitForTimeout(300);
  // Watch for a thumbnail in the rail's box that is empty at a frame.
  await page.evaluate(() => {
    const w = window as unknown as { __emptyInView: number };
    w.__emptyInView = 0;
    const list = document.querySelector('.pptx-rail__list')!;
    const check = (): void => {
      const r = list.getBoundingClientRect();
      for (const t of document.querySelectorAll('.pptx-rail__thumb')) {
        const b = t.getBoundingClientRect();
        if (b.bottom > r.top && b.top < r.bottom && !t.firstElementChild) w.__emptyInView += 1;
      }
      requestAnimationFrame(check);
    };
    requestAnimationFrame(check);
  });
  // Undo puts slide 3 back above the view, with slide 101 still current.
  const top = await scrolled();
  expect(top).toBeGreaterThan(5000);
  expect(await run({ type: 'undo' })).toBe(true);
  await expect.poll(painted, { timeout: 30_000 }).toBe(200);
  expect(Math.abs((await scrolled()) - top)).toBeLessThan(400);
  expect(
    await page.evaluate(() => (window as unknown as { __emptyInView: number }).__emptyInView)
  ).toBe(0);
});

test('IME: the slide shows the composing text in place, and the model takes it once', async ({
  page,
}) => {
  // Chromium's own composition through the DevTools protocol: real `composition*` and
  // `insertCompositionText` events, and real layout of the composed text.
  await ready(page);
  const title = page.locator('.pptx-surface__slide [data-kind="shape"]', {
    hasText: 'PPTX Editor',
  });
  const id = await title.first().getAttribute('data-node-id');
  await page.evaluate((nodeId) => {
    const at = { paragraph: 0, offset: 4 };
    window.__pptxEditor.setSelection({ kind: 'text', nodeId: nodeId!, anchor: at, focus: at });
    document.querySelector<HTMLElement>('.pptx-surface__input')!.focus();
  }, id);
  const revision = () => page.evaluate(() => window.__pptxEditor.snapshot().revision);
  const before = await revision();
  const slide = page.locator('.pptx-surface__slide');
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Input.imeSetComposition', {
    text: ' にほん',
    selectionStart: 4,
    selectionEnd: 4,
  });
  await expect(slide).toContainText('PPTX にほん Editor');
  await expect(page.locator('.pptx-surface__composition line')).toHaveCount(1);
  // The composed text is not in the model yet: no revision, so no undo step and no peer op.
  expect(await revision()).toBe(before);
  await cdp.send('Input.insertText', { text: ' 日本' });
  await expect(slide).toContainText('PPTX 日本 Editor');
  await expect(page.locator('.pptx-surface__composition line')).toHaveCount(0);
  expect(await revision()).toBe(before + 1);
  expect(await page.evaluate(() => window.__pptxEditor.find.matches('日本').length)).toBe(1);
});
