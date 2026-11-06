import { chromium } from '@playwright/test';
const S =
  '/private/tmp/claude-501/-Users-jedrzejblaszyk-pptx-editor/ea404932-33d7-4472-a47f-9ec5c29e2207/scratchpad/ux2/';
const b = await chromium.launch();
const p = await b.newPage({ viewport: { width: 1500, height: 950 } });
p.on('pageerror', (e) => console.log('PAGEERR', e.message));
p.on('console', (m) => {
  if (m.type() === 'error') console.log('CONSOLE', m.text());
});
await p.goto('http://localhost:5391/');
await p.waitForTimeout(3000);
const info = async (tag) => {
  const r = await p.evaluate(() => {
    const q = (s) =>
      [...document.querySelectorAll(s)].map((e) => {
        const r = e.getBoundingClientRect();
        return [Math.round(r.x), Math.round(r.y), Math.round(r.width), Math.round(r.height)].join(
          ','
        );
      });
    return {
      caret: q('.pptx-surface__caret'),
      sel: q('.pptx-surface__cell-selection'),
      handle: q('.pptx-surface__table-handle'),
      line: q('.pptx-surface__table-line'),
      ins: q('.pptx-surface__table-insert'),
      active: document.activeElement?.className,
    };
  });
  console.log(tag, JSON.stringify(r));
};
const menuItems = () =>
  p.$$eval('[role=menu]', (ms) =>
    ms.map((m) =>
      [...m.querySelectorAll('[role^=menuitem]')].map(
        (e) =>
          (e.getAttribute('aria-disabled') === 'true' || e.disabled ? '(x)' : '') +
          e.textContent.trim()
      )
    )
  );
const fp = async (tag) => {
  const vals = await p.$$eval('input', (els) =>
    els
      .filter((e) => e.offsetParent && e.type !== 'checkbox' && e.closest('[class*=panel], aside'))
      .map((e) => e.value)
  );
  console.log(tag, 'panel', JSON.stringify(vals));
};
const newTable = async (r = 3, c = 3) => {
  await p.getByRole('button', { name: 'New slide', exact: true }).click();
  await p.waitForTimeout(500);
  await p.getByText('Insert', { exact: true }).first().click();
  await p
    .getByRole('menuitem', { name: /^Table/ })
    .first()
    .hover();
  await p.waitForTimeout(300);
  await p
    .getByRole('gridcell')
    .nth((r - 1) * 10 + (c - 1))
    .click();
  await p.waitForTimeout(600);
};
const fillBasic = async () => {
  for (const [i, t] of [
    'Name',
    'Qty',
    'Price',
    'Apple',
    '3',
    '1.20',
    'Pear',
    '5',
    '0.90',
  ].entries()) {
    await p.keyboard.type(t);
    if (i < 8) await p.keyboard.press('Tab');
  }
};
await newTable();
await fillBasic();
await p.mouse.move(582, 450);
await p.waitForTimeout(150);
console.log(
  'ancestors cursor',
  await p.evaluate(() => {
    const out = [];
    let e = document.elementFromPoint(582, 450);
    while (e) {
      out.push(
        getComputedStyle(e).cursor + '@' + String(e.getAttribute('class') || e.tagName).slice(0, 30)
      );
      e = e.parentElement;
    }
    return out.join(' < ');
  })
);
const tbBtn = (n) => p.getByRole('toolbar').getByRole('button', { name: n, exact: true });
console.log(
  'pen buttons',
  await tbBtn('Pen color').count(),
  JSON.stringify(
    await tbBtn('Pen color').evaluateAll((es) =>
      es.map((e) => [e.getAttribute('aria-pressed'), e.getAttribute('aria-haspopup'), e.title])
    )
  )
);
await tbBtn('Pen color').nth(1).click();
await p.waitForTimeout(300);
await p.screenshot({ path: S + '60-pen-palette.png' });
const blue = p.locator('[aria-label="Blue"], [title="Blue"]').first();
console.log('blue count', await blue.count());
const swatches = await p.$$eval(
  '[role=dialog] [aria-label], .pptx-color-picker [aria-label], [class*=swatch]',
  (els) => els.slice(0, 20).map((e) => e.getAttribute('aria-label') || e.title)
);
console.log('swatches', JSON.stringify(swatches));
const red = p.locator('[aria-label*="Red" i]').first();
if (await red.count()) await red.click();
await p.waitForTimeout(250);
console.log(
  'pen pressed after pick',
  JSON.stringify(
    await tbBtn('Pen color').evaluateAll((es) => es.map((e) => e.getAttribute('aria-pressed')))
  )
);
await p.mouse.move(582, 440);
await p.waitForTimeout(100);
console.log(
  'cursor in pen mode',
  await p.evaluate(() => {
    const out = [];
    let e = document.elementFromPoint(582, 440);
    while (e) {
      const c = getComputedStyle(e).cursor;
      if (c !== 'auto')
        out.push(c + '@' + String(e.getAttribute('class') || e.tagName).slice(0, 30));
      e = e.parentElement;
    }
    return out.slice(0, 3).join(' < ');
  })
);
// draw along col line 582 from row1 to row2
await p.mouse.move(582, 415);
await p.mouse.down();
await p.mouse.move(582, 440, { steps: 5 });
await p.mouse.move(582, 462, { steps: 5 });
await p.mouse.up();
await p.waitForTimeout(250);
await p.mouse.move(1050, 620);
await p.screenshot({ path: S + '61-pen-drawn.png' });
await p.keyboard.press('Escape');
await p.waitForTimeout(200);
console.log(
  'pen pressed after esc',
  JSON.stringify(
    await tbBtn('Pen color').evaluateAll((es) => es.map((e) => e.getAttribute('aria-pressed')))
  )
);
await p.mouse.click(700, 425);
await p.waitForTimeout(200);
await info('click after esc');
await p.keyboard.press('Meta+z');
await p.waitForTimeout(250);
await p.mouse.move(1050, 620);
await p.screenshot({ path: S + '62-pen-undo.png' });

await b.close();
