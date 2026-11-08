import { expect, test } from '@playwright/test';
import { buildPptx, para, run, shape } from '../scripts/fixtures/builder.ts';
import type { PptxEditorInstance } from '../packages/core/src/editor/index.ts';

test('an empty list caret paints an oversized bullet without blocking input', async ({ page }) => {
  const bullet = 'a' + '\u0301'.repeat(50_000);
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: 0,
          y: 0,
          cx: 6_000_000,
          cy: 4_000_000,
          text: para(run(''), `<a:pPr><a:buChar char="${bullet}"/></a:pPr>`),
          bodyPr: '<a:bodyPr wrap="none"/>',
        }),
      },
    ],
  });
  await page.route('**/bullet.pptx', (route) => route.fulfill({ body: Buffer.from(bytes) }));
  await page.goto('/?deck=basic');
  await page.waitForSelector('.pptx-surface__slide');
  const result = await page.evaluate(async () => {
    const editor = (window as Window & { __pptxEditor: PptxEditorInstance }).__pptxEditor;
    await editor.load(new Uint8Array(await (await fetch('/bullet.pptx')).arrayBuffer()));
    const item = editor.session!.layout(0)!.items.find((i) => i.kind === 'shape' && i.text);
    if (!item || item.kind !== 'shape') throw new Error('missing shape');
    const start = performance.now();
    editor.setSelection({
      kind: 'text',
      nodeId: item.element.nodeId,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 0 },
    });
    await new Promise((resolve) => setTimeout(resolve, 100));
    const ghost = document.querySelector('.pptx-surface__ghost-bullet');
    const ms = performance.now() - start;
    return {
      ms,
      text: ghost?.textContent,
      pieces: [...(ghost?.querySelectorAll('text') ?? [])].map((t) => t.textContent?.length ?? 0),
    };
  });
  console.log(JSON.stringify({ id: 'ghost-bullet-combining-marks-50k', ms: result.ms }));
  expect(result.ms).toBeLessThan(5000);
  expect(result.text).toBe(bullet);
  expect(result.pieces.every((length) => length <= 256)).toBe(true);
});
