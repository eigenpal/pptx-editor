// The Vue slide show for a screen reader: the slide is not one unnamed image. Its text and the
// alt text of its pictures are in the tree, as in the React twin (`slideshow.test.tsx`).

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, type App, type Component } from 'vue';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import {
  PptxEditorRoot,
  PptxEditorSlideshow,
  PptxEditorViewport,
  useSlideshow,
  type UseSlideshowReturn,
} from '../src/index.ts';

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const PICTURE =
  '<p:pic><p:nvPicPr><p:cNvPr id="3" name="Lake" descr="A lake at dusk"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
  '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
  '<p:spPr><a:xfrm><a:off x="5486400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';

function deck(): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({ id: 2, x: IN, y: IN, cx: 3 * IN, cy: IN, text: para(run('One')) }) + PICTURE,
        rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
      },
    ],
    media: [
      { name: 'image1.png', bytes: new Uint8Array([137, 80, 78, 71]), contentType: 'image/png' },
    ],
  });
}

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
});

describe('PptxEditorSlideshow for a screen reader', () => {
  test('reads the slide text and the picture alt text', async () => {
    let show: UseSlideshowReturn | null = null;
    const Probe = defineComponent({
      setup() {
        show = useSlideshow();
        return () => null;
      },
    });
    let ready: PptxEditorInstance | null = null;
    host = document.createElement('div');
    document.body.append(host);
    app = createApp(
      defineComponent({
        setup: () => () =>
          h(PptxEditorRoot, { onReady: (e: PptxEditorInstance) => (ready = e) }, () =>
            h(PptxEditorViewport, null, () => [h(Probe), h(PptxEditorSlideshow)])
          ),
      }) as Component
    );
    app.mount(host);
    await nextTick();
    const editor = ready as PptxEditorInstance | null;
    expect((await editor!.load(deck())).ok).toBe(true);
    await nextTick();
    show!.start('start');
    await nextTick();
    const svg = document.querySelector('.pptx-show__slide svg')!;
    expect(svg.getAttribute('role')).toBe('group');
    expect(svg.textContent).toContain('One');
    expect(svg.querySelector('[role="img"]')?.getAttribute('aria-label')).toBe('A lake at dusk');
  });
});
