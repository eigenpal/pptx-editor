// A value in a deck can be any string, also the name of an Object prototype member. A font or a
// compound line type named `constructor` or `__proto__` lays out and paints like any unknown
// name.

import { afterAll, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor } from '../../editor/index.ts';
import { paintSlide } from '../paint.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IN = 914_400;

test.each(['constructor', '__proto__', 'toString', 'hasOwnProperty'])(
  'a font and a line type named %s lay out and paint',
  async (name) => {
    const deck = buildPptx({
      slides: [
        {
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 3 * IN,
            cy: IN,
            line: `<a:ln w="38100" cmpd="${name}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`,
            text: para(run('Text', '', `<a:latin typeface="${name}"/>`)),
          }),
        },
      ],
    });
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(deck)).ok).toBe(true);
    const layout = editor.session!.layout(0)!;
    expect(layout.items.length).toBe(1);
    expect(editor.snapshot().slides.length).toBe(1);
    const svg = paintSlide(doc, layout, { images: () => undefined, idPrefix: 'x' });
    expect(svg.textContent).toContain('Text');
    editor.destroy();
  }
);
