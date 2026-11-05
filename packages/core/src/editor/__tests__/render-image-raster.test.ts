// A slide or shape picture with no DOM: the editor's own paint goes into a DOM-free SVG
// document, and a rasterizer that the host supplies draws it. Here a stub rasterizer records the
// request and answers a PNG header of the asked size.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { CommandRefusal } from '../../automation/commands/op-batch.ts';
import { openPresentation, type DocumentSession } from '../index.ts';
import { renderImageWith, type ImageRasterRequest } from '../render-image.ts';

const IN = 914_400;

/** The PNG signature and an `IHDR` chunk: enough for the size check. */
function pngOf(width: number, height: number): Uint8Array {
  const b = new Uint8Array(33);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  const view = new DataView(b.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  b.set([8, 6, 0, 0, 0], 24);
  return b;
}

async function session(): Promise<DocumentSession> {
  const opened = await openPresentation(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            name: 'Box',
            x: IN,
            y: IN,
            cx: 2 * IN,
            cy: IN,
            fill: solid('4472C4'),
            text: para(run('Hello')),
          }),
        },
      ],
    })
  );
  if (!opened.ok) throw new Error('fixture');
  return opened.session;
}

function recording(): {
  requests: ImageRasterRequest[];
  rasterize: (r: ImageRasterRequest) => Uint8Array;
} {
  const requests: ImageRasterRequest[] = [];
  return {
    requests,
    rasterize: (request) => {
      requests.push(request);
      return pngOf(request.width, request.height);
    },
  };
}

/** The node id of the first shape of slide 1. */
function boxNode(s: DocumentSession): string {
  return s.layout(0)!.items[0]!.element.nodeId;
}

describe('pictures with a rasterizer', () => {
  test('a slide picture is the painted SVG at the asked size, with its fonts', async () => {
    const s = await session();
    const { requests, rasterize } = recording();
    const png = await renderImageWith(s, 0, { width: 640 }, rasterize);
    expect(png).toEqual(pngOf(640, 360));
    const [request] = requests;
    expect([request!.width, request!.height]).toEqual([640, 360]);
    expect(request!.svg.startsWith('<svg xmlns="http://www.w3.org/2000/svg"')).toBe(true);
    expect(request!.svg).toContain('width="640"');
    expect(request!.svg).toContain('data-kind="background"');
    expect(request!.svg).toContain('Hello');
    // Each face the SVG names is listed with its bytes, and inline as an @font-face rule.
    expect(request!.fonts.length).toBeGreaterThan(0);
    for (const font of request!.fonts) {
      expect(request!.svg).toContain(`font-family="${font.family}`);
      expect(request!.svg).toContain(`@font-face{font-family:&quot;${font.family}&quot;`);
      expect(font.bytes.length).toBeGreaterThan(0);
    }
    s.dispose();
  });

  test('only faces that may travel go to the rasterizer, in the SVG and in fonts', async () => {
    const real = await session();
    // The same faces as fonts of this machine (`local`): none may leave the engine.
    const faces = real.faces.map((face) => ({ ...face, origin: 'local' as const }));
    const s = new Proxy(real, {
      get: (target, key) => {
        if (key === 'faces') return faces;
        const value: unknown = Reflect.get(target, key, target);
        return typeof value === 'function' ? value.bind(target) : value;
      },
    });
    const { requests, rasterize } = recording();
    await renderImageWith(s, 0, { width: 320 }, rasterize);
    const [request] = requests;
    expect(request!.svg).toContain('Hello');
    expect(request!.fonts).toEqual([]);
    expect(request!.svg).not.toContain('@font-face');
    real.dispose();
  });

  test('a shape picture has no slide background and is cut to the shape box', async () => {
    const s = await session();
    const { requests, rasterize } = recording();
    await renderImageWith(s, 0, { shapeNodeId: boxNode(s) }, rasterize);
    const [request] = requests;
    // Two inches by one inch at 96 pixels per inch.
    expect([request!.width, request!.height]).toEqual([192, 96]);
    expect(request!.svg).not.toContain('data-kind="background"');
    expect(request!.svg).toContain(`viewBox="${IN} ${IN} ${2 * IN} ${IN}"`);
    s.dispose();
  });

  test('refuses an answer that is not a PNG of the asked size, and a rasterizer that throws', async () => {
    const s = await session();
    const answers: (() => Uint8Array)[] = [
      () => pngOf(10, 10),
      () => new TextEncoder().encode('not a picture'),
      () => {
        throw new Error('rasterizer down');
      },
    ];
    for (const answer of answers) {
      const error = await renderImageWith(s, 0, {}, answer).catch((e: unknown) => e);
      expect(error).toBeInstanceOf(CommandRefusal);
    }
    const wrongType = await renderImageWith(s, 0, {}, () => 'png' as unknown as Uint8Array).catch(
      (e: unknown) => e
    );
    expect(wrongType).toBeInstanceOf(CommandRefusal);
    s.dispose();
  });

  test('a shape that the slide does not draw refuses', async () => {
    const s = await session();
    const error = await renderImageWith(s, 0, { shapeNodeId: 'none' }, () => pngOf(1, 1)).catch(
      (e: unknown) => e
    );
    expect(error).toBeInstanceOf(CommandRefusal);
    s.dispose();
  });
});
