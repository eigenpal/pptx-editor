// An SVG picture as drawing records for the PDF export (svg-picture-drawing.ts): the icons the
// editor draws as vector art read into shapes and groups, and every other picture, including
// hostile ones, gives undefined so that the export draws the raster alternative.

import { describe, expect, test } from 'bun:test';
import { MAX_DRAWING_COMMANDS, parseSvgPath } from '../svg-picture-geometry.ts';
import {
  svgPictureDrawing,
  type SvgPictureGroup,
  type SvgPictureShape,
} from '../svg-picture-drawing.ts';

const enc = (s: string) => new TextEncoder().encode(s);
const NS = 'xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"';
const svg = (body: string, root = 'viewBox="0 0 96 96"') => enc(`<svg ${NS} ${root}>${body}</svg>`);
const shapes = (bytes: Uint8Array) =>
  svgPictureDrawing(bytes)!.items.filter((i): i is SvgPictureShape => i.kind === 'shape');

describe('svgPictureDrawing', () => {
  test('reads shapes with their solid paints in view box space', () => {
    const d = svgPictureDrawing(
      svg(
        '<rect width="96" height="96" fill="#DEEBF7"/><circle cx="48" cy="48" r="36" fill="#2E75B6"/>',
        'width="96" height="96" viewBox="0 0 96 96"'
      )
    )!;
    expect(d.viewBox).toEqual([0, 0, 96, 96]);
    const [rect, circle] = d.items as SvgPictureShape[];
    expect(rect!.fill!.color).toEqual({ r: 0xde, g: 0xeb, b: 0xf7, a: 1 });
    expect(rect!.path.map((c) => c.op).join('')).toBe('MLLLZ');
    expect(circle!.path.filter((c) => c.op === 'C')).toHaveLength(4);
    expect(circle!.fill!.color).toEqual({ r: 0x2e, g: 0x75, b: 0xb6, a: 1 });
  });

  test('takes the width and height as the view box when there is none', () => {
    expect(
      svgPictureDrawing(svg('<rect width="4" height="4"/>', 'width="40" height="20"'))!.viewBox
    ).toEqual([0, 0, 40, 20]);
  });

  test('applies style sheet rules by specificity, then the style attribute', () => {
    const [a, b, c] = shapes(
      svg(
        '<style>.accent{fill:#4472C4} path{fill:red} #me{fill:lime} .s{stroke:navy;stroke-width:3}</style>' +
          '<path class="accent" fill="#000" d="M0 0h10v10z"/>' +
          '<path id="me" class="accent" d="M0 0h10v10z"/>' +
          '<path class="s" style="fill:none" d="M0 0h10v10z"/>'
      )
    );
    expect(a!.fill!.color).toEqual({ r: 0x44, g: 0x72, b: 0xc4, a: 1 });
    expect(b!.fill!.color).toEqual({ r: 0, g: 255, b: 0, a: 1 });
    expect(c!.fill).toBeUndefined();
    expect(c!.stroke).toMatchObject({ color: { r: 0, g: 0, b: 128, a: 1 }, width: 3 });
  });

  test('inherits paint from groups and composes transforms', () => {
    const [s] = shapes(
      svg(
        '<g fill="rgb(20, 40, 160)" transform="translate(8 8) scale(2)">' +
          '<g transform="rotate(90)"><rect width="10" height="5" fill-opacity="0.5"/></g></g>'
      )
    );
    expect(s!.fill!.color).toEqual({ r: 20, g: 40, b: 160, a: 0.5 });
    const m = s!.matrix;
    expect(m[0]).toBeCloseTo(0);
    expect(m[1]).toBeCloseTo(2);
    expect(m[2]).toBeCloseTo(-2);
    expect(m[4]).toBe(8);
  });

  test('draws a line and a polyline as strokes only', () => {
    const [line, poly] = shapes(
      svg(
        '<line x1="8" y1="48" x2="80" y2="48" stroke="currentColor" stroke-width="6" stroke-linecap="round"/>' +
          '<polyline points="60,28 82,48 60,68" fill="none" stroke="#000" stroke-linejoin="round"/>'
      )
    );
    expect(line!.fill).toBeUndefined();
    expect(line!.stroke).toMatchObject({ width: 6, cap: 'round' });
    expect(poly!.stroke!.join).toBe('round');
  });

  test('draws an element opacity as a layer when it has a fill and a stroke', () => {
    const d = svgPictureDrawing(
      svg(
        '<rect width="10" height="10" fill="red" opacity="0.5"/>' +
          '<rect width="10" height="10" fill="red" stroke="blue" opacity="0.5"/>' +
          '<g opacity="0.25"><rect width="10" height="10"/><rect width="5" height="5"/></g>'
      )
    )!;
    expect(d.items[0]).toMatchObject({ kind: 'shape', fill: { color: { a: 0.5 } } });
    expect(d.items[1]).toMatchObject({ kind: 'group', opacity: 0.5 });
    expect((d.items[2] as SvgPictureGroup).items).toHaveLength(2);
  });

  test('expands use links and symbols, and clips to a clip path', () => {
    const d = svgPictureDrawing(
      svg(
        '<defs><symbol id="dot" viewBox="0 0 10 10"><circle cx="5" cy="5" r="4"/></symbol>' +
          '<clipPath id="c"><rect width="96" height="96" rx="12"/></clipPath></defs>' +
          '<g clip-path="url(#c)" style="fill:navy;stroke:none">' +
          '<use xlink:href="#dot" x="10" y="10" width="20" height="20"/></g>'
      )
    )!;
    const group = d.items[0] as SvgPictureGroup;
    expect(group.kind).toBe('group');
    expect(group.clip!.path.length).toBeGreaterThan(5);
    const symbol = group.items[0] as SvgPictureGroup;
    const dot = symbol.items[0] as SvgPictureShape;
    // The symbol's view box of 10 fits the 20 × 20 viewport at 10, 10.
    expect(dot.matrix).toEqual([2, 0, 0, 2, 10, 10]);
    expect(dot.fill!.color).toEqual({ r: 0, g: 0, b: 128, a: 1 });
  });

  test('skips hidden elements and elements of other namespaces', () => {
    const d = svgPictureDrawing(
      svg(
        '<rect width="4" height="4" display="none"/><g visibility="hidden"><rect width="4" height="4"/></g>' +
          '<x:data xmlns:x="urn:example"/><title>T</title><rect width="2" height="2"/>'
      )
    )!;
    expect(d.items).toHaveLength(1);
  });

  test('refuses what the editor draws as its raster alternative', () => {
    const refused = {
      script: svg('<script>alert(1)</script><rect width="4" height="4"/>'),
      externalUse: svg('<use href="http://example.com/a.svg#x"/>'),
      externalImage: svg('<image href="http://example.com/a.png" width="4" height="4"/>'),
      dataImage: svg('<image href="data:image/png;base64,AAAA" width="4" height="4"/>'),
      gradient: svg(
        '<linearGradient id="g"><stop offset="0" stop-color="red"/></linearGradient>' +
          '<rect width="4" height="4" fill="url(#g)"/>'
      ),
      entity: enc(
        '<?xml version="1.0"?><!DOCTYPE svg [<!ENTITY a "aaaa">]>' +
          `<svg ${NS} viewBox="0 0 4 4"><text>&a;</text></svg>`
      ),
      animation: svg('<rect width="4" height="4"><animate attributeName="x" to="4"/></rect>'),
      noViewBox: svg('<rect width="4" height="4"/>', ''),
      notSvg: enc('<html><body/></html>'),
      styleImport: svg('<style>@import url(http://example.com/a.css);</style>'),
    };
    for (const [name, bytes] of Object.entries(refused))
      expect([name, svgPictureDrawing(bytes)]).toEqual([name, undefined]);
  });

  test('refuses icon features it does not draw', () => {
    const refused = {
      nested: svg('<svg viewBox="0 0 4 4"><rect width="4" height="4"/></svg>'),
      boxClip: svg(
        '<clipPath id="c" clipPathUnits="objectBoundingBox"><rect width="1" height="1"/></clipPath>' +
          '<rect width="4" height="4" clip-path="url(#c)"/>'
      ),
      twoShapeClip: svg(
        '<clipPath id="c"><rect width="1" height="1"/><rect x="2" width="1" height="1"/></clipPath>' +
          '<rect width="4" height="4" clip-path="url(#c)"/>'
      ),
    };
    for (const [name, bytes] of Object.entries(refused))
      expect([name, svgPictureDrawing(bytes)]).toEqual([name, undefined]);
  });

  test('refuses hostile nesting, long paths, and use cycles quickly', () => {
    const deep = svg(`${'<g>'.repeat(400)}<rect width="4" height="4"/>${'</g>'.repeat(400)}`);
    const longPath = svg(`<path d="M0 0${' l1 1'.repeat(50_000)}"/>`);
    const cycle = svg('<g id="a"><use href="#b"/></g><g id="b"><use href="#a"/></g>');
    // Ten links to the level below, eight levels deep: 10^8 empty groups if expanded.
    let fan = '<g id="l0"/>';
    for (let k = 1; k <= 8; k += 1)
      fan += `<g id="l${k}">${`<use href="#l${k - 1}"/>`.repeat(10)}</g>`;
    const start = performance.now();
    expect(svgPictureDrawing(deep)).toBeUndefined();
    expect(svgPictureDrawing(longPath)).toBeUndefined();
    expect(svgPictureDrawing(cycle)).toBeUndefined();
    expect(svgPictureDrawing(svg(fan))).toBeUndefined();
    expect(performance.now() - start).toBeLessThan(2_000);
  });
});

describe('parseSvgPath', () => {
  test('reads relative commands, smooth curves, and packed arc flags', () => {
    const path = parseSvgPath('m10 10h5v5H10zM0 0q5 -5 10 0t10 0c0 5 5 5 5 0s5 -5 5 0a5 5 0 105 5');
    expect(path[0]).toEqual({ op: 'M', x: 10, y: 10 });
    expect(path[1]).toEqual({ op: 'L', x: 15, y: 10 });
    expect(path[2]).toEqual({ op: 'L', x: 15, y: 15 });
    expect(path[3]).toEqual({ op: 'L', x: 10, y: 15 });
    expect(path[4]).toEqual({ op: 'Z' });
    expect(path.at(-1)).toMatchObject({ op: 'C', x: 35, y: 5 });
    expect(path.every((c) => c.op !== 'Q')).toBe(true);
  });

  test('draws an arc as curves that end on the arc', () => {
    const path = parseSvgPath('M0 0 A10 10 0 0 1 20 0');
    const curves = path.filter((c) => c.op === 'C');
    expect(curves).toHaveLength(2);
    // The half circle passes through its lowest point at 10, -10 (sweep 1 turns clockwise in y down).
    expect(curves[0]).toMatchObject({ x: expect.closeTo(10, 6), y: expect.closeTo(-10, 6) });
    expect(curves[1]).toMatchObject({ x: 20, y: 0 });
  });

  test('keeps the path up to an error, as SVG renders it', () => {
    expect(parseSvgPath('M0 0 L10 10 L oops 5')).toHaveLength(2);
    expect(parseSvgPath('L10 10')).toHaveLength(0);
  });

  test('stops a path past the command limit', () => {
    expect(() => parseSvgPath(`M0 0${' l1 1'.repeat(MAX_DRAWING_COMMANDS + 1)}`)).toThrow();
  });
});
