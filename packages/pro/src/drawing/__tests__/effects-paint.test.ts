/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The engine's SVG paint with and without the Pro drawing renderers: glow, inner shadow,
// reflection, and soft edge paint only after `installDrawing` (or `drawingModule`); the outer
// shadow paints in both.

import { afterAll, describe, expect, spyOn, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '@pptx-editor.dev/core/store';
import { resolveSlide } from '@pptx-editor.dev/core/cascade';
import {
  createFixedMeasurer,
  installDrawingRenderers,
  layoutSlide,
} from '@pptx-editor.dev/core/layout';
import { paintSlide } from '@pptx-editor.dev/core/output';
import { drawingModule, installDrawing } from '../index.ts';
import { drawingEffectsRenderer } from '../effects-renderer.ts';

// A DOM for this file only.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

function paint(shapes: string): SVGSVGElement {
  const r = readPptx(buildPptx({ slides: [{ shapes, layout: 'blank' }] }));
  if (!r.ok) throw new Error('open failed');
  const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
  return paintSlide(doc, layout, { images: () => undefined, idPrefix: 't' });
}

const SHADOW =
  '<a:outerShdw blurRad="38100" dist="38100" dir="2700000"><a:srgbClr val="000000"/></a:outerShdw>';

/** A filled shape with every effect: glow, inner shadow, reflection, soft edge, and a shadow. */
const everyEffect = shape({
  id: 2,
  x: 1_000_000,
  y: 1_000_000,
  cx: 2_000_000,
  cy: 1_000_000,
  fill: '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>',
  effects:
    '<a:effectLst><a:glow rad="63500"><a:srgbClr val="FF0000"/></a:glow>' +
    '<a:innerShdw blurRad="63500" dist="50800" dir="5400000"><a:srgbClr val="000000"/></a:innerShdw>' +
    SHADOW +
    '<a:reflection blurRad="6350" stA="50000" endPos="55000" dist="5000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>' +
    '<a:softEdge rad="12700"/></a:effectLst>',
});

/** The effects that a painted slide shows. */
function painted(root: SVGSVGElement): Record<string, boolean> {
  const results = [...root.querySelectorAll('filter > *')].map((e) => e.getAttribute('result'));
  return {
    shadow: root.querySelector('[data-effect="shadow"]') !== null,
    reflection:
      root.querySelector('[data-effect="reflection"]') !== null &&
      root.querySelector('mask') !== null,
    glow: results.includes('glow'),
    innerShadow: results.includes('inner'),
    softEdge: results.includes('soft'),
  };
}

// Each test file runs in its own process (`scripts/test/run-parallel.mjs`), so the order below
// is the whole story: the free tier first, then the module.
describe('the Pro drawing tier', () => {
  test('without the module, only the outer shadow paints', () => {
    expect(painted(paint(everyEffect))).toEqual({
      shadow: true,
      reflection: false,
      glow: false,
      innerShadow: false,
      softEdge: false,
    });
  });

  test('the drawing module restores glow, inner shadow, reflection, and soft edge', () => {
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    const module = drawingModule({ licenseKey: 'test' });
    expect(module.id).toBe('drawing');
    // The install comes after the first paint above, so it warns once; the same renderers again
    // stay silent.
    expect(warn).toHaveBeenCalledTimes(1);
    installDrawing();
    expect(warn).toHaveBeenCalledTimes(1);
    warn.mockRestore();
    expect(painted(paint(everyEffect))).toEqual({
      shadow: true,
      reflection: true,
      glow: true,
      innerShadow: true,
      softEdge: true,
    });
  });

  test('a renderer that throws paints the shape without its effects', () => {
    const fail = (): never => {
      throw new Error('boom');
    };
    installDrawingRenderers({ effects: { body: fail, glow: fail, reflection: fail } });
    const root = paint(everyEffect);
    expect(painted(root)).toMatchObject({ shadow: true, glow: false, reflection: false });
    expect(root.querySelector('[data-kind="shape"] path')).not.toBeNull();
    installDrawingRenderers({ effects: drawingEffectsRenderer });
    expect(painted(paint(everyEffect)).glow).toBe(true);
  });

  test('removing the renderer returns to the free tier', () => {
    installDrawingRenderers({ effects: undefined });
    expect(painted(paint(everyEffect))).toMatchObject({ shadow: true, glow: false });
    installDrawing();
  });
});

describe('effect filter regions', () => {
  test('a shadow and a glow keep the outline that paints past the box', () => {
    const root = paint(
      shape({
        id: 2,
        x: 1_000_000,
        y: 1_000_000,
        cx: 2_000_000,
        cy: 1_000_000,
        fill: '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>',
        line: '<a:ln w="88900"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
        effects:
          '<a:effectLst><a:glow rad="12700"><a:srgbClr val="FF0000"/></a:glow>' +
          '<a:outerShdw blurRad="0" dist="0"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>',
      })
    );
    const filters = [...root.querySelectorAll('filter')];
    expect(filters.length).toBe(2);
    // Half the 7 pt outline lies outside the box; neither region may cut it off.
    for (const f of filters)
      expect(Number(f.getAttribute('x'))).toBeLessThanOrEqual(1_000_000 - 88_900);
  });

  test('the region keeps a miter join and arrowheads', () => {
    const region = (line: string, prst = 'rect'): number => {
      const root = paint(
        shape({
          id: 2,
          x: 1_000_000,
          y: 1_000_000,
          cx: 2_000_000,
          cy: 1_000_000,
          prst,
          line,
          effects:
            '<a:effectLst><a:glow rad="12700"><a:srgbClr val="FF0000"/></a:glow></a:effectLst>',
        })
      );
      return Number(root.querySelector('filter')!.getAttribute('x'));
    };
    const solid = '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>';
    // A miter limit of 8 reaches up to 4 widths past the corner.
    expect(region(`<a:ln w="25400">${solid}<a:miter lim="800000"/></a:ln>`)).toBeLessThanOrEqual(
      1_000_000 - 4 * 25_400
    );
    // An arrowhead on an open path reaches 3 widths past its end.
    expect(
      region(`<a:ln w="25400">${solid}<a:round/><a:tailEnd type="triangle"/></a:ln>`, 'line')
    ).toBeLessThanOrEqual(1_000_000 - 3 * 25_400);
  });

  test('a soft edge and an inner shadow keep a callout tail outside the box', () => {
    const root = paint(
      shape({
        id: 2,
        x: 1_000_000,
        y: 1_000_000,
        cx: 2_000_000,
        cy: 1_000_000,
        prst: 'wedgeRectCallout',
        // The tail tip lies 1.5 box heights below the box.
        avLst: '<a:gd name="adj1" fmla="val 0"/><a:gd name="adj2" fmla="val 200000"/>',
        fill: '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>',
        effects:
          '<a:effectLst><a:innerShdw blurRad="63500" dist="50800" dir="5400000">' +
          '<a:srgbClr val="000000"/></a:innerShdw><a:softEdge rad="12700"/></a:effectLst>',
      })
    );
    const f = root.querySelector('filter')!;
    const bottom = Number(f.getAttribute('y')) + Number(f.getAttribute('height'));
    expect(bottom).toBeGreaterThanOrEqual(1_000_000 + 2.5 * 1_000_000);
  });

  test('a turned callout keeps its tail in the region too', () => {
    const root = paint(
      shape({
        id: 2,
        x: 1_000_000,
        y: 1_000_000,
        cx: 2_000_000,
        cy: 1_000_000,
        rot: 5_400_000,
        prst: 'wedgeRectCallout',
        avLst: '<a:gd name="adj1" fmla="val 0"/><a:gd name="adj2" fmla="val 200000"/>',
        fill: '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>',
        effects: '<a:effectLst><a:softEdge rad="12700"/></a:effectLst>',
      })
    );
    // Turned 90° about the box center (2 000 000, 1 500 000), the tail tip moves to x = 0.
    const f = root.querySelector('filter')!;
    expect(Number(f.getAttribute('x'))).toBeLessThanOrEqual(0);
  });
});

describe('glow paint', () => {
  test('a skewed reflection keeps its linear coefficients to 1e-6', () => {
    const root = paint(
      shape({
        id: 2,
        x: 1_000_000,
        y: 1_000_000,
        cx: 2_000_000,
        cy: 1_000_000,
        fill: '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>',
        effects:
          '<a:effectLst><a:reflection stA="50000" endPos="55000" dir="5400000" sy="-100000"' +
          ' kx="-1200000" algn="bl" rotWithShape="0"/></a:effectLst>',
      })
    );
    const transform = root
      .querySelector('[data-effect="reflection"] use')
      ?.getAttribute('transform');
    const linear = /matrix\(([^)]*)\)/
      .exec(transform ?? '')![1]!
      .split(' ')
      .slice(0, 4)
      .map(Number);
    // The skew of 20 degrees writes tan 20 degrees, not its value to two decimals.
    expect(linear.some((v) => Math.abs(Math.abs(v) - Math.tan(Math.PI / 9)) < 1e-5)).toBe(true);
  });

  test('the glow is a copy under the content, and the text paints with no filter', () => {
    const root = paint(
      shape({
        id: 2,
        x: 1_000_000,
        y: 1_000_000,
        cx: 2_000_000,
        cy: 1_000_000,
        text: '<a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>Glow</a:t></a:r></a:p>',
        effects:
          '<a:effectLst><a:glow rad="63500"><a:srgbClr val="FF0000"/></a:glow></a:effectLst>',
      })
    );
    const copy = root.querySelector('[data-effect="glow"]')!;
    expect(copy.firstElementChild?.tagName.toLowerCase()).toBe('use');
    // The glow filter paints the glow alone: no merge with the source graphic.
    expect(root.querySelector('feMerge')).toBeNull();
    const text = root.querySelector('text');
    expect(text).not.toBeNull();
    for (let n: Element | null = text; n && n !== root; n = n.parentElement)
      expect(n.hasAttribute('filter')).toBe(false);
  });
});
