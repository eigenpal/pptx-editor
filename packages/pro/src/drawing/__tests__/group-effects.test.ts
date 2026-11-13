/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The effects of a group (`p:grpSpPr/a:effectLst`) with and without the Pro drawing renderers:
// glow, inner shadow, reflection, soft edge, and blur paint over the group's content as one
// picture only after `installDrawing`; the outer shadow paints in both.

import { afterAll, describe, expect, spyOn, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, group, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '@pptx-editor.dev/core/store';
import { resolveSlide } from '@pptx-editor.dev/core/cascade';
import { createFixedMeasurer, layoutSlide } from '@pptx-editor.dev/core/layout';
import { paintSlide } from '@pptx-editor.dev/core/output';
import { installDrawing } from '../index.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IN = 914_400;

function paint(shapes: string): SVGSVGElement {
  const r = readPptx(buildPptx({ slides: [{ shapes, layout: 'blank' }] }));
  if (!r.ok) throw new Error('open failed');
  const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
  return paintSlide(doc, layout, { images: () => undefined, idPrefix: 't', hitTargets: false });
}

const effects =
  '<a:effectLst><a:blur rad="38100"/><a:glow rad="63500"><a:srgbClr val="FF0000"/></a:glow>' +
  '<a:innerShdw blurRad="63500" dist="50800" dir="5400000"><a:srgbClr val="000000"/></a:innerShdw>' +
  '<a:outerShdw blurRad="38100" dist="38100" dir="2700000"><a:srgbClr val="000000"/></a:outerShdw>' +
  '<a:reflection blurRad="6350" stA="50000" endPos="55000" dist="5000" dir="5400000" ' +
  'sy="-100000" algn="bl" rotWithShape="0"/><a:softEdge rad="12700"/></a:effectLst>';

const deck = group(
  { id: 10, x: IN, y: IN, cx: 3 * IN, cy: IN, rot: 600_000, effects },
  shape({ id: 11, x: IN, y: IN, cx: IN, cy: IN, fill: solid('4472C4') }) +
    shape({
      id: 12,
      x: 2.5 * IN,
      y: IN,
      cx: 1.5 * IN,
      cy: IN,
      prst: 'ellipse',
      fill: solid('ED7D31'),
    })
);

/** The effects that the group paints, and how many shapes each copy holds. */
function painted(root: SVGSVGElement) {
  const node = root.querySelector('[data-kind="group"]')!;
  const results = [...root.querySelectorAll('filter > *')].map((e) => e.getAttribute('result'));
  const copied = (selector: string): number => {
    const href = node.querySelector(`${selector} use`)?.getAttribute('href');
    const target = href ? root.querySelector(`[id="${href.slice(1)}"]`) : null;
    return target?.querySelectorAll('[data-kind="shape"]').length ?? 0;
  };
  return {
    shadow: copied('[data-effect="shadow"]'),
    reflection: copied('[data-effect="reflection"]'),
    glow: copied('[data-effect="glow"]'),
    innerShadow: results.includes('inner'),
    softEdge: results.includes('soft'),
    // Each shape of the group paints once, outside the copies.
    shapes: node.querySelectorAll(':scope [data-kind="shape"]').length,
  };
}

// Each test file runs in its own process, so the free tier comes first, then the module.
describe('group effects and the Pro drawing tier', () => {
  test('without the module, only the outer shadow paints, of both shapes', () => {
    expect(painted(paint(deck))).toEqual({
      shadow: 2,
      reflection: 0,
      glow: 0,
      innerShadow: false,
      softEdge: false,
      shapes: 2,
    });
  });

  test('with the module, every effect paints over the group content as one picture', () => {
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    installDrawing();
    warn.mockRestore();
    const root = paint(deck);
    expect(painted(root)).toEqual({
      shadow: 2,
      reflection: 2,
      glow: 2,
      innerShadow: true,
      softEdge: true,
      shapes: 2,
    });
    // The children carry no effects of their own: the group's filters wrap them both.
    expect(root.querySelectorAll('[data-kind="shape"] [data-effect]').length).toBe(0);
    // The blur filters the content, and the body filter (inner shadow and soft edge) inside it
    // holds both shapes.
    const shown = root.querySelector('[data-kind="group"]')!.lastElementChild!;
    expect(shown.getAttribute('filter')).toMatch(/^url\(#/);
    const body = shown.querySelector('[id] > [filter]')!;
    expect(body.querySelectorAll('[data-kind="shape"]').length).toBe(2);
  });
});
