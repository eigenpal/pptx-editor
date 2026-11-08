// The geometry glue over the canonical tree. The evaluator itself (guide formulas, arcs, the
// path-start rule, the caps, and the preset table) is tested in
// packages/ooxml-core/src/__tests__/drawingml-geometry.test.ts.

import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { parseXmlPart, readPptx } from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import {
  createFixedMeasurer,
  evaluateGeometry,
  layoutSlide,
  presetNames,
  type EvaluatedGeometry,
  type PathCommand,
} from '../index.ts';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

/** Evaluate an `a:custGeom` of the canonical tree, given its guide and path list markup. */
function custom(gd: string, paths: string, w: number, h: number): EvaluatedGeometry {
  const r = parseXmlPart(
    '/geom.xml',
    `<a:custGeom ${NS}><a:avLst/><a:gdLst>${gd}</a:gdLst><a:ahLst/><a:cxnLst/>` +
      `<a:rect l="l" t="t" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>`
  );
  if (!r.ok) throw new Error('bad xml');
  return evaluateGeometry({ kind: 'custom', element: r.part.root }, w, h);
}

const finite = (c: PathCommand): boolean =>
  Object.values(c).every((v) => typeof v === 'string' || Number.isFinite(v));

describe('custom geometry over the canonical tree', () => {
  test('reads guides, path sizes, and commands, and draws nothing of a path without a first moveTo', () => {
    const g = custom(
      '<a:gd name="half" fmla="*/ w 1 2"/>',
      '<a:path w="100" h="100" fill="none"><a:lnTo><a:pt x="100" y="0"/></a:lnTo>' +
        '<a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo>' +
        '<a:close/></a:path><a:path><a:moveTo><a:pt x="half" y="0"/></a:moveTo></a:path>',
      200,
      100
    );
    // The path starts with lnTo, so it draws nothing, also after its later moveTo.
    expect(g.paths[0]).toEqual({ commands: [], fill: 'none', stroke: true });
    expect(g.paths[1].commands).toEqual([{ op: 'M', x: 100, y: 0 }]);
    expect(g.textRect).toEqual({ l: 0, t: 0, r: 200, b: 100 });
  });

  test('a preset spec draws the preset with its adjust values', () => {
    const rect = evaluateGeometry({ kind: 'preset', preset: 'rect', adjust: new Map() }, 10, 20);
    expect(rect.paths[0].commands.map((c) => c.op).join('')).toBe('MLLLZ');
    const unknown = evaluateGeometry({ kind: 'preset', preset: 'nope', adjust: new Map() }, 1, 1);
    expect(unknown.unknownPreset).toBe('nope');
    expect(presetNames()).toContain('upArrow');
  });
});

describe('arc sweeps from guide formulas (R1)', () => {
  const arcShape = (fmla: string): string =>
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="F"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' +
    '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>' +
    `<a:custGeom><a:avLst/><a:gdLst><a:gd name="g1" fmla="${fmla}"/></a:gdLst><a:ahLst/><a:cxnLst/>` +
    '<a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="1000000" h="1000000">' +
    '<a:moveTo><a:pt x="1000000" y="500000"/></a:moveTo>' +
    '<a:arcTo wR="500000" hR="500000" stAng="0" swAng="g1"/><a:close/></a:path></a:pathLst>' +
    '</a:custGeom></p:spPr></p:sp>';

  test('a 2^31 × 2^31 sweep opens in well under 50 ms with a bounded path', () => {
    const r = readPptx(
      buildPptx({ slides: [{ shapes: arcShape('*/ 2147483647 2147483647 1'), layout: 'blank' }] })
    );
    if (!r.ok) throw new Error('open failed');
    const slide = resolveSlide(r.pkg, 0)!;
    const t0 = performance.now();
    const layout = layoutSlide(slide, createFixedMeasurer());
    const ms = performance.now() - t0;
    const item = layout.items[0];
    if (item.kind !== 'shape') throw new Error('not a shape');
    const cmds = item.geometry.paths[0].commands;
    expect(ms).toBeLessThan(50);
    // Move, at most two turns of quarter-arc Béziers, close.
    expect(cmds.length).toBeLessThanOrEqual(11);
    expect(cmds.every(finite)).toBe(true);
  });
});
