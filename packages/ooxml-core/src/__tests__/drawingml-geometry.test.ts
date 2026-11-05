// Tests for the DrawingML geometry evaluator: guide formulas, arcs, the path-start rule, the
// output caps, the preset table, and the `a:custGeom` reader over a neutral element.

import { describe, expect, test } from 'bun:test';
import {
  customGeometryDefinition,
  evaluateFormula,
  evaluateGeometry,
  evaluatePresetGeometry,
  presetNames,
  type DrawingGeometryElement,
  type EvaluatedGeometry,
  type PathCommand,
  type PresetCommand,
  type PresetShapeDefinition,
} from '../drawingml-geometry.ts';
import { readXml, type XmlNode } from '../xml-reader.ts';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NAMESPACES: Readonly<Record<string, string>> = { a: A, x: 'urn:example:other' };

type XmlElement = Extract<XmlNode, { type: 'element' }>;

/** A neutral geometry element over the bounded reader's tree (prefix `a` is DrawingML). */
function neutral(el: XmlElement): DrawingGeometryElement {
  const colon = el.name.indexOf(':');
  const prefix = colon < 0 ? '' : el.name.slice(0, colon);
  return {
    namespaceUri: NAMESPACES[prefix] ?? '',
    localName: el.name.slice(colon + 1),
    attribute: (name) => (Object.hasOwn(el.attributes, name) ? el.attributes[name] : undefined),
    childElements: () =>
      el.children.filter((c): c is XmlElement => c.type === 'element').map(neutral),
  };
}

/** Evaluate an `a:custGeom` given its guide list and path list markup. */
function custom(gd: string, paths: string, w = 1_000_000, h = 1_000_000): EvaluatedGeometry {
  const r = readXml(
    `<a:custGeom><a:avLst/><a:gdLst>${gd}</a:gdLst><a:ahLst/><a:cxnLst/>` +
      `<a:rect l="l" t="t" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>`
  );
  const root = r.ok ? r.nodes.find((n): n is XmlElement => n.type === 'element') : undefined;
  if (!root) throw new Error('bad xml');
  return evaluateGeometry(customGeometryDefinition(neutral(root)), w, h);
}

const finite = (c: PathCommand): boolean =>
  Object.values(c).every((v) => typeof v === 'string' || Number.isFinite(v));

describe('guide formulas (ECMA-376 Part 1, 20.1.9.11)', () => {
  const vars = new Map([
    ['w', 200],
    ['h', 100],
  ]);

  test.each([
    ['val 7', 7],
    ['*/ w 3 4', 150],
    ['+- w h 50', 250],
    ['+/ w h 3', 100],
    ['?: h 1 2', 1],
    ['?: 0 1 2', 2],
    ['abs -5', 5],
    ['max w h', 200],
    ['min w h', 100],
    ['mod 3 4 0', 5],
    ['pin 0 w 150', 150],
    ['pin 0 -5 150', 0],
    ['sqrt 16', 4],
    ['at2 0 1', 5_400_000],
    ['cos 10 0', 10],
    ['sin 10 5400000', 10],
  ])('%s gives %d', (fmla, value) => {
    expect(evaluateFormula(fmla, vars)).toBeCloseTo(value, 6);
  });

  test('a division by zero, a negative root, and an unknown operator give 0', () => {
    expect(evaluateFormula('*/ w h 0', vars)).toBe(0);
    expect(evaluateFormula('+/ w h 0', vars)).toBe(0);
    expect(evaluateFormula('sqrt -4', vars)).toBe(0);
    expect(evaluateFormula('nope w h', vars)).toBe(0);
  });

  test('an unknown guide name reads as 0', () => {
    expect(evaluateFormula('+- missing 5 0', vars)).toBe(5);
  });
});

describe('arc sweeps from guide formulas (R1)', () => {
  test('a sweep beyond one turn keeps its end point', () => {
    const quarter = custom(
      '',
      '<a:path w="100" h="100"><a:moveTo><a:pt x="100" y="50"/></a:moveTo>' +
        '<a:arcTo wR="50" hR="50" stAng="0" swAng="5400000"/></a:path>'
    );
    const fiveQuarters = custom(
      '',
      '<a:path w="100" h="100"><a:moveTo><a:pt x="100" y="50"/></a:moveTo>' +
        '<a:arcTo wR="50" hR="50" stAng="0" swAng="27000000"/></a:path>'
    );
    const end = (g: EvaluatedGeometry): PathCommand => g.paths[0].commands.at(-1)!;
    const a = end(quarter);
    const b = end(fiveQuarters);
    if (a.op !== 'C' || b.op !== 'C') throw new Error('expected arcs');
    expect(b.x).toBeCloseTo(a.x, 3);
    expect(b.y).toBeCloseTo(a.y, 3);
  });

  test('a sweep of 2^31 squared draws at most two turns of Beziers', () => {
    const g = custom(
      '<a:gd name="g1" fmla="*/ 2147483647 2147483647 1"/>',
      '<a:path w="1000000" h="1000000"><a:moveTo><a:pt x="1000000" y="500000"/></a:moveTo>' +
        '<a:arcTo wR="500000" hR="500000" stAng="0" swAng="g1"/><a:close/></a:path>'
    );
    const cmds = g.paths[0].commands;
    expect(cmds.length).toBeLessThanOrEqual(11);
    expect(cmds.every(finite)).toBe(true);
  });

  test('an overflowing guide counts as 0, never as Infinity or NaN', () => {
    const g = custom(
      '<a:gd name="big" fmla="*/ 1e308 1e308 1"/><a:gd name="nan" fmla="+- big 0 big"/>',
      '<a:path><a:moveTo><a:pt x="big" y="nan"/></a:moveTo>' +
        '<a:arcTo wR="big" hR="big" stAng="big" swAng="big"/><a:lnTo><a:pt x="nan" y="10"/></a:lnTo></a:path>'
    );
    const cmds = g.paths[0].commands;
    expect(cmds.every(finite)).toBe(true);
    expect(cmds[0]).toEqual({ op: 'M', x: 0, y: 0 });
  });

  test('the output of one geometry is capped', () => {
    const arcs = '<a:arcTo wR="10" hR="10" stAng="0" swAng="21600000"/>'.repeat(30_000);
    const g = custom('', `<a:path><a:moveTo><a:pt x="20" y="10"/></a:moveTo>${arcs}</a:path>`);
    const total = g.paths.reduce((n, p) => n + p.commands.length, 0);
    expect(total).toBeLessThanOrEqual(100_000);
    expect(total).toBeGreaterThan(90_000);
  });
});

describe('resource caps on hostile definitions', () => {
  /** A definition with these guides and one path, built directly: no XML element cap applies. */
  const def = (
    gd: readonly (readonly [string, string])[],
    cmds: readonly PresetCommand[]
  ): PresetShapeDefinition => ({ av: [], gd, paths: [{ cmds }] });

  test('guides past the 4096th are not defined, so a reference to one reads as 0', () => {
    const gd = Array.from({ length: 5000 }, (_, i) => [`g${i}`, `val ${i + 1}`] as const);
    const g = evaluateGeometry(def(gd, [['M', 'g4095', 'g4096']]), 10, 10);
    expect(g.paths[0].commands).toEqual([{ op: 'M', x: 4096, y: 0 }]);
  });

  test('path commands past the 100 000th of a path are not read', () => {
    // 100 000 lines with no current point draw nothing; the move after them is not read.
    const lines = Array.from({ length: 100_000 }, (): PresetCommand => ['L', '1', '1']);
    const g = evaluateGeometry(def([], [...lines, ['M', '5', '5'], ['L', '6', '6']]), 10, 10);
    expect(g.paths[0].commands).toEqual([]);
  });

  test('a guide that names itself or a later guide reads the value so far, or 0', () => {
    const g = evaluateGeometry(
      def(
        [
          ['a', '+- a 1 0'],
          ['a', '+- a 1 0'],
          ['b', '+- c 1 0'],
          ['c', 'val 7'],
        ],
        [['M', 'a', 'b']]
      ),
      10,
      10
    );
    // Guides are defined in order (ECMA-376 Part 1, 20.1.9.11): no cycle can form.
    expect(g.paths[0].commands).toEqual([{ op: 'M', x: 2, y: 1 }]);
  });

  test('a chain of 4096 doubling guides stays finite', () => {
    const gd = [
      ['g0', 'val 1'] as const,
      ...Array.from({ length: 4095 }, (_, i) => [`g${i + 1}`, `+- g${i} g${i} 0`] as const),
    ];
    const g = evaluateGeometry(def(gd, [['M', 'g4095', 'g1023']]), 10, 10);
    expect(g.paths[0].commands.every(finite)).toBe(true);
  });

  test('literals that are not finite numbers read as 0', () => {
    for (const lit of ['NaN', 'Infinity', '-Infinity', '1e400', '0x10z', '']) {
      const g = evaluateGeometry(def([['v', `val ${lit}`]], [['M', 'v', lit]]), 10, 10);
      expect({ lit, cmd: g.paths[0].commands[0] }).toEqual({ lit, cmd: { op: 'M', x: 0, y: 0 } });
    }
  });

  test('tan at a right angle and a zero-size shape give finite commands', () => {
    const g = evaluateGeometry(
      def(
        [['t', 'tan 1000 5400000']],
        [
          ['M', 't', 't'],
          ['L', 'w', 'h'],
        ]
      ),
      0,
      0
    );
    expect(g.paths[0].commands.every(finite)).toBe(true);
  });
});

describe('arcs in path units (R7)', () => {
  const cloud = (w: number, h: number): EvaluatedGeometry =>
    evaluatePresetGeometry('cloud', w, h, new Map());

  test('cloud at 4:1 is the square cloud scaled, so it closes as the spec data does', () => {
    const square = cloud(1_000_000, 1_000_000).paths[0].commands;
    const wide = cloud(4_000_000, 1_000_000).paths[0].commands;
    expect(wide.length).toBe(square.length);
    let worst = 0;
    for (let i = 0; i < square.length; i += 1) {
      const a = square[i];
      const b = wide[i];
      expect(b.op).toBe(a.op);
      if (a.op === 'Z' || b.op === 'Z') continue;
      worst = Math.max(worst, Math.abs(b.x - a.x * 4), Math.abs(b.y - a.y));
    }
    expect(worst).toBeLessThan(1);
    // The spec's own path data ends 139 of 43 200 path units from its start; `close` draws
    // that chord. Before arcs were evaluated in path units, the gap at 4:1 was 155 746 EMU.
    const first = wide[0];
    const last = wide[wide.length - 2];
    if (first.op !== 'M' || last.op !== 'C') throw new Error('unexpected cloud path');
    expect(Math.hypot(first.x - last.x, first.y - last.y)).toBeLessThan(5_000);
  });
});

describe('drawing commands before the first moveTo', () => {
  // The reference PDF of fixture `custom-geometry` slide 2 shows no part of such a path.
  test('a path that starts with lnTo draws nothing', () => {
    const g = custom(
      '',
      '<a:path w="100" h="100"><a:lnTo><a:pt x="100" y="0"/></a:lnTo>' +
        '<a:lnTo><a:pt x="100" y="100"/></a:lnTo>' +
        '<a:arcTo wR="50" hR="50" stAng="5400000" swAng="5400000"/><a:close/></a:path>'
    );
    expect(g.paths[0].commands).toEqual([]);
  });

  test('a path that starts with arcTo, quadBezTo, or cubicBezTo draws nothing', () => {
    const g = custom(
      '',
      '<a:path w="100" h="100"><a:arcTo wR="50" hR="50" stAng="10800000" swAng="5400000"/></a:path>' +
        '<a:path w="100" h="100"><a:quadBezTo><a:pt x="50" y="0"/><a:pt x="100" y="100"/></a:quadBezTo></a:path>' +
        '<a:path w="100" h="100"><a:cubicBezTo><a:pt x="0" y="50"/><a:pt x="50" y="0"/>' +
        '<a:pt x="100" y="100"/></a:cubicBezTo></a:path>'
    );
    expect(g.paths.map((p) => p.commands.length)).toEqual([0, 0, 0]);
  });

  test('a later moveTo starts the drawing', () => {
    const g = custom(
      '',
      '<a:path w="100" h="100"><a:lnTo><a:pt x="100" y="0"/></a:lnTo>' +
        '<a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo>' +
        '<a:close/></a:path>',
      100,
      100
    );
    expect(g.paths[0].commands).toEqual([
      { op: 'M', x: 0, y: 100 },
      { op: 'L', x: 100, y: 100 },
      { op: 'Z' },
    ]);
  });

  test('presets with an explicit move are unchanged', () => {
    const rect = evaluatePresetGeometry('rect', 10, 20, new Map());
    expect(rect.paths[0].commands.map((c) => c.op).join('')).toBe('MLLLZ');
  });
});

describe('the custGeom reader', () => {
  test('reads path sizes, fill modes, stroke flags, and the text rectangle', () => {
    const r = readXml(
      '<a:custGeom><a:gdLst><a:gd name="q" fmla="*/ w 1 4"/></a:gdLst>' +
        '<a:rect l="q" t="t" r="r" b="nope"/><a:pathLst>' +
        '<a:path w="10" h="20" fill="darken" stroke="0"><a:moveTo><a:pt x="10" y="20"/></a:moveTo></a:path>' +
        '<a:path w="abc" fill="glow"><a:moveTo><a:pt x="1" y="1"/></a:moveTo></a:path></a:pathLst></a:custGeom>'
    );
    const root = r.ok ? r.nodes.find((n): n is XmlElement => n.type === 'element') : undefined;
    if (!root) throw new Error('bad xml');
    const g = evaluateGeometry(customGeometryDefinition(neutral(root)), 400, 200);
    expect(g.paths[0]).toEqual({
      commands: [{ op: 'M', x: 400, y: 200 }],
      fill: 'darken',
      stroke: false,
    });
    // A fill mode outside ST_PathFillMode reads as `norm`.
    expect(g.paths[1].fill).toBe('norm');
    // A path size that is not an integer reads as absent: the shape's own EMU space.
    expect(g.paths[1].commands).toEqual([{ op: 'M', x: 1, y: 1 }]);
    // A text rectangle reference that is no guide reads as the shape's edge.
    expect(g.textRect).toEqual({ l: 100, t: 0, r: 400, b: 200 });
  });

  test('reads guides and points only in the DrawingML main namespace', () => {
    const g = custom(
      '<x:gd name="p" fmla="val 50"/>',
      '<a:path w="100" h="100"><a:moveTo><x:pt x="9" y="9"/><a:pt x="p" y="5"/></a:moveTo>' +
        '<x:lnTo><a:pt x="7" y="7"/></x:lnTo></a:path>',
      100,
      100
    );
    // `x:gd` is no guide, so `p` reads as 0; `x:pt` is no point. A path verb is read by its
    // local name in any namespace.
    expect(g.paths[0].commands).toEqual([
      { op: 'M', x: 0, y: 5 },
      { op: 'L', x: 7, y: 7 },
    ]);
  });
});

describe('the preset table', () => {
  test('every preset evaluates to finite commands at a non-square size', () => {
    for (const name of presetNames()) {
      const g = evaluatePresetGeometry(name, 3_000_000, 1_000_000, new Map());
      expect(g.unknownPreset).toBeUndefined();
      for (const p of g.paths) expect(p.commands.every(finite)).toBe(true);
    }
  });

  test('upArrow is downArrow mirrored vertically, with the same adjust values', () => {
    const adjust = new Map([['adj2', 30_000]]);
    const up = evaluatePresetGeometry('upArrow', 100, 200, adjust);
    const down = evaluatePresetGeometry('downArrow', 100, 200, adjust);
    const ys = (g: EvaluatedGeometry): number[] =>
      g.paths[0].commands.flatMap((c) => (c.op === 'M' || c.op === 'L' ? [c.y] : []));
    expect(ys(up)).toEqual(ys(down).map((y) => 200 - y));
    expect(up.textRect.t).toBeCloseTo(200 - down.textRect.b, 9);
  });

  test.each(['noSuchShape', '__proto__', 'constructor', 'toString'])(
    'an unknown name (%s) draws a rectangle and is reported',
    (name) => {
      const g = evaluatePresetGeometry(name, 10, 20, new Map());
      expect(g.unknownPreset).toBe(name);
      expect(g.paths[0].commands.map((c) => c.op).join('')).toBe('MLLLZ');
    }
  );

  test('an adjust value overrides the preset default, and a non-finite one reads as 0', () => {
    const at = (adj: number): EvaluatedGeometry =>
      evaluatePresetGeometry('roundRect', 1000, 1000, new Map([['adj', adj]]));
    const first = (g: EvaluatedGeometry): PathCommand => g.paths[0].commands[0];
    expect(first(at(50_000))).toEqual({ op: 'M', x: 0, y: 500 });
    expect(first(at(Number.NaN))).toEqual({ op: 'M', x: 0, y: 0 });
  });
});
