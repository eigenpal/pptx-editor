// The dash and line end rules that the SVG paint and the PDF export share. The expected values
// are measured on the reference PDF of the `line-ends-dashes` fixture.

import { describe, expect, test } from 'bun:test';
import type { LineEnd, ResolvedLine } from '../../cascade/index.ts';
import {
  dashArray,
  lineEndReach,
  outsideDashArray,
  planStroke,
  strokeCap,
  type PathCommand,
} from '../index.ts';

const PT = 12_700;

function line(o: Partial<ResolvedLine> = {}): ResolvedLine {
  return {
    fill: { type: 'solid', color: { r: 0, g: 0, b: 0, a: 1 } },
    width: PT,
    cap: 'flat',
    join: 'round',
    miterLimit: 8,
    dash: 'solid',
    compound: 'sng',
    ...o,
  } as ResolvedLine;
}

const end = (type: string, w: LineEnd['w'] = 'med', len: LineEnd['len'] = 'med'): LineEnd => ({
  type,
  w,
  len,
});

/** A horizontal path from 0 to 100 pt. */
const STRAIGHT: PathCommand[] = [
  { op: 'M', x: 0, y: 0 },
  { op: 'L', x: 100 * PT, y: 0 },
];

/** The points of a drawn end's path, in points, rounded to 0.001 pt. */
function points(commands: readonly PathCommand[]): [number, number][] {
  const r = (v: number): number => Math.round((v / PT) * 1000) / 1000 + 0;
  return commands.flatMap((c) => (c.op === 'Z' ? [] : [[r(c.x), r(c.y)] as [number, number]]));
}

describe('dashes', () => {
  test('a flat dash is the preset in line widths', () => {
    expect(dashArray(line({ dash: 'dashDot' }), 2)).toEqual([8, 6, 2, 6]);
    expect(strokeCap(line({ dash: 'dashDot' }))).toBe('butt');
    expect(dashArray(line({ dash: 'solid' }), 2)).toBeUndefined();
    // A file value that is not a preset draws solid, also a prototype name.
    expect(dashArray(line({ dash: 'constructor' }), 2)).toBeUndefined();
  });

  test('round caps take one width from each dash and give it to each gap', () => {
    const round = line({ dash: 'custom', dashes: [3, 1.5, 1, 1.5], cap: 'rnd' });
    expect(dashArray(round, 2)).toEqual([4, 5, 0.001, 5]);
    expect(strokeCap(round)).toBe('round');
  });

  test('square caps show only at the ends of a preset dash, and on every custom dash', () => {
    expect(strokeCap(line({ dash: 'sysDot', cap: 'sq' }))).toBe('butt');
    expect(dashArray(line({ dash: 'sysDot', cap: 'sq' }), 2)).toEqual([2, 2]);
    const custom = line({ dash: 'custom', dashes: [6, 2], cap: 'sq' });
    expect(strokeCap(custom)).toBe('square');
    expect(dashArray(custom, 2)).toEqual([12, 4]);
    expect(strokeCap(line({ cap: 'sq' }))).toBe('square');
  });

  test("a picture border's dashes get butt caps and keep a centered line's visible length", () => {
    expect(outsideDashArray(line({ dash: 'dash', cap: 'rnd' }), 2)).toEqual([8, 6]);
    expect(outsideDashArray(line({ dash: 'dash' }), 2)).toEqual([8, 6]);
    expect(outsideDashArray(line(), 2)).toBeUndefined();
  });

  test('a square-capped preset dash gets a cap at its start, and at its end inside a dash', () => {
    // `dot` at 1 pt: 1 pt on, 3 pt off. 100 pt ends at the start of a dot; 102 pt in a gap.
    const dotted = line({ dash: 'dot', cap: 'sq' });
    expect(planStroke(STRAIGHT, dotted).caps.map(points)).toEqual([
      [
        [0, 0],
        [-0.5, 0],
      ],
      [
        [100, 0],
        [100.5, 0],
      ],
    ]);
    const gap: PathCommand[] = [STRAIGHT[0], { op: 'L', x: 102 * PT, y: 0 }];
    expect(planStroke(gap, dotted).caps.length).toBe(1);
    expect(planStroke(STRAIGHT, line({ dash: 'dot' })).caps).toEqual([]);
  });

  test('each subpath starts its dash pattern again, so each gets its own caps', () => {
    const two: PathCommand[] = [
      ...STRAIGHT,
      { op: 'M', x: 0, y: 10 * PT },
      { op: 'L', x: 100 * PT, y: 10 * PT },
    ];
    expect(planStroke(two, line({ dash: 'dot', cap: 'sq' })).caps.length).toBe(4);
  });
});

describe('line ends', () => {
  test('a line thinner than 2 pt takes the heads of a 2 pt line', () => {
    const plan = planStroke(
      STRAIGHT,
      line({ width: 0.75 * PT, tail: end('triangle', 'sm', 'lg') })
    );
    // Tip on the end point, 10 pt long and 4 pt wide; the line stops 1 pt inside the base.
    expect(points(plan.ends[0].commands)).toEqual([
      [100, 0],
      [90, 2],
      [90, -2],
    ]);
    expect(points(plan.commands)).toEqual([
      [0, 0],
      [91, 0],
    ]);
  });

  test('a stealth arrow has its notch 1, 2, or 3 units from the tip, and the line stops there', () => {
    const plan = planStroke(STRAIGHT, line({ width: 4.5 * PT, head: end('stealth', 'med', 'lg') }));
    expect(points(plan.ends[0].commands)).toEqual([
      [0, 0],
      [22.5, -6.75],
      [13.5, 0],
      [22.5, 6.75],
    ]);
    expect(points(plan.commands)[0]).toEqual([13.5, 0]);
  });

  test('an open arrow is a stroke as wide as the line, with its mitered tip on the end point', () => {
    const plan = planStroke(STRAIGHT, line({ width: 2.25 * PT, tail: end('arrow', 'sm', 'sm') }));
    const arrow = plan.ends[0];
    expect(arrow.kind).toBe('stroke');
    if (arrow.kind !== 'stroke') return;
    expect(arrow.width).toBe(2.25 * PT);
    // The vertex lies half a width over sin(32°) behind the tip; the arms reach 2 and 1.25 units.
    const vertex = 100 - 1.125 / Math.sin(Math.atan2(1.25, 2));
    const [a, v, b] = points(arrow.commands);
    expect(v[0]).toBeCloseTo(vertex, 2);
    expect(a[0]).toBeCloseTo(vertex - 4.5, 2);
    expect(Math.abs(a[1])).toBeCloseTo(2.8125, 2);
    expect(b[1]).toBeCloseTo(-a[1], 2);
    expect(points(plan.commands)[1][0]).toBeCloseTo(vertex, 2);
  });

  test('ovals and diamonds center on the end point, and the line runs to it', () => {
    const plan = planStroke(
      STRAIGHT,
      line({ width: 2 * PT, head: end('diamond', 'lg', 'sm'), tail: end('oval', 'sm', 'lg') })
    );
    expect(points(plan.ends[0].commands).sort()).toEqual([
      [-2, 0],
      [0, -5],
      [0, 5],
      [2, 0],
    ]);
    const oval = points(plan.ends[1].commands);
    expect(Math.min(...oval.map((p) => p[0]))).toBeCloseTo(95, 6);
    expect(Math.max(...oval.map((p) => p[1]))).toBeCloseTo(2, 6);
    expect(plan.commands).toBe(STRAIGHT);
  });

  test('a head on a curve points along the chord from where the line stops', () => {
    const curve: PathCommand[] = [
      { op: 'M', x: 0, y: 0 },
      { op: 'C', x1: 50 * PT, y1: 0, x2: 100 * PT, y2: 50 * PT, x: 100 * PT, y: 100 * PT },
    ];
    const plan = planStroke(curve, line({ width: 2 * PT, tail: end('triangle', 'med', 'lg') }));
    const [stop] = points(plan.commands).slice(-1);
    // The line stops 9 pt from the end point, and the triangle's axis runs through that point.
    expect(Math.hypot(100 - stop[0], 100 - stop[1])).toBeCloseTo(9, 3);
    const [tip, c1, c2] = points(plan.ends[0].commands);
    const mid = [(c1[0] + c2[0]) / 2, (c1[1] + c2[1]) / 2];
    const cross = (mid[0] - tip[0]) * (stop[1] - tip[1]) - (mid[1] - tip[1]) * (stop[0] - tip[0]);
    expect(Math.abs(cross)).toBeLessThan(0.01);
    expect(plan.commands.at(-1)?.op).toBe('C');
  });

  test('a closed path and an unknown type draw no line end', () => {
    const closed: PathCommand[] = [...STRAIGHT, { op: 'L', x: 0, y: 50 * PT }, { op: 'Z' }];
    expect(planStroke(closed, line({ tail: end('triangle') })).ends).toEqual([]);
    const unknown = planStroke(STRAIGHT, line({ tail: end('star') }));
    expect(unknown.ends).toEqual([]);
    expect(unknown.commands).toBe(STRAIGHT);
  });

  test('a lone move at an end of the path does not take the line end', () => {
    const trailing: PathCommand[] = [...STRAIGHT, { op: 'M', x: 50 * PT, y: 50 * PT }];
    const leading: PathCommand[] = [{ op: 'M', x: 50 * PT, y: 50 * PT }, ...STRAIGHT];
    for (const commands of [trailing, leading]) {
      const plan = planStroke(commands, line({ head: end('triangle'), tail: end('triangle') }));
      expect(plan.ends.length).toBe(2);
      // Both ends of the drawn subpath stop 2.5 units (5 pt) inside the heads.
      expect(points(plan.commands)).toEqual([
        [5, 0],
        [95, 0],
      ]);
    }
  });

  test('a line shorter than its line end draws only the end', () => {
    const short: PathCommand[] = [STRAIGHT[0], { op: 'L', x: 2 * PT, y: 0 }];
    const plan = planStroke(short, line({ tail: end('triangle') }));
    expect(plan.ends.length).toBe(1);
    expect(plan.commands.some((c) => c.op === 'L')).toBe(false);
    // With a head too, both ends draw, and the tail still points along the line.
    const both = planStroke(short, line({ head: end('triangle'), tail: end('triangle') }));
    expect(both.ends.length).toBe(2);
    expect(points(both.ends[1].commands)[0]).toEqual([2, 0]);
  });

  test('the reach bounds every point of every line end, also on a line shorter than its head', () => {
    expect(lineEndReach(line())).toBe(0);
    const short: PathCommand[] = [STRAIGHT[0], { op: 'L', x: 2 * PT, y: 0 }];
    const tip = short[1] as { x: number; y: number };
    const sizes = ['sm', 'med', 'lg'] as const;
    for (const type of ['triangle', 'stealth', 'diamond', 'oval', 'arrow'])
      for (const w of sizes)
        for (const len of sizes)
          for (const width of [PT / 4, PT, 6 * PT]) {
            const l = line({ width, tail: end(type, w, len) });
            const reach = lineEndReach(l);
            const [shape] = planStroke(short, l).ends;
            const pad = shape.kind === 'stroke' ? shape.width / 2 : 0;
            for (const c of shape.commands)
              if (c.op !== 'Z')
                expect(Math.hypot(c.x - tip.x, c.y - tip.y) + pad).toBeLessThanOrEqual(
                  reach + 1e-6
                );
          }
    // A lg triangle reaches its base corners, more than half its length from the tip.
    const lg = lineEndReach(line({ tail: end('triangle', 'lg', 'lg') }));
    expect(lg).toBeCloseTo(Math.hypot(5, 2.5) * 25_400, 3);
  });
});
