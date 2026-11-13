// The dash and line end rules that the SVG paint and the PDF export share. The expected values
// are measured on the reference PDFs of the `line-ends-dashes` and `picture-adjustments` fixtures.

import { describe, expect, test } from 'bun:test';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx, writePptx } from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import type { LineEnd, ResolvedLine } from '../../cascade/index.ts';
import {
  createFixedMeasurer,
  layoutSlide,
  dashArray,
  lineEndReach,
  pictureBorder,
  unmeasuredDashStart,
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

/** A closed 20 × 10 pt rectangle, clockwise from the top-left. */
const SQUARE: PathCommand[] = [
  { op: 'M', x: 0, y: 0 },
  { op: 'L', x: 20 * PT, y: 0 },
  { op: 'L', x: 20 * PT, y: 10 * PT },
  { op: 'L', x: 0, y: 10 * PT },
  { op: 'Z' },
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

  test('a closed path with round or square caps starts its dashes at the closing segment', () => {
    const rnd = line({ dash: 'lgDashDot', cap: 'rnd' });
    const plan = planStroke(SQUARE, rnd);
    // Bottom-left first, then the closing segment up to the top-left, then the path as it was.
    expect(points(plan.commands)).toEqual([
      [0, 10],
      [0, 0],
      [20, 0],
      [20, 10],
      [0, 10],
    ]);
    expect(plan.commands[plan.commands.length - 1]).toEqual({ op: 'Z' });
    // Round dashes are [7, 4, 0, 4] widths: the pattern starts one width before its end.
    expect(plan.dashPhase).toBeCloseTo(14 * PT, 1);
    const sq = planStroke(SQUARE, line({ dash: 'lgDashDot', cap: 'sq' }));
    expect(points(sq.commands)[0]).toEqual([0, 10]);
    expect(sq.dashPhase).toBeUndefined();
  });

  test('a closed path that ends at its first point starts at the last chord', () => {
    // A closing segment shorter than one EMU is no segment: the last line is the last chord.
    const loop: PathCommand[] = [...SQUARE.slice(0, 4), { op: 'L', x: 0.4, y: 0.2 }, { op: 'Z' }];
    const rnd = line({ dash: 'sysDot', cap: 'rnd' });
    const plan = planStroke(loop, rnd);
    expect(points(plan.commands).slice(0, 3)).toEqual([
      [0, 10],
      [0, 0],
      [20, 0],
    ]);
    expect(plan.dashPhase).toBeCloseTo(PT, 1);
    expect(unmeasuredDashStart(loop, rnd)).toBe(true);
    expect(unmeasuredDashStart(loop, line({ dash: 'sysDot' }))).toBe(false);
    expect(unmeasuredDashStart(SQUARE, rnd)).toBe(false);
  });

  test('a closed curve starts one chord before its end: 16 or 32 steps a quarter circle', () => {
    // A circle of four cubic quarter arcs, from the left point, clockwise (y down).
    const circle = (r: number): PathCommand[] => {
      const k = 0.5522847498 * r;
      return [
        { op: 'M', x: -r, y: 0 },
        { op: 'C', x1: -r, y1: -k, x2: -k, y2: -r, x: 0, y: -r },
        { op: 'C', x1: k, y1: -r, x2: r, y2: -k, x: r, y: 0 },
        { op: 'C', x1: r, y1: k, x2: k, y2: r, x: 0, y: r },
        { op: 'C', x1: -k, y1: r, x2: -r, y2: k, x: -r, y: 0 },
        { op: 'Z' },
      ];
    };
    const sq = line({ dash: 'dash', cap: 'sq' });
    // The angle before the left point where the dashes start, in degrees.
    const before = (r: number): number => {
      const start = planStroke(circle(r), sq).commands[0]!;
      if (start.op !== 'M') throw new Error('no move');
      return (Math.atan2(start.y, -start.x) * 180) / Math.PI;
    };
    // Reference PDFs: about 6° on a 0.8 in circle, about 2.8° on 1.7 in and 3 in circles.
    expect(before(0.4 * 914_400)).toBeCloseTo(90 / 16, 0);
    expect(before(0.85 * 914_400)).toBeCloseTo(90 / 32, 0);
    expect(before(1.5 * 914_400)).toBeCloseTo(90 / 32, 0);
  });

  test('a closed path with flat caps, a solid line, or a custom dash keeps its start', () => {
    for (const l of [
      line({ dash: 'lgDashDot' }),
      line({ cap: 'rnd' }),
      line({ dash: 'custom', dashes: [3, 1], cap: 'rnd' }),
    ]) {
      const plan = planStroke(SQUARE, l);
      expect(plan.commands).toBe(SQUARE);
      expect(plan.dashPhase).toBeUndefined();
    }
  });

  test("a rectangle picture's border is the outline grown by the line width", () => {
    const border = pictureBorder(
      [SQUARE],
      true,
      line({ dash: 'sysDot', cap: 'rnd' }),
      20 * PT,
      10 * PT
    );
    expect(border.kind).toBe('grown');
    if (border.kind !== 'grown') return;
    expect(border.clip).toBe(false);
    // Half a width out on each side, from the closing segment, as a shape's outline.
    expect(points(border.paths[0]!.commands)).toEqual([
      [-0.5, 10.5],
      [-0.5, -0.5],
      [20.5, -0.5],
      [20.5, 10.5],
      [-0.5, 10.5],
    ]);
    expect(border.paths[0]!.dashPhase).toBeCloseTo(PT, 1);
    // A solid rectangle border is grown too: its joins are the joins of the grown outline.
    expect(pictureBorder([SQUARE], true, line(), 20 * PT, 10 * PT).kind).toBe('grown');
  });

  test('a border off a rectangle is clipped outside; a solid one is a doubled stroke', () => {
    const dashed = pictureBorder([SQUARE], false, line({ dash: 'dash' }), 20 * PT, 10 * PT);
    expect(dashed.kind === 'grown' && dashed.clip).toBe(true);
    expect(pictureBorder([SQUARE], false, line(), 20 * PT, 10 * PT)).toEqual({
      kind: 'doubled',
      paths: [SQUARE],
    });
    // An empty box cannot grow.
    expect(pictureBorder([SQUARE], true, line({ dash: 'dash' }), 0, 10 * PT).kind).toBe('doubled');
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

describe('native dash expansion limits', () => {
  function deck(width: number, preset = 'line') {
    const opened = readPptx(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              x: 0,
              y: 0,
              cx: width,
              cy: preset === 'line' ? 0 : width,
              prst: preset,
              line: '<a:ln w="1"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="sysDot"/></a:ln>',
            }),
          },
        ],
      })
    );
    if (!opened.ok) throw new Error('open failed');
    return opened.pkg;
  }

  test('keeps tiny dashes on short paths and at the expansion bound', () => {
    for (const width of [100, 199_998]) {
      const layout = layoutSlide(resolveSlide(deck(width), 0)!, createFixedMeasurer());
      const item = layout.items[0];
      if (item.kind !== 'shape') throw new Error('no shape');
      expect(item.element.line?.dash).toBe('sysDot');
      expect(layout.diagnostics).not.toContain('approximation: dash-complexity');
    }
  });

  test('draws over-budget straight and curved outlines solid without changing saved properties', () => {
    for (const [width, preset] of [
      [200_000, 'line'],
      [2147483647, 'ellipse'],
    ] as const) {
      const pkg = deck(width, preset);
      const resolved = resolveSlide(pkg, 0)!;
      const layout = layoutSlide(resolved, createFixedMeasurer());
      const item = layout.items[0];
      if (item.kind !== 'shape') throw new Error('no shape');
      expect(item.paintLine?.dash).toBe('solid');
      expect(item.element.line?.dash).toBe('sysDot');
      expect(layout.diagnostics).toContain('approximation: dash-complexity');
      const source = resolved.elements[0];
      if (source.kind !== 'shape') throw new Error('no source shape');
      expect(source.line?.dash).toBe('sysDot');
      const reopened = readPptx(writePptx(pkg));
      if (!reopened.ok) throw new Error('reopen failed');
      const kept = resolveSlide(reopened.pkg, 0)!.elements[0];
      if (kept.kind !== 'shape') throw new Error('no saved shape');
      expect(kept.line?.dash).toBe('sysDot');
    }
  });
});

test('picture outlines bound even custom paths marked without a stroke', () => {
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          '<p:pic><p:nvPicPr><p:cNvPr id="2" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/>' +
          '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>' +
          '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>' +
          '<a:pathLst><a:path w="1000000" h="1000000" stroke="0"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
          '<a:lnTo><a:pt x="1000000" y="0"/></a:lnTo><a:lnTo><a:pt x="1000000" y="1000000"/></a:lnTo>' +
          '<a:close/></a:path></a:pathLst></a:custGeom>' +
          '<a:ln w="1"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="sysDot"/></a:ln>' +
          '</p:spPr></p:pic>',
      },
    ],
  });
  const opened = readPptx(bytes);
  if (!opened.ok) throw new Error('open failed');
  const layout = layoutSlide(resolveSlide(opened.pkg, 0)!, createFixedMeasurer());
  const picture = layout.items[0];
  if (picture.kind !== 'picture') throw new Error('no picture');
  expect(picture.geometry.paths[0].stroke).toBe(false);
  expect(picture.element.line?.dash).toBe('sysDot');
  expect(picture.paintLine?.dash).toBe('solid');
});
