// Connector ends glue to the connection sites of shapes (`a:stCxn`, `a:endCxn`), and glued
// connectors follow the shapes they glue to. Straight, elbow, and curved connectors alike.

import { describe, expect, test } from 'bun:test';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { presetConnectionSites } from '@eigenpal/ooxml-core';
import { apply, type LayoutItem } from '../../layout/index.ts';
import { serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { findItem, lineEnds } from '../../automation/commands/placement.ts';
import { elbowPoints, placedSites } from '../../automation/commands/connectors.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { dragStep, startDrag } from '../surface-drag.ts';
import { preview } from '../drag-preview.ts';

const IN = 914_400;

/** A connector (`p:cxnSp`) with a preset, optionally glued (`a:stCxn`, `a:endCxn`). */
function connector(o: {
  id: number;
  prst: string;
  x: number;
  y: number;
  cx: number;
  cy: number;
  st?: [number, number];
  end?: [number, number];
}): string {
  const glue =
    (o.st ? `<a:stCxn id="${o.st[0]}" idx="${o.st[1]}"/>` : '') +
    (o.end ? `<a:endCxn id="${o.end[0]}" idx="${o.end[1]}"/>` : '');
  return (
    `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${o.id}" name="C${o.id}"/><p:cNvCxnSpPr>${glue}</p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${o.x}" y="${o.y}"/><a:ext cx="${o.cx}" cy="${o.cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${o.prst}"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr></p:cxnSp>`
  );
}

/**
 * Two boxes: A at (1 in, 2 in) and B at (6 in, 3 in), each 2 in × 1 in. A rect's sites are
 * top (0), left (1), bottom (2), right (3).
 */
function deck(connectors: string): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({ id: 2, name: 'A', x: IN, y: 2 * IN, cx: 2 * IN, cy: IN }) +
          shape({ id: 3, name: 'B', x: 6 * IN, y: 3 * IN, cx: 2 * IN, cy: IN }) +
          connectors,
      },
    ],
  });
}

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const items = (e: PptxEditorInstance): readonly LayoutItem[] => e.session!.layout(0)!.items;
const named = (e: PptxEditorInstance, name: string): LayoutItem =>
  items(e).find((i) => i.element.name === name)!;
const slideXml = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!);
const near = (a: { x: number; y: number }, b: { x: number; y: number }, tol = 50): boolean =>
  Math.abs(a.x - b.x) <= tol && Math.abs(a.y - b.y) <= tol;
const endsOf = (e: PptxEditorInstance, name: string) => lineEnds(named(e, name))!;
/** The corners of a connector's path on the slide, in inches, rounded to 0.01 in. */
const cornersOf = (e: PptxEditorInstance, name: string): number[][] => {
  const item = named(e, name);
  if (item.kind !== 'shape') return [];
  return item.geometry.paths[0].commands
    .filter((c) => c.op === 'M' || c.op === 'L')
    .map((c) => apply(item.matrix, (c as { x: number }).x, (c as { y: number }).y))
    .map((p) => [Math.round((p.x / IN) * 100) / 100, Math.round((p.y / IN) * 100) / 100]);
};
const presetOf = (e: PptxEditorInstance, name: string): string => {
  const g = (named(e, name).element as { geometry: { kind: string; preset?: string } }).geometry;
  return g.kind === 'preset' ? g.preset! : g.kind;
};

/** Drag one end of connector `name` to slide point `to` and release: one `setTransforms`. */
function dragEnd(
  e: PptxEditorInstance,
  name: string,
  which: 'start' | 'end',
  to: { x: number; y: number }
): ReturnType<typeof dragStep> {
  const item = named(e, name);
  const from = lineEnds(item)![which];
  const d = startDrag(e.session!.layout(0), which, [item.element.nodeId], from, 'g')!;
  const step = dragStep(d, to, {
    shift: false,
    snapTo: null,
    glueTo: 0.1 * IN,
    aspectLocked: () => false,
  });
  expect(e.exec({ type: 'setTransforms', items: step.items }).ok).toBe(true);
  return step;
}

describe('connection sites', () => {
  test('a preset lists its a:cxnLst sites with their angles', () => {
    const sites = presetConnectionSites('rect', 200, 100, new Map());
    expect(sites).toEqual([
      { x: 100, y: 0, angle: 270 },
      { x: 0, y: 50, angle: 180 },
      { x: 100, y: 100, angle: 90 },
      { x: 200, y: 50, angle: 0 },
    ]);
  });

  test('sites are placed on the slide with the shape transform', async () => {
    const e = await open(deck(''));
    const right = placedSites(named(e, 'A'))[3];
    expect(near(right, { x: 3 * IN, y: 2.5 * IN }, 1)).toBe(true);
    expect([right.dx, right.dy]).toEqual([1, 0]);
  });
});

describe('elbow routes', () => {
  const s = { x: 0, y: 0 };
  test('facing sites turn half way (three segments)', () => {
    const pts = elbowPoints(s, { x: 1, y: 0 }, { x: 1000, y: 500 }, { x: -1, y: 0 }, 100);
    expect(pts).toEqual([s, { x: 500, y: 0 }, { x: 500, y: 500 }, { x: 1000, y: 500 }]);
  });
  test('perpendicular sites turn once (two segments)', () => {
    const pts = elbowPoints(s, { x: 1, y: 0 }, { x: 1000, y: 500 }, { x: 0, y: -1 }, 100);
    expect(pts).toEqual([s, { x: 1000, y: 0 }, { x: 1000, y: 500 }]);
  });
  test('sites that face the same way go around both ends', () => {
    const pts = elbowPoints(s, { x: 0, y: -1 }, { x: 1000, y: 200 }, { x: 0, y: -1 }, 100);
    expect(pts).toEqual([s, { x: 0, y: -100 }, { x: 1000, y: -100 }, { x: 1000, y: 200 }]);
  });
  test('back to back sites run out of both and cross between them (five segments)', () => {
    const pts = elbowPoints(s, { x: 1, y: 0 }, { x: -1000, y: 600 }, { x: -1, y: 0 }, 100);
    expect(pts.length).toBe(6);
  });
});

describe('gluing a connector end', () => {
  for (const prst of ['straightConnector1', 'bentConnector3', 'curvedConnector3']) {
    test(`a ${prst} end dragged near a site glues to it`, async () => {
      const e = await open(
        deck(connector({ id: 4, prst, x: 3.5 * IN, y: 5 * IN, cx: 2 * IN, cy: IN }))
      );
      // A's right site is at (3 in, 2.5 in); B's left site at (6 in, 3.5 in).
      const step = dragEnd(e, 'C4', 'start', { x: 3.05 * IN, y: 2.52 * IN });
      expect(step.guides.some((g) => g.kind === 'site' && g.glued)).toBe(true);
      dragEnd(e, 'C4', 'end', { x: 5.96 * IN, y: 3.47 * IN });
      const xml = slideXml(e);
      expect(xml).toContain('<a:stCxn id="2" idx="3"/>');
      expect(xml).toContain('<a:endCxn id="3" idx="1"/>');
      const ends = endsOf(e, 'C4');
      expect(near(ends.start, { x: 3 * IN, y: 2.5 * IN })).toBe(true);
      expect(near(ends.end, { x: 6 * IN, y: 3.5 * IN })).toBe(true);
      // An elbow from a right site to a left site turns half way.
      if (prst !== 'straightConnector1') expect(presetOf(e, 'C4')).toBe(prst);
    });
  }

  test('an end dragged away from every site lets go', async () => {
    const e = await open(
      deck(
        connector({
          id: 4,
          prst: 'bentConnector3',
          x: 3 * IN,
          y: 2.5 * IN,
          cx: 3 * IN,
          cy: IN,
          st: [2, 3],
          end: [3, 1],
        })
      )
    );
    dragEnd(e, 'C4', 'end', { x: 9 * IN, y: 6 * IN });
    const xml = slideXml(e);
    expect(xml).toContain('<a:stCxn id="2" idx="3"/>');
    expect(xml).not.toContain('endCxn');
  });

  test('an elbow glued to a top site leaves upward and picks its preset', async () => {
    const e = await open(
      deck(connector({ id: 4, prst: 'bentConnector3', x: 3.5 * IN, y: 5 * IN, cx: 2 * IN, cy: IN }))
    );
    dragEnd(e, 'C4', 'start', { x: 3 * IN, y: 2.5 * IN });
    // B's top site (7 in, 3 in): right out of A, then down into B's top.
    dragEnd(e, 'C4', 'end', { x: 7 * IN, y: 3 * IN });
    expect(slideXml(e)).toContain('<a:endCxn id="3" idx="0"/>');
    expect(presetOf(e, 'C4')).toBe('bentConnector2');
    expect(cornersOf(e, 'C4')).toEqual([
      [3, 2.5],
      [7, 2.5],
      [7, 3],
    ]);
    const ends = endsOf(e, 'C4');
    expect(near(ends.start, { x: 3 * IN, y: 2.5 * IN })).toBe(true);
    expect(near(ends.end, { x: 7 * IN, y: 3 * IN })).toBe(true);
  });
});

describe('elbow connectors route from site to site', () => {
  test('facing sites: the route turns half way', async () => {
    const e = await open(
      deck(
        connector({
          id: 4,
          prst: 'bentConnector3',
          x: 0,
          y: 0,
          cx: IN,
          cy: IN,
          st: [2, 3],
          end: [3, 1],
        })
      )
    );
    const a = named(e, 'A');
    // Any move of A routes the connector again.
    const x = a.element.xfrm;
    const same = { nodeId: a.element.nodeId, x: x.x, y: x.y + 1, cx: x.cx, cy: x.cy, rot: 0 };
    expect(e.exec({ type: 'setTransforms', items: [same] }).ok).toBe(true);
    expect(cornersOf(e, 'C4')).toEqual([
      [3, 2.5],
      [4.5, 2.5],
      [4.5, 3.5],
      [6, 3.5],
    ]);
  });

  test('a bottom site leaves downward: the box turns 90°', async () => {
    const e = await open(
      deck(
        connector({
          id: 4,
          prst: 'bentConnector3',
          x: 0,
          y: 0,
          cx: IN,
          cy: IN,
          st: [2, 2],
          end: [3, 1],
        })
      )
    );
    const a = named(e, 'A');
    const x = a.element.xfrm;
    const same = { nodeId: a.element.nodeId, x: x.x, y: x.y + 1, cx: x.cx, cy: x.cy, rot: 0 };
    expect(e.exec({ type: 'setTransforms', items: [same] }).ok).toBe(true);
    expect(presetOf(e, 'C4')).toBe('bentConnector2');
    expect(cornersOf(e, 'C4')).toEqual([
      [2, 3],
      [2, 3.5],
      [6, 3.5],
    ]);
    expect(slideXml(e)).toContain('rot="5400000"');
  });

  test('a curve routes as an elbow does, with the curved preset', async () => {
    const e = await open(
      deck(
        connector({
          id: 4,
          prst: 'curvedConnector3',
          x: 0,
          y: 0,
          cx: IN,
          cy: IN,
          st: [2, 0],
          end: [3, 0],
        })
      )
    );
    const a = named(e, 'A');
    const x = a.element.xfrm;
    const same = { nodeId: a.element.nodeId, x: x.x, y: x.y + 1, cx: x.cx, cy: x.cy, rot: 0 };
    expect(e.exec({ type: 'setTransforms', items: [same] }).ok).toBe(true);
    // Both top sites face up: around both ends, three segments.
    expect(presetOf(e, 'C4')).toBe('curvedConnector3');
    const ends = endsOf(e, 'C4');
    expect(near(ends.start, { x: 2 * IN, y: 2 * IN })).toBe(true);
    expect(near(ends.end, { x: 7 * IN, y: 3 * IN })).toBe(true);
  });
});

describe('glued connectors follow their shapes', () => {
  const glued = (prst: string): string =>
    connector({ id: 4, prst, x: 3 * IN, y: 2.5 * IN, cx: 3 * IN, cy: IN, st: [2, 3], end: [3, 1] });

  for (const prst of ['straightConnector1', 'bentConnector3', 'curvedConnector3']) {
    test(`a ${prst} reroutes when a glued shape moves`, async () => {
      const e = await open(deck(glued(prst)));
      const b = named(e, 'B');
      const x = b.element.xfrm;
      // B moves down 2 in and right 1 in.
      const moved = {
        nodeId: b.element.nodeId,
        x: x.x + IN,
        y: x.y + 2 * IN,
        cx: x.cx,
        cy: x.cy,
        rot: 0,
      };
      expect(e.exec({ type: 'setTransforms', items: [moved] }).ok).toBe(true);
      const ends = endsOf(e, 'C4');
      expect(near(ends.start, { x: 3 * IN, y: 2.5 * IN })).toBe(true);
      expect(near(ends.end, { x: 7 * IN, y: 5.5 * IN })).toBe(true);
      // The glue stays, and one undo puts the shape and the connector back.
      expect(slideXml(e)).toContain('<a:endCxn id="3" idx="1"/>');
      e.exec({ type: 'undo' });
      expect(near(endsOf(e, 'C4').end, { x: 6 * IN, y: 3.5 * IN })).toBe(true);
    });
  }

  test('an elbow re-picks its preset when the shapes no longer face each other', async () => {
    const e = await open(deck(glued('bentConnector3')));
    const b = named(e, 'B');
    const x = b.element.xfrm;
    // B moves to the left of A: the right site of A and the left site of B are back to back.
    const moved = { nodeId: b.element.nodeId, x: 0, y: 5 * IN, cx: x.cx, cy: x.cy, rot: 0 };
    expect(e.exec({ type: 'setTransforms', items: [moved] }).ok).toBe(true);
    expect(presetOf(e, 'C4')).toBe('bentConnector5');
    const ends = endsOf(e, 'C4');
    expect(near(ends.start, { x: 3 * IN, y: 2.5 * IN })).toBe(true);
    expect(near(ends.end, { x: 0, y: 5.5 * IN })).toBe(true);
  });

  test('a resize moves the site the connector is glued to', async () => {
    const e = await open(deck(glued('bentConnector3')));
    const a = named(e, 'A');
    const x = a.element.xfrm;
    const resized = {
      nodeId: a.element.nodeId,
      x: x.x,
      y: x.y,
      cx: x.cx + IN,
      cy: x.cy + IN,
      rot: 0,
    };
    expect(e.exec({ type: 'setTransforms', items: [resized] }).ok).toBe(true);
    expect(near(endsOf(e, 'C4').start, { x: 4 * IN, y: 3 * IN })).toBe(true);
  });

  test('a connector moved on its own lets go of its ends', async () => {
    const e = await open(deck(glued('bentConnector3')));
    e.setSelection({ kind: 'shapes', nodeIds: [named(e, 'C4').element.nodeId] });
    expect(e.exec({ type: 'nudge', dx: IN, dy: 0 }).ok).toBe(true);
    expect(slideXml(e)).not.toContain('Cxn id=');
  });

  test('the drag preview shows the connector at its new route', async () => {
    const e = await open(deck(glued('bentConnector3')));
    const b = named(e, 'B');
    const layout = e.session!.layout(0)!;
    const d = startDrag(layout, 'move', [b.element.nodeId], { x: 7 * IN, y: 3.5 * IN }, 'g')!;
    const step = dragStep(
      d,
      { x: 8 * IN, y: 5.5 * IN },
      {
        shift: false,
        snapTo: null,
        aspectLocked: () => false,
      }
    );
    const { relayoutItem } = await import('../../layout/item-layout.ts');
    preview(d, step.items, (item, xfrm) => relayoutItem(item, xfrm, e.session!.measurer));
    const c = d.shown.get(named(e, 'C4').element.nodeId);
    expect(c).toBeDefined();
    expect(near(lineEnds(c!)!.end, { x: 7 * IN, y: 5.5 * IN })).toBe(true);
    // The preview writes nothing.
    expect(findItem(e.session!.layout(0)!.items, named(e, 'C4').element.nodeId)).toBe(
      named(e, 'C4')
    );
  });
});
