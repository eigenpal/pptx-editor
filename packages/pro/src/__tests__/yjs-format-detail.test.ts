import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  formatLineState,
  type PptxEditorInstance,
  type ShapeFormatPatch,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, shape } from '../../../../scripts/fixtures/builder.ts';
import { fingerprint, link, replica } from './convergence-harness.ts';

const deck = () =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          prst: 'line',
          x: 0,
          y: 0,
          cx: 914400,
          cy: 914400,
          line: '<a:ln w="39687"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:headEnd type="triangle" w="lg" len="sm"/><a:tailEnd type="oval" w="sm" len="lg"/></a:ln>',
          fill: '<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="accent1"><a:alpha val="80000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>',
          effects:
            '<a:effectLst><a:outerShdw blurRad="12700" dist="25400" dir="0"><a:schemeClr val="accent2"><a:alpha val="50000"/></a:schemeClr></a:outerShdw></a:effectLst>',
        }),
      },
    ],
  });
function read(e: PptxEditorInstance) {
  const id = e.session!.layout(0)!.items[0]!.element.nodeId;
  const result = e.shapeFormat(id);
  if (!result.ok) throw new Error(result.reason);
  return result.format;
}
function patch(e: PptxEditorInstance, value: ShapeFormatPatch) {
  return e.exec({ type: 'patchShapeFormat', target: read(e).target, patch: value });
}

test('concurrent endpoint type edits preserve both ends and undo independently', async () => {
  const docA = new Y.Doc(),
    docB = new Y.Doc();
  const sync = link(docA, docB);
  const bytes = deck();
  const a = await replica(docA, 'a', bytes),
    b = await replica(docB, 'b', bytes);
  try {
    a.setSelection({ kind: 'shapes', nodeIds: [read(a).target.nodeId] });
    b.setSelection({ kind: 'shapes', nodeIds: [read(b).target.nodeId] });
    sync.pause();
    expect(a.exec({ type: 'setLineEnds', head: 'diamond' }).ok).toBe(true);
    expect(b.exec({ type: 'setLineEnds', tail: 'stealth' }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(formatLineState(a)).toMatchObject({ head: 'diamond', tail: 'stealth' });
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(formatLineState(b)).toMatchObject({ head: 'triangle', tail: 'stealth' });
    expect(fingerprint(a)).toBe(fingerprint(b));
  } finally {
    a.destroy();
    b.destroy();
    docA.destroy();
    docB.destroy();
  }
});

test('concurrent shadow fields and gradient stop fields retain both edits and undo independently', async () => {
  const docA = new Y.Doc(),
    docB = new Y.Doc();
  const sync = link(docA, docB);
  const bytes = deck();
  const a = await replica(docA, 'a', bytes),
    b = await replica(docB, 'b', bytes);
  try {
    sync.pause();
    expect(patch(a, { kind: 'shadow', blur: 6 }).ok).toBe(true);
    expect(patch(b, { kind: 'shadow', distance: 9 }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(read(a).shadow).toMatchObject({ blur: 6, distance: 9 });
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(read(b).shadow).toMatchObject({ blur: 1, distance: 9 });
    sync.pause();
    expect(patch(a, { kind: 'gradient', angle: 90 }).ok).toBe(true);
    expect(
      patch(b, { kind: 'gradient', stop: { id: read(b).gradient!.stops[0]!.id, opacity: 0.3 } }).ok
    ).toBe(true);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(read(a).gradient?.angle).toBe(90);
    expect(read(a).gradient?.stops[0]?.opacity).toBe(0.3);
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(read(b).gradient?.angle).toBe(0);
    expect(read(b).gradient?.stops[0]?.opacity).toBe(0.3);
  } finally {
    a.destroy();
    b.destroy();
    docA.destroy();
    docB.destroy();
  }
});

test('a deleted gradient stop refuses a stale patch without creating another stop', async () => {
  const docA = new Y.Doc(),
    docB = new Y.Doc();
  const sync = link(docA, docB);
  const bytes = deck();
  const a = await replica(docA, 'a', bytes),
    b = await replica(docB, 'b', bytes);
  try {
    expect(patch(a, { kind: 'gradient', addStop: { position: 0.5, color: '#FF0000' } }).ok).toBe(
      true
    );
    const stale = read(b).gradient!.stops[1]!.id;
    expect(patch(a, { kind: 'gradient', removeStop: read(a).gradient!.stops[1]!.id }).ok).toBe(
      true
    );
    const before = fingerprint(b);
    expect(patch(b, { kind: 'gradient', stop: { id: stale, opacity: 0.1 } }).ok).toBe(false);
    sync.flush();
    expect(fingerprint(b)).toBe(before);
    expect(read(b).gradient?.stops).toHaveLength(2);
  } finally {
    a.destroy();
    b.destroy();
    docA.destroy();
    docB.destroy();
  }
});

test('concurrent fill transparency and fill color edits converge and undo independently', async () => {
  const docA = new Y.Doc(),
    docB = new Y.Doc();
  const sync = link(docA, docB);
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: 0,
          y: 0,
          cx: 914400,
          cy: 914400,
          fill: '<a:solidFill><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:solidFill>',
        }),
      },
    ],
  });
  const a = await replica(docA, 'a', bytes),
    b = await replica(docB, 'b', bytes);
  try {
    a.setSelection({ kind: 'shapes', nodeIds: [read(a).target.nodeId] });
    b.setSelection({ kind: 'shapes', nodeIds: [read(b).target.nodeId] });
    sync.pause();
    expect(patch(a, { kind: 'solid', opacity: 0.3 }).ok).toBe(true);
    expect(patch(b, { kind: 'solid', opacity: 0.8 }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(read(a).solid).toEqual(read(b).solid);
    sync.pause();
    expect(patch(a, { kind: 'solid', opacity: 0.6 }).ok).toBe(true);
    expect(b.exec({ type: 'setShapeFill', color: { theme: 'accent2' } }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(read(a).solid).toEqual(read(b).solid);
    expect(read(b).solid?.color).toEqual({ theme: 'accent2' });
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(read(a).solid?.color).toEqual({ theme: 'accent1' });
  } finally {
    a.destroy();
    b.destroy();
    docA.destroy();
    docB.destroy();
  }
});
