import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  readPptx,
  serializeXmlPart,
  xmlPartOf,
} from '@pptx-editor.dev/core/store';
import { buildPptx, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

// The G4 shape commands (effects, styles, flips, turns, the aspect lock, hide, and rename)
// through two editors on linked Y.Docs, as AGENTS.md "Collaboration verification" asks.

const IN = 914_400;
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({ id: 2, name: 'A', x: IN, y: IN, cx: 2 * IN, cy: IN, fill: solid('4472C4') }) +
          shape({ id: 3, name: 'B', x: 4 * IN, y: IN, cx: 2 * IN, cy: IN, fill: solid('ED7D31') }),
      },
    ],
  });

/** Two docs and a sync link that can pause, to make edits concurrent. */
function link(a: Y.Doc, b: Y.Doc) {
  const queue: [Y.Doc, Uint8Array][] = [];
  let paused = false;
  const send = (to: Y.Doc) => (update: Uint8Array, origin: unknown) => {
    if (origin === 'link') return;
    if (paused) queue.push([to, update]);
    else Y.applyUpdate(to, update, 'link');
  };
  a.on('update', send(b));
  b.on('update', send(a));
  return {
    pause: () => (paused = true),
    flush: () => {
      paused = false;
      for (const [to, update] of queue.splice(0)) Y.applyUpdate(to, update, 'link');
    },
  };
}

async function replica(doc: Y.Doc, actorId: string, bytes: Uint8Array) {
  const session = createYjsCollaborationSession({
    doc,
    documentId: 'deck',
    identity: { actorId, name: actorId },
  });
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [collaborationModule({ session })],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return { editor, session };
}

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
const xml = (e: PptxEditorInstance) =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!);
const idOf = (e: PptxEditorInstance, name: string) =>
  e.snapshot().shapes!.find((s) => s.name === name)!.nodeId;
const on = (e: PptxEditorInstance, name: string) =>
  e.setSelection({ kind: 'shapes', nodeIds: [idOf(e, name)] });

async function pair() {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const l = link(docA, docB);
  const a = await replica(docA, 'a', deck());
  const b = await replica(docB, 'b', deck());
  return { a: a.editor, b: b.editor, l };
}

describe('shape commands on two replicas', () => {
  test('concurrent edits on different shapes converge and survive save and reopen', async () => {
    const { a, b, l } = await pair();
    l.pause();
    on(a, 'A');
    expect(a.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 8 } }).ok).toBe(true);
    expect(
      a.exec({ type: 'setShapeStyle', style: { row: 'intenseEffect', color: 'accent2' } }).ok
    ).toBe(true);
    on(b, 'B');
    expect(b.exec({ type: 'flipObjects', axis: 'horizontal' }).ok).toBe(true);
    expect(b.exec({ type: 'rotateObjects', degrees: 90 }).ok).toBe(true);
    expect(b.exec({ type: 'setLockAspectRatio', lock: true }).ok).toBe(true);
    expect(b.exec({ type: 'renameShape', nodeId: idOf(b, 'B'), name: 'Renamed' }).ok).toBe(true);
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    const text = xml(a);
    expect(text).toContain('<p:cNvPr id="3" name="Renamed"/>');
    expect(text).toContain('<a:spLocks noChangeAspect="1"/>');
    expect(text).toContain('<a:effectRef idx="3">');
    for (const e of [a, b]) {
      const r = readPptx(e.save()!);
      expect(r.ok).toBe(true);
    }
  });

  test('undo of a shape command reaches the other replica', async () => {
    const { a, b } = await pair();
    on(a, 'A');
    expect(
      a.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'outerTop' } }).ok
    ).toBe(true);
    expect(xml(b)).toContain('<a:outerShdw');
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(xml(b)).not.toContain('<a:outerShdw');
    expect(fingerprint(a)).toBe(fingerprint(b));
  });

  test('hiding a shape while the other replica deletes it converges', async () => {
    const { a, b, l } = await pair();
    l.pause();
    expect(a.exec({ type: 'setShapesVisible', nodeIds: [idOf(a, 'A')], visible: false }).ok).toBe(
      true
    );
    on(b, 'A');
    expect(b.exec({ type: 'delete' }).ok).toBe(true);
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(a.snapshot().shapes!.some((s) => s.name === 'A')).toBe(false);
  });

  test('concurrent changes to an existing style and shadow keep one of each', async () => {
    const { a, b, l } = await pair();
    on(a, 'A');
    a.exec({ type: 'setShapeStyle', style: { row: 'coloredFill', color: 'accent1' } });
    a.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'outerTop' } });
    a.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 5 } });
    l.pause();
    on(a, 'A');
    on(b, 'A');
    a.exec({ type: 'setShapeStyle', style: { row: 'intenseEffect', color: 'accent2' } });
    b.exec({ type: 'setShapeStyle', style: { row: 'subtleEffect', color: 'accent6' } });
    l.flush();
    expect(xml(a).match(/<p:style>/g)).toHaveLength(1);
    // A style removes the shape's own effects; set them again before the effect phase.
    a.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'outerTop' } });
    a.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 5 } });
    l.pause();
    a.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'outerBottom' } });
    b.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'outerLeft' } });
    a.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 8 } });
    b.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 18 } });
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    const text = xml(a);
    expect(text.match(/<p:style>/g)).toHaveLength(1);
    expect(text.match(/<a:outerShdw /g)).toHaveLength(1);
    expect(text.match(/<a:glow /g)).toHaveLength(1);
  });

  // Schema singletons (the first a:effectLst, p:style, a:spLocks, or one effect's color) take
  // a key from their parent and slot in the Yjs mirror: of two concurrent first copies, one wins
  // and the other stays a hidden companion. The loser's effect does not show.
  test('concurrent first effects, styles, and locks keep one winner of each', async () => {
    const { a, b, l } = await pair();
    l.pause();
    on(a, 'A');
    on(b, 'A');
    a.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'outerTop' } });
    b.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 8 } });
    a.exec({ type: 'setLockAspectRatio', lock: true });
    b.exec({ type: 'setLockAspectRatio', lock: true });
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    let text = xml(a);
    expect(text.match(/<a:effectLst>/g)).toHaveLength(1);
    // One peer's first effect list wins: the shadow or the glow, not both.
    expect(
      (text.match(/<a:outerShdw /g) ?? []).length + (text.match(/<a:glow /g) ?? []).length
    ).toBe(1);
    expect(text.match(/<a:spLocks /g)).toHaveLength(1);
    const styles = (xml(a).match(/<p:style>/g) ?? []).length;
    l.pause();
    on(a, 'B');
    on(b, 'B');
    a.exec({ type: 'setShapeStyle', style: { row: 'intenseEffect', color: 'accent2' } });
    b.exec({ type: 'setShapeStyle', style: { row: 'subtleEffect', color: 'accent6' } });
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    text = xml(a);
    // B had no style: one of the two peers' first styles wins.
    expect(text.match(/<p:style>/g)).toHaveLength(styles + 1);
    expect(readPptx(a.save()!).ok).toBe(true);
  });

  test('concurrent glow colors leave one color in the glow', async () => {
    const { a, b, l } = await pair();
    on(a, 'A');
    a.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 8 } });
    l.pause();
    on(a, 'A');
    on(b, 'A');
    a.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 8, color: '00FF00' } });
    b.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 8, color: 'FF0000' } });
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    const glow = /<a:glow [\s\S]*?<\/a:glow>/.exec(xml(a))![0];
    expect(glow.match(/<a:(srgbClr|schemeClr)\b/g)).toHaveLength(1);
  });
});
