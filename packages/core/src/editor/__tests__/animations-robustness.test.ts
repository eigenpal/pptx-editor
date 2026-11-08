// The animation and transition commands with hostile file values and unusual timing trees:
// names from the file index no object prototype, and wrappers keep their
// markup. The collaboration cases are in `packages/pro/src/__tests__/yjs-animations.test.ts`.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { timingWriter } from '../../../../../scripts/fixtures/timing.ts';
import { presentationInfo, serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const r = await e.load(bytes);
  if (!r.ok) throw new Error('load failed');
  return e;
}

const slideXml = (e: PptxEditorInstance): string => {
  const pkg = e.session!.store.package;
  return serializeXmlPart(xmlPartOf(pkg, presentationInfo(pkg).slides[0]!.partName)!);
};

const box = (id: number, name: string, x: number, text = true): string =>
  shape({
    id,
    name,
    x: x * IN,
    y: IN,
    cx: 2 * IN,
    cy: IN,
    ...(text ? { text: para(run(name)) } : {}),
  });

/** A slide with text shapes A (2) and B (3), shape C (4), and `ext` after the color map. */
const deck = (ext = ''): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: box(2, 'A', 0.5) + box(3, 'B', 3) + box(4, 'C', 6, false),
        ...(ext ? { ext } : {}),
      },
    ],
  });

const effects = (e: PptxEditorInstance) => e.snapshot().animations!.effects;
const select = (e: PptxEditorInstance, name: string): void =>
  e.setSelection({
    kind: 'shapes',
    nodeIds: [e.snapshot().shapes!.find((s) => s.name === name)!.nodeId],
  });
const ids = (xml: string): string[] =>
  [...xml.matchAll(/<p:cTn [^>]*?id="(\d+)"/g)].map((m) => m[1]);

/** The main sequence of `steps` (click steps of `timingWriter`), with builds for A and B. */
function timing(steps: (w: ReturnType<typeof timingWriter>) => string): string {
  const w = timingWriter();
  return w.timing(w.mainSeq(steps(w)), [2, 3]);
}

describe('values that name object keys', () => {
  test('setTransition refuses a kind that is a prototype key', async () => {
    const e = await open(deck());
    for (const kind of ['__proto__', 'constructor', 'toString', 'hasOwnProperty'])
      expect(e.exec({ type: 'setTransition', kind } as never)).toMatchObject({
        ok: false,
        reason: 'errors.unsupported',
      });
    expect(slideXml(e)).not.toContain('p:transition');
  });

  test('file values that are prototype keys read as unknown', async () => {
    const t = timing((w) => w.click(w.effect(2)))
      .replace('presetClass="entr"', 'presetClass="constructor"')
      .replace('nodeType="clickEffect"', 'nodeType="toString"');
    const e = await open(deck('<p:transition spd="valueOf"><p:toString/></p:transition>' + t));
    const animations = e.snapshot().animations!;
    // A snapshot holds plain data only.
    expect(() => structuredClone(animations)).not.toThrow();
    expect(animations.transition).toMatchObject({ kind: 'other', durationMs: 500 });
    expect(animations.effects[0]).toMatchObject({ effectClass: 'other', trigger: 'withPrevious' });
    // An effect with an unknown `nodeType` makes the main sequence read-only.
    expect(animations.editable).toBe(false);
  });

  test('a main sequence id past xsd:unsignedInt is read-only', async () => {
    const t = timing(() => '').replace('<p:cTn id="2"', '<p:cTn id="9999999999"');
    const e = await open(deck(t));
    expect(e.snapshot().animations!.editable).toBe(false);
    select(e, 'A');
    expect(e.exec({ type: 'addAnimation', preset: 'fade', trigger: 'withPrevious' }).ok).toBe(
      false
    );
  });
});

describe('the timing tree', () => {
  test('a shape target deep in an effect is found', async () => {
    let inner = timingWriter().show(2).replace('id="10"', 'id="3"');
    for (let i = 0; i < 30; i += 1)
      inner = `<p:par><p:cTn id="${1000 + i}"><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`;
    const t = timing((w) => w.click(w.effect(2).replace(/<p:set>.*<\/p:set>/, inner)));
    const e = await open(deck(t));
    expect(effects(e)[0].shapeId).toBe('2');
    select(e, 'B');
    expect(e.exec({ type: 'addAnimation', preset: 'fade' }).ok).toBe(true);
    const all = ids(slideXml(e));
    expect(new Set(all).size).toBe(all.length);
  });

  test('a step reads from its start conditions', async () => {
    // The first step starts with the slide, though its effect says "on click".
    const t = timing((w) => w.click(w.effect(2)) + w.click(w.effect(3))).replace(
      '<p:cond delay="indefinite"/>',
      '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>'
    );
    const e = await open(deck(t));
    expect(effects(e).map((x) => x.step)).toEqual([0, 1]);
  });

  test('a new order keeps the markup and the ids of the step and group wrappers', async () => {
    // Steps 12 (A) and 16 (B); each wrapper has an attribute and a condition of its own.
    const t = timing((w) => w.click(w.effect(2)) + w.click(w.effect(3)))
      .replace('<p:cTn id="12" fill="hold">', '<p:cTn id="12" fill="hold" accel="10000">')
      .replace(
        '<p:cTn id="17" fill="hold"><p:stCondLst><p:cond delay="0"/>',
        '<p:cTn id="17" fill="hold"><p:stCondLst><p:cond delay="0"/><p:cond evt="onEnd" delay="0"><p:tn val="12"/></p:cond>'
      );
    const e = await open(deck(t));
    const b = effects(e)[1].id;
    expect(e.exec({ type: 'moveAnimation', effectId: b, to: 0 }).ok).toBe(true);
    const xml = slideXml(e);
    expect(effects(e).map((x) => x.shapeName)).toEqual(['B', 'A']);
    expect(xml).toContain('<p:cTn id="12" fill="hold" accel="10000">');
    expect(xml).toContain('<p:cond evt="onEnd" delay="0"><p:tn val="12"/></p:cond>');
    // Only the wrappers moved: each effect is in its own step still.
    expect(xml.indexOf('id="16"')).toBeLessThan(xml.indexOf('id="12"'));
    expect(xml.indexOf('id="12"')).toBeLessThan(xml.indexOf('id="10"'));
  });

  test('a removal keeps when the other effects play', async () => {
    // Step 1: A on click, B with it. A goes: B still plays on that click.
    const t = timing((w) =>
      w.click(w.effect(2) + w.effect(3).replace('nodeType="clickEffect"', 'nodeType="withEffect"'))
    );
    const e = await open(deck(t));
    expect(e.exec({ type: 'removeAnimation', effectId: effects(e)[0].id }).ok).toBe(true);
    expect(effects(e).map((x) => [x.shapeName, x.trigger, x.step])).toEqual([['B', 'onClick', 1]]);
  });

  test('the last removal takes the whole timing', async () => {
    // A build entry for C, which has no effect.
    const w = timingWriter();
    const e = await open(deck(w.timing(w.mainSeq(w.click(w.effect(2))), [2, 4])));
    expect(e.exec({ type: 'removeAnimation', effectId: effects(e)[0].id }).ok).toBe(true);
    expect(slideXml(e)).not.toContain('p:timing');
  });
});

describe('transitions', () => {
  test('a duration that is no speed keeps its precision', async () => {
    const e = await open(deck('<p:transition spd="fast" advTm="3000"><p:fade/></p:transition>'));
    expect(e.exec({ type: 'setTransition', durationMs: 1200 }).ok).toBe(true);
    expect(e.snapshot().animations!.transition).toMatchObject({
      kind: 'fade',
      durationMs: 1200,
      advanceAfterMs: 3000,
    });
    const xml = slideXml(e);
    expect(xml).toContain('p14:dur="1200"');
    expect(xml).toContain('<mc:Fallback><p:transition spd="slow" advTm="3000"><p:fade/>');
    // A speed value stays plain.
    const f = await open(deck('<p:transition spd="fast"><p:fade/></p:transition>'));
    expect(f.exec({ type: 'setTransition', durationMs: 1000 }).ok).toBe(true);
    expect(slideXml(f)).toContain('<p:transition spd="slow"><p:fade/></p:transition>');
  });
});

describe('the save', () => {
  test('a changed timing saves one build entry for each shape and group', async () => {
    const w = timingWriter();
    const e = await open(deck(w.timing(w.mainSeq(w.click(w.effect(2))), [2, 2])));
    // The file's own timing saves as the file wrote it.
    expect(slideXml(await open(e.save()!)).match(/<p:bldP spid="2"/g)).toHaveLength(2);
    select(e, 'B');
    expect(e.exec({ type: 'addAnimation', preset: 'fade' }).ok).toBe(true);
    expect(slideXml(await open(e.save()!)).match(/<p:bldP spid="2"/g)).toHaveLength(1);
  });
});
