// The animation and transition commands: the effects of the main sequence (ECMA-376 Part 1,
// §19.5) and the slide transition (§19.3.1.50).

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { clickEntrances } from '../../../../../scripts/fixtures/timing.ts';
import { presentationInfo, serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { animationMarks } from '../surface-animation-marks.ts';

const IN = 914_400;
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const slideXml = (e: PptxEditorInstance, index = 0): string => {
  const pkg = e.session!.store.package;
  return serializeXmlPart(xmlPartOf(pkg, presentationInfo(pkg).slides[index]!.partName)!);
};

const box = (id: number, name: string, x: number, text?: string): string =>
  shape({
    id,
    name,
    x: x * IN,
    y: IN,
    cx: 2 * IN,
    cy: IN,
    ...(text !== undefined ? { text: para(run(text)) } : {}),
  });

/** A slide with text shapes A and B, a shape C without text, and `ext` after the color map. */
function deck(ext = ''): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: box(2, 'A', 0.5, 'Alpha') + box(3, 'B', 3, 'Beta') + box(4, 'C', 6),
        ...(ext ? { ext } : {}),
      },
    ],
  });
}

const node = (e: PptxEditorInstance, name: string): string =>
  e.snapshot().shapes!.find((s) => s.name === name)!.nodeId;

function select(e: PptxEditorInstance, ...names: string[]): void {
  e.setSelection({ kind: 'shapes', nodeIds: names.map((n) => node(e, n)) });
}

const effects = (e: PptxEditorInstance) => e.snapshot().animations!.effects;

/** Why a command was refused, or undefined when it ran. */
const reasonOf = (r: { ok: boolean; reason?: string }): string | undefined => r.reason;

/** Every `p:cTn/@id` of a slide's XML. */
const timeNodeIds = (xml: string): string[] =>
  [...xml.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => m[1]);

describe('add an effect', () => {
  test('writes the timing tree, a build entry for text, and lists the effect', async () => {
    const e = await open(deck());
    expect(effects(e)).toEqual([]);
    expect(e.snapshot().animations!.editable).toBe(true);
    select(e, 'A');
    expect(e.exec({ type: 'addAnimation', preset: 'flyIn' }).ok).toBe(true);
    const xml = slideXml(e);
    expect(xml).toContain('nodeType="tmRoot"');
    expect(xml).toContain('nodeType="mainSeq"');
    expect(xml).toContain(
      'presetID="2" presetClass="entr" presetSubtype="4" fill="hold" grpId="0" nodeType="clickEffect"'
    );
    expect(xml).toContain('<p:attrName>ppt_y</p:attrName>');
    expect(xml).toContain('<p:bldP spid="2" grpId="0" animBg="1"/>');
    // `CT_Slide` order: the timing after the color map and before the extension list.
    expect(xml.indexOf('</p:clrMapOvr>')).toBeLessThan(xml.indexOf('<p:timing>'));
    const [fly] = effects(e);
    expect(fly).toMatchObject({
      shapeName: 'A',
      nodeId: node(e, 'A'),
      effectClass: 'entrance',
      preset: 'flyIn',
      trigger: 'onClick',
      sequence: 'main',
      step: 1,
      delayMs: 0,
      durationMs: 500,
      editable: true,
    });
  });

  test('writes each class, keeps ids unique, and gives a shape one group id per effect', async () => {
    const e = await open(deck());
    select(e, 'C');
    for (const preset of ['appear', 'pulse', 'spin', 'growShrink', 'fadeOut', 'zoomOut'] as const)
      expect(e.exec({ type: 'addAnimation', preset }).ok).toBe(true);
    const list = effects(e);
    expect(list.map((x) => x.effectClass)).toEqual([
      'entrance',
      'emphasis',
      'emphasis',
      'emphasis',
      'exit',
      'exit',
    ]);
    expect(list.map((x) => x.durationMs)).toEqual([null, 500, 2000, 2000, 500, 500]);
    const xml = slideXml(e);
    const ids = timeNodeIds(xml);
    expect(new Set(ids).size).toBe(ids.length);
    expect([...xml.matchAll(/grpId="(\d+)" nodeType/g)].map((m) => m[1])).toEqual([
      '0',
      '1',
      '2',
      '3',
      '4',
      '5',
    ]);
    // C has no text: no build entry.
    expect(xml).not.toContain('<p:bldLst>');
    // The exit hides the shape at the end of the effect.
    expect(xml).toContain('<p:cond delay="499"/>');
  });

  test('with previous and after previous join the click step, and groups start in turn', async () => {
    const e = await open(deck());
    select(e, 'A');
    e.exec({ type: 'addAnimation', preset: 'fade' });
    select(e, 'B');
    e.exec({ type: 'addAnimation', preset: 'zoom', trigger: 'withPrevious' });
    select(e, 'C');
    e.exec({ type: 'addAnimation', preset: 'spin', trigger: 'afterPrevious' });
    expect(effects(e).map((x) => [x.shapeName, x.trigger, x.step])).toEqual([
      ['A', 'onClick', 1],
      ['B', 'withPrevious', 1],
      ['C', 'afterPrevious', 1],
    ]);
    const xml = slideXml(e);
    // One click step: two time groups, the second at the end of the first (500 ms).
    expect(xml.match(/<p:cond delay="indefinite"\/>/g)?.length).toBe(1);
    expect(xml).toMatch(/<p:cTn id="\d+" fill="hold"><p:stCondLst><p:cond delay="500"\/>/);
  });

  test('a first effect that plays with the slide waits for the slide start', async () => {
    const e = await open(deck());
    select(e, 'A');
    e.exec({ type: 'addAnimation', preset: 'appear', trigger: 'afterPrevious' });
    expect(effects(e)[0].step).toBe(0);
    expect(slideXml(e)).toMatch(/<p:cond evt="onBegin" delay="0"><p:tn val="\d+"\/><\/p:cond>/);
  });

  test('a shape in a group animates its group, and a selection of two adds two effects', async () => {
    const e = await open(deck());
    select(e, 'A', 'B');
    expect(e.exec({ type: 'group' }).ok).toBe(true);
    const group = e.snapshot().shapes!.find((s) => s.children?.length)!;
    const a = group.children!.find((s) => s.name === 'A')!;
    e.setSelection({ kind: 'shapes', nodeIds: [a.nodeId], group: group.nodeId });
    e.exec({ type: 'addAnimation', preset: 'wipe' });
    expect(effects(e)[0].nodeId).toBe(group.nodeId);
    select(e, 'C');
    e.setSelection({ kind: 'shapes', nodeIds: [node(e, 'C'), group.nodeId] });
    e.exec({ type: 'addAnimation', preset: 'pulse' });
    expect(effects(e).map((x) => [x.preset, x.trigger])).toEqual([
      ['wipe', 'onClick'],
      ['pulse', 'onClick'],
      ['pulse', 'withPrevious'],
    ]);
  });

  test('refuses with no selection, an unknown preset, and in view mode', async () => {
    const e = await open(deck());
    expect(e.exec({ type: 'addAnimation', preset: 'fade' }).ok).toBe(false);
    select(e, 'A');
    expect(e.can({ type: 'addAnimation', preset: 'bounce' as never }).ok).toBe(false);
    e.setMode('view');
    expect(e.can({ type: 'addAnimation', preset: 'fade' }).ok).toBe(false);
  });
});

describe('edit effects', () => {
  async function three(): Promise<PptxEditorInstance> {
    const e = await open(deck());
    for (const [name, preset] of [
      ['A', 'fade'],
      ['B', 'flyIn'],
      ['C', 'spin'],
    ] as const) {
      select(e, name);
      e.exec({ type: 'addAnimation', preset });
    }
    return e;
  }

  test('reorders, and undo restores the exact XML', async () => {
    const e = await three();
    const before = slideXml(e);
    const c = effects(e)[2];
    expect(e.exec({ type: 'moveAnimation', effectId: c.id, to: 0 }).ok).toBe(true);
    expect(effects(e).map((x) => x.shapeName)).toEqual(['C', 'A', 'B']);
    expect(e.exec({ type: 'moveAnimation', effectId: c.id, to: 0 }).ok).toBe(false);
    e.exec({ type: 'undo' });
    expect(slideXml(e)).toBe(before);
  });

  test('changes trigger, delay, and duration', async () => {
    const e = await three();
    const b = effects(e)[1];
    expect(
      e.exec({
        type: 'setAnimationTiming',
        effectId: b.id,
        trigger: 'withPrevious',
        delayMs: 250,
        durationMs: 1000,
      }).ok
    ).toBe(true);
    expect(effects(e)[1]).toMatchObject({
      trigger: 'withPrevious',
      delayMs: 250,
      durationMs: 1000,
      step: 1,
    });
    expect(effects(e)[2].step).toBe(2);
    // The same values again change nothing.
    expect(e.exec({ type: 'setAnimationTiming', effectId: b.id, delayMs: 250 }).ok).toBe(false);
    expect(e.exec({ type: 'setAnimationTiming', effectId: b.id, durationMs: 5 }).ok).toBe(false);
    expect(e.exec({ type: 'setAnimationTiming', effectId: b.id, delayMs: Number.NaN }).ok).toBe(
      false
    );
  });

  test('refuses a duration for an instant effect', async () => {
    const e = await open(deck());
    select(e, 'A');
    e.exec({ type: 'addAnimation', preset: 'appear' });
    const [appear] = effects(e);
    expect(
      reasonOf(e.exec({ type: 'setAnimationTiming', effectId: appear.id, durationMs: 800 }))
    ).toBe('disabled.instantAnimation');
  });

  test('a removal takes the build entry, and the last removal takes the timing', async () => {
    const e = await three();
    const before = slideXml(e);
    const [a, b, c] = effects(e);
    e.exec({ type: 'removeAnimation', effectId: a.id });
    expect(slideXml(e)).not.toContain('<p:bldP spid="2"');
    expect(slideXml(e)).toContain('<p:bldP spid="3"');
    e.exec({ type: 'removeAnimation', effectId: b.id });
    expect(slideXml(e)).not.toContain('<p:bldLst>');
    e.exec({ type: 'removeAnimation', effectId: c.id });
    expect(slideXml(e)).not.toContain('<p:timing>');
    expect(effects(e)).toEqual([]);
    e.exec({ type: 'undo' });
    e.exec({ type: 'undo' });
    e.exec({ type: 'undo' });
    expect(slideXml(e)).toBe(before);
    expect(reasonOf(e.exec({ type: 'removeAnimation', effectId: '999999' }))).toBe(
      'disabled.noAnimation'
    );
  });

  test('effects survive save and reopen', async () => {
    const e = await three();
    const again = await open(e.save()!);
    expect(effects(again).map((x) => [x.shapeName, x.preset])).toEqual([
      ['A', 'fade'],
      ['B', 'flyIn'],
      ['C', 'spin'],
    ]);
  });
});

describe('kept timing', () => {
  test('lists the effects of a file, and edits next to them', async () => {
    const e = await open(deck(clickEntrances([2, 3])));
    expect(effects(e).map((x) => [x.shapeName, x.preset, x.step])).toEqual([
      ['A', 'appear', 1],
      ['B', 'appear', 2],
    ]);
    select(e, 'C');
    e.exec({ type: 'addAnimation', preset: 'zoom' });
    expect(effects(e).map((x) => x.shapeName)).toEqual(['A', 'B', 'C']);
    const ids = timeNodeIds(slideXml(e));
    expect(new Set(ids).size).toBe(ids.length);
  });

  test('a paragraph effect shows its paragraphs', async () => {
    const timing =
      '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>' +
      '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' +
      '<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>' +
      '<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
      '<p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">' +
      '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
      '<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="6" dur="500"/>' +
      '<p:tgtEl><p:spTgt spid="2"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect>' +
      '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>' +
      '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
      '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>' +
      '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst><p:bldP spid="2" grpId="0" build="p"/></p:bldLst></p:timing>';
    const e = await open(deck(timing));
    expect(effects(e)[0]).toMatchObject({ preset: 'fade', paragraphs: { start: 0, end: 0 } });
  });

  test('a main sequence id that is not a number is read-only, so it never reaches new markup', async () => {
    const e = await open(
      deck(clickEntrances([2]).replace('<p:cTn id="2"', '<p:cTn id="2&quot;&gt;"'))
    );
    expect(e.snapshot().animations!.editable).toBe(false);
    select(e, 'B');
    expect(reasonOf(e.exec({ type: 'addAnimation', preset: 'fade' }))).toBe(
      'disabled.animationsReadOnly'
    );
  });

  test('timing in markup-compatibility content is read-only', async () => {
    const wrapped =
      `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:p14="${P14}" Requires="p14">` +
      clickEntrances([2]) +
      '</mc:Choice><mc:Fallback>' +
      clickEntrances([2]) +
      '</mc:Fallback></mc:AlternateContent>';
    const e = await open(deck(wrapped));
    const animations = e.snapshot().animations!;
    expect(animations.editable).toBe(false);
    expect(animations.effects.map((x) => [x.shapeName, x.editable])).toEqual([['A', false]]);
    select(e, 'B');
    expect(reasonOf(e.exec({ type: 'addAnimation', preset: 'fade' }))).toBe(
      'disabled.animationsReadOnly'
    );
  });

  test('an interactive sequence is listed read-only with its trigger shape', async () => {
    const interactive =
      '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>' +
      '<p:seq concurrent="1" nextAc="seek"><p:cTn id="7" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">' +
      '<p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="4"/></p:tgtEl></p:cond></p:stCondLst>' +
      '<p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst>' +
      '<p:par><p:cTn id="8" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">' +
      '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="9" dur="1" fill="hold">' +
      '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="2"/></p:tgtEl>' +
      '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>' +
      '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>' +
      '<p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="4"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>' +
      '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>';
    const f = await open(deck(interactive));
    expect(effects(f)).toEqual([
      expect.objectContaining({
        shapeName: 'A',
        sequence: 'interactive',
        triggerShapeId: '4',
        step: null,
        editable: false,
      }),
    ]);
    // A new effect adds a main sequence before the interactive one.
    select(f, 'B');
    expect(f.exec({ type: 'addAnimation', preset: 'fade' }).ok).toBe(true);
    expect(effects(f).map((x) => [x.shapeName, x.sequence])).toEqual([
      ['B', 'main'],
      ['A', 'interactive'],
    ]);
  });
});

describe('transitions', () => {
  test('sets a kind with a duration and advance, and clears it', async () => {
    const e = await open(deck());
    expect(e.snapshot().animations!.transition.kind).toBe('none');
    expect(
      e.exec({ type: 'setTransition', kind: 'push', durationMs: 1200, advanceAfterMs: 3000 }).ok
    ).toBe(true);
    expect(e.snapshot().animations!.transition).toEqual({
      kind: 'push',
      durationMs: 1200,
      advanceOnClick: true,
      advanceAfterMs: 3000,
    });
    const xml = slideXml(e);
    expect(xml).toContain('Requires="p14"');
    expect(xml).toContain('p14:dur="1200"');
    expect(xml).toContain('<mc:Fallback><p:transition spd="slow" advTm="3000"><p:push dir="u"/>');
    expect(e.exec({ type: 'setTransition', advanceOnClick: false }).ok).toBe(true);
    expect(slideXml(e).match(/advClick="0"/g)?.length).toBe(2);
    expect(
      e.exec({ type: 'setTransition', kind: 'none', advanceOnClick: true, advanceAfterMs: null }).ok
    ).toBe(true);
    expect(slideXml(e)).not.toContain('transition');
    const again = await open(e.save()!);
    expect(again.snapshot().animations!.transition.kind).toBe('none');
  });

  test('the transition goes before the timing, and every slide can take it', async () => {
    const two = buildPptx({
      slides: [
        { layout: 'blank', shapes: box(2, 'A', 1, 'A'), ext: clickEntrances([2]) },
        { layout: 'blank', shapes: box(2, 'A', 1, 'A') },
      ],
    });
    const e = await open(two);
    expect(e.exec({ type: 'setTransition', kind: 'fade', indices: [0, 1] }).ok).toBe(true);
    const xml = slideXml(e, 0);
    expect(xml.indexOf('<mc:AlternateContent')).toBeLessThan(xml.indexOf('<p:timing>'));
    expect(slideXml(e, 1)).toContain('<p:fade/>');
    expect(e.exec({ type: 'setTransition', kind: 'fade', indices: [0, 1] }).ok).toBe(false);
    expect(e.exec({ type: 'setTransition', kind: 'fade', indices: [5] }).ok).toBe(false);
  });

  test('advance and duration edits keep a kind the editor does not write', async () => {
    const P159 = 'http://schemas.microsoft.com/office/powerpoint/2015/09/main';
    const morph =
      `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:p159="${P159}" Requires="p159">` +
      '<p:transition spd="slow"><p159:morph option="byObject"/></p:transition></mc:Choice>' +
      '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>';
    const e = await open(deck(morph));
    expect(e.snapshot().animations!.transition.kind).toBe('other');
    expect(e.exec({ type: 'setTransition', advanceAfterMs: 2000, durationMs: 500 }).ok).toBe(true);
    const xml = slideXml(e);
    expect(xml).toContain('<p159:morph option="byObject"/>');
    expect(xml.match(/advTm="2000"/g)?.length).toBe(2);
    expect(xml.match(/spd="fast"/g)?.length).toBe(2);
    // A new kind replaces the whole markup.
    expect(e.exec({ type: 'setTransition', kind: 'wipe' }).ok).toBe(true);
    expect(slideXml(e)).not.toContain('morph');
    expect(e.snapshot().animations!.transition).toMatchObject({
      kind: 'wipe',
      advanceAfterMs: 2000,
    });
  });

  test('a save leaves the transitions that the file wrote', async () => {
    const two =
      '<p:transition><p:fade/></p:transition>' +
      clickEntrances([2]) +
      '<p:transition><p:cut/></p:transition>';
    const e = await open(deck(two));
    expect(slideXml(await open(e.save()!))).toBe(slideXml(e));
  });

  test('refuses values out of range', async () => {
    const e = await open(deck());
    expect(e.exec({ type: 'setTransition', kind: 'fade', durationMs: 0 }).ok).toBe(false);
    expect(e.exec({ type: 'setTransition', kind: 'fade', advanceAfterMs: -1 }).ok).toBe(false);
    expect(e.exec({ type: 'setTransition', kind: 'spin' as never }).ok).toBe(false);
  });
});

describe('order marks', () => {
  test('is a view setting with no undo step', async () => {
    const e = await open(deck());
    expect(e.snapshot().animationOrder).toBe(false);
    expect(e.exec({ type: 'showAnimationOrder', show: true }).ok).toBe(true);
    expect(e.snapshot().animationOrder).toBe(true);
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('tags each animated shape with the steps of its effects', async () => {
    const e = await open(deck());
    select(e, 'A');
    e.exec({ type: 'addAnimation', preset: 'fade', trigger: 'afterPrevious' });
    e.exec({ type: 'addAnimation', preset: 'spin' });
    select(e, 'B');
    e.exec({ type: 'addAnimation', preset: 'zoom', trigger: 'withPrevious' });
    expect(animationMarks(effects(e))).toEqual([
      { nodeId: node(e, 'A'), label: '0, 1' },
      { nodeId: node(e, 'B'), label: '1' },
    ]);
  });
});
