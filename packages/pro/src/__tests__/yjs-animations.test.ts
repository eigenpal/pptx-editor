/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Animation and transition edits of two collaborators: removals of two effects both stay, a
// removal wins over a move, effects added at the same time keep unique ids and build groups,
// and two transitions set at the same time save as one.

import { describe, expect, test } from 'bun:test';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { presentationInfo, serializeXmlPart, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { clickEntrances } from '../../../../scripts/fixtures/timing.ts';
import { openPeers } from '../../../../scripts/fidelity/collaboration-kit.ts';

const IN = 914_400;

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
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

/** A slide with text shapes A (2) and B (3), and shape C (4), with click entrances on A and B. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: box(2, 'A', 0.5) + box(3, 'B', 3) + box(4, 'C', 6, false),
        ext: clickEntrances([2, 3]),
      },
    ],
  });

const effects = (e: PptxEditorInstance) => e.snapshot().animations!.effects;
const select = (e: PptxEditorInstance, name: string): void =>
  e.setSelection({
    kind: 'shapes',
    nodeIds: [e.snapshot().shapes!.find((s) => s.name === name)!.nodeId],
  });
const timeNodeIds = (xml: string): string[] =>
  [...xml.matchAll(/<p:cTn [^>]*?id="(\d+)"/g)].map((m) => m[1]!);
const savedXml = async (e: PptxEditorInstance) => slideXml(await open(e.save()!));

describe('animations in a collaboration room', () => {
  test('removals of two different effects both stay', async () => {
    const { editor: a, peer: b, link, destroy } = await openPeers(deck());
    try {
      link.pause();
      expect(a.exec({ type: 'removeAnimation', effectId: effects(a)[0]!.id }).ok).toBe(true);
      expect(b.exec({ type: 'removeAnimation', effectId: effects(b)[1]!.id }).ok).toBe(true);
      link.flush();
      expect(slideXml(a)).toBe(slideXml(b));
      expect(effects(a)).toEqual([]);
      expect(await savedXml(a)).not.toContain('p:timing');
    } finally {
      destroy();
    }
  });

  test('a removal wins over a move of the same effect', async () => {
    const { editor: a, peer: b, link, destroy } = await openPeers(deck());
    try {
      link.pause();
      expect(a.exec({ type: 'moveAnimation', effectId: effects(a)[1]!.id, to: 0 }).ok).toBe(true);
      expect(b.exec({ type: 'removeAnimation', effectId: effects(b)[1]!.id }).ok).toBe(true);
      link.flush();
      expect(slideXml(a)).toBe(slideXml(b));
      expect(effects(a).map((x) => x.shapeName)).toEqual(['A']);
      expect(await savedXml(a)).not.toContain('spid="3"');
    } finally {
      destroy();
    }
  });

  test('two effects added to one shape take two build groups', async () => {
    const { editor: a, peer: b, link, destroy } = await openPeers(deck());
    try {
      link.pause();
      select(a, 'A');
      expect(a.exec({ type: 'addAnimation', preset: 'pulse' }).ok).toBe(true);
      select(b, 'A');
      expect(b.exec({ type: 'addAnimation', preset: 'spin' }).ok).toBe(true);
      link.flush();
      expect(effects(a).map((x) => x.preset)).toEqual(['appear', 'appear', 'pulse', 'spin']);
      const saved = await savedXml(a);
      const groups = [...saved.matchAll(/<p:cTn [^>]*grpId="(\d+)"[^>]*presetClass="emph"/g)];
      expect(new Set(groups.map((m) => m[1])).size).toBe(2);
      const builds = saved.match(/<p:bldP [^>]*>/g) ?? [];
      const keys = builds.map((x) => /spid="\d+"/.exec(x)![0] + /grpId="\d+"/.exec(x)![0]);
      expect(keys.filter((k) => k.startsWith('spid="2"'))).toHaveLength(3);
      expect(new Set(keys).size).toBe(keys.length);
      const all = timeNodeIds(saved);
      expect(new Set(all).size).toBe(all.length);
    } finally {
      destroy();
    }
  });

  test('concurrent effect edits converge with unique time node ids', async () => {
    const { editor: a, peer: b, link, destroy } = await openPeers(deck());
    try {
      link.pause();
      select(a, 'C');
      expect(a.exec({ type: 'addAnimation', preset: 'fade', trigger: 'withPrevious' }).ok).toBe(
        true
      );
      expect(b.exec({ type: 'removeAnimation', effectId: effects(b)[0]!.id }).ok).toBe(true);
      select(b, 'A');
      expect(b.exec({ type: 'addAnimation', preset: 'spin', trigger: 'afterPrevious' }).ok).toBe(
        true
      );
      link.flush();
      const xml = slideXml(a);
      expect(xml).toBe(slideXml(b));
      const ids = timeNodeIds(xml);
      expect(new Set(ids).size).toBe(ids.length);
      expect(effects(a).map((x) => x.preset)).toContain('spin');
      expect(effects(a).map((x) => x.preset)).toContain('fade');
    } finally {
      destroy();
    }
  });

  test('two transitions set at the same time save as one, before the timing', async () => {
    const { editor: a, peer: b, link, destroy } = await openPeers(deck());
    try {
      link.pause();
      expect(a.exec({ type: 'setTransition', kind: 'fade' }).ok).toBe(true);
      expect(b.exec({ type: 'setTransition', kind: 'push' }).ok).toBe(true);
      link.flush();
      const xml = await savedXml(a);
      expect(xml.match(/<mc:AlternateContent/g)?.length).toBe(1);
      expect(xml.indexOf('<mc:AlternateContent')).toBeLessThan(xml.indexOf('<p:timing>'));
    } finally {
      destroy();
    }
  });
});
