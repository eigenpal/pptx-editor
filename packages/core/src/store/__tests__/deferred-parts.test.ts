// Deferred XML parts: a slide read with `defer` parses on first read, to the tree an eager read
// gives; the slide list reads `show` without parsing; a part that does not parse reads as a
// stand-in, is reported, and refuses the write.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { readZip, writeZip } from '@eigenpal/ooxml-core';
import {
  canonicalPackageFingerprint,
  deferredRejection,
  isDeferredPart,
  isDeferredPending,
  parseDeferredParts,
  pendingDeferredBytes,
  peekDeferredRootAttribute,
  presentationInfo,
  readPptx,
  serializeXmlPart,
  writePptx,
  PPTX_ZIP_LIMITS,
  type PptxPackage,
} from '../index.ts';

const IN = 914_400;
const deck = (n: number): Uint8Array =>
  buildPptx({
    slides: Array.from({ length: n }, (_, i) => ({
      shapes: shape({
        id: 2,
        name: 'T',
        x: IN,
        y: IN,
        cx: 4 * IN,
        cy: IN,
        text: para(run(`S${i}`)),
      }),
    })),
  });

const deferAll = (bytes: Uint8Array) => {
  const r = readPptx(bytes, PPTX_ZIP_LIMITS, { defer: () => true });
  if (!r.ok) throw new Error(JSON.stringify(r.error));
  return r.pkg;
};

/** Every XML part serialized, by name. */
const serialized = (pkg: PptxPackage): Map<string, string> =>
  new Map(
    [...pkg.parts.values()].map((p) => [
      p.partName,
      p.kind === 'xml' ? serializeXmlPart(p.xml) : `${p.bytes.length}`,
    ])
  );

/** `bytes` with entry `name` replaced by `text`. */
function withEntry(bytes: Uint8Array, name: string, text: string): Uint8Array {
  const zip = readZip(bytes);
  if (!zip.ok) throw new Error('zip');
  const entries = new Map(zip.entries);
  entries.set(name, new TextEncoder().encode(text));
  return writeZip(entries);
}

describe('deferred parts', () => {
  test('only slides stay unparsed, and each parses on first read to the eager tree', () => {
    const bytes = deck(3);
    const pkg = deferAll(bytes);
    const slides = presentationInfo(pkg).slides.map((s) => s.partName);
    expect(slides.length).toBe(3);
    for (const name of slides) expect(isDeferredPending(pkg.parts.get(name))).toBe(true);
    // Masters, layouts, themes, and the presentation part parse as before.
    for (const part of pkg.parts.values())
      if (!slides.includes(part.partName)) expect(isDeferredPart(part)).toBe(false);
    expect(pendingDeferredBytes(pkg)).toBeGreaterThan(0);
    const eager = readPptx(bytes);
    if (!eager.ok) throw new Error('eager');
    // Reading the slide list did not parse any slide.
    expect(isDeferredPending(pkg.parts.get(slides[1]!))).toBe(true);
    expect(serialized(pkg)).toEqual(serialized(eager.pkg));
    for (const name of slides) expect(isDeferredPending(pkg.parts.get(name))).toBe(false);
    expect(pendingDeferredBytes(pkg)).toBe(0);
    expect(canonicalPackageFingerprint(pkg)).toBe(canonicalPackageFingerprint(eager.pkg));
  });

  test('a deferred part keeps one tree once read', () => {
    const pkg = deferAll(deck(2));
    const name = presentationInfo(pkg).slides[0]!.partName;
    const part = pkg.parts.get(name)!;
    if (part.kind !== 'xml') throw new Error('kind');
    expect(part.xml).toBe(part.xml);
  });

  test('the slide list reads `show` from the start tag, without parsing', () => {
    const base = deck(2);
    const pkg0 = deferAll(base);
    const name = presentationInfo(pkg0).slides[1]!.partName;
    const part = pkg0.parts.get(name)!;
    if (part.kind !== 'xml') throw new Error('kind');
    const xml = serializeXmlPart(part.xml).replace('<p:sld ', '<p:sld show="0" ');
    const pkg = deferAll(withEntry(base, name, xml));
    const info = presentationInfo(pkg);
    expect(info.slides.map((s) => s.hidden)).toEqual([false, true]);
    expect(isDeferredPending(pkg.parts.get(name))).toBe(true);
    expect(peekDeferredRootAttribute(pkg.parts.get(name)!, 'show')).toBe('0');
    expect(peekDeferredRootAttribute(pkg.parts.get(name)!, 'absent')).toBeUndefined();
  });

  test('a start tag that cannot be peeked answers null, and the reader parses instead', () => {
    const base = deck(2);
    const pkg0 = deferAll(base);
    const name = presentationInfo(pkg0).slides[1]!.partName;
    const part = pkg0.parts.get(name)!;
    if (part.kind !== 'xml') throw new Error('kind');
    const xml = serializeXmlPart(part.xml).replace('<p:sld ', '<p:sld show="&#48;" ');
    const pkg = deferAll(withEntry(base, name, xml));
    expect(peekDeferredRootAttribute(pkg.parts.get(name)!, 'show')).toBeNull();
    expect(presentationInfo(pkg).slides[1]!.hidden).toBe(true);
    expect(isDeferredPending(pkg.parts.get(name))).toBe(false);
  });

  test('a slide that does not parse reads as a stand-in, is reported, and refuses the write', () => {
    const base = deck(3);
    const pkg0 = deferAll(base);
    const name = presentationInfo(pkg0).slides[2]!.partName;
    const broken = withEntry(base, name, '<p:sld xmlns:p="urn:x"><p:cSld>');
    // An eager read refuses the deck.
    expect(readPptx(broken)).toMatchObject({ ok: false, error: { reason: 'xml', partName: name } });
    const pkg = deferAll(broken);
    expect(deferredRejection(pkg)).toBeUndefined();
    parseDeferredParts(pkg);
    expect(deferredRejection(pkg)).toMatchObject({ reason: 'xml', partName: name });
    const part = pkg.parts.get(name)!;
    if (part.kind !== 'xml') throw new Error('kind');
    expect(part.xml.root.local).toBe('sld');
    expect(() => writePptx(pkg)).toThrow();
  });

  test('a write before every part was read parses the rest and saves what an eager read saves', () => {
    const bytes = deck(4);
    const eager = readPptx(bytes);
    if (!eager.ok) throw new Error('eager');
    expect(writePptx(deferAll(bytes))).toEqual(writePptx(eager.pkg));
  });
});
