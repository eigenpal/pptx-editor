import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx, para, run, shape } from '../fixtures/builder.ts';
import { parseXmlPart, readPptx, type Element } from '../../packages/core/src/store/index.ts';
import {
  censusDeck,
  EDIT_MOVE_EMU,
  editResult,
  firstDifference,
  LIST_LIMIT,
  MachineFault,
  KNOWN_TYPES,
  packageDifferences,
  partType,
  TYPE_PREFIXES,
  typeLabel,
} from './roundtrip-census.ts';

const IN = 914_400;

const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        shapes:
          shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: IN, text: para(run('Private words')) }) +
          shape({ id: 3, x: IN, y: 3 * IN, cx: IN, cy: IN }),
      },
    ],
  });

const root = (xml: string): Element => {
  const r = parseXmlPart('/ppt/slides/slide1.xml', xml);
  if (!r.ok) throw new Error(r.reason);
  return r.part.root;
};

describe('round-trip census', () => {
  test('a clean deck round-trips, with and without the scripted edit', async () => {
    const kept: string[] = [];
    const r = await censusDeck(deck(), { keep: (kind) => kept.push(kind) });
    expect(r.error).toBeNull();
    expect(r.slides).toBe(1);
    expect(r.roundTrip).toMatchObject({
      fingerprintEqual: true,
      digestStable: true,
      differences: { count: 0 },
      namespaceDifferences: { count: 0 },
    });
    expect(r.edit?.edits).toEqual([
      { edit: 'setText', result: 'applied' },
      { edit: 'move', result: 'applied' },
    ]);
    expect(r.edit).toMatchObject({
      editedParts: 1,
      fingerprintEqual: true,
      lost: 0,
      undoRestores: true,
    });
    expect(r.roundTrip?.xmlParts).toBeGreaterThan(5);
    expect((r.roundTrip?.xmlParts ?? 0) + (r.roundTrip?.binaryParts ?? 0)).toBe(r.parts!);
    expect(r.edit?.collateral.count).toBe(0);
    expect(kept).toEqual(['saved', 'edited']);
  });

  test('the result holds no deck text', async () => {
    const r = await censusDeck(deck());
    expect(JSON.stringify(r)).not.toContain('Private');
  });

  test('a deck without runs or transforms reports no target, not a pass', async () => {
    const r = await censusDeck(buildPptx({ slides: [{ shapes: '', layout: 'blank' }] }));
    expect(r.edit?.edits.map((e) => e.result)).toEqual(['no-target', 'no-target']);
    expect(r.edit?.editedParts).toBe(0);
    expect(r.edit?.fingerprintEqual).toBeNull();
  });

  test('bytes that are not a deck fail at open', async () => {
    const r = await censusDeck(new TextEncoder().encode('not a zip'));
    expect(r.error).toMatchObject({ stage: 'open', reason: 'zip' });
    expect(r.roundTrip).toBeNull();
  });

  test('the first difference names elements and attributes, never text', () => {
    const ns = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    const a = root(`<a:p ${ns}><a:r><a:rPr lang="en-US"/><a:t>one</a:t></a:r></a:p>`);
    const b = root(`<a:p ${ns}><a:r><a:rPr lang="de-DE"/><a:t>one</a:t></a:r></a:p>`);
    const c = root(`<a:p ${ns}><a:r><a:rPr lang="en-US"/><a:t>two</a:t></a:r></a:p>`);
    const d = root(`<a:p ${ns}><a:r><a:rPr lang="en-US"/><a:t>one</a:t></a:r><a:br/></a:p>`);
    expect(firstDifference(a, a)).toBeUndefined();
    expect(firstDifference(a, b)).toEqual({ path: 'a:p/a:r/a:rPr', what: 'attribute lang value' });
    expect(firstDifference(a, c)).toEqual({ path: 'a:p/a:r/a:t', what: 'text' });
    expect(firstDifference(a, d)).toEqual({ path: 'a:p', what: 'child added a:br' });
  });

  test('a changed child is found before a later added or removed child', () => {
    const ns = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    const r = (lang: string): string => `<a:r><a:rPr lang="${lang}"/><a:t>x</a:t></a:r>`;
    const a = root(`<a:p ${ns}>${r('en-US')}${r('en-US')}<a:br/></a:p>`);
    const b = root(`<a:p ${ns}>${r('de-DE')}${r('en-US')}</a:p>`);
    expect(firstDifference(a, b)).toEqual({ path: 'a:p/a:r/a:rPr', what: 'attribute lang value' });
    // A reordering names the element whose count changed, not the first one out of place.
    const c = root(`<a:p ${ns}>${r('en-US')}<a:br/></a:p>`);
    const d = root(`<a:p ${ns}><a:br/>${r('en-US')}${r('en-US')}</a:p>`);
    expect(firstDifference(c, d)).toEqual({ path: 'a:p', what: 'child added a:r' });
  });

  test('only schema names reach the output', () => {
    const ns = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:x"';
    const b = root(`<a:p ${ns}/>`);
    const named = (inner: string): string | undefined =>
      firstDifference(root(`<a:p ${ns}>${inner}</a:p>`), b)?.what;
    expect(named('<v:ProjectName/>')).toBe('child removed x:?');
    expect(named('<a:ProjectFalcon/>')).toBe('child removed a:?');
    expect(named('<Codename xmlns=""/>')).toBe('child removed x:?');
    expect(named('<a:br/>')).toBe('child removed a:br');
    const c = root(`<a:p ${ns} secretWord="1"/>`);
    expect(firstDifference(c, b)?.what).toBe('attribute ? removed');
  });

  test('a pass-1 difference in the edited part is not counted again in pass 2', async () => {
    const one = buildPptx({
      slides: [
        {
          shapes:
            shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, text: para(run('first')) }) +
            shape({ id: 3, x: IN, y: 2 * IN, cx: IN, cy: IN, text: para(run('MARK')) }),
        },
      ],
    });
    const files = unzipSync(one);
    const name = 'ppt/slides/slide1.xml';
    files[name] = strToU8(strFromU8(files[name]).replace('>MARK<', '>a<![CDATA[b]]>c<'));
    const r = await censusDeck(zipSync(files));
    expect(r.roundTrip?.differences.count).toBe(1);
    expect(r.edit).toMatchObject({
      editedParts: 1,
      fingerprintEqual: true,
      lost: 0,
      collateral: { count: 0 },
    });
  });

  test('a fault of the machine while a copy is kept goes up to the caller', async () => {
    const keep = (kind: 'saved' | 'edited'): void => {
      if (kind === 'edited') throw new MachineFault('cannot write a kept copy (ENOSPC)');
    };
    await expect(censusDeck(deck(), { keep })).rejects.toBeInstanceOf(MachineFault);
  });

  test('a refused edit names the store code and leaves the package as it was', async () => {
    const r = await censusDeck(deck(), { editText: '\u0001' });
    expect(r.edit?.edits[0]).toEqual({ edit: 'setText', result: 'refused', code: 'invalid-text' });
    expect(r.edit?.edits[1].result).toBe('applied');
    expect(r.edit?.undoRestores).toBe(true);
  });

  test('pass 2 names edits not kept, lost edits, and collateral parts', async () => {
    const pkg = (text: string, other: string): ReturnType<typeof readPptx> =>
      readPptx(
        buildPptx({
          slides: [
            { shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, text: para(run(text)) }) },
            { shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, text: para(run(other)) }) },
          ],
        })
      );
    const before = pkg('one', 'two');
    const edited = pkg('edit', 'two');
    const keptCopy = pkg('edit', 'two');
    const lostCopy = pkg('one', 'changed');
    if (!before.ok || !edited.ok || !keptCopy.ok || !lostCopy.ok) throw new Error('fixture');
    const touched = new Set(['/ppt/slides/slide1.xml']);
    expect(editResult(before.pkg, edited.pkg, keptCopy.pkg, touched)).toMatchObject({
      fingerprintEqual: true,
      lost: 0,
      collateral: { count: 0 },
    });
    const bad = editResult(before.pkg, edited.pkg, lostCopy.pkg, touched);
    expect(bad).toMatchObject({ fingerprintEqual: false, lost: 1, collateral: { count: 1 } });
    expect(bad.differences.items[0]).toMatchObject({ type: 'pml:slide', what: 'text' });
    // A collateral item names the element path and what differs, as a pass-1 item does.
    expect(bad.collateral.items[0]).toMatchObject({
      type: 'pml:slide',
      change: 'changed',
      what: 'text',
    });
    expect(bad.collateral.items[0]?.path).toBeTruthy();
  });

  test('a list holds at most LIST_LIMIT items, and the count holds them all', () => {
    const slides = (text: string): Uint8Array =>
      buildPptx({
        slides: Array.from({ length: LIST_LIMIT + 5 }, () => ({
          shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, text: para(run(text)) }),
        })),
      });
    const a = readPptx(slides('a'));
    const b = readPptx(slides('b'));
    if (!a.ok || !b.ok) throw new Error('fixture');
    const d = packageDifferences(a.pkg, b.pkg);
    expect(d.count).toBe(LIST_LIMIT + 5);
    expect(d.items).toHaveLength(LIST_LIMIT);
  });

  test('a slide that does not parse fails at open with its part type', async () => {
    const files = unzipSync(deck());
    files['ppt/slides/slide1.xml'] = strToU8('<p:sld');
    const r = await censusDeck(zipSync(files));
    expect(r.error).toMatchObject({ stage: 'open', reason: 'xml', partType: 'ppt/slides' });
  });

  test('a pass-1 difference is not counted again in pass 2', async () => {
    const two = buildPptx({
      slides: [
        { shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, text: para(run('first')) }) },
        { shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, text: para(run('MARK')) }) },
      ],
    });
    // A CDATA section splits the text into three nodes; the saved copy has one.
    const files = unzipSync(two);
    const name = 'ppt/slides/slide2.xml';
    files[name] = strToU8(strFromU8(files[name]).replace('>MARK<', '>a<![CDATA[b]]>c<'));
    const r = await censusDeck(zipSync(files));
    expect(r.roundTrip?.differences.count).toBe(1);
    expect(r.edit?.edits.map((e) => e.result)).toEqual(['applied', 'applied']);
    expect(r.edit).toMatchObject({ fingerprintEqual: true, lost: 0, collateral: { count: 0 } });
  });

  test('the move takes the first group or graphic frame, and stays inside the coordinate range', async () => {
    const group = (x: string): string =>
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="1" cy="1"/><a:chOff x="0" y="0"/>` +
      `<a:chExt cx="1" cy="1"/></a:xfrm></p:grpSpPr></p:grpSp>`;
    const frame =
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="F"/><p:cNvGraphicFramePr/><p:nvPr/>` +
      `</p:nvGraphicFramePr><p:xfrm><a:off x="27273042316900" y="0"/><a:ext cx="1" cy="1"/></p:xfrm>` +
      `<a:graphic><a:graphicData uri="urn:x"/></a:graphic></p:graphicFrame>`;
    const moved = async (shapes: string): Promise<string | undefined> => {
      let edited: Uint8Array | undefined;
      const r = await censusDeck(buildPptx({ slides: [{ shapes }] }), {
        keep: (kind, bytes) => {
          if (kind === 'edited') edited = bytes;
        },
      });
      if (r.edit?.edits[1].result !== 'applied' || !edited) return undefined;
      const xml = strFromU8(unzipSync(edited)['ppt/slides/slide1.xml']);
      return /name="[GF]".*?<a:off x="(-?\d+)"/s.exec(xml)?.[1];
    };
    expect(await moved(group('100'))).toBe(String(100 + EDIT_MOVE_EMU));
    expect(await moved(frame)).toBe(String(27_273_042_316_900 - EDIT_MOVE_EMU));
    // A value that is not an ST_Coordinate integer is never a move target.
    expect(await moved(group('0x10'))).toBeUndefined();
  });

  test('prefix and declaration changes count only in namespace mode', () => {
    const a = root('<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>');
    const b = root('<x:p xmlns:x="http://schemas.openxmlformats.org/drawingml/2006/main"/>');
    expect(firstDifference(a, b)).toBeUndefined();
    expect(firstDifference(a, b, true)).toEqual({ path: 'a:p', what: 'element prefix' });
  });

  test('part types come from content types, after a fixed prefix', () => {
    const r = readPptx(deck());
    if (!r.ok) throw new Error('fixture');
    expect(partType(r.pkg, '/ppt/slides/slide1.xml')).toBe('pml:slide');
    expect(partType(r.pkg, '/ppt/slides/_rels/slide1.xml.rels')).toBe('rels:pml:slide');
    expect(partType(r.pkg, '/_rels/.rels')).toBe('rels:package');
    expect(partType(r.pkg, '/[Content_Types].xml')).toBe('content-types');
    expect(typeLabel('application/vnd.openxmlformats-officedocument.theme+xml')).toBe(
      'office:theme'
    );
    expect(typeLabel('image/png')).toBe('image/png');
    expect(typeLabel('text/secret words')).toBe('other');
    expect(typeLabel('image/secret words')).toBe('image/other');
    // Only listed types keep their name: a deck can declare any word as a content type.
    expect(typeLabel('application/x-projectfalcon')).toBe('application/other');
    expect(typeLabel('application/vnd.ms-ignore-previous-instructions')).toBe('ms:other');
    expect(typeLabel(undefined)).toBe('no-content-type');
    // Every listed type keeps its listed spelling, in any case.
    for (const name of KNOWN_TYPES.values()) {
      const [prefix, label] = TYPE_PREFIXES.find(([, l]) => name.startsWith(l))!;
      const type = `${prefix}${name.slice(label.length)}`;
      expect(typeLabel(type)).toBe(name);
      expect(typeLabel(`${type.toLowerCase()}+xml`)).toBe(name);
    }
    expect(
      typeLabel('application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml')
    ).toBe('pml:slideLayout');
  });
});
