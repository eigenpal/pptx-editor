import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import {
  createPptxEditor,
  formatControlMixed,
  formatLineState,
  formatSectionsOf,
  formatSelectionCategory,
  formatSectionOpen,
  setFormatSectionOpen,
  listenFormatSectionExpansion,
  toolbarCommandState,
  type PptxEditorInstance,
} from '../index.ts';

const outline = (end: string) =>
  '<a:ln w="39687"><a:solidFill><a:srgbClr val="123456"/></a:solidFill>' +
  '<a:prstDash val="dashDot"/><a:headEnd type="triangle" w="lg" len="sm" xmlns:x="urn:test" x:tracking="head"><x:kept/></a:headEnd>' +
  `<a:tailEnd type="${end}" w="sm" len="lg"/></a:ln>`;
function deck() {
  const presets = ['line', 'lineInv', 'bentConnector3', 'curvedConnector3', 'rect'];
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: presets
          .map((prst, index) =>
            shape({
              id: index + 2,
              name: prst,
              prst,
              x: 914400,
              y: 914400 * index,
              cx: 914400,
              cy: 200000,
              line: outline(index ? 'diamond' : 'oval'),
            })
          )
          .join(''),
      },
    ],
  });
}
async function open(bytes = deck()) {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}
const ids = (e: PptxEditorInstance) =>
  e
    .session!.layout(0)!
    .items.filter((i) => i.element.layer === 'slide')
    .map((i) => i.element.nodeId);
const xml = (e: PptxEditorInstance) =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0]!.partName)!);

test('all line-shaped presets expose independent common and mixed endpoint values', async () => {
  const e = await open();
  const lines = ids(e).slice(0, 4);
  for (const id of lines) {
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    expect(formatSelectionCategory(e.snapshot())).toBe('line');
    expect(formatLineState(e)?.width).toBe(39687 / 12700);
    expect(formatSectionsOf(e.snapshot())[0]!.id).toBe('line');
  }
  e.setSelection({ kind: 'shapes', nodeIds: lines });
  expect(formatSectionsOf(e.snapshot()).find((s) => s.id === 'line')?.slots).toContain(
    'shape.lineStart'
  );
  expect(formatSectionsOf(e.snapshot()).find((s) => s.id === 'line')?.slots).toContain(
    'shape.lineEnd'
  );
  expect(formatLineState(e)).toMatchObject({
    head: 'triangle',
    headMixed: false,
    tail: null,
    tailMixed: true,
    endsSupported: true,
  });
  expect(formatControlMixed(e, 'shape.lineStart')).toBe(false);
  expect(formatControlMixed(e, 'shape.lineEnd')).toBe(true);
  expect(toolbarCommandState(e, 'shape.lineStart').value).toBe('triangle');
  expect(toolbarCommandState(e, 'shape.lineEnd').value).toBeUndefined();
  expect(e.exec({ type: 'setLineEnds', tail: 'stealth' }).ok).toBe(true);
  expect(formatLineState(e)?.tail).toBe('stealth');
  expect(e.exec({ type: 'undo' }).ok).toBe(true);
  expect(formatLineState(e)?.tailMixed).toBe(true);
  e.destroy();
});

test('mixed line and rectangle selections refuse endpoint edits atomically', async () => {
  const e = await open();
  const all = ids(e);
  e.setSelection({ kind: 'shapes', nodeIds: [all[0]!, all[4]!] });
  const before = xml(e);
  expect(formatSectionsOf(e.snapshot()).find((s) => s.id === 'line')?.slots).not.toContain(
    'shape.lineStart'
  );
  expect(formatLineState(e)?.endsSupported).toBe(false);
  expect(e.can({ type: 'setLineEnds', head: 'diamond' }).ok).toBe(false);
  expect(e.exec({ type: 'setLineEnds', head: 'diamond' }).ok).toBe(false);
  expect(xml(e)).toBe(before);
  e.destroy();
});

test('changing an endpoint type preserves dimensions, extension XML, opposite end, and line attributes', async () => {
  const e = await open();
  e.setSelection({ kind: 'shapes', nodeIds: [ids(e)[0]!] });
  const before = xml(e);
  expect(e.exec({ type: 'setLineEnds', head: 'diamond' }).ok).toBe(true);
  expect(xml(e)).toBe(before.replace('type="triangle"', 'type="diamond"'));
  const saved = e.save()!;
  const reopened = await open(saved);
  reopened.setSelection({ kind: 'shapes', nodeIds: [ids(reopened)[0]!] });
  expect(formatLineState(reopened)).toMatchObject({ head: 'diamond', tail: 'oval' });
  expect(xml(reopened)).toContain('w="lg" len="sm"');
  expect(xml(reopened)).toContain('x:tracking="head"');
  expect(e.exec({ type: 'undo' }).ok).toBe(true);
  expect(xml(e)).toBe(before);
  reopened.destroy();
  e.destroy();
});

test('group children retain endpoint eligibility while the group itself does not', async () => {
  const e = await open();
  const lines = ids(e).slice(0, 2);
  e.setSelection({ kind: 'shapes', nodeIds: lines });
  expect(e.exec({ type: 'group' }).ok).toBe(true);
  const group = e.snapshot().geometry!.nodeId;
  expect(formatSelectionCategory(e.snapshot())).toBe('group');
  expect(formatSectionsOf(e.snapshot()).some((s) => s.id === 'line')).toBe(false);
  e.setSelection({ kind: 'shapes', nodeIds: lines, group });
  expect(formatSectionsOf(e.snapshot()).find((s) => s.id === 'line')?.slots).toContain(
    'shape.lineStart'
  );
  expect(e.exec({ type: 'setLineEnds', head: 'oval' }).ok).toBe(true);
  expect(formatLineState(e)?.head).toBe('oval');
  e.destroy();
});

test('an inherited line keeps its dimensions, extensions, and opposite endpoint after a type edit', async () => {
  const files = unzipSync(deck());
  const original = strFromU8(files['ppt/slides/slide1.xml']!);
  const placeholder = '<p:nvPr><p:ph type="body" idx="8"/></p:nvPr>';
  files['ppt/slideLayouts/slideLayout2.xml'] = strToU8(
    original
      .replace(/<p:sld\b/, '<p:sldLayout')
      .replace('</p:sld>', '</p:sldLayout>')
      .replace('<p:nvPr></p:nvPr>', placeholder)
  );
  files['ppt/slides/slide1.xml'] = strToU8(
    original
      .replace('<p:nvPr></p:nvPr>', placeholder)
      .replace(outline('oval'), '')
      .replace(/<a:prstGeom\b.*?<\/a:prstGeom>/, '')
  );
  const e = await open(zipSync(files));
  e.setSelection({ kind: 'shapes', nodeIds: [ids(e)[0]!] });
  const source = serializeXmlPart(
    xmlPartOf(e.session!.store.package, '/ppt/slideLayouts/slideLayout2.xml')!
  );
  expect(formatSelectionCategory(e.snapshot())).toBe('line');
  expect(e.exec({ type: 'setLineEnds', head: 'stealth' }).ok).toBe(true);
  expect(formatLineState(e)).toMatchObject({ head: 'stealth', tail: 'oval', width: 39687 / 12700 });
  expect(xml(e)).toContain('type="stealth" w="lg" len="sm"');
  expect(xml(e)).toContain('x:tracking="head"');
  expect(
    serializeXmlPart(xmlPartOf(e.session!.store.package, '/ppt/slideLayouts/slideLayout2.xml')!)
  ).toBe(source);
  e.destroy();
});

test('disclosures remember editor and category preferences without writing document state', async () => {
  const e = await open(),
    other = await open();
  const before = xml(e);
  let calls = 0;
  const stop = listenFormatSectionExpansion(e, () => {
    calls += 1;
  });
  expect(formatSectionOpen(e, 'line', 'line')).toBe(true);
  expect(formatSectionOpen(e, 'line', 'geometry')).toBe(false);
  setFormatSectionOpen(e, 'line', 'geometry', true);
  expect(formatSectionOpen(e, 'line', 'geometry')).toBe(true);
  expect(formatSectionOpen(e, 'shape', 'geometry')).toBe(false);
  expect(formatSectionOpen(other, 'line', 'geometry')).toBe(false);
  expect(calls).toBe(1);
  stop();
  setFormatSectionOpen(e, 'line', 'geometry', false);
  expect(calls).toBe(1);
  expect(xml(e)).toBe(before);
  e.destroy();
  other.destroy();
});
