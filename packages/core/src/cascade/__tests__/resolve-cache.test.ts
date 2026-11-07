// The resolve caches keep unchanged elements by identity. An edit that changes what an element
// reads outside its own XML (here, the slide's color map) must still resolve it again: the
// incremental result equals a fresh parse of the same package.

import { expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { PptxDocumentStore, readPptx, writePptx, xmlPartOf } from '../../store/index.ts';
import { resolveSlide } from '../index.ts';

interface Node {
  readonly id?: string;
  readonly local?: string;
  readonly children?: readonly Node[];
}

function find(el: Node | undefined, local: string): Node | undefined {
  if (!el || typeof el !== 'object') return undefined;
  if (el.local === local) return el;
  for (const c of el.children ?? []) {
    const found = find(c, local);
    if (found) return found;
  }
  return undefined;
}

const text = (v: unknown) => JSON.stringify(v, (_k, x) => (x instanceof Map ? [...x] : x));

const SHAPE =
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
  '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"><a:solidFill>' +
  '<a:schemeClr val="tx1"/></a:solidFill></a:rPr><a:t>Hi</a:t></a:r></a:p></p:txBody></p:sp>';
const OVERRIDE =
  '<a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" ' +
  'accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" ' +
  'accent6="accent6" hlink="hlink" folHlink="folHlink"/>';

test('a color map edit reaches the cached text of the slide', () => {
  const opened = readPptx(
    buildPptx({ slides: [{ shapes: SHAPE, layout: 'blank', clrMapOvr: OVERRIDE }] })
  );
  if (!opened.ok) throw new Error('open failed');
  const store = new PptxDocumentStore(opened.pkg);
  const before = text(resolveSlide(store.package, 0)!.elements.at(-1));
  const part = resolveSlide(store.package, 0)!.partName;
  const mapping = find(xmlPartOf(store.package, part)!.root as Node, 'overrideClrMapping')!;
  const edit = { ns: '', local: 'tx1' };
  expect(
    store.transact([{ op: 'setAttribute', part, node: mapping.id!, name: edit, value: 'lt1' }]).ok
  ).toBe(true);
  const after = text(resolveSlide(store.package, 0)!.elements.at(-1));
  const reopened = readPptx(writePptx(store.package));
  if (!reopened.ok) throw new Error('reopen failed');
  expect(after).not.toBe(before);
  expect(after).toBe(text(resolveSlide(reopened.pkg, 0)!.elements.at(-1)));
});
