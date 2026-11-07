// The placeholder match of a layout change takes linear time and gives the same result as a
// scan of the layout places in order: random slides and layouts, and one with 24000 places.

import { expect, test } from 'bun:test';
import { mapPlaceholders } from '../commands/slide-layouts.ts';
import { parseXmlPart, type Element } from '../../store/package/tree.ts';

const TYPES = ['title', 'ctrTitle', 'body', 'subTitle', 'obj', 'pic', 'tbl', 'dt', 'ftr', 'sldNum'];
const TITLES = new Set(['title', 'ctrTitle']);
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

interface Place {
  readonly el: Element;
  readonly ph: Element;
  readonly type: string;
  readonly idx: string | undefined;
}

/** Shapes with placeholders of the given types and indexes; `text` gives them text. */
function places(specs: readonly { type: string; idx?: string; text: boolean }[]): Place[] {
  const shapes = specs
    .map(
      (s, i) =>
        `<p:sp><p:nvSpPr><p:cNvPr id="${i + 2}" name="S"/><p:cNvSpPr/><p:nvPr>` +
        `<p:ph type="${s.type}"${s.idx === undefined ? '' : ` idx="${s.idx}"`}/></p:nvPr></p:nvSpPr>` +
        `<p:spPr/><p:txBody><a:bodyPr/><a:p>${s.text ? '<a:r><a:t>x</a:t></a:r>' : ''}</a:p></p:txBody></p:sp>`
    )
    .join('');
  const r = parseXmlPart(
    '/ppt/slides/slide1.xml',
    `<p:spTree xmlns:p="${NS_P}" xmlns:a="${NS_A}">${shapes}</p:spTree>`
  );
  if (!r.ok) throw new Error('parse');
  const els = r.part.root.children.filter((c): c is Element => c.type === 'element');
  return els.map((el, i) => {
    const ph = (el.children[0] as Element).children
      .filter((c): c is Element => c.type === 'element')
      .find((c) => c.local === 'nvPr')!
      .children.find((c): c is Element => c.type === 'element')!;
    return { el, ph, type: specs[i]!.type, idx: specs[i]!.idx };
  });
}

/** The match as an in-order scan of the layout, per rule: the reference. */
function reference(
  slide: readonly Place[],
  layout: readonly Place[],
  fits: (s: Place, t: Place) => boolean
) {
  const out = new Map<Place, Place>();
  const used = new Set<Place>();
  const text = (s: Place) => s.el.children.length > 0 && JSON.stringify(s.el).includes('"x"');
  const pass = (
    rule: (s: Place, t: Place) => boolean,
    which: (s: Place) => boolean = () => true
  ) => {
    for (const s of slide) {
      if (out.has(s) || !which(s)) continue;
      const t = layout.find((l) => !used.has(l) && fits(s, l) && rule(s, l));
      if (t) {
        out.set(s, t);
        used.add(t);
      }
    }
  };
  pass((s, t) => s.type === t.type && s.idx === t.idx);
  pass((s, t) => TITLES.has(s.type) && TITLES.has(t.type));
  pass((s, t) => s.idx !== undefined && s.idx !== '0' && s.idx === t.idx);
  pass((s, t) => s.type === t.type, text);
  pass(() => true, text);
  pass(
    (s, t) => s.type === t.type,
    (s) => !text(s)
  );
  pass(
    () => true,
    (s) => !text(s)
  );
  return out;
}

function random(seed: number) {
  let x = seed;
  return (n: number) => {
    x = (x * 1103515245 + 12345) & 0x7fffffff;
    return x % n;
  };
}

const spec = (r: (n: number) => number) => ({
  type: TYPES[r(TYPES.length)]!,
  ...(r(3) ? { idx: String(r(6)) } : {}),
  text: r(2) === 0,
});

test('the match equals an in-order scan of the layout', () => {
  for (let seed = 1; seed <= 400; seed += 1) {
    const r = random(seed);
    const slide = places(Array.from({ length: 1 + r(12) }, () => spec(r)));
    const layout = places(Array.from({ length: 1 + r(12) }, () => spec(r)));
    const got = mapPlaceholders(slide, layout);
    // The engine's own `fits` decides, so the reference takes the pairs the engine allows:
    // a pair is allowed when the engine maps that type pair for a lone placeholder.
    const allowed = (s: Place, t: Place) => {
      const [one] = places([{ type: s.type, text: JSON.stringify(s.el).includes('"x"') }]);
      const [place] = places([{ type: t.type }].map((p) => ({ ...p, text: false })));
      return mapPlaceholders([one!], [place!]).size === 1;
    };
    const want = reference(slide, layout, allowed);
    const index = (m: Map<Place, Place>) =>
      slide.map((s) => (m.has(s) ? layout.indexOf(m.get(s)!) : -1));
    expect(index(got)).toEqual(index(want));
  }
});

test('24000 places that match nothing take linear time', () => {
  const slide = places(
    Array.from({ length: 24_000 }, (_, i) => ({ type: 'body', idx: `s${i}`, text: true }))
  );
  const layout = places(
    Array.from({ length: 24_000 }, (_, i) => ({ type: 'dt', idx: `l${i}`, text: false }))
  );
  const t0 = performance.now();
  expect(mapPlaceholders(slide, layout).size).toBe(0);
  expect(performance.now() - t0).toBeLessThan(500);
});
