// Computed document refresh changes (review/refresh-changes.ts) over random edits: every change
// id is unique, and the editor's own check accepts every computed location. The decks hold
// repeated shape ids, groups, wrapped shapes, several paragraphs, emoji, notes, hidden slides,
// and repeated creation ids. The seed is fixed, so a failure repeats.

import { expect, test } from 'bun:test';
import { readPptx, type PptxPackage } from '@pptx-editor.dev/core/store';
import { resolveRefreshChanges } from '../../../core/src/editor/document-refresh-changes.ts';
import { buildPptx, group, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { computeRefreshChanges } from '../review/refresh-changes.ts';

interface ShapeModel {
  kind: 'box' | 'group' | 'wrapped';
  id: number;
  childId: number;
  x: number;
  paras: string[];
}
interface SlideModel {
  shapes: ShapeModel[];
  notes?: string;
  hidden: boolean;
  cid?: number;
}

/** A small seeded generator (mulberry32). */
function random(seed: number) {
  let s = seed >>> 0;
  const next = () => {
    s = (s + 0x6d2b79f5) >>> 0;
    let t = s;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  const int = (n: number) => Math.floor(next() * n);
  return { next, int, pick: <T>(xs: readonly T[]): T => xs[int(xs.length)]! };
}
type Rng = ReturnType<typeof random>;

const WORDS = ['alpha', 'beta', '$4.0M', 'Hello,', '😀', '東京都', 'x😀y', 'naïve', '', '42%'];
const sentence = (r: Rng) =>
  Array.from({ length: 1 + r.int(4) }, () => r.pick(WORDS))
    .join(' ')
    .trim();

function newShape(r: Rng): ShapeModel {
  return {
    kind: r.pick(['box', 'box', 'box', 'group', 'wrapped'] as const),
    id: 2 + r.int(6),
    childId: 2 + r.int(8),
    x: r.int(5) * 100_000,
    paras: Array.from({ length: 1 + r.int(3) }, () => sentence(r)),
  };
}

function newSlide(r: Rng): SlideModel {
  return {
    shapes: Array.from({ length: 1 + r.int(4) }, () => newShape(r)),
    ...(r.next() < 0.3 ? { notes: sentence(r) } : {}),
    hidden: r.next() < 0.1,
    ...(r.next() < 0.5 ? { cid: 100 + r.int(4) } : {}),
  };
}

const clone = (slides: readonly SlideModel[]): SlideModel[] =>
  slides.map((s) => ({ ...s, shapes: s.shapes.map((sh) => ({ ...sh, paras: [...sh.paras] })) }));

function mutate(r: Rng, slides: SlideModel[]): void {
  const slide = slides.length > 0 ? r.pick(slides) : undefined;
  const sh = slide && slide.shapes.length > 0 ? r.pick(slide.shapes) : undefined;
  switch (r.int(12)) {
    case 0:
    case 1:
    case 2:
      if (sh && sh.paras.length > 0) {
        const p = r.int(sh.paras.length);
        const words = sh.paras[p]!.split(' ');
        words[r.int(words.length)] = r.pick(WORDS);
        sh.paras[p] = words.join(' ');
      }
      break;
    case 3:
      if (sh) sh.paras.splice(r.int(sh.paras.length + 1), 0, sentence(r));
      break;
    case 4:
      if (sh && sh.paras.length > 1) sh.paras.splice(r.int(sh.paras.length), 1);
      break;
    case 5:
      slide?.shapes.splice(r.int(slide.shapes.length + 1), 0, newShape(r));
      break;
    case 6:
      if (slide && slide.shapes.length > 0) slide.shapes.splice(r.int(slide.shapes.length), 1);
      break;
    case 7:
      if (sh) sh.x += 50_000;
      break;
    case 8:
      if (slide) slide.hidden = !slide.hidden;
      break;
    case 9:
      if (slide) {
        if (slide.notes === undefined || r.next() < 0.5) slide.notes = sentence(r);
        else delete slide.notes;
      }
      break;
    case 10:
      slides.splice(
        r.int(slides.length + 1),
        0,
        r.next() < 0.5 || !slide ? newSlide(r) : clone([slide])[0]!
      );
      break;
    default:
      if (slides.length > 1) {
        const [moved] = slides.splice(r.int(slides.length), 1);
        if (r.next() < 0.5) slides.splice(r.int(slides.length + 1), 0, moved!);
      }
  }
}

const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const creation = (v: number) =>
  `<p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="${v}"/></p:ext></p:extLst>`;

function shapeXml(s: ShapeModel): string {
  const text = s.paras.map((p) => para(run(p))).join('');
  const box = (id: number) => shape({ id, x: s.x, y: 1, cx: 100, cy: 100, text });
  if (s.kind === 'group')
    return group({ id: s.id, x: s.x, y: 0, cx: 100, cy: 100 }, box(s.childId));
  if (s.kind === 'wrapped')
    return `<mc:AlternateContent xmlns:mc="${NS_MC}"><mc:Choice Requires="p14">${box(s.id)}</mc:Choice><mc:Fallback>${box(s.id)}</mc:Fallback></mc:AlternateContent>`;
  return box(s.id);
}

function bytesOf(slides: readonly SlideModel[]): Uint8Array {
  return buildPptx({
    slides: slides.map((s) => ({
      shapes: s.shapes.map(shapeXml).join(''),
      hidden: s.hidden,
      ...(s.notes !== undefined ? { notes: s.notes } : {}),
      ...(s.cid !== undefined ? { ext: creation(s.cid) } : {}),
    })),
  });
}

const pkg = (bytes: Uint8Array): PptxPackage => {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('fixture did not open');
  return r.pkg;
};

test('random edits give unique ids and locations that the editor accepts', () => {
  const r = random(20261116);
  for (let round = 0; round < 150; round += 1) {
    const before = Array.from({ length: 1 + r.int(5) }, () => newSlide(r));
    const after = clone(before);
    for (let k = 1 + r.int(4); k > 0; k -= 1) mutate(r, after);
    const returned = pkg(bytesOf(after));
    const { changes } = computeRefreshChanges(pkg(bytesOf(before)), returned);
    expect(new Set(changes.map((c) => c.id)).size).toBe(changes.length);
    const located = resolveRefreshChanges(returned, 'r', changes, new Map(), () => true);
    for (const l of located) {
      const status = l.change.status;
      const located = changes.find((c) => c.id === l.change.id)!.location !== undefined;
      // A located change is available; one without a location is deleted or unavailable.
      expect([round, l.change.id, status, l.change.diagnostic?.code]).toEqual([
        round,
        l.change.id,
        located ? 'available' : status,
        located ? undefined : l.change.diagnostic?.code,
      ]);
    }
  }
});
