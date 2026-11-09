// Seeded fuzz of cut: random text boxes, groups, selections, keyboard and menu cuts, pastes,
// undo, and redo on a slide with a chart that the clipboard cannot carry. After each step:
//
// - the chart is still on the slide, at any group depth (no cut ever loses it);
// - a cut that went through removed exactly the shapes that its copy carried;
// - a refused cut changed nothing.
//
// Knobs: CUT_FUZZ_SEEDS (default 24), CUT_FUZZ_FROM (default 1), CUT_FUZZ_STEPS (default 30).

import { afterAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync, zipSync } from 'fflate';
import { GlobalWindow } from 'happy-dom';
import { runClipboardSlot } from '../chrome-actions.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { xmlPartOf, type Element } from '../../store/index.ts';
import type { ClipboardNode, ClipboardPayload } from '../../contracts/index.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const knob = (name: string, fallback: number): number => {
  const v = Number(process.env[`CUT_FUZZ_${name}`]);
  return Number.isInteger(v) && v > 0 ? v : fallback;
};
const SEEDS = knob('SEEDS', 24);
const FROM = knob('FROM', 1);
const STEPS = knob('STEPS', 30);

const PACKAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';
const deckBytes = (() => {
  const files = unzipSync(
    new Uint8Array(readFileSync(join(import.meta.dir, '../../../../../fixtures/charts.pptx')))
  );
  files['ppt/charts/_rels/chart1.xml.rels'] = new TextEncoder().encode(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdBook" Type="${PACKAGE}" Target="file:///book.xlsx" TargetMode="External"/>` +
      '</Relationships>'
  );
  return zipSync(files);
})();

/** A small deterministic generator (mulberry32). */
function random(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** The `cNvPr` names of every shape on slide 1, at any group depth, sorted. */
function allNames(e: PptxEditorInstance): string[] {
  const part = e.session!.info.slides[0]!.partName;
  const out: string[] = [];
  const walk = (el: Element): void => {
    for (const c of el.children) {
      if (c.type !== 'element') continue;
      if (c.local === 'cNvPr') {
        const name = c.attrs.find((a) => a.local === 'name')?.value;
        if (name) out.push(name);
      }
      walk(c);
    }
  };
  walk(xmlPartOf(e.session!.store.package, part)!.root);
  return out.filter(Boolean).sort();
}

/** The names that `before` holds more often than `after`. */
function removed(before: string[], after: string[]): string[] {
  const left = [...after];
  const out: string[] = [];
  for (const n of before) {
    const i = left.indexOf(n);
    if (i >= 0) left.splice(i, 1);
    else out.push(n);
  }
  return out.sort();
}

/** The names of every shape in a clipboard payload, at any group depth, sorted. */
function payloadNames(payload: ClipboardPayload): string[] {
  const out: string[] = [];
  const walk = (n: ClipboardNode): void => {
    if (n.local === 'cNvPr') {
      const name = n.attrs?.find((a) => a.local === 'name')?.value;
      if (name) out.push(name);
    }
    for (const c of n.children ?? []) if (!('text' in c)) walk(c);
  };
  for (const s of payload.shapes) walk(s);
  return out.sort();
}

function nativeCut(host: HTMLElement): boolean {
  const data = new Map<string, string>();
  const event = new win.Event('cut', { bubbles: true, cancelable: true }) as unknown as Event;
  Object.defineProperty(event, 'clipboardData', {
    value: {
      setData: (t: string, v: string) => void data.set(t, v),
      getData: (t: string) => data.get(t) ?? '',
      items: [],
      files: [],
      types: [],
    },
  });
  (host.querySelector('.pptx-surface__input') as HTMLElement).dispatchEvent(event);
  return data.size > 0;
}

async function runSeed(seed: number): Promise<void> {
  const rand = random(seed);
  const pick = <T>(xs: readonly T[]): T => xs[Math.floor(rand() * xs.length)]!;
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(deckBytes)).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  e.attach(host);
  const trail: string[] = [];
  try {
    for (let step = 0; step < STEPS; step++) {
      const top = (e.snapshot().shapes ?? []).map((s) => s.nodeId);
      const action = pick([
        'textBox',
        'select',
        'select',
        'group',
        'ungroup',
        'cut',
        'cut',
        'menuCut',
        'paste',
        'undo',
        'redo',
      ] as const);
      trail.push(action);
      if (action === 'textBox') e.exec({ type: 'insertTextBox' });
      else if (action === 'select' && top.length) {
        const chosen = top.filter(() => rand() < 0.5);
        e.setSelection({ kind: 'shapes', nodeIds: chosen.length ? chosen : [pick(top)] });
      } else if (action === 'group') e.exec({ type: 'group' });
      else if (action === 'ungroup') e.exec({ type: 'ungroup' });
      else if (action === 'paste' && e.lastCopied) e.exec({ type: 'paste', payload: e.lastCopied });
      else if (action === 'undo') e.exec({ type: 'undo' });
      else if (action === 'redo') e.exec({ type: 'redo' });
      else if (action === 'cut' || action === 'menuCut') {
        const sel = e.snapshot().selection;
        if (sel.kind !== 'shapes' || !sel.nodeIds.length) continue;
        const before = allNames(e);
        const revision = e.session!.store.revision;
        const copied = e.lastCopied;
        let wrote: boolean;
        if (action === 'cut') wrote = nativeCut(host);
        else {
          runClipboardSlot(e, 'edit.cut', host.querySelector('.pptx-surface__input'));
          wrote = e.lastCopied !== copied;
        }
        const gone = removed(before, allNames(e));
        if (gone.length) {
          // It went through: what it removed is exactly what it copied.
          expect(e.lastCopied).not.toBe(copied);
          expect(gone).toEqual(payloadNames(e.lastCopied!));
        } else if (action === 'cut') {
          // Refused: the store and the clipboard are as they were.
          expect(e.session!.store.revision).toBe(revision);
          expect(wrote).toBe(false);
          expect(e.lastCopied).toBe(copied);
        }
      }
      expect(allNames(e)).toContain('Chart 1');
    }
  } catch (error) {
    throw new Error(`cut fuzz seed ${seed} failed after ${trail.join(', ')}: ${String(error)}`);
  } finally {
    e.destroy();
  }
}

describe('cut fuzz', () => {
  for (let seed = FROM; seed < FROM + SEEDS; seed++)
    test(`seed ${seed}: no cut loses a shape that its copy did not carry`, () => runSeed(seed));
});
