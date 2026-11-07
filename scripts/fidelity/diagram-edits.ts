// The diagram edits: the save evidence of the diagram row's own edits. Each recipe types in a
// diagram shape, or adds or deletes a diagram shape, through the public editor API, saves, and
// opens the saved copy again. Its check reads both parts the edit writes: the data model
// (`dgm:dataModel`), which a presentation application lays the diagram out from, and the saved
// drawing (`dsp:drawing`), which the editor draws. Every connection of the data model must name
// points that exist, and every drawing shape must name a point. `FEATURE_EDITS['objects.diagrams']`
// makes a text edit, an add, and a delete at once; the recipes here make one edit each.
//
// `diagram-edits.test.ts` runs every recipe. `bun scripts/fidelity/diagram-edits.ts --out <dir>`
// writes each edited copy as `<id>.pptx`, to check for a repair prompt (CONTRIBUTING.md,
// "Fidelity harness").

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import { bytesText } from './chart-data-edits.ts';
import {
  all,
  has,
  lacks,
  runFeatureEdit,
  type EditTools,
  type FeatureEdit,
  type PartRule,
  type SavedDeck,
} from './feature-edit-kit.ts';

/** The parts a diagram edit may change: the data model and the saved drawing parts. */
export const DIAGRAM_PARTS: readonly PartRule[] = [/^\/ppt\/diagrams\/(data|drawing)\d+\.xml$/];

/** The current slide's diagram frame, and the text each of its node shapes shows by node. */
function diagramOf(t: EditTools): { frame: string; texts: Map<string, string> } {
  const e = t.editor;
  const layout = e.session!.layout(e.snapshot().currentSlide);
  const frame = layout?.items.find((i) => i.kind === 'frame' && i.drawing);
  if (frame?.kind !== 'frame') throw new Error('no diagram on the current slide');
  const out = new Map<string, string>();
  const walk = (items: readonly any[]): void => {
    for (const it of items) {
      const el = it.element;
      if (el.kind === 'shape' && el.diagramNode)
        out.set(
          el.diagramNode,
          (el.text?.paragraphs ?? [])
            .map((p: any) => (p.runs ?? []).map((r: any) => r.text ?? '').join(''))
            .join('\n')
        );
      if (it.children) walk(it.children);
    }
  };
  walk(frame.drawing ?? []);
  return { frame: frame.element.nodeId, texts: out };
}

/** Put the caret in the diagram shape that shows `text`, at `offset` (default: its end). */
export function caretInDiagram(t: EditTools, text: string, offset = text.length): void {
  const { frame, texts } = diagramOf(t);
  const node = [...texts].find(([, shown]) => shown === text)?.[0];
  if (!node) throw new Error(`no diagram shape shows "${text}"`);
  const at = { paragraph: 0, offset };
  t.editor.setSelection({ kind: 'text', nodeId: frame, diagramNode: node, anchor: at, focus: at });
}

/** Run a diagram command at the caret, then type `text` in the new shape, if any. */
export function diagramCommand(
  t: EditTools,
  command:
    | { type: 'addDiagramShape'; position: 'after' | 'before' | 'below' }
    | {
        type: 'deleteDiagramShape';
      },
  text?: string
): void {
  t.run(command);
  if (text) t.type(text);
}

const ids = (xml: string, re: RegExp): string[] => [...xml.matchAll(re)].map((m) => m[1]!);

/**
 * A failure when the data model of `n` names a point that it does not hold, or a drawing shape
 * of `n` names no point of it.
 */
export function consistent(d: SavedDeck, n: number): string | undefined {
  const data = bytesText(d.pkg, `/ppt/diagrams/data${n}.xml`);
  const drawing = bytesText(d.pkg, `/ppt/diagrams/drawing${n}.xml`);
  const points = new Set(ids(data, /<dgm:pt modelId="([^"]+)"/g));
  const named = ids(data, /(?:srcId|destId|parTransId|sibTransId|presAssocID|cxnId)="([^"]+)"/g);
  const missing = named.filter((id) => !points.has(id) && !data.includes(`modelId="${id}"`));
  const shapes = ids(drawing, /<dsp:sp modelId="([^"]+)"/g).filter((id) => !points.has(id));
  return all(
    missing.length ? `the data model names missing points ${missing.join(', ')}` : undefined,
    shapes.length ? `drawing shapes name missing points ${shapes.join(', ')}` : undefined
  );
}

/** The top-level node texts of data part `n`, in `srcOrd` order. */
export function topTexts(d: SavedDeck, n: number): string[] {
  const data = bytesText(d.pkg, `/ppt/diagrams/data${n}.xml`);
  const doc = /<dgm:pt modelId="([^"]+)" type="doc"/.exec(data)?.[1];
  const textOf = (id: string): string => {
    const pt = new RegExp(
      `<dgm:pt modelId="${id.replace(/[{}]/g, '\\$&')}"[^>]*>(.*?)</dgm:pt>`
    ).exec(data)?.[1];
    return [...(pt ?? '').matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((t) => t[1]).join('');
  };
  return [...data.matchAll(/<dgm:cxn ([^>]*)\/>/g)]
    .map((m) => m[1]!)
    .filter((a) => !/ type="/.test(a) && a.includes(`srcId="${doc}"`))
    .map((a) => ({
      ord: Number(/srcOrd="(\d+)"/.exec(a)?.[1]),
      dest: /destId="([^"]+)"/.exec(a)![1]!,
    }))
    .sort((p, q) => p.ord - q.ord)
    .map((c) => textOf(c.dest));
}

const order = (actual: string[], expected: string[]): string | undefined =>
  actual.join('|') === expected.join('|')
    ? undefined
    : `the top-level nodes are ${JSON.stringify(actual)}, not ${JSON.stringify(expected)}`;

const count = (xml: string, re: RegExp): number => (xml.match(re) ?? []).length;

const recipe = (r: Omit<FeatureEdit, 'parts'>): FeatureEdit => ({ parts: DIAGRAM_PARTS, ...r });

/** The recipes, by `objects.diagrams: <edit>`. */
export const DIAGRAM_EDITS: Readonly<Record<string, FeatureEdit>> = {
  'objects.diagrams: edit text': recipe({
    fixture: 'diagram-editing',
    edit: (t) => {
      caretInDiagram(t, 'Build');
      t.type(' & review');
    },
    check: (d) =>
      all(
        has(bytesText(d.pkg, '/ppt/diagrams/data1.xml'), '<a:t>Build &amp; review</a:t>'),
        has(
          bytesText(d.pkg, '/ppt/diagrams/drawing1.xml'),
          'sz="2000" kern="1200"/><a:t>Build &amp; review</a:t>'
        ),
        consistent(d, 1)
      ),
  }),
  'objects.diagrams: add a shape after': recipe({
    fixture: 'diagram-editing',
    edit: (t) => {
      caretInDiagram(t, 'Test');
      diagramCommand(t, { type: 'addDiagramShape', position: 'after' }, 'Review');
    },
    check: (d) => {
      const drawing = bytesText(d.pkg, '/ppt/diagrams/drawing1.xml');
      return all(
        order(topTexts(d, 1), ['Plan', 'Build', 'Test', 'Review', 'Ship']),
        has(drawing, '<a:t>Review</a:t>'),
        // Five boxes and four arrows.
        count(drawing, /prst="roundRect"/g) === 5 ? undefined : 'not five boxes',
        count(drawing, /prst="rightArrow"/g) === 4 ? undefined : 'not four arrows',
        consistent(d, 1)
      );
    },
  }),
  'objects.diagrams: add a shape before': recipe({
    fixture: 'diagram-editing',
    slide: 1,
    edit: (t) => {
      caretInDiagram(t, 'Goals');
      diagramCommand(t, { type: 'addDiagramShape', position: 'before' }, 'Context');
    },
    check: (d) =>
      all(
        order(topTexts(d, 2), ['Context', 'Goals', 'Risks', 'Owners']),
        has(bytesText(d.pkg, '/ppt/diagrams/drawing2.xml'), '<a:t>Context</a:t>'),
        consistent(d, 2)
      ),
  }),
  'objects.diagrams: add a shape below': recipe({
    fixture: 'diagrams',
    edit: (t) => {
      caretInDiagram(t, 'Ship');
      diagramCommand(t, { type: 'addDiagramShape', position: 'below' }, 'Review');
    },
    check: (d) => {
      const data = bytesText(d.pkg, '/ppt/diagrams/data1.xml');
      return all(
        order(topTexts(d, 1), ['Plan', 'Write', 'Ship']),
        has(data, '<a:t>Review</a:t>'),
        has(bytesText(d.pkg, '/ppt/diagrams/drawing1.xml'), '<a:t>Review</a:t>'),
        consistent(d, 1)
      );
    },
  }),
  'objects.diagrams: delete a shape': recipe({
    fixture: 'diagram-editing',
    edit: (t) => {
      caretInDiagram(t, 'Test');
      diagramCommand(t, { type: 'deleteDiagramShape' });
    },
    check: (d) => {
      const data = bytesText(d.pkg, '/ppt/diagrams/data1.xml');
      const drawing = bytesText(d.pkg, '/ppt/diagrams/drawing1.xml');
      return all(
        order(topTexts(d, 1), ['Plan', 'Build', 'Ship']),
        lacks(data, '<a:t>Test</a:t>'),
        lacks(drawing, '<a:t>Test</a:t>'),
        count(drawing, /prst="rightArrow"/g) === 2 ? undefined : 'not two arrows',
        consistent(d, 1)
      );
    },
  }),
};

if (import.meta.main) {
  const { values } = parseArgs({ options: { out: { type: 'string' } } });
  if (values.out) mkdirSync(values.out, { recursive: true });
  let failed = 0;
  for (const [id, r] of Object.entries(DIAGRAM_EDITS)) {
    const result = await runFeatureEdit(id, r);
    if (!result.ok) failed++;
    console.log(`${id}: ${result.ok ? `ok ${result.fingerprint}` : `FAIL ${result.failure}`}`);
    const file = id.replace(/[^a-z0-9.-]+/gi, '-');
    if (values.out && result.bytes) writeFileSync(join(values.out, `${file}.pptx`), result.bytes);
  }
  process.exitCode = failed ? 1 : 0;
}
