// The answers of `can` for New Slide and Add Section are cached on the deck's structure, not on
// the package (`STRUCTURE_COMMANDS` in `can-cache.ts`), so typing keeps them. These tests
// compare each cached answer with the dry run it stands for (the command compiled on a copy),
// in every state that can change the answer, and after edits that keep the cache.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { textDeck } from '../../../../../scripts/bench/text-deck.ts';
import type { CanResult, EditorCommand } from '../../contracts/index.ts';
import { CommandRefusal } from '../../automation/commands/op-batch.ts';
import {
  parseXmlPart,
  presentationInfo,
  relationshipsOf,
  relsPartName,
  serializeXmlPart,
  xmlPartOf,
} from '../../store/index.ts';
import { compile, type CommandContext } from '../commands.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { renderKey } from '../slide-dependencies.ts';

const win = new GlobalWindow();
afterAll(() => win.happyDOM.close());

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const r = await editor.load(bytes);
  expect(r.ok).toBe(true);
  return editor;
}

/** The dry run `can` stands for: the command compiled as a probe, outside every cache. */
function dryRun(editor: PptxEditorInstance, command: EditorCommand): CanResult {
  const snap = editor.snapshot();
  const ctx: CommandContext = {
    session: editor.session!,
    slide: snap.currentSlide,
    selectedSlides: snap.selectedSlides,
    selection: snap.selection,
    readOnly: snap.mode === 'view',
    ...(snap.presentationView === 'slideSorter' ? { slideSorter: true } : {}),
    probe: true,
    active: {
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      bullets: false,
      numbering: false,
    },
  };
  try {
    compile(ctx, command);
    return { ok: true };
  } catch (e) {
    return { ok: false, reason: e instanceof CommandRefusal ? e.reason : 'errors.unsupported' };
  }
}

/** The structure commands, with a listed layout, an unknown layout, and named sections. */
function commands(editor: PptxEditorInstance): EditorCommand[] {
  const layouts = editor.snapshot().layouts ?? [];
  return [
    { type: 'addSlide' },
    ...layouts.slice(0, 2).map((l) => ({ type: 'addSlide' as const, layoutId: l.id })),
    { type: 'addSlide', layoutId: '/ppt/slideLayouts/none.xml' },
    { type: 'addSection' },
    { type: 'addSection', name: 'Part two' },
    { type: 'addSection', index: 0 },
    { type: 'addSection', index: 99 },
  ];
}

/** Every structure answer equals its dry run. Asked twice: the second answer is cached. */
function expectExact(editor: PptxEditorInstance, label: string): void {
  for (const command of commands(editor))
    for (let pass = 0; pass < 2; pass += 1)
      expect({ label, command, answer: editor.can(command) }).toEqual({
        label,
        command,
        answer: dryRun(editor, command),
      });
}

/** Put the caret at the end of the first paragraph of shape `name` on slide `index`. */
function caretIn(editor: PptxEditorInstance, index: number, name = 'Body'): void {
  editor.goToSlide(index);
  const lay = editor.session!.layout(index)!;
  const item = lay.items.find((i) => i.kind === 'shape' && i.element.name === name)!;
  const end = item.kind === 'shape' ? (item.text?.paragraphLengths[0] ?? 0) : 0;
  const caret = { paragraph: 0, offset: end };
  editor.setSelection({ kind: 'text', nodeId: item.element.nodeId, anchor: caret, focus: caret });
}

/** Set an attribute on the root of a part, on the store directly. */
function setRootAttr(editor: PptxEditorInstance, part: string, local: string, value: string) {
  const store = editor.session!.store;
  const root = xmlPartOf(store.package, part)!.root;
  const r = store.transact([
    { op: 'setAttribute', part, node: root.id, name: { ns: '', local }, value },
  ]);
  expect(r.ok).toBe(true);
}

/** Count the probes `can` compiles, by the probe batches it opens. */
function countProbes(editor: PptxEditorInstance): { readonly count: () => number } {
  const session = editor.session!;
  let n = 0;
  const probeBatch = session.probeBatch.bind(session);
  session.probeBatch = (base) => ((n += 1), probeBatch(base));
  return { count: () => n };
}

describe('structure answers of can', () => {
  test('equal the dry run across slides, modes, views, selections, and history', async () => {
    const editor = await open(textDeck(4));
    expectExact(editor, 'opened');
    for (const i of [0, 2, 3]) {
      editor.goToSlide(i);
      expectExact(editor, `slide ${i}`);
    }
    caretIn(editor, 1);
    expectExact(editor, 'text selection');
    expect(editor.input('text', 'abc').ok).toBe(true);
    expectExact(editor, 'after typing');
    const shape = editor.session!.layout(1)!.items[0]!;
    editor.setSelection({ kind: 'shapes', nodeIds: [shape.element.nodeId] });
    expectExact(editor, 'shape selection');
    editor.selectSlide(3, 'range');
    expectExact(editor, 'slide range');
    editor.setMode('view');
    expectExact(editor, 'view mode');
    editor.setMode('edit');
    expect(editor.exec({ type: 'setPresentationView', view: 'slideSorter' }).ok).toBe(true);
    expectExact(editor, 'slide sorter');
    expect(editor.exec({ type: 'setPresentationView', view: 'normal' }).ok).toBe(true);
    expect(editor.exec({ type: 'addSection', name: 'One' }).ok).toBe(true);
    expectExact(editor, 'one section');
    editor.exec({ type: 'addSection', name: 'One again' });
    expectExact(editor, 'a second section at the slide');
    caretIn(editor, 2);
    editor.input('text', 'x');
    expectExact(editor, 'typing in a section');
    expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
    expectExact(editor, 'after New Slide');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expectExact(editor, 'after undo');
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expectExact(editor, 'after redo');
    while (editor.snapshot().canUndo) editor.exec({ type: 'undo' });
    expectExact(editor, 'empty history');
    editor.destroy();
  });

  test('equal the dry run in a deck with no slides', async () => {
    const editor = await open(textDeck(2));
    // Empty the slide list on the store: the editor keeps one slide at least.
    const store = editor.session!.store;
    const info = presentationInfo(store.package);
    const r = store.transact(
      info.slides.map((sl) => ({
        op: 'removeNode' as const,
        part: info.presentationPart,
        node: sl.nodeId,
      }))
    );
    expect(r.ok).toBe(true);
    expect(editor.snapshot().slides).toHaveLength(0);
    expectExact(editor, 'no slides');
    expect(editor.can({ type: 'addSection' }).ok).toBe(false);
    expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
    expectExact(editor, 'one new slide');
    editor.destroy();
  });

  test('change when a slide loses its layout, and hold after slide content edits', async () => {
    const editor = await open(textDeck(3));
    const pkg = editor.session!.store.package;
    const slide = presentationInfo(pkg).slides[1]!.partName;
    editor.goToSlide(1);
    expect(editor.can({ type: 'addSlide' })).toEqual({ ok: true });
    // Hidden or not, a slide gives New Slide its layout: the answer holds.
    setRootAttr(editor, slide, 'show', '0');
    expect(editor.snapshot().slides[1]!.hidden).toBe(true);
    expectExact(editor, 'hidden slide');
    // Without its layout relationship, New Slide after it is refused.
    const rels = relsPartName(slide);
    const layoutRel = xmlPartOf(pkg, rels)!.root.children.find(
      (c) => c.type === 'element' && c.attrs.some((a) => a.value.endsWith('/slideLayout'))
    )!;
    expect(relationshipsOf(pkg, slide).length).toBeGreaterThan(0);
    const r = editor.session!.store.transact([
      { op: 'removeNode', part: rels, node: layoutRel.id },
    ]);
    expect(r.ok).toBe(true);
    expect(editor.can({ type: 'addSlide' })).toEqual({ ok: false, reason: 'errors.malformed' });
    expectExact(editor, 'slide without a layout');
    editor.goToSlide(0);
    expectExact(editor, 'the slide before it');
    editor.destroy();
  });

  test('typing compiles no structure probe again', async () => {
    const editor = await open(textDeck(6));
    caretIn(editor, 2);
    for (const c of commands(editor)) editor.can(c);
    const probes = countProbes(editor);
    for (let i = 0; i < 3; i += 1) {
      expect(editor.input('text', 'x').ok).toBe(true);
      for (const c of commands(editor)) editor.can(c);
    }
    expect(probes.count()).toBe(0);
    // A structure edit compiles them again.
    expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
    for (const c of commands(editor)) editor.can(c);
    expect(probes.count()).toBe(commands(editor).length);
    editor.destroy();
  });

  test('moving between slides on one layout compiles no New Slide probe again', async () => {
    const editor = await open(textDeck(6));
    const addSlide: EditorCommand = { type: 'addSlide' };
    editor.goToSlide(1);
    editor.can(addSlide);
    const probes = countProbes(editor);
    for (const i of [2, 4, 5, 0, 3]) {
      editor.goToSlide(i);
      const before = probes.count();
      const answer = editor.can(addSlide);
      expect(probes.count()).toBe(before);
      expect(answer).toEqual(dryRun(editor, addSlide));
    }
    editor.destroy();
  });
});

/** A small seeded generator (mulberry32), so a failing seed replays. */
function random(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4_294_967_296;
  };
}

/** Replace a part with a parsed copy of itself, as a collaboration replica publishes it. */
function replaceAsRemote(editor: PptxEditorInstance, part: string): void {
  const store = editor.session!.store;
  const xml = xmlPartOf(store.package, part);
  if (!xml) return;
  const copy = parseXmlPart(part, serializeXmlPart(xml));
  if (!copy.ok) throw new Error(`reparse ${part}`);
  expect(store.replaceParts([{ kind: 'xml', partName: part, xml: copy.part }]).ok).toBe(true);
}

describe('structure answers under random edits', () => {
  const SEEDS = [1, 2, 3, 4, 5, 6, 7, 8];
  for (const seed of SEEDS)
    test(`seed ${seed}: every answer equals the dry run after each edit`, async () => {
      const rnd = random(seed);
      const pick = <T>(xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)]!;
      const editor = await open(textDeck(5));
      const info = () => presentationInfo(editor.session!.store.package);
      const anySlide = () => Math.floor(rnd() * info().slides.length);
      const edits: Record<string, () => void> = {
        type: () => {
          const i = anySlide();
          if (!editor.session!.layout(i)?.items.some((it) => it.element.name === 'Body')) return;
          caretIn(editor, i);
          editor.input('text', pick(['x', 'yz', ' ']));
        },
        goTo: () => editor.goToSlide(anySlide()),
        addSlide: () => void editor.exec({ type: 'addSlide' }),
        deleteSlide: () => {
          if (info().slides.length > 1)
            editor.exec({ type: 'deleteSlides', indices: [anySlide()] });
        },
        moveSlide: () =>
          void editor.exec({ type: 'moveSlides', indices: [anySlide()], to: anySlide() }),
        addSection: () => void editor.exec({ type: 'addSection', name: `S${seed}` }),
        undo: () => void editor.exec({ type: 'undo' }),
        redo: () => void editor.exec({ type: 'redo' }),
        showFlip: () => {
          const slide = info().slides[anySlide()];
          if (slide) setRootAttr(editor, slide.partName, 'show', pick(['0', '1']));
        },
        layoutEdit: () => {
          const layout = info().slides[anySlide()]?.layoutPart;
          if (layout) setRootAttr(editor, layout, 'preserve', pick(['0', '1']));
        },
        masterEdit: () => {
          const master = info().slides[0]?.masterPart;
          if (master) setRootAttr(editor, master, 'preserve', pick(['0', '1']));
        },
        themeEdit: () => {
          const theme = info().slides[0]?.themePart;
          if (theme) setRootAttr(editor, theme, 'name', pick(['A', 'B']));
        },
        layoutRelDrop: () => {
          // A slide loses its layout: New Slide after it is refused.
          const slide = info().slides[anySlide()];
          if (!slide) return;
          const rels = relsPartName(slide.partName);
          const rel = xmlPartOf(editor.session!.store.package, rels)?.root.children.find(
            (c) => c.type === 'element' && c.attrs.some((a) => a.value.endsWith('/slideLayout'))
          );
          if (rel) editor.session!.store.transact([{ op: 'removeNode', part: rels, node: rel.id }]);
        },
        remoteSlide: () => {
          const slide = info().slides[anySlide()];
          if (slide) replaceAsRemote(editor, slide.partName);
        },
        remotePresentation: () => replaceAsRemote(editor, info().presentationPart),
        remoteRels: () => replaceAsRemote(editor, relsPartName(info().presentationPart)),
        mode: () => editor.setMode(pick(['edit', 'view'] as const)),
        view: () =>
          void editor.exec({
            type: 'setPresentationView',
            view: pick(['normal', 'slideSorter'] as const),
          }),
      };
      const names = Object.keys(edits);
      const done: string[] = [];
      for (let step = 0; step < 40; step += 1) {
        const name = pick(names);
        done.push(name);
        edits[name]!();
        expectExact(editor, `seed ${seed}, steps ${done.join(' ')}`);
      }
      editor.destroy();
    });
});

describe('slide summaries after an edit', () => {
  test('have the keys a full pass gives', async () => {
    const editor = await open(textDeck(5));
    const fullKeys = () => {
      const session = editor.session!;
      const pkg = session.store.package;
      return presentationInfo(pkg).slides.map(
        (_, i) => renderKey(pkg, i, session.paintGeneration) ?? ''
      );
    };
    editor.snapshot();
    caretIn(editor, 3);
    editor.input('text', 'x');
    expect(editor.snapshot().slides.map((s) => s.key)).toEqual(fullKeys());
    // A master edit changes every key.
    const pkg = editor.session!.store.package;
    const master = presentationInfo(pkg).slides[0]!.masterPart!;
    setRootAttr(editor, master, 'preserve', '1');
    const keys = editor.snapshot().slides.map((s) => s.key);
    expect(keys).toEqual(fullKeys());
    // Hiding a slide changes its summary.
    setRootAttr(editor, presentationInfo(pkg).slides[4]!.partName, 'show', '0');
    expect(editor.snapshot().slides[4]!.hidden).toBe(true);
    expect(editor.snapshot().slides.map((s) => s.key)).toEqual(fullKeys());
    expect(editor.exec({ type: 'moveSlides', indices: [4], to: 0 }).ok).toBe(true);
    expect(editor.snapshot().slides.map((s) => s.key)).toEqual(fullKeys());
    editor.destroy();
  });

  test('keep the summaries of slides the edit did not touch', async () => {
    const editor = await open(
      buildPptx({
        slides: [0, 1, 2].map((i) => ({
          shapes: shape({
            id: 2,
            name: 'Body',
            x: 0,
            y: 0,
            cx: 914_400,
            cy: 914_400,
            text: para(run(`S${i}`)),
          }),
        })),
      })
    );
    const before = editor.snapshot().slides;
    caretIn(editor, 0);
    editor.input('text', 'x');
    const after = editor.snapshot().slides;
    expect(after[0]).not.toBe(before[0]!);
    expect(after[1]).toBe(before[1]!);
    expect(after[2]).toBe(before[2]!);
    editor.destroy();
  });
});
