// The automation host over a live editor, after DOCX Editor's `createBrowserAutomationHost`.
//
// It writes to the editor's own store, so a batch is one undo step in the editor's history and
// the editor repaints and keeps its slide and selection valid, as for any change made on the
// store directly. It follows `load`: a new document is a new session, and the revision stays
// monotonic across sessions so a stale read never matches a new document.
//
// Selection is the reader's: reads map the editor's selection to slide ids, shape ids, and plain
// text offsets, and writes move it after the batch commits.

import type { EditorSelection } from '../contracts/index.ts';
import { cNvPrOf, findShape } from '../automation/targets.ts';
import { textOffsetOf, textPointOf } from '../automation/text-offsets.ts';
import { CommandRefusal } from '../automation/commands/op-batch.ts';
import {
  createAutomationHost,
  type AutomationDocumentPort,
  type AutomationHost,
  type AutomationSelectionPort,
} from '../automation/index.ts';
import {
  attr,
  child,
  nodeById,
  presentationInfo,
  shapeTreeOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../store/index.ts';
import type { DocumentSession } from './document-session.ts';
import type { HistoryMark } from './history-marks.ts';
import { renderImagePng } from './render-image.ts';
import { isCollaborating } from './modules.ts';
import type { PptxEditorInstance } from './pptx-editor-types.ts';

/** The label of the history entry each automation batch commits. */
const LABEL = 'automation';

function elementOf(pkg: PptxPackage, part: string, id: string): Element | undefined {
  const xml = xmlPartOf(pkg, part);
  const node = xml ? nodeById(xml, id) : undefined;
  return node?.type === 'element' ? node : undefined;
}

function selectionPort(editor: PptxEditorInstance): AutomationSelectionPort {
  const slides = (pkg: PptxPackage) => presentationInfo(pkg).slides;
  const goTo = (pkg: PptxPackage, sldId: string): string | undefined => {
    const index = slides(pkg).findIndex((s) => s.sldId === sldId);
    if (index < 0) return undefined;
    if (editor.snapshot().currentSlide !== index) editor.selectSlide(index);
    return slides(pkg)[index]!.partName;
  };
  return {
    read() {
      const pkg = editor.session?.store.package;
      const snap = editor.snapshot();
      if (!pkg) return { slides: [], shapes: null, text: null };
      const all = slides(pkg);
      const current = all[snap.currentSlide];
      const selected = snap.selectedSlides.flatMap((i) => (all[i] ? [all[i].sldId] : []));
      if (!current) return { slides: selected, shapes: null, text: null };
      const shapeIdOf = (nodeId: string): string | undefined => {
        const el = elementOf(pkg, current.partName, nodeId);
        return el ? attr(cNvPrOf(el), 'id') : undefined;
      };
      const sel: EditorSelection = snap.selection;
      if (sel.kind === 'shapes') {
        const shapeIds = sel.nodeIds.flatMap((id) => shapeIdOf(id) ?? []);
        return { slides: selected, shapes: { sldId: current.sldId, shapeIds }, text: null };
      }
      if (sel.kind === 'text' || sel.kind === 'cells') {
        const shapeId = shapeIdOf(sel.nodeId);
        if (!shapeId) return { slides: selected, shapes: null, text: null };
        const shapes = { sldId: current.sldId, shapeIds: [shapeId] };
        const el = elementOf(pkg, current.partName, sel.nodeId);
        const txBody = sel.kind === 'text' && !sel.cell ? child(el, 'p:txBody') : undefined;
        if (!txBody || sel.kind !== 'text') return { slides: selected, shapes, text: null };
        const number = snap.currentSlide + 1;
        const a = textOffsetOf(txBody, sel.anchor, number);
        const f = textOffsetOf(txBody, sel.focus, number);
        const text = {
          sldId: current.sldId,
          shapeId,
          start: Math.min(a, f),
          length: Math.abs(a - f),
        };
        return { slides: selected, shapes, text };
      }
      return { slides: selected, shapes: null, text: null };
    },
    selectSlides(sldIds) {
      const pkg = editor.session?.store.package;
      if (!pkg) return;
      // Each slide once: a repeated id would toggle its slide off again.
      const indices = [
        ...new Set(
          sldIds.map((id) => slides(pkg).findIndex((s) => s.sldId === id)).filter((i) => i >= 0)
        ),
      ];
      indices.forEach((index, i) => editor.selectSlide(index, i === 0 ? 'replace' : 'toggle'));
    },
    selectShapes(sldId, shapeIds) {
      const pkg = editor.session?.store.package;
      const part = pkg ? goTo(pkg, sldId) : undefined;
      if (!pkg || !part) return;
      const tree = shapeTreeOf(pkg, part);
      const hits = shapeIds.flatMap((id) => findShape(tree, id) ?? []);
      const parent = hits[0]?.parent;
      const same = hits.filter((h) => h.parent === parent);
      if (!same.length) return editor.setSelection({ kind: 'none' });
      editor.setSelection({
        kind: 'shapes',
        nodeIds: same.map((h) => h.el.id),
        ...(same[0]!.level > 0 ? { group: parent!.id } : {}),
      });
    },
    selectText(sldId, shapeId, start, length) {
      const pkg = editor.session?.store.package;
      const part = pkg ? goTo(pkg, sldId) : undefined;
      if (!pkg || !part) return;
      const hit = findShape(shapeTreeOf(pkg, part), shapeId);
      const txBody = child(hit?.el, 'p:txBody');
      if (!hit || !txBody) return;
      const number = editor.snapshot().currentSlide + 1;
      const anchor = textPointOf(txBody, start, number);
      const focus = textPointOf(txBody, start + length, number);
      if (!anchor || !focus) return;
      editor.setSelection({ kind: 'text', nodeId: hit.el.id, anchor, focus });
    },
  };
}

/**
 * The slide and selection of the editor now, as the history mark its own edits record: undo
 * goes back to it, as for an edit made in the editor.
 */
function markOf(editor: PptxEditorInstance, pkg: PptxPackage): HistoryMark {
  const snap = editor.snapshot();
  return {
    part: presentationInfo(pkg).slides[snap.currentSlide]?.partName,
    selection: snap.selection,
  };
}

/** A document port over whichever session the editor holds now. */
function editorPort(editor: PptxEditorInstance): AutomationDocumentPort {
  let seen: DocumentSession | null = editor.session;
  let base = 0;
  let generation = 0;
  const follow = (): DocumentSession | null => {
    const now = editor.session;
    if (now !== seen) {
      // A new document: carry the revision on so no request built on the old one matches.
      base += (seen?.store.revision ?? 0) + 1;
      generation += 1;
      seen = now;
    }
    return now;
  };
  // The store token of the request group in progress: consecutive commits of one group are one
  // undo step. Any other commit closes the group in the store.
  let group: { readonly id: string; readonly token: symbol } | null = null;
  const groupToken = (id: string): symbol => {
    if (group?.id !== id) group = { id, token: Symbol('automation-sync') };
    return group.token;
  };
  const listeners = new Set<() => void>();
  let offStore: (() => void) | undefined;
  let bound: DocumentSession | null = null;
  const bind = (): void => {
    const now = follow();
    if (now === bound) return;
    offStore?.();
    bound = now;
    offStore = now?.store.subscribe(() => {
      for (const l of [...listeners]) l();
    });
  };
  const offEditor = editor.subscribe(() => {
    const before = bound;
    bind();
    if (bound !== before) for (const l of [...listeners]) l();
  });
  bind();
  return {
    document: () => (follow(), generation),
    revision: () => base + (follow()?.store.revision ?? 0),
    currentPackage: () => follow()?.store.package ?? null,
    batch: (pkg) => follow()!.batch(pkg),
    apply(ops, historyGroup) {
      const session = follow();
      if (!session) return { ok: false, reason: 'disabled.noEditor' };
      if (editor.snapshot().mode === 'view') return { ok: false, reason: 'disabled.viewing' };
      const before = session.store.revision;
      const mark = markOf(editor, session.store.package);
      const r = session.store.transact(ops, {
        label: LABEL,
        selectionBefore: mark,
        selectionAfter: mark,
        ...(historyGroup === undefined ? {} : { historyGroup: groupToken(historyGroup) }),
      });
      if (!r.ok) return { ok: false, reason: r.rejection.code };
      return { ok: true, changed: session.store.revision !== before };
    },
    committed() {
      // The selection writes of the batch ran after the commit: redo restores where they left it.
      const session = follow();
      if (session) session.store.setSelectionAfter(markOf(editor, session.store.package));
    },
    save: () => editor.save(),
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    measurer: () => follow()?.measurer,
    collaborating: () => isCollaborating(editor),
    selection: selectionPort(editor),
    renderImage(target) {
      const session = follow();
      if (!session) return Promise.reject(new CommandRefusal('errors.exportFailed', 'no editor'));
      return renderImagePng(session, target.slideIndex, {
        ...(target.shapeNodeId === undefined ? {} : { shapeNodeId: target.shapeNodeId }),
        ...(target.width === undefined ? {} : { width: target.width }),
        ...(target.height === undefined ? {} : { height: target.height }),
      });
    },
    dispose() {
      offEditor();
      offStore?.();
      listeners.clear();
    },
  };
}

/**
 * An automation host over a live editor. Each batch writes to the editor's store as one undo
 * step; selection operations read and move the reader's selection. The host does not save: the
 * editor owns its document.
 *
 * @public
 */
export function createBrowserAutomationHost(editor: PptxEditorInstance): AutomationHost {
  return createAutomationHost({
    port: editorPort(editor),
    capabilities: {
      document: true,
      save: false,
      events: true,
      selection: true,
      autofit: true,
      images: typeof document !== 'undefined',
    },
  });
}
