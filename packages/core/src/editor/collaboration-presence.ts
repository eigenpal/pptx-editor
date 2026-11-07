// Presence between the editor and an attached collaboration session, after DOCX Editor's
// surface wiring (`paginated-surface.ts`): the editor publishes its selection on every change,
// follows its own caret when a peer's edit moves it, returns to the selection of an undone
// change, and hands the surface the peers' selections to paint.
//
// The editor addresses text as (shape, paragraph index, offset in the paragraph); a session
// addresses it as (text node, offset in that node), because only a node has a shared key in
// the room. This module converts between the two.

import type { EditorSelection, TableCellAddress, TextPoint } from '../contracts/index.ts';
import {
  ancestorsOf,
  children,
  nodeById,
  xmlPartOf,
  type Element,
  type PptxPackage,
  type XmlPart,
} from '../store/index.ts';
import type {
  CollaborationLocalSelection,
  EditorCollaborationSession,
  EditorModuleRegistry,
} from './modules.ts';
import { cellTxBody } from './table-cells.ts';
import { tblOf } from '../automation/commands/table-model.ts';
import { paragraphsOf, slotsOf, type Slot } from '../automation/commands/text-model.ts';

type Point = { readonly node: string; readonly offset: number };
type Selected = Exclude<EditorSelection, { kind: 'none' }>;

/** A peer's selection in editor coordinates, for the surface to paint. */
export interface PaintedPresence {
  readonly actorId: string;
  readonly name: string;
  /** A checked `#rgb`/`#rrggbb` color, or undefined for the theme accent. */
  readonly color?: string;
  readonly selection: Selected;
}

const COLOR = /^#(?:[0-9a-f]{3}|[0-9a-f]{6})$/i;

/** The participant color if it is a plain hex color; anything else falls back to CSS. */
export function safeParticipantColor(value: string | undefined): string | undefined {
  return value !== undefined && COLOR.test(value) ? value : undefined;
}

/** The text node inside a run or field slot (`a:r/a:t/#text`), if it has one. */
function textNodeOf(slot: Slot): { id: string; length: number } | undefined {
  for (const c of slot.el.children) {
    if (c.type !== 'element' || c.local !== 't') continue;
    const text = c.children.find((n) => n.type === 'text');
    if (text?.type === 'text') return { id: text.id, length: text.value.length };
  }
  return undefined;
}

function editorToNode(paragraph: Element, offset: number): Point {
  const slots = slotsOf(paragraph).filter((s) => s.kind !== 'break');
  const inside =
    slots.find((s) => s.start <= offset && offset < s.end) ??
    [...slots].reverse().find((s) => s.end === offset);
  const node = inside ? textNodeOf(inside) : undefined;
  if (!inside || !node) return { node: paragraph.id, offset: 0 };
  return { node: node.id, offset: Math.min(offset - inside.start, node.length) };
}

/** The local editor selection as a session selection, or null when there is nothing to send. */
export function toCollaborationSelection(
  pkg: PptxPackage,
  part: string,
  selection: EditorSelection
): CollaborationLocalSelection | null {
  if (selection.kind === 'none') return null;
  if (selection.kind === 'shapes')
    return selection.nodeIds.length > 0
      ? { kind: 'shapes', part, shapes: selection.nodeIds }
      : null;
  // A cell range shows as its table selected.
  if (selection.kind === 'cells') return { kind: 'shapes', part, shapes: [selection.nodeId] };
  // Text in a diagram shape is in the diagram's own parts: peers see the diagram selected.
  if (selection.diagramNode !== undefined)
    return { kind: 'shapes', part, shapes: [selection.nodeId] };
  const xml = xmlPartOf(pkg, part);
  const shape = xml ? nodeById(xml, selection.nodeId) : undefined;
  if (shape?.type !== 'element') return null;
  const body = selection.cell
    ? cellTxBody(shape, selection.cell)
    : shape.children.find((c) => c.type === 'element' && c.kind === 'txBody');
  if (body?.type !== 'element') return null;
  const paragraphs = paragraphsOf(body);
  const point = (p: TextPoint): Point | null => {
    const paragraph = paragraphs[p.paragraph];
    return paragraph ? editorToNode(paragraph, p.offset) : null;
  };
  const anchor = point(selection.anchor);
  const head = point(selection.focus);
  return anchor && head ? { kind: 'text', part, anchor, head } : null;
}

function nodeToEditor(
  xml: XmlPart,
  p: Point
): { shape: string; cell?: TableCellAddress; point: TextPoint } | null {
  const node = nodeById(xml, p.node);
  if (!node) return null;
  const chain = ancestorsOf(xml, p.node);
  const self = node.type === 'element' ? [node] : [];
  const up = [...chain, ...self].reverse();
  const paragraph = up.find((e) => e.kind === 'paragraph');
  const body = up.find((e) => e.kind === 'txBody');
  const shape = up.find((e) => e.kind === 'sp' || e.kind === 'graphicFrame');
  if (!paragraph || !body || !shape) return null;
  // Text in a table cell: the cell's row and column in the frame's table.
  const tc = up.find((e) => e.local === 'tc');
  const tr = up.find((e) => e.local === 'tr');
  const rows = children(tblOf(shape), 'a:tr');
  const cell =
    shape.kind === 'graphicFrame' && tc && tr
      ? { row: rows.indexOf(tr), column: children(tr, 'a:tc').indexOf(tc) }
      : undefined;
  if (shape.kind === 'graphicFrame' && (!cell || cell.row < 0 || cell.column < 0)) return null;
  // A cell with two text bodies draws them as one (`cellTxBody`): the caret's paragraph counts
  // in the joined body, as `toCollaborationSelection` counts it.
  const drawn = cell ? cellTxBody(shape, cell) : body;
  const index = drawn ? paragraphsOf(drawn).indexOf(paragraph) : -1;
  if (index < 0) return null;
  let offset = 0;
  if (node.type === 'text') {
    const slotEl = up.find((e) => e.kind === 'run' || e.kind === 'field');
    const slot = slotsOf(paragraph).find((s) => s.el === slotEl);
    if (slot) offset = slot.start + Math.min(Math.max(0, p.offset), slot.end - slot.start);
  }
  return { shape: shape.id, ...(cell ? { cell } : {}), point: { paragraph: index, offset } };
}

/** A session selection in editor coordinates, or null when it does not resolve here. */
export function fromCollaborationSelection(
  pkg: PptxPackage,
  selection: CollaborationLocalSelection
): Selected | null {
  const xml = xmlPartOf(pkg, selection.part);
  if (!xml) return null;
  if (selection.kind === 'shapes') {
    const ids = selection.shapes.filter((id) => nodeById(xml, id)?.type === 'element');
    return ids.length > 0 ? { kind: 'shapes', nodeIds: ids } : null;
  }
  const anchor = nodeToEditor(xml, selection.anchor);
  const head = nodeToEditor(xml, selection.head);
  if (!anchor || !head || anchor.shape !== head.shape) return null;
  if (JSON.stringify(anchor.cell) !== JSON.stringify(head.cell)) return null;
  return {
    kind: 'text',
    nodeId: anchor.shape,
    ...(anchor.cell ? { cell: anchor.cell } : {}),
    anchor: anchor.point,
    focus: head.point,
  };
}

/** The editor's side of presence: one per editor instance. */
export class PresenceLink {
  readonly #registry: EditorModuleRegistry;
  #published = '';
  #painted: {
    source: unknown;
    pkg: unknown;
    part: string;
    out: readonly PaintedPresence[];
  } | null = null;

  constructor(registry: EditorModuleRegistry) {
    this.#registry = registry;
  }

  #session(): EditorCollaborationSession | undefined {
    const session = this.#registry.collaboration?.session;
    return session?.attached ? session : undefined;
  }

  /** Forget what was published: the next `publish` sends again (after an attach). */
  reset(): void {
    this.#published = '';
  }

  /** Publish the editor selection when it changed. */
  publish(
    pkg: PptxPackage | undefined,
    part: string | undefined,
    selection: EditorSelection
  ): void {
    const session = this.#session();
    if (!session?.setLocalSelection) return;
    const next = pkg && part ? toCollaborationSelection(pkg, part, selection) : null;
    const key = JSON.stringify(next);
    if (key === this.#published) return;
    this.#published = key;
    session.setLocalSelection(next);
  }

  /** Peers' selections on `part`, in editor coordinates. Same array until something changes. */
  painted(pkg: PptxPackage | undefined, part: string | undefined): readonly PaintedPresence[] {
    const remote = this.#session()?.remoteSelections?.() ?? [];
    if (!pkg || !part || remote.length === 0) return [];
    const c = this.#painted;
    if (c && c.source === remote && c.pkg === pkg && c.part === part) return c.out;
    const out: PaintedPresence[] = [];
    for (const peer of remote) {
      if (peer.selection.part !== part) continue;
      const selection = fromCollaborationSelection(pkg, peer.selection);
      if (!selection) continue;
      const color = safeParticipantColor(peer.color);
      out.push({ actorId: peer.actorId, name: peer.name, ...(color ? { color } : {}), selection });
    }
    this.#painted = { source: remote, pkg, part, out };
    return out;
  }

  /** Listen to peers' selection changes. */
  subscribe(listener: () => void): () => void {
    const session = this.#registry.collaboration?.session;
    return session?.subscribeRemoteSelections?.(() => listener()) ?? (() => {});
  }

  /**
   * The local text selection after a remote change, moved with its characters. The editor
   * publishes on every change, so the session's `from` is the selection the editor shows.
   */
  followRemote(
    pkg: PptxPackage,
    part: string | undefined,
    selection: EditorSelection
  ): EditorSelection {
    if (selection.kind !== 'text') return selection;
    const move = this.#session()?.remoteSelectionMove?.();
    if (!move || move.to.part !== part) return selection;
    const next = fromCollaborationSelection(pkg, move.to);
    if (!next) return selection;
    this.#published = JSON.stringify(move.to);
    return next;
  }

  /**
   * Where a session undo or redo goes, as a store undo does: the selection published when the
   * change was made, on its slide. Without one, the first slide the step changed, unless the
   * current slide changed too.
   */
  historyTarget(
    before: PptxPackage,
    pkg: PptxPackage,
    slides: readonly string[],
    current: number
  ): { slide: number; selection: Selected | null } {
    const back = this.#session()?.historySelection?.();
    const selection = back ? fromCollaborationSelection(pkg, back) : null;
    const at = back && selection ? slides.indexOf(back.part) : -1;
    if (at >= 0) return { slide: at, selection };
    const changed = (part: string | undefined) =>
      !!part && before.parts.get(part) !== pkg.parts.get(part);
    if (changed(slides[current])) return { slide: current, selection: null };
    const first = slides.findIndex(changed);
    return { slide: first >= 0 ? first : current, selection: null };
  }
}
