// The editable slide surface. The painted slide IS the surface: pointer and keyboard input are
// read here and re-expressed as editor commands; the DOM is never edited directly. The caret and
// selection are engine-painted from layout geometry, like docx-editor's `surface-caret.ts`.

import { paintSlide } from '../output/index.ts';
import { apply, type LayoutItem, type Matrix, type SlideLayout } from '../layout/index.ts';
import type { EditorSelection, ImageCrop, TransformItem } from '../contracts/index.ts';
import { SVG_NS } from '../output/index.ts';
import { ghostBullet } from './surface-ghost.ts';
import { caretAt, linkedAt, linkedAtPoint, pointAt, textFrameMatrix } from './surface-geometry.ts';

export { caretAt, pointAt, textFrameMatrix };
import { selectionRects } from './surface-text-selection.ts';
import { cellsOverlay, TableGestures } from './surface-table.ts';
import { textItemOf } from './table-cells.ts';
import { paragraphUnits, wordSelectionAt } from '../automation/commands/text-units.ts';
import {
  dragSelection,
  PressCounter,
  unitOfPress,
  unitRange,
  type TextDrag,
  type TextDragUnit,
} from './text-drag.ts';
import { presenceElements } from './surface-presence.ts';
import type { PaintedPresence } from './collaboration-presence.ts';
import {
  boxOf,
  cornersOf,
  deltaToParent,
  findItem,
  selectionScope,
  isLineItem,
  lineEnds,
  lineTransform,
  rotationAfterDrag,
  rotationHandle,
  slideCenter,
  toFile,
  toParent,
  unionBounds,
  type Box,
} from '../automation/commands/placement.ts';
import { bindSurfaceInput, type SurfaceInputHost } from './surface-input.ts';
import { selectForContextMenu } from './surface-context-menu.ts';
import { guideElements, snapContext, snapMove, snapResize, SNAP_PX } from './surface-guides.ts';
import type { SnapContext } from './surface-guides.ts';
import type { SnapGuide, SnapResult } from './snapping.ts';
import { cropAt, cropOverlay, onCropImage, startCropDrag } from './surface-crop.ts';
import type { CropDrag, CropHandle } from './surface-crop.ts';
import { withHalos } from './surface-halo.ts';
import { hitItem, inText, onBorder } from './surface-hit.ts';
import { outsideVeil, scrollBoxIntoView } from './surface-view.ts';
import { clampZoom, wheelZoomFactor } from './zoom.ts';
import { resizeOne, scaleMany } from './surface-resize.ts';

const EMU_PER_PX = 9_525;

/** What the surface asks of the editor. Input and clipboard calls are in `SurfaceInputHost`. */
export interface SurfaceHost extends SurfaceInputHost {
  /** Unique to the editor instance: SVG ids of two editors on one page never collide. */
  readonly idPrefix: string;
  readonly layout: () => SlideLayout | undefined;
  /** The layout to paint, when it differs from `layout` (placeholder prompts). */
  readonly paintLayout?: (layout: SlideLayout) => SlideLayout;
  readonly selection: () => EditorSelection;
  readonly readOnly: () => boolean;
  readonly imageUrl: (part: string, rId: string) => string | undefined;
  readonly imageSize?: (part: string, rId: string) => { cx: number; cy: number } | undefined;
  readonly label: (key: string, params?: Record<string, string | number>) => string;
  select(selection: EditorSelection): void;
  /** Collaboration peers on this slide, painted as editing chrome. */
  presence?(): readonly PaintedPresence[];
  /** New boxes in file coordinates; one gesture merges into one undo step. */
  transform(items: readonly TransformItem[], gesture: string): void;
  fitChanged(zoom: number): void;
  zoom(): number;
  /** Set the zoom; `fit` follows the window. */
  setZoom(zoom: number | 'fit'): void;
  slideCount(): number;
  /** Whether drags snap to smart guides (View > Smart guides). */
  smartGuides(): boolean;
  /** True while the format painter is on: the pointer shows the brush, and clicks never drag. */
  painting?(): boolean;
  /** A pointer gesture that started on the slide ended (the format painter paints then). */
  pointerUp?(): void;
  /** The picture in crop mode, or null (`crop-mode.ts`). */
  cropping?(): string | null;
  /** One step of a crop drag; one gesture merges into one undo step. */
  crop?(crop: ImageCrop, gesture: string): void;
  /** Whether an element's aspect ratio is locked: every resize handle keeps the ratio. */
  aspectLocked?(nodeId: string): boolean;
}

type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w' | 'rot' | 'start' | 'end';
type BoxHandle = Exclude<Handle, 'start' | 'end'>;
const HANDLES: readonly BoxHandle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w', 'rot'];
const HANDLE_POS: Record<Exclude<Handle, 'start' | 'end'>, [number, number]> = {
  nw: [0, 0],
  n: [0.5, 0],
  ne: [1, 0],
  e: [1, 0.5],
  se: [1, 1],
  s: [0.5, 1],
  sw: [0, 1],
  w: [0, 0.5],
  rot: [0.5, 0],
};

/** Find a slide item by node id, inside groups too. */
export function itemById(layout: SlideLayout | undefined, id: string): LayoutItem | undefined {
  return layout ? findItem(layout.items, id) : undefined;
}

/** A shape being dragged: its laid-out box in parent space and its slide-space center. */
interface DragItem {
  readonly item: LayoutItem;
  readonly box: Box;
  readonly center: { x: number; y: number };
}

const xy = (p: { x: number; y: number }): [number, number] => [p.x, p.y];

/** Bounds as `x`, `y`, right, bottom, in slide EMU. */
type Bounds = { x: number; y: number; r: number; b: number };

/** A pointer drag of a selection or one of its handles. */
interface Drag {
  readonly handle: Handle | 'move';
  readonly items: readonly DragItem[];
  readonly start: { x: number; y: number };
  readonly bounds: Bounds;
  readonly gesture: string;
  readonly snap: SnapContext;
}

/** The editable slide surface bound to one host element. */
export class SlideSurface {
  readonly root: HTMLDivElement;
  readonly #stage: HTMLDivElement;
  readonly #overlay: SVGSVGElement;
  readonly #input: HTMLTextAreaElement;
  readonly #live: HTMLDivElement;
  readonly #resize: ResizeObserver;
  #drag: Drag | null = null;
  #cropDrag: CropDrag | null = null;
  #marquee: null | { start: { x: number; y: number }; at: { x: number; y: number } } = null;
  #textDrag: TextDrag | null = null;
  /** The last press selected a word or a paragraph, so the browser's `dblclick` keeps it. */
  #unitPress = false;
  readonly #table: TableGestures;
  readonly #unbindInput: () => void;
  #zoomAnchor: { x: number; y: number } | null = null;
  #renderedZoom: number | null = null;
  #paintKey: unknown = null;
  #gestures = 0;
  /** The smart guides of the drag in progress. */
  #guides: readonly SnapGuide[] = [];
  #pressed = false;
  /** The caret the view last followed (`#keepCaretInView`). */
  #caretKey = '';
  /** The presses of a multi-click: their count, time, and place. */
  #presses = new PressCounter();

  constructor(
    readonly host: HTMLElement,
    readonly editor: SurfaceHost
  ) {
    const doc = host.ownerDocument;
    this.root = doc.createElement('div');
    this.root.className = 'pptx-surface';
    this.#stage = doc.createElement('div');
    this.#stage.className = 'pptx-surface__stage';
    this.#overlay = doc.createElementNS(SVG_NS, 'svg') as SVGSVGElement;
    this.#overlay.setAttribute('class', 'pptx-surface__overlay');
    this.#table = new TableGestures(editor);
    this.#input = doc.createElement('textarea');
    this.#input.className = 'pptx-surface__input';
    // The input receives keystrokes and IME composition for the engine; it is focusable and
    // therefore not aria-hidden. Both focusable elements take their names from `en.json`.
    this.#input.setAttribute('aria-label', editor.label('canvas.inputAriaLabel'));
    this.#input.setAttribute('aria-multiline', 'true');
    this.#input.setAttribute('autocomplete', 'off');
    this.#input.setAttribute('autocorrect', 'off');
    this.#input.spellcheck = false;
    this.#input.tabIndex = -1; // Not a Tab stop: the slide is; text selection focuses the input.
    // A screen reader hears which cell the caret is in after Tab or an arrow moves it.
    this.#live = doc.createElement('div');
    this.#live.className = 'pptx-sr-only';
    this.#live.setAttribute('aria-live', 'polite');
    // The input sits in the stage, on the caret, so an IME candidate window opens there.
    this.#stage.append(this.#overlay, this.#input);
    this.root.append(this.#stage, this.#live);
    host.append(this.root);
    this.#stage.addEventListener('pointerdown', this.#onPointerDown);
    this.#stage.addEventListener('dblclick', this.#onDoubleClick);
    this.#stage.addEventListener('contextmenu', (e) =>
      selectForContextMenu(e, editor, this.#hit.bind(this), this.#slidePoint.bind(this), (id, p) =>
        this.#cellCaret(id, p)
      )
    );
    doc.addEventListener('pointermove', this.#onPointerMove);
    doc.addEventListener('pointerup', this.#onPointerUp);
    doc.addEventListener('pointercancel', this.#onPointerCancel);
    this.#unbindInput = bindSurfaceInput(this.#input, this.root, editor);
    this.root.tabIndex = 0;
    this.root.setAttribute('role', 'group');
    // Pinch (ctrl + wheel on trackpads) and Ctrl/Cmd + wheel zoom around the pointer.
    host.addEventListener('wheel', this.#onWheel, { passive: false });
    this.#resize = new ResizeObserver(() => this.#reportFit());
    this.#resize.observe(host);
  }

  destroy(): void {
    const doc = this.host.ownerDocument;
    doc.removeEventListener('pointermove', this.#onPointerMove);
    doc.removeEventListener('pointerup', this.#onPointerUp);
    doc.removeEventListener('pointercancel', this.#onPointerCancel);
    // A drag in progress ends here: its history group closes with it.
    this.#table.up();
    this.#resize.disconnect();
    this.host.removeEventListener('wheel', this.#onWheel);
    this.#unbindInput();
    this.root.remove();
  }

  /** Keyboard focus for typing. */
  focus(): void {
    const sel = this.editor.selection();
    (sel.kind === 'text' ? this.#input : this.root).focus({ preventScroll: true });
  }

  #reportFit(): void {
    const layout = this.editor.layout();
    if (!layout) return;
    const w = layout.slide.width / EMU_PER_PX;
    const h = layout.slide.height / EMU_PER_PX;
    // The host's own padding (an open find bar reserves its top) is not room for the slide.
    const pad = this.host.ownerDocument.defaultView?.getComputedStyle(this.host);
    const padX =
      (parseFloat(pad?.paddingLeft ?? '') || 0) + (parseFloat(pad?.paddingRight ?? '') || 0);
    const padY =
      (parseFloat(pad?.paddingTop ?? '') || 0) + (parseFloat(pad?.paddingBottom ?? '') || 0);
    const availW = Math.max(100, this.host.clientWidth - padX - 48);
    const availH = Math.max(100, this.host.clientHeight - padY - 48);
    this.editor.fitChanged(Math.max(0.1, Math.min(availW / w, availH / h)));
  }

  /** Zoom by a factor, keeping the slide point under (`clientX`, `clientY`) in place. */
  zoomBy(factor: number, clientX?: number, clientY?: number): void {
    const r = this.host.getBoundingClientRect();
    this.#zoomAnchor = { x: clientX ?? r.left + r.width / 2, y: clientY ?? r.top + r.height / 2 };
    this.editor.setZoom(clampZoom(this.editor.zoom() * factor));
  }

  /** Repaint the slide (when the layout changed) and the overlay. */
  render(): void {
    this.root.classList.toggle('pptx-surface--painting', this.editor.painting?.() ?? false);
    // A read-only input still takes focus, keys, and copy, but opens no on-screen keyboard.
    this.#input.readOnly = this.editor.readOnly();
    const layout = this.editor.layout();
    const zoom = this.editor.zoom();
    if (!layout) {
      this.#stage.querySelector('.pptx-surface__slide')?.remove();
      this.#paintKey = null;
      this.#overlay.replaceChildren();
      return;
    }
    // A zoom change keeps the slide point under its anchor (the pointer, or the view center).
    let keep: { client: { x: number; y: number }; fx: number; fy: number } | null = null;
    if (this.#renderedZoom !== null && Math.abs(this.#renderedZoom - zoom) > 1e-6) {
      const host = this.host.getBoundingClientRect();
      const client = this.#zoomAnchor ?? {
        x: host.left + host.width / 2,
        y: host.top + host.height / 2,
      };
      const st = this.#stage.getBoundingClientRect();
      if (st.width > 0 && st.height > 0)
        keep = { client, fx: (client.x - st.left) / st.width, fy: (client.y - st.top) / st.height };
    }
    this.#zoomAnchor = null;
    this.#renderedZoom = zoom;
    const name = this.editor.label('canvas.ariaLabel', {
      index: layout.slide.index + 1,
      total: this.editor.slideCount(),
    });
    if (this.root.getAttribute('aria-label') !== name) this.root.setAttribute('aria-label', name);
    const sel = this.editor.selection();
    const cell = sel.kind === 'text' ? sel.cell : sel.kind === 'cells' ? sel.focus : undefined;
    // A shape that Tab (or a click) selects is named, so a screen reader hears where Tab went.
    const one = sel.kind === 'shapes' && sel.nodeIds.length === 1 ? sel.nodeIds[0] : undefined;
    const shapeName = one ? itemById(layout, one)?.element.name : undefined;
    const said = cell
      ? this.editor.label('canvas.tableCell', { row: cell.row + 1, column: cell.column + 1 })
      : shapeName
        ? this.editor.label('canvas.shapeSelected', { name: shapeName })
        : '';
    if (this.#live.textContent !== said) this.#live.textContent = said;
    const w = (layout.slide.width / EMU_PER_PX) * zoom;
    const h = (layout.slide.height / EMU_PER_PX) * zoom;
    this.#stage.style.width = `${w}px`;
    this.#stage.style.height = `${h}px`;
    if (keep) {
      const st = this.#stage.getBoundingClientRect();
      this.host.scrollLeft += st.left + keep.fx * st.width - keep.client.x;
      this.host.scrollTop += st.top + keep.fy * st.height - keep.client.y;
    }
    const painted = this.editor.paintLayout?.(layout) ?? layout;
    if (this.#paintKey !== painted) {
      this.#paintKey = painted;
      const svg = paintSlide(this.host.ownerDocument, painted, {
        images: this.editor.imageUrl,
        ...(this.editor.imageSize ? { imageSize: this.editor.imageSize } : {}),
        idPrefix: `${this.editor.idPrefix}s${layout.slide.index}`,
      });
      svg.setAttribute('class', 'pptx-surface__slide');
      const old = this.#stage.querySelector('.pptx-surface__slide');
      if (old) old.replaceWith(svg);
      else this.#stage.prepend(svg);
    }
    for (const svg of [this.#stage.querySelector('.pptx-surface__slide'), this.#overlay]) {
      svg?.setAttribute('width', String(w));
      svg?.setAttribute('height', String(h));
    }
    this.#overlay.setAttribute('viewBox', `0 0 ${layout.slide.width} ${layout.slide.height}`);
    this.#renderOverlay(layout, zoom);
    this.#reportFit();
  }

  #renderOverlay(layout: SlideLayout, zoom: number): void {
    const doc = this.host.ownerDocument;
    const ns = SVG_NS;
    const px = EMU_PER_PX / zoom;
    // The area outside the slide, dimmed over the parts of shapes that lie there.
    const parts: Element[] = [outsideVeil(doc, layout.slide.width, layout.slide.height)];
    const sel = this.editor.selection();
    const ids = sel.kind === 'shapes' ? sel.nodeIds : sel.kind === 'none' ? [] : [sel.nodeId];
    const scope = selectionScope(layout.items, sel);
    const scopeItem = scope ? itemById(layout, scope) : undefined;
    if (scopeItem) {
      // The entered group's box, dashed.
      const g = this.#outline(doc, scopeItem, px);
      g.firstElementChild?.setAttribute('class', 'pptx-surface__group-outline');
      g.firstElementChild?.setAttribute('stroke-dasharray', `${px * 2} ${px * 2}`);
      parts.push(g);
    }
    const selected = ids.map((id) => itemById(layout, id)).filter((i): i is LayoutItem => !!i);
    // Crop mode shows the whole image and the crop handles in place of the box handles.
    const crop = this.#cropItem(layout);
    if (crop) {
      const fill = crop.kind === 'picture' ? crop.element.fill : undefined;
      // As paint: the SVG when it loads, else the raster image beside it.
      const url =
        fill?.type === 'picture'
          ? (this.editor.imageUrl(crop.element.partName, fill.svgRId ?? fill.rId) ??
            this.editor.imageUrl(crop.element.partName, fill.rId))
          : undefined;
      parts.push(...cropOverlay(doc, crop, url, px, this.editor.idPrefix));
    }
    // A single line is edited by its two end points: no box, no rotation.
    const line = selected.length === 1 && isLineItem(selected[0]) ? lineEnds(selected[0]) : null;
    for (const item of line ? [] : selected) {
      const g = this.#outline(doc, item, px * (sel.kind === 'text' ? 1 : 1.5));
      if (sel.kind === 'text')
        g.firstElementChild?.setAttribute('stroke-dasharray', `${px * 4} ${px * 3}`);
      parts.push(g);
    }
    // Handles show while text is edited too, so a text box can be resized without leaving it.
    if (line && !this.editor.readOnly()) {
      for (const [handle, p] of [
        ['start', line.start],
        ['end', line.end],
      ] as const) {
        const node = doc.createElementNS(ns, 'circle');
        node.setAttribute('class', 'pptx-surface__handle pptx-surface__handle--point');
        node.setAttribute('data-handle', handle);
        node.setAttribute('cx', String(p.x));
        node.setAttribute('cy', String(p.y));
        node.setAttribute('r', String(px * 5));
        node.setAttribute('stroke-width', String(px));
        parts.push(node);
      }
    } else if (sel.kind !== 'none' && selected.length > 0 && !crop && !this.editor.readOnly()) {
      const single = selected.length === 1 ? selected[0] : undefined;
      const u = unionBounds(selected);
      // One shape: handles on its own (rotated) box. Several: on their common bounds.
      const m: Matrix = single ? single.matrix : [1, 0, 0, 1, u.x, u.y];
      const cx = single ? single.element.xfrm.cx : u.r - u.x;
      const cy = single ? single.element.xfrm.cy : u.b - u.y;
      if (!single) {
        const box = doc.createElementNS(ns, 'rect');
        box.setAttribute('class', 'pptx-surface__group-outline');
        box.setAttribute('x', String(u.x));
        box.setAttribute('y', String(u.y));
        box.setAttribute('width', String(cx));
        box.setAttribute('height', String(cy));
        box.setAttribute('stroke-width', String(px));
        box.setAttribute('stroke-dasharray', `${px * 2} ${px * 2}`);
        parts.push(box);
      }
      for (const handle of HANDLES) {
        if (handle === 'rot' && !single) continue;
        const [hx, hy] = HANDLE_POS[handle];
        const p = apply(m, hx * cx, hy * cy);
        const hp = handle === 'rot' && single ? rotationHandle(single, 22 * px) : p;
        const node = doc.createElementNS(ns, handle === 'rot' ? 'circle' : 'rect');
        node.setAttribute('class', `pptx-surface__handle pptx-surface__handle--${handle}`);
        node.setAttribute('data-handle', handle);
        if (handle === 'rot') {
          node.setAttribute('cx', String(hp.x));
          node.setAttribute('cy', String(hp.y));
          node.setAttribute('r', String(px * 5));
        } else {
          node.setAttribute('x', String(p.x - px * 4));
          node.setAttribute('y', String(p.y - px * 4));
          node.setAttribute('width', String(px * 8));
          node.setAttribute('height', String(px * 8));
        }
        node.setAttribute('stroke-width', String(px));
        parts.push(node);
      }
    }
    parts.push(...presenceElements(doc, layout, this.editor.presence?.() ?? [], px));
    parts.push(...guideElements(doc, this.#guides, px));
    if (this.#marquee) {
      const { start: a, at: b } = this.#marquee;
      const r = doc.createElementNS(ns, 'rect');
      r.setAttribute('class', 'pptx-surface__marquee');
      r.setAttribute('x', String(Math.min(a.x, b.x)));
      r.setAttribute('y', String(Math.min(a.y, b.y)));
      r.setAttribute('width', String(Math.abs(b.x - a.x)));
      r.setAttribute('height', String(Math.abs(b.y - a.y)));
      r.setAttribute('stroke-width', String(px));
      parts.push(r);
    }
    parts.push(...cellsOverlay(doc, layout, sel));
    if (sel.kind === 'text') {
      const item = textItemOf(layout, sel);
      if (item) {
        parts.push(
          ...selectionRects(doc, item, sel.anchor, sel.focus, 'pptx-surface__text-selection')
        );
        const caret = caretAt(item, sel.focus);
        if (caret) {
          const line = doc.createElementNS(ns, 'line');
          // In text, the caret takes the text's color and needs no halo: it shows wherever the
          // text does. In an empty body it keeps the slide caret color and its halo.
          line.setAttribute(
            'class',
            caret.color ? 'pptx-surface__caret pptx-surface__caret--text' : 'pptx-surface__caret'
          );
          if (caret.color) line.style.stroke = caret.color;
          line.setAttribute('x1', String(caret.x1));
          line.setAttribute('y1', String(caret.y1));
          line.setAttribute('x2', String(caret.x2));
          line.setAttribute('y2', String(caret.y2));
          line.setAttribute('stroke-width', String(px * 1.5));
          parts.push(line);
          const ghost =
            item.kind === 'shape' ? ghostBullet(doc, item, sel, textFrameMatrix(item)) : null;
          if (ghost) parts.push(ghost);
          this.#placeInput(caret, zoom / EMU_PER_PX);
          this.#keepCaretInView(sel, caret, zoom / EMU_PER_PX);
        }
      }
    }
    this.#overlay.replaceChildren(...withHalos(parts));
  }

  /**
   * Keep the hidden input on the caret, as tall as the caret line: the stage holds it, so this
   * follows zoom and scroll, and an IME candidate window opens under the caret.
   */
  #placeInput(caret: { x1: number; y1: number; x2: number; y2: number }, scale: number): void {
    const top = Math.min(caret.y1, caret.y2) * scale;
    const height = Math.max(1, Math.abs(caret.y2 - caret.y1) * scale);
    const style = this.#input.style;
    style.left = `${Math.min(caret.x1, caret.x2) * scale}px`;
    style.top = `${top}px`;
    style.height = `${height}px`;
    style.fontSize = style.lineHeight = `${height}px`;
  }

  /**
   * Scroll the caret (the focus end of a range) into view when it moves by a key: typing at a
   * high zoom, or an arrow key. A pointer press does not scroll: the caret is under it.
   */
  #keepCaretInView(
    sel: Extract<EditorSelection, { kind: 'text' }>,
    caret: { x1: number; y1: number; x2: number; y2: number },
    scale: number
  ): void {
    const key = JSON.stringify([sel.nodeId, sel.cell, sel.focus]);
    if (key === this.#caretKey) return;
    this.#caretKey = key;
    if (this.#pressed) return;
    const st = this.#stage.getBoundingClientRect();
    const box = {
      left: st.left + Math.min(caret.x1, caret.x2) * scale,
      right: st.left + Math.max(caret.x1, caret.x2) * scale,
      top: st.top + Math.min(caret.y1, caret.y2) * scale,
      bottom: st.top + Math.max(caret.y1, caret.y2) * scale,
    };
    scrollBoxIntoView(this.host, box, 32);
  }

  /** A caret in the table cell under `p`, for a right-click; none in view mode. */
  #cellCaret(id: string, p: { x: number; y: number }) {
    const layout = this.editor.layout();
    const at = layout && !this.editor.readOnly() ? this.#table.cellItemAt(layout, id, p) : null;
    const point = at ? pointAt(at.item, p.x, p.y) : undefined;
    if (!at || !point) return undefined;
    return { kind: 'text' as const, nodeId: id, cell: at.cell, anchor: point, focus: point };
  }

  /** The picture in crop mode, when it is on this slide. */
  #cropItem(layout: SlideLayout | undefined): LayoutItem | undefined {
    const id = this.editor.readOnly() ? null : this.editor.cropping?.();
    return id ? itemById(layout, id) : undefined;
  }

  /** An item's box outline in slide space. */
  #outline(doc: Document, item: LayoutItem, width: number): SVGGElement {
    const { cx, cy } = item.element.xfrm;
    const g = doc.createElementNS(SVG_NS, 'g') as SVGGElement;
    g.setAttribute('transform', `matrix(${item.matrix.join(' ')})`);
    const rect = doc.createElementNS(SVG_NS, 'rect');
    rect.setAttribute('class', 'pptx-surface__outline');
    rect.setAttribute('width', String(cx));
    rect.setAttribute('height', String(Math.max(cy, 1)));
    rect.setAttribute('stroke-width', String(width));
    g.append(rect);
    return g;
  }

  /** Pointer position in slide EMU. */
  #slidePoint(e: { clientX: number; clientY: number }): { x: number; y: number } {
    const layout = this.editor.layout();
    const rect = this.#stage.getBoundingClientRect();
    const w = layout?.slide.width ?? 1;
    return {
      x: ((e.clientX - rect.left) / rect.width) * w,
      y: ((e.clientY - rect.top) / rect.width) * w,
    };
  }

  /** The editable element under a slide point in a scope (`surface-hit.ts`). */
  #hit(p: { x: number; y: number }, scope?: string): string | undefined {
    return hitItem(this.editor.layout(), p, this.editor.zoom(), scope);
  }

  /** Start dragging the selected items (or a handle of their box). */
  #startDrag(
    handle: Handle | 'move',
    ids: readonly string[],
    start: { x: number; y: number }
  ): void {
    const layout = this.editor.layout();
    const items = ids
      .map((id) => itemById(layout, id))
      .filter((i): i is LayoutItem => !!i)
      .map((item) => ({ item, box: boxOf(item), center: slideCenter(item) }));
    if (items.length === 0) return;
    this.#drag = {
      handle,
      items,
      start,
      bounds: unionBounds(items.map((i) => i.item)),
      gesture: `g${++this.#gestures}`,
      snap: layout ? snapContext(layout, ids) : { targets: [], slide: { width: 0, height: 0 } },
    };
  }

  #onPointerDown = (e: PointerEvent): void => {
    // On a Mac, Ctrl+click is a right-click: the context menu handles it, not a press.
    if (e.button !== 0 || (e.ctrlKey && /Mac|iP(hone|ad)/.test(navigator.platform))) return;
    const layout = this.editor.layout();
    if (!layout) return;
    this.#pressed = true;
    this.#unitPress = false;
    // The format painter only selects: a click never moves or resizes a shape.
    const painting = this.editor.painting?.() ?? false;
    const target = e.target as Element;
    const handle = target.getAttribute?.('data-handle') as Handle | null;
    const p = this.#slidePoint(e);
    const sel = this.editor.selection();
    // In crop mode, a crop handle crops and a press on the image moves the image.
    const crop = painting ? undefined : this.#cropItem(layout);
    const cropHandle = target.getAttribute?.('data-crop-handle') as CropHandle | null;
    if (crop && (cropHandle || onCropImage(crop, p))) {
      this.#cropDrag = startCropDrag(crop, cropHandle ?? 'move', p, `g${++this.#gestures}`);
      this.root.focus({ preventScroll: true });
      e.preventDefault();
      return;
    }
    if (handle && !painting && sel.kind !== 'none') {
      this.#startDrag(handle, sel.kind === 'shapes' ? sel.nodeIds : [sel.nodeId], p);
      e.preventDefault();
      return;
    }
    // The border band of the selected shape, inside or just outside its box, moves it.
    const editing =
      sel.kind === 'text' || sel.kind === 'cells'
        ? sel.nodeId
        : sel.kind === 'shapes' && sel.nodeIds.length === 1
          ? sel.nodeIds[0]
          : undefined;
    const editingItem = editing ? itemById(layout, editing) : undefined;
    if (
      editing &&
      editingItem &&
      !e.shiftKey &&
      !painting &&
      !this.editor.readOnly() &&
      this.#onBorder(editingItem, p)
    ) {
      const group = sel.kind === 'shapes' ? sel.group : undefined;
      this.editor.select({ kind: 'shapes', nodeIds: [editing], ...(group ? { group } : {}) });
      this.root.focus({ preventScroll: true });
      this.#startDrag('move', [editing], p);
      e.preventDefault();
      return;
    }
    // A second press on the text being edited selects its word, a third its paragraph; a drag
    // then extends by that unit (text-drag.ts).
    const clicks = this.#presses.count(e);
    if (clicks >= 2 && !e.shiftKey && this.#selectUnit(layout, sel, p, unitOfPress(clicks))) {
      e.preventDefault();
      return;
    }
    const scope = selectionScope(layout.items, sel);
    // Inside an entered group, a click outside its children leaves the group.
    const inScope = scope ? this.#hit(p, scope) : undefined;
    const id = inScope ?? this.#hit(p);
    const group = inScope ? scope : undefined;
    // A table: a caret in the cell, a cell range, or a grid line drag (surface-table.ts).
    // Shift+click on a table adds it to a shape selection, as Shift+click on any shape does.
    const addShape = e.shiftKey && sel.kind === 'shapes';
    // The format painter paints a cell's word or a cell range, as on a text box.
    const table = id && !addShape ? this.#table.down(layout, id, p, e.shiftKey, painting) : null;
    if (table) {
      e.preventDefault();
      if (table === 'text') this.#input.focus({ preventScroll: true });
      else this.root.focus({ preventScroll: true });
      return;
    }
    if (sel.kind === 'text' && id === sel.nodeId) {
      const item = itemById(layout, id);
      const at = item ? pointAt(item, p.x, p.y) : undefined;
      // A viewer's caret goes only into linked text; another click selects the shape.
      const point = at && (!this.editor.readOnly() || linkedAt(item!, at)) ? at : undefined;
      if (point) {
        const anchor = e.shiftKey ? sel.anchor : point;
        this.#textDrag = { id, unit: 'character', anchor, anchorEnd: anchor };
        this.editor.select({ kind: 'text', nodeId: id, anchor, focus: point });
        e.preventDefault();
        this.#input.focus({ preventScroll: true });
        return;
      }
    }
    // A click on a shape's text places the caret there; its border moves it. A viewer's click
    // places it only in linked text, where the link panel then opens; it edits nothing.
    const textItem = id && !e.shiftKey ? itemById(layout, id) : undefined;
    const at = textItem && this.#inText(textItem, p) ? pointAt(textItem, p.x, p.y) : undefined;
    const caret =
      at && textItem && (!this.editor.readOnly() || linkedAt(textItem, at)) ? at : undefined;
    if (id && caret) {
      this.#textDrag = { id, unit: 'character', anchor: caret, anchorEnd: caret };
      this.editor.select({ kind: 'text', nodeId: id, anchor: caret, focus: caret });
      e.preventDefault();
      this.#input.focus({ preventScroll: true });
      return;
    }
    if (!id) {
      if (!e.shiftKey) this.editor.select({ kind: 'none' });
      this.#marquee = { start: p, at: p };
      this.root.focus({ preventScroll: true });
      e.preventDefault();
      return;
    }
    const sameScope = sel.kind === 'shapes' && sel.group === group;
    const current = sameScope ? sel.nodeIds : [];
    const next = e.shiftKey
      ? current.includes(id)
        ? current.filter((x) => x !== id)
        : [...current, id]
      : current.includes(id)
        ? current
        : [id];
    this.editor.select(
      next.length
        ? { kind: 'shapes', nodeIds: next, ...(group ? { group } : {}) }
        : { kind: 'none' }
    );
    this.root.focus({ preventScroll: true });
    if (!this.editor.readOnly() && !painting && next.includes(id)) this.#startDrag('move', next, p);
    e.preventDefault();
  };

  /**
   * Select the word or paragraph under `p` in the text being edited, and start a drag that
   * extends by it; false when `p` is not on that text or the editor is read-only.
   */
  #selectUnit(
    layout: SlideLayout,
    sel: EditorSelection,
    p: { x: number; y: number },
    unit: TextDragUnit
  ): boolean {
    if (sel.kind !== 'text' || this.editor.readOnly()) return false;
    const item = textItemOf(layout, sel);
    const scope = selectionScope(layout.items, sel);
    const hit = (scope ? this.#hit(p, scope) : undefined) ?? this.#hit(p);
    if (!item || item.kind !== 'shape' || hit !== sel.nodeId) return false;
    if (sel.cell && !this.#table.textPressIn(layout, sel.nodeId, sel.cell, p)) return false;
    const point = pointAt(item, p.x, p.y);
    if (!point) return false;
    const [anchor, anchorEnd] = unitRange(item, point, unit);
    this.#textDrag = { id: sel.nodeId, unit, anchor, anchorEnd };
    this.#unitPress = true;
    this.editor.select({ ...sel, anchor, focus: anchorEnd });
    this.#input.focus({ preventScroll: true });
    return true;
  }

  #inText(item: LayoutItem, p: { x: number; y: number }): boolean {
    return inText(item, p, this.editor.zoom());
  }

  #onBorder(item: LayoutItem, p: { x: number; y: number }): boolean {
    return onBorder(item, p, this.editor.zoom());
  }

  /** The pointer shows what a click does: a text cursor over text, a move cursor on shapes. */
  #updateCursor(p: { x: number; y: number }): void {
    // The format painter's brush comes from the stylesheet (`pptx-surface--painting`).
    if (this.editor.painting?.()) {
      this.#stage.style.cursor = '';
      return;
    }
    const layout = this.editor.layout();
    const sel = this.editor.selection();
    const scope = layout ? selectionScope(layout.items, sel) : undefined;
    const selectedId =
      sel.kind === 'shapes' ? sel.nodeIds[0] : sel.kind === 'none' ? undefined : sel.nodeId;
    const selectedItem = selectedId ? itemById(layout, selectedId) : undefined;
    const crop = this.#cropItem(layout);
    if (
      (crop && onCropImage(crop, p)) ||
      (selectedItem && !this.editor.readOnly() && this.#onBorder(selectedItem, p))
    ) {
      if (this.#stage.style.cursor !== 'move') this.#stage.style.cursor = 'move';
      return;
    }
    const id = (scope ? this.#hit(p, scope) : undefined) ?? this.#hit(p);
    const item = id ? itemById(layout, id) : undefined;
    const table = id && layout ? this.#table.cursor(layout, id, p) : undefined;
    const cursor = table
      ? table
      : !item
        ? ''
        : !this.editor.readOnly() && this.#inText(item, p)
          ? 'text'
          : this.editor.readOnly()
            ? this.#onLink(item, p)
              ? 'pointer'
              : 'default'
            : 'move';
    if (this.#stage.style.cursor !== cursor) this.#stage.style.cursor = cursor;
  }

  /** The point is on linked text of `item`, where a viewer's click opens the link panel. */
  #onLink(item: LayoutItem, p: { x: number; y: number }): boolean {
    return this.#inText(item, p) && linkedAtPoint(item, p.x, p.y);
  }

  #onPointerMove = (e: PointerEvent): void => {
    const layout = this.editor.layout();
    if (!layout) return;
    const p = this.#slidePoint(e);
    if (
      !this.#drag &&
      !this.#textDrag &&
      !this.#table.active &&
      !this.#marquee &&
      e.target instanceof Node &&
      this.#stage.contains(e.target)
    )
      this.#updateCursor(p);
    if (this.#table.move(layout, p)) return;
    if (this.#textDrag) {
      const drag = this.#textDrag;
      const sel = this.editor.selection();
      // A word or paragraph drag in a table cell runs over the cell's own text item.
      const item =
        drag.unit !== 'character' && sel.kind === 'text'
          ? textItemOf(layout, sel)
          : itemById(layout, drag.id);
      const point = item ? pointAt(item, p.x, p.y) : undefined;
      if (item && point)
        this.editor.select(
          drag.unit !== 'character' && sel.kind === 'text'
            ? { ...sel, ...dragSelection(drag, item, point) }
            : { kind: 'text', nodeId: drag.id, ...dragSelection(drag, item, point) }
        );
      return;
    }
    if (this.#marquee) {
      this.#marquee = { ...this.#marquee, at: p };
      this.#renderOverlay(layout, this.editor.zoom());
      return;
    }
    const c = this.#cropDrag;
    if (c) {
      if (Math.abs(p.x - c.start.x) + Math.abs(p.y - c.start.y) >= 2 * EMU_PER_PX)
        this.editor.crop?.(cropAt(c, p), c.gesture);
      return;
    }
    const d = this.#drag;
    if (!d) return;
    let dx = p.x - d.start.x;
    let dy = p.y - d.start.y;
    if (Math.abs(dx) + Math.abs(dy) < 2 * EMU_PER_PX) return;
    if (d.handle === 'move') {
      ({ dx, dy } = this.#guided(snapMove(d.snap, d.bounds, dx, dy, e.shiftKey, this.#snapTo(e))));
      this.editor.transform(
        d.items.map(({ item, box }) => {
          const dp = deltaToParent(item, dx, dy);
          return toFile(item, { ...box, x: box.x + dp.x, y: box.y + dp.y });
        }),
        d.gesture
      );
      return;
    }
    if (d.handle === 'start' || d.handle === 'end') {
      this.editor.transform([this.#moveLineEnd(d, p, e.shiftKey)], d.gesture);
      return;
    }
    const [hx, hy] = HANDLE_POS[d.handle];
    // A resize snaps its dragged edges.
    if (d.handle !== 'rot') {
      const items = d.items.map((i) => i.item);
      const tolerance = this.#snapTo(e);
      const s = snapResize(d.snap, items, d.bounds, [hx, hy], [dx, dy], e.shiftKey, tolerance);
      ({ dx, dy } = this.#guided(s));
    }
    if (d.items.length > 1) {
      const many = scaleMany(d.bounds, d.items, [hx, hy], dx, dy, e.shiftKey);
      this.editor.transform(many, d.gesture);
      return;
    }
    // One shape: work in its parent's space, where its box lives (`surface-resize.ts`).
    const { item, box: x } = d.items[0];
    if (d.handle === 'rot') {
      const rot = rotationAfterDrag(item, x, d.start, p, e.shiftKey);
      this.editor.transform([toFile(item, { ...x, rot })], d.gesture);
      return;
    }
    const s0 = toParent(item, d.start.x, d.start.y);
    const s1 = toParent(item, d.start.x + dx, d.start.y + dy);
    const locked = !!this.editor.aspectLocked?.(item.element.nodeId);
    const keep = { shift: e.shiftKey, locked };
    const one = resizeOne(item, x, [hx, hy], s1.x - s0.x, s1.y - s0.y, keep);
    this.editor.transform([one], d.gesture);
  };

  /** The snap distance in EMU, or null when Alt is down or View > Smart guides is off. */
  #snapTo(e: PointerEvent): number | null {
    if (e.altKey || !this.editor.smartGuides()) return null;
    return (SNAP_PX * EMU_PER_PX) / this.editor.zoom();
  }

  /** Keep a snap's guides for the overlay. */
  #guided(snap: SnapResult): SnapResult {
    this.#guides = snap.guides;
    return snap;
  }

  /** A line with one end at the pointer; Shift snaps its angle to 45° steps. */
  #moveLineEnd(d: Drag, p: { x: number; y: number }, snap: boolean): TransformItem {
    const { item } = d.items[0];
    const ends = lineEnds(item)!;
    const fixed = toParent(item, ...xy(d.handle === 'start' ? ends.end : ends.start));
    let moved = toParent(item, p.x, p.y);
    if (snap) {
      const len = Math.hypot(moved.x - fixed.x, moved.y - fixed.y);
      const angle =
        Math.round(Math.atan2(moved.y - fixed.y, moved.x - fixed.x) / (Math.PI / 4)) *
        (Math.PI / 4);
      moved = { x: fixed.x + Math.cos(angle) * len, y: fixed.y + Math.sin(angle) * len };
    }
    return d.handle === 'start'
      ? lineTransform(item, moved, fixed)
      : lineTransform(item, fixed, moved);
  }

  #onPointerUp = (e: PointerEvent): void => {
    const pressed = this.#pressed;
    this.#pressed = false;
    this.#endGesture(e);
    if (pressed) this.editor.pointerUp?.();
  };

  /**
   * The browser took the pointer (a touch scroll, a system gesture): the drag ends where it is,
   * keeping the steps it applied as one undo step. A marquee selects nothing, and the format
   * painter does not paint.
   */
  #onPointerCancel = (e: PointerEvent): void => {
    this.#pressed = false;
    if (this.#marquee) {
      this.#marquee = null;
      this.render();
    }
    this.#endGesture(e);
  };

  #endGesture(e: PointerEvent): void {
    this.#drag = null;
    this.#cropDrag = null;
    this.#textDrag = null;
    this.#table.up();
    if (this.#guides.length) {
      this.#guides = [];
      this.render();
    }
    const m = this.#marquee;
    if (!m) return;
    this.#marquee = null;
    const layout = this.editor.layout();
    if (!layout) return;
    const p = this.#slidePoint(e);
    const box = {
      x: Math.min(m.start.x, p.x),
      y: Math.min(m.start.y, p.y),
      r: Math.max(m.start.x, p.x),
      b: Math.max(m.start.y, p.y),
    };
    if (box.r - box.x < 2 * EMU_PER_PX && box.b - box.y < 2 * EMU_PER_PX) {
      this.render();
      return;
    }
    // A shape is selected when the marquee encloses it completely.
    const inside = layout.items
      .filter((item) => item.element.layer === 'slide')
      .filter((item) =>
        cornersOf(item).every((c) => c.x >= box.x && c.x <= box.r && c.y >= box.y && c.y <= box.b)
      )
      .map((item) => item.element.nodeId);
    const sel = this.editor.selection();
    const keep = e.shiftKey && sel.kind === 'shapes' && !sel.group ? sel.nodeIds : [];
    const ids = [...new Set([...keep, ...inside])];
    this.editor.select(ids.length ? { kind: 'shapes', nodeIds: ids } : { kind: 'none' });
    this.render();
  }

  #onDoubleClick = (e: MouseEvent): void => {
    const layout = this.editor.layout();
    if (!layout) return;
    const p = this.#slidePoint(e);
    const scope = selectionScope(layout.items, this.editor.selection());
    const id = (scope ? this.#hit(p, scope) : undefined) ?? this.#hit(p);
    let item = id ? itemById(layout, id) : undefined;
    if (!item) return;
    if (item.kind === 'group') {
      // Enter the group, selecting the child under the pointer.
      const inner = this.#hit(p, item.element.nodeId);
      if (inner)
        this.editor.select({ kind: 'shapes', nodeIds: [inner], group: item.element.nodeId });
      return;
    }
    // In a table, the cell under the pointer is the text item; a grid line drags, not selects.
    const cell = id && item.kind === 'frame' ? this.#table.cellItemAt(layout, id, p) : undefined;
    if (cell && id && !this.#table.textPressIn(layout, id, cell.cell, p)) return;
    if (cell) item = cell.item;
    if (this.editor.readOnly() || item.kind !== 'shape' || !item.element.text) return;
    // The second press already selected the word, and a drag may have extended it since: the
    // browser's late `dblclick` must not replace that with the word at the release point.
    if (this.#unitPress) return;
    const point = pointAt(item, p.x, p.y) ?? { paragraph: 0, offset: 0 };
    // A double-click on text selects the word under the pointer (UAX #29), or the run of
    // spaces or punctuation there.
    const units = paragraphUnits(item.element.text.paragraphs[point.paragraph]);
    const [from, to] = wordSelectionAt(units, point.offset);
    this.editor.select({
      kind: 'text',
      nodeId: item.element.nodeId,
      ...(cell ? { cell: cell.cell } : {}),
      anchor: { paragraph: point.paragraph, offset: from },
      focus: { paragraph: point.paragraph, offset: to },
    });
    this.#input.focus({ preventScroll: true });
  };

  #onWheel = (e: WheelEvent): void => {
    if (!e.ctrlKey && !e.metaKey) return;
    e.preventDefault();
    this.zoomBy(wheelZoomFactor(e.deltaY, e.deltaMode), e.clientX, e.clientY);
  };
}
