// The editable slide surface translates input into editor commands and paints layout geometry.

import { SlidePainter, type AltTextPort } from '../output/index.ts';
import type { LayoutItem, SlideLayout } from '../layout/index.ts';
import type { EditorSelection, ImageCrop, TransformItem } from '../contracts/index.ts';
import { SVG_NS } from '../output/index.ts';
import { ghostBullet } from './surface-ghost.ts';
import { caretAt, linkedAt, linkedAtPoint, pointAt, textFrameMatrix } from './surface-geometry.ts';

export { caretAt, pointAt, textFrameMatrix };
import { selectionRects } from './surface-text-selection.ts';
import { cellsOverlay, TableGestures } from './surface-table.ts';
import { TableChrome } from './surface-table-chrome.ts';
import { PlaceholderChrome } from './surface-placeholder-chrome.ts';
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
  cornersOf,
  findItem,
  selectionScope,
  isLineItem,
  lineEnds,
} from '../automation/commands/placement.ts';
import { bindSurfaceInput, type SurfaceInputHost } from './surface-input.ts';
import { selectForContextMenu } from './surface-context-menu.ts';
import { guideElements, SNAP_PX } from './surface-guides.ts';
import { cropAt, cropOverlay, onCropImage, startCropDrag } from './surface-crop.ts';
import type { CropDrag, CropHandle } from './surface-crop.ts';
import { withHalos } from './surface-halo.ts';
import { hitItem, inText, onBorder } from './surface-hit.ts';
import { fitZoom, outsideVeil, scrollBoxIntoView } from './surface-view.ts';
import { clampZoom, wheelZoomFactor } from './zoom.ts';
import { selectionAnnouncement } from './surface-announce.ts';
import { ShapeDrag, type Handle } from './surface-drag.ts';
import { boxOutline, handleElements, memberOutline } from './surface-handles.ts';
import {
  adjustAt,
  adjustDragLive,
  adjustHandleElements,
  startAdjustDrag,
  type AdjustDrag,
} from './surface-adjust.ts';
import type { Relayout } from './drag-preview.ts';
import * as paint from './surface-paint.ts';
import { CursorLayer, NextFrame, PAINT_BUDGET, setAttr, setStyle } from './surface-paint.ts';

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
  /** Changes when the slide paints differently with no edit (`DocumentSession.paintKey`). */
  readonly paintKey?: () => string | undefined;
  readonly imageSize?: (part: string, rId: string) => { cx: number; cy: number } | undefined;
  readonly label: (key: string, params?: Record<string, string | number>) => string;
  /** The alt text of an element: pictures are named by it, and the live region says it. */
  readonly altText?: AltTextPort;
  select(selection: EditorSelection): void;
  /** Collaboration peers on this slide, painted as editing chrome. */
  presence?(): readonly PaintedPresence[];
  /** The boxes a shape drag writes on release, in file coordinates: one undo step. */
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
  /** One step of an adjust handle drag (`setAdjustments`); one gesture is one undo step. */
  adjust?(values: Readonly<Record<string, number>>, gesture: string): void;
  /** Whether an element's aspect ratio is locked: every resize handle keeps the ratio. */
  aspectLocked?(nodeId: string): boolean;
  /** Lay out a dragged item at a new box, for the drag preview (drag-preview.ts). */
  relayout?: Relayout;
  /** Ask the user for a picture file and insert it (an empty placeholder's picture button). */
  pickImage?(): void;
  /** The placeholder type a new table or picture can fill on this slide (`emptyPlaceholderType`). */
  fillableType?(nodeId: string): string | undefined;
}

/** Find a slide item by node id, inside groups too. */
export function itemById(layout: SlideLayout | undefined, id: string): LayoutItem | undefined {
  return layout ? findItem(layout.items, id) : undefined;
}

/** The editable slide surface bound to one host element. */
export class SlideSurface {
  readonly root: HTMLDivElement;
  readonly #stage: HTMLDivElement;
  readonly #overlay: SVGSVGElement;
  /** The pointer's cursor, on a layer over the slide (surface-paint.ts). */
  readonly #cursor: CursorLayer;
  /** The frame that paints the items a paint budget left out (`PAINT_BUDGET`). */
  readonly #paintRest: NextFrame;
  /** The view for the window, and the repaint when the window grows (surface-paint.ts). */
  readonly #reach: paint.WindowReach;
  readonly #input: HTMLTextAreaElement;
  readonly #live: HTMLDivElement;
  readonly #resize: ResizeObserver;
  /** A shape drag: it previews while the pointer moves and writes on release (surface-drag.ts). */
  readonly #drag: ShapeDrag;
  #cropDrag: CropDrag | null = null;
  #adjustDrag: AdjustDrag | null = null;
  #marquee: null | { start: { x: number; y: number }; at: { x: number; y: number } } = null;
  #textDrag: TextDrag | null = null;
  /** The last press selected a word or a paragraph, so the browser's `dblclick` keeps it. */
  #unitPress = false;
  readonly #table: TableGestures;
  readonly #tableChrome: TableChrome;
  readonly #placeholderChrome: PlaceholderChrome;
  readonly #unbindInput: () => void;
  #zoomAnchor: { x: number; y: number } | null = null;
  #renderedZoom: number | null = null;
  #paintKey: unknown = null;
  /** Paints the slide, reusing the nodes of unchanged items (`SlidePainter`). */
  #painter: SlidePainter | null = null;
  /**
   * The slide size the fit zoom was last computed for. The fit reads computed styles and the
   * host's size, which make the browser lay out the whole document. A render that keeps the
   * slide size keeps the fit, so a keystroke reads no style. The ResizeObserver refits when the
   * host's size changes, and when the surface's content box changes: a new zoom, or a new
   * surface padding (`--doc-slide-padding`, a style sheet that loads late). At the fit zoom the
   * surface fills the host, so a new padding changes its content box at once; zoomed in, the
   * fit follows at the next zoom change.
   */
  #fitSlide = '';
  #gestures = 0;
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
    this.#overlay.setAttribute('aria-hidden', 'true'); // Selection chrome: the live region says it.
    this.#table = new TableGestures(editor);
    this.#tableChrome = new TableChrome(editor);
    this.#placeholderChrome = new PlaceholderChrome(editor);
    this.#input = doc.createElement('textarea');
    this.#input.className = 'pptx-surface__input';
    // The input receives keystrokes and IME composition for the engine; it is focusable and
    // therefore not aria-hidden. Both focusable elements take their names from `en.json`.
    this.#input.setAttribute('aria-label', editor.label('canvas.inputAriaLabel'));
    this.#input.setAttribute('aria-multiline', 'true');
    this.#input.setAttribute('autocomplete', 'off');
    this.#input.setAttribute('autocorrect', 'off');
    Object.assign(this.#input, { name: 'pptx-slide-text', spellcheck: false }); // Named, never submitted.
    this.#input.tabIndex = -1; // Not a Tab stop: the slide is; text selection focuses the input.
    // A screen reader hears which cell the caret is in after Tab or an arrow moves it.
    this.#live = doc.createElement('div');
    this.#live.className = 'pptx-sr-only';
    this.#live.setAttribute('aria-live', 'polite');
    // The input sits in the stage, on the caret, so an IME candidate window opens there.
    this.#cursor = new CursorLayer(doc);
    this.#paintRest = new NextFrame(doc.defaultView as Window | null);
    const shown = () => ((l) => l && { slide: l.slide, zoom: editor.zoom() })(editor.layout());
    this.#reach = new paint.WindowReach(doc.defaultView, EMU_PER_PX, shown, () => this.render());
    this.#stage.append(this.#cursor.element, this.#overlay, this.#input);
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
    this.#drag = new ShapeDrag(
      {
        label: (key) => editor.label(key),
        layout: () => editor.layout(),
        transform: (items, gesture) => editor.transform(items, gesture),
        ...(editor.relayout ? { relayout: editor.relayout } : {}),
        ...(editor.paintLayout ? { paintLayout: editor.paintLayout } : {}),
        slide: () => this.#stage.querySelector('.pptx-surface__slide'),
        paintOptions: (layout) => this.#paintOptions(layout),
        repaint: () => this.render(),
        overlay: (layout) => this.#renderOverlay(layout, editor.zoom()),
      },
      doc
    );
    this.#unbindInput = bindSurfaceInput(this.#input, this.root, editor);
    this.root.tabIndex = 0;
    this.root.setAttribute('role', 'group');
    // Pinch (ctrl + wheel on trackpads) and Ctrl/Cmd + wheel zoom around the pointer.
    host.addEventListener('wheel', this.#onWheel, { passive: false });
    this.#resize = new ResizeObserver(() => this.#reportFit());
    this.#resize.observe(host);
    this.#resize.observe(this.root);
  }

  destroy(): void {
    const doc = this.host.ownerDocument;
    doc.removeEventListener('pointermove', this.#onPointerMove);
    doc.removeEventListener('pointerup', this.#onPointerUp);
    doc.removeEventListener('pointercancel', this.#onPointerCancel);
    // A drag in progress ends here: its history group closes with it; a preview writes nothing.
    this.#drag.dispose();
    this.#table.up();
    this.#resize.disconnect();
    this.host.removeEventListener('wheel', this.#onWheel);
    this.#unbindInput();
    this.#paintRest.stop();
    this.#reach.destroy();
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
    this.#fitSlide = `${layout.slide.width}x${layout.slide.height}`;
    this.editor.fitChanged(fitZoom(this.host, this.root, layout.slide));
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
    // A dragged shape that changed under the drag (a peer's edit) ends the drag.
    this.#drag.check(layout);
    this.#syncRotating();
    if (!layout) {
      this.#stage.querySelector('.pptx-surface__slide')?.remove();
      this.#paintKey = null;
      this.#painter = null;
      this.#overlay.replaceChildren();
      // Focus can still land here (Close on an open failure): give the canvas a true name.
      const empty = this.editor.label('canvas.emptyAriaLabel');
      if (this.root.getAttribute('aria-label') !== empty)
        this.root.setAttribute('aria-label', empty);
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
    const said = selectionAnnouncement(layout, this.editor);
    // A canceled drag is said in place of the selection, until the selection changes.
    this.#drag.speak(this.#live, said);
    const w = (layout.slide.width / EMU_PER_PX) * zoom;
    const h = (layout.slide.height / EMU_PER_PX) * zoom;
    setStyle(this.#stage, 'width', `${w}px`);
    setStyle(this.#stage, 'height', `${h}px`);
    if (keep) {
      const st = this.#stage.getBoundingClientRect();
      this.host.scrollLeft += st.left + keep.fx * st.width - keep.client.x;
      this.host.scrollTop += st.top + keep.fy * st.height - keep.client.y;
    }
    const painted = this.editor.paintLayout?.(layout) ?? layout;
    const paintKey = this.editor.paintKey?.();
    // Converted pictures, new faces, or a wider view: paint again, the old nodes standing in.
    if (this.#painter?.invalidate(paintKey)) this.#paintKey = null;
    if (this.#painter?.view(this.#view(layout))) this.#paintKey = null;
    if (this.#paintKey !== painted) {
      this.#paintKey = painted;
      const opts = {
        ...this.#paintOptions(layout),
        ...(this.editor.altText ? { altText: this.editor.altText } : {}),
      };
      // One painter for the view: an edit paints the items it changed, not the whole slide.
      if (!this.#painter?.paints(opts)) {
        this.#painter = new SlidePainter(this.host.ownerDocument, opts);
        this.#painter.invalidate(paintKey);
      }
      this.#drag.beforeRepaint();
      const svg = this.#painter.paint(painted, PAINT_BUDGET);
      // Items left out paint on the next frame, and the painter keeps what it has.
      if (this.#painter.pending)
        this.#paintRest.request(() => {
          if (!this.#painter?.pending) return;
          this.#paintKey = null;
          this.render();
        });
      paint.placeSlide(this.#stage, svg);
      // A repaint during a drag (a peer's edit to another shape) keeps the preview.
      this.#drag.repainted(layout);
    }
    paint.sizeAll([this.#stage.querySelector('.pptx-surface__slide'), this.#overlay], w, h);
    setAttr(this.#overlay, 'viewBox', `0 0 ${layout.slide.width} ${layout.slide.height}`);
    this.#renderOverlay(this.#drag.overlayLayout(layout), zoom);
    if (this.#fitSlide !== `${layout.slide.width}x${layout.slide.height}`) this.#reportFit();
  }

  #renderOverlay(layout: SlideLayout, zoom: number): void {
    const doc = this.host.ownerDocument;
    const ns = SVG_NS;
    const px = EMU_PER_PX / zoom;
    // The area outside the slide, dimmed over the parts of shapes that lie there.
    const parts: Element[] = [
      outsideVeil(doc, layout.slide.width, layout.slide.height),
      ...this.#placeholderChrome.elements(doc, layout, px),
    ];
    const sel = this.editor.selection();
    const ids = sel.kind === 'shapes' ? sel.nodeIds : sel.kind === 'none' ? [] : [sel.nodeId];
    const scope = selectionScope(layout.items, sel);
    const scopeItem = scope ? itemById(layout, scope) : undefined;
    if (scopeItem) {
      // The entered group's box, dashed.
      const g = boxOutline(doc, scopeItem, px);
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
    // One shape of several: a thin line just outside it, under the common box's handles.
    if (selected.length > 1) parts.push(...selected.map((item) => memberOutline(doc, item, px)));
    for (const item of line || selected.length > 1 ? [] : selected) {
      const g = boxOutline(doc, item, px * (sel.kind === 'text' ? 1 : 1.5));
      if (sel.kind === 'text')
        g.firstElementChild?.setAttribute('stroke-dasharray', `${px * 4} ${px * 3}`);
      parts.push(g);
    }
    // Handles show while text is edited too, so a text box can be resized without leaving it.
    if (sel.kind !== 'none' && !crop && !this.editor.readOnly()) {
      parts.push(...handleElements(doc, selected, line, px));
      // One selected preset shape also shows its adjust handles, over the box handles.
      if (sel.kind === 'shapes' && selected.length === 1 && !line && this.editor.adjust)
        parts.push(
          ...adjustHandleElements(doc, selected[0], px, this.editor.label('canvas.adjustHandle'))
        );
    }
    parts.push(...presenceElements(doc, layout, this.editor.presence?.() ?? [], px));
    parts.push(...guideElements(doc, this.#drag.guides, px));
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
    parts.push(...cellsOverlay(doc, layout, sel), ...this.#tableChrome.elements(doc, layout));
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
    // Read the stage position at the next frame, where the browser lays out the page anyway. A
    // read here would force a layout of the whole slide on every keystroke.
    const pending = !this.#caretScroll;
    this.#caretScroll = { caret, scale };
    const view = this.host.ownerDocument.defaultView;
    if (!pending) return;
    const scroll = () => {
      const next = this.#caretScroll;
      this.#caretScroll = null;
      if (!next || !this.root.isConnected) return;
      const c = next.caret;
      const st = this.#stage.getBoundingClientRect();
      const box = {
        left: st.left + Math.min(c.x1, c.x2) * next.scale,
        right: st.left + Math.max(c.x1, c.x2) * next.scale,
        top: st.top + Math.min(c.y1, c.y2) * next.scale,
        bottom: st.top + Math.max(c.y1, c.y2) * next.scale,
      };
      scrollBoxIntoView(this.host, box, 32);
    };
    if (view?.requestAnimationFrame) view.requestAnimationFrame(scroll);
    else scroll();
  }

  /** The caret scroll that waits for the next frame (`#keepCaretInView`). */
  #caretScroll: {
    caret: { x1: number; y1: number; x2: number; y2: number };
    scale: number;
  } | null = null;

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
    const layout = this.editor.layout();
    const zoom = this.editor.zoom();
    // Just outside the selected table, its grid edge comes first: a near miss never takes a
    // shape behind it.
    return this.#table.edgeAt(layout, p, true) ?? hitItem(layout, p, zoom, scope);
  }

  /** Start dragging the selected items (or a handle of their box). */
  #startDrag(
    handle: Handle | 'move',
    ids: readonly string[],
    start: { x: number; y: number }
  ): void {
    this.#drag.start(handle, ids, start, `g${++this.#gestures}`);
    this.#syncRotating();
  }

  /** A rotation drag keeps the rotate pointer, off the handle too (`pptx-surface--rotating`). */
  #syncRotating(): void {
    this.root.classList.toggle('pptx-surface--rotating', this.#drag.rotating);
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
    // An empty placeholder's buttons insert a table or a picture that fills it.
    if (!painting && this.#placeholderChrome.press(layout, target)) {
      this.focus();
      e.preventDefault();
      return;
    }
    // The table chrome: the select handle selects and drags the table; "+" inserts a row or column.
    const chrome = painting ? null : this.#tableChrome.press(layout, target);
    if (chrome) {
      if (chrome.kind === 'move') {
        this.editor.select({ kind: 'shapes', nodeIds: [chrome.id], ...chrome.scope });
        this.#startDrag('move', [chrome.id], p);
      }
      this.focus();
      e.preventDefault();
      return;
    }
    const adjustHandle = target.getAttribute?.('data-adjust-handle');
    if (adjustHandle !== null && adjustHandle !== undefined && !painting && sel.kind === 'shapes') {
      const item = itemById(layout, sel.nodeIds[0] ?? '');
      this.#adjustDrag = startAdjustDrag(item, Number(adjustHandle), `g${++this.#gestures}`);
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

  /** The border band moves a shape; a table's band is outside its grid, so a cell takes a press. */
  #onBorder(item: LayoutItem, p: { x: number; y: number }): boolean {
    if (this.#table.onLine(item, p) || this.#table.inGrid(item, p)) return false;
    return onBorder(item, p, this.editor.zoom());
  }

  /** The pointer shows what a click does: a text cursor over text, a move cursor on shapes. */
  #updateCursor(p: { x: number; y: number }): void {
    // The format painter's brush comes from the stylesheet (`pptx-surface--painting`).
    if (this.editor.painting?.()) {
      this.#setCursor('', p);
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
      this.#setCursor('move', p);
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
    this.#setCursor(cursor, p);
  }

  /** Show `cursor` at slide point `p`, also past the slide's edge (`CursorLayer.show`). */
  #setCursor(cursor: string, p: { x: number; y: number }): void {
    const slide = this.editor.layout()?.slide;
    const outside = !slide || p.x < 0 || p.y < 0 || p.x > slide.width || p.y > slide.height;
    this.#cursor.show(cursor, this.#stage.querySelector('.pptx-surface__slide'), outside);
  }

  /** The point is on linked text of `item`, where a viewer's click opens the link panel. */
  #onLink(item: LayoutItem, p: { x: number; y: number }): boolean {
    return this.#inText(item, p) && linkedAtPoint(item, p.x, p.y);
  }

  #onPointerMove = (e: PointerEvent): void => {
    const layout = this.editor.layout();
    if (!layout) return;
    const p = this.#slidePoint(e);
    const idle =
      !this.#drag.active &&
      !this.#textDrag &&
      !this.#marquee &&
      !this.#cropDrag &&
      !this.#adjustDrag;
    const free = idle && !this.#table.active;
    const inStage = free && e.target instanceof Node && this.#stage.contains(e.target);
    if (idle && this.#tableChrome.hover(layout, p, inStage ? this.#hit(p) : undefined))
      this.#renderOverlay(layout, this.editor.zoom());
    if (inStage) this.#updateCursor(p);
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
    // A press that ended out of sight (a lost pointerup) ends a handle drag: a hover edits nothing.
    if ((this.#adjustDrag || this.#cropDrag) && (e.buttons & 1) === 0) {
      this.#adjustDrag = this.#cropDrag = null;
      this.#pressed = false;
      return;
    }
    const a = this.#adjustDrag;
    if (a) {
      // The drag edits the shape it started on. Escape, an undo, or a peer's edit that changes
      // the selection or removes the shape ends it.
      if (!adjustDragLive(a, this.editor.selection(), layout))
        return void (this.#adjustDrag = null);
      const values = adjustAt(a, p);
      if (values) this.editor.adjust?.(values, a.gesture);
      return;
    }
    const c = this.#cropDrag;
    if (c) {
      if (Math.abs(p.x - c.start.x) + Math.abs(p.y - c.start.y) >= 2 * EMU_PER_PX)
        this.editor.crop?.(cropAt(c, p), c.gesture);
      return;
    }
    // A preview, not an edit: the release writes the boxes (surface-drag.ts).
    this.#drag.move(layout, p, 2 * EMU_PER_PX, {
      shift: e.shiftKey,
      snapTo: this.#snapTo(e),
      glueTo: e.altKey ? null : (2 * SNAP_PX * EMU_PER_PX) / this.editor.zoom(),
      aspectLocked: (id) => !!this.editor.aspectLocked?.(id),
    });
  };

  #paintOptions(layout: SlideLayout) {
    return {
      images: this.editor.imageUrl,
      ...(this.editor.imageSize ? { imageSize: this.editor.imageSize } : {}),
      idPrefix: `${this.editor.idPrefix}s${layout.slide.index}`,
      view: this.#view(layout),
    };
  }

  /** What the surface can show at this zoom (`surfaceView`). */
  #view = (layout: SlideLayout) => this.#reach.view(layout.slide, this.editor.zoom());

  /**
   * End a drag (an undo or a redo). A shape drag ends with no edit, and says so when it moved. A
   * crop or adjust handle drag keeps the steps it wrote, which the undo then reverts.
   */
  cancelDrag(): void {
    this.#drag.cancel();
    this.#adjustDrag = this.#cropDrag = null;
  }

  /** The snap distance in EMU, or null when Alt is down or View > Smart guides is off. */
  #snapTo(e: PointerEvent): number | null {
    if (e.altKey || !this.editor.smartGuides()) return null;
    return (SNAP_PX * EMU_PER_PX) / this.editor.zoom();
  }

  #onPointerUp = (e: PointerEvent): void => {
    const pressed = this.#pressed;
    this.#pressed = false;
    // A release over an empty placeholder's Picture button asks for the file.
    if (pressed) this.#placeholderChrome.release(this.#slidePoint(e));
    this.#endGesture(e);
    if (pressed) this.editor.pointerUp?.();
  };

  /**
   * The browser took the pointer (a touch scroll, a system gesture): a shape drag writes its
   * preview as one undo step, and a crop or grid line drag keeps the steps it applied as one. A
   * marquee selects nothing, and the format painter does not paint.
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
    this.#cropDrag = null;
    this.#adjustDrag = null;
    this.#textDrag = null;
    // After a grid line drag, a caret that stayed in the table takes the keys again.
    if (this.#table.up()) this.focus();
    // A shape drag writes its last boxes as one edit.
    this.#drag.release();
    this.#syncRotating();
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
