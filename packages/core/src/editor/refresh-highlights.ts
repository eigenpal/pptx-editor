// The change flash of a document refresh: what `highlightChanges()` shows, after DOCX Editor's
// refresh paragraph highlights (docx-editor/packages/core/src/editor/paragraph-highlights.ts at
// b63cbea4a), with the timing and options of its anchor highlights.
//
// Three surfaces flash together, from one start time:
//
// - The slide rail and the slide sorter: each thumbnail of a changed slide gets a ring, a light
//   wash, and a corner marker (`data-refresh-highlight` on the `.pptx-rail__thumb` container),
//   and its item gets the accessible description "Changed". The marker is a shape, so color is
//   not the only signal. A thumbnail that mounts during the hold (a scroll, a long rail that
//   mounts in steps) shows the flash for the rest of the hold, without the entrance.
// - The canvas: on the slide on screen, a ring inside the slide edge for a slide change, a ring
//   around each changed shape, and a band over each line of changed text, in a layer of their
//   own under the selection overlay. They follow the slide on screen, the zoom, and the layout.
// - The notes pane: a ring around the pane while the slide on screen has changed notes.
//
// The flash is CSS on chrome the editor owns: it never paints a slide or a thumbnail again, and
// it never reaches the slide paint, the thumbnails' images, the slide show, an export, or a
// saved file. A fade is an opacity transition; reduced motion caps it at 125 milliseconds.
//
// A flash that shows when a new result arrives continues through it. The controller holds the
// flash before the commit (`hold`): while the commit runs, the flash follows each slide by its
// part and draws only nodes that the new layout has, so it never marks a moved or removed
// place. After the commit, `adopt` moves it to the new places of the same shapes and slides.
// A flash request that follows (`show`) starts the places that show already in the shown
// state, and only new places run the entrance.

import type { SlideLayout } from '../layout/index.ts';
import { SVG_NS } from '../output/index.ts';
import { findItem } from '../automation/commands/placement.ts';
import type { PptxEditorInstance } from './pptx-editor-types.ts';
import type {
  AnchorHighlightAnimation,
  ClearRefreshHighlightsOptions,
  RefreshHighlightOptions,
  RefreshHighlightSurface,
} from './document-refresh-types.ts';
import type { RefreshTarget } from './document-refresh-changes.ts';
import { presentationInfo } from '../store/index.ts';
import { selectionRects } from './surface-text-selection.ts';
import { surfaceRootOf } from './surface-registry.ts';
import { translateOf } from './document-replace.ts';

const MARK = 'data-refresh-highlight';
const DEFAULT_COLOR = 'var(--doc-refresh-highlight-color)';
const SURFACES: readonly RefreshHighlightSurface[] = ['thumbnails', 'canvas', 'notes'];
const REDUCED_MS = 125;
const MAX_TIMEOUT = 2147483647;
/** EMU in one CSS pixel at 100% zoom. */
const EMU_PER_PX = 9525;

interface Settings {
  readonly color: string;
  readonly borderColor: string;
  readonly borderStyle: 'solid' | 'dashed' | 'dotted';
  readonly borderWidth: number | null;
  readonly borderRadius: number;
  readonly opacity: number;
  readonly padding: number;
  readonly className: string;
  readonly enterMs: number;
  readonly exitMs: number;
  readonly easing: string;
  readonly timeoutMs: number | null;
  readonly surfaces: ReadonlySet<RefreshHighlightSurface>;
}

const finite = (name: string, v: unknown, min: number, max = Infinity): number => {
  if (typeof v !== 'number' || !Number.isFinite(v))
    throw new RangeError(`${name} must be a finite number.`);
  if (v < min || v > max) throw new RangeError(`${name} must be between ${min} and ${max}.`);
  return v;
};
const text = (name: string, v: unknown): string => {
  if (typeof v !== 'string') throw new TypeError(`${name} must be a string.`);
  return v;
};
/** A color for a CSS custom property: never one that makes the page fetch a URL. */
const color = (name: string, v: unknown): string => {
  const value = text(name, v);
  // CSS escapes can spell a function name, so a backslash is refused as well.
  if (/url\s*\(|image-set\s*\(|image\s*\(|\\/i.test(value))
    throw new TypeError(`${name} cannot load a URL.`);
  return value;
};

/** The settings of `options`, or a thrown TypeError or RangeError. Each call starts from defaults. */
export function refreshSettings(options: RefreshHighlightOptions | undefined): Settings {
  const o = options ?? {};
  if (o.includePrevious !== undefined && typeof o.includePrevious !== 'boolean')
    throw new TypeError('includePrevious must be a boolean.');
  if (o.changeIds !== undefined) {
    if (!Array.isArray(o.changeIds)) throw new TypeError('changeIds must be an array.');
    if (o.changeIds.length > 10000) throw new RangeError('changeIds holds at most 10000 ids.');
    if (o.changeIds.some((id) => typeof id !== 'string' || id === ''))
      throw new TypeError('changeIds must hold nonempty strings.');
  }
  let surfaces: ReadonlySet<RefreshHighlightSurface> = new Set(SURFACES);
  if (o.surfaces !== undefined) {
    if (!Array.isArray(o.surfaces) || o.surfaces.some((s) => !SURFACES.includes(s)))
      throw new TypeError("surfaces must hold 'thumbnails', 'canvas', or 'notes'.");
    surfaces = new Set(o.surfaces);
  }
  const fill = o.color === undefined ? DEFAULT_COLOR : color('color', o.color);
  const style = o.borderStyle ?? 'solid';
  if (style !== 'solid' && style !== 'dashed' && style !== 'dotted')
    throw new TypeError("borderStyle must be 'solid', 'dashed', or 'dotted'.");
  let anim: AnchorHighlightAnimation | false = {};
  if (o.animation === false) anim = false;
  else if (o.animation !== undefined && o.animation !== true) {
    if (typeof o.animation !== 'object' || o.animation === null)
      throw new TypeError('animation must be a boolean or an object.');
    anim = o.animation;
  }
  const enter =
    anim === false ? 0 : finite('animation.durationMs', anim.durationMs ?? 180, 0, 10000);
  const exit =
    anim === false ? 0 : finite('animation.exitDurationMs', anim.exitDurationMs ?? enter, 0, 10000);
  const easing =
    anim === false || anim.easing === undefined
      ? 'var(--doc-motion-ease-out, ease-out)'
      : text('animation.easing', anim.easing);
  if (anim !== false && anim.easing !== undefined && /var\s*\(/i.test(easing))
    throw new TypeError('animation.easing cannot use CSS variables.');
  return {
    color: fill,
    borderColor: o.borderColor === undefined ? fill : color('borderColor', o.borderColor),
    borderStyle: style,
    borderWidth: o.borderWidth === undefined ? null : finite('borderWidth', o.borderWidth, 0),
    borderRadius: finite('borderRadius', o.borderRadius ?? 6, 0),
    opacity: finite('opacity', o.opacity ?? 0.14, 0, 1),
    padding: finite('padding', o.padding ?? 4, 0),
    className: o.className === undefined ? '' : text('className', o.className),
    enterMs: enter,
    exitMs: exit,
    easing,
    timeoutMs:
      o.timeoutMs === null ? null : finite('timeoutMs', o.timeoutMs ?? 3000, 0, MAX_TIMEOUT),
    surfaces,
  };
}

/** Validate dismissal options. */
function clearSettings(options: ClearRefreshHighlightsOptions | undefined, last: Settings | null) {
  const a = options?.animation;
  if (a === false) return { exitMs: 0, easing: last?.easing ?? 'ease-out' };
  if (a !== undefined && a !== true && (typeof a !== 'object' || a === null))
    throw new TypeError('animation must be a boolean or an object.');
  const o = typeof a === 'object' ? a : {};
  const exitMs =
    o.exitDurationMs === undefined
      ? (last?.exitMs ?? 180)
      : finite('animation.exitDurationMs', o.exitDurationMs, 0, 10000);
  const easing =
    o.easing === undefined ? (last?.easing ?? 'ease-out') : text('animation.easing', o.easing);
  if (o.easing !== undefined && /var\s*\(/i.test(easing))
    throw new TypeError('animation.easing cannot use CSS variables.');
  return { exitMs, easing };
}

function reducedMotion(win: Window | null | undefined): boolean {
  try {
    return !!win?.matchMedia?.('(prefers-reduced-motion: reduce)').matches;
  } catch {
    return false;
  }
}

type Phase = 'entering' | 'shown' | 'exiting';

/** The place of a target: its part, notes or not, and its shape. */
const placeOf = (t: RefreshTarget): string => `${t.part}|${t.notes ? 1 : 0}|${t.nodeId ?? ''}`;
/** The place of a target with its text range. */
const keyOf = (t: RefreshTarget): string =>
  t.text ? `${placeOf(t)}|${t.text.paragraph}:${t.text.start}:${t.text.end}` : placeOf(t);

/** What showed before a flash request that continues the flash. */
interface Kept {
  /** The keys of the targets (`keyOf`). */
  readonly keys: ReadonlySet<string>;
  /** The slide indexes whose thumbnails showed. */
  readonly slides: ReadonlySet<number>;
  /** The slide indexes whose notes showed. */
  readonly notes: ReadonlySet<number>;
  /** The canvas layer showed marks. */
  readonly layer: boolean;
}

/** The change flash of one editor. @internal */
export class RefreshHighlights {
  #targets: readonly RefreshTarget[] = [];
  #settings: Settings | null = null;
  #phase: Phase | null = null;
  #timer: ReturnType<typeof setTimeout> | null = null;
  #exitTimer: ReturnType<typeof setTimeout> | null = null;
  #layer: SVGSVGElement | null = null;
  #marked = new Set<Element>();
  #observer: MutationObserver | null = null;
  #unsubscribe: (() => void) | null = null;
  /** The `aria-description` each marked item had before the flash: null for none. */
  #descriptions = new Map<Element, string | null>();
  #disposed = false;
  /** A commit runs: the flash follows slides by part (`hold`). */
  #holding = false;
  /** What showed before the current request, while its entrance runs. */
  #kept: Kept | null = null;

  constructor(
    private readonly editor: PptxEditorInstance,
    private readonly onExpire: () => void
  ) {}

  validate(options: RefreshHighlightOptions | undefined): Settings {
    return refreshSettings(options);
  }

  /** Flash `targets`. An empty list fades the current flash out. */
  show(targets: readonly RefreshTarget[], settings: Settings): void {
    if (this.#disposed) return;
    if (targets.length === 0) return this.hide();
    const root = surfaceRootOf(this.editor);
    const win = root?.ownerDocument.defaultView;
    const reduced = reducedMotion(win);
    const cap = (ms: number) => (reduced ? Math.min(ms, REDUCED_MS) : ms);
    // Shown already: the timeout starts again, and the places that show do not enter again.
    const before = this.visible ? this.#targets : null;
    this.#holding = false;
    this.#kept = before
      ? {
          keys: new Set(before.map(keyOf)),
          slides: new Set(before.map((t) => t.slideIndex)),
          notes: new Set(before.filter((t) => t.notes).map((t) => t.slideIndex)),
          layer: !!this.#layer?.firstChild && this.#layer.getAttribute(MARK) !== 'exiting',
        }
      : null;
    this.#settings = { ...settings, enterMs: cap(settings.enterMs), exitMs: cap(settings.exitMs) };
    this.#targets = targets;
    this.#clearTimers();
    const kept = this.#kept;
    const fresh = !kept || targets.some((t) => !kept.keys.has(keyOf(t)));
    this.#phase = fresh && this.#settings.enterMs > 0 ? 'entering' : 'shown';
    if (this.#phase === 'shown') this.#kept = null;
    this.#watch();
    this.draw();
    if (this.#phase === 'entering') {
      const next = () => {
        if (this.#phase !== 'entering') return;
        this.#phase = 'shown';
        this.#kept = null;
        this.#applyPhase();
      };
      // The next frame, so the browser sees the start state and runs the transition.
      if (win?.requestAnimationFrame)
        win.requestAnimationFrame(() => win.requestAnimationFrame(next));
      else setTimeout(next, 0);
    }
    if (this.#settings.timeoutMs !== null)
      this.#timer = setTimeout(() => {
        this.#timer = null;
        this.hide();
        this.onExpire();
      }, this.#settings.timeoutMs);
  }

  /** Fade the flash out, then remove it. */
  hide(options?: ClearRefreshHighlightsOptions): void {
    const { exitMs, easing } = clearSettings(options, this.#settings);
    if (this.#phase === null) return;
    this.#clearTimers();
    if (this.#settings) this.#settings = { ...this.#settings, exitMs, easing };
    if (exitMs === 0) return this.clear();
    this.#phase = 'exiting';
    this.#applyPhase();
    this.#exitTimer = setTimeout(() => this.clear(), exitMs);
  }

  /**
   * Keep the flash through a commit: until `adopt` or `release`, it follows each slide by its
   * part and draws only nodes that the layout has. A flash that does not show goes.
   */
  hold(): void {
    if (!this.visible) return this.clear();
    this.#holding = true;
  }

  /** End a hold with no commit: the document did not change. */
  release(): void {
    this.#holding = false;
  }

  get holding(): boolean {
    return this.#holding;
  }

  /**
   * End a hold after the commit: each held target moves to the targets of the accepted result
   * at the same place (`placeOf`), with their text ranges. A place the result does not name
   * goes. Returns whether the flash still shows.
   */
  adopt(targets: readonly RefreshTarget[]): boolean {
    if (!this.#holding) return this.visible;
    this.#holding = false;
    const held = new Set(this.#targets.map(placeOf));
    const next = targets.filter((t) => held.has(placeOf(t)));
    if (!this.visible || next.length === 0) {
      this.clear();
      return false;
    }
    const seen = new Set<string>();
    this.#targets = next.filter((t) => !seen.has(keyOf(t)) && !!seen.add(keyOf(t)));
    this.draw();
    return true;
  }

  /** Move each held target to the index of its part; a removed part takes its target away. */
  #followParts(): void {
    const pkg = this.editor.session?.store.package;
    if (!pkg) return;
    const at = new Map<string, number>();
    presentationInfo(pkg).slides.forEach((slide, i) => {
      at.set(slide.partName, i);
      if (slide.notesPart) at.set(slide.notesPart, i);
    });
    this.#targets = this.#targets.flatMap((t) => {
      const i = at.get(t.part);
      if (i === undefined) return [];
      return i === t.slideIndex ? [t] : [{ ...t, slideIndex: i }];
    });
  }

  /** Remove the flash at once: a document edit, a reload, a new result. */
  clear(): void {
    this.#clearTimers();
    this.#holding = false;
    this.#kept = null;
    this.#phase = null;
    this.#targets = [];
    for (const el of this.#marked) this.#unmark(el);
    this.#marked.clear();
    this.#layer?.remove();
    this.#layer = null;
    this.#observer?.disconnect();
    this.#observer = null;
    this.#unsubscribe?.();
    this.#unsubscribe = null;
  }

  /** Remove the flash and show none again: the editor is destroyed. */
  dispose(): void {
    this.clear();
    this.#disposed = true;
  }

  get visible(): boolean {
    return this.#phase !== null && this.#phase !== 'exiting';
  }

  #clearTimers(): void {
    if (this.#timer) clearTimeout(this.#timer);
    if (this.#exitTimer) clearTimeout(this.#exitTimer);
    this.#timer = this.#exitTimer = null;
  }

  /** Follow the slide on screen, the zoom, and thumbnails that mount during the hold. */
  #watch(): void {
    if (!this.#unsubscribe)
      this.#unsubscribe = this.editor.subscribe(() => {
        if (this.#holding) this.#followParts();
        this.draw();
      });
    const root = this.#editorRoot();
    const MO = root?.ownerDocument.defaultView?.MutationObserver;
    if (!this.#observer && root && MO) {
      this.#observer = new MO((records) => {
        // Only new rail items: the flash's own attribute writes are not child changes.
        if (records.some((r) => r.addedNodes.length > 0)) this.#drawThumbnails(false);
      });
      this.#observer.observe(root, { childList: true, subtree: true });
    }
  }

  #editorRoot(): HTMLElement | null {
    const root = surfaceRootOf(this.editor);
    return (root?.closest('.pptx-editor') as HTMLElement | null) ?? root;
  }

  /** Draw every surface for the current state. */
  draw(): void {
    if (this.#phase === null || !this.#settings) return;
    this.#drawThumbnails(true);
    this.#drawCanvas();
    this.#drawNotes();
  }

  #style(el: Element & ElementCSSInlineStyle, s: Settings, borderWidth: number): void {
    el.style.setProperty('--pptx-refresh-color', s.color);
    el.style.setProperty('--pptx-refresh-border-color', s.borderColor);
    el.style.setProperty('--pptx-refresh-border-style', s.borderStyle);
    el.style.setProperty('--pptx-refresh-border-width', `${borderWidth}px`);
    el.style.setProperty('--pptx-refresh-radius', `${s.borderRadius}px`);
    el.style.setProperty('--pptx-refresh-opacity', String(s.opacity));
    el.style.setProperty('--pptx-refresh-padding', `${s.padding}px`);
    el.style.setProperty('--pptx-refresh-enter', `${s.enterMs}ms`);
    el.style.setProperty('--pptx-refresh-exit', `${s.exitMs}ms`);
    el.style.setProperty('--pptx-refresh-easing', s.easing);
    el.setAttribute('data-refresh-border', s.borderStyle);
  }

  /** The phase of a mark: one with no entrance, or one that showed already, starts shown. */
  #phaseOf(entrance: boolean): Phase {
    return !entrance && this.#phase === 'entering' ? 'shown' : this.#phase!;
  }

  #mark(el: HTMLElement, kind: string, entrance: boolean): void {
    const s = this.#settings!;
    const phase = this.#phaseOf(entrance);
    if (!this.#marked.has(el)) {
      this.#marked.add(el);
      this.#style(el, s, s.borderWidth ?? 2);
      for (const name of s.className.split(/\s+/).filter(Boolean)) el.classList.add(name);
      el.classList.add(`pptx-refresh-${kind}`);
    }
    el.setAttribute(MARK, phase);
  }

  #unmark(el: Element): void {
    el.removeAttribute(MARK);
    // The adapter's own description comes back.
    if (this.#descriptions.has(el)) {
      const before = this.#descriptions.get(el);
      this.#descriptions.delete(el);
      if (before === null || before === undefined) el.removeAttribute('aria-description');
      else el.setAttribute('aria-description', before);
    }
    el.removeAttribute('data-refresh-border');
    for (const name of (this.#settings?.className ?? '').split(/\s+/).filter(Boolean))
      el.classList.remove(name);
    el.classList.remove('pptx-refresh-thumbnail', 'pptx-refresh-notes', 'pptx-refresh-item');
  }

  #applyPhase(): void {
    for (const el of this.#marked) if (el.hasAttribute(MARK)) el.setAttribute(MARK, this.#phase!);
    this.#layer?.setAttribute(MARK, this.#phase!);
    // The canvas marks of new places in a layer that showed already.
    for (const g of this.#layer?.querySelectorAll(`[${MARK}]`) ?? [])
      g.setAttribute(MARK, this.#phase!);
  }

  #drawThumbnails(entrance: boolean): void {
    const s = this.#settings;
    const root = this.#editorRoot();
    if (!s || !root || this.#phase === null) return;
    const snap = this.editor.snapshot();
    const slides = new Set(
      s.surfaces.has('thumbnails') && snap.presentationView !== 'slideMaster'
        ? this.#targets.map((t) => t.slideIndex)
        : []
    );
    const label = translateOf(this.editor)?.('refresh.changed') ?? 'Changed';
    const want = new Set<Element>();
    for (const index of slides)
      for (const item of root.querySelectorAll<HTMLElement>(`[data-index="${index}"]`)) {
        const thumb = item.querySelector<HTMLElement>('.pptx-rail__thumb');
        if (!thumb) continue;
        want.add(thumb);
        want.add(item);
        const enters = entrance && !this.#kept?.slides.has(index);
        this.#mark(thumb, 'thumbnail', enters);
        if (!this.#marked.has(item)) {
          this.#marked.add(item);
          item.classList.add('pptx-refresh-item');
        }
        item.setAttribute(MARK, this.#phaseOf(enters));
        if (!this.#descriptions.has(item))
          this.#descriptions.set(item, item.getAttribute('aria-description'));
        item.setAttribute('aria-description', label);
      }
    // A slide that moved since, or a surface turned off: its old marks go.
    for (const el of [...this.#marked])
      if (
        (el.classList.contains('pptx-refresh-thumbnail') ||
          el.classList.contains('pptx-refresh-item')) &&
        !want.has(el)
      ) {
        this.#unmark(el);
        this.#marked.delete(el);
      }
  }

  #drawCanvas(): void {
    const s = this.#settings!;
    const root = surfaceRootOf(this.editor);
    const stage = root?.querySelector<HTMLElement>('.pptx-surface__stage');
    const session = this.editor.session;
    const snap = this.editor.snapshot();
    const layout = session?.layout(snap.currentSlide);
    const targets = s.surfaces.has('canvas')
      ? this.#targets.filter((t) => !t.notes && t.slideIndex === snap.currentSlide)
      : [];
    // The master view shows a layout, not the slide the targets name.
    if (
      !stage ||
      !layout ||
      targets.length === 0 ||
      snap.presentationView === 'slideSorter' ||
      snap.presentationView === 'slideMaster'
    ) {
      this.#layer?.replaceChildren();
      return;
    }
    const doc = stage.ownerDocument;
    if (!this.#layer) {
      this.#layer = doc.createElementNS(SVG_NS, 'svg') as SVGSVGElement;
      this.#layer.setAttribute('class', 'pptx-surface__refresh');
      this.#layer.setAttribute('aria-hidden', 'true');
      this.#style(this.#layer, s, s.borderWidth ?? 2);
    }
    const overlay = stage.querySelector(':scope > .pptx-surface__overlay');
    if (this.#layer.parentNode !== stage) stage.insertBefore(this.#layer, overlay);
    // A layer that showed marks stays shown: only the marks of new places enter.
    const kept = this.#kept?.layer ? this.#kept : null;
    this.#layer.setAttribute(MARK, this.#phaseOf(!kept));
    const { width, height } = layout.slide;
    this.#layer.setAttribute('viewBox', `0 0 ${width} ${height}`);
    this.#layer.setAttribute('preserveAspectRatio', 'none');
    const px = EMU_PER_PX / (snap.zoom || 1);
    const enters =
      kept && this.#phase === 'entering' ? (t: RefreshTarget) => !kept.keys.has(keyOf(t)) : null;
    this.#layer.replaceChildren(...canvasMarks(doc, layout, targets, s, px, enters));
  }

  #drawNotes(): void {
    const s = this.#settings!;
    const root = this.#editorRoot();
    const snap = this.editor.snapshot();
    const pane = root?.querySelector<HTMLElement>('.pptx-notes') ?? null;
    const on =
      s.surfaces.has('notes') &&
      this.#targets.some((t) => t.notes && t.slideIndex === snap.currentSlide);
    for (const el of [...this.#marked])
      if (el.classList.contains('pptx-refresh-notes') && (el !== pane || !on)) {
        this.#unmark(el);
        this.#marked.delete(el);
      }
    if (pane && on) this.#mark(pane, 'notes', !this.#kept?.notes.has(snap.currentSlide));
  }
}

/**
 * The canvas marks of `targets` on `layout`: slide rings, shape rings, and text bands. The
 * marks of a target that `enters` names start in the entering state.
 */
function canvasMarks(
  doc: Document,
  layout: SlideLayout,
  targets: readonly RefreshTarget[],
  s: Settings,
  px: number,
  enters: ((t: RefreshTarget) => boolean) | null
): Element[] {
  const out: Element[] = [];
  for (const t of targets) {
    const marks = targetMarks(doc, layout, t, s, px);
    if (enters?.(t)) for (const el of marks) el.setAttribute(MARK, 'entering');
    out.push(...marks);
  }
  return out;
}

/** The canvas marks of one target: a slide ring, a shape ring, or text bands. */
function targetMarks(
  doc: Document,
  layout: SlideLayout,
  t: RefreshTarget,
  s: Settings,
  px: number
): Element[] {
  const stroke = (s.borderWidth ?? 2) * px;
  const pad = s.padding * px;
  const radius = s.borderRadius * px;
  if (!t.nodeId) {
    const rect = doc.createElementNS(SVG_NS, 'rect');
    rect.setAttribute('class', 'pptx-refresh-ring pptx-refresh-ring--slide');
    rect.setAttribute('x', String(stroke / 2));
    rect.setAttribute('y', String(stroke / 2));
    rect.setAttribute('width', String(Math.max(0, layout.slide.width - stroke)));
    rect.setAttribute('height', String(Math.max(0, layout.slide.height - stroke)));
    rect.setAttribute('rx', String(radius));
    rect.setAttribute('stroke-width', String(stroke));
    return [rect];
  }
  const item = findItem(layout.items, t.nodeId);
  if (!item) return [];
  if (t.text && t.text.end > t.text.start) {
    const from = { paragraph: t.text.paragraph, offset: t.text.start };
    const to = { paragraph: t.text.paragraph, offset: t.text.end };
    const rects = selectionRects(doc, item, from, to, 'pptx-refresh-band');
    for (const g of rects) g.setAttribute('class', 'pptx-refresh-band-group');
    if (rects.length > 0) return rects;
  }
  const { cx, cy } = item.element.xfrm;
  const [a, b, c, d] = item.matrix;
  const sx = Math.hypot(a, b) || 1;
  const sy = Math.hypot(c, d) || 1;
  const g = doc.createElementNS(SVG_NS, 'g');
  g.setAttribute('class', 'pptx-refresh-ring-group');
  g.setAttribute('transform', `matrix(${item.matrix.join(' ')})`);
  const rect = doc.createElementNS(SVG_NS, 'rect');
  rect.setAttribute('class', 'pptx-refresh-ring');
  rect.setAttribute('x', String(-pad / sx));
  rect.setAttribute('y', String(-pad / sy));
  rect.setAttribute('width', String(cx + (2 * pad) / sx));
  rect.setAttribute('height', String(Math.max(cy, 1) + (2 * pad) / sy));
  rect.setAttribute('rx', String(radius / sx));
  rect.setAttribute('ry', String(radius / sy));
  rect.setAttribute('stroke-width', String(stroke / Math.sqrt(sx * sy)));
  g.append(rect);
  return [g];
}
