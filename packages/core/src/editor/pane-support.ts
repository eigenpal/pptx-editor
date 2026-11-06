// Framework-neutral behavior of the editor's panes: the slide rail, the slide sorter, their
// section headers, drag to reorder, and the speaker notes field. Each adapter keeps only its
// glue (state, effects, elements), and calls these for the keys, the drops, and the focus
// moves, so the React and Vue panes act the same. They take a DOM `KeyboardEvent` or a React
// synthetic one (`PaneKeyEvent`).
//
// Every export is adapter glue, not a stable API: `@internal`.

import type { Translate } from '@pptx-editor.dev/i18n';
import type { NotesFormat, SlideSection } from '../contracts/index.ts';
import type { PptxEditorInstance } from './pptx-editor.ts';
import { runSlideStep, sectionOfSlide, stepCommand, type SlideStep } from './adapter-support.ts';
import { focusSurfaceOf } from './surface-registry.ts';
import {
  railOrientationOf,
  railParentKey,
  railStepOfKey,
  scrollIntoNearestScroller,
  type SlideRailOrientation,
} from './rail-orientation.ts';
import {
  insertNotesText,
  notesCommittedOffset,
  notesFieldText,
  typeArmedNotes,
  typeAtParagraphStart,
} from './notes-dom.ts';

/** The fields of a key event that the pane handlers read. @internal */
export interface PaneKeyEvent {
  readonly key: string;
  readonly metaKey: boolean;
  readonly ctrlKey: boolean;
  readonly shiftKey: boolean;
  readonly altKey: boolean;
  readonly target: EventTarget | null;
  preventDefault(): void;
  stopPropagation(): void;
}

const NO_IDS: readonly string[] = [];

/** How a click selects slides: Shift+click a range, Ctrl/Cmd+click one more or one less. @internal */
export function slideClickMode(e: {
  shiftKey: boolean;
  metaKey: boolean;
  ctrlKey: boolean;
}): 'range' | 'toggle' | 'replace' {
  return e.shiftKey ? 'range' : e.metaKey || e.ctrlKey ? 'toggle' : 'replace';
}

/**
 * The accessible name of a slide item in the rail and the slide sorter: "Slide 3", or "Slide 3
 * (hidden)". The item gets it from hidden text, not `aria-label`, so the visible number is part
 * of the name and the thumbnail art, hidden from assistive technology, is not. @internal
 */
export function slideItemName(
  t: Translate,
  slide: { readonly index: number; readonly hidden: boolean }
): string {
  return t(slide.hidden ? 'rail.hiddenSlideLabel' : 'rail.slideLabel', { index: slide.index + 1 });
}

/** The zoom levels of the zoom pill's menu, after Fit. @internal */
export const ZOOM_MENU_LEVELS: readonly number[] = [0.5, 0.75, 1, 1.5, 2];

// --- Drag to reorder ----------------------------------------------------------------------------

/** A drag in progress in a slide list: the slides that move, and the slide and side under the pointer. @internal */
export interface SlideDrag {
  readonly moving: readonly number[];
  readonly over: number;
  readonly side: 'before' | 'after';
}

/**
 * Start a drag on slide `index`. Dragging a selected slide moves the whole selection; another
 * slide moves alone, and becomes the current slide. @internal
 */
export function startSlideDrag(
  editor: PptxEditorInstance | null,
  selected: readonly number[],
  index: number
): SlideDrag {
  const moving = selected.includes(index) ? selected : [index];
  if (moving.length === 1) editor?.goToSlide(index);
  return { moving, over: index, side: 'before' };
}

/**
 * The drag after the pointer moved over slide `index` (box `rect`). `axis` is the direction in
 * which the list reads: `y` for the rail, `x` for the sorter's rows. @internal
 */
export function dragOverSlide(
  drag: SlideDrag,
  index: number,
  axis: 'x' | 'y',
  rect: { left: number; top: number; width: number; height: number },
  point: { clientX: number; clientY: number }
): SlideDrag {
  const before =
    axis === 'y'
      ? point.clientY < rect.top + rect.height / 2
      : point.clientX < rect.left + rect.width / 2;
  return { ...drag, over: index, side: before ? 'before' : 'after' };
}

/** The drop marker of slide `index` (`data-drop`), or undefined. @internal */
export function slideDropMarker(
  drag: SlideDrag | null,
  index: number
): SlideDrag['side'] | undefined {
  return drag && drag.over === index && !drag.moving.includes(index) ? drag.side : undefined;
}

/**
 * Drop the drag: one `moveSlides` command, so one undo step. With sections, the drop joins the
 * section of the slide under the pointer: a drop before a section's first slide (below its
 * header) puts the block into that section. @internal
 */
export function dropSlideDrag(
  editor: PptxEditorInstance | null,
  drag: SlideDrag | null,
  grouped: readonly SlideSection[] | null
): void {
  if (!drag || !editor) return;
  // The drop point, counted among the slides that do not move.
  const point = drag.side === 'before' ? drag.over : drag.over + 1;
  let to = 0;
  for (let i = 0; i < point; i += 1) if (!drag.moving.includes(i)) to += 1;
  // The slide under the pointer stays where it is, so its section is next to the block.
  const section = drag.moving.includes(drag.over)
    ? undefined
    : sectionOfSlide(grouped, drag.over)?.id;
  editor.exec({ type: 'moveSlides', indices: drag.moving, to, ...(section ? { section } : {}) });
}

// --- Focus --------------------------------------------------------------------------------------

/**
 * Focus the current slide of `root`, or the tab stop of its tree (the header of a shut section
 * that holds the current slide). @internal
 */
export function focusCurrentSlideIn(root: HTMLElement | null | undefined): void {
  const target =
    root?.querySelector<HTMLElement>('[aria-current="true"]') ??
    root?.querySelector<HTMLElement>('[role="treeitem"][tabindex="0"]');
  target?.focus({ preventScroll: true });
}

/** Focus `el` and say whether it took focus: a hidden element (a closed overlay) does not. */
function tookFocus(el: HTMLElement | null | undefined): boolean {
  if (!el) return false;
  el.focus({ preventScroll: true });
  return el.ownerDocument.activeElement === el;
}

/**
 * Focus the current slide after a view change took the focused part away: the current slide in
 * the slide sorter, else in the slide rail, else the slide itself. A rail that is absent, or
 * hidden (a narrow window shows it only as an overlay), cannot take focus, so focus goes to the
 * slide and never stays on the page body. `root` holds the views. @internal
 */
export function focusCurrentSlideOrCanvas(
  root: HTMLElement | null | undefined,
  editor: PptxEditorInstance | null
): void {
  if (!root) return;
  for (const list of ['.pptx-sorter__list', '.pptx-rail__list']) {
    const el = root.querySelector<HTMLElement>(list);
    const target =
      el?.querySelector<HTMLElement>('[aria-current="true"]') ??
      el?.querySelector<HTMLElement>('[role="treeitem"][tabindex="0"]');
    if (tookFocus(target)) return;
  }
  if (editor && focusSurfaceOf(editor)) return;
  tookFocus(root.querySelector<HTMLElement>('.pptx-surface'));
}

/**
 * A slide sorter that appears (View > Slide sorter) takes focus on its current slide when focus
 * is on nothing, on a removed element, or in the menu that opened it; focus a host put
 * elsewhere stays. It checks again after the next frame, when the menu has closed. Returns the
 * function that cancels the check. @internal
 */
export function focusSorterOnShow(list: HTMLElement): () => void {
  const doc = list.ownerDocument;
  const go = (): void => {
    const active = doc.activeElement;
    const lost =
      !active ||
      active === doc.body ||
      !active.isConnected ||
      active.closest('[role="menu"], [role="menubar"]') !== null;
    if (lost && list.isConnected) focusCurrentSlideIn(list);
  };
  go();
  const view = doc.defaultView;
  if (view && typeof view.requestAnimationFrame === 'function') {
    const frame = view.requestAnimationFrame(go);
    return () => view.cancelAnimationFrame(frame);
  }
  const timer = setTimeout(go, 0);
  return () => clearTimeout(timer);
}

/**
 * After sections collapse, the element `was` that had focus can be gone: focus moves to the
 * current slide, or to the header of its section. Focus that went elsewhere stays. @internal
 */
export function refocusAfterCollapse(list: HTMLElement | null, was: HTMLElement | null): void {
  if (!was || was.isConnected || !list) return;
  const active = was.ownerDocument.activeElement;
  if (active && active !== was.ownerDocument.body) return;
  focusCurrentSlideIn(list);
}

/**
 * Keyboard focus follows the current slide of a slide list while one of its items (`item`, a
 * selector) has focus. When the focused item `focused` went away, focus goes to the current
 * item too, so the next key still acts on the list. @internal
 */
export function followCurrentSlide(
  list: HTMLElement | null,
  item: string,
  focused: HTMLElement | null
): void {
  const target = list?.querySelector<HTMLElement>(`${item}[aria-current="true"]`);
  // The rail scrolls, not the page: the editor can sit in a panel that slides in from off screen.
  if (target) scrollIntoNearestScroller(target, list?.closest('.pptx-rail, .pptx-sorter') ?? list);
  const active = target?.ownerDocument.activeElement;
  const onItem = !!active?.closest?.(item) && !!list?.contains(active);
  const gone =
    !!focused && !focused.isConnected && (!active || active === active.ownerDocument.body);
  if (target && (onItem || gone)) target.focus({ preventScroll: true });
}

/** Scroll the current slide of `list` into view in the rail or the sorter, never the page. */
function revealCurrentSlide(list: HTMLElement, item: string): void {
  const target = list.querySelector<HTMLElement>(`${item}[aria-current="true"]`);
  if (target) scrollIntoNearestScroller(target, list.closest('.pptx-rail, .pptx-sorter') ?? list);
}

/**
 * Keep the current slide of a slide list in view while the list changes size, as when a host
 * narrows the panel that holds a horizontal rail. `followCurrentSlide` covers a change of the
 * current slide; this covers a resize. Returns the function that stops it. Without
 * `ResizeObserver`, it does nothing. @internal
 */
export function keepCurrentSlideInView(list: HTMLElement | null, item: string): () => void {
  if (!list || typeof ResizeObserver !== 'function') return () => undefined;
  const observer = new ResizeObserver(() => revealCurrentSlide(list, item));
  observer.observe(list);
  return () => observer.disconnect();
}

// --- Sections -----------------------------------------------------------------------------------

/**
 * A section selected by its name, and the current slide that the selection made: that change
 * of the current slide leaves the section as it is, collapsed or not.
 */
const selectedByName = new WeakMap<PptxEditorInstance, { section: string; slide: number }>();

/**
 * Select every slide of a section, as a click on its name does. The first slide of the section
 * is the current slide: the range runs from the last slide (the anchor) up to the first, so a
 * later Shift+click or Shift+Arrow extends from the far end, as any range does. @internal
 */
export function selectSlideSection(editor: PptxEditorInstance | null, section: SlideSection): void {
  if (!editor || !section.slideCount) return;
  editor.selectSlide(section.firstSlide + section.slideCount - 1);
  editor.selectSlide(section.firstSlide, 'range');
  selectedByName.set(editor, { section: section.id, slide: editor.snapshot().currentSlide });
}

/**
 * The current slide changed to `current`: its section opens, so the current slide is in view,
 * unless a click on the section's name made it current. @internal
 */
export function openSectionOfCurrent(
  editor: PptxEditorInstance,
  grouped: readonly SlideSection[] | null,
  current: number
): void {
  const own = sectionOfSlide(grouped, current);
  const named = selectedByName.get(editor);
  if (named && named.slide === current && named.section === own?.id) return;
  selectedByName.delete(editor);
  if (own && (editor.snapshot().collapsedSections ?? NO_IDS).includes(own.id))
    editor.exec({ type: 'setSectionsCollapsed', collapsed: false, sectionIds: [own.id] });
}

/** The slide sorters of each editor that are mounted. */
const sorters = new WeakMap<PptxEditorInstance, number>();

/** Count a mounted slide sorter of `editor`. Returns the function that uncounts it. @internal */
export function registerSlideSorter(editor: PptxEditorInstance): () => void {
  sorters.set(editor, (sorters.get(editor) ?? 0) + 1);
  return () => void sorters.set(editor, (sorters.get(editor) ?? 1) - 1);
}

/** Whether a slide sorter of `editor` is mounted, in either adapter. @internal */
export function slideSorterMounted(editor: PptxEditorInstance): boolean {
  return (sorters.get(editor) ?? 0) > 0;
}

/** What a section header's keys need. @internal */
export interface SectionKeyContext {
  readonly editor: PptxEditorInstance | null;
  readonly readOnly: boolean;
  readonly section: SlideSection;
  readonly collapsed: boolean;
  /** Collapse or expand the section. */
  toggle(): void;
  /** Start renaming the section in place. */
  rename(): void;
  /** Open the section menu at the header, from the keyboard. */
  openMenu(owner: HTMLElement): void;
}

/**
 * The keys of a focused section header (`node`): Right and Left expand and collapse, Enter
 * and Space select its slides, F2 renames, Shift+F10 and the Menu key open the menu. In a
 * horizontal rail, Left and Right step along the row, so Down expands and Up collapses.
 * @internal
 */
export function sectionHeaderKeyDown(
  e: PaneKeyEvent,
  node: HTMLElement,
  ctx: SectionKeyContext
): void {
  if (e.target !== node) return;
  const { editor, section, collapsed } = ctx;
  const row = railOrientationOf(node) === 'horizontal';
  const expand = row ? 'ArrowDown' : 'ArrowRight';
  const collapse = row ? 'ArrowUp' : 'ArrowLeft';
  if (e.key === 'ContextMenu' || (e.shiftKey && e.key === 'F10')) ctx.openMenu(node);
  else if (e.key === 'F2' && !ctx.readOnly) ctx.rename();
  else if (e.key === 'Enter' || e.key === ' ') selectSlideSection(editor, section);
  else if (e.key === expand && collapsed) ctx.toggle();
  else if (e.key === expand && section.slideCount) {
    // Into the section: its first slide becomes current, so slide keys act on what has focus.
    editor?.goToSlide(section.firstSlide);
    node.querySelector<HTMLElement>('.pptx-rail__item, .pptx-sorter__item')?.focus();
  } else if (e.key === collapse && !collapsed) ctx.toggle();
  else return;
  e.preventDefault();
  e.stopPropagation();
}

/**
 * Collapse or expand a section (`node` is its header). A slide of the section that has focus
 * goes away when it collapses: its header takes focus. @internal
 */
export function toggleSlideSection(
  editor: PptxEditorInstance | null,
  section: SlideSection,
  collapsed: boolean,
  node: HTMLElement | null
): void {
  const active = node?.ownerDocument.activeElement;
  const inside = !collapsed && !!active && active !== node && !!node?.contains(active);
  editor?.exec({ type: 'setSectionsCollapsed', collapsed: !collapsed, sectionIds: [section.id] });
  if (inside) node?.focus({ preventScroll: true });
}

/**
 * A click on a section's name selects its slides. Focus goes to the current slide of the
 * selection, so the slide keys act on the slides that the click selected. A shut section keeps
 * focus. @internal
 */
export function selectSectionByName(
  editor: PptxEditorInstance | null,
  section: SlideSection,
  collapsed: boolean,
  node: HTMLElement | null
): void {
  selectSlideSection(editor, section);
  const current = editor?.snapshot().currentSlide;
  const slide = collapsed
    ? null
    : node?.querySelector<HTMLElement>(`[role="group"] [data-index="${current}"]`);
  (slide ?? node)?.focus({ preventScroll: true });
}

// --- Slide list keys ----------------------------------------------------------------------------

/** What the keys of a slide list (the rail, the sorter) need. @internal */
export interface SlideListKeyContext {
  readonly editor: PptxEditorInstance | null;
  readonly readOnly: boolean;
  readonly selected: readonly number[];
  readonly current: number;
  readonly count: number;
  readonly grouped: readonly SlideSection[] | null;
  /** The list element (`ol`). */
  readonly list: HTMLElement | null;
  /** Where the rail runs; `vertical` by default. A horizontal rail steps with Left and Right. */
  readonly orientation?: SlideRailOrientation;
  /** Where the selected slides land one step up or down (`slideStep`), or null at an edge. */
  step(step: -1 | 1): SlideStep | null;
  /** Open the slide menu on slide `index` at its item, from the keyboard. */
  openMenu(index: number, item: HTMLElement): void;
  /** Move focus to the current slide after the next render. */
  focusCurrent(): void;
}

const handled = (e: PaneKeyEvent): void => {
  e.preventDefault();
  e.stopPropagation();
};

/**
 * Whether a key event comes from outside the list `list`: a menu a framework portals out of the
 * list (a React portal's events bubble through the component tree) keeps its own keys, so Enter
 * on a menu row runs the row, not New Slide. @internal
 */
export function fromOutsideList(e: PaneKeyEvent, list: HTMLElement | null): boolean {
  const target = e.target as Node | null;
  return !!list && !!target && !list.contains(target);
}

/**
 * Whether a key of a slide list edits the deck: delete, duplicate, undo and redo, New Slide
 * (Enter in the rail, Ctrl+M), and Cmd/Ctrl+Arrow to move slides. In the viewing mode those
 * keys show why they do nothing, as the buttons do.
 */
function slideListEditKey(e: PaneKeyEvent, enterAdds: boolean): boolean {
  const mod = e.metaKey || e.ctrlKey;
  const lower = e.key.toLowerCase();
  return (
    e.key === 'Backspace' ||
    e.key === 'Delete' ||
    (mod && (lower === 'd' || lower === 'z' || lower === 'y' || e.key.startsWith('Arrow'))) ||
    (enterAdds && e.key === 'Enter' && !mod) ||
    (e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey && lower === 'm')
  );
}

/**
 * Up and Down through a rail with sections, as in a tree: the headers are in the order, so Up
 * from a section's first slide goes to its header. A slide becomes the current slide; a header
 * only takes focus. Left on a slide goes to its header. Returns whether the key moved.
 */
function railTreeStep(e: PaneKeyEvent, ctx: SlideListKeyContext): boolean {
  const { editor, grouped, list } = ctx;
  if (!grouped || !editor || e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return false;
  const rows = [
    ...(list?.querySelectorAll<HTMLElement>('.pptx-rail__section, .pptx-rail__item') ?? []),
  ];
  const from = (e.target as HTMLElement).closest<HTMLElement>(
    '.pptx-rail__item, .pptx-rail__section'
  );
  const at = from ? rows.indexOf(from) : -1;
  const isItem = !!from?.classList.contains('pptx-rail__item');
  let to: HTMLElement | undefined;
  const orientation = ctx.orientation ?? railOrientationOf(list);
  const step = railStepOfKey(e.key, orientation, list);
  if (step === 1) to = rows[at + 1];
  else if (step === -1) to = rows[at - 1];
  else if (e.key === 'Home') to = rows[0];
  else if (e.key === 'End') to = rows[rows.length - 1];
  else if (e.key === railParentKey(orientation) && isItem)
    to = from!.parentElement?.closest<HTMLElement>('.pptx-rail__section') ?? undefined;
  else return false;
  if (!to) return true;
  if (to.classList.contains('pptx-rail__item')) editor.goToSlide(Number(to.dataset.index));
  to.focus();
  return true;
}

/** The slide shortcuts of the rail while it has focus. @internal */
export function railKeyDown(e: PaneKeyEvent, ctx: SlideListKeyContext): void {
  const target = e.target as HTMLElement;
  // A section's rename field keeps its own keys, and so does a menu portaled out of the rail.
  if (fromOutsideList(e, ctx.list) || target.closest('.pptx-rail__section-name-field')) return;
  const { editor, readOnly, selected, current, count } = ctx;
  if (editor && (e.key === 'ContextMenu' || (e.shiftKey && e.key === 'F10'))) {
    const item = target.closest<HTMLElement>('.pptx-rail__item');
    if (item) ctx.openMenu(Number(item.dataset.index), item);
    handled(e);
    return;
  }
  if (railTreeStep(e, ctx)) {
    handled(e);
    return;
  }
  if (!editor) return;
  const mod = e.metaKey || e.ctrlKey;
  const step = railStepOfKey(e.key, ctx.orientation ?? railOrientationOf(ctx.list), ctx.list);
  // A focused section header is not a slide: keys that act on the selected slides stop there.
  const onHeader = target.classList.contains('pptx-rail__section');
  const lower = e.key.toLowerCase();
  if (readOnly && !onHeader && slideListEditKey(e, true)) {
    editor.reject('disabled.viewing');
    handled(e);
    return;
  }
  if (
    onHeader &&
    (e.key === 'Backspace' ||
      e.key === 'Delete' ||
      (mod && (lower === 'd' || lower === 'a' || step !== 0)) ||
      (e.shiftKey && step !== 0))
  ) {
    handled(e);
    return;
  }
  if (e.key === 'Backspace' || e.key === 'Delete')
    editor.exec({ type: 'deleteSlides', indices: selected });
  else if (mod && lower === 'd') editor.exec({ type: 'duplicateSlides', indices: selected });
  else if (mod && lower === 'z') editor.exec({ type: e.shiftKey ? 'redo' : 'undo' });
  else if (mod && lower === 'y') editor.exec({ type: 'redo' });
  // New Slide: Enter in the slide pane, or Ctrl+M.
  else if (
    (e.key === 'Enter' && !mod) ||
    (e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey && lower === 'm')
  )
    editor.exec({ type: 'addSlide' });
  else if (mod && lower === 'a') {
    editor.selectSlide(0);
    editor.selectSlide(count - 1, 'range');
  } else if (step && mod) {
    // A step across a section header moves the item into another group: focus follows it.
    runSlideStep(editor, stepCommand(selected, ctx.step(step)));
    ctx.focusCurrent();
  } else if (step && e.shiftKey) editor.selectSlide(current + step, 'range');
  else if (step) editor.goToSlide(current + step);
  // Home and End go to the first and the last slide (a rail with sections: `railTreeStep`).
  else if ((e.key === 'Home' || e.key === 'End') && !mod && !e.altKey) {
    const to = e.key === 'Home' ? 0 : count - 1;
    if (e.shiftKey) editor.selectSlide(to, 'range');
    else editor.goToSlide(to);
  } else return;
  handled(e);
}

const SORTER_HEADER = '.pptx-sorter__section[role="treeitem"]';

/**
 * Where an arrow key goes in the sorter. On a tile: Left and Right go to the tile before or
 * after; Up and Down go to the nearest tile in the row above or below in the same section, and
 * past the section's first or last row to its header or the next row. On a section header: Up
 * and Down go to the row before or after it. Home and End go to the first and last row.
 * @internal
 */
export function sorterKeyTarget(
  list: HTMLElement | null,
  from: HTMLElement | null,
  key: string
): HTMLElement | null {
  const rows = [
    ...(list?.querySelectorAll<HTMLElement>(`${SORTER_HEADER}, .pptx-sorter__item`) ?? []),
  ];
  if (!rows.length) return null;
  if (key === 'Home') return rows[0]!;
  if (key === 'End') return rows[rows.length - 1]!;
  const at = from ? rows.indexOf(from) : -1;
  if (at < 0) return rows[0]!;
  const up = key === 'ArrowUp';
  if (!from!.classList.contains('pptx-sorter__item'))
    return key === 'ArrowUp' || key === 'ArrowDown' ? (rows[at + (up ? -1 : 1)] ?? null) : null;
  const tiles = rows.filter((r) => r.classList.contains('pptx-sorter__item'));
  const t = tiles.indexOf(from!);
  if (key === 'ArrowLeft') return tiles[t - 1] ?? null;
  if (key === 'ArrowRight') return tiles[t + 1] ?? null;
  const grid = [...(from!.parentElement?.children ?? [])] as HTMLElement[];
  const box = from!.getBoundingClientRect();
  const x = box.left + box.width / 2;
  let best: HTMLElement | null = null;
  let bestRow = up ? -Infinity : Infinity;
  let bestDx = Infinity;
  for (const tile of grid) {
    const r = tile.getBoundingClientRect();
    if (up ? r.top >= box.top - 1 : r.top <= box.top + 1) continue;
    const dx = Math.abs(r.left + r.width / 2 - x);
    const nearer = up ? r.top > bestRow + 1 : r.top < bestRow - 1;
    if (nearer || (Math.abs(r.top - bestRow) <= 1 && dx < bestDx)) {
      best = tile;
      bestRow = r.top;
      bestDx = dx;
    }
  }
  if (best) return best;
  // Past the section's first row: its header; past its last row: the row after the section.
  if (up) return from!.closest<HTMLElement>(SORTER_HEADER);
  return rows[rows.indexOf(grid[grid.length - 1]!) + 1] ?? null;
}

/** The keys of the slide sorter. `openSlide` opens a slide in the Normal view. @internal */
export function sorterKeyDown(
  e: PaneKeyEvent,
  ctx: SlideListKeyContext & { openSlide(index: number): void }
): void {
  const target = e.target as HTMLElement;
  // A section's rename field keeps its own keys, and so does a menu portaled out of the sorter.
  if (fromOutsideList(e, ctx.list) || target.closest('.pptx-rail__section-name-field')) return;
  const { editor, readOnly, selected, current, count } = ctx;
  const item = target.closest<HTMLElement>('.pptx-sorter__item');
  const mod = e.metaKey || e.ctrlKey;
  const lower = e.key.toLowerCase();
  if (editor && item && (e.key === 'ContextMenu' || (e.shiftKey && e.key === 'F10'))) {
    ctx.openMenu(Number(item.dataset.index), item);
  } else if (!editor) return;
  else if (mod && item && !readOnly && e.key.startsWith('Arrow')) {
    // Cmd/Ctrl+Arrow moves the selected slides one step in grid order, as the rail does.
    const step = e.key === 'ArrowLeft' || e.key === 'ArrowUp' ? -1 : 1;
    runSlideStep(editor, stepCommand(selected, ctx.step(step)));
    // A step across a section header moves the tile into another group: focus follows it.
    ctx.focusCurrent();
  } else if (['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'].includes(e.key)) {
    const to = sorterKeyTarget(ctx.list, item ?? target, e.key);
    if (to?.classList.contains('pptx-sorter__item'))
      editor.selectSlide(Number(to.dataset.index), e.shiftKey ? 'range' : 'replace');
    to?.focus();
  } else if (mod && lower === 'a') {
    editor.selectSlide(0);
    editor.selectSlide(count - 1, 'range');
  } else if (!item) return;
  else if (e.key === 'Enter') ctx.openSlide(current);
  else if (readOnly && slideListEditKey(e, false)) editor.reject('disabled.viewing');
  else if (readOnly) return;
  else if (e.key === 'Backspace' || e.key === 'Delete')
    editor.exec({ type: 'deleteSlides', indices: selected });
  else if (mod && lower === 'd') editor.exec({ type: 'duplicateSlides', indices: selected });
  else if (mod && lower === 'z') editor.exec({ type: e.shiftKey ? 'redo' : 'undo' });
  else if (mod && lower === 'y') editor.exec({ type: 'redo' });
  // New Slide (Ctrl+M).
  else if (e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey && lower === 'm')
    editor.exec({ type: 'addSlide' });
  else return;
  handled(e);
}

/**
 * Open slide `index` in the Normal view, as a double-click or Enter in the sorter does. The
 * sorter goes away with its focus, so focus goes to the current slide in the rail, or to the
 * slide when the rail is absent or hidden, once the Normal view rendered. `from` is the sorter's list.
 * @internal
 */
export function openSlideInNormalView(
  editor: PptxEditorInstance | null,
  index: number,
  from?: HTMLElement | null
): void {
  const scope = from?.closest<HTMLElement>('.pptx-editor') ?? null;
  editor?.goToSlide(index);
  editor?.exec({ type: 'setPresentationView', view: 'normal' });
  focusNormalViewAfterRender(scope, editor);
}

/**
 * The slide sorter `list` is about to leave the page (View > Normal, the Normal view button).
 * When one of its slides has focus, focus goes to the current slide in the rail, or to the
 * slide, once the Normal view rendered, so it never stays on the page body. Focus that a host
 * put elsewhere stays. Call it before the sorter leaves the page. @internal
 */
export function refocusAfterSorterHides(
  list: HTMLElement,
  editor: PptxEditorInstance | null
): void {
  if (!list.contains(list.ownerDocument.activeElement)) return;
  focusNormalViewAfterRender(list.closest<HTMLElement>('.pptx-editor'), editor);
}

/** After the adapter rendered the Normal view (the rail mounts with it), focus its current slide when focus is lost. */
function focusNormalViewAfterRender(
  scope: HTMLElement | null,
  editor: PptxEditorInstance | null
): void {
  const view = scope?.ownerDocument.defaultView;
  if (!editor || !scope || !view) return;
  const go = (): void => {
    const active = scope.ownerDocument.activeElement;
    if (active && active !== scope.ownerDocument.body && active.isConnected) return;
    if (scope.isConnected) focusCurrentSlideOrCanvas(scope, editor);
  };
  if (typeof view.requestAnimationFrame === 'function') view.requestAnimationFrame(go);
  else setTimeout(go, 0);
}

/**
 * Where a section menu opened from the keyboard opens: under the header row of the section
 * `owner`, not under its slides. @internal
 */
export function sectionMenuPoint(owner: HTMLElement): { x: number; y: number } {
  const head = owner.querySelector<HTMLElement>('.pptx-rail__section-head') ?? owner;
  const r = head.getBoundingClientRect();
  return { x: r.left + 12, y: r.bottom - 4 };
}

/** Where a slide menu opened from the keyboard opens: at the slide item. @internal */
export function slideMenuPoint(
  item: HTMLElement,
  at: { x: number; y: number } | null
): { x: number; y: number } {
  if (at) return at;
  const r = item.getBoundingClientRect();
  return { x: r.left + 12, y: r.bottom - 8 };
}

// --- Speaker notes ------------------------------------------------------------------------------

/** A word character: a letter, a digit, an apostrophe, or a hyphen inside the word. */
const WORD = /[\p{L}\p{N}\p{M}_'’-]/u;

/**
 * The word that has offset `at` of `text` inside it, or null. A caret at the start or the end
 * of a word is not inside it: there the format applies to the next typed text
 * (`armNotesFormat`). @internal
 */
export function notesWordAt(text: string, at: number): { start: number; end: number } | null {
  if (!WORD.test(text[at - 1] ?? '') || !WORD.test(text[at] ?? '')) return null;
  let start = at;
  let end = at;
  while (start > 0 && WORD.test(text[start - 1]!)) start -= 1;
  while (end < text.length && WORD.test(text[end]!)) end += 1;
  return start < end ? { start, end } : null;
}

/** No notes format: what the B, I, and U buttons show without a focused field. @internal */
export const NO_NOTES_FORMATS: Readonly<Record<NotesFormat, boolean>> = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
};

/** The notes pane's format buttons: format, icon, and label key. @internal */
export const NOTES_FORMAT_BUTTONS: readonly (readonly [
  NotesFormat,
  string,
  'notes.bold' | 'notes.italic' | 'notes.underline',
])[] = [
  ['bold', 'format_bold', 'notes.bold'],
  ['italic', 'format_italic', 'notes.italic'],
  ['underline', 'format_underlined', 'notes.underline'],
];

type ModifierKeys = Pick<PaneKeyEvent, 'key' | 'metaKey' | 'ctrlKey' | 'shiftKey' | 'altKey'>;

/** The keyboard shortcut of a notes format (Ctrl/Cmd + B, I, U), or null. @internal */
export function notesFormatShortcut(e: ModifierKeys): NotesFormat | null {
  if (!(e.metaKey || e.ctrlKey) || e.altKey || e.shiftKey) return null;
  const key = e.key.toLowerCase();
  return key === 'b' ? 'bold' : key === 'i' ? 'italic' : key === 'u' ? 'underline' : null;
}

/** Undo or redo for a key (Ctrl/Cmd+Z, Ctrl/Cmd+Shift+Z, Ctrl+Y), or null. @internal */
export function notesHistoryShortcut(e: ModifierKeys): 'undo' | 'redo' | null {
  if (!(e.metaKey || e.ctrlKey) || e.altKey) return null;
  const key = e.key.toLowerCase();
  if (key === 'z') return e.shiftKey ? 'redo' : 'undo';
  return key === 'y' && !e.shiftKey ? 'redo' : null;
}

/** Whether the notes field `el`, or an element inside it, has the keyboard focus. @internal */
export function notesFieldFocused(el: HTMLElement): boolean {
  const active = el.ownerDocument.activeElement;
  return !!active && (active === el || el.contains(active));
}

/** What the notes field's keys need. @internal */
export interface NotesKeyContext {
  readonly readOnly: boolean;
  onDraft(text: string): void;
  onFormat(format: NotesFormat): void;
  onHistory(kind: 'undo' | 'redo'): void;
  /** New Slide (Ctrl+M), as on the canvas. */
  readonly onNewSlide?: (() => void) | undefined;
}

/** The keys of the notes field `el`. A composition in progress keeps its keys. @internal */
export function notesFieldKeyDown(e: PaneKeyEvent, el: HTMLElement, ctx: NotesKeyContext): void {
  const format = notesFormatShortcut(e);
  const history = notesHistoryShortcut(e);
  if (format) {
    e.preventDefault();
    if (!ctx.readOnly) ctx.onFormat(format);
  } else if (history) {
    // The deck's history, not the browser's: a format is a deck edit too.
    e.preventDefault();
    if (!ctx.readOnly) ctx.onHistory(history);
  } else if (
    ctx.onNewSlide &&
    e.ctrlKey &&
    !e.metaKey &&
    !e.altKey &&
    !e.shiftKey &&
    e.key.toLowerCase() === 'm'
  ) {
    e.preventDefault();
    if (!ctx.readOnly) ctx.onNewSlide();
  } else if (e.key === 'Enter' && !ctx.readOnly) {
    // Every browser writes a new paragraph the same way: one `\n`.
    e.preventDefault();
    insertNotesText(el, '\n');
    ctx.onDraft(notesFieldText(el));
  } else if (e.key === 'Escape') {
    // Out of the notes, to the current slide of the rail, so the keyboard stays in the editor.
    // A current slide in a collapsed section has no tile: the rail's tab stop (the section
    // header) takes focus, as after a re-render. Without a rail the field only loses focus.
    const editor = el.closest('.pptx-editor');
    const slide =
      editor?.querySelector<HTMLElement>('.pptx-rail__item[aria-current="true"]') ??
      editor?.querySelector<HTMLElement>('.pptx-rail [role="treeitem"][tabindex="0"]');
    if (slide) slide.focus();
    else el.blur();
  }
}

/**
 * The notes field's `beforeinput`: typed text at a caret with an armed format (`armNotesFormat`)
 * goes into a run with that format, and typed text at the start of a paragraph goes into its
 * first run (`typeAtParagraphStart`); the draft follows. Other input as usual.
 * @internal
 */
export function notesFieldBeforeInput(
  e: InputEvent,
  el: HTMLElement,
  ctx: Pick<NotesKeyContext, 'readOnly' | 'onDraft'>
): void {
  if (ctx.readOnly || e.isComposing || e.inputType !== 'insertText' || !e.data) return;
  if (!typeArmedNotes(el, e.data) && !typeAtParagraphStart(el, e.data)) return;
  e.preventDefault();
  ctx.onDraft(notesFieldText(el));
}

/** A selection range in the notes text, in the offsets of the written text. @internal */
export interface NotesRange {
  readonly start: number;
  readonly end: number;
}

/** A field selection `selected` of field text `text`, in the offsets of the text `setNotes` writes. @internal */
export function notesCommittedRange(text: string, selected: NotesRange | null): NotesRange | null {
  return (
    selected && {
      start: notesCommittedOffset(text, selected.start),
      end: notesCommittedOffset(text, selected.end),
    }
  );
}

/**
 * The range a notes format acts on: the selection, or, for a caret in a word, the word. Null
 * when there is nothing to format. @internal
 */
export function notesFormatRange(text: string, selected: NotesRange | null): NotesRange | null {
  return selected && selected.start === selected.end ? notesWordAt(text, selected.start) : selected;
}

/**
 * Whether the B, I, and U buttons take a pick: selected text formats it, a caret in a word
 * formats the word, and any other caret arms the format for the next typed text. @internal
 */
export function notesCanFormat(text: string, range: NotesRange | null): boolean {
  return !!range && range.start <= text.length;
}

/** Whether two format readings are the same, so a pane need not render again. @internal */
export function sameNotesFormats(
  a: Readonly<Record<NotesFormat, boolean>>,
  b: Readonly<Record<NotesFormat, boolean>>
): boolean {
  return (Object.keys(b) as NotesFormat[]).every((f) => a[f] === b[f]);
}
