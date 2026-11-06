/** @jsxImportSource react */
// The slide rail, speaker notes, zoom pill, and side panel. Each is a hook consumer over the
// editor; none holds document state of its own.

import type { ChromeTranslate } from '@pptx-editor.dev/core/editor';
import {
  memo,
  useEffect,
  useId,
  useLayoutEffect,
  useMemo,
  useReducer,
  useRef,
  useState,
  type CSSProperties,
  type DragEvent,
  type FocusEvent,
  type KeyboardEvent,
  type MouseEvent,
  type ReactNode,
  type RefObject,
} from 'react';
import {
  ZOOM_MENU_LEVELS,
  createProgressiveMount,
  tooltipAttributes,
  followCurrentSlide,
  groupSlideSections,
  keepCurrentSlideInView,
  radioGroupStep,
  railDragAxis,
  railKeyDown,
  railOrientationChrome,
  railRightToLeft,
  watchRailFold,
  resolveRailOrientation,
  runSlideStep,
  sectionOfSlide,
  slideClickMode,
  slideItemName,
  slideMenuPoint,
  stepCommand,
  zoomStep,
  zoomLimitReason,
  elementKindLabelKey,
  platformShortcuts,
  type SlideRailOrientation,
  type SlideSummary,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { Icon, MenuItem, Popover, ToolButton, keepFocus, useAriaDisabled } from './chrome.tsx';
import { useEditorDialog } from './popups.tsx';
import { RailMenuHost, useSlideStep, type RailMenuState } from './RailContextMenu.tsx';
import {
  RailActionsContext,
  railActionContent,
  railExpandContent,
  type RailActions,
} from './rail-parts.tsx';
import { useSlideDrag } from './slide-drag.ts';
import { LayoutGallery } from './SlideControls.tsx';
import {
  SectionNode,
  useCollapsedSections,
  useFocusAfterCollapse,
  useFocusCurrentAfterRender,
  useOpenCurrentSection,
  useSectionRenameRequests,
  useSlideSections,
} from './RailSections.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';

/** The side panel's tabs, in order. */
const PANEL_TABS = ['format', 'comments'] as const;

const sameSlides = (a: readonly SlideSummary[], b: readonly SlideSummary[]): boolean =>
  a.length === b.length && a.every((s, i) => s === b[i]);
const sameIndices = (a: readonly number[], b: readonly number[]): boolean =>
  a.length === b.length && a.every((v, i) => v === b[i]);

/** How a click selects slides: Shift+click a range, Ctrl/Cmd+click one more or one less. */
export const clickMode = slideClickMode;

/**
 * A slide thumbnail, repainted only when its slide renders differently. It is hidden from
 * assistive technology: the slide item that holds it gives the name ("Slide 3"). @internal
 */
export const Thumbnail = memo(function Thumbnail({
  slide,
  ratio,
}: {
  slide: SlideSummary;
  ratio: number;
}) {
  const editor = usePptxEditor();
  const box = useRef<HTMLDivElement>(null);
  // A deferred paint: a long list paints the thumbnails in view first. The disposer runs on
  // unmount and on an editor change, so the old editor forgets the container.
  useEffect(() => {
    if (!editor || !box.current) return undefined;
    return editor.renderThumbnail(slide.index, box.current, { defer: true });
  }, [editor, slide]);
  return (
    <div
      ref={box}
      className="pptx-rail__thumb"
      style={{ aspectRatio: String(1 / ratio) }}
      aria-hidden="true"
    />
  );
});

/** The rail's handlers for its items: one object for the rail's life, which reads its latest render. */
interface RailItemHandlers {
  click(index: number, e: MouseEvent<HTMLLIElement>): void;
  menu(index: number, e: MouseEvent<HTMLLIElement>): void;
  dragStart(index: number, e: DragEvent<HTMLElement>): void;
  dragOver(index: number, e: DragEvent<HTMLElement>): void;
  dragEnd(index: number): void;
  drop(index: number, e: DragEvent<HTMLElement>): void;
}

/**
 * One slide of the rail. Memoized on its props, so a keystroke that changes one slide, or a
 * long deck whose rail grows while it opens, renders only the items that changed. @internal
 */
const RailItem = memo(function RailItem(props: {
  slide: SlideSummary;
  ratio: number;
  selected: boolean;
  current: boolean;
  role: 'treeitem' | 'option';
  /** The item's name for assistive technology (`slideItemName`). */
  name: string;
  drop: string | undefined;
  draggable: boolean;
  on: RailItemHandlers;
}) {
  const { slide: s, on } = props;
  return (
    <li
      className="pptx-rail__item"
      role={props.role}
      aria-selected={props.selected}
      aria-current={props.current}
      data-hidden={s.hidden ? '' : undefined}
      data-index={s.index}
      data-drop={props.drop}
      draggable={props.draggable}
      onDragStart={(e) => on.dragStart(s.index, e)}
      onDragOver={(e) => on.dragOver(s.index, e)}
      onDragEnd={() => on.dragEnd(s.index)}
      onDrop={(e) => on.drop(s.index, e)}
      tabIndex={props.current ? 0 : -1}
      onClick={(e) => on.click(s.index, e)}
      onContextMenu={(e) => on.menu(s.index, e)}
    >
      {/* The name comes from the hidden text, not `aria-label`: the visible number is part of
          it, and the thumbnail art is hidden from assistive technology. */}
      <span className="pptx-rail__number" aria-hidden="true">
        {s.index + 1}
      </span>
      <span className="pptx-sr-only">{props.name}</span>
      <Thumbnail slide={s} ratio={props.ratio} />
    </li>
  );
});

/**
 * Keyboard focus follows the current slide of a slide list (the rail, the slide sorter) while
 * one of its items `item` has focus. When the focused item goes away (Delete on the last
 * slide, an undo that removes one), focus goes to the current item too, so the next key still
 * acts on the list. Returns the list's `onFocus` handler.
 */
export function useFocusFollowsCurrent(
  list: RefObject<HTMLElement | null>,
  item: string,
  current: number,
  count: number,
  shown = true
): (e: FocusEvent) => void {
  const focused = useRef<HTMLElement | null>(null);
  useEffect(() => {
    followCurrentSlide(list.current, item, focused.current);
  }, [list, item, current, count, shown]);
  return (e) => {
    focused.current = (e.target as HTMLElement).closest?.<HTMLElement>(item) ?? null;
  };
}

/**
 * Props of the editor's panes, after DOCX Editor's `Navigation` (`open`, `onOpenChange`,
 * `paneWidth`). A pane given `open` is controlled by the host. @public
 */
export interface PaneProps {
  /** A class name added to the pane. */
  readonly className?: string;
  /** Whether the pane is open, controlled. Use with `onOpenChange`. */
  readonly open?: boolean;
  /** Whether the pane is open first, uncontrolled. Default `true`. */
  readonly defaultOpen?: boolean;
  /** Called when the user opens or closes the pane. */
  readonly onOpenChange?: (open: boolean) => void;
  /** The pane's width in CSS pixels (the slide rail). */
  readonly paneWidth?: number;
  /** Inline styles for the pane, after DOCX Editor's `Navigation`; a `width` here wins over `paneWidth`. */
  readonly style?: CSSProperties;
}

/** Props of `SlideRail`: the pane props, DOCX Editor's `toggle`, and `orientation`. @public */
export interface SlideRailProps extends PaneProps {
  /**
   * Your translation function for this compound and its parts, as DOCX Editor's `t`. Without
   * it, the root's `locale` gives the labels.
   */
  readonly t?: ChromeTranslate;
  /**
   * Where the rail runs. `vertical` (the default) is a column beside the slide. `horizontal` is
   * a row under it, for a tall, narrow host: the thumbnails scroll sideways, Left and Right
   * step, a drag reorders along the row, and Hide thumbnails folds the rail down.
   */
  readonly orientation?: SlideRailOrientation;
  /**
   * The built-in Hide thumbnails and Show thumbnails buttons. `false` removes them, for a host
   * that opens and closes the rail from its own chrome (`open`, `onOpenChange`).
   */
  readonly toggle?: boolean;
  /**
   * `true` (default): the packaged action buttons, with named part children
   * (`SlideRail.AddSlide`, `SlideRail.Delete`, …) in their places and other children after
   * Delete. `false`: only the children, in their order.
   */
  readonly preset?: boolean;
  /** Named part children replace their buttons; other children go after **Delete**. */
  readonly children?: ReactNode;
}

/** The grid columns of a horizontal rail's two action rows (`--pptx-rail-action-columns`). */
function actionColumns(shown: number): CSSProperties {
  return {
    '--pptx-rail-action-columns': String(Math.max(1, Math.ceil(shown / 2))),
  } as CSSProperties;
}

/** The collapsed rail strip keeps the host's styles but not a width: it stays a narrow strip. */
function stripStyle(style: CSSProperties | undefined): CSSProperties | undefined {
  if (!style || style.width === undefined) return style;
  const { width: _width, ...rest } = style;
  return rest;
}

/** A pane's inline style: its width from `paneWidth`, then the host's own styles. */
function paneStyle(width: number | undefined, style: CSSProperties | undefined) {
  if (width === undefined) return style;
  return { width, ...style };
}

/**
 * How many rail items to mount while a long deck opens: the first screen with the first slide,
 * then more after each painted frame (core's `createProgressiveMount`).
 */
function useProgressiveMount(total: number, needed: number, opening: boolean): number {
  const [, grow] = useReducer((n: number) => n + 1, 0);
  const [mount] = useState(() => createProgressiveMount(grow));
  useEffect(() => () => mount.dispose(), [mount]);
  const shown = mount.limit(total, needed, opening);
  useEffect(() => mount.follow(total));
  return shown;
}

/** Open state of a pane: controlled by `open`, else its own. */
function usePaneOpen(props: PaneProps) {
  const [own, setOwn] = useState(props.defaultOpen ?? true);
  const open = props.open ?? own;
  const setOpen = (next: boolean): void => {
    if (props.open === undefined) setOwn(next);
    props.onOpenChange?.(next);
  };
  return [open, setOpen] as const;
}

/**
 * The slide thumbnail rail with drag-to-reorder. Its last button hides it to a narrow strip
 * with a button that shows it again. @public
 */
export function SlideRail(props: SlideRailProps = {}) {
  const [open, setPaneOpen] = usePaneOpen(props);
  const toggle = props.toggle !== false;
  const rail = useRef<HTMLElement>(null);
  const listId = useId();
  // A toggle unmounts its own button: focus moves to the other toggle, never to the page.
  const focusToggle = useRef(false);
  const setOpen = (next: boolean): void => {
    focusToggle.current = rail.current?.contains(document.activeElement) ?? false;
    setPaneOpen(next);
  };
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const slides = useEditorState((s) => s.slides, sameSlides);
  const current = useEditorState((s) => s.currentSlide);
  const selected = useEditorState((s) => s.selectedSlides, sameIndices);
  const ratio = useEditorState((s) => s.slideHeight / s.slideWidth);
  const readOnly = useEditorState((s) => s.mode === 'view');
  const list = useRef<HTMLOListElement>(null);
  const count = slides.length;
  const opening = useEditorState((s) => (s.pendingSlides ?? 0) > 0);
  const mounted = useProgressiveMount(count, current, opening);
  const [menu, setMenu] = useState<RailMenuState | null>(null);
  const [gallery, setGallery] = useState<HTMLElement | null>(null);
  const actions = useRef<HTMLDivElement>(null);
  const sections = useSlideSections();
  // Sections group the rail when they cover every slide in order; otherwise it stays flat.
  const grouped = groupSlideSections(sections, count);
  const orientation = resolveRailOrientation(props.orientation, 'PptxEditor.SlideRail');
  // A right-to-left row mirrors the move buttons, as its arrow keys do. The direction comes
  // from the mounted rail, so it is read after the rail mounts or changes.
  const [rightToLeft, setRightToLeft] = useState(false);
  useLayoutEffect(() => setRightToLeft(railRightToLeft(rail.current)), [orientation, open]);
  const chrome = railOrientationChrome(orientation, rightToLeft);
  // A narrow horizontal rail (a phone) folds its secondary actions into one menu.
  const [folded, setFolded] = useState(false);
  useEffect(() => watchRailFold(rail.current, orientation, setFolded), [orientation, open]);
  const {
    itemProps: dragProps,
    cancel: cancelDrag,
    marker: dropMarker,
  } = useSlideDrag(railDragAxis(orientation), selected, readOnly, grouped);
  const collapsed = useCollapsedSections();
  const [renaming, setRenaming] = useState<string | null>(null);
  useOpenCurrentSection(grouped);
  // An open rail renames in place, also while it is flat: Add Section in a deck without sections
  // asks for the name before the new header renders. The id waits in `renaming` until the
  // header mounts with its field. A closed rail has no header, so it does not count.
  useSectionRenameRequests(setRenaming, open);
  const focusCurrent = useFocusCurrentAfterRender(list);
  useFocusAfterCollapse(list, collapsed);
  const background = useEditorDialog('background');
  const stepTarget = useSlideStep(selected, count);
  /**
   * Open the context menu on slide `index`. A slide outside the selection becomes the
   * selection; a slide inside a multi-selection keeps it.
   */
  const openMenu = (index: number, item: HTMLElement, at: { x: number; y: number } | null) => {
    if (!editor) return;
    if (!editor.snapshot().selectedSlides.includes(index)) editor.selectSlide(index);
    setMenu({ at: slideMenuPoint(item, at), index, keyboard: at === null, owner: item });
  };
  const moveBy = (step: -1 | 1): void => {
    runSlideStep(editor, stepCommand(selected, stepTarget(step)));
  };
  const trackFocus = useFocusFollowsCurrent(list, '.pptx-rail__item', current, count, open);
  // A resize (a host narrowing the panel) keeps the current slide in view too.
  useEffect(() => keepCurrentSlideInView(list.current, '.pptx-rail__item'), [open]);
  /** Why the rail's slide buttons are off, or null when they work. */
  const why = !editor ? t('disabled.noEditor') : readOnly ? t('disabled.viewing') : null;
  useEffect(() => {
    // Only when the toggle's own unmount lost focus: a host that vetoed the change and closes
    // the rail later from its own chrome keeps its focus.
    const active = document.activeElement;
    const lost = !active || active === document.body || rail.current?.contains(active);
    if (focusToggle.current && lost)
      rail.current?.querySelector<HTMLElement>('.pptx-rail__collapse')?.focus();
    // A closed rail drops its open menu, drag, and rename.
    if (!open) {
      setMenu(null);
      cancelDrag();
      setRenaming(null);
    }
  }, [open, cancelDrag]);
  // A toggle the host refused leaves no flag behind for a later change to act on.
  useEffect(() => {
    focusToggle.current = false;
  });
  /** Slide shortcuts while the rail has focus (core's `railKeyDown`). */
  const onKeyDown = (e: KeyboardEvent): void =>
    railKeyDown(e, {
      editor,
      readOnly,
      selected,
      current,
      count,
      grouped,
      list: list.current,
      orientation,
      step: stepTarget,
      openMenu: (index, item) => openMenu(index, item, null),
      focusCurrent,
    });
  // The items' handlers read the latest render, so the object stays the same and a memoized
  // item does not render again for a new closure.
  const latest = useRef({ editor, dragProps, openMenu });
  useLayoutEffect(() => {
    latest.current = { editor, dragProps, openMenu };
  });
  const on = useMemo<RailItemHandlers>(
    () => ({
      click: (index, e) => {
        e.currentTarget.focus({ preventScroll: true });
        latest.current.editor?.selectSlide(index, clickMode(e));
      },
      menu: (index, e) => {
        e.preventDefault();
        // A keyboard-opened menu reports no pointer position: open at the slide.
        const pointer = e.clientX !== 0 || e.clientY !== 0;
        const at = pointer ? { x: e.clientX, y: e.clientY } : null;
        latest.current.openMenu(index, e.currentTarget, at);
      },
      dragStart: (index, e) => latest.current.dragProps(index).onDragStart(e),
      dragOver: (index, e) => latest.current.dragProps(index).onDragOver(e),
      dragEnd: (index) => latest.current.dragProps(index).onDragEnd(),
      drop: (index, e) => latest.current.dragProps(index).onDrop(e),
    }),
    []
  );
  /** One slide of the rail. */
  const item = (s: SlideSummary) => (
    <RailItem
      key={s.id}
      slide={s}
      ratio={ratio}
      selected={selected.includes(s.index)}
      current={s.index === current}
      role={grouped ? 'treeitem' : 'option'}
      name={slideItemName(t, s)}
      drop={dropMarker(s.index)}
      draggable={!readOnly}
      on={on}
    />
  );
  const railActions: RailActions = {
    t,
    editor,
    why,
    selected,
    count,
    chrome,
    step: stepTarget,
    moveBy,
    galleryOpen: !!gallery,
    // A second click closes the gallery, as the toolbar's menu buttons do.
    // In a folded rail, the gallery opens at the overflow button that held its row.
    toggleGallery: () =>
      setGallery((shown) =>
        shown
          ? null
          : (actions.current?.querySelector<HTMLElement>('.pptx-rail__layouts, .pptx-rail__more') ??
            null)
      ),
    listId,
    open,
    toggle,
    collapse: () => setOpen(false),
    expand: () => setOpen(true),
  };
  const { content: actionContent, shown } = railActionContent(
    props.children,
    props.preset !== false,
    toggle,
    folded
  );
  const className = ['pptx-rail', props.className].filter(Boolean).join(' ');
  // Closed with no toggle, the rail leaves the page: the host's chrome opens it again.
  if (!open && !toggle) return null;
  if (!open)
    return (
      <nav
        ref={rail}
        className={`${className} pptx-rail--collapsed`}
        aria-label={t('rail.ariaLabel')}
        data-open="false"
        data-orientation={orientation}
        style={stripStyle(props.style)}
        onMouseDown={keepFocus}
      >
        <RailActionsContext.Provider value={railActions}>
          {railExpandContent(props.children)}
        </RailActionsContext.Provider>
      </nav>
    );
  return (
    <nav
      ref={rail}
      className={className}
      aria-label={t('rail.ariaLabel')}
      data-open="true"
      data-orientation={orientation}
      style={paneStyle(props.paneWidth, props.style)}
    >
      <ol
        ref={list}
        id={listId}
        className="pptx-rail__list"
        role={grouped ? 'tree' : 'listbox'}
        aria-multiselectable="true"
        aria-orientation={orientation}
        aria-label={t('rail.ariaLabel')}
        onKeyDown={onKeyDown}
        onFocus={trackFocus}
      >
        {grouped
          ? grouped.map((sec) => {
              const shut = collapsed.includes(sec.id);
              return (
                <SectionNode
                  key={sec.id}
                  section={sec}
                  collapsed={shut}
                  renaming={renaming === sec.id}
                  variant="rail"
                  // A collapsed section that holds the current slide takes Tab in its place.
                  focusable={shut && sectionOfSlide(grouped, current)?.id === sec.id}
                  onRename={() => setRenaming(sec.id)}
                  onRenameEnd={() => setRenaming(null)}
                  onRemoved={focusCurrent}
                >
                  {slides
                    .slice(sec.firstSlide, Math.min(sec.firstSlide + sec.slideCount, mounted))
                    .map(item)}
                </SectionNode>
              );
            })
          : slides.slice(0, mounted).map(item)}
      </ol>
      <RailActionsContext.Provider value={railActions}>
        <div
          ref={actions}
          className="pptx-rail__actions"
          style={actionColumns(shown)}
          onMouseDown={keepFocus}
        >
          {actionContent}
        </div>
      </RailActionsContext.Provider>
      {gallery ? (
        <Popover anchor={gallery} onClose={() => setGallery(null)} label={t('rail.chooseLayout')}>
          <LayoutGallery
            onChange={(layoutId) => {
              setGallery(null);
              editor?.exec({ type: 'addSlide', layoutId });
            }}
          />
        </Popover>
      ) : null}
      <RailMenuHost
        open={
          menu
            ? {
                menu,
                onClose: () => setMenu(null),
                onSectionAdded: (id: string) => setRenaming(id),
                ...(background.isAvailable
                  ? { onBackground: () => background.open(menu.owner) }
                  : {}),
              }
            : null
        }
      />
    </nav>
  );
}

export { NotesPane, type NotesPaneProps } from './NotesPane.tsx';

/** Props of `ZoomPill`. @public */
export interface ZoomPillProps {
  /** A class name added to the pill. */
  readonly className?: string;
}

/** The zoom pill (`– 100% ▾ +`). @public */
export function ZoomPill({ className }: ZoomPillProps = {}) {
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const zoom = useEditorState((s) => s.zoom);
  const fit = useEditorState((s) => s.fitZoom);
  // The pill zooms the slide; a thumbnail zoom for the slide sorter is a later change (decisions).
  const off = useEditorState((s) => s.presentationView === 'slideSorter');
  const reason = off ? t('disabled.sorterZoom') : null;
  // The sorter draws no slide at the zoom: the level shows no number there, only why it is off.
  const levelLabel = off
    ? t('zoom.ariaLabel')
    : t('zoom.levelAriaLabel', { percent: Math.round(zoom * 100) });
  const anchor = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  // A view switch closes the zoom menu, so it does not open again on the way back.
  useEffect(() => {
    if (off) setOpen(false);
  }, [off]);
  const step = (dir: 1 | -1): void => {
    const next = zoomStep(zoom, dir);
    if (next) editor?.setZoom(next);
  };
  const levelOff = useAriaDisabled(off, reason);
  // A zoom button at the end of the range is off, and says why (zoom.ts).
  const limit = (dir: 1 | -1): string | null => {
    const key = zoomLimitReason(zoom, dir);
    return key ? t(key) : null;
  };
  return (
    <div
      className={['pptx-zoom', className].filter(Boolean).join(' ')}
      role="group"
      aria-label={t('zoom.ariaLabel')}
      onMouseDown={keepFocus}
    >
      <ToolButton
        label={t('zoom.zoomOut')}
        icon="remove"
        onClick={() => step(-1)}
        disabled={off || limit(-1) !== null}
        disabledReason={reason ?? limit(-1)}
      />
      <button
        ref={anchor}
        type="button"
        className="pptx-zoom__level"
        aria-label={levelLabel}
        {...tooltipAttributes(levelLabel, off, reason)}
        aria-haspopup="menu"
        aria-expanded={open && !off}
        // Off stays focusable, so its tooltip says why; the reason is its
        // description too.
        {...levelOff.attrs}
        data-disabled={off ? '' : undefined}
        onClick={() => !off && setOpen((o) => !o)}
      >
        {off ? '\u2013' : t('zoom.level', { percent: Math.round(zoom * 100) })}
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>
      {levelOff.note}
      <ToolButton
        label={t('zoom.zoomIn')}
        icon="add"
        onClick={() => step(1)}
        disabled={off || limit(1) !== null}
        disabledReason={reason ?? limit(1)}
      />
      {open && !off ? (
        <Popover anchor={anchor.current} onClose={() => setOpen(false)} align="end">
          <MenuItem checked={fit} onSelect={() => (editor?.setZoom('fit'), setOpen(false))}>
            {t('zoom.fit')}
          </MenuItem>
          {ZOOM_MENU_LEVELS.map((z) => (
            <MenuItem
              key={z}
              checked={!fit && Math.abs(z - zoom) < 0.001}
              onSelect={() => (editor?.setZoom(z), setOpen(false))}
            >
              {t('zoom.level', { percent: Math.round(z * 100) })}
            </MenuItem>
          ))}
        </Popover>
      ) : null}
    </div>
  );
}

const EMU_PER_INCH = 914_400;
const inches = (emu: number): string => (emu / EMU_PER_INCH).toFixed(2);

/** A number field that commits on Enter or blur and shows the document value otherwise. */
function NumberField({
  label,
  value,
  unit,
  disabled,
  onCommit,
}: {
  label: string;
  value: string;
  unit: string;
  disabled: boolean;
  onCommit: (value: number) => void;
}) {
  const [draft, setDraft] = useState<string | null>(null);
  const commit = (): void => {
    // An empty field keeps the value: `Number('')` is 0, which would move the shape.
    const n = draft === null || draft.trim() === '' ? NaN : Number(draft.replace(',', '.'));
    if (Number.isFinite(n) && draft !== value) onCommit(n);
    setDraft(null);
  };
  return (
    <label className="pptx-field">
      {label}
      <span className="pptx-field__control">
        <input
          inputMode="decimal"
          aria-label={label}
          value={draft ?? value}
          disabled={disabled}
          onChange={(e) => setDraft(e.target.value)}
          onBlur={commit}
          onKeyDown={(e) => {
            if (e.key === 'Enter') (e.currentTarget as HTMLInputElement).blur();
            if (e.key === 'Escape') setDraft(null);
          }}
        />
        <span className="pptx-field__unit">{unit}</span>
      </span>
    </label>
  );
}

/** Position, size, rotation, and text fitting of the selected element. */
function FormatTab() {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const geometry = useEditorState((s) => s.geometry);
  const table = useEditorState((s) => s.table);
  const readOnly = useEditorState((s) => s.mode === 'view');
  const selected = useEditorState((s) =>
    s.selection.kind === 'shapes' ? s.selection.nodeIds.length : 0
  );
  const background = useEditorDialog('background');
  if (!geometry)
    return selected > 1 ? (
      <div className="pptx-empty">
        <p className="pptx-empty__title">{t('panel.format.multipleTitle', { count: selected })}</p>
        <p className="pptx-empty__body">{platformShortcuts(t('panel.format.multipleBody'))}</p>
      </div>
    ) : (
      <>
        <div className="pptx-panel__section">
          <h3 className="pptx-panel__title">{t('panel.format.slide')}</h3>
          {background.isAvailable ? (
            <button
              type="button"
              className="pptx-button"
              disabled={readOnly}
              onClick={(e) => background.open(e.currentTarget)}
            >
              {t('menu.background')}
            </button>
          ) : null}
        </div>
        <div className="pptx-empty">
          <p className="pptx-empty__title">{t('panel.format.emptyTitle')}</p>
          <p className="pptx-empty__body">{t('panel.format.emptyBody')}</p>
        </div>
      </>
    );
  const g = geometry;
  const move = (patch: Partial<Record<'x' | 'y' | 'cx' | 'cy' | 'rot', number>>): void => {
    editor?.exec({
      type: 'setTransform',
      nodeId: g.nodeId,
      x: patch.x ?? g.x,
      y: patch.y ?? g.y,
      cx: Math.max(0, patch.cx ?? g.cx),
      cy: Math.max(0, patch.cy ?? g.cy),
      rot: patch.rot ?? g.rot,
    });
  };
  const inch = (v: number): number => Math.round(v * EMU_PER_INCH);
  const unit = t('panel.format.inchUnit');
  const field = (key: 'x' | 'y' | 'cx' | 'cy', labelKey: TranslationKey) => (
    <NumberField
      label={t(labelKey)}
      value={inches(g[key])}
      unit={unit}
      disabled={readOnly}
      onCommit={(v) => move({ [key]: inch(v) })}
    />
  );
  return (
    <>
      <div className="pptx-panel__section">
        <h3 className="pptx-panel__title">{t('panel.format.element')}</h3>
        <span className="pptx-chip">{t(elementKindLabelKey(g, table))}</span>
      </div>
      <div className="pptx-panel__section">
        <h3 className="pptx-panel__title">{t('panel.format.size')}</h3>
        <div className="pptx-panel__grid">
          {field('cx', 'panel.format.width')}
          {field('cy', 'panel.format.height')}
        </div>
      </div>
      <div className="pptx-panel__section">
        <h3 className="pptx-panel__title">{t('panel.format.position')}</h3>
        <div className="pptx-panel__grid">
          {field('x', 'panel.format.x')}
          {field('y', 'panel.format.y')}
        </div>
      </div>
      <div className="pptx-panel__section">
        <h3 className="pptx-panel__title">{t('panel.format.rotation')}</h3>
        <div className="pptx-panel__grid">
          <NumberField
            label={t('panel.format.angle')}
            value={String(Math.round(g.rot * 10) / 10)}
            unit="°"
            disabled={readOnly}
            onCommit={(v) => move({ rot: ((v % 360) + 360) % 360 })}
          />
        </div>
      </div>
      {g.textFitting ? (
        <div className="pptx-panel__section">
          <h3 className="pptx-panel__title">{t('panel.format.textFitting')}</h3>
          <div className="pptx-radios" role="radiogroup" aria-label={t('panel.format.textFitting')}>
            {(['none', 'shrink', 'shape'] as const).map((mode) => (
              <label key={mode} className="pptx-radio">
                <input
                  type="radio"
                  name="pptx-text-fitting"
                  checked={g.textFitting === mode}
                  disabled={readOnly}
                  onChange={() => editor?.exec({ type: 'setAutofit', mode })}
                />
                {t(`panel.format.fitting.${mode}`)}
              </label>
            ))}
          </div>
        </div>
      ) : null}
    </>
  );
}

/** Props of `SidePanel`. The host shows or hides the panel by rendering it. @public */
export interface SidePanelProps {
  /** Shows a close button that calls it. */
  readonly onClose?: () => void;
  /** A class name added to the panel. */
  readonly className?: string;
  /** The panel's width in CSS pixels. */
  readonly paneWidth?: number;
  /** Inline styles for the panel; a `width` here wins over `paneWidth`. */
  readonly style?: CSSProperties;
}

/** The side panel: Format · Comments. @public */
export function SidePanel({ onClose, className, paneWidth, style }: SidePanelProps) {
  const { t } = useTranslation();
  const [tab, setTab] = useState<'format' | 'comments'>('format');
  return (
    <aside
      className={['pptx-panel', className].filter(Boolean).join(' ')}
      aria-label={t('panel.ariaLabel')}
      style={paneStyle(paneWidth, style)}
      onMouseDown={keepFocus}
    >
      <div className="pptx-panel__head">
        <div
          className="pptx-tabs"
          role="tablist"
          // One Tab stop: the selected tab; the arrow keys move between tabs.
          onKeyDown={(e) => {
            const next = radioGroupStep(PANEL_TABS, tab, e.key);
            if (!next) return;
            e.preventDefault();
            setTab(next);
            e.currentTarget.querySelector<HTMLElement>(`[data-tab="${next}"]`)?.focus();
          }}
        >
          {PANEL_TABS.map((id) => (
            <button
              key={id}
              type="button"
              role="tab"
              className="pptx-tabs__tab"
              data-tab={id}
              aria-selected={tab === id}
              tabIndex={tab === id ? 0 : -1}
              onClick={() => setTab(id)}
            >
              {t(`panel.tabs.${id}`)}
            </button>
          ))}
        </div>
        {onClose ? <ToolButton label={t('panel.close')} icon="close" onClick={onClose} /> : null}
      </div>
      <div className="pptx-panel__body" role="tabpanel">
        {tab === 'format' ? <FormatTab /> : null}
        {tab === 'comments' ? (
          <div className="pptx-empty">
            <p className="pptx-empty__title">{t('panel.comments.emptyTitle')}</p>
            <p className="pptx-empty__body">{t('panel.comments.emptyBody')}</p>
          </div>
        ) : null}
      </div>
    </aside>
  );
}
