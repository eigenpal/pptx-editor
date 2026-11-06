/** @jsxImportSource vue */
// The slide rail, the zoom pill, and the side panel: the Vue twins of the React adapter's
// `panes.tsx`, with the same markup, classes, ARIA, and keys, so the one core stylesheet styles
// both. Each reads the editor through the composables; none holds document state of its own.
// The keys, the drag, and the focus moves come from core (`pane-support.ts`).

import { chromeTranslateProp, useChromeTranslateScope } from './headless-hooks.ts';
import {
  computed,
  defineComponent,
  nextTick,
  onBeforeUnmount,
  onMounted,
  onUnmounted,
  onUpdated,
  provide,
  ref,
  shallowRef,
  useId,
  watch,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
  type Ref,
  type StyleValue,
} from 'vue';
import {
  ZOOM_MENU_LEVELS,
  createProgressiveMount,
  mountedSections,
  tooltipAttributes,
  followCurrentSlide,
  groupSlideSections,
  keepCurrentSlideInView,
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
  slideStep,
  stepCommand,
  zoomStep,
  zoomLimitReason,
  type SlideSection,
  type SlideStep,
  type SlideRailOrientation,
  type SlideSummary,
} from '@pptx-editor.dev/core/editor';
import { Icon, MenuItem, Popover, ToolButton, keepFocus, useAriaDisabled } from './chrome.tsx';
import { useEditorDialog } from './popups.ts';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { RailMenuHost, type RailMenuState } from './RailContextMenu.tsx';
import {
  SlideRailAddSlide,
  SlideRailChooseLayout,
  SlideRailCollapse,
  SlideRailDelete,
  SlideRailDuplicate,
  SlideRailExpand,
  SlideRailMoveEarlier,
  SlideRailMoveLater,
  railActionContent,
  railActionsKey,
  railExpandContent,
} from './rail-parts.tsx';
import { ToolbarAction } from './toolbar-parts.ts';
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

export { PptxEditorSidePanel, type SidePanelProps } from './SidePanel.tsx';

const sameSlides = (a: readonly SlideSummary[], b: readonly SlideSummary[]): boolean =>
  a.length === b.length && a.every((s, i) => s === b[i]);
const sameIndices = (a: readonly number[], b: readonly number[]): boolean =>
  a.length === b.length && a.every((v, i) => v === b[i]);

/** How a click selects slides: Shift+click a range, Ctrl/Cmd+click one more or one less. @internal */
export const clickMode = slideClickMode;

/**
 * A slide thumbnail, repainted only when its slide renders differently. It is hidden from
 * assistive technology: the slide item that holds it gives the name ("Slide 3"). @internal
 */
export const Thumbnail = defineComponent({
  name: 'PptxThumbnail',
  props: {
    slide: { type: Object as PropType<SlideSummary>, required: true },
    ratio: { type: Number, required: true },
  },
  setup(props) {
    const editor = usePptxEditor();
    const box = ref<HTMLDivElement | null>(null);
    // A deferred paint: a long list paints the thumbnails in view first. The disposer runs on
    // unmount and before each new paint, so an old editor forgets the container.
    let dispose: (() => void) | null = null;
    const paint = (): void => {
      dispose?.();
      dispose =
        editor.value && box.value
          ? editor.value.renderThumbnail(props.slide.index, box.value, { defer: true })
          : null;
    };
    onMounted(paint);
    watch(() => [editor.value, props.slide] as const, paint, { flush: 'post' });
    onBeforeUnmount(() => {
      dispose?.();
      dispose = null;
    });
    return () => (
      <div
        ref={box}
        class="pptx-rail__thumb"
        data-pending={props.slide.pending ? '' : undefined}
        style={{ aspectRatio: String(1 / props.ratio) }}
        aria-hidden="true"
      />
    );
  },
});

/** The rail's handlers for its items: one object for the rail's life. */
interface RailItemHandlers {
  click(index: number, e: MouseEvent): void;
  menu(index: number, e: MouseEvent): void;
  dragStart(index: number, e: DragEvent): void;
  dragOver(index: number, e: DragEvent): void;
  dragEnd(index: number): void;
  drop(index: number, e: DragEvent): void;
}

/**
 * One slide of the rail, as React's memoized `RailItem`: a component, so a keystroke that
 * changes one slide, or a long deck whose rail grows while it opens, patches only the items
 * whose props changed. @internal
 */
const RailItem = defineComponent({
  name: 'PptxRailItem',
  props: {
    slide: { type: Object as PropType<SlideSummary>, required: true },
    ratio: { type: Number, required: true },
    selected: { type: Boolean, required: true },
    current: { type: Boolean, required: true },
    role: { type: String as PropType<'treeitem' | 'option'>, required: true },
    /** The item's name for assistive technology (`slideItemName`). */
    name: { type: String, required: true },
    drop: { type: String as PropType<'before' | 'after'>, default: undefined },
    draggable: { type: Boolean, required: true },
    on: { type: Object as PropType<RailItemHandlers>, required: true },
  },
  setup(props) {
    return () => {
      const { slide: s, on } = props;
      return (
        <li
          class="pptx-rail__item"
          role={props.role}
          aria-selected={props.selected}
          aria-current={props.current}
          data-hidden={s.hidden ? '' : undefined}
          data-index={s.index}
          data-drop={props.drop}
          draggable={props.draggable}
          onDragstart={(e: DragEvent) => on.dragStart(s.index, e)}
          onDragover={(e: DragEvent) => on.dragOver(s.index, e)}
          onDragend={() => on.dragEnd(s.index)}
          onDrop={(e: DragEvent) => on.drop(s.index, e)}
          tabindex={props.current ? 0 : -1}
          onClick={(e: MouseEvent) => on.click(s.index, e)}
          onContextmenu={(e: MouseEvent) => on.menu(s.index, e)}
        >
          {/* The name comes from the hidden text, not `aria-label`: the visible number is
              part of it, and the thumbnail art is hidden from assistive technology. */}
          <span class="pptx-rail__number" aria-hidden="true">
            {s.index + 1}
          </span>
          <span class="pptx-sr-only">{props.name}</span>
          <Thumbnail slide={s} ratio={props.ratio} />
        </li>
      );
    };
  },
});

/**
 * Keyboard focus follows the current slide of a slide list (the rail, the slide sorter) while
 * one of its items `item` has focus. When the focused item goes away, focus goes to the current
 * item too, so the next key still acts on the list. Returns the list's `focusin` handler.
 * @internal
 */
export function useFocusFollowsCurrent(
  list: Ref<HTMLElement | null>,
  item: string,
  watched: () => readonly (number | boolean | null)[]
): (e: FocusEvent) => void {
  let focused: HTMLElement | null = null;
  // Compare the values, not a new array: a deck edit gives new slide summaries, and the list
  // must not scroll back to the current slide when its index, the count, and the pane hold.
  watch(
    () => watched().join('|'),
    () => followCurrentSlide(list.value, item, focused),
    { flush: 'post' }
  );
  return (e) => {
    focused = (e.target as HTMLElement).closest?.<HTMLElement>(item) ?? null;
  };
}

/** Where the selected slides land one step up or down, or null at an edge. @internal */
export function useSlideStepTarget(
  selected: () => readonly number[],
  count: () => number,
  grouped: () => readonly SlideSection[] | null
): (step: -1 | 1) => SlideStep | null {
  return (step) => slideStep(selected(), count(), step, grouped());
}

/** The runtime props of every pane, after DOCX Editor's `Navigation`. */
const paneProps = {
  /** Whether the pane is open, controlled. Use with the `update:open` event (`v-model:open`). */
  open: { type: Boolean, default: undefined },
  /** Whether the pane is open first, uncontrolled. Default `true`. */
  defaultOpen: { type: Boolean, default: undefined },
  /** The pane's width in CSS pixels (the slide rail). */
  paneWidth: { type: Number, default: undefined },
  /** A class name added to the pane, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
  /** Inline styles for the pane, as React's `style`; a `width` here wins over `paneWidth`. */
  style: { type: Object as PropType<CSSProperties>, default: undefined },
} as const;

/**
 * Props of the editor's panes, the twin of React's `PaneProps`. A host `class` and `style` go
 * on the pane; a `width` in `style` wins over `paneWidth`. @public
 */
export type PaneProps = ExtractPublicPropTypes<typeof paneProps>;

const slideRailProps = {
  ...paneProps,
  ...chromeTranslateProp,
  /**
   * The built-in Hide thumbnails and Show thumbnails buttons. `false` removes them, for a host
   * that opens and closes the rail from its own chrome (`open`, `update:open`).
   */
  toggle: { type: Boolean, default: undefined },
  /**
   * Where the rail runs. `vertical` (the default) is a column beside the slide. `horizontal` is
   * a row under it, for a tall, narrow host: the thumbnails scroll sideways, Left and Right
   * step, a drag reorders along the row, and Hide thumbnails folds the rail down.
   */
  orientation: { type: String as PropType<SlideRailOrientation>, default: undefined },
  /**
   * `true` (default): the packaged action buttons, with named part children
   * (`PptxEditorSlideRail.AddSlide`, `.Delete`, …) in their places and other children after
   * Delete. `false`: only the default slot, in its order.
   */
  preset: { type: Boolean, default: true },
} as const;

/** Props of `PptxEditorSlideRail`, the twin of React's `SlideRailProps`. @public */
export type SlideRailProps = ExtractPublicPropTypes<typeof slideRailProps>;

/** A pane's inline style: its width from `paneWidth`, then the host's own styles. @internal */
export function paneStyle(width: number | undefined, style: unknown): StyleValue {
  return width === undefined
    ? (style as StyleValue)
    : [{ width: `${width}px` }, style as StyleValue];
}

/** The collapsed rail strip keeps the host's styles but not a width: it stays a narrow strip. */
function stripStyle(style: unknown): StyleValue {
  return style ? [style as StyleValue, { width: undefined }] : undefined;
}

/** The fallthrough attributes of a pane, without `class` and `style`, which it places itself. */
function restAttrs(attrs: Record<string, unknown>): Record<string, unknown> {
  const { class: _class, style: _style, ...rest } = attrs;
  return rest;
}

/** Open state of a pane: controlled by `open`, else its own. Emits `update:open`. @internal */
export function usePaneOpen(
  props: { readonly open: boolean | undefined; readonly defaultOpen: boolean | undefined },
  emit: (event: 'update:open', open: boolean) => void
): { readonly open: Readonly<Ref<boolean>>; setOpen(next: boolean): void } {
  const own = ref(props.defaultOpen ?? true);
  return {
    open: computed(() => props.open ?? own.value),
    setOpen: (next) => {
      if (props.open === undefined) own.value = next;
      emit('update:open', next);
    },
  };
}

/**
 * The slide thumbnail rail with drag-to-reorder, the twin of React's `PptxEditor.SlideRail`.
 * Its last button hides it to a narrow strip with a button that shows it again. Emits
 * `update:open`; a host `class` and `style` go on the rail. @public
 */
const SlideRailRoot = defineComponent({
  name: 'PptxEditorSlideRail',
  inheritAttrs: false,
  props: slideRailProps,
  emits: { 'update:open': (_open: boolean) => true },
  setup(props, { attrs, emit, slots }) {
    const pane = usePaneOpen(props, emit);
    const rail = ref<HTMLElement | null>(null);
    const listId = useId();
    // A toggle unmounts its own button: focus moves to the other toggle, never to the page.
    let focusToggle = false;
    const setOpen = (next: boolean): void => {
      focusToggle = rail.value?.contains(document.activeElement) ?? false;
      pane.setOpen(next);
      // A toggle the host refused leaves no flag behind for a later change to act on.
      void nextTick(() => {
        focusToggle = false;
      });
    };
    const editor = usePptxEditor();
    const { t } = useChromeTranslateScope(() => props.t);
    const slides = useEditorState((s) => s.slides, sameSlides);
    const current = useEditorState((s) => s.currentSlide);
    const selected = useEditorState((s) => s.selectedSlides, sameIndices);
    const ratio = useEditorState((s) => s.slideHeight / s.slideWidth);
    const readOnly = useEditorState((s) => s.mode === 'view');
    const list = ref<HTMLOListElement | null>(null);
    const count = (): number => slides.value.length;
    // While a long deck opens: the first screen with the first slide, then more after each
    // painted frame (core's `createProgressiveMount`).
    const opening = useEditorState((s) => (s.pendingSlides ?? 0) > 0);
    const grown = ref(0);
    const mount = createProgressiveMount(() => (grown.value += 1));
    onMounted(() => mount.follow(count()));
    onUpdated(() => mount.follow(count()));
    onUnmounted(() => mount.dispose());
    const menu = shallowRef<RailMenuState | null>(null);
    const gallery = shallowRef<HTMLElement | null>(null);
    const actions = ref<HTMLDivElement | null>(null);
    const sections = useSlideSections();
    // Sections group the rail when they cover every slide in order; otherwise it stays flat.
    const grouped = computed(() => groupSlideSections(sections.value, count()));
    const orientation = (): SlideRailOrientation =>
      resolveRailOrientation(props.orientation, 'PptxEditorSlideRail');
    const drag = useSlideDrag(() => railDragAxis(orientation()), selected, readOnly, grouped);
    // A right-to-left row mirrors the move buttons, as its arrow keys do. The direction comes
    // from the mounted rail, so it is read after each render.
    // A narrow horizontal rail (a phone) folds its secondary actions into one menu.
    const folded = shallowRef(false);
    watch(
      () => [rail.value, orientation(), pane.open.value] as const,
      ([el, axis], _old, onCleanup) =>
        onCleanup(
          watchRailFold(el, axis, (next) => {
            folded.value = next;
          })
        ),
      { flush: 'post', immediate: true }
    );
    const rightToLeft = shallowRef(false);
    const readDirection = (): void => {
      rightToLeft.value = railRightToLeft(rail.value);
    };
    onMounted(readDirection);
    onUpdated(readDirection);
    const collapsed = useCollapsedSections();
    const renaming = ref<string | null>(null);
    useOpenCurrentSection(grouped);
    // An open rail renames in place, also while it is flat: Add Section in a deck without
    // sections asks for the name before the new header renders. A closed rail has no header,
    // so it does not count.
    useSectionRenameRequests(
      (id) => {
        renaming.value = id;
      },
      () => pane.open.value
    );
    const focusCurrent = useFocusCurrentAfterRender(list);
    const trackCollapse = useFocusAfterCollapse(list, collapsed);
    const background = useEditorDialog('background');
    const stepTarget = useSlideStepTarget(
      () => selected.value,
      count,
      () => grouped.value
    );
    /**
     * Open the context menu on slide `index`. A slide outside the selection becomes the
     * selection; a slide inside a multi-selection keeps it.
     */
    const openMenu = (index: number, item: HTMLElement, at: { x: number; y: number } | null) => {
      const ed = editor.value;
      if (!ed) return;
      if (!ed.snapshot().selectedSlides.includes(index)) ed.selectSlide(index);
      menu.value = { at: slideMenuPoint(item, at), index, keyboard: at === null, owner: item };
    };
    const moveBy = (step: -1 | 1): void => {
      runSlideStep(editor.value, stepCommand(selected.value, stepTarget(step)));
    };
    const trackFocus = useFocusFollowsCurrent(list, '.pptx-rail__item', () => [
      current.value,
      count(),
      pane.open.value,
    ]);
    // A resize (a host narrowing the panel) keeps the current slide in view too.
    watch(
      list,
      (el, _old, onCleanup) => onCleanup(keepCurrentSlideInView(el, '.pptx-rail__item')),
      { flush: 'post' }
    );
    watch(
      () => pane.open.value,
      (open) => {
        // Only when the toggle's own unmount lost focus: a host that vetoed the change and
        // closes the rail later from its own chrome keeps its focus.
        const active = document.activeElement;
        const lost = !active || active === document.body || rail.value?.contains(active);
        if (focusToggle && lost)
          rail.value?.querySelector<HTMLElement>('.pptx-rail__collapse')?.focus();
        // A closed rail drops its open menu, drag, and rename.
        if (!open) {
          menu.value = null;
          drag.cancel();
          renaming.value = null;
        }
      },
      { flush: 'post' }
    );
    provide(railActionsKey, {
      t,
      editor: () => editor.value,
      why: () =>
        !editor.value ? t('disabled.noEditor') : readOnly.value ? t('disabled.viewing') : null,
      selected: () => selected.value,
      count,
      chrome: () => railOrientationChrome(orientation(), rightToLeft.value),
      step: stepTarget,
      moveBy,
      galleryOpen: () => !!gallery.value,
      // A second click closes the gallery, as the toolbar's menu buttons do.
      toggleGallery: () => {
        gallery.value = gallery.value
          ? null
          : // In a folded rail, the gallery opens at the overflow button that held its row.
            (actions.value?.querySelector<HTMLElement>('.pptx-rail__layouts, .pptx-rail__more') ??
            null);
      },
      listId,
      open: () => pane.open.value,
      toggle: () => props.toggle !== false,
      collapse: () => setOpen(false),
      expand: () => setOpen(true),
    });
    /** Slide shortcuts while the rail has focus (core's `railKeyDown`). */
    const onKeydown = (e: KeyboardEvent): void =>
      railKeyDown(e, {
        editor: editor.value,
        readOnly: readOnly.value,
        selected: selected.value,
        current: current.value,
        count: count(),
        grouped: grouped.value,
        list: list.value,
        orientation: orientation(),
        step: stepTarget,
        openMenu: (index, item) => openMenu(index, item, null),
        focusCurrent,
      });
    // The items' handlers: one object for the rail's life, so an item's props stay the same.
    const on: RailItemHandlers = {
      click: (index, e) => {
        (e.currentTarget as HTMLElement).focus({ preventScroll: true });
        editor.value?.selectSlide(index, clickMode(e));
      },
      menu: (index, e) => {
        e.preventDefault();
        // A keyboard-opened menu reports no pointer position: open at the slide.
        const pointer = e.clientX !== 0 || e.clientY !== 0;
        const at = pointer ? { x: e.clientX, y: e.clientY } : null;
        openMenu(index, e.currentTarget as HTMLElement, at);
      },
      dragStart: (index, e) => drag.itemProps(index).onDragstart(e),
      dragOver: (index, e) => drag.itemProps(index).onDragover(e),
      dragEnd: (index) => drag.itemProps(index).onDragend(),
      drop: (index, e) => drag.itemProps(index).onDrop(e),
    };
    /** One slide of the rail. */
    const item = (s: SlideSummary) => (
      <RailItem
        key={s.id}
        slide={s}
        ratio={ratio.value}
        selected={selected.value.includes(s.index)}
        current={s.index === current.value}
        role={grouped.value ? 'treeitem' : 'option'}
        name={slideItemName(t, s)}
        drop={drag.marker(s.index)}
        draggable={!readOnly.value}
        on={on}
      />
    );
    return () => {
      const toggle = props.toggle !== false;
      const open = pane.open.value;
      const className = ['pptx-rail', attrs.class, props.className];
      const style = [attrs.style as StyleValue, props.style];
      // Closed with no toggle, the rail leaves the page: the host's chrome opens it again.
      if (!open && !toggle) return null;
      if (!open)
        return (
          <nav
            {...restAttrs(attrs)}
            ref={rail}
            class={[...className, 'pptx-rail--collapsed']}
            aria-label={t('rail.ariaLabel')}
            data-open="false"
            data-orientation={orientation()}
            style={stripStyle(style)}
            onMousedown={keepFocus}
          >
            {railExpandContent(slots.default?.())}
          </nav>
        );
      const { content, shown } = railActionContent(
        slots.default?.(),
        props.preset !== false,
        toggle,
        folded.value
      );
      const groups = grouped.value;
      const currentSection = groups ? sectionOfSlide(groups, current.value)?.id : undefined;
      void grown.value;
      const mounted = mount.limit(count(), current.value, opening.value);
      // A grouped rail mounts whole sections, so no header shows with no slides under it.
      const inView = groups ? mountedSections(groups, mounted) : null;
      return (
        <nav
          {...restAttrs(attrs)}
          ref={rail}
          class={className}
          aria-label={t('rail.ariaLabel')}
          data-open="true"
          data-orientation={orientation()}
          style={paneStyle(props.paneWidth, style)}
        >
          <ol
            ref={list}
            id={listId}
            class="pptx-rail__list"
            role={groups ? 'tree' : 'listbox'}
            aria-multiselectable="true"
            aria-orientation={orientation()}
            aria-label={t('rail.ariaLabel')}
            onKeydown={onKeydown}
            onFocusin={(e: FocusEvent) => {
              trackFocus(e);
              trackCollapse(e);
            }}
          >
            {groups && inView
              ? inView.sections.map((sec) => {
                  const shut = collapsed.value.includes(sec.id);
                  return (
                    <SectionNode
                      key={sec.id}
                      section={sec}
                      collapsed={shut}
                      renaming={renaming.value === sec.id}
                      variant="rail"
                      // A collapsed section that holds the current slide takes Tab in its place.
                      focusable={shut && currentSection === sec.id}
                      onRename={() => {
                        renaming.value = sec.id;
                      }}
                      onRenameEnd={() => {
                        renaming.value = null;
                      }}
                      onRemoved={focusCurrent}
                    >
                      {() =>
                        slides.value
                          .slice(
                            sec.firstSlide,
                            Math.min(sec.firstSlide + sec.slideCount, inView.limit)
                          )
                          .map(item)
                      }
                    </SectionNode>
                  );
                })
              : slides.value.slice(0, mounted).map(item)}
          </ol>
          <div
            ref={actions}
            class="pptx-rail__actions"
            style={{ '--pptx-rail-action-columns': String(Math.max(1, Math.ceil(shown / 2))) }}
            onMousedown={keepFocus}
          >
            {content}
          </div>
          {gallery.value ? (
            <Popover
              anchor={gallery.value}
              onClose={() => {
                gallery.value = null;
              }}
              label={t('rail.chooseLayout')}
            >
              {() => (
                <LayoutGallery
                  onChange={(layoutId: string) => {
                    gallery.value = null;
                    editor.value?.exec({ type: 'addSlide', layoutId });
                  }}
                />
              )}
            </Popover>
          ) : null}
          <RailMenuHost
            open={
              menu.value
                ? {
                    menu: menu.value,
                    onClose: () => {
                      menu.value = null;
                    },
                    onSectionAdded: (id: string) => {
                      renaming.value = id;
                    },
                    ...(background.isAvailable.value
                      ? { onBackground: () => background.open(menu.value?.owner ?? null) }
                      : {}),
                  }
                : null
            }
          />
        </nav>
      );
    };
  },
});

/**
 * The slide rail with its action parts, as React's `PptxEditor.SlideRail`: a named part child
 * replaces its button, `hidden` removes it, other children go after Delete, and `preset: false`
 * renders only the default slot. `Action` is `ToolbarAction`. @public
 */
export const PptxEditorSlideRail = Object.assign(SlideRailRoot, {
  AddSlide: SlideRailAddSlide,
  ChooseLayout: SlideRailChooseLayout,
  Duplicate: SlideRailDuplicate,
  MoveEarlier: SlideRailMoveEarlier,
  MoveLater: SlideRailMoveLater,
  Delete: SlideRailDelete,
  Collapse: SlideRailCollapse,
  Expand: SlideRailExpand,
  Action: ToolbarAction,
});

const zoomPillProps = {
  /** A class name added to the pill, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `PptxEditorZoomPill`, the twin of React's `ZoomPillProps`. @public */
export type ZoomPillProps = ExtractPublicPropTypes<typeof zoomPillProps>;

/** The zoom pill (`– 100% ▾ +`), the twin of React's `PptxEditor.ZoomPill`. @public */
export const PptxEditorZoomPill = defineComponent({
  name: 'PptxEditorZoomPill',
  props: zoomPillProps,
  setup(props) {
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const zoom = useEditorState((s) => s.zoom);
    const fit = useEditorState((s) => s.fitZoom);
    // The pill zooms the slide; a thumbnail zoom for the slide sorter is a later change.
    const off = useEditorState((s) => s.presentationView === 'slideSorter');
    const reason = (): string | null => (off.value ? t('disabled.sorterZoom') : null);
    const anchor = ref<HTMLButtonElement | null>(null);
    const open = ref(false);
    // A view switch closes the zoom menu, so it does not open again on the way back.
    watch(off, (now) => {
      if (now) open.value = false;
    });
    const step = (dir: 1 | -1): void => {
      const next = zoomStep(zoom.value, dir);
      if (next) editor.value?.setZoom(next);
    };
    const levelOff = useAriaDisabled(off, reason);
    const pick = (z: number | 'fit'): void => {
      editor.value?.setZoom(z);
      open.value = false;
    };
    // A zoom button at the end of the range is off, and says why (zoom.ts).
    const limit = (dir: 1 | -1): string | null => {
      const key = zoomLimitReason(zoom.value, dir);
      return key ? t(key) : null;
    };
    return () => {
      // The sorter draws no slide at the zoom: the level shows no number there, only why it is off.
      const levelLabel = off.value
        ? t('zoom.ariaLabel')
        : t('zoom.levelAriaLabel', { percent: Math.round(zoom.value * 100) });
      return (
        <div
          class={['pptx-zoom', props.className]}
          role="group"
          aria-label={t('zoom.ariaLabel')}
          onMousedown={keepFocus}
        >
          <ToolButton
            label={t('zoom.zoomOut')}
            icon="remove"
            onClick={() => step(-1)}
            disabled={off.value || limit(-1) !== null}
            disabledReason={reason() ?? limit(-1)}
          />
          <button
            ref={anchor}
            type="button"
            class="pptx-zoom__level"
            aria-label={levelLabel}
            {...tooltipAttributes(levelLabel, off.value, reason())}
            aria-haspopup="menu"
            aria-expanded={open.value && !off.value}
            // Off stays focusable, so its tooltip says why; the reason is its
            // description too.
            {...levelOff.attrs.value}
            data-disabled={off.value ? '' : undefined}
            onClick={() => {
              if (!off.value) open.value = !open.value;
            }}
          >
            {off.value ? '\u2013' : t('zoom.level', { percent: Math.round(zoom.value * 100) })}
            <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
          </button>
          {levelOff.note()}
          <ToolButton
            label={t('zoom.zoomIn')}
            icon="add"
            onClick={() => step(1)}
            disabled={off.value || limit(1) !== null}
            disabledReason={reason() ?? limit(1)}
          />
          {open.value && !off.value ? (
            <Popover
              anchor={anchor.value}
              onClose={() => {
                open.value = false;
              }}
              align="end"
            >
              {() => [
                <MenuItem checked={fit.value} onSelect={() => pick('fit')}>
                  {() => t('zoom.fit')}
                </MenuItem>,
                ...ZOOM_MENU_LEVELS.map((z) => (
                  <MenuItem
                    key={z}
                    checked={!fit.value && Math.abs(z - zoom.value) < 0.001}
                    onSelect={() => pick(z)}
                  >
                    {() => t('zoom.level', { percent: Math.round(z * 100) })}
                  </MenuItem>
                )),
              ]}
            </Popover>
          ) : null}
        </div>
      );
    };
  },
});
