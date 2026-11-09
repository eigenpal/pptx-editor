/** @jsxImportSource vue */
// View > Slide Sorter: the Vue twin of the React adapter's `SlideSorter.tsx`, with the same
// markup, classes, ARIA, and keys. Every slide shows as a thumbnail in a grid, grouped by
// section. Click, Shift+click, and Ctrl/Cmd+click select slides; a drag moves the selection;
// Delete, Ctrl/Cmd+D, and the context menu act on it; a double-click or Enter opens the slide in
// the Normal view. Every edit goes through the same commands as the rail, so it is one undo
// step. The keys come from core (`sorterKeyDown`).

import {
  computed,
  defineComponent,
  onBeforeUnmount,
  onMounted,
  onUpdated,
  ref,
  shallowRef,
  watch,
  type ExtractPublicPropTypes,
  type VNodeChild,
} from 'vue';
import {
  STATUS_BAR_VIEW_OPTIONS,
  focusSorterOnShow,
  tooltipAttributes,
  groupSlideSections,
  focusCurrentSlideOrCanvas,
  openSlideInNormalView,
  refocusAfterSorterHides,
  sectionOfSlide,
  slideItemName,
  slideMenuPoint,
  slideSorterMounted,
  sorterKeyDown,
  trackNotesInset,
  translateKeyOrText,
  warnSlideSorterUnmounted,
  type NotesInsetTracker,
  type SlideSummary,
} from '@pptx-editor.dev/core/editor';
import { Icon, keepFocus, useAriaDisabled } from './chrome.tsx';
import { useEditorCommand, useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { useEditorDialog } from './popups.ts';
import {
  PptxEditorZoomPill,
  Thumbnail,
  clickMode,
  useFocusFollowsCurrent,
  useSlideStepTarget,
} from './panes.tsx';
import { RailMenuHost, type RailMenuState } from './RailContextMenu.tsx';
import {
  SectionNode,
  useCollapsedSections,
  useFocusAfterCollapse,
  useFocusCurrentAfterRender,
  useOpenCurrentSection,
  useSectionRenameRequests,
  useSlideSections,
  useSorterRegistration,
} from './RailSections.tsx';
import { useSlideDrag } from './slide-drag.ts';
import { PptxEditorEditNotice } from './EditNotice.tsx';

const sameList = <T,>(a: readonly T[], b: readonly T[]): boolean =>
  a.length === b.length && a.every((v, i) => v === b[i]);

const slideSorterProps = {
  /** A class name added to the sorter, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `PptxEditorSlideSorter`, the twin of React's `SlideSorterProps`. @public */
export type SlideSorterProps = ExtractPublicPropTypes<typeof slideSorterProps>;

/**
 * The slide sorter, the twin of React's `PptxEditor.SlideSorter`. Put it where the slide and the
 * rail go in the Normal view, while `snapshot().presentationView` is `slideSorter`. When the
 * view becomes `slideSorter` and no sorter of the editor is mounted, a development warning says
 * so. @public
 */
export const PptxEditorSlideSorter = defineComponent({
  name: 'PptxEditorSlideSorter',
  props: slideSorterProps,
  setup(props) {
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const slides = useEditorState((s) => s.slides, sameList);
    const current = useEditorState((s) => s.currentSlide);
    const selected = useEditorState((s) => s.selectedSlides, sameList);
    const ratio = useEditorState((s) => s.slideHeight / s.slideWidth);
    const readOnly = useEditorState((s) => s.mode === 'view');
    const sections = useSlideSections();
    const count = (): number => slides.value.length;
    const grouped = computed(() => groupSlideSections(sections.value, count()));
    const collapsed = useCollapsedSections();
    useOpenCurrentSection(grouped);
    useSorterRegistration();
    const list = ref<HTMLOListElement | null>(null);
    const drag = useSlideDrag('x', selected, readOnly, grouped);
    const menu = shallowRef<RailMenuState | null>(null);
    const renaming = ref<string | null>(null);
    // The sorter renames in place, also while it is flat: Add Section in a deck without
    // sections asks for the name before the new header renders.
    useSectionRenameRequests((id) => {
      renaming.value = id;
    });
    const focusCurrent = useFocusCurrentAfterRender(list);
    const trackCollapse = useFocusAfterCollapse(list, collapsed);
    // Shown from a menu, the sorter takes focus on the current slide (core's
    // `focusSorterOnShow`), after the hooks above listen for focus.
    let stopFocus = (): void => {};
    onMounted(() => {
      if (list.value) stopFocus = focusSorterOnShow(list.value);
    });
    // Hidden while one of its slides has focus, the sorter hands focus to the Normal view
    // (core's `refocusAfterSorterHides`). `onBeforeUnmount` runs while the list is on the page.
    onBeforeUnmount(() => {
      stopFocus();
      if (list.value) refocusAfterSorterHides(list.value, editor.value);
    });
    const background = useEditorDialog('background');
    const stepTo = useSlideStepTarget(
      () => selected.value,
      count,
      () => grouped.value
    );
    // Keyboard focus follows the current slide while a tile has it (after Ctrl+D, a drop, a
    // Delete of the last tile).
    const trackFocus = useFocusFollowsCurrent(list, '.pptx-sorter__item', () => [
      current.value,
      count(),
    ]);
    const openSlide = (index: number): void =>
      openSlideInNormalView(editor.value, index, list.value);
    const openMenu = (index: number, item: HTMLElement, at: { x: number; y: number } | null) => {
      const ed = editor.value;
      if (!ed) return;
      if (!ed.snapshot().selectedSlides.includes(index)) ed.selectSlide(index);
      menu.value = { at: slideMenuPoint(item, at), index, keyboard: !at, owner: item };
    };
    const onKeydown = (e: KeyboardEvent): void =>
      sorterKeyDown(e, {
        editor: editor.value,
        readOnly: readOnly.value,
        selected: selected.value,
        current: current.value,
        count: count(),
        grouped: grouped.value,
        list: list.value,
        step: stepTo,
        openMenu: (index, item) => openMenu(index, item, null),
        focusCurrent,
        openSlide,
      });
    return () => {
      const groups = grouped.value;
      // Tab goes to the current slide, or to its section's header when that section is shut.
      const own = sectionOfSlide(groups, current.value);
      const shut = !!own && collapsed.value.includes(own.id);
      const item = (s: SlideSummary): VNodeChild => (
        <li
          key={s.id}
          class="pptx-sorter__item"
          role={groups ? 'treeitem' : 'option'}
          aria-selected={selected.value.includes(s.index)}
          aria-current={s.index === current.value}
          data-hidden={s.hidden ? '' : undefined}
          data-index={s.index}
          {...drag.itemProps(s.index)}
          tabindex={s.index === current.value && !shut ? 0 : -1}
          onClick={(e: MouseEvent) => {
            (e.currentTarget as HTMLElement).focus({ preventScroll: true });
            editor.value?.selectSlide(s.index, clickMode(e));
          }}
          onDblclick={() => openSlide(s.index)}
          onContextmenu={(e: MouseEvent) => {
            e.preventDefault();
            const pointer = e.clientX !== 0 || e.clientY !== 0;
            openMenu(
              s.index,
              e.currentTarget as HTMLElement,
              pointer ? { x: e.clientX, y: e.clientY } : null
            );
          }}
        >
          <Thumbnail slide={s} ratio={ratio.value} />
          <span class="pptx-sorter__number" aria-hidden="true">
            {s.index + 1}
          </span>
          <span class="pptx-sr-only">{slideItemName(t, s)}</span>
        </li>
      );
      return (
        <section class={['pptx-sorter', props.className]} aria-label={t('sorter.ariaLabel')}>
          <ol
            ref={list}
            class="pptx-sorter__list"
            role={groups ? 'tree' : 'listbox'}
            aria-multiselectable="true"
            aria-label={t('sorter.ariaLabel')}
            onKeydown={onKeydown}
            onFocusin={(e: FocusEvent) => {
              trackFocus(e);
              trackCollapse(e);
            }}
          >
            {groups ? (
              groups.map((sec) => (
                <SectionNode
                  key={sec.id}
                  section={sec}
                  collapsed={collapsed.value.includes(sec.id)}
                  renaming={renaming.value === sec.id}
                  variant="sorter"
                  focusable={shut && own?.id === sec.id}
                  onRename={() => {
                    renaming.value = sec.id;
                  }}
                  onRenameEnd={() => {
                    renaming.value = null;
                  }}
                  onRemoved={focusCurrent}
                >
                  {() =>
                    slides.value.slice(sec.firstSlide, sec.firstSlide + sec.slideCount).map(item)
                  }
                </SectionNode>
              ))
            ) : (
              <li role="presentation" class="pptx-sorter__section">
                <ol role="group" class="pptx-sorter__grid">
                  {slides.value.map(item)}
                </ol>
              </li>
            )}
          </ol>
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
        </section>
      );
    };
  },
});

/**
 * Says that the open deck is read-only (`EditorSnapshot.readOnlyReason`), as React's
 * `ReadOnlyBadge`. Nothing when the deck can be edited. @internal
 */
export const ReadOnlyBadge = defineComponent({
  name: 'ReadOnlyBadge',
  setup() {
    const { t } = useTranslation();
    const damaged = useEditorState((s) => s.readOnlyReason === 'damaged-slide');
    return () =>
      damaged.value ? (
        <span class="pptx-statusbar__read-only" role="status" data-reason="damaged-slide">
          {t('statusBar.readOnlyDamaged')}
        </span>
      ) : null;
  },
});

/**
 * Normal and Slide Sorter as buttons, as in a status bar: the default editor's view switch, a
 * one-of-two choice (`radiogroup`). View > Slide Master opens the slide master view. Hosts
 * place the `view.presentationView` slot instead. The twin of React's `ViewSwitch`. @internal
 */
export const ViewSwitch = defineComponent({
  name: 'PptxViewSwitch',
  setup() {
    const { t } = useTranslation();
    const cmd = useEditorCommand('view.presentationView');
    const options = STATUS_BAR_VIEW_OPTIONS;
    const off = computed(() => !cmd.isEnabled.value);
    const reason = computed(() =>
      cmd.disabledReason.value ? translateKeyOrText(t, cmd.disabledReason.value) : null
    );
    // Off stays focusable, with the reason as its description.
    const offAttrs = useAriaDisabled(off, reason);
    const pick = (e: KeyboardEvent, step: number): void => {
      if (off.value) return;
      const at = options.findIndex((o) => o.value === cmd.value.value);
      const next = options[(at + step + options.length) % options.length];
      if (!next) return;
      cmd.execute(next.value);
      const buttons = (e.currentTarget as HTMLElement).parentElement?.querySelectorAll<HTMLElement>(
        '[role="radio"]'
      );
      buttons?.[options.indexOf(next)]?.focus();
    };
    const STEPS: Readonly<Record<string, number>> = {
      ArrowLeft: -1,
      ArrowUp: -1,
      ArrowRight: 1,
      ArrowDown: 1,
    };
    return () => (
      <div
        class="pptx-zoom pptx-view-switch"
        role="radiogroup"
        aria-label={t('toolbar.presentationView')}
        onMousedown={keepFocus}
      >
        {options.map((o, i) => {
          const on = cmd.value.value === o.value;
          // In the slide master view no option is checked: the first one keeps the tab stop.
          const tabStop = on || (i === 0 && !options.some((x) => x.value === cmd.value.value));
          const label = t(o.labelKey);
          return (
            <button
              key={o.value}
              type="button"
              role="radio"
              aria-checked={on}
              aria-label={label}
              {...offAttrs.attrs.value}
              {...tooltipAttributes(label, off.value, reason.value)}
              tabindex={tabStop ? 0 : -1}
              class="pptx-toolbar__button pptx-view-switch__option"
              data-active={on ? '' : undefined}
              data-disabled={off.value ? '' : undefined}
              onClick={() => {
                if (!off.value) cmd.execute(o.value);
              }}
              onKeydown={(e: KeyboardEvent) => {
                const step = STEPS[e.key];
                if (!step) return;
                e.preventDefault();
                pick(e, step);
              }}
            >
              <Icon paths={o.paths ?? null} />
            </button>
          );
        })}
        {offAttrs.note()}
      </div>
    );
  },
});

/**
 * The default editor's workspace, the twin of React's `DefaultWorkspace`: the `normal` slot
 * (the rail, the slide, and the notes) in the Normal view, the slide sorter in the Slide Sorter
 * view. The `status` slot (by default the view switch and the zoom) stays mounted at the bottom
 * right in both, so a view switch keeps the focus it has. When the view changes and focus was
 * in the part that went away, focus goes to the current slide in the new view. @internal
 */
export const DefaultWorkspace = defineComponent({
  name: 'PptxDefaultWorkspace',
  setup(_props, { slots }) {
    const editor = usePptxEditor();
    const view = useEditorState((s) => s.presentationView ?? 'normal');
    const box = ref<HTMLDivElement | null>(null);
    watch(
      view,
      () => {
        const active = box.value?.ownerDocument.activeElement;
        if (active && active !== active.ownerDocument.body && active.isConnected) return;
        focusCurrentSlideOrCanvas(box.value, editor.value);
      },
      { flush: 'post' }
    );
    // The status area and the font notice float above the notes card, at every notes height.
    let inset: NotesInsetTracker | null = null;
    onMounted(() => {
      if (box.value) inset = trackNotesInset(box.value);
    });
    // A view change or another `normal` slot can add or remove the notes card.
    onUpdated(() => inset?.refresh());
    onBeforeUnmount(() => {
      inset?.dispose();
      inset = null;
    });
    return () => (
      <div ref={box} class="pptx-default__views">
        {view.value === 'slideSorter' ? (
          <div class="pptx-default__workspace">
            <div class="pptx-default__main">
              <PptxEditorSlideSorter />
              <PptxEditorEditNotice />
            </div>
          </div>
        ) : (
          slots.normal?.()
        )}
        {slots.overlay?.()}
        {slots.status ? (
          slots.status()
        ) : (
          <div class="pptx-statusbar">
            <ReadOnlyBadge />
            <ViewSwitch />
            <PptxEditorZoomPill />
          </div>
        )}
      </div>
    );
  },
});

/**
 * Warns in development when the view becomes Slide Sorter and no slide sorter of the editor is
 * mounted, the twin of React's `SlideSorterWatch`. `PptxEditorRoot` renders it after its
 * default slot, so a sorter that mounts with the view change counts. @internal
 */
export const SlideSorterWatch = defineComponent({
  name: 'PptxSlideSorterWatch',
  setup() {
    const editor = usePptxEditor();
    const sorter = useEditorState((s) => s.presentationView === 'slideSorter');
    watch(
      [sorter, editor],
      ([on, ed]) => {
        if (on && ed && !slideSorterMounted(ed)) warnSlideSorterUnmounted();
      },
      { flush: 'post', immediate: true }
    );
    return () => null;
  },
});
