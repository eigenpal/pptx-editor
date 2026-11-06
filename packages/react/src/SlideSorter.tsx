/** @jsxImportSource react */
// View > Slide Sorter: every slide as a thumbnail in a grid, grouped by section, as a
// Slide Sorter view shows them. Click, Shift+click, and Ctrl/Cmd+click select slides; a drag
// moves the selection; Delete, Ctrl/Cmd+D, and the context menu act on it; a double-click or
// Enter opens the slide in the Normal view. Every edit goes through the same commands as the
// rail, so it is one undo step.

import {
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type KeyboardEvent,
  type ReactNode,
} from 'react';
import {
  chromeControlForSlot,
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
import { useEditorDialog } from './popups.tsx';
import { useEditorCommand, useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { Icon, keepFocus, useAriaDisabled } from './chrome.tsx';
import { clickMode, Thumbnail, useFocusFollowsCurrent, ZoomPill } from './panes.tsx';
import { RailMenuHost, useSlideStep, type RailMenuState } from './RailContextMenu.tsx';
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
import { EditNotice } from './EditNotice.tsx';

const sameList = <T,>(a: readonly T[], b: readonly T[]): boolean =>
  a.length === b.length && a.every((v, i) => v === b[i]);

/** Props of {@link SlideSorter}. @public */
export interface SlideSorterProps {
  /** A class name added to the sorter. */
  readonly className?: string;
}

/**
 * The slide sorter. Put it where the slide and the rail go in the Normal view, while
 * `snapshot().presentationView` is `slideSorter`. When the view becomes `slideSorter` and no
 * sorter of the editor is mounted, a development warning says so. @public
 */
export function SlideSorter({ className }: SlideSorterProps = {}) {
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const slides = useEditorState((s) => s.slides, sameList);
  const current = useEditorState((s) => s.currentSlide);
  const selected = useEditorState((s) => s.selectedSlides, sameList);
  const ratio = useEditorState((s) => s.slideHeight / s.slideWidth);
  const readOnly = useEditorState((s) => s.mode === 'view');
  const sections = useSlideSections();
  const grouped = groupSlideSections(sections, slides.length);
  const collapsed = useCollapsedSections();
  useOpenCurrentSection(grouped);
  useSorterRegistration();
  const list = useRef<HTMLOListElement>(null);
  const { itemProps: dragProps } = useSlideDrag('x', selected, readOnly, grouped);
  const [menu, setMenu] = useState<RailMenuState | null>(null);
  const [renaming, setRenaming] = useState<string | null>(null);
  // The sorter renames in place, also while it is flat: Add Section in a deck without sections
  // asks for the name before the new header renders. The id waits in `renaming` until the
  // header mounts with its field.
  useSectionRenameRequests(setRenaming);
  const focusCurrent = useFocusCurrentAfterRender(list);
  useFocusAfterCollapse(list, collapsed);
  // Shown from a menu, the sorter takes focus on the current slide (core's `focusSorterOnShow`),
  // after the hooks above listen for focus.
  useEffect(() => (list.current ? focusSorterOnShow(list.current) : undefined), []);
  // Hidden while one of its slides has focus, the sorter hands focus to the Normal view (core's
  // `refocusAfterSorterHides`). A layout effect's cleanup runs while the list is still on the page.
  const editorNow = useRef(editor);
  editorNow.current = editor;
  useLayoutEffect(() => {
    const el = list.current;
    return () => {
      if (el) refocusAfterSorterHides(el, editorNow.current);
    };
  }, []);
  // Tab goes to the current slide, or to its section's header when that section is shut.
  const own = sectionOfSlide(grouped, current);
  const shut = !!own && collapsed.includes(own.id);
  const background = useEditorDialog('background');
  const count = slides.length;
  const stepTo = useSlideStep(selected, count);
  // Keyboard focus follows the current slide while a tile has it (after Ctrl+D, a drop, a
  // Delete of the last tile).
  const trackFocus = useFocusFollowsCurrent(list, '.pptx-sorter__item', current, count);

  /** Open slide `index` in the Normal view. */
  const openSlide = (index: number): void => openSlideInNormalView(editor, index, list.current);
  const openMenu = (index: number, item: HTMLElement, at: { x: number; y: number } | null) => {
    if (!editor) return;
    if (!editor.snapshot().selectedSlides.includes(index)) editor.selectSlide(index);
    setMenu({ at: slideMenuPoint(item, at), index, keyboard: !at, owner: item });
  };
  /** The sorter's keys (core's `sorterKeyDown`). */
  const onKeyDown = (e: KeyboardEvent): void =>
    sorterKeyDown(e, {
      editor,
      readOnly,
      selected,
      current,
      count,
      grouped,
      list: list.current,
      step: stepTo,
      openMenu: (index, item) => openMenu(index, item, null),
      focusCurrent,
      openSlide,
    });

  const item = (s: SlideSummary): ReactNode => (
    <li
      key={s.id}
      className="pptx-sorter__item"
      role={grouped ? 'treeitem' : 'option'}
      aria-selected={selected.includes(s.index)}
      aria-current={s.index === current}
      data-hidden={s.hidden ? '' : undefined}
      data-index={s.index}
      {...dragProps(s.index)}
      tabIndex={s.index === current && !shut ? 0 : -1}
      onClick={(e) => {
        e.currentTarget.focus({ preventScroll: true });
        editor?.selectSlide(s.index, clickMode(e));
      }}
      onDoubleClick={() => openSlide(s.index)}
      onContextMenu={(e) => {
        e.preventDefault();
        const pointer = e.clientX !== 0 || e.clientY !== 0;
        openMenu(s.index, e.currentTarget, pointer ? { x: e.clientX, y: e.clientY } : null);
      }}
    >
      <Thumbnail slide={s} ratio={ratio} />
      <span className="pptx-sorter__number" aria-hidden="true">
        {s.index + 1}
      </span>
      <span className="pptx-sr-only">{slideItemName(t, s)}</span>
    </li>
  );

  return (
    <section
      className={['pptx-sorter', className].filter(Boolean).join(' ')}
      aria-label={t('sorter.ariaLabel')}
    >
      <ol
        ref={list}
        className="pptx-sorter__list"
        role={grouped ? 'tree' : 'listbox'}
        aria-multiselectable="true"
        aria-label={t('sorter.ariaLabel')}
        onKeyDown={onKeyDown}
        onFocus={trackFocus}
      >
        {grouped ? (
          grouped.map((sec) => (
            <SectionNode
              key={sec.id}
              section={sec}
              collapsed={collapsed.includes(sec.id)}
              renaming={renaming === sec.id}
              variant="sorter"
              focusable={shut && own?.id === sec.id}
              onRename={() => setRenaming(sec.id)}
              onRenameEnd={() => setRenaming(null)}
              onRemoved={focusCurrent}
            >
              {slides.slice(sec.firstSlide, sec.firstSlide + sec.slideCount).map(item)}
            </SectionNode>
          ))
        ) : (
          <li role="presentation" className="pptx-sorter__section">
            <ol role="group" className="pptx-sorter__grid">
              {slides.map(item)}
            </ol>
          </li>
        )}
      </ol>
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
    </section>
  );
}

/**
 * Normal and Slide Sorter as two buttons, as in a status bar: the default editor's
 * view switch, a one-of-two choice (`radiogroup`). Hosts place the `view.presentationView` slot
 * instead. @internal
 */
export function ViewSwitch() {
  const { t } = useTranslation();
  const cmd = useEditorCommand('view.presentationView');
  const options = chromeControlForSlot('view.presentationView')?.options ?? [];
  const off = !cmd.isEnabled;
  const reason = cmd.disabledReason ? translateKeyOrText(t, cmd.disabledReason) : null;
  // Off stays focusable, with the reason as its description.
  const offAttrs = useAriaDisabled(off, reason);
  const pick = (e: KeyboardEvent<HTMLElement>, step: number): void => {
    if (off) return;
    const at = options.findIndex((o) => o.value === cmd.value);
    const next = options[(at + step + options.length) % options.length];
    if (!next) return;
    cmd.execute(next.value);
    const buttons = e.currentTarget.parentElement?.querySelectorAll<HTMLElement>('[role="radio"]');
    buttons?.[options.indexOf(next)]?.focus();
  };
  return (
    <div
      className="pptx-zoom pptx-view-switch"
      role="radiogroup"
      aria-label={t('toolbar.presentationView')}
      onMouseDown={keepFocus}
    >
      {options.map((o) => {
        const on = cmd.value === o.value;
        const label = t(o.labelKey);
        return (
          <button
            key={o.value}
            type="button"
            role="radio"
            aria-checked={on}
            aria-label={label}
            {...offAttrs.attrs}
            {...tooltipAttributes(label, off, reason)}
            tabIndex={on ? 0 : -1}
            className="pptx-toolbar__button pptx-view-switch__option"
            data-active={on ? '' : undefined}
            data-disabled={off ? '' : undefined}
            onClick={() => !off && cmd.execute(o.value)}
            onKeyDown={(e) => {
              const step = { ArrowLeft: -1, ArrowUp: -1, ArrowRight: 1, ArrowDown: 1 }[e.key];
              if (!step) return;
              e.preventDefault();
              pick(e, step);
            }}
          >
            <Icon paths={o.paths ?? null} />
          </button>
        );
      })}
      {offAttrs.note}
    </div>
  );
}

/**
 * The default editor's workspace: `normal` (the rail, the slide, and the notes) in the Normal
 * view, the slide sorter in the Slide Sorter view. `status` (by default the view switch and the
 * zoom) stays mounted at the bottom right in both, so a view switch keeps the focus it has. When
 * the view changes and focus was in the part that went away, focus goes to the current slide in
 * the new view. `overlay` (the open-failure notice) shows over both views. @internal
 */
export function DefaultWorkspace({
  normal,
  status,
  overlay,
}: {
  normal: ReactNode;
  status?: ReactNode;
  overlay?: ReactNode;
}) {
  const editor = usePptxEditor();
  const view = useEditorState((s) => s.presentationView ?? 'normal');
  const box = useRef<HTMLDivElement>(null);
  const was = useRef(view);
  useEffect(() => {
    if (was.current === view) return;
    was.current = view;
    const active = box.current?.ownerDocument.activeElement;
    if (active && active !== active.ownerDocument.body && active.isConnected) return;
    focusCurrentSlideOrCanvas(box.current, editor);
  }, [view, editor]);
  // The status area and the font notice float above the notes card, at every notes height.
  const inset = useRef<NotesInsetTracker | null>(null);
  useEffect(() => {
    if (!box.current) return;
    const tracker = trackNotesInset(box.current);
    inset.current = tracker;
    return () => {
      tracker.dispose();
      inset.current = null;
    };
  }, []);
  // A view change or another `normal` part can add or remove the notes card.
  useEffect(() => {
    inset.current?.refresh();
  });
  return (
    <div ref={box} className="pptx-default__views">
      {view === 'slideSorter' ? (
        <div className="pptx-default__workspace">
          <div className="pptx-default__main">
            <SlideSorter />
            <EditNotice />
          </div>
        </div>
      ) : (
        normal
      )}
      {overlay}
      {status === undefined ? (
        <div className="pptx-statusbar">
          <ReadOnlyBadge />
          <ViewSwitch />
          <ZoomPill />
        </div>
      ) : (
        status
      )}
    </div>
  );
}

/**
 * Says that the open deck is read-only (`EditorSnapshot.readOnlyReason`): a slide of a long
 * deck did not parse after edits. Nothing when the deck can be edited. @internal
 */
export function ReadOnlyBadge(): ReactNode {
  const { t } = useTranslation();
  const damaged = useEditorState((s) => s.readOnlyReason === 'damaged-slide');
  if (!damaged) return null;
  return (
    <span className="pptx-statusbar__read-only" role="status" data-reason="damaged-slide">
      {t('statusBar.readOnlyDamaged')}
    </span>
  );
}

/**
 * Warns in development when the view becomes Slide Sorter and no slide sorter of the editor is
 * mounted. `PptxEditor.Root` renders it after its children, so a sorter that mounts with the
 * view change counts. @internal
 */
export function SlideSorterWatch(): null {
  const editor = usePptxEditor();
  const sorter = useEditorState((s) => s.presentationView === 'slideSorter');
  useEffect(() => {
    if (sorter && editor && !slideSorterMounted(editor)) warnSlideSorterUnmounted();
  }, [sorter, editor]);
  return null;
}
