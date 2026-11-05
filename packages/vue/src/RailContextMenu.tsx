/** @jsxImportSource vue */
// The slide rail's context menu and the section header's menu: the Vue twins of the React
// adapter's `RailContextMenu.tsx` and `SectionContextMenu` (`RailSections.tsx`). Right-click a
// thumbnail (or press Shift+F10 or the Menu key on it): a slide outside the selection becomes
// the selection first; a slide inside a multi-selection keeps it, so Duplicate, Delete, and
// Move act on every selected slide. Every row asks the engine whether it can run
// (`commandRowReason`, `menu-arrangement.ts`). The rail and the sorter (`panes.tsx`,
// `RailSections.tsx`) open these menus.

import { defineComponent, type ExtractPublicPropTypes, type PropType, type VNodeChild } from 'vue';
import {
  addSectionAt,
  addSectionCommand,
  commandRowReason,
  groupSlideSections,
  layoutRowProbe,
  runSlideStep,
  slideStep,
  stepCommand,
  translateKeyOrText,
  type EditorCommand,
  type SlideSection,
} from '@pptx-editor.dev/core/editor';
import type { SlideLayoutSummary } from '@pptx-editor.dev/core/contracts';
import { ContextMenu, type ContextMenuPoint } from './ContextMenu.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { MenuRow, MenuSeparator, MenuSubmenu } from './menu-row.tsx';
import { usePopupEntry } from './popups.ts';
import { LayoutGallery } from './SlideControls.tsx';
import { useSlideshow } from './Slideshow.tsx';

const NO_LAYOUTS: readonly SlideLayoutSummary[] = [];
const NO_SECTIONS: readonly SlideSection[] = [];

const sameIndices = (a: readonly number[], b: readonly number[]): boolean =>
  a.length === b.length && a.every((v, i) => v === b[i]);

/** An open rail menu: where, on which slide, and how it opened. @public */
export interface RailMenuState {
  /** Where the menu opens: the pointer, or the slide for a keyboard open. */
  readonly at: ContextMenuPoint;
  /** The slide the menu opened on, from 0. */
  readonly index: number;
  /** Opened from the keyboard: focus goes to the first row. */
  readonly keyboard: boolean;
  /** The element the menu opened on; its editor scope carries the theme tokens. */
  readonly owner: HTMLElement | null;
}

const railContextMenuProps = {
  /** Where the menu opens, on which slide, and how. */
  menu: { type: Object as PropType<RailMenuState>, required: true as const },
  /** Closes the menu. */
  onClose: { type: Function as PropType<() => void>, required: true as const },
  /** A section was added at the slide: the rail starts renaming it. */
  onSectionAdded: { type: Function as PropType<(sectionId: string) => void>, default: undefined },
  /** Opens the background dialog; the row is left out when the dialog is turned off. */
  onBackground: { type: Function as PropType<() => void>, default: undefined },
} as const;

/**
 * Props of the context menu of a slide in the rail or the sorter, and of a
 * `popups.railContextMenu` renderer, the twin of React's `RailContextMenuProps`. @public
 */
export type RailContextMenuProps = ExtractPublicPropTypes<typeof railContextMenuProps>;

/** One row bound to a command: disabled with the engine's reason when `can` refuses it. */
interface CommandRow {
  readonly label: string;
  readonly icon: string;
  readonly slot: string;
  readonly shortcut?: string;
}

/**
 * The slide rail's context menu: new, duplicate, delete, and move slides, background, and
 * present from here. Rows act on every selected slide. @public
 */
export const RailContextMenu = defineComponent({
  name: 'RailContextMenu',
  props: railContextMenuProps,
  setup(props) {
    const editor = usePptxEditor();
    const show = useSlideshow();
    const { t } = useTranslation();
    // Re-render on every change, so each row's state is the engine's current answer.
    const revision = useEditorState((s) => s.revision);
    const selectedRef = useEditorState((s) => s.selectedSlides, sameIndices);
    const countRef = useEditorState((s) => s.slides.length);
    const readOnly = useEditorState((s) => s.mode === 'view');
    const currentRef = useEditorState((s) => s.slides[s.currentSlide]);
    const slidesRef = useEditorState((s) => s.slides);
    const layoutsRef = useEditorState((s) => s.layouts ?? NO_LAYOUTS);
    const sectionsRef = useEditorState((s) => s.sections ?? NO_SECTIONS);
    return () => {
      void revision.value;
      const close = (): void => props.onClose();
      const selected = selectedRef.value;
      const count = countRef.value;
      const many = selected.length > 1;
      const row = (command: EditorCommand | null, spec: CommandRow): VNodeChild => {
        const reason = commandRowReason(editor.value, command);
        return (
          <MenuRow
            {...spec}
            disabled={!!reason}
            reason={reason ? translateKeyOrText(t, reason) : null}
            onSelect={() => {
              close();
              // A move step opens a collapsed section it lands in (`runSlideStep`).
              if (command?.type === 'moveSlides') runSlideStep(editor.value, command);
              else if (command) editor.value?.exec(command);
            }}
          />
        );
      };
      const grouped = groupSlideSections(sectionsRef.value, count);
      const up = stepCommand(selected, slideStep(selected, count, -1, grouped));
      const down = stepCommand(selected, slideStep(selected, count, 1, grouped));
      const canPresent = count > 0 && !!editor.value;
      const addSection = addSectionCommand(props.menu.index, t);
      const sectionCan = editor.value ? editor.value.can(addSection) : null;
      const layoutReason = commandRowReason(
        editor.value,
        layoutRowProbe(layoutsRef.value, currentRef.value?.layoutId, selected)
      );
      // Checked when every selected slide is hidden; the row then shows them, else hides them.
      const hidden = selected.length > 0 && selected.every((i) => slidesRef.value[i]?.hidden);
      const hide: EditorCommand = { type: 'setSlidesHidden', indices: selected, hidden: !hidden };
      const hideCan = editor.value ? editor.value.can(hide) : null;
      return (
        <ContextMenu
          at={props.menu.at}
          onClose={close}
          label={t('contextMenu.railLabel')}
          keyboard={props.menu.keyboard}
          owner={props.menu.owner}
        >
          {{
            default: () => (
              <>
                {row(
                  { type: 'addSlide' },
                  { label: t('contextMenu.newSlide'), icon: 'add', slot: 'slide.add' }
                )}
                {row(
                  { type: 'duplicateSlides', indices: selected },
                  {
                    label: many
                      ? t('contextMenu.duplicateSlides', { count: selected.length })
                      : t('rail.duplicateSlide'),
                    icon: 'content_copy',
                    shortcut: t('contextMenu.shortcuts.duplicate'),
                    slot: 'slide.duplicate',
                  }
                )}
                {row(
                  { type: 'deleteSlides', indices: selected },
                  {
                    label: many
                      ? t('contextMenu.deleteSlides', { count: selected.length })
                      : t('rail.deleteSlide'),
                    icon: 'delete',
                    shortcut: t('contextMenu.shortcuts.delete'),
                    slot: 'slide.delete',
                  }
                )}
                <MenuRow
                  label={t('sections.add')}
                  icon="view_agenda"
                  slot="section.add"
                  disabled={!sectionCan?.ok}
                  reason={
                    sectionCan && !sectionCan.ok ? translateKeyOrText(t, sectionCan.reason) : null
                  }
                  onSelect={() => {
                    close();
                    if (!editor.value) return;
                    const added = addSectionAt(editor.value, addSection, props.menu.index);
                    if (added) props.onSectionAdded?.(added);
                  }}
                />
                <MenuSeparator />
                <MenuSubmenu
                  label={t('contextMenu.layout')}
                  icon="dashboard"
                  disabled={!!layoutReason}
                  reason={layoutReason ? translateKeyOrText(t, layoutReason) : null}
                >
                  {{
                    default: () => (
                      <LayoutGallery
                        value={currentRef.value?.layoutId ?? ''}
                        onChange={(layoutId: string) => {
                          close();
                          editor.value?.exec({ type: 'applyLayout', indices: selected, layoutId });
                        }}
                      />
                    ),
                  }}
                </MenuSubmenu>
                {row(
                  { type: 'resetSlides', indices: selected },
                  { label: t('contextMenu.resetSlide'), icon: 'restart_alt', slot: 'slide.reset' }
                )}
                {props.onBackground ? (
                  <MenuRow
                    label={t('contextMenu.background')}
                    icon="wallpaper"
                    slot="slide.background"
                    disabled={readOnly.value}
                    reason={t('disabled.viewing')}
                    onSelect={() => {
                      close();
                      props.onBackground?.();
                    }}
                  />
                ) : null}
                <MenuRow
                  label={t('contextMenu.hideSlide')}
                  icon="visibility_off"
                  slot="slide.hide"
                  checked={hidden}
                  disabled={!hideCan?.ok}
                  reason={hideCan && !hideCan.ok ? translateKeyOrText(t, hideCan.reason) : null}
                  onSelect={() => {
                    close();
                    editor.value?.exec(hide);
                  }}
                />
                <MenuSeparator />
                {row(up, {
                  label: t('contextMenu.moveUp'),
                  icon: 'keyboard_arrow_up',
                  shortcut: t('contextMenu.shortcuts.moveUp'),
                  slot: 'slide.moveUp',
                })}
                {row(down, {
                  label: t('contextMenu.moveDown'),
                  icon: 'keyboard_arrow_down',
                  shortcut: t('contextMenu.shortcuts.moveDown'),
                  slot: 'slide.moveDown',
                })}
                <MenuSeparator />
                <MenuRow
                  label={t('contextMenu.presentFromHere')}
                  icon="slideshow"
                  slot="slide.present"
                  disabled={!canPresent}
                  reason={t('disabled.noEditor')}
                  onSelect={() => {
                    close();
                    editor.value?.goToSlide(props.menu.index);
                    show.start('current');
                  }}
                />
              </>
            ),
          }}
        </ContextMenu>
      );
    };
  },
});

/**
 * The slide menu that the rail and the sorter open, as the `railContextMenu` popup entry says:
 * `false` shows none, a function renders the host's menu, otherwise the packaged menu.
 * @internal
 */
export const RailMenuHost = defineComponent({
  name: 'PptxRailMenuHost',
  props: {
    open: { type: Object as PropType<RailContextMenuProps | null>, default: null },
  },
  setup(props) {
    const entry = usePopupEntry('railContextMenu');
    return () => {
      const renderer = entry();
      if (!props.open || renderer === false) return null;
      return typeof renderer === 'function' ? (
        renderer(props.open)
      ) : (
        <RailContextMenu {...props.open} />
      );
    };
  },
});

/** An open section menu: where it opens, how it opened, and the element it belongs to. @public */
export interface SectionMenuState {
  /** The point where the menu opens, in viewport pixels. */
  readonly at: ContextMenuPoint;
  /**
   * Whether a key opened the menu (Shift+F10 or the Menu key). The menu then puts focus on its
   * first row; after a right-click, on the menu itself.
   */
  readonly keyboard: boolean;
  /** The element the menu belongs to: the section header. Focus goes back to it on close. */
  readonly owner: HTMLElement | null;
}

const sectionContextMenuProps = {
  /** The section whose header opened the menu. */
  section: { type: Object as PropType<SlideSection>, required: true as const },
  /** Where the menu opens, and how. */
  menu: { type: Object as PropType<SectionMenuState>, required: true as const },
  /** Closes the menu. Call it when a row runs or the user dismisses the menu. */
  onClose: { type: Function as PropType<() => void>, required: true as const },
  /** Starts renaming the section in place, in the rail or the sorter that opened the menu. */
  onRename: { type: Function as PropType<() => void>, required: true as const },
  /** Call after a row removes sections: focus goes to the current slide after the render. */
  onRemoved: { type: Function as PropType<() => void>, required: true as const },
} as const;

/**
 * Props of the section menu, and of a `popups.sectionContextMenu` renderer, the twin of
 * React's `SectionContextMenuProps`. @public
 */
export type SectionContextMenuProps = ExtractPublicPropTypes<typeof sectionContextMenuProps>;

/**
 * The section menu: Rename, Remove, Remove with slides, Remove all, Move up and down, Collapse
 * and Expand all. Every row asks the engine whether it can run. @public
 */
export const SectionContextMenu = defineComponent({
  name: 'SectionContextMenu',
  props: sectionContextMenuProps,
  setup(props) {
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const revision = useEditorState((s) => s.revision);
    return () => {
      void revision.value;
      const id = props.section.id;
      const row = (
        command: EditorCommand,
        label: string,
        icon: string,
        slot: string,
        then?: () => void
      ): VNodeChild => {
        const reason = commandRowReason(editor.value, command);
        return (
          <MenuRow
            label={label}
            icon={icon}
            slot={slot}
            disabled={!!reason}
            reason={reason ? translateKeyOrText(t, reason) : null}
            onSelect={() => {
              props.onClose();
              if (editor.value?.exec(command).ok) then?.();
            }}
          />
        );
      };
      const removed = (): void => props.onRemoved();
      // Rename runs no command until the name is typed; it needs editing mode. The probe uses a
      // fixed valid name: the section's own name can come from a file and be invalid.
      const rename = editor.value?.can({ type: 'renameSection', sectionId: id, name: 'x' });
      return (
        <ContextMenu
          at={props.menu.at}
          onClose={() => props.onClose()}
          label={t('sections.menuLabel')}
          keyboard={props.menu.keyboard}
          owner={props.menu.owner}
        >
          {{
            default: () => (
              <>
                <MenuRow
                  label={t('sections.rename')}
                  icon="edit_note"
                  slot="section.rename"
                  disabled={!rename?.ok}
                  reason={rename && !rename.ok ? translateKeyOrText(t, rename.reason) : null}
                  onSelect={() => {
                    props.onClose();
                    props.onRename();
                  }}
                />
                {row(
                  { type: 'removeSection', sectionId: id },
                  t('sections.remove'),
                  'remove',
                  'section.remove',
                  removed
                )}
                {row(
                  { type: 'removeSection', sectionId: id, withSlides: true },
                  t('sections.removeWithSlides'),
                  'delete',
                  'section.removeWithSlides',
                  removed
                )}
                {row(
                  { type: 'removeAllSections' },
                  t('sections.removeAll'),
                  'delete',
                  'section.removeAll',
                  removed
                )}
                <MenuSeparator />
                {row(
                  { type: 'moveSection', sectionId: id, direction: -1 },
                  t('sections.moveUp'),
                  'keyboard_arrow_up',
                  'section.moveUp'
                )}
                {row(
                  { type: 'moveSection', sectionId: id, direction: 1 },
                  t('sections.moveDown'),
                  'keyboard_arrow_down',
                  'section.moveDown'
                )}
                <MenuSeparator />
                {row(
                  { type: 'setSectionsCollapsed', collapsed: true },
                  t('sections.collapseAll'),
                  'unfold_less',
                  'section.collapseAll'
                )}
                {row(
                  { type: 'setSectionsCollapsed', collapsed: false },
                  t('sections.expandAll'),
                  'unfold_more',
                  'section.expandAll'
                )}
              </>
            ),
          }}
        </ContextMenu>
      );
    };
  },
});
