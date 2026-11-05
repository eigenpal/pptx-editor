/** @jsxImportSource react */
// The slide rail's context menu: right-click a thumbnail (or press Shift+F10 or the Menu key on
// it). A slide outside the selection becomes the selection first; a slide inside a
// multi-selection keeps it, so Duplicate, Delete, and Move act on every selected slide, as in
// other slide editors such as Google Slides.

import type { ReactNode } from 'react';
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
  type SlideStep,
} from '@pptx-editor.dev/core/editor';
import type { EditorCommand } from '@pptx-editor.dev/core/editor';
import type { SlideLayoutSummary } from '@pptx-editor.dev/core/contracts';
import {
  ContextMenu,
  MenuRow,
  MenuSeparator,
  MenuSubmenu,
  type ContextMenuPoint,
} from './ContextMenu.tsx';
import { LayoutGallery } from './SlideControls.tsx';
import { useEditorState, usePptxEditor, useSlideshow, useTranslation } from './context.ts';
import { usePopupEntry } from './popups.tsx';
import { useSlideSections } from './RailSections.tsx';

const NO_LAYOUTS: readonly SlideLayoutSummary[] = [];

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

/**
 * Where the selected slides land one step up or down, or null at the top or bottom
 * (`slideStep`): a step across a section header changes the section first.
 */
export function useSlideStep(
  selected: readonly number[],
  count: number
): (step: -1 | 1) => SlideStep | null {
  const grouped = groupSlideSections(useSlideSections(), count);
  return (step) => slideStep(selected, count, step, grouped);
}

/**
 * The slide menu that the rail and the sorter open, as the `railContextMenu` popup entry says:
 * `false` shows none, a function renders the host's menu, otherwise the packaged menu.
 */
export function RailMenuHost({ open }: { open: RailContextMenuProps | null }): ReactNode {
  const entry = usePopupEntry('railContextMenu');
  if (!open || entry === false) return null;
  const props = open;
  return typeof entry === 'function' ? entry(props) : <RailContextMenu {...props} />;
}

/** Props of the context menu of a slide in the rail or the sorter, and of a `popups.railContextMenu` renderer. @public */
export interface RailContextMenuProps {
  /** Where the menu opens, on which slide, and how. */
  readonly menu: RailMenuState;
  /** Called when the menu closes. */
  readonly onClose: () => void;
  /** A section was added at the slide: the rail starts renaming it. */
  readonly onSectionAdded?: (sectionId: string) => void;
  /** Opens the background dialog; the row is left out when the dialog is turned off. */
  readonly onBackground?: () => void;
}

/**
 * The slide rail's context menu: new, duplicate, delete, and move slides, background, and
 * present from here. Rows act on every selected slide. @public
 */
export function RailContextMenu({
  menu,
  onClose,
  onBackground,
  onSectionAdded,
}: RailContextMenuProps) {
  const editor = usePptxEditor();
  const show = useSlideshow();
  const { t } = useTranslation();
  // Re-render on every change, so each row's state is the engine's current answer.
  useEditorState((s) => s.revision);
  const selected = useEditorState((s) => s.selectedSlides, sameIndices);
  const count = useEditorState((s) => s.slides.length);
  const readOnly = useEditorState((s) => s.mode === 'view');
  const current = useEditorState((s) => s.slides[s.currentSlide]);
  const slides = useEditorState((s) => s.slides);
  const layouts = useEditorState((s) => s.layouts ?? NO_LAYOUTS);
  const many = selected.length > 1;
  /** A row bound to a command: disabled with the engine's reason when `can` refuses it. */
  const row = (
    command: EditorCommand | null,
    props: { label: string; icon: string; shortcut?: string; slot: string }
  ) => {
    const reason = commandRowReason(editor, command);
    return (
      <MenuRow
        {...props}
        disabled={!!reason}
        reason={reason ? translateKeyOrText(t, reason) : null}
        onSelect={() => {
          onClose();
          // A move step opens a collapsed section it lands in (`runSlideStep`).
          if (command?.type === 'moveSlides') runSlideStep(editor, command);
          else if (command) editor?.exec(command);
        }}
      />
    );
  };
  const stepTo = useSlideStep(selected, count);
  const up = stepCommand(selected, stepTo(-1));
  const down = stepCommand(selected, stepTo(1));
  const canPresent = count > 0 && !!editor;
  return (
    <ContextMenu
      at={menu.at}
      onClose={onClose}
      label={t('contextMenu.railLabel')}
      keyboard={menu.keyboard}
      owner={menu.owner}
    >
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
      {(() => {
        const command = addSectionCommand(menu.index, t);
        const can = editor ? editor.can(command) : null;
        return (
          <MenuRow
            label={t('sections.add')}
            icon="view_agenda"
            slot="section.add"
            disabled={!can?.ok}
            reason={can && !can.ok ? translateKeyOrText(t, can.reason) : null}
            onSelect={() => {
              onClose();
              if (!editor) return;
              const added = addSectionAt(editor, command, menu.index);
              if (added) onSectionAdded?.(added);
            }}
          />
        );
      })()}
      <MenuSeparator />
      {(() => {
        // The engine's answer for a listed layout tells whether any layout can be picked: the
        // slide's own one when the deck lists it, else the first one.
        const reason = commandRowReason(
          editor,
          layoutRowProbe(layouts, current?.layoutId, selected)
        );
        return (
          <MenuSubmenu
            label={t('contextMenu.layout')}
            icon="dashboard"
            disabled={!!reason}
            reason={reason ? translateKeyOrText(t, reason) : null}
          >
            <LayoutGallery
              value={current?.layoutId ?? ''}
              onChange={(layoutId) => {
                onClose();
                editor?.exec({ type: 'applyLayout', indices: selected, layoutId });
              }}
            />
          </MenuSubmenu>
        );
      })()}
      {row(
        { type: 'resetSlides', indices: selected },
        { label: t('contextMenu.resetSlide'), icon: 'restart_alt', slot: 'slide.reset' }
      )}
      {onBackground ? (
        <MenuRow
          label={t('contextMenu.background')}
          icon="wallpaper"
          slot="slide.background"
          disabled={readOnly}
          reason={t('disabled.viewing')}
          onSelect={() => {
            onClose();
            onBackground();
          }}
        />
      ) : null}
      {(() => {
        // Checked when every selected slide is hidden; the row then shows them, else hides them.
        const hidden = selected.length > 0 && selected.every((i) => slides[i]?.hidden);
        const command: EditorCommand = {
          type: 'setSlidesHidden',
          indices: selected,
          hidden: !hidden,
        };
        const can = editor ? editor.can(command) : null;
        return (
          <MenuRow
            label={t('contextMenu.hideSlide')}
            icon="visibility_off"
            slot="slide.hide"
            checked={hidden}
            disabled={!can?.ok}
            reason={can && !can.ok ? translateKeyOrText(t, can.reason) : null}
            onSelect={() => {
              onClose();
              editor?.exec(command);
            }}
          />
        );
      })()}
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
          onClose();
          editor?.goToSlide(menu.index);
          show.start('current');
        }}
      />
    </ContextMenu>
  );
}
