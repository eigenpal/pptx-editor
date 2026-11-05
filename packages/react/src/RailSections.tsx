/** @jsxImportSource react */
// Sections in the slide rail and the slide sorter (`p14:sectionLst`): a header per
// section that collapses its slides (the arrow), selects them (a click on the name), renames in
// place (double-click, F2, or the menu), and opens the section menu (Rename, Remove, Remove
// with slides, Remove all, Move up and down, Collapse and Expand all). Which sections are
// collapsed is the engine's view setting (`collapsedSections`): the rail and the sorter share
// it, and it is never saved.
//
// In the rail and the sorter the sections are a tree (`tree` > `treeitem` > `group` >
// `treeitem`), so the headers are in the keyboard order.

import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  useSyncExternalStore,
  type KeyboardEvent,
  type ReactNode,
  type RefObject,
} from 'react';
import {
  canRenameSectionInPlace,
  focusCurrentSlideIn,
  openSectionOfCurrent,
  commandRowReason,
  refocusAfterCollapse,
  registerSectionRenamer,
  registerSlideSorter,
  requestSectionRenameIn,
  sectionHeaderKeyDown,
  sectionMenuPoint,
  selectSectionByName,
  subscribeSectionRenamers,
  toggleSlideSection,
  translateKeyOrText,
  type EditorCommand,
  type PptxEditorInstance,
  type SlideSection,
} from '@pptx-editor.dev/core/editor';
import { Icon, keepFocus } from './chrome.tsx';
import { ContextMenu, MenuRow, MenuSeparator, type ContextMenuPoint } from './ContextMenu.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { usePopupEntry } from './popups.tsx';

const EMPTY: readonly SlideSection[] = [];
const NO_IDS: readonly string[] = [];

/** The deck's sections, reference-stable while they do not change. */
export function useSlideSections(): readonly SlideSection[] {
  return useEditorState((s) => s.sections ?? EMPTY);
}

/** The collapsed section ids, reference-stable while they do not change. */
export function useCollapsedSections(): readonly string[] {
  return useEditorState((s) => s.collapsedSections ?? NO_IDS);
}

/**
 * When the current slide changes, its section opens, so the current slide is in view. A
 * section the user collapses stays collapsed until then, and a click on the name of a
 * collapsed section selects its slides and keeps it collapsed.
 */
export function useOpenCurrentSection(grouped: readonly SlideSection[] | null): void {
  const editor = usePptxEditor();
  const current = useEditorState((s) => s.currentSlide);
  const was = useRef(current);
  useEffect(() => {
    if (was.current === current || !editor) return;
    was.current = current;
    openSectionOfCurrent(editor, grouped, current);
  }, [current, grouped, editor]);
}

/**
 * Returns a function that moves focus to the current slide of `list` after the next render:
 * after a section menu row removes sections, the section that had focus is gone only then. A
 * current slide in a collapsed section has no tile, so its section header takes focus.
 */
export function useFocusCurrentAfterRender(list: RefObject<HTMLElement | null>): () => void {
  const [request, setRequest] = useState(0);
  useEffect(() => {
    if (request) focusCurrentSlideIn(list.current);
  }, [request, list]);
  return useCallback(() => setRequest((n) => n + 1), []);
}

/**
 * After sections collapse (the arrow, Collapse all from a menu or the toolbar), a slide that had
 * focus can be gone: focus moves to the current slide, or to the header of its section when
 * that section is shut. Focus that went elsewhere on purpose stays there.
 */
export function useFocusAfterCollapse(
  list: RefObject<HTMLElement | null>,
  collapsed: readonly string[]
): void {
  const focused = useRef<HTMLElement | null>(null);
  useEffect(() => {
    const root = list.current;
    if (!root) return;
    const onFocus = (e: FocusEvent): void => {
      focused.current = e.target as HTMLElement;
    };
    root.addEventListener('focusin', onFocus);
    return () => root.removeEventListener('focusin', onFocus);
    // Each render: the list element can mount after the first one (a rail that opens later).
  });
  useEffect(() => {
    refocusAfterCollapse(list.current, focused.current);
  }, [collapsed, list]);
}

/** Count a mounted slide sorter of the editor, for the Slide Sorter view's warning. */
export function useSorterRegistration(): void {
  const editor = usePptxEditor();
  useEffect(() => (editor ? registerSlideSorter(editor) : undefined), [editor]);
}

/**
 * Ask the slide rail or the slide sorter of `editor` to start renaming section `sectionId` in
 * place, as Home > Section > Rename Section does: what the `section.rename` slot runs. Use it
 * when you build your own section chrome. When neither is mounted, nothing opens and a
 * development warning says so. @public
 */
export function requestSectionRename(editor: PptxEditorInstance | null, sectionId: string): void {
  requestSectionRenameIn(editor, sectionId);
}

/**
 * Start renaming when `requestSectionRename` asks for this editor, while `enabled`: a closed
 * rail has no header to rename in, so it does not count. A flat rail or sorter counts: the
 * section that Add Section creates renders its header right after the request.
 */
export function useSectionRenameRequests(
  onRequest: (sectionId: string) => void,
  enabled = true
): void {
  const editor = usePptxEditor();
  useEffect(() => {
    if (!editor || !enabled) return;
    return registerSectionRenamer(editor, onRequest);
  }, [editor, onRequest, enabled]);
}

/** Whether a mounted rail or sorter of this editor renames a section in place. */
export function useSectionRenameAvailable(): boolean {
  const editor = usePptxEditor();
  return useSyncExternalStore(
    subscribeSectionRenamers,
    () => canRenameSectionInPlace(editor),
    () => false
  );
}

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

/**
 * Props of the section menu, and of a `popups.sectionContextMenu` renderer: the menu of a
 * section header in the slide rail and in the slide sorter. @public
 */
export interface SectionContextMenuProps {
  /** The section whose header opened the menu. */
  readonly section: SlideSection;
  /** Where the menu opens, and how. */
  readonly menu: SectionMenuState;
  /** Closes the menu. Call it when a row runs or the user dismisses the menu. */
  readonly onClose: () => void;
  /** Starts renaming the section in place, in the rail or the sorter that opened the menu. */
  readonly onRename: () => void;
  /** Call after a row removes sections: focus goes to the current slide after the render. */
  readonly onRemoved: () => void;
}

/**
 * One section: its header and its slides, a `treeitem` that takes focus and the section keys
 * (Right and Left expand and collapse, Enter selects its slides, F2 renames, Shift+F10 opens
 * the menu). Up and Down go to the rail or the sorter.
 */
export function SectionNode({
  section,
  collapsed,
  renaming,
  variant,
  focusable,
  onRename,
  onRenameEnd,
  onRemoved,
  children,
}: {
  section: SlideSection;
  collapsed: boolean;
  renaming: boolean;
  /** The slide rail (a list of thumbnails) or the slide sorter (a grid of tiles). */
  variant: 'rail' | 'sorter';
  focusable: boolean;
  onRename: () => void;
  onRenameEnd: () => void;
  /** A menu row removed sections: the rail or the sorter moves focus after its render. */
  onRemoved: () => void;
  children: ReactNode;
}) {
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const readOnly = useEditorState((s) => s.mode === 'view');
  const [menu, setMenu] = useState<SectionMenuState | null>(null);
  const node = useRef<HTMLLIElement>(null);
  const menuEntry = usePopupEntry('sectionContextMenu');
  // Move up and Move down reorder the headers: the browser drops focus from a moved element,
  // so the header that had it takes it back after the move.
  const hadFocus = useRef(false);
  useLayoutEffect(() => {
    const el = node.current;
    const active = el?.ownerDocument.activeElement;
    if (hadFocus.current && el?.isConnected && (!active || active === el.ownerDocument.body))
      el.focus({ preventScroll: true });
  }, [section.firstSlide]);
  const label = t(section.slideCount === 1 ? 'sections.headerOne' : 'sections.header', {
    name: section.name,
    count: section.slideCount,
  });
  // A slide of the section that has focus goes away when it collapses: its header takes focus.
  const toggle = (): void => toggleSlideSection(editor, section, collapsed, node.current);
  const openMenu = (owner: HTMLElement, at: ContextMenuPoint | null): void =>
    setMenu({ at: at ?? sectionMenuPoint(owner), keyboard: !at, owner });
  const onKeyDown = (e: KeyboardEvent<HTMLLIElement>): void =>
    sectionHeaderKeyDown(e, e.currentTarget, {
      editor,
      readOnly,
      section,
      collapsed,
      toggle,
      rename: onRename,
      openMenu: (owner) => openMenu(owner, null),
    });
  const head = (
    <SectionHead
      section={section}
      collapsed={collapsed}
      renaming={renaming}
      label={label}
      onToggle={toggle}
      // Focus goes to the current slide of the selection, so Delete, Ctrl+D, and the other
      // slide keys act on the slides that the click selected. A shut section keeps focus.
      onSelect={() => selectSectionByName(editor, section, collapsed, node.current)}
      onRename={() => !readOnly && onRename()}
      onRenameEnd={(keyboard) => {
        onRenameEnd();
        // After Enter or Escape, focus goes back to the section. A blur keeps the focus where
        // the user put it.
        if (keyboard) node.current?.focus({ preventScroll: true });
      }}
      onMenu={(owner, at) => openMenu(owner, at)}
    />
  );
  const menuProps: SectionContextMenuProps | null = menu
    ? { section, menu, onClose: () => setMenu(null), onRename, onRemoved }
    : null;
  const menuPart =
    !menuProps || menuEntry === false ? null : typeof menuEntry === 'function' ? (
      menuEntry(menuProps)
    ) : (
      <SectionContextMenu {...menuProps} />
    );
  const sorter = variant === 'sorter';
  return (
    <li
      ref={node}
      onFocus={(e) => {
        if (e.target === e.currentTarget) hadFocus.current = true;
      }}
      onBlur={(e) => {
        // Focus that leaves is gone on purpose: to another element, to nothing (a click on an
        // empty part of the page), or out of the window. Two blurs keep it: one into the
        // section's menu, which gives focus back after its row runs, and one of a disconnected
        // element, which is the move of this header.
        const to = e.relatedTarget as HTMLElement | null;
        if (
          e.target === e.currentTarget &&
          e.currentTarget.isConnected &&
          !to?.closest('[role="menu"]')
        )
          hadFocus.current = false;
      }}
      role="treeitem"
      aria-expanded={!collapsed}
      aria-label={label}
      tabIndex={focusable ? 0 : -1}
      className={sorter ? 'pptx-sorter__section' : 'pptx-rail__section'}
      data-section={section.id}
      onKeyDown={onKeyDown}
      onContextMenu={(e) => {
        if (e.target !== e.currentTarget) return;
        e.preventDefault();
        openMenu(e.currentTarget, null);
      }}
    >
      {head}
      {collapsed ? null : (
        <ol role="group" className={sorter ? 'pptx-sorter__grid' : 'pptx-rail__section-list'}>
          {children}
        </ol>
      )}
      {menuPart}
    </li>
  );
}

/** A section's header row: the arrow, the name (or its rename field), and the slide count. */
function SectionHead({
  section,
  collapsed,
  renaming,
  label,
  onToggle,
  onSelect,
  onRename,
  onRenameEnd,
  onMenu,
}: {
  section: SlideSection;
  collapsed: boolean;
  renaming: boolean;
  label: string;
  onToggle: () => void;
  onSelect: () => void;
  onRename: () => void;
  /** `keyboard`: Enter or Escape ended the rename, not a blur. */
  onRenameEnd: (keyboard: boolean) => void;
  onMenu: (owner: HTMLElement, at: ContextMenuPoint | null) => void;
}) {
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const [name, setName] = useState(section.name);
  const field = useRef<HTMLInputElement>(null);
  // Escape cancels: the field's blur that follows must not write the name.
  const cancelled = useRef(false);
  // The name when the rename starts; a rename from another replica does not reset the field.
  const current = useRef(section.name);
  current.current = section.name;
  // The field selects its whole name once it shows the name it starts from: a rename that
  // Escape cancelled leaves the typed text in the state until then.
  const [selectAll, setSelectAll] = useState(false);
  useEffect(() => {
    if (!renaming) return;
    cancelled.current = false;
    setName(current.current);
    setSelectAll(true);
  }, [renaming]);
  useLayoutEffect(() => {
    if (!selectAll || !renaming) return;
    setSelectAll(false);
    field.current?.focus();
    field.current?.select();
  }, [selectAll, renaming, name]);
  const commitName = (keyboard: boolean): void => {
    if (cancelled.current) return;
    cancelled.current = true;
    const next = name.trim();
    if (next && next !== section.name)
      editor?.exec({ type: 'renameSection', sectionId: section.id, name: next });
    onRenameEnd(keyboard);
  };
  return (
    <div
      className="pptx-rail__section-head"
      data-collapsed={collapsed ? '' : undefined}
      // No tooltip while renaming: it would cover the slides under the field.
      data-pptx-tooltip={renaming ? undefined : label}
      onContextMenu={(e) => {
        e.preventDefault();
        e.stopPropagation();
        onMenu(e.currentTarget, { x: e.clientX, y: e.clientY });
      }}
    >
      <span
        className="pptx-rail__section-toggle"
        aria-hidden="true"
        onMouseDown={keepFocus}
        onClick={(e) => {
          e.stopPropagation();
          onToggle();
        }}
      >
        <Icon name={collapsed ? 'chevron_right' : 'expand_more'} />
      </span>
      {renaming ? (
        <input
          ref={field}
          className="pptx-rail__section-name-field"
          aria-label={t('sections.nameLabel')}
          value={name}
          maxLength={255}
          onChange={(e) => setName(e.target.value)}
          onBlur={() => commitName(false)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') commitName(true);
            else if (e.key === 'Escape') {
              cancelled.current = true;
              onRenameEnd(true);
            } else {
              // Editing keys stay in the field: they never act on the rail.
              e.stopPropagation();
              return;
            }
            e.preventDefault();
            e.stopPropagation();
          }}
        />
      ) : (
        <span
          className="pptx-rail__section-name"
          onMouseDown={keepFocus}
          onClick={onSelect}
          onDoubleClick={onRename}
        >
          {section.name}
        </span>
      )}
      <span className="pptx-rail__section-count" aria-hidden="true">
        {section.slideCount}
      </span>
    </div>
  );
}

/**
 * The section menu: Rename, Remove, Remove with slides, Remove all, Move up
 * and down, Collapse and Expand all. Every row asks the engine whether it can run. @public
 */
export function SectionContextMenu({
  section,
  menu,
  onClose,
  onRename,
  onRemoved,
}: SectionContextMenuProps) {
  const editor = usePptxEditor();
  const { t } = useTranslation();
  useEditorState((s) => s.revision);
  const id = section.id;
  const row = (
    command: EditorCommand,
    label: string,
    icon: string,
    slot: string,
    then?: () => void
  ) => {
    const reason = commandRowReason(editor, command);
    return (
      <MenuRow
        label={label}
        icon={icon}
        slot={slot}
        disabled={!!reason}
        reason={reason ? translateKeyOrText(t, reason) : null}
        onSelect={() => {
          onClose();
          if (editor?.exec(command).ok) then?.();
        }}
      />
    );
  };
  // Rename runs no command until the name is typed; it needs editing mode. The probe uses a
  // fixed valid name: the section's own name can come from a file and be invalid.
  const rename = editor?.can({ type: 'renameSection', sectionId: id, name: 'x' });
  return (
    <ContextMenu
      at={menu.at}
      onClose={onClose}
      label={t('sections.menuLabel')}
      keyboard={menu.keyboard}
      owner={menu.owner}
    >
      <MenuRow
        label={t('sections.rename')}
        icon="edit_note"
        slot="section.rename"
        disabled={!rename?.ok}
        reason={rename && !rename.ok ? translateKeyOrText(t, rename.reason) : null}
        onSelect={() => {
          onClose();
          onRename();
        }}
      />
      {row(
        { type: 'removeSection', sectionId: id },
        t('sections.remove'),
        'remove',
        'section.remove',
        onRemoved
      )}
      {row(
        { type: 'removeSection', sectionId: id, withSlides: true },
        t('sections.removeWithSlides'),
        'delete',
        'section.removeWithSlides',
        onRemoved
      )}
      {row(
        { type: 'removeAllSections' },
        t('sections.removeAll'),
        'delete',
        'section.removeAll',
        onRemoved
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
    </ContextMenu>
  );
}
