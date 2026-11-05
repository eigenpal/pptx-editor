/** @jsxImportSource vue */
// Sections in the slide rail and the slide sorter (`p14:sectionLst`): the Vue twin of the React
// adapter's `RailSections.tsx`, with the same markup, classes, ARIA, and keys. A header per
// section collapses its slides (the arrow), selects them (a click on the name), renames in
// place (double-click, F2, or the menu), and opens the section menu. Which sections are
// collapsed is the engine's view setting (`collapsedSections`): the rail and the sorter share
// it, and it is never saved. The keys, the focus moves, and the menu rows come from core
// (`pane-support.ts`), so the two adapters act the same.
//
// In the rail and the sorter the sections are a tree (`tree` > `treeitem` > `group` >
// `treeitem`), so the headers are in the keyboard order.

import {
  defineComponent,
  getCurrentScope,
  nextTick,
  onMounted,
  onScopeDispose,
  ref,
  shallowRef,
  toValue,
  watch,
  type MaybeRefOrGetter,
  type PropType,
  type Ref,
  type ShallowRef,
  type VNodeChild,
} from 'vue';
import {
  canRenameSectionInPlace,
  focusCurrentSlideIn,
  openSectionOfCurrent,
  refocusAfterCollapse,
  registerSectionRenamer,
  registerSlideSorter,
  requestSectionRenameIn,
  sectionHeaderKeyDown,
  sectionMenuPoint,
  selectSectionByName,
  subscribeSectionRenamers,
  toggleSlideSection,
  type PptxEditorInstance,
  type SlideSection,
} from '@pptx-editor.dev/core/editor';
import { Icon, keepFocus } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { usePopupEntry } from './popups.ts';
import {
  SectionContextMenu,
  type SectionContextMenuProps,
  type SectionMenuState,
} from './RailContextMenu.tsx';
import type { ContextMenuPoint } from './ContextMenu.tsx';

const EMPTY: readonly SlideSection[] = [];
const NO_IDS: readonly string[] = [];

/** The deck's sections, reference-stable while they do not change. @internal */
export function useSlideSections(): Readonly<ShallowRef<readonly SlideSection[]>> {
  return useEditorState((s) => s.sections ?? EMPTY);
}

/** The collapsed section ids, reference-stable while they do not change. @internal */
export function useCollapsedSections(): Readonly<ShallowRef<readonly string[]>> {
  return useEditorState((s) => s.collapsedSections ?? NO_IDS);
}

/**
 * When the current slide changes, its section opens, so the current slide is in view. A
 * section the user collapses stays collapsed until then, and a click on the name of a
 * collapsed section selects its slides and keeps it collapsed. @internal
 */
export function useOpenCurrentSection(
  grouped: MaybeRefOrGetter<readonly SlideSection[] | null>
): void {
  const editor = usePptxEditor();
  const current = useEditorState((s) => s.currentSlide);
  watch(current, (now, was) => {
    if (now === was || !editor.value) return;
    openSectionOfCurrent(editor.value, toValue(grouped), now);
  });
}

/**
 * Returns a function that moves focus to the current slide of `list` after the next render:
 * after a section menu row removes sections, the section that had focus is gone only then. A
 * current slide in a collapsed section has no tile, so its section header takes focus.
 * @internal
 */
export function useFocusCurrentAfterRender(list: Ref<HTMLElement | null>): () => void {
  return () => void nextTick(() => focusCurrentSlideIn(list.value));
}

/**
 * After sections collapse (the arrow, Collapse all from a menu or the toolbar), a slide that had
 * focus can be gone: focus moves to the current slide, or to the header of its section when
 * that section is shut. Focus that went elsewhere on purpose stays there. Put the returned
 * handler on the list's `focusin`. @internal
 */
export function useFocusAfterCollapse(
  list: Ref<HTMLElement | null>,
  collapsed: Readonly<ShallowRef<readonly string[]>>
): (e: FocusEvent) => void {
  let focused: HTMLElement | null = null;
  watch(collapsed, () => refocusAfterCollapse(list.value, focused), { flush: 'post' });
  return (e) => {
    focused = e.target as HTMLElement;
  };
}

/** Count a mounted slide sorter of the editor, for the Slide Sorter view's warning. @internal */
export function useSorterRegistration(): void {
  const editor = usePptxEditor();
  watch(
    editor,
    (next, _old, onCleanup) => {
      if (next) onCleanup(registerSlideSorter(next));
    },
    { immediate: true }
  );
}

/**
 * Ask the slide rail or the slide sorter of `editor` to start renaming section `sectionId` in
 * place, as Home > Section > Rename Section does: what the `section.rename` slot runs. Use it
 * when you build your own section chrome. When neither is mounted, nothing opens and a
 * development warning says so. The twin of React's `requestSectionRename`. @public
 */
export function requestSectionRename(editor: PptxEditorInstance | null, sectionId: string): void {
  requestSectionRenameIn(editor, sectionId);
}

/**
 * Start renaming when `requestSectionRename` asks for this editor, while `enabled`: a closed
 * rail has no header to rename in, so it does not count. A flat rail or sorter counts: the
 * section that Add Section creates renders its header right after the request. @internal
 */
export function useSectionRenameRequests(
  onRequest: (sectionId: string) => void,
  enabled: MaybeRefOrGetter<boolean> = true
): void {
  const editor = usePptxEditor();
  watch(
    () => [editor.value, toValue(enabled)] as const,
    ([next, on], _old, onCleanup) => {
      if (next && on) onCleanup(registerSectionRenamer(next, onRequest));
    },
    { immediate: true }
  );
}

/** Whether a mounted rail or sorter of this editor renames a section in place. @internal */
export function useSectionRenameAvailable(): Readonly<Ref<boolean>> {
  const editor = usePptxEditor();
  const available = ref(canRenameSectionInPlace(editor.value));
  const read = (): void => {
    available.value = canRenameSectionInPlace(editor.value);
  };
  if (!getCurrentScope()) return available;
  watch(editor, read);
  onScopeDispose(subscribeSectionRenamers(read));
  return available;
}

/** A section's header row: the arrow, the name (or its rename field), and the slide count. */
const SectionHead = defineComponent({
  name: 'PptxSectionHead',
  props: {
    section: { type: Object as PropType<SlideSection>, required: true },
    collapsed: { type: Boolean, required: true },
    renaming: { type: Boolean, required: true },
    label: { type: String, required: true },
    onToggle: { type: Function as PropType<() => void>, required: true },
    onSelect: { type: Function as PropType<() => void>, required: true },
    onRename: { type: Function as PropType<() => void>, required: true },
    /** `keyboard`: Enter or Escape ended the rename, not a blur. */
    onRenameEnd: { type: Function as PropType<(keyboard: boolean) => void>, required: true },
    onMenu: {
      type: Function as PropType<(owner: HTMLElement, at: ContextMenuPoint | null) => void>,
      required: true,
    },
  },
  setup(props) {
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const name = ref(props.section.name);
    const field = ref<HTMLInputElement | null>(null);
    // Escape cancels: the field's blur that follows must not write the name.
    let cancelled = false;
    const start = (): void => {
      cancelled = false;
      // The name when the rename starts; a rename from another replica does not reset the field.
      name.value = props.section.name;
      // The field selects its whole name once it shows that name: a rename that Escape
      // cancelled leaves the typed text in the field until the next render.
      void nextTick(() => {
        field.value?.focus();
        field.value?.select();
      });
    };
    onMounted(() => {
      if (props.renaming) start();
    });
    watch(
      () => props.renaming,
      (on) => {
        if (on) start();
      },
      { flush: 'post' }
    );
    const commitName = (keyboard: boolean): void => {
      if (cancelled) return;
      cancelled = true;
      const next = name.value.trim();
      if (next && next !== props.section.name)
        editor.value?.exec({ type: 'renameSection', sectionId: props.section.id, name: next });
      props.onRenameEnd(keyboard);
    };
    const onFieldKey = (e: KeyboardEvent): void => {
      if (e.key === 'Enter') commitName(true);
      else if (e.key === 'Escape') {
        cancelled = true;
        props.onRenameEnd(true);
      } else {
        // Editing keys stay in the field: they never act on the rail.
        e.stopPropagation();
        return;
      }
      e.preventDefault();
      e.stopPropagation();
    };
    return () => (
      <div
        class="pptx-rail__section-head"
        data-collapsed={props.collapsed ? '' : undefined}
        // No tooltip while renaming: it would cover the slides under the field.
        data-pptx-tooltip={props.renaming ? undefined : props.label}
        onContextmenu={(e: MouseEvent) => {
          e.preventDefault();
          e.stopPropagation();
          props.onMenu(e.currentTarget as HTMLElement, { x: e.clientX, y: e.clientY });
        }}
      >
        <span
          class="pptx-rail__section-toggle"
          aria-hidden="true"
          onMousedown={keepFocus}
          onClick={(e: MouseEvent) => {
            e.stopPropagation();
            props.onToggle();
          }}
        >
          <Icon name={props.collapsed ? 'chevron_right' : 'expand_more'} />
        </span>
        {props.renaming ? (
          <input
            ref={field}
            class="pptx-rail__section-name-field"
            aria-label={t('sections.nameLabel')}
            value={name.value}
            maxlength={255}
            onInput={(e: Event) => {
              name.value = (e.target as HTMLInputElement).value;
            }}
            onBlur={() => commitName(false)}
            onKeydown={onFieldKey}
          />
        ) : (
          <span
            class="pptx-rail__section-name"
            onMousedown={keepFocus}
            onClick={() => props.onSelect()}
            onDblclick={() => props.onRename()}
          >
            {props.section.name}
          </span>
        )}
        <span class="pptx-rail__section-count" aria-hidden="true">
          {props.section.slideCount}
        </span>
      </div>
    );
  },
});

/**
 * One section: its header and its slides, a `treeitem` that takes focus and the section keys
 * (Right and Left expand and collapse, Enter selects its slides, F2 renames, Shift+F10 opens
 * the menu). Up and Down go to the rail or the sorter. The default slot is its slides.
 * @internal
 */
export const SectionNode = defineComponent({
  name: 'PptxSectionNode',
  props: {
    section: { type: Object as PropType<SlideSection>, required: true },
    collapsed: { type: Boolean, required: true },
    renaming: { type: Boolean, required: true },
    /** The slide rail (a list of thumbnails) or the slide sorter (a grid of tiles). */
    variant: { type: String as PropType<'rail' | 'sorter'>, required: true },
    focusable: { type: Boolean, required: true },
    onRename: { type: Function as PropType<() => void>, required: true },
    onRenameEnd: { type: Function as PropType<() => void>, required: true },
    /** A menu row removed sections: the rail or the sorter moves focus after its render. */
    onRemoved: { type: Function as PropType<() => void>, required: true },
  },
  setup(props, { slots }) {
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const readOnly = useEditorState((s) => s.mode === 'view');
    const menu = shallowRef<SectionMenuState | null>(null);
    const node = ref<HTMLLIElement | null>(null);
    const menuEntry = usePopupEntry('sectionContextMenu');
    // Move up and Move down reorder the headers: the browser drops focus from a moved element,
    // so the header that had it takes it back after the move.
    let hadFocus = false;
    watch(
      () => props.section.firstSlide,
      () => {
        const el = node.value;
        const active = el?.ownerDocument.activeElement;
        if (hadFocus && el?.isConnected && (!active || active === el.ownerDocument.body))
          el.focus({ preventScroll: true });
      },
      { flush: 'post' }
    );
    const toggle = (): void =>
      toggleSlideSection(editor.value, props.section, props.collapsed, node.value);
    const openMenu = (owner: HTMLElement, at: ContextMenuPoint | null): void => {
      menu.value = { at: at ?? sectionMenuPoint(owner), keyboard: !at, owner };
    };
    const onKeydown = (e: KeyboardEvent): void => {
      if (!node.value) return;
      sectionHeaderKeyDown(e, node.value, {
        editor: editor.value,
        readOnly: readOnly.value,
        section: props.section,
        collapsed: props.collapsed,
        toggle,
        rename: () => props.onRename(),
        openMenu: (owner) => openMenu(owner, null),
      });
    };
    const menuPart = (): VNodeChild => {
      const entry = menuEntry();
      if (!menu.value || entry === false) return null;
      const menuProps: SectionContextMenuProps = {
        section: props.section,
        menu: menu.value,
        onClose: () => {
          menu.value = null;
        },
        onRename: () => props.onRename(),
        onRemoved: () => props.onRemoved(),
      };
      return typeof entry === 'function' ? entry(menuProps) : <SectionContextMenu {...menuProps} />;
    };
    return () => {
      const { section, collapsed } = props;
      const label = t(section.slideCount === 1 ? 'sections.headerOne' : 'sections.header', {
        name: section.name,
        count: section.slideCount,
      });
      const sorter = props.variant === 'sorter';
      return (
        <li
          ref={node}
          onFocus={() => {
            hadFocus = true;
          }}
          onBlur={(e: FocusEvent) => {
            // Focus that leaves is gone on purpose: to another element, to nothing (a click on
            // an empty part of the page), or out of the window. Two blurs keep it: one into the
            // section's menu, which gives focus back after its row runs, and one of a
            // disconnected element, which is the move of this header.
            // Vue moves a keyed header while it is still connected, and the browser blurs it
            // then: the blur counts only when focus is still away after the render that
            // moved it.
            const to = e.relatedTarget as HTMLElement | null;
            if (to?.closest?.('[role="menu"]')) return;
            setTimeout(() => {
              const el = node.value;
              if (el?.isConnected && el.ownerDocument.activeElement !== el) hadFocus = false;
            }, 0);
          }}
          role="treeitem"
          aria-expanded={!collapsed}
          aria-label={label}
          tabindex={props.focusable ? 0 : -1}
          class={sorter ? 'pptx-sorter__section' : 'pptx-rail__section'}
          data-section={section.id}
          onKeydown={onKeydown}
          onContextmenu={(e: MouseEvent) => {
            if (e.target !== e.currentTarget) return;
            e.preventDefault();
            openMenu(e.currentTarget as HTMLElement, null);
          }}
        >
          <SectionHead
            section={section}
            collapsed={collapsed}
            renaming={props.renaming}
            label={label}
            onToggle={toggle}
            // Focus goes to the current slide of the selection, so Delete, Ctrl+D, and the
            // other slide keys act on the slides that the click selected. A shut section keeps
            // focus.
            onSelect={() => selectSectionByName(editor.value, section, collapsed, node.value)}
            onRename={() => {
              if (!readOnly.value) props.onRename();
            }}
            onRenameEnd={(keyboard: boolean) => {
              props.onRenameEnd();
              // After Enter or Escape, focus goes back to the section. A blur keeps the focus
              // where the user put it.
              if (keyboard) node.value?.focus({ preventScroll: true });
            }}
            onMenu={openMenu}
          />
          {collapsed ? null : (
            <ol role="group" class={sorter ? 'pptx-sorter__grid' : 'pptx-rail__section-list'}>
              {slots.default?.()}
            </ol>
          )}
          {menuPart()}
        </li>
      );
    };
  },
});
