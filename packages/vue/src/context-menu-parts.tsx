/** @jsxImportSource vue */
// `PptxEditorContextMenu`: the slide canvas's context menu as library parts, the Vue twin of
// the React adapter's `PptxEditor.ContextMenu` (`context-menu-parts.tsx`), after DOCX Editor's
// `DocxEditorContextMenu` (docx-editor/docs/site/content/customization/menus.mdx).
//
// The canvas opens a menu request (right-click, Shift+F10, the Menu key) for what is under the
// pointer: shapes, pictures, text being edited, or empty slide area. The menu renders core's
// `CANVAS_CONTEXT_MENUS` rows for that kind. `PptxEditorContextMenu.Slot slot-id` replaces or
// hides one row in place, other children add rows at the end, and `:preset="false"` keeps only
// the children. The `popups.contextMenu` registry entry decides what the canvas renders.

import {
  defineComponent,
  inject,
  provide,
  type ExtractPublicPropTypes,
  type InjectionKey,
  type PropType,
  type VNode,
} from 'vue';
import {
  CANVAS_CONTEXT_MENUS,
  arrangeMenuRows,
  canvasMenuLabelKey,
  warnMenuPart,
  type CanvasMenuKind,
} from '@pptx-editor.dev/core/editor';
import { ContextMenu, type ContextMenuPoint } from './ContextMenu.tsx';
import { useTranslation } from './composables.ts';
import { MenuRow, MenuSeparator, MenuSubmenu } from './menu-row.tsx';
import { MenuEntries, VUE_MENU_PARTS } from './menu-rows.tsx';
import {
  MenuCloseScope,
  MenuGroup,
  MenuSlotItem,
  flattenMenuChildren,
  slotRowOf,
} from './menu-parts.tsx';

/** What a canvas context menu opened on, and where. @public */
export interface ContextMenuRequest {
  /**
   * Shapes, pictures only (`picture`, the picture menu), text being edited, or empty slide
   * area. A selection of pictures only reports `picture`, not `shapes`.
   */
  readonly kind: CanvasMenuKind;
  /** Where the menu opens, in client pixels. */
  readonly at: ContextMenuPoint;
  /** Opened from the keyboard: focus goes to the first row. */
  readonly keyboard: boolean;
  /** The canvas element the menu opened on. */
  readonly owner: HTMLElement;
  /** Close the menu. */
  readonly close: () => void;
}

const requestKey: InjectionKey<ContextMenuRequest | null> = Symbol('pptxContextMenuRequest');

/** Provides an open menu request to `PptxEditorContextMenu`. @internal */
export const ContextMenuTargetProvider = defineComponent({
  name: 'PptxContextMenuTargetProvider',
  props: { value: { type: Object as PropType<ContextMenuRequest>, required: true as const } },
  setup(props, { slots }) {
    // The canvas mounts one provider per request, so the request is fixed for its life.
    provide(requestKey, props.value);
    return () => slots.default?.();
  },
});

/** What `useContextMenuRequest` returns: the open request, or null. @public */
export type UseContextMenuRequestReturn = ContextMenuRequest | null;

/**
 * The open canvas context menu, for a host's own rows: what it opened on, where, and how to
 * close it. Null outside an open menu. Each request mounts the menu anew, so the value is fixed
 * for the menu's life. @public
 */
export function useContextMenuRequest(): UseContextMenuRequestReturn {
  return inject(requestKey, null);
}

/** What `useContextMenuTarget` returns: the element the menu opened on, or null. @public */
export type UseContextMenuTargetReturn = HTMLElement | null;

/**
 * The element the open context menu opened on (the slide canvas), or null while no menu is
 * open. The same composable as DOCX Editor's. @public
 */
export function useContextMenuTarget(): UseContextMenuTargetReturn {
  return inject(requestKey, null)?.owner ?? null;
}

const pptxEditorContextMenuProps = {
  /**
   * `true` (default): the built-in rows for the target, with `PptxEditorContextMenu.Slot`
   * children replacing their rows in place and other children added at the end. `false`: only
   * the children.
   */
  preset: { type: Boolean, default: true },
  /** A class name added to the menu, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/**
 * Props for `PptxEditorContextMenu`, the twin of React's `PptxEditorContextMenuProps`.
 * `className` (or a host `class`) lands on the menu; the default slot holds the host's rows. @public
 */
export type PptxEditorContextMenuProps = ExtractPublicPropTypes<typeof pptxEditorContextMenuProps>;

/**
 * The slide canvas's context menu. It renders while the canvas has a menu request open: put it
 * in the `popups.contextMenu` entry to change the rows.
 */
const ContextMenuRoot = defineComponent({
  name: 'PptxEditorContextMenu',
  inheritAttrs: false,
  props: pptxEditorContextMenuProps,
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    const target = useContextMenuRequest();
    return () => {
      if (!target) {
        warnMenuPart('PptxEditorContextMenu', 'outsideContextMenu', undefined, VUE_MENU_PARTS);
        return null;
      }
      const content = (): unknown => {
        const children = flattenMenuChildren(slots.default?.());
        if (!props.preset) return children;
        // A hidden row for a row this target lacks (a text row on a shape) is no extra row.
        const { entries, overrides, extras } = arrangeMenuRows(
          CANVAS_CONTEXT_MENUS[target.kind],
          children,
          slotRowOf
        );
        const renders = new Map<string, (() => VNode) | null>();
        for (const [key, child] of overrides) renders.set(key, child ? () => child : null);
        return [
          <MenuEntries entries={entries} overrides={renders} />,
          extras.length ? <MenuSeparator /> : null,
          ...extras,
        ];
      };
      return (
        <ContextMenu
          at={target.at}
          onClose={target.close}
          label={t(canvasMenuLabelKey(target.kind))}
          keyboard={target.keyboard}
          owner={target.owner}
          class={attrs.class}
          className={props.className}
        >
          {{
            default: () => (
              <MenuCloseScope close={target.close}>{{ default: content }}</MenuCloseScope>
            ),
          }}
        </ContextMenu>
      );
    };
  },
});

/** The canvas context menu with its parts as statics, the twin of React's. @public */
export type ContextMenuNamespace = typeof ContextMenuRoot & {
  /** A registry row: replace, change, or hide it in place. */
  readonly Slot: typeof MenuSlotItem;
  /** A host row. */
  readonly Item: typeof MenuRow;
  /** A labeled set of rows. */
  readonly Group: typeof MenuGroup;
  /** A line between rows. */
  readonly Separator: typeof MenuSeparator;
  /** A row that opens a nested panel. */
  readonly Submenu: typeof MenuSubmenu;
};

/**
 * The slide canvas's context menu (`PptxEditorContextMenu`), with DOCX Editor's parts (`Slot`,
 * `Item`, `Group`, `Separator`, `Submenu`). Render it in `popups.contextMenu`. @public
 *
 * @example
 * ```ts
 * const popups: PptxEditorPopups = {
 *   contextMenu: (props) =>
 *     h(PptxEditorContextMenu, props, () => [
 *       h(PptxEditorContextMenu.Slot, { slotId: 'format.painter', hidden: true }),
 *       h(PptxEditorContextMenu.Item, { label: 'Ask legal', icon: 'info', onSelect: askLegal }),
 *     ]),
 * };
 * ```
 */
export const PptxEditorContextMenu: ContextMenuNamespace = Object.assign(ContextMenuRoot, {
  Slot: MenuSlotItem,
  Item: MenuRow,
  Group: MenuGroup,
  Separator: MenuSeparator,
  Submenu: MenuSubmenu,
});
