/** @jsxImportSource vue */
// `PptxEditorMenu`: the menu bar as library parts, the Vue twin of the React adapter's
// `PptxEditor.Menu` (`menu-parts.tsx`), after DOCX Editor's `DocxEditorMenu`
// (docx-editor/docs/site/content/customization/menus.mdx, packages/vue/src/editor/menu/).
//
// The preset bar renders core's `CHROME_MENUS` (File, Edit, View, Insert, Format, Arrange,
// Slide Show). A named menu child (`PptxEditorMenu.File`) changes its menu in place: `hidden`
// removes it, `MenuSlotItem slot-id` replaces or hides one row, other children add rows at its
// end, and `:preset="false"` keeps only its children. `MenuEntry` with another id adds a host
// menu. `:preset="false"` on the bar renders only the children. Hiding a row changes the
// interface only: its command still runs from the toolbar, the keyboard, and the API. Which
// rows and menus show, and how the bar's keys move, come from core (`menu-arrangement.ts`).

import {
  Comment,
  Fragment,
  Text,
  defineComponent,
  inject,
  isVNode,
  normalizeClass,
  provide,
  ref,
  watch,
  type DefineComponent,
  type ExtractPublicPropTypes,
  type PropType,
  type VNode,
  type VNodeArrayChildren,
  type VNodeChild,
} from 'vue';
import {
  arrangeMenuBar,
  arrangeMenuRows,
  menuBarArrow,
  menuBarClosedOnEscape,
  menuBarRestoresFocus,
  menuBarTriggerKey,
  menuPartName,
  translateKeyOrText,
  warnMenuPart,
  type ChromeMenu,
  type ChromeMenuId,
  type ChromeSlotId,
  type MenuSlotRowSpec,
} from '@pptx-editor.dev/core/editor';
import { Icon, Popover, keepFocus } from './chrome.tsx';
import { useTranslation } from './composables.ts';
import { MenuRow, MenuSeparator, MenuSubmenu, menuCloseKey } from './menu-row.tsx';
import { MenuEntries, SlotMenuRow, VUE_MENU_PARTS, rowDefaultsKey } from './menu-rows.tsx';
import { MenuTableGrid } from './TableControls.tsx';

// ---------------------------------------------------------------------------------------------
// Reading host children

/** A vnode prop by its camelCase or kebab-case name (templates pass either). */
function prop(vnode: VNode, camel: string): unknown {
  const props = (vnode.props ?? {}) as Record<string, unknown>;
  if (camel in props) return props[camel];
  return props[camel.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`)];
}

/** A boolean vnode prop: present with no value (`hidden`) means true, as Vue casts it. */
function flag(vnode: VNode, camel: string): boolean | undefined {
  const value = prop(vnode, camel);
  if (value === undefined) return undefined;
  return value === '' || value === true || value === 'true';
}

const text = (value: unknown): string | undefined =>
  typeof value === 'string' ? value : undefined;

/** The direct children of a slot, with fragments opened and comments and blank text dropped. */
export function flattenMenuChildren(
  children: VNodeArrayChildren | VNodeChild | undefined,
  out: VNode[] = []
): VNode[] {
  if (Array.isArray(children)) {
    for (const child of children) flattenMenuChildren(child as VNodeChild, out);
    return out;
  }
  if (!isVNode(children)) return out;
  if (children.type === Comment) return out;
  if (children.type === Text && String(children.children ?? '').trim() === '') return out;
  if (children.type === Fragment && Array.isArray(children.children))
    return flattenMenuChildren(children.children as VNodeArrayChildren, out);
  out.push(children);
  return out;
}

/** The default slot content of a component vnode, as vnodes. */
function slotChildren(vnode: VNode): VNode[] {
  const c = vnode.children as unknown;
  if (Array.isArray(c)) return flattenMenuChildren(c as VNodeArrayChildren);
  const slot = (c as { default?: () => VNodeArrayChildren } | null)?.default;
  return typeof slot === 'function' ? flattenMenuChildren(slot()) : [];
}

/** The registry row a `MenuSlotItem` child names, or null for any other child. @internal */
export function slotRowOf(vnode: VNode): MenuSlotRowSpec | null {
  if (vnode.type !== MenuSlotItem) return null;
  return {
    slotId: String(prop(vnode, 'slotId')),
    value: text(prop(vnode, 'value')),
    hidden: flag(vnode, 'hidden') ?? false,
  };
}

// ---------------------------------------------------------------------------------------------
// MenuSlotItem

const menuSlotItemProps = {
  /** The slot the row runs, for example `text.bold`. */
  slotId: { type: String as PropType<ChromeSlotId>, required: true as const },
  /** The value a value slot runs with (`arrange.order` with `front`). */
  value: { type: String, default: undefined },
  /** A translation key for the row's label, in place of the registry label. */
  labelKey: { type: String, default: undefined },
  /** A translation key for the shortcut shown at the end of the row. */
  shortcutKey: { type: String, default: undefined },
  /** In place of the registry icon: a Material Symbol name, or icon content (a vnode). */
  icon: { type: [String, Object] as PropType<string | VNode>, default: undefined },
  /**
   * Run the slot's command at once instead of opening its choices (a value slot such as
   * `slide.add`). Unset, the row runs as the registry row it replaces.
   */
  direct: { type: Boolean, default: undefined },
  /** Remove the row. Its command keeps working elsewhere. */
  hidden: { type: Boolean, default: false },
  /** A class name added to the row, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/**
 * Props for `PptxEditorMenu.Item` and `PptxEditorContextMenu.Slot`: one registry row, the twin
 * of React's `MenuSlotItemProps`. `className` (or a host `class`) lands on the row. @public
 */
export type MenuSlotItemProps = ExtractPublicPropTypes<typeof menuSlotItemProps>;

/**
 * A menu row for a registered slot: its label, icon, shortcut, enabled state, and action come
 * from the chrome registry. A slot that collects a value opens a submenu of its choices.
 * @public
 */
export const MenuSlotItem = defineComponent({
  name: 'MenuSlotItem',
  // The row may be a submenu (a fragment), so the host's `class` goes onto the row here.
  inheritAttrs: false,
  props: menuSlotItemProps,
  setup(props, { attrs }) {
    // In place of a registry row, the row keeps that row's own label and shortcut.
    const replaced = inject(rowDefaultsKey, null);
    return () => {
      if (props.hidden) return null;
      const labelKey = props.labelKey ?? replaced?.labelKey;
      // In place of a registry row that runs its command, the row runs it too, unless it says not to.
      const inherited =
        !!replaced?.direct && replaced.slot === props.slotId && props.value === undefined;
      const direct = props.direct ?? inherited;
      // A row that opens choices does not show the shortcut of the direct row it replaces.
      const shortcutKey =
        props.shortcutKey ?? (inherited && !direct ? undefined : replaced?.shortcutKey);
      const className = normalizeClass([attrs.class, props.className]);
      return (
        <SlotMenuRow
          entry={{
            kind: 'item',
            slot: props.slotId,
            ...(props.value !== undefined ? { value: props.value } : {}),
            ...(direct ? { direct: true as const } : {}),
          }}
          overrides={{
            ...(labelKey ? { labelKey } : {}),
            ...(shortcutKey ? { shortcutKey } : {}),
            ...(className ? { className } : {}),
            ...(props.icon !== undefined ? { icon: props.icon } : {}),
          }}
        />
      );
    };
  },
});

// ---------------------------------------------------------------------------------------------
// MenuGroup

const menuGroupProps = {
  /** A heading shown above the rows. Wins over `labelKey`. */
  label: { type: String, default: undefined },
  /** A translation key for the heading. */
  labelKey: { type: String, default: undefined },
  /** Remove the group and its rows. */
  hidden: { type: Boolean, default: false },
  /** A class name added to the group, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props for `PptxEditorMenu.Group`, the twin of React's `MenuGroupProps`. @public */
export type MenuGroupProps = ExtractPublicPropTypes<typeof menuGroupProps>;

/** A labeled set of rows inside a menu. A host `class` lands on the group. @public */
export const MenuGroup = defineComponent({
  name: 'MenuGroup',
  props: menuGroupProps,
  setup(props, { slots }) {
    const { t } = useTranslation();
    return () => {
      if (props.hidden) return null;
      const heading =
        props.label ?? (props.labelKey ? translateKeyOrText(t, props.labelKey) : undefined);
      return (
        <div role="group" aria-label={heading} class={['pptx-menu__group', props.className]}>
          {heading ? (
            <div class="pptx-menu__heading" aria-hidden="true">
              {heading}
            </div>
          ) : null}
          {slots.default?.()}
        </div>
      );
    };
  },
});

// ---------------------------------------------------------------------------------------------
// MenuEntry and the named menus

const menuPartProps = {
  /** The menu's title. Wins over `labelKey`. Built-in menus keep their own title without it. */
  label: { type: String, default: undefined },
  /** A translation key for the menu's title. */
  labelKey: { type: String, default: undefined },
  /** Icon content before the title: a Material Symbol name, or a vnode. */
  icon: { type: [String, Object] as PropType<string | VNode>, default: undefined },
  /** Remove the menu from the bar. */
  hidden: { type: Boolean, default: false },
  /**
   * `true` (default): the built-in rows, with `MenuSlotItem` children replacing their rows in
   * place and other children added at the end. `false`: only the children.
   */
  preset: { type: Boolean, default: undefined },
  /** For a host menu: the menu it follows. Without it, host menus follow the built-in ones. */
  after: { type: String as PropType<ChromeMenuId | (string & {})>, default: undefined },
  /** A class name added to the menu's panel, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

const menuEntryProps = {
  /** A built-in menu id changes that menu; any other id adds a host menu. */
  id: { type: String as PropType<ChromeMenuId | (string & {})>, required: true as const },
  ...menuPartProps,
} as const;

/**
 * Props for `PptxEditorMenu.Entry`, the twin of React's `MenuEntryProps`. `className` (or a
 * host `class`) goes onto the menu's panel; the default slot is the menu's rows. @public
 */
export type MenuEntryProps = ExtractPublicPropTypes<typeof menuEntryProps>;

/**
 * One menu of the bar. Inside `PptxEditorMenu` it is a description the bar reads; the bar
 * renders it. @public
 */
export const MenuEntry = defineComponent({
  name: 'MenuEntry',
  props: menuEntryProps,
  setup() {
    // `PptxEditorMenu` reads its entries' props; rendered anywhere else, an entry says so.
    return () => {
      warnMenuPart('PptxEditorMenu.Entry', 'outsideBar', undefined, VUE_MENU_PARTS);
      return null;
    };
  },
});

/** Props of a named menu: `MenuEntryProps` without `id`. */
type MenuPartProps = ExtractPublicPropTypes<typeof menuPartProps>;

/** A named menu (`PptxEditorMenu.File`): a `MenuEntry` with its id pinned. @public */
export type MenuPartComponent = DefineComponent<MenuPartProps> & {
  /** The menu the part stands for. */
  readonly pptxMenu: ChromeMenuId;
};

function defineMenu(id: ChromeMenuId): MenuPartComponent {
  const name = menuPartName(id);
  return Object.assign(
    defineComponent({
      name: name.replace('.', ''),
      props: menuPartProps,
      setup() {
        return () => {
          warnMenuPart(`PptxEditor${name}`, 'outsideBar', undefined, VUE_MENU_PARTS);
          return null;
        };
      },
    }),
    { pptxMenu: id }
  ) as unknown as MenuPartComponent;
}

/** The named menus. */
const MENU_PARTS = {
  File: defineMenu('file'),
  Edit: defineMenu('edit'),
  View: defineMenu('view'),
  Insert: defineMenu('insert'),
  Format: defineMenu('format'),
  Arrange: defineMenu('arrange'),
  SlideShow: defineMenu('slideShow'),
} as const;

/** What the bar reads from a `MenuEntry` or named menu child. */
interface EntrySpec {
  readonly id: string;
  readonly label: string | undefined;
  readonly labelKey: string | undefined;
  readonly icon: string | VNode | undefined;
  readonly className: string | undefined;
  readonly hidden: boolean;
  readonly preset: boolean | undefined;
  readonly after: string | undefined;
  readonly children: () => VNode[];
}

/** The menu a child describes: `MenuEntry`, or a named menu with its pinned id. */
function entryOf(vnode: VNode): EntrySpec | null {
  const pinned = (vnode.type as { pptxMenu?: ChromeMenuId } | null)?.pptxMenu;
  if (!pinned && vnode.type !== MenuEntry) return null;
  const icon = prop(vnode, 'icon');
  return {
    id: pinned ?? String(prop(vnode, 'id')),
    label: text(prop(vnode, 'label')),
    labelKey: text(prop(vnode, 'labelKey')),
    icon: typeof icon === 'string' || isVNode(icon) ? icon : undefined,
    className: normalizeClass([prop(vnode, 'class'), prop(vnode, 'className')]) || undefined,
    hidden: flag(vnode, 'hidden') ?? false,
    preset: flag(vnode, 'preset'),
    after: text(prop(vnode, 'after')),
    children: () => slotChildren(vnode),
  };
}

/** The rows of a built-in menu after the host's children. */
function builtInRows(menu: ChromeMenu, spec: EntrySpec | undefined): VNodeChild {
  if (spec?.preset === false) return spec.children();
  const { entries, overrides, extras } = arrangeMenuRows(
    menu.entries,
    spec?.children() ?? [],
    slotRowOf,
    // A hidden row this menu does not have removes nothing, and says so.
    (slot) =>
      warnMenuPart(
        `PptxEditor${menuPartName(menu.id)} > PptxEditorMenu.Item`,
        'unmatchedHidden',
        slot,
        VUE_MENU_PARTS
      )
  );
  const renders = new Map<string, (() => VNode) | null>();
  for (const [key, child] of overrides) renders.set(key, child ? () => child : null);
  return [<MenuEntries entries={entries} overrides={renders} />, ...extras];
}

// ---------------------------------------------------------------------------------------------
// The bar

/** Provides the close function to the rows of one open menu. @internal */
export const MenuCloseScope = defineComponent({
  name: 'PptxMenuCloseScope',
  props: { close: { type: Function as PropType<() => void>, required: true as const } },
  setup(props, { slots }) {
    provide(menuCloseKey, () => props.close());
    return () => slots.default?.();
  },
});

/** One menu as the bar draws it. */
interface BarMenu {
  readonly id: string;
  readonly title: string;
  readonly icon: string | VNode | undefined;
  readonly className: string | undefined;
  readonly content: () => VNodeChild;
}

const menuProps = {
  /**
   * `true` (default): the built-in menus, changed by named menu and `MenuEntry` children;
   * other children follow the menus. `false`: only the children.
   */
  preset: { type: Boolean, default: true },
  /** A class name added to the bar, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/**
 * Props for `PptxEditorMenu`, the twin of React's `MenuProps`. `className` (or a host `class`)
 * lands on the bar; the default slot holds the named menus, `MenuEntry` children, and other bar content.
 * @public
 */
export type MenuProps = ExtractPublicPropTypes<typeof menuProps>;

/**
 * The menu bar: File, Edit, View, Insert, Format, Arrange, and Slide Show, from core's
 * `CHROME_MENUS`. Every row runs a registered slot, so its state and action are the toolbar's.
 * @public
 */
const MenuRoot = defineComponent({
  name: 'PptxEditorMenu',
  props: menuProps,
  setup(props, { slots }) {
    const { t } = useTranslation();
    const open = ref<string | null>(null);
    // Roving tab stop, as DOCX Editor's menu bar: the bar is one Tab stop, the arrows move.
    const active = ref<string | null>(null);
    const triggers = new Map<string, HTMLButtonElement>();
    // Set when hover switched the bar to a menu: the click that follows must not close it.
    let switchedByHover = false;
    // A pointer open puts focus on the trigger so the arrow keys reach the menu; closing gives
    // focus back to where it was (the slide's text input, the slide). A keyboard open, and
    // Escape, return to the trigger.
    let restoreTo: HTMLElement | null = null;
    const bar = ref<HTMLDivElement | null>(null);
    const close = (): void => {
      open.value = null;
    };
    watch(
      open,
      (now) => {
        if (now !== null) return;
        const target = restoreTo;
        restoreTo = null;
        // Focus that a row moved on purpose (a dialog, the find bar, the link popup) stays there.
        if (target && menuBarRestoresFocus(target, bar.value))
          target.focus({ preventScroll: true });
      },
      { flush: 'post' }
    );
    const rtl = (): boolean => {
      const first = triggers.values().next().value;
      return !!first && getComputedStyle(first).direction === 'rtl';
    };
    return () => {
      const specs = new Map<string, EntrySpec>();
      const loose: VNode[] = [];
      for (const child of flattenMenuChildren(slots.default?.())) {
        const spec = entryOf(child);
        if (spec) specs.set(spec.id, spec);
        else loose.push(child);
      }
      // With `preset` off, a named built-in menu still brings its title and rows, as in DOCX
      // Editor; `MenuEntry :preset="false"` keeps only its children.
      const menus: BarMenu[] = arrangeMenuBar(specs, props.preset, t).map(
        ({ id, title, spec, registry }) => ({
          id,
          title,
          icon: spec?.icon,
          className: spec?.className,
          content: () => (registry ? builtInRows(registry, spec) : (spec?.children() ?? null)),
        })
      );
      const stop = menus.some((m) => m.id === active.value) ? active.value : (menus[0]?.id ?? null);
      const step = (from: string, by: -1 | 1, keepOpen: boolean): void => {
        const index = menus.findIndex((m) => m.id === from);
        const next = menus[(index + by + menus.length) % menus.length];
        if (!next) return;
        active.value = next.id;
        // The trigger takes focus first, so an open popover moves focus to its first row.
        triggers.get(next.id)?.focus();
        if (keepOpen) open.value = next.id;
      };
      const onTriggerKey = (id: string, e: KeyboardEvent): void => {
        const action = menuBarTriggerKey(e.key, rtl());
        if (!action) return;
        e.preventDefault();
        if (action.kind === 'step') step(id, action.by, open.value !== null);
        else if (action.kind === 'open') {
          active.value = id;
          open.value = id;
        } else {
          const target = action.edge === 'first' ? menus[0] : menus[menus.length - 1];
          if (!target) return;
          active.value = target.id;
          triggers.get(target.id)?.focus();
        }
      };
      const onPanelKey = (id: string, e: KeyboardEvent): void => {
        const by = menuBarArrow(e.key, rtl());
        if (!by) return;
        e.preventDefault();
        step(id, by, true);
      };
      const onClick = (id: string): void => {
        if (switchedByHover) {
          switchedByHover = false;
          return;
        }
        active.value = id;
        const trigger = triggers.get(id);
        if (open.value === null) {
          const focused = trigger?.ownerDocument.activeElement;
          restoreTo =
            focused instanceof HTMLElement && focused !== focused.ownerDocument.body
              ? focused
              : null;
        }
        // A pointer open puts focus on the trigger, so the arrow keys reach the menu instead of
        // moving a shape on the slide (the popover then focuses its first row).
        trigger?.focus({ preventScroll: true });
        open.value = open.value === id ? null : id;
      };
      return (
        <div
          ref={bar}
          role="menubar"
          aria-label={t('menu.ariaLabel')}
          class={['pptx-menubar', props.className]}
        >
          {menus.map((menu) => [
            <button
              key={`trigger-${menu.id}`}
              ref={(el) => {
                if (el) triggers.set(menu.id, el as HTMLButtonElement);
                else triggers.delete(menu.id);
              }}
              type="button"
              role="menuitem"
              class="pptx-menubar__trigger"
              aria-haspopup="menu"
              aria-expanded={open.value === menu.id}
              tabindex={stop === menu.id ? 0 : -1}
              data-menu={menu.id}
              onMousedown={keepFocus}
              onFocus={() => {
                active.value = menu.id;
              }}
              onKeydown={(e: KeyboardEvent) => onTriggerKey(menu.id, e)}
              onMouseenter={() => {
                if (open.value !== null && open.value !== menu.id) {
                  switchedByHover = true;
                  active.value = menu.id;
                  // Focus follows the open menu, so the old panel's cleanup cannot pull it back.
                  triggers.get(menu.id)?.focus({ preventScroll: true });
                  open.value = menu.id;
                }
              }}
              onMouseleave={() => {
                switchedByHover = false;
              }}
              onClick={() => onClick(menu.id)}
            >
              {/* The same icon rule as the rows: a string is a Material Symbol name. */}
              {typeof menu.icon === 'string' ? <Icon name={menu.icon} /> : (menu.icon ?? null)}
              {menu.title}
            </button>,
            open.value === menu.id ? (
              <Popover
                key={`panel-${menu.id}`}
                anchor={triggers.get(menu.id) ?? null}
                onClose={() => {
                  // Escape leaves focus on the menubar item, not on the slide (WAI-ARIA menubar).
                  if (menuBarClosedOnEscape(triggers.get(menu.id))) restoreTo = null;
                  close();
                }}
                label={menu.title}
              >
                {{
                  default: () => (
                    <MenuCloseScope close={close}>
                      {{
                        default: () => (
                          <div
                            class={['pptx-menubar__panel', menu.className]}
                            data-menu={menu.id}
                            onKeydown={(e: KeyboardEvent) => onPanelKey(menu.id, e)}
                          >
                            {menu.content()}
                          </div>
                        ),
                      }}
                    </MenuCloseScope>
                  ),
                }}
              </Popover>
            ) : null,
          ])}
          {loose}
        </div>
      );
    };
  },
});

/** The menu bar with its parts as statics, the twin of React's `MenuNamespace`. @public */
export type MenuNamespace = typeof MenuRoot & {
  /** A host menu, or a built-in one by id. */
  readonly Entry: typeof MenuEntry;
  /** A registry row: replace, change, or hide it in place. */
  readonly Item: typeof MenuSlotItem;
  /** A host row. */
  readonly Row: typeof MenuRow;
  /** A labeled set of rows. */
  readonly Group: typeof MenuGroup;
  /** A line between rows. */
  readonly Separator: typeof MenuSeparator;
  /** A row that opens a nested panel. */
  readonly Submenu: typeof MenuSubmenu;
  /** The Insert > Table grid, for a host's own Table submenu (DOCX Editor's `Menu.TableGrid`). */
  readonly TableGrid: typeof MenuTableGrid;
  readonly File: MenuPartComponent;
  readonly Edit: MenuPartComponent;
  readonly View: MenuPartComponent;
  readonly Insert: MenuPartComponent;
  readonly Format: MenuPartComponent;
  readonly Arrange: MenuPartComponent;
  readonly SlideShow: MenuPartComponent;
};

/**
 * The menu bar (`PptxEditorMenu`): File, Edit, View, Insert, Format, Arrange, and Slide Show,
 * with DOCX Editor's parts (`Entry`, `Item`, `Row`, `Group`, `Separator`, `Submenu`) and one
 * named part per menu (`PptxEditorMenu.File`, …). @public
 *
 * @example
 * ```vue
 * <PptxEditorMenu>
 *   <PptxEditorMenu.File>
 *     <PptxEditorMenu.Row label="Open…" icon="folder_open" @select="openFile" />
 *   </PptxEditorMenu.File>
 *   <PptxEditorMenu.Edit>
 *     <PptxEditorMenu.Item slot-id="more.delete" hidden />
 *   </PptxEditorMenu.Edit>
 *   <PptxEditorMenu.Arrange hidden />
 * </PptxEditorMenu>
 * ```
 */
export const PptxEditorMenu: MenuNamespace = Object.assign(MenuRoot, {
  ...MENU_PARTS,
  Entry: MenuEntry,
  Item: MenuSlotItem,
  Row: MenuRow,
  Group: MenuGroup,
  Separator: MenuSeparator,
  Submenu: MenuSubmenu,
  TableGrid: MenuTableGrid,
});
