/** @jsxImportSource react */
// `PptxEditor.Menu`: the menu bar as library parts, after DOCX Editor's `DocxEditor.Menu`
// (docx-editor/docs/site/content/customization/menus.mdx, packages/react/src/editor/menu/). DOCX
// Editor commit adf0e178.
//
// The preset bar renders core's `CHROME_MENUS` (File, Edit, View, Insert, Format, Arrange,
// Slide Show). A named menu child (`Menu.File`) changes its menu in place: `hidden` removes
// it, `Menu.Item slotId` replaces or hides one row, other children add rows at its end, and
// `preset={false}` keeps only its children. `Menu.Entry` with another id adds a host menu.
// `preset={false}` on the bar renders only the children. Hiding a row changes the interface
// only: its command still runs from the toolbar, the keyboard, and the API.

import type { ChromeTranslate } from '@pptx-editor.dev/core/editor';
import {
  Children,
  Fragment,
  isValidElement,
  useCallback,
  useContext,
  useEffect,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type ReactNode,
} from 'react';
import {
  arrangeMenuBar,
  arrangeMenuRows,
  menuBarClosedOnEscape,
  menuBarRestoresFocus,
  menuBarTriggerKey,
  menuBarArrow,
  menuPartName,
  translateKeyOrText,
  warnMenuPart,
  MENU_NAMED_PARTS,
  type ChromeMenu,
  type ChromeMenuId,
  type ChromeSlotId,
  type MenuNamedRow,
} from '@pptx-editor.dev/core/editor';
import { Icon, Popover, keepFocus } from './chrome.tsx';
import { MenuCloseContext } from './ContextMenu.tsx';
import { useTranslation } from './context.ts';
import { MenuEntries, RowDefaultsContext, SlotMenuRow } from './menu-rows.tsx';

/** Props for `Menu.Item` and `ContextMenu.Slot`: one registry row. @public */
export interface MenuSlotItemProps {
  /** The slot the row runs, for example `text.bold`. */
  readonly slotId: ChromeSlotId;
  /** The value a value slot runs with (`arrange.order` with `front`). */
  readonly value?: string;
  /** A translation key for the row's label, in place of the registry label. */
  readonly labelKey?: string;
  /** A translation key for the shortcut shown at the end of the row. */
  readonly shortcutKey?: string;
  /**
   * In place of the registry icon: a Material Symbol name (as `Menu.Row` reads it), or icon
   * content (an element).
   */
  readonly icon?: ReactNode;
  /**
   * Run the slot's command at once instead of opening its choices (a value slot such as
   * `slide.add`). Unset, the row runs as the registry row it replaces: Insert > New slide adds a
   * slide at once; `direct={false}` opens the layout gallery there.
   */
  readonly direct?: boolean;
  /** A class name added to the row. */
  readonly className?: string;
  /** Remove the row. Its command keeps working elsewhere. */
  readonly hidden?: boolean;
}

/**
 * A menu row for a registered slot: its label, icon, shortcut, enabled state, and action come
 * from the chrome registry. A slot that collects a value opens a submenu of its choices.
 * @public
 */
export function MenuSlotItem(props: MenuSlotItemProps): ReactNode {
  // In place of a registry row, the row keeps that row's own label and shortcut.
  const replaced = useContext(RowDefaultsContext);
  if (props.hidden) return null;
  const labelKey = props.labelKey ?? replaced?.labelKey;
  // In place of a registry row that runs its command, the row runs it too, unless it says not to.
  const inherited =
    !!replaced?.direct && replaced.slot === props.slotId && props.value === undefined;
  const direct = props.direct ?? inherited;
  // A row that opens choices does not show the shortcut of the direct row it replaces.
  const shortcutKey =
    props.shortcutKey ?? (inherited && !direct ? undefined : replaced?.shortcutKey);
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
        ...(props.className ? { className: props.className } : {}),
        ...(props.icon !== undefined ? { icon: props.icon } : {}),
      }}
    />
  );
}

/** Props for a named menu row (`Menu.Cut`, `ContextMenu.Cut`, …), as DOCX Editor's. @public */
export interface MenuRowPartProps {
  /** The row's text, or a translation key, in place of the registry label. */
  readonly label?: string;
  /** A translation key for the row's label, in place of the registry label. */
  readonly labelKey?: string;
  /** A translation key for the shortcut shown at the end of the row. */
  readonly shortcutKey?: string;
  /** In place of the registry icon: a Material Symbol name, or icon content (an element). */
  readonly icon?: ReactNode;
  /** A class name added to the row. */
  readonly className?: string;
  /** Remove the row. Its command keeps working elsewhere. */
  readonly hidden?: boolean;
}

/** A named menu row: a `Menu.Item` with its slot pinned. @public */
export interface MenuRowPartComponent {
  (props: MenuRowPartProps): ReactNode;
  /** The registry row the part stands for. */
  readonly pptxMenuRow: MenuNamedRow;
}

/**
 * A named row, after DOCX Editor's `defineCommandRow`: a `Menu.Item` with its slot and value
 * pinned, and the row as a static (`pptxMenuRow`), so a menu finds it among its children.
 * @internal
 */
export function defineMenuRowPart(row: MenuNamedRow): MenuRowPartComponent {
  const Part = ({ label, labelKey, ...rest }: MenuRowPartProps): ReactNode => (
    <MenuSlotItem
      {...rest}
      slotId={row.slot}
      {...(row.value !== undefined ? { value: row.value } : {})}
      {...((label ?? labelKey) !== undefined ? { labelKey: label ?? labelKey } : {})}
    />
  );
  return Object.assign(Part, { pptxMenuRow: row });
}

/** The named rows of `table`, by name. @internal */
export function defineMenuRowParts<K extends string>(
  table: Readonly<Record<K, MenuNamedRow>>
): Record<K, MenuRowPartComponent> {
  return Object.fromEntries(
    Object.entries<MenuNamedRow>(table).map(([name, row]) => [name, defineMenuRowPart(row)])
  ) as Record<K, MenuRowPartComponent>;
}

/**
 * The registry row a menu child names: a `Menu.Item`, or a named row with its row pinned. Null
 * for any other child. @internal
 */
export function menuSlotRowOf(node: ReactNode): MenuSlotItemProps | null {
  if (!isValidElement<NodeProps>(node)) return null;
  if (node.type === MenuSlotItem) return node.props as unknown as MenuSlotItemProps;
  const row = (node.type as { pptxMenuRow?: MenuNamedRow } | null)?.pptxMenuRow;
  if (!row) return null;
  return {
    ...(node.props as unknown as MenuRowPartProps),
    slotId: row.slot,
    ...(row.value !== undefined ? { value: row.value } : {}),
  };
}

/** The named rows of the menu bar, from core's table. */
const MENU_ROW_PARTS = defineMenuRowParts(MENU_NAMED_PARTS);

/** Props for `Menu.Group`. @public */
export interface MenuGroupProps {
  /** A heading shown above the rows. Wins over `labelKey`. */
  readonly label?: string;
  /** A translation key for the heading. */
  readonly labelKey?: string;
  /** A class name added to the group. */
  readonly className?: string;
  /** Remove the group and its rows. */
  readonly hidden?: boolean;
  /** The group's rows. */
  readonly children?: ReactNode;
}

/** A labeled set of rows inside a menu. @public */
export function MenuGroup({ label, labelKey, className, hidden, children }: MenuGroupProps) {
  const { t } = useTranslation();
  if (hidden) return null;
  const heading = label ?? (labelKey ? translateKeyOrText(t, labelKey) : undefined);
  return (
    <div
      role="group"
      aria-label={heading}
      className={['pptx-menu__group', className].filter(Boolean).join(' ')}
    >
      {heading ? (
        <div className="pptx-menu__heading" aria-hidden="true">
          {heading}
        </div>
      ) : null}
      {children}
    </div>
  );
}

/** Props for `Menu.Entry` and the named menus (`Menu.File`, …). @public */
export interface MenuEntryProps {
  /** A built-in menu id changes that menu; any other id adds a host menu. */
  readonly id: ChromeMenuId | (string & {});
  /** The menu's title. Wins over `labelKey`. Built-in menus keep their own title without it. */
  readonly label?: string;
  /** A translation key for the menu's title. */
  readonly labelKey?: string;
  /** Icon content before the title. */
  readonly icon?: ReactNode;
  /** A class name added to the menu's panel. */
  readonly className?: string;
  /** Remove the menu from the bar. */
  readonly hidden?: boolean;
  /**
   * `true` (default): the built-in rows, with `Menu.Item` children replacing their rows in
   * place and other children added at the end. `false`: only the children.
   */
  readonly preset?: boolean;
  /** For a host menu: the menu it follows. Without it, host menus follow the built-in ones. */
  readonly after?: ChromeMenuId | (string & {});
  /** The menu's rows: `Menu.Item` children change built-in rows; others add rows. */
  readonly children?: ReactNode;
}

/**
 * One menu of the bar. Inside `PptxEditor.Menu` it is a description the bar reads; the bar
 * renders it. @public
 */
export function MenuEntry(_props: MenuEntryProps): ReactNode {
  // `PptxEditor.Menu` reads its entries' props; rendered anywhere else, an entry says so.
  warnMenuPart('Menu.Entry', 'outsideBar');
  return null;
}

/** A named menu: a `Menu.Entry` with its id pinned. @public */
export interface MenuPartComponent {
  /** Changes the built-in menu the part stands for, as a `Menu.Entry` with its id. */
  (props: Omit<MenuEntryProps, 'id'>): ReactNode;
  /** The menu the part stands for. */
  readonly pptxMenu: ChromeMenuId;
}

function defineMenu(id: ChromeMenuId): MenuPartComponent {
  const Part = (_props: Omit<MenuEntryProps, 'id'>): ReactNode => {
    warnMenuPart(`Menu.${id[0]!.toUpperCase()}${id.slice(1)}`, 'outsideBar');
    return null;
  };
  return Object.assign(Part, { pptxMenu: id });
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

type NodeProps = Record<string, unknown> & { children?: ReactNode };

/** The direct children of `nodes`, with fragments opened. */
const flatten = (nodes: ReactNode): ReactNode[] =>
  Children.toArray(nodes).flatMap((node) =>
    isValidElement<NodeProps>(node) && node.type === Fragment
      ? flatten(node.props.children)
      : [node]
  );

/** The menu a child describes: `Menu.Entry`, or a named menu with its pinned id. */
function entryOf(node: ReactNode): MenuEntryProps | null {
  if (!isValidElement<NodeProps>(node)) return null;
  const pinned = (node.type as { pptxMenu?: ChromeMenuId } | null)?.pptxMenu;
  if (pinned) return { ...(node.props as Omit<MenuEntryProps, 'id'>), id: pinned };
  return node.type === MenuEntry ? (node.props as unknown as MenuEntryProps) : null;
}

/** One menu as the bar draws it. */
interface BarMenu {
  readonly id: string;
  readonly title: string;
  readonly icon?: ReactNode;
  readonly className?: string;
  readonly content: ReactNode;
}

/** The rows of a built-in menu after the host's children. */
function builtInRows(menu: ChromeMenu, spec: MenuEntryProps | undefined): ReactNode {
  if (spec?.preset === false) return spec.children;
  const { entries, overrides, extras } = arrangeMenuRows(
    menu.entries,
    flatten(spec?.children),
    menuSlotRowOf,
    // A hidden row this menu does not have removes nothing, and says so.
    (slot) => warnMenuPart(`${menuPartName(menu.id)} > Menu.Item`, 'unmatchedHidden', slot)
  );
  return (
    <>
      <MenuEntries entries={entries} overrides={overrides} />
      {extras}
    </>
  );
}

/** Props for `PptxEditor.Menu`. @public */
export interface MenuProps {
  /**
   * Your translation function for this compound and its parts, as DOCX Editor's `t`. Without
   * it, the root's `locale` gives the labels.
   */
  readonly t?: ChromeTranslate;
  /** A class name added to the bar. */
  readonly className?: string;
  /**
   * `true` (default): the built-in menus, changed by named menu and `Menu.Entry` children;
   * other children follow the menus. `false`: only the children.
   */
  readonly preset?: boolean;
  /** Named menus and `Menu.Entry` children, and other content after the menus. */
  readonly children?: ReactNode;
}

/**
 * The menu bar: File, Edit, View, Insert, Format, Arrange, and Slide Show, from core's
 * `CHROME_MENUS`. Every row runs a registered slot, so its state and action are the toolbar's.
 *
 * @example
 * ```tsx
 * <PptxEditor.Menu>
 *   <PptxEditor.Menu.File>
 *     <PptxEditor.Menu.Row label="Open…" icon="folder_open" onSelect={openFile} />
 *   </PptxEditor.Menu.File>
 *   <PptxEditor.Menu.Edit>
 *     <PptxEditor.Menu.Item slotId="more.delete" hidden />
 *   </PptxEditor.Menu.Edit>
 *   <PptxEditor.Menu.Arrange hidden />
 * </PptxEditor.Menu>
 * ```
 *
 * @public
 */
export function Menu({ className, preset = true, children }: MenuProps): ReactNode {
  const { t } = useTranslation();
  const specs = new Map<string, MenuEntryProps>();
  const loose: ReactNode[] = [];
  for (const child of flatten(children)) {
    const spec = entryOf(child);
    if (spec) specs.set(spec.id, spec);
    else loose.push(child);
  }
  // With `preset={false}`, a named built-in menu still brings its title and rows, as in DOCX
  // Editor; `Menu.Entry preset={false}` keeps only its children.
  const menus: BarMenu[] = arrangeMenuBar(specs, preset, t).map(
    ({ id, title, spec, registry }) => ({
      id,
      title,
      ...(spec?.icon !== undefined ? { icon: spec.icon } : {}),
      ...(spec?.className ? { className: spec.className } : {}),
      content: registry ? builtInRows(registry, spec) : spec?.children,
    })
  );
  return <MenuBar className={className} menus={menus} extras={loose} />;
}

function MenuBar({
  className,
  menus,
  extras,
}: {
  className: string | undefined;
  menus: readonly BarMenu[];
  extras: readonly ReactNode[];
}): ReactNode {
  const { t } = useTranslation();
  const [open, setOpen] = useState<string | null>(null);
  // Roving tab stop, as DOCX Editor's menu bar: the bar is one Tab stop, the arrows move.
  const [active, setActive] = useState<string | null>(null);
  const stop = menus.some((m) => m.id === active) ? active : (menus[0]?.id ?? null);
  const triggers = useRef(new Map<string, HTMLButtonElement>());
  // Set when hover switched the bar to a menu: the click that follows must not close it
  // (DOCX Editor's `switchedByHover`).
  const switchedByHover = useRef(false);
  // A pointer open puts focus on the trigger so the arrow keys reach the menu; closing gives
  // focus back to where it was (the slide's text input, the slide), as the context menu does
  // and as DOCX Editor keeps pointer focus. A keyboard open, and Escape, return to the trigger.
  const restoreTo = useRef<HTMLElement | null>(null);
  const bar = useRef<HTMLDivElement>(null);
  const close = useCallback(() => setOpen(null), []);
  useEffect(() => {
    if (open !== null) return;
    const target = restoreTo.current;
    restoreTo.current = null;
    // Focus that a row moved on purpose (a dialog, the find bar, the link popup) stays there.
    if (target && menuBarRestoresFocus(target, bar.current)) target.focus({ preventScroll: true });
  }, [open]);
  const step = (from: string, by: -1 | 1, keepOpen: boolean): void => {
    const index = menus.findIndex((m) => m.id === from);
    const next = menus[(index + by + menus.length) % menus.length];
    if (!next) return;
    setActive(next.id);
    // The trigger takes focus first, so an open popover moves focus to its first row.
    triggers.current.get(next.id)?.focus();
    if (keepOpen) setOpen(next.id);
  };
  const rtl = (): boolean => {
    const first = triggers.current.values().next().value;
    return !!first && getComputedStyle(first).direction === 'rtl';
  };
  const onTriggerKey = (id: string, e: ReactKeyboardEvent): void => {
    const action = menuBarTriggerKey(e.key, rtl());
    if (!action) return;
    e.preventDefault();
    if (action.kind === 'step') step(id, action.by, open !== null);
    else if (action.kind === 'open') {
      setActive(id);
      setOpen(id);
    } else {
      const target = action.edge === 'first' ? menus[0] : menus[menus.length - 1];
      if (!target) return;
      setActive(target.id);
      triggers.current.get(target.id)?.focus();
    }
  };
  const onPanelKey = (id: string, e: ReactKeyboardEvent): void => {
    const by = menuBarArrow(e.key, rtl());
    if (!by) return;
    e.preventDefault();
    step(id, by, true);
  };
  return (
    <div
      ref={bar}
      role="menubar"
      aria-label={t('menu.ariaLabel')}
      className={['pptx-menubar', className].filter(Boolean).join(' ')}
    >
      {menus.map((menu) => (
        <Fragment key={menu.id}>
          <button
            ref={(el) => {
              if (el) triggers.current.set(menu.id, el);
              else triggers.current.delete(menu.id);
            }}
            type="button"
            role="menuitem"
            className="pptx-menubar__trigger"
            aria-haspopup="menu"
            aria-expanded={open === menu.id}
            tabIndex={stop === menu.id ? 0 : -1}
            data-menu={menu.id}
            onMouseDown={keepFocus}
            onFocus={() => setActive(menu.id)}
            onKeyDown={(e) => onTriggerKey(menu.id, e)}
            onMouseEnter={() => {
              if (open !== null && open !== menu.id) {
                switchedByHover.current = true;
                setActive(menu.id);
                // Focus follows the open menu, so the old panel's cleanup cannot pull it back.
                triggers.current.get(menu.id)?.focus({ preventScroll: true });
                setOpen(menu.id);
              }
            }}
            onMouseLeave={() => (switchedByHover.current = false)}
            onClick={() => {
              if (switchedByHover.current) {
                switchedByHover.current = false;
                return;
              }
              setActive(menu.id);
              const trigger = triggers.current.get(menu.id);
              if (open === null) {
                const active = trigger?.ownerDocument.activeElement;
                restoreTo.current =
                  active instanceof HTMLElement && active !== active.ownerDocument.body
                    ? active
                    : null;
              }
              // A pointer open puts focus on the trigger, so the arrow keys reach the menu
              // instead of moving a shape on the slide (the popover then focuses its first row).
              trigger?.focus({ preventScroll: true });
              setOpen((cur) => (cur === menu.id ? null : menu.id));
            }}
          >
            {/* The same icon rule as the rows: a string is a Material Symbol name. */}
            {typeof menu.icon === 'string' ? <Icon name={menu.icon} /> : (menu.icon ?? null)}
            {menu.title}
          </button>
          {open === menu.id ? (
            <Popover
              anchor={triggers.current.get(menu.id) ?? null}
              onClose={() => {
                // Escape leaves focus on the menubar item, not on the slide (WAI-ARIA menubar).
                if (menuBarClosedOnEscape(triggers.current.get(menu.id))) restoreTo.current = null;
                close();
              }}
              label={menu.title}
            >
              <MenuCloseContext.Provider value={close}>
                <div
                  className={['pptx-menubar__panel', menu.className].filter(Boolean).join(' ')}
                  data-menu={menu.id}
                  onKeyDown={(e) => onPanelKey(menu.id, e)}
                >
                  {menu.content}
                </div>
              </MenuCloseContext.Provider>
            </Popover>
          ) : null}
        </Fragment>
      ))}
      {extras}
    </div>
  );
}

export { MENU_PARTS, MENU_ROW_PARTS };
