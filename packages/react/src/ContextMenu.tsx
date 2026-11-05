/** @jsxImportSource react */
// Right-click menus and menu rows with an icon column, a label, and a shortcut hint.
//
// Adapted from DOCX Editor (eigenpal/docx-editor, packages/react/src/editor/contextmenu/
// DocxEditorContextMenu.tsx and parts.tsx, and the `.docx-contextmenu` and
// `.docx-menubar__item` rules of packages/core/src/styles/editor.css): the menu opens at the
// pointer and flips to stay inside the window, focus moves into it and returns on close, and a
// disabled row stays in the menu with the engine's reason as its tooltip. The rows share the
// popover menu's keyboard rules (`menu-keyboard.ts`). PPTX Editor adds `MenuSubmenu`, the
// cascading panel of the demo's Insert menu.

import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type ReactNode,
} from 'react';
import { createPortal } from 'react-dom';
import { sanitizeHref } from '@pptx-editor.dev/core/store';
import { Icon, keepFocus } from './chrome.tsx';
import {
  tooltipAttributes,
  focusEdge,
  focusMenuRowAtPointer,
  focusPanelFromPointer,
  handleMenuKey,
  listenContextMenuDismiss,
  panelKeyPressed,
  focusSurfaceOf,
  placeInViewport as placeAt,
  panelItems,
  placeSubmenu,
  submenuPanelKey,
  translateKeyOrText,
} from '@pptx-editor.dev/core/editor';
import { PptxEditorContext, TranslationContext } from './context.ts';

/** Closes the menu a row is in; `null` outside a packaged menu. @internal */
export const MenuCloseContext = createContext<(() => void) | null>(null);

/** A point in client (viewport) pixels. @public */
export interface ContextMenuPoint {
  /** The horizontal position in CSS pixels from the left of the window. */
  readonly x: number;
  /** The vertical position in CSS pixels from the top of the window. */
  readonly y: number;
}

/** Props of `ContextMenu`. @public */
export interface ContextMenuProps {
  /** Where the menu opens: the pointer, or the focused element for a keyboard open. */
  readonly at: ContextMenuPoint;
  /** Called when the menu closes: a row ran, Escape, Tab, or a click outside. */
  readonly onClose: () => void;
  /** The accessible name of the menu. */
  readonly label: string;
  /** Opened from the keyboard (Shift+F10, the Menu key): focus goes to the first row. */
  readonly keyboard?: boolean;
  /** The element the menu opened on; its editor scope carries the theme tokens. */
  readonly owner?: HTMLElement | null;
  /** A class name added to the menu. */
  readonly className?: string;
  /** The menu's rows. */
  readonly children: ReactNode;
}

/**
 * Place a panel of size `w` × `h` at `at`, flipped and clamped to stay inside the window. The
 * rule lives in core (`placeInViewport`), shared with the Vue adapter.
 */
export const placeInViewport: (
  at: ContextMenuPoint,
  w: number,
  h: number,
  vw: number,
  vh: number
) => { left: number; top: number } = placeAt;

/** The editor scope an element lives in, so a portaled panel inherits its tokens. */
const scopeOf = (el: HTMLElement | null | undefined): HTMLElement =>
  (el?.closest('.pptx-editor') as HTMLElement | null) ?? document.body;

/**
 * A context menu at a point. Closed by Escape (focus returns to where it was), Tab, a click
 * outside, or a window resize. The arrow keys, Home, and End move through the rows. @public
 */
export function ContextMenu({
  at,
  onClose,
  label,
  keyboard,
  owner,
  className,
  children,
}: ContextMenuProps) {
  const ref = useRef<HTMLDivElement>(null);
  const editor = useContext(PptxEditorContext);
  // Read on close: the menu belongs to one editor for its whole life.
  const editorRef = useRef(editor);
  editorRef.current = editor;
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null);
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;
  useLayoutEffect(() => {
    const panel = ref.current;
    if (!panel) return;
    setPos(
      placeInViewport(
        at,
        panel.offsetWidth,
        panel.offsetHeight,
        window.innerWidth,
        window.innerHeight
      )
    );
  }, [at]);
  useLayoutEffect(() => {
    const panel = ref.current;
    if (!panel) return;
    const back = panel.ownerDocument.activeElement as HTMLElement | null;
    return () => {
      // Focus returns only when it is still in the menu: a click elsewhere keeps its focus.
      const active = panel.ownerDocument.activeElement;
      const lost = !active || active === panel.ownerDocument.body || panel.contains(active);
      if (!lost) return;
      // On the slide, focus goes where typing works: the text input while text is selected
      // (a right press moves focus to the slide frame first).
      const owner = editorRef.current;
      if (back?.closest('.pptx-surface') && owner && focusSurfaceOf(owner)) return;
      if (back?.isConnected) back.focus({ preventScroll: true });
    };
  }, []);
  // Focus moves in once the panel is placed: a hidden panel cannot take focus.
  const placed = pos !== null;
  useEffect(() => {
    const panel = ref.current;
    if (!panel || !placed) return;
    // A pointer open focuses the panel, with no ring until a key is pressed.
    if (!keyboard || !focusEdge(panelItems(panel), 'first')) focusPanelFromPointer(panel);
  }, [placed, keyboard]);
  // A press outside, Escape, a blur, a resize, or a scroll closes it (core's
  // `listenContextMenuDismiss`).
  useEffect(
    () =>
      listenContextMenuDismiss(
        () => ref.current,
        () => onCloseRef.current()
      ),
    []
  );
  const onKeyDown = (e: ReactKeyboardEvent): void => {
    panelKeyPressed(ref.current);
    if (!ref.current?.contains(e.target as Node)) return;
    if (e.key === 'Escape' || e.key === 'Tab') {
      e.preventDefault();
      e.stopPropagation();
      onCloseRef.current();
      return;
    }
    if (handleMenuKey(ref.current, e.key)) {
      e.preventDefault();
      e.stopPropagation();
    }
  };
  return createPortal(
    <div
      ref={ref}
      className={['pptx-menu pptx-contextmenu', className].filter(Boolean).join(' ')}
      role="menu"
      aria-label={label}
      tabIndex={-1}
      style={pos ?? { visibility: 'hidden', left: at.x, top: at.y }}
      onMouseDown={keepFocus}
      // The pointer moves the focused row, so one row looks active.
      onMouseMove={(e) => focusMenuRowAtPointer(ref.current, e)}
      onContextMenu={(e) => e.preventDefault()}
      onKeyDown={onKeyDown}
    >
      {children}
    </div>,
    scopeOf(owner)
  );
}

/** Props of `MenuRow`. @public */
export interface MenuRowProps {
  /** The row's text, and its accessible name. */
  readonly label: string;
  /**
   * A Material Symbol name from `ICON_PATHS`, or icon content (an element), as DOCX Editor's
   * `ContextMenuItem` and `Toolbar.Action` take.
   */
  readonly icon?: string | ReactNode;
  /** SVG path data for the icon, in place of `icon`. */
  readonly paths?: readonly string[] | null;
  /** Other icon content (a drawn shape or line), in place of `icon` and `paths`. */
  readonly glyph?: ReactNode;
  /** The shortcut hint, already in this platform's form (`useTranslation` gives it). */
  readonly shortcut?: string;
  /** The row cannot run. It stays focusable and names `reason`. */
  readonly disabled?: boolean;
  /** Why the row is disabled; shown as its tooltip. */
  readonly reason?: string | null;
  /** The same as `reason`, under DOCX Editor's and `Toolbar.Action`'s name. */
  readonly disabledReason?: string | null;
  /** A toggle that is on (Bold on bold text). */
  readonly active?: boolean;
  /** An on/off setting (View > Smart guides): a `menuitemcheckbox` with a check when on. */
  readonly checked?: boolean;
  /**
   * With `checked`, one choice of several (a Shadow preset, or View > Normal and Slide sorter):
   * a `menuitemradio` instead of a `menuitemcheckbox`.
   */
  readonly radio?: boolean;
  /** The chrome slot or command the row runs, as `data-slot` for hosts and tests. */
  readonly slot?: string;
  /** A class name added to the row. */
  readonly className?: string;
  /**
   * A web link: the row opens it in a new tab. Only `http`, `https`, `mailto`, `tel`, and `ftp`
   * links pass `sanitizeHref`; any other link disables the row.
   */
  readonly href?: string;
  /** Runs the row. Inside `PptxEditor.Menu` or `PptxEditor.ContextMenu`, the menu closes after. */
  readonly onSelect?: () => void;
}

/**
 * A menu row: icon, label, and shortcut. A disabled row stays focusable with `aria-disabled`,
 * so a keyboard or pointer user can read why. @public
 */
export function MenuRow(props: MenuRowProps) {
  const { label, icon, paths, glyph, shortcut, active, checked, slot, href } = props;
  const close = useContext(MenuCloseContext);
  const link = href === undefined ? null : sanitizeHref(href);
  const disabled = props.disabled || (link !== null && !link.ok);
  const reason = props.reason ?? props.disabledReason;
  const checkbox = checked === undefined ? {} : { 'aria-checked': checked };
  const shared = {
    role: checked === undefined ? 'menuitem' : props.radio ? 'menuitemradio' : 'menuitemcheckbox',
    ...checkbox,
    tabIndex: -1,
    className: ['pptx-menu__item pptx-menu__row', props.className].filter(Boolean).join(' '),
    'aria-disabled': disabled || undefined,
    // The row shows its name; the tooltip adds only the reason, which is also its description.
    ...(disabled && reason ? { 'aria-description': reason, 'data-pptx-tooltip': reason } : {}),
    'data-slot': slot,
    'data-active': active ? '' : undefined,
    'data-disabled': disabled ? '' : undefined,
    onClick: () => {
      if (disabled) return;
      props.onSelect?.();
      close?.();
    },
  };
  const content = (
    <>
      <span className="pptx-menu__icon" aria-hidden="true">
        {glyph ??
          (typeof icon === 'string' || paths ? (
            <Icon name={typeof icon === 'string' ? icon : undefined} paths={paths} />
          ) : (
            (icon ?? null)
          ))}
      </span>
      <span className="pptx-menu__label">{label}</span>
      {checked ? <Icon name="check" className="pptx-icon pptx-menu__check" /> : null}
      {/* The shortcut is a visual hint; the row's accessible name is its label. */}
      {shortcut ? (
        <span className="pptx-menu__shortcut" aria-hidden="true">
          {shortcut}
        </span>
      ) : null}
    </>
  );
  if (link?.ok && !props.disabled)
    return (
      <a {...shared} href={link.href} target="_blank" rel="noopener noreferrer">
        {content}
      </a>
    );
  return (
    <button type="button" {...shared}>
      {content}
    </button>
  );
}

/** A line between groups of rows. @public */
export function MenuSeparator() {
  return <div className="pptx-menu__sep" role="separator" />;
}

/** Props of `MenuSubmenu`. @public */
export interface MenuSubmenuProps {
  /** The row's label. Wins over `labelKey`. */
  readonly label?: string;
  /** A translation key for the label, as DOCX Editor's `Menu.Submenu` takes. */
  readonly labelKey?: string;
  /** A Material Symbol name, or icon content (an element). */
  readonly icon?: string | ReactNode;
  /** SVG path data for the icon, in place of `icon`. */
  readonly paths?: readonly string[] | null;
  /** Other icon content (a drawn shape or line), in place of `icon` and `paths`. */
  readonly glyph?: ReactNode;
  /** The row cannot open its panel. It stays focusable and names `reason`. */
  readonly disabled?: boolean;
  /** Why the row is disabled; shown as its tooltip. */
  readonly reason?: string | null;
  /** The submenu's own class, for a grid of shapes. */
  readonly className?: string;
  /** The rows of the cascading panel. */
  readonly children: ReactNode;
}

/**
 * A row that opens a cascading panel to its right (to its left at the window's edge). It opens
 * on hover, click, Enter, and ArrowRight, and closes on ArrowLeft and Escape, which give focus
 * back to the row. @public
 */
export function MenuSubmenu(props: MenuSubmenuProps) {
  const { icon, paths, glyph, disabled, reason, className, children } = props;
  const t = useContext(TranslationContext);
  const label = props.label ?? (props.labelKey ? translateKeyOrText(t, props.labelKey) : '');
  const row = useRef<HTMLButtonElement>(null);
  const panel = useRef<HTMLDivElement>(null);
  const [open, setOpen] = useState<false | 'pointer' | 'keyboard'>(false);
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null);
  const close = useCallback((focusRow: boolean): void => {
    setOpen(false);
    setPos(null);
    if (focusRow) row.current?.focus();
  }, []);
  useLayoutEffect(() => {
    const r = row.current?.getBoundingClientRect();
    const p = panel.current;
    if (!open || !r || !p) return;
    setPos(
      placeSubmenu(
        r,
        { width: p.offsetWidth, height: p.offsetHeight },
        { width: window.innerWidth, height: window.innerHeight }
      )
    );
  }, [open]);
  // Focus moves in once the panel is placed: a hidden panel cannot take focus.
  useEffect(() => {
    if (pos && open === 'keyboard' && panel.current) focusEdge(panelItems(panel.current), 'first');
  }, [pos, open]);
  // Hovering another row of the parent menu closes this one, as in Google Slides.
  useEffect(() => {
    if (!open) return;
    const over = (e: Event): void => {
      const target = e.target as Node;
      const parent = row.current?.closest('.pptx-menu');
      if (!parent?.contains(target) || row.current?.contains(target)) return;
      close(false);
    };
    document.addEventListener('mouseover', over);
    return () => document.removeEventListener('mouseover', over);
  }, [open, close]);
  const onRowKey = (e: ReactKeyboardEvent): void => {
    if (disabled || (e.key !== 'ArrowRight' && e.key !== 'Enter' && e.key !== ' ')) return;
    e.preventDefault();
    e.stopPropagation();
    setOpen('keyboard');
  };
  const onPanelKey = (e: ReactKeyboardEvent): void => {
    if (submenuPanelKey(panel.current, e)) close(true);
  };
  return (
    <>
      <button
        ref={row}
        type="button"
        role="menuitem"
        tabIndex={-1}
        className="pptx-menu__item pptx-menu__row"
        aria-haspopup="menu"
        aria-expanded={!!open}
        aria-disabled={disabled || undefined}
        {...(disabled && reason
          ? { 'aria-description': reason, ...tooltipAttributes(label, true, reason) }
          : {})}
        data-open={open ? '' : undefined}
        onMouseEnter={() => !disabled && setOpen((o) => o || 'pointer')}
        onClick={() => !disabled && setOpen('keyboard')}
        onKeyDown={onRowKey}
      >
        <span className="pptx-menu__icon" aria-hidden="true">
          {glyph ??
            (typeof icon === 'string' || paths ? (
              <Icon name={typeof icon === 'string' ? icon : undefined} paths={paths} />
            ) : (
              (icon ?? null)
            ))}
        </span>
        <span className="pptx-menu__label">{label}</span>
        <Icon name="chevron_right" className="pptx-icon pptx-menu__caret" />
      </button>
      {open
        ? createPortal(
            <div
              ref={panel}
              className={`pptx-menu pptx-menu--submenu${className ? ` ${className}` : ''}`}
              role="menu"
              aria-label={label}
              style={pos ?? { visibility: 'hidden', left: 0, top: 0 }}
              onMouseDown={keepFocus}
              onMouseMove={(e) => focusMenuRowAtPointer(panel.current, e)}
              onKeyDown={onPanelKey}
            >
              {children}
            </div>,
            scopeOf(row.current)
          )
        : null}
    </>
  );
}
