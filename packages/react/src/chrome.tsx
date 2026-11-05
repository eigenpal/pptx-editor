/** @jsxImportSource react */
// Chrome building blocks: icons, buttons, and popover menus shared by the toolbar, rail, notes,
// panel, and zoom pill. Every mousedown on chrome is prevented (except form fields) so it never
// steals the slide selection or the text caret.

import {
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type ReactElement,
  type ReactNode,
  type MouseEvent as ReactMouseEvent,
  useId,
} from 'react';
import { createPortal } from 'react-dom';
import {
  chromeKeepsFocus,
  tooltipAttributes,
  focusMenuRowAtPointer,
  focusPopoverOnOpen,
  focusTextAfterChrome,
  ICON_PATHS,
  listenPopoverDismiss,
  popoverKeyDown,
  popoverScope,
  positionPopover,
  restorePopoverFocus,
  warnUnknownIcon,
  warnUnsupportedOverride,
  type PopoverAnchorRect,
} from '@pptx-editor.dev/core/editor';
import {
  buttonRoleProps,
  renderAsChild,
  slotPartNames,
  useSlotOverride,
  type HostElementProps,
} from './slot-override.ts';

/**
 * Prevent chrome mousedown from moving focus (inputs keep their behavior). Chrome building
 * block, not a stable API. @internal
 */
export function keepFocus(e: ReactMouseEvent): void {
  if (chromeKeepsFocus(e.target)) e.preventDefault();
}

/**
 * Give the caret back to the slide text after a toolbar field or picker is done, as DOCX
 * Editor's pickers call `editor.focus()`: a format armed at the caret applies to what is typed,
 * and the next key (Cmd+Z) reaches the slide. Chrome building block. @internal
 */
export function focusText(editor: { focus(): void } | null, from?: Element | null): void {
  focusTextAfterChrome(editor, from);
}

/** The props of `Icon`. @public */
export interface IconProps {
  /** A Material Symbol name from `ICON_PATHS`. */
  name?: string;
  /** SVG path data, in place of `name`. */
  paths?: readonly string[] | null;
  /** The class of the `svg` element. Defaults to `pptx-icon`. */
  className?: string;
}

/** An inline Material Symbol. @public */
export function Icon({ name, paths, className = 'pptx-icon' }: IconProps) {
  const known = name ? ICON_PATHS[name] : undefined;
  if (name && !paths && !known) warnUnknownIcon(name);
  const d = paths ?? known ?? [];
  return (
    <svg className={className} viewBox="0 -960 960 960" aria-hidden="true">
      {d.map((p, i) => (
        <path key={i} d={p} />
      ))}
    </svg>
  );
}

/**
 * `aria-disabled` and the spoken reason for a raw toolbar trigger (a picker, a field), the same
 * rule as `ToolButton`: it stays a Tab and arrow stop, ignores clicks,
 * and names its reason through `aria-describedby`. Render `note` beside the element. @internal
 */
export function useAriaDisabled(
  disabled: boolean,
  reason: string | null | undefined
): {
  attrs: {
    'aria-disabled'?: true;
    'aria-describedby'?: string;
    onClickCapture?: (e: { preventDefault(): void; stopPropagation(): void }) => void;
  };
  note: ReactNode;
} {
  const id = useId();
  const said = disabled && reason ? reason : null;
  return {
    attrs: disabled
      ? {
          'aria-disabled': true,
          ...(said ? { 'aria-describedby': id } : {}),
          onClickCapture: (e) => (e.preventDefault(), e.stopPropagation()),
        }
      : {},
    note: said ? (
      <span id={id} className="pptx-sr-only" hidden>
        {said}
      </span>
    ) : null,
  };
}

/** The props of `ToolButton`. @internal */
export interface ToolButtonProps {
  /** The accessible name and tooltip of the button. */
  label: string;
  /** SVG path data for the icon, in place of `icon`. */
  paths?: readonly string[] | null;
  /** A Material Symbol name from `ICON_PATHS`. */
  icon?: string;
  /** Called when the person activates the button. */
  onClick: () => void;
  /** The button cannot run. It stays focusable and names `disabledReason`. */
  disabled?: boolean;
  /** Why the button is disabled; read after its name. */
  disabledReason?: string | null;
  /** For a toggle: whether it is on (`aria-pressed`). */
  pressed?: boolean;
  /** For a button that shows or hides a region: whether it is shown (`aria-expanded`). */
  expanded?: boolean;
  /** The id of the region the button shows or hides (`aria-controls`). */
  controls?: string;
  /** Content after the icon. */
  children?: ReactNode;
  /** A class name added to the button. */
  className?: string;
  /** The chrome slot this button drives, as `data-slot` for hosts and tests. */
  slot?: string;
  /** The popup this button opens (`aria-haspopup`); `expanded` says whether it is open. */
  hasPopup?: 'menu' | 'grid' | 'dialog';
  /** False for the main half of a color split: a host element cannot stand in for it. */
  asChild?: boolean;
  /** A host element that renders the button, from a part's own `asChild` (no slot override). */
  host?: ReactElement<HostElementProps> | null;
}

/** A toolbar icon button with tooltip and disabled reason. Not a stable API. @internal */
export function ToolButton(props: ToolButtonProps) {
  // `Toolbar.Slot` overrides for this slot: a host icon, a class name, or a host element.
  const override = useSlotOverride(props.slot);
  const reasonId = useId();
  if (override?.asChild && props.asChild === false)
    warnUnsupportedOverride(props.slot ?? '', 'asChild', slotPartNames(override));
  // A disabled control stays focusable (`aria-disabled`), so a keyboard user reaches it in the
  // toolbar and hears its reason through `aria-describedby`, as DOCX Editor's menus and table
  // chrome do. The reason text is `hidden`: a description still reads it, and the accessibility
  // tree shows it once, on the button, not again as loose text in the toolbar.
  const reason = props.disabled && props.disabledReason ? props.disabledReason : null;
  const onClick = props.disabled ? undefined : props.onClick;
  const className = ['pptx-toolbar__button', props.className, override?.className]
    .filter(Boolean)
    .join(' ');
  const icon =
    override?.icon !== undefined ? (
      override.icon
    ) : props.paths || props.icon ? (
      <Icon paths={props.paths} name={props.icon} />
    ) : null;
  const attributes = {
    className,
    'aria-label': props.label,
    ...tooltipAttributes(props.label, props.disabled, props.disabledReason),
    'data-slot': props.slot,
    // State as data attributes, the first step of the customization ladder.
    'data-active': props.pressed ? '' : undefined,
    'data-disabled': props.disabled ? '' : undefined,
    'aria-pressed': props.pressed === undefined ? undefined : props.pressed,
    'aria-haspopup': props.hasPopup,
    'aria-expanded': props.hasPopup ? !!props.expanded : props.expanded,
    'aria-controls': props.controls,
    'aria-disabled': props.disabled ? true : undefined,
    'aria-describedby': reason ? reasonId : undefined,
    onMouseDown: keepFocus,
  };
  const described = reason ? (
    <span id={reasonId} className="pptx-sr-only" hidden>
      {reason}
    </span>
  ) : null;
  const hostElement =
    props.host ?? (override?.asChild && props.asChild !== false ? override.asChild : null);
  if (hostElement) {
    // The host element renders the control, after DOCX Editor's toolbar `Slot`: the props merge
    // into the child's (classes join, handlers run child-first). A host element that is not a
    // button gets `role="button"`, a tab stop, and Enter and Space, so the label and the
    // pressed state stay valid ARIA and the keyboard can run it. It
    // cannot be disabled natively, so it says so with `aria-disabled` and gets no handler.
    const isButton = hostElement.type === 'button';
    return (
      <>
        {renderAsChild(hostElement, {
          ...attributes,
          ...(isButton ? { type: 'button' } : buttonRoleProps(onClick)),
          onClick,
        })}
        {described}
      </>
    );
  }
  return (
    <>
      <button type="button" {...attributes} onClick={onClick}>
        {icon}
        {props.children}
      </button>
      {described}
    </>
  );
}

/** The props of `Popover`. @internal */
export interface PopoverProps {
  /** The element the panel opens under. */
  anchor: HTMLElement | null;
  /** Called when the panel closes: an outside click, Escape, or Tab. */
  onClose: () => void;
  /** The panel's content. */
  children: ReactNode;
  /** The edge of the anchor that the panel lines up with. Defaults to `start`. */
  align?: 'start' | 'end';
  /** The accessible name; defaults to the trigger's `aria-label`. */
  label?: string;
  /** `menu` (default) for rows, or `dialog` for a form or controls. */
  role?: 'menu' | 'dialog';
  /** The element id, for the trigger's `aria-controls`. */
  id?: string;
  /**
   * A dialog panel opened from the keyboard takes focus at its first control (a menu always
   * does), and Tab past its first or last control closes it and returns to the trigger.
   */
  focusFirst?: boolean;
  /**
   * A dialog with a field to type in (the font search): its field takes focus when it opens,
   * for pointer and keyboard users alike.
   */
  field?: boolean;
  /** A panel that takes no focus when it opens: the keys stay in the slide text. */
  passive?: boolean;
  /** A class name added to the panel. */
  className?: string;
  /**
   * A box in client pixels to open at instead of the anchor's own box (the link popup at the
   * selected text). Focus and dismissal still use `anchor`.
   */
  anchorRect?: PopoverAnchorRect | null;
}

/**
 * A popover anchored under an element, closed on outside click, Escape, or Tab.
 *
 * As a `menu` (the default) it follows the WAI-ARIA menu pattern, as DOCX Editor's menus do:
 * the arrow keys, Home, and End move through the rows (by grid rows in a swatch grid), and
 * every row is one `tabIndex=-1` stop. As a `dialog` it holds a form or controls (the font
 * search, the link field, the More card). With `field`, its field takes focus when it opens,
 * for pointer and keyboard users. Otherwise, when the trigger has focus at open (a keyboard
 * user), focus moves to the first row or control; after a pointer open the panel itself takes
 * focus (core's `focusPopoverOnOpen`), so typing and the arrow keys reach the popover, not the
 * slide text. Escape returns focus to the trigger. Another close gives focus back to where it
 * was before a pointer open (the slide text), else to the trigger, unless the popover moved it
 * on purpose. @internal
 */
export function Popover({
  anchor,
  onClose,
  children,
  align = 'start',
  label,
  role = 'menu',
  id,
  focusFirst = false,
  field = false,
  passive = false,
  className,
  anchorRect = null,
}: PopoverProps) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{
    left: number;
    top: number;
    maxHeight: number;
    scroll: boolean;
  } | null>(null);
  // Whether focus was inside the popover, so closing it gives focus back to the trigger
  // instead of dropping it on <body> when the popover unmounts.
  const focusInside = useRef(false);
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;
  // Before paint, and written to the node at once: React runs the focus effect below before
  // it re-renders with `pos`, and a browser does not focus into a `visibility: hidden` panel.
  useLayoutEffect(() => {
    if (!anchor) return;
    setPos(positionPopover(ref.current, anchor, align, anchorRect));
  }, [anchor, align, anchorRect]);
  // The element that had focus before the open, for the same trigger. StrictMode runs the
  // effect below twice: the second run finds focus already inside the panel and saves nothing,
  // so the first saved element is kept instead of falling back to the trigger on close.
  const back = useRef<{ anchor: HTMLElement | null; el: HTMLElement | null } | null>(null);
  useEffect(() => {
    // A pointer open moves focus into the panel too, so the keys reach the menu; closing gives
    // focus back to where it was (the slide text), and Escape to the trigger.
    const saved = focusPopoverOnOpen(ref.current, anchor, { role, focusFirst, field, passive });
    if (saved || back.current?.anchor !== anchor) back.current = { anchor, el: saved };
    return () => restorePopoverFocus(anchor, focusInside.current, back.current?.el ?? null);
  }, [anchor, role, focusFirst, field, passive]);
  useEffect(
    () =>
      listenPopoverDismiss({
        panel: () => ref.current,
        anchor,
        focusInside: () => focusInside.current,
        onClose: () => onCloseRef.current(),
      }),
    [anchor]
  );
  // The menu inherits tokens from the nearest editor scope by rendering inside it.
  const scope = popoverScope(anchor);
  return createPortal(
    <div
      ref={ref}
      id={id}
      className={className ? `pptx-menu ${className}` : 'pptx-menu'}
      role={role}
      aria-label={label ?? anchor?.getAttribute('aria-label') ?? undefined}
      tabIndex={-1}
      style={
        pos
          ? {
              left: pos.left,
              top: pos.top,
              maxHeight: pos.maxHeight,
              overflowY: pos.scroll ? 'auto' : undefined,
            }
          : { visibility: 'hidden', left: 0, top: 0 }
      }
      onMouseDown={keepFocus}
      // In a menu the pointer moves the focused row, so one row looks active.
      onMouseMove={role === 'menu' ? (e) => focusMenuRowAtPointer(ref.current, e) : undefined}
      onFocus={() => {
        focusInside.current = true;
      }}
      onKeyDown={(e) =>
        // Keys from a popover opened inside this one (React events bubble through portals) are
        // left alone.
        popoverKeyDown(ref.current, anchor, e, {
          role,
          focusFirst,
          onClose: () => onCloseRef.current(),
        })
      }
    >
      {children}
    </div>,
    scope
  );
}

/** The props of `MenuItem`. @internal */
export interface MenuItemProps {
  /** The row's content. */
  children: ReactNode;
  /** Called when the person activates the row. */
  onSelect: () => void;
  /** For a choice row: whether it is the current choice (`aria-checked`). */
  checked?: boolean;
  /** The row cannot run. */
  disabled?: boolean;
  /**
   * Why a disabled row cannot run. The row then stays focusable (`aria-disabled`) and names
   * the reason, so a keyboard or screen-reader user meets it too; it still does not run.
   */
  reason?: string;
  /** The row's tooltip. */
  title?: string;
  /** A row of a `listbox` (the font list) instead of a menu. */
  option?: boolean;
  /** An on/off setting: `menuitemcheckbox`. */
  checkbox?: boolean;
  /**
   * A hint read after the row's name (`aria-description`), so the name stays the label alone:
   * mark a visible hint in the row `aria-hidden`.
   */
  description?: string;
}

/**
 * A menu row: `menuitem`, or `menuitemradio` when `checked` is given (`menuitemcheckbox` with
 * `checkbox`, for an on/off setting), or `option` in a list box. Rows are `tabIndex=-1`: the
 * arrow keys move between them. @internal
 */
export function MenuItem({
  children,
  onSelect,
  checked,
  disabled,
  reason,
  title,
  option,
  checkbox,
  description,
}: MenuItemProps) {
  const state = option
    ? { role: 'option', 'aria-selected': checked ?? false }
    : checked === undefined
      ? { role: 'menuitem' }
      : { role: checkbox ? 'menuitemcheckbox' : 'menuitemradio', 'aria-checked': checked };
  return (
    <button
      type="button"
      {...state}
      tabIndex={-1}
      className="pptx-menu__item"
      // A disabled row stays focusable, as `MenuRow` does, so the arrow keys reach it and a
      // keyboard or screen-reader user can read why; a `reason` is spoken and shown.
      aria-disabled={disabled || undefined}
      {...(disabled && reason
        ? { 'aria-description': reason, 'data-pptx-tooltip': reason }
        : title
          ? { 'data-pptx-tooltip': title }
          : {})}
      {...(description && !(disabled && reason) ? { 'aria-description': description } : {})}
      onClick={disabled ? undefined : onSelect}
    >
      {children}
      {/* The checked row of a menu shows a check mark, as DOCX Editor's menus and the context
          menu do. */}
      {checked && !option ? <Icon name="check" className="pptx-icon pptx-menu__check" /> : null}
    </button>
  );
}
