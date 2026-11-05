/** @jsxImportSource vue */
// Chrome building blocks: icons, buttons, and popover menus shared by the toolbar, rail, notes,
// panel, and zoom pill. The Vue twin of the React adapter's `chrome.tsx`: the same markup,
// classes, ARIA, and keys, with the behavior in core (`chrome-popover.ts`), so the two cannot
// drift. Every mousedown on chrome is prevented (except form fields) so it never steals the
// slide selection or the text caret.

import {
  Teleport,
  computed,
  defineComponent,
  normalizeClass,
  onMounted,
  onUnmounted,
  ref,
  shallowRef,
  toValue,
  useId,
  watch,
  type ComputedRef,
  type MaybeRefOrGetter,
  type PropType,
  type VNode,
  type ExtractPublicPropTypes,
} from 'vue';
import {
  ICON_PATHS,
  chromeKeepsFocus,
  tooltipAttributes,
  focusMenuRowAtPointer,
  focusPopoverOnOpen,
  focusTextAfterChrome,
  listenPopoverDismiss,
  popoverKeyDown,
  popoverScope,
  positionPopover,
  restorePopoverFocus,
  warnUnknownIcon,
  warnUnsupportedOverride,
  type PopoverPlacement,
} from '@pptx-editor.dev/core/editor';
import { buttonRoleProps, renderAsChild, slotPartNames, useSlotOverride } from './slot-override.ts';

/**
 * Prevent chrome mousedown from moving focus (inputs keep their behavior). Chrome building
 * block, not a stable API. @internal
 */
export function keepFocus(e: MouseEvent): void {
  if (chromeKeepsFocus(e.target)) e.preventDefault();
}

/**
 * Give the caret back to the slide text after a toolbar field or picker is done, as React's
 * `focusText`. Chrome building block. @internal
 */
export function focusText(editor: { focus(): void } | null, from?: Element | null): void {
  focusTextAfterChrome(editor, from);
}

const iconProps = {
  /** A Material Symbol name from `ICON_PATHS`. */
  name: { type: String, default: undefined },
  /** SVG path data, in place of `name`. */
  paths: { type: Array as PropType<readonly string[] | null>, default: undefined },
  className: { type: String, default: 'pptx-icon' },
} as const;

/** The props of `Icon`, the twin of React's `IconProps`. @public */
export type IconProps = ExtractPublicPropTypes<typeof iconProps>;

/** An inline Material Symbol. @public */
export const Icon = defineComponent({
  name: 'Icon',
  props: iconProps,
  setup(props) {
    return () => {
      const known = props.name ? ICON_PATHS[props.name] : undefined;
      if (props.name && !props.paths && !known) warnUnknownIcon(props.name);
      const d = props.paths ?? known ?? [];
      return (
        <svg class={props.className} viewBox="0 -960 960 960" aria-hidden="true">
          {d.map((p, i) => (
            <path key={i} d={p} />
          ))}
        </svg>
      );
    };
  },
});

/** What `useAriaDisabled` returns. @internal */
export interface UseAriaDisabledReturn {
  /** The attributes for the trigger element. */
  readonly attrs: ComputedRef<{
    'aria-disabled'?: true;
    'aria-describedby'?: string;
    onClickCapture?: (e: Event) => void;
  }>;
  /** The spoken reason, to render beside the element. */
  readonly note: () => VNode | null;
}

/**
 * `aria-disabled` and the spoken reason for a raw toolbar trigger (a picker, a field), the same
 * rule as `ToolButton`: it stays a Tab and arrow stop, ignores clicks,
 * and names its reason through `aria-describedby`. Render `note()` beside the element.
 * @internal
 */
export function useAriaDisabled(
  disabled: MaybeRefOrGetter<boolean>,
  reason: MaybeRefOrGetter<string | null | undefined>
): UseAriaDisabledReturn {
  const id = useId();
  const said = computed(() => (toValue(disabled) && toValue(reason) ? toValue(reason)! : null));
  return {
    attrs: computed(() =>
      toValue(disabled)
        ? {
            'aria-disabled': true as const,
            ...(said.value ? { 'aria-describedby': id } : {}),
            onClickCapture: (e: Event) => (e.preventDefault(), e.stopPropagation()),
          }
        : {}
    ),
    note: () =>
      said.value ? (
        <span id={id} class="pptx-sr-only" hidden>
          {said.value}
        </span>
      ) : null,
  };
}

const toolButtonProps = {
  label: { type: String, required: true as const },
  paths: { type: Array as PropType<readonly string[] | null>, default: undefined },
  icon: { type: String, default: undefined },
  onClick: { type: Function as PropType<() => void>, default: undefined },
  disabled: { type: Boolean, default: false },
  disabledReason: { type: String as PropType<string | null>, default: undefined },
  pressed: { type: Boolean, default: undefined },
  /** For a button that shows or hides a region: whether it is shown (`aria-expanded`). */
  expanded: { type: Boolean, default: undefined },
  /** The id of the region the button shows or hides (`aria-controls`). */
  controls: { type: String, default: undefined },
  className: { type: String, default: undefined },
  /** The chrome slot this button drives, as `data-slot` for hosts and tests. */
  slot: { type: String, default: undefined },
  /** The popup this button opens (`aria-haspopup`); `expanded` says whether it is open. */
  hasPopup: { type: String as PropType<'menu' | 'grid' | 'dialog'>, default: undefined },
  /** False for the main half of a color split: a host element cannot stand in for it. */
  asChild: { type: Boolean, default: undefined },
} as const;

/** The props of `ToolButton`, the twin of React's `ToolButtonProps`. @internal */
export type ToolButtonProps = ExtractPublicPropTypes<typeof toolButtonProps>;

/**
 * A toolbar icon button with tooltip and disabled reason. The default slot adds content after
 * the icon. Not a stable API. @internal
 */
export const ToolButton = defineComponent({
  name: 'ToolButton',
  // Two root nodes (the button and its reason text): a host `class` and other attributes go on
  // the button, with `className`, as React's.
  inheritAttrs: false,
  props: toolButtonProps,
  setup(props, { slots, attrs }) {
    // `ToolbarSlot` overrides for this slot: a host icon, a class name, or a host element.
    const override = useSlotOverride(props.slot);
    const reasonId = useId();
    if (override?.asChild && props.asChild === false)
      warnUnsupportedOverride(props.slot ?? '', 'asChild', slotPartNames(override));
    return () => {
      // A disabled control stays focusable (`aria-disabled`), so a keyboard user reaches it in
      // the toolbar and hears its reason through `aria-describedby`. The reason text is
      // `hidden`: a description still reads it, and the tree shows it once, on the button.
      const reason = props.disabled && props.disabledReason ? props.disabledReason : null;
      const onClick = props.disabled ? undefined : props.onClick;
      const icon =
        override?.icon !== undefined ? (
          override.icon
        ) : props.paths || props.icon ? (
          <Icon paths={props.paths} name={props.icon} />
        ) : null;
      const { class: hostClass, ...hostAttrs } = attrs;
      const attributes = {
        ...hostAttrs,
        class: normalizeClass([
          'pptx-toolbar__button',
          hostClass,
          props.className,
          override?.className,
        ]),
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
        onMousedown: keepFocus,
      };
      const described = reason ? (
        <span id={reasonId} class="pptx-sr-only" hidden>
          {reason}
        </span>
      ) : null;
      if (override?.asChild && props.asChild !== false) {
        // The host element renders the control, as React's `renderAsChild`: a host element that
        // is not a button gets `role="button"`, a tab stop, and Enter and Space.
        const isButton = override.asChild.type === 'button';
        return [
          renderAsChild(override.asChild, {
            ...attributes,
            ...(isButton ? { type: 'button' } : buttonRoleProps(onClick)),
            onClick,
          }),
          described,
        ];
      }
      return [
        <button type="button" {...attributes} onClick={onClick}>
          {icon}
          {slots.default?.()}
        </button>,
        described,
      ];
    };
  },
});

const popoverProps = {
  // No runtime type check: an element from another window (an iframe) is not `Object`.
  anchor: { type: null as unknown as PropType<HTMLElement | null>, default: null },
  onClose: { type: Function as PropType<() => void>, required: true as const },
  align: { type: String as PropType<'start' | 'end'>, default: 'start' },
  /** The accessible name; defaults to the trigger's `aria-label`. */
  label: { type: String, default: undefined },
  role: { type: String as PropType<'menu' | 'dialog'>, default: 'menu' },
  /** The element id, for the trigger's `aria-controls`. */
  id: { type: String, default: undefined },
  /**
   * A dialog panel opened from the keyboard takes focus at its first control (a menu always
   * does), and Tab past its first or last control closes it and returns to the trigger.
   */
  focusFirst: { type: Boolean, default: false },
  /** A dialog with a field to type in: its field takes focus when it opens. */
  field: { type: Boolean, default: false },
  /** A panel that takes no focus when it opens: the keys stay in the slide text. */
  passive: { type: Boolean, default: false },
  /** A class name added to the panel. */
  className: { type: String, default: undefined },
} as const;

/** The props of `Popover`, the twin of React's `PopoverProps`. @internal */
export type PopoverProps = ExtractPublicPropTypes<typeof popoverProps>;

/**
 * A popover anchored under an element, closed on outside click, Escape, or Tab, as React's
 * `Popover`. As a `menu` (the default) it follows the WAI-ARIA menu pattern; as a `dialog` it
 * holds a form or controls. It renders in the anchor's editor scope, so the tokens apply.
 * @internal
 */
export const Popover = defineComponent({
  name: 'Popover',
  // The root is a teleport: a host `class` and other attributes go on the panel, with
  // `className`, as React's.
  inheritAttrs: false,
  props: popoverProps,
  setup(props, { slots, attrs }) {
    const panel = ref<HTMLDivElement | null>(null);
    const pos = shallowRef<PopoverPlacement | null>(null);
    // Whether focus was inside the popover, so closing it gives focus back to the trigger.
    let focusInside = false;
    const close = (): void => props.onClose();
    const place = (): void => {
      if (props.anchor) pos.value = positionPopover(panel.value, props.anchor, props.align);
    };
    let stopListening: (() => void) | null = null;
    const listen = (anchor: HTMLElement | null): void => {
      stopListening?.();
      stopListening = listenPopoverDismiss({
        panel: () => panel.value,
        anchor,
        focusInside: () => focusInside,
        onClose: close,
      });
    };
    // The element that had focus before a pointer open: closing gives focus back to it.
    let back: HTMLElement | null = null;
    onMounted(() => {
      // Placed before focus moves in: a browser does not focus into a hidden panel.
      place();
      back = focusPopoverOnOpen(panel.value, props.anchor, props);
      listen(props.anchor);
    });
    watch(
      () => [props.anchor, props.align] as const,
      ([anchor], [old]) => {
        place();
        if (anchor === old) return;
        restorePopoverFocus(old, focusInside, back);
        back = focusPopoverOnOpen(panel.value, anchor, props);
        listen(anchor);
      },
      { flush: 'post' }
    );
    // After the panel left the DOM, as React's effect cleanup: focus inside it has dropped to
    // `<body>`, and goes back to where it was before a pointer open, else to the trigger.
    onUnmounted(() => {
      stopListening?.();
      restorePopoverFocus(props.anchor, focusInside, back);
    });
    return () => {
      const p = pos.value;
      return (
        <Teleport to={popoverScope(props.anchor)}>
          <div
            {...attrs}
            ref={panel}
            id={props.id}
            class={normalizeClass(['pptx-menu', attrs.class, props.className])}
            role={props.role}
            aria-label={props.label ?? props.anchor?.getAttribute('aria-label') ?? undefined}
            tabindex={-1}
            style={
              p
                ? {
                    left: `${p.left}px`,
                    top: `${p.top}px`,
                    maxHeight: `${p.maxHeight}px`,
                    overflowY: p.scroll ? 'auto' : undefined,
                  }
                : { visibility: 'hidden', left: '0px', top: '0px' }
            }
            onMousedown={keepFocus}
            // In a menu the pointer moves the focused row, so one row looks active.
            onMousemove={
              props.role === 'menu'
                ? (e: MouseEvent) => focusMenuRowAtPointer(panel.value, e)
                : undefined
            }
            onFocusin={() => {
              focusInside = true;
            }}
            onKeydown={(e: KeyboardEvent) =>
              popoverKeyDown(panel.value, props.anchor, e, { ...props, onClose: close })
            }
          >
            {slots.default?.()}
          </div>
        </Teleport>
      );
    };
  },
});

const menuItemProps = {
  onSelect: { type: Function as PropType<() => void>, required: true as const },
  checked: { type: Boolean, default: undefined },
  disabled: { type: Boolean, default: false },
  /**
   * Why a disabled row cannot run. The row then stays focusable (`aria-disabled`) and names
   * the reason, so a keyboard or screen-reader user meets it too; it still does not run.
   */
  reason: { type: String, default: undefined },
  title: { type: String, default: undefined },
  /** A row of a `listbox` (the font list) instead of a menu. */
  option: { type: Boolean, default: false },
  /** An on/off setting: `menuitemcheckbox`. */
  checkbox: { type: Boolean, default: false },
  /**
   * A hint read after the row's name (`aria-description`), so the name stays the label alone:
   * mark a visible hint in the row `aria-hidden`.
   */
  description: { type: String, default: undefined },
} as const;

/** The props of `MenuItem`, the twin of React's `MenuItemProps`. @internal */
export type MenuItemProps = ExtractPublicPropTypes<typeof menuItemProps>;

/**
 * A menu row: `menuitem`, or `menuitemradio` when `checked` is given (`menuitemcheckbox` with
 * `checkbox`), or `option` in a list box. Rows are `tabindex=-1`: the arrow keys move between
 * them. The default slot is the row's content. @internal
 */
export const MenuItem = defineComponent({
  name: 'MenuItem',
  props: menuItemProps,
  setup(props, { slots }) {
    return () => {
      const state = props.option
        ? { role: 'option', 'aria-selected': props.checked ?? false }
        : props.checked === undefined
          ? { role: 'menuitem' }
          : {
              role: props.checkbox ? 'menuitemcheckbox' : 'menuitemradio',
              'aria-checked': props.checked,
            };
      return (
        <button
          type="button"
          {...state}
          tabindex={-1}
          class="pptx-menu__item"
          // A disabled row stays focusable, so the arrow keys reach it and a keyboard or
          // screen-reader user can read why; a `reason` is spoken and shown.
          aria-disabled={props.disabled || undefined}
          {...(props.disabled && props.reason
            ? { 'aria-description': props.reason, 'data-pptx-tooltip': props.reason }
            : props.title
              ? { 'data-pptx-tooltip': props.title }
              : {})}
          {...(props.description && !(props.disabled && props.reason)
            ? { 'aria-description': props.description }
            : {})}
          onClick={props.disabled ? undefined : props.onSelect}
        >
          {slots.default?.()}
          {/* The checked row of a menu shows a check mark. */}
          {props.checked && !props.option ? (
            <Icon name="check" className="pptx-icon pptx-menu__check" />
          ) : null}
        </button>
      );
    };
  },
});
