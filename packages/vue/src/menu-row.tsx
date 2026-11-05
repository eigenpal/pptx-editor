/** @jsxImportSource vue */
// Menu rows with an icon column, a label, and a shortcut hint, and the cascading submenu: the
// Vue twin of `MenuRow`, `MenuSeparator`, and `MenuSubmenu` in the React adapter's
// `ContextMenu.tsx`, with the same markup and classes. A disabled row stays in the menu with
// the engine's reason as its tooltip. The rows share the popover menu's keyboard rules
// (`menu-keyboard.ts`), and the submenu's place and keys come from core (`chrome-popover.ts`).

import {
  Teleport,
  defineComponent,
  inject,
  nextTick,
  normalizeClass,
  onBeforeUnmount,
  ref,
  shallowRef,
  watch,
  type ComponentOptionsMixin,
  type DefineComponent,
  type ExtractPublicPropTypes,
  type InjectionKey,
  type PropType,
  type VNode,
  type VNodeChild,
} from 'vue';
import { sanitizeHref } from '@pptx-editor.dev/core/store';
import {
  tooltipAttributes,
  focusEdge,
  focusMenuRowAtPointer,
  panelItems,
  placeSubmenu,
  popoverScope,
  submenuPanelKey,
  translateKeyOrText,
} from '@pptx-editor.dev/core/editor';
import { Icon, keepFocus } from './chrome.tsx';
import { useTranslation } from './composables.ts';
import type { JSX } from 'vue/jsx-runtime';

/** Closes the menu a row is in; `null` outside a packaged menu. @internal */
export const menuCloseKey: InjectionKey<(() => void) | null> = Symbol('pptxMenuClose');

/** The icon column of a row: other content, a named or drawn icon, or icon content. */
function rowIcon(
  glyph: VNodeChild | undefined,
  icon: string | VNodeChild | undefined,
  paths: readonly string[] | null | undefined
): VNodeChild {
  if (glyph !== undefined && glyph !== null) return glyph;
  if (typeof icon === 'string' || paths)
    return <Icon name={typeof icon === 'string' ? icon : undefined} paths={paths} />;
  return icon ?? null;
}

const menuRowProps = {
  label: { type: String, required: true as const },
  /** A Material Symbol name from `ICON_PATHS`, or icon content (a vnode). */
  icon: { type: [String, Object] as PropType<string | VNode>, default: undefined },
  paths: { type: Array as PropType<readonly string[] | null>, default: undefined },
  /** Other icon content (a drawn shape or line), in place of `icon` and `paths`. */
  glyph: { type: Object as PropType<VNode>, default: undefined },
  /** The shortcut hint, already in this platform's form (`useTranslation` gives it). */
  shortcut: { type: String, default: undefined },
  disabled: { type: Boolean, default: false },
  /** Why the row is disabled; shown as its tooltip. */
  reason: { type: String as PropType<string | null>, default: undefined },
  /** The same as `reason`, under DOCX Editor's and `ToolbarAction`'s name. */
  disabledReason: { type: String as PropType<string | null>, default: undefined },
  /** A toggle that is on (Bold on bold text). */
  active: { type: Boolean, default: false },
  /** An on/off setting: a `menuitemcheckbox` with a check when on. */
  checked: { type: Boolean, default: undefined },
  /** With `checked`, one choice of several: a `menuitemradio` instead. */
  radio: { type: Boolean, default: false },
  /** The chrome slot or command the row runs, as `data-slot` for hosts and tests. */
  slot: { type: String, default: undefined },
  /** A class name added to the row. */
  className: { type: String, default: undefined },
  /**
   * A web link: the row opens it in a new tab. Only `http`, `https`, `mailto`, `tel`, and `ftp`
   * links pass `sanitizeHref`; any other link disables the row.
   */
  href: { type: String, default: undefined },
} as const;

/** Props of `MenuRow`, the twin of React's `MenuRowProps`. @public */
export type MenuRowProps = ExtractPublicPropTypes<typeof menuRowProps>;

const menuRow = defineComponent({
  name: 'MenuRow',
  props: menuRowProps,
  emits: { select: () => true },
  setup(props, { emit }) {
    const close = inject(menuCloseKey, null);
    return () => {
      const { href } = props;
      const link = href === undefined ? null : sanitizeHref(href);
      const disabled = props.disabled || (link !== null && !link.ok);
      const reason = props.reason ?? props.disabledReason;
      const shared = {
        role:
          props.checked === undefined
            ? 'menuitem'
            : props.radio
              ? 'menuitemradio'
              : 'menuitemcheckbox',
        ...(props.checked === undefined ? {} : { 'aria-checked': props.checked }),
        tabindex: -1,
        class: ['pptx-menu__item pptx-menu__row', props.className].filter(Boolean).join(' '),
        'aria-disabled': disabled || undefined,
        // The row shows its name; the tooltip adds only the reason, which is also its description.
        ...(disabled && reason ? { 'aria-description': reason, 'data-pptx-tooltip': reason } : {}),
        'data-slot': props.slot,
        'data-active': props.active ? '' : undefined,
        'data-disabled': disabled ? '' : undefined,
        onClick: () => {
          if (disabled) return;
          emit('select');
          close?.();
        },
      };
      const content = [
        <span class="pptx-menu__icon" aria-hidden="true">
          {rowIcon(props.glyph, props.icon, props.paths)}
        </span>,
        <span class="pptx-menu__label">{props.label}</span>,
        props.checked ? <Icon name="check" className="pptx-icon pptx-menu__check" /> : null,
        // The shortcut is a visual hint; the row's accessible name is its label.
        props.shortcut ? (
          <span class="pptx-menu__shortcut" aria-hidden="true">
            {props.shortcut}
          </span>
        ) : null,
      ];
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
    };
  },
});

// An explicit type: the inferred one lists the defaults of `reason` and `disabledReason` in an
// order that changes between builds, and the API report must not change with it.
/**
 * A menu row: icon, label, and shortcut. A disabled row stays focusable with `aria-disabled`,
 * so a keyboard or pointer user can read why. Emits `select`; inside a packaged menu, the menu
 * closes after. @public
 */
export const MenuRow: DefineComponent<
  typeof menuRowProps,
  () => JSX.Element,
  {},
  {},
  {},
  ComponentOptionsMixin,
  ComponentOptionsMixin,
  { select: () => true }
> = menuRow;

/** A line between groups of rows. @public */
export const MenuSeparator = defineComponent({
  name: 'MenuSeparator',
  setup() {
    return () => <div class="pptx-menu__sep" role="separator" />;
  },
});

const menuSubmenuProps = {
  /** The row's label. Wins over `labelKey`. */
  label: { type: String, default: undefined },
  /** A translation key for the label, as DOCX Editor's `Menu.Submenu` takes. */
  labelKey: { type: String, default: undefined },
  /** A Material Symbol name, or icon content (a vnode). */
  icon: { type: [String, Object] as PropType<string | VNode>, default: undefined },
  paths: { type: Array as PropType<readonly string[] | null>, default: undefined },
  glyph: { type: Object as PropType<VNode>, default: undefined },
  disabled: { type: Boolean, default: false },
  reason: { type: String as PropType<string | null>, default: undefined },
  /** The submenu's own class, for a grid of shapes. */
  className: { type: String, default: undefined },
} as const;

/** Props of `MenuSubmenu`, the twin of React's `MenuSubmenuProps`. @public */
export type MenuSubmenuProps = ExtractPublicPropTypes<typeof menuSubmenuProps>;

const menuSubmenu = defineComponent({
  name: 'MenuSubmenu',
  // The row and a teleported panel are two root nodes: a host `class` goes on the panel, with
  // `className`, as React's; other attributes go on the row.
  inheritAttrs: false,
  props: menuSubmenuProps,
  setup(props, { slots, attrs }) {
    const { t } = useTranslation();
    const row = ref<HTMLButtonElement | null>(null);
    const panel = ref<HTMLDivElement | null>(null);
    const open = ref<false | 'pointer' | 'keyboard'>(false);
    const pos = shallowRef<{ left: number; top: number } | null>(null);
    const close = (focusRow: boolean): void => {
      open.value = false;
      pos.value = null;
      if (focusRow) row.value?.focus();
    };
    // Place the panel once it renders, then move focus in: a hidden panel cannot take focus.
    watch(
      open,
      (now) => {
        const r = row.value?.getBoundingClientRect();
        const p = panel.value;
        if (!now || !r || !p) return;
        pos.value = placeSubmenu(
          r,
          { width: p.offsetWidth, height: p.offsetHeight },
          { width: window.innerWidth, height: window.innerHeight }
        );
      },
      { flush: 'post' }
    );
    watch(
      [pos, open],
      ([placed, now]) => {
        if (!placed || now !== 'keyboard') return;
        // This watcher runs before the render that clears `visibility: hidden`, so wait for
        // that render, as React's effect runs after its commit.
        void nextTick(() => {
          if (pos.value && open.value === 'keyboard' && panel.value)
            focusEdge(panelItems(panel.value), 'first');
        });
      },
      { flush: 'post' }
    );
    // Hovering another row of the parent menu closes this one.
    const over = (e: Event): void => {
      const target = e.target as Node;
      const parent = row.value?.closest('.pptx-menu');
      if (!parent?.contains(target) || row.value?.contains(target)) return;
      close(false);
    };
    watch(open, (now, _old, onCleanup) => {
      if (!now) return;
      document.addEventListener('mouseover', over);
      onCleanup(() => document.removeEventListener('mouseover', over));
    });
    onBeforeUnmount(() => document.removeEventListener('mouseover', over));
    const onRowKey = (e: KeyboardEvent): void => {
      if (props.disabled || (e.key !== 'ArrowRight' && e.key !== 'Enter' && e.key !== ' ')) return;
      e.preventDefault();
      e.stopPropagation();
      open.value = 'keyboard';
    };
    return () => {
      const label = props.label ?? (props.labelKey ? translateKeyOrText(t, props.labelKey) : '');
      const { class: hostClass, ...hostAttrs } = attrs;
      return [
        <button
          {...hostAttrs}
          ref={row}
          type="button"
          role="menuitem"
          tabindex={-1}
          class="pptx-menu__item pptx-menu__row"
          aria-haspopup="menu"
          aria-expanded={!!open.value}
          aria-disabled={props.disabled || undefined}
          {...(props.disabled && props.reason
            ? { 'aria-description': props.reason, ...tooltipAttributes(label, true, props.reason) }
            : {})}
          data-open={open.value ? '' : undefined}
          onMouseenter={() => {
            if (!props.disabled) open.value = open.value || 'pointer';
          }}
          onClick={() => {
            if (!props.disabled) open.value = 'keyboard';
          }}
          onKeydown={onRowKey}
        >
          <span class="pptx-menu__icon" aria-hidden="true">
            {rowIcon(props.glyph, props.icon, props.paths)}
          </span>
          <span class="pptx-menu__label">{label}</span>
          <Icon name="chevron_right" className="pptx-icon pptx-menu__caret" />
        </button>,
        open.value ? (
          <Teleport to={popoverScope(row.value)}>
            <div
              ref={panel}
              class={normalizeClass(['pptx-menu pptx-menu--submenu', hostClass, props.className])}
              role="menu"
              aria-label={label}
              style={
                pos.value
                  ? { left: `${pos.value.left}px`, top: `${pos.value.top}px` }
                  : { visibility: 'hidden', left: '0px', top: '0px' }
              }
              onMousedown={keepFocus}
              onMousemove={(e: MouseEvent) => focusMenuRowAtPointer(panel.value, e)}
              onKeydown={(e: KeyboardEvent) => {
                if (submenuPanelKey(panel.value, e)) close(true);
              }}
            >
              {slots.default?.()}
            </div>
          </Teleport>
        ) : null,
      ];
    };
  },
});

// An explicit type, as `MenuRow`'s, so the API report does not change between builds.
/**
 * A row that opens a cascading panel to its right (to its left at the window's edge). It opens
 * on hover, click, Enter, and ArrowRight, and closes on ArrowLeft and Escape, which give focus
 * back to the row. The default slot is the panel's rows. @public
 */
export const MenuSubmenu: DefineComponent<
  typeof menuSubmenuProps,
  () => (JSX.Element | null)[],
  {},
  {},
  {},
  ComponentOptionsMixin,
  ComponentOptionsMixin,
  {}
> = menuSubmenu;
