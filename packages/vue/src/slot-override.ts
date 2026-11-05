// A part-level injection for one toolbar slot's host overrides (`ToolbarSlot`): the icon, the
// class name, and an `asChild` element. The Vue twin of the React adapter's `slot-override.ts`:
// every control reads it for its own slot, so a customization reaches the control without new
// props, as in DOCX Editor's toolbar parts.

import { cloneVNode, inject, type InjectionKey, type VNode, type VNodeChild } from 'vue';
import { warnUnsupportedOverride, type ToolbarPartNames } from '@pptx-editor.dev/core/editor';

/** The overrides one `ToolbarSlot` passes to the control it wraps. @internal */
export interface SlotOverride {
  readonly slotId: string;
  /** Icon content in place of the registry icon. */
  readonly icon?: VNodeChild;
  readonly className?: string;
  /** With `asChild`, the host element that the control renders as. */
  readonly asChild?: VNode;
  /** The named part the override came from, for warning text. */
  readonly part?: string;
}

/** Provided by a `ToolbarSlot` that overrides its control. @internal */
export const slotOverrideKey: InjectionKey<SlotOverride | null> = Symbol('pptxSlotOverride');

/** The overrides for `slot`, or null when the nearest override names another slot. @internal */
export function useSlotOverride(slot: string | undefined): SlotOverride | null {
  const override = inject(slotOverrideKey, null);
  return override && slot !== undefined && override.slotId === slot ? override : null;
}

/** The names warnings use for a slot: the named part when there is one. @internal */
export function slotPartNames(override: SlotOverride | null): ToolbarPartNames | undefined {
  return override?.part ? { slot: override.part, group: 'ToolbarGroup' } : undefined;
}

/** State and class attributes for a control's trigger element. @internal */
export interface TriggerAttributes {
  readonly class: string;
  readonly 'data-active'?: '';
  readonly 'data-disabled'?: '';
}

/**
 * The class and state attributes for a control that renders a picker trigger, not a
 * `ToolButton`: `icon` and `asChild` are refused with a development warning, as in React.
 * @internal
 */
export function useTriggerAttributes(
  slot: string,
  base: string,
  state: () => { readonly disabled?: boolean; readonly active?: boolean }
): () => TriggerAttributes {
  const override = useSlotOverride(slot);
  if (override?.icon !== undefined) warnUnsupportedOverride(slot, 'icon', slotPartNames(override));
  if (override?.asChild) warnUnsupportedOverride(slot, 'asChild', slotPartNames(override));
  return () => {
    const now = state();
    return {
      class: [base, override?.className].filter(Boolean).join(' '),
      ...(now.active ? { 'data-active': '' as const } : {}),
      ...(now.disabled ? { 'data-disabled': '' as const } : {}),
    };
  };
}

/**
 * The props that make a host element that is not a `<button>` act as one (WCAG 2.1.1): the
 * button role, a tab stop, and Enter and Space to activate. A link keeps its own Enter.
 * @internal
 */
export function buttonRoleProps(
  activate: ((target: Element) => void) | undefined
): Record<string, unknown> {
  return {
    role: 'button',
    tabindex: 0,
    onKeydown: (event: KeyboardEvent) => {
      const target = event.currentTarget as Element;
      const link = target instanceof HTMLAnchorElement && target.href;
      if (event.key !== ' ' && (event.key !== 'Enter' || link)) return;
      event.preventDefault();
      activate?.(target);
    },
  };
}

function composeHandlers(childHandler: unknown, slotHandler: unknown): unknown {
  if (typeof slotHandler !== 'function') return childHandler;
  if (typeof childHandler !== 'function') return slotHandler;
  return (event: { defaultPrevented: boolean }) => {
    childHandler(event);
    if (!event.defaultPrevented) slotHandler(event);
  };
}

/**
 * Render `child` with the control's props merged in, as React's `renderAsChild`: classes join,
 * styles merge (child wins per property), handlers run child-first (the control's handler is
 * skipped when the child called `preventDefault()`), and other props keep the child's value.
 * @internal
 */
export function renderAsChild(child: VNode, slotProps: Record<string, unknown>): VNode {
  const childProps = (child.props ?? {}) as Record<string, unknown>;
  const merged: Record<string, unknown> = { ...slotProps, ...childProps };
  for (const key of Object.keys(slotProps)) {
    if (/^on[A-Z]/.test(key)) merged[key] = composeHandlers(childProps[key], slotProps[key]);
    else if (key === 'class') merged[key] = [slotProps[key], childProps[key]];
    else if (key === 'style') merged[key] = [slotProps[key], childProps[key]];
  }
  // `cloneVNode` without merge: `merged` already holds the child's own props.
  return cloneVNode({ ...child, props: null } as VNode, merged);
}
