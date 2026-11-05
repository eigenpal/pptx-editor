// A part-level context for one toolbar slot's host overrides (`Toolbar.Slot`): the icon, the
// class name, and an `asChild` element. Every control reads it for its own slot, so a
// customization reaches the control without new props, as in DOCX Editor's toolbar parts.

import { cloneElement, createContext, useContext, type ReactElement, type ReactNode } from 'react';
import { warnUnsupportedOverride, type ToolbarPartNames } from '@pptx-editor.dev/core/editor';

/** Props of an element a host passes with `asChild`. */
export type HostElementProps = Record<string, unknown>;

/** The part a named part's props came from (`Toolbar.Bold`), for warning text. @internal */
export const NAMED_PART = Symbol('pptxNamedPart');

/** The overrides one `Toolbar.Slot` passes to the control it wraps. */
export interface SlotOverride {
  readonly slotId: string;
  readonly icon?: ReactNode;
  readonly className?: string;
  /** With `asChild`, the host element that the control renders as. */
  readonly asChild?: ReactElement<HostElementProps>;
  /** The named part (`Toolbar.Bold`) the override came from, for warning text. */
  readonly [NAMED_PART]?: string;
}

export const SlotOverrideContext = createContext<SlotOverride | null>(null);

/** The overrides for `slot`, or null when the nearest override names another slot. */
export function useSlotOverride(slot: string | undefined): SlotOverride | null {
  const override = useContext(SlotOverrideContext);
  return override && slot !== undefined && override.slotId === slot ? override : null;
}

/** State and class attributes for a control's trigger element. */
export interface TriggerAttributes {
  readonly className: string;
  readonly 'data-active'?: '';
  readonly 'data-disabled'?: '';
}

/**
 * The class name and state attributes for a control that renders a picker trigger, not a
 * `ToolButton`: a font list, a size stepper, a dropdown. Such a control has no single icon to
 * replace and no single element to hand over, so `icon` and `asChild` are refused with a
 * development warning; `className` and the state attributes apply.
 */
export function useTriggerAttributes(
  slot: string,
  base: string,
  state: { readonly disabled?: boolean; readonly active?: boolean }
): TriggerAttributes {
  const override = useSlotOverride(slot);
  if (override?.icon !== undefined) warnUnsupportedOverride(slot, 'icon', slotPartNames(override));
  if (override?.asChild) warnUnsupportedOverride(slot, 'asChild', slotPartNames(override));
  return {
    className: [base, override?.className].filter(Boolean).join(' '),
    ...(state.active ? { 'data-active': '' as const } : {}),
    ...(state.disabled ? { 'data-disabled': '' as const } : {}),
  };
}

/**
 * The props that make a host element that is not a `<button>` act as one (WCAG 2.1.1): the
 * button role, a tab stop, and Enter and Space to activate. A link keeps its own Enter, which
 * already clicks it. `activate` gets the element. @internal
 */
export function buttonRoleProps(
  activate: ((target: Element) => void) | undefined
): HostElementProps {
  return {
    role: 'button',
    tabIndex: 0,
    onKeyDown: (event: KeyboardEvent & { currentTarget: Element }) => {
      const link = event.currentTarget instanceof HTMLAnchorElement && event.currentTarget.href;
      if (event.key !== ' ' && (event.key !== 'Enter' || link)) return;
      event.preventDefault();
      activate?.(event.currentTarget);
    },
  };
}

/** The names warnings use for a slot: the named part when there is one. @internal */
export function slotPartNames(props: object | null | undefined): ToolbarPartNames | undefined {
  const named = (props as { [NAMED_PART]?: string } | null | undefined)?.[NAMED_PART];
  return named ? { slot: named, group: 'Toolbar.Group' } : undefined;
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
 * Render `child` with the control's props merged in, after DOCX Editor's toolbar `Slot`
 * (docx-editor/packages/react/src/editor/toolbar/Slot.tsx): class names join, styles merge
 * (child wins per property), handlers run child-first (the control's handler is skipped when
 * the child called `preventDefault()`), and other props keep the child's own value.
 */
export function renderAsChild(
  child: ReactElement<HostElementProps>,
  slotProps: HostElementProps
): ReactElement {
  const childProps = child.props;
  const merged: HostElementProps = { ...slotProps, ...childProps };
  for (const key of Object.keys(slotProps)) {
    if (/^on[A-Z]/.test(key)) merged[key] = composeHandlers(childProps[key], slotProps[key]);
    else if (key === 'className')
      merged[key] = [slotProps[key], childProps[key]].filter(Boolean).join(' ');
    else if (key === 'style') merged[key] = Object.assign({}, slotProps[key], childProps[key]);
  }
  // The child keeps its ref: `cloneElement` keeps it in React 18, and React 19 keeps it in props.
  return cloneElement(child, merged);
}
