/** @jsxImportSource react */
// Host-owned toolbar controls, after DOCX Editor's `ToolbarAction` and `ToolbarSeparator`
// (docx-editor/packages/react/src/editor/toolbar/ToolbarAction.tsx, parts.tsx).
//
// `Toolbar.Slot` and the named parts are bound to a registered slot: the engine owns their
// label, icon, and enabled state. A host's own action ("Send for review") has no slot and
// never will. `Toolbar.Action` gives it the toolbar's button styling, state attributes, and
// the caret guard (a chrome mousedown never moves the slide selection), so a host does not
// copy private class names. It is not a way to fake a slot: enabled and pressed state are the
// host's to supply.

import { Children, Fragment, isValidElement, type ReactElement, type ReactNode } from 'react';
import { tooltipAttributes, warnAsChildChild } from '@pptx-editor.dev/core/editor';
import { keepFocus, useAriaDisabled } from './chrome.tsx';
import { buttonRoleProps, renderAsChild, type HostElementProps } from './slot-override.ts';

/** Props for `PptxEditor.Toolbar.Action`. @public */
export interface ToolbarActionProps {
  /** The accessible name and tooltip: a resolved string from the host's own catalog. */
  readonly label: string;
  /** Icon content, about 18 pixels, as the packaged icons. */
  readonly icon?: ReactNode;
  /** The pressed state, for an action that toggles: `aria-pressed` and `data-active`. */
  readonly active?: boolean;
  /** Turn the action off: the button stays focusable (`aria-disabled`) and `onSelect` does not run. */
  readonly disabled?: boolean;
  /** Why the action is disabled: in the tooltip and spoken through `aria-describedby`. */
  readonly disabledReason?: string;
  /** Runs the host's action. */
  readonly onSelect?: () => void;
  /** Merge the behavior onto the single child element instead of rendering a `<button>`. */
  readonly asChild?: boolean;
  /** A class name added to the button. */
  readonly className?: string;
  /** Visible content when there is no `icon`, or the element for `asChild`. */
  readonly children?: ReactNode;
}

/**
 * A host-owned toolbar action, styled and behaved like the packaged controls. Put it in a
 * `Toolbar.Group` so it collapses into the More menu with its group. A direct child of a
 * preset toolbar goes after the groups and never collapses, as in DOCX Editor. Without an
 * `icon` or children, it shows its `label` as text.
 *
 * @example
 * ```tsx
 * <PptxEditor.Toolbar>
 *   <PptxEditor.Toolbar.Group id="review" label="Review" priority={15}>
 *     <PptxEditor.Toolbar.Action label="Send for review" icon={<SendIcon />} onSelect={send} />
 *   </PptxEditor.Toolbar.Group>
 * </PptxEditor.Toolbar>
 * ```
 *
 * @public
 */
export function ToolbarAction(props: ToolbarActionProps): ReactNode {
  const { label, icon, active, disabled, disabledReason, onSelect, asChild, className, children } =
    props;
  // Focusable when disabled, with its reason spoken, as every toolbar control.
  const off = useAriaDisabled(!!disabled, disabledReason);
  const shared = {
    className: ['pptx-toolbar__button', className].filter(Boolean).join(' '),
    'aria-label': label,
    ...tooltipAttributes(label, disabled, disabledReason),
    // The same presence attributes as the registered controls, so one stylesheet selects both.
    'data-active': active ? '' : undefined,
    'data-disabled': disabled ? '' : undefined,
    'aria-pressed': active === undefined ? undefined : active,
    'aria-describedby': off.attrs['aria-describedby'],
    onMouseDown: keepFocus,
  };
  const only = asChild ? Children.toArray(children) : [];
  const child =
    only.length === 1 && isValidElement<HostElementProps>(only[0])
      ? (only[0] as ReactElement<HostElementProps>)
      : null;
  if (asChild && !child) warnAsChildChild(label, ACTION_PARTS);
  // `onSelect` takes no argument: the event or the element never reaches it.
  const select = disabled ? undefined : () => onSelect?.();
  if (child) {
    const isButton = child.type === 'button';
    return (
      <>
        {renderAsChild(child, {
          ...shared,
          ...(isButton ? { type: 'button' } : buttonRoleProps(select)),
          'aria-disabled': disabled ? true : undefined,
          onClick: select,
        })}
        {off.note}
      </>
    );
  }
  // A misused `asChild` falls back to the packaged button with the label.
  const own = asChild ? undefined : children;
  const content = icon ?? (own === null || own === false ? undefined : own);
  // Text needs room, laid out beside any icon: the label, or children with a text node in them
  // (`<SendIcon /> Send`). An element alone is icon content. The Vue twin uses the same rule.
  const text = icon == null && (content === undefined || hasText(content));
  const button = (
    <button
      type="button"
      {...shared}
      className={[shared.className, text ? 'pptx-toolbar__button--text' : '']
        .filter(Boolean)
        .join(' ')}
      aria-disabled={disabled ? true : undefined}
      onClick={select}
    >
      {content ?? label}
    </button>
  );
  return off.note ? (
    <>
      {button}
      {off.note}
    </>
  ) : (
    button
  );
}

/** Whether children hold a text node at the top level (through fragments and arrays). */
function hasText(node: ReactNode): boolean {
  if (typeof node === 'string' || typeof node === 'number') return String(node).trim() !== '';
  if (Array.isArray(node)) return node.some(hasText);
  return (
    isValidElement<{ children?: ReactNode }>(node) &&
    node.type === Fragment &&
    hasText(node.props.children)
  );
}

/** Warning names for `Toolbar.Action`: it is a part of its own, not a slot. */
const ACTION_PARTS = { slot: 'Toolbar.Action', group: 'Toolbar.Group' } as const;

/** Props for `PptxEditor.Toolbar.Separator`. @public */
export interface ToolbarSeparatorProps {
  /** A class name added to the separator. */
  readonly className?: string;
}

/** A vertical rule between controls, for a `preset={false}` toolbar. @public */
export function ToolbarSeparator({ className }: ToolbarSeparatorProps): ReactNode {
  return (
    <span
      role="separator"
      aria-orientation="vertical"
      className={['pptx-toolbar__sep', className].filter(Boolean).join(' ')}
    />
  );
}
