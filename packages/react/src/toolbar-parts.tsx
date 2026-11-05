/** @jsxImportSource react */
// Toolbar customization parts, after DOCX Editor's toolbar ladder
// (docx-editor/docs/site/content/customization/toolbar.mdx):
//
//   1. `className`, and state as `data-active` / `data-disabled` on every control.
//   2. `icon` on `Toolbar.Slot`.
//   3. `asChild` on `Toolbar.Slot`: the host element renders the control.
//   4. `Toolbar.Slot slotId` (replace one control, or `hidden`), `Toolbar.Group id` (change a
//      built-in group, or add a host group), and `preset={false}` (render only the children).
//   5. The raw hooks (`useEditorCommand`, `useEditorState`) for chrome built from nothing.
//
// Hiding a control changes the interface only: the command still runs from menus, keys, and
// the API.

import { Children, isValidElement, type ReactElement, type ReactNode } from 'react';
import type { ChromeGroupId, ChromeSlotId } from '@pptx-editor.dev/core/editor';
import {
  NAMED_PART,
  SlotOverrideContext,
  slotPartNames,
  type HostElementProps,
} from './slot-override.ts';
export { NAMED_PART, slotPartNames } from './slot-override.ts';
import {
  TOOLBAR_COMPOUND_SLOTS,
  chromeControlForSlot,
  warnAsChildChild,
  warnIgnoredSlotAttributes,
  warnReplacedControlSetting,
} from '@pptx-editor.dev/core/editor';
import { FontFamilyControl } from './FontFamilyControl.tsx';

/** What every `Toolbar.Slot` takes. @public */
export interface ToolbarSlotBaseProps {
  /** The registered control, for example `text.bold`. */
  readonly slotId: ChromeSlotId;
  /** Remove the control from the toolbar. Its command keeps working elsewhere. */
  readonly hidden?: boolean;
  /** A host icon in place of the registry icon. Button controls only. */
  readonly icon?: ReactNode;
  /** A class name added to the control. */
  readonly className?: string;
}

/**
 * Props for `PptxEditor.Toolbar.Slot`. With `asChild`, the single child element renders the
 * control and receives its label, state, and click handler. Without it, children replace the
 * control in its place. @public
 */
export type ToolbarSlotProps = ToolbarSlotBaseProps &
  (
    | {
        /** Render the control as the single child element. Button controls only. */
        readonly asChild: true;
        /** The host element that renders the control. */
        readonly children: ReactElement;
      }
    | {
        readonly asChild?: false;
        /** Content that replaces the control in its place. */
        readonly children?: ReactNode;
      }
  );

/**
 * Props for the named button parts (`Toolbar.Bold`, `Toolbar.Undo`, …), as DOCX Editor's: the
 * slot is pinned, and children are icon content (`icon`, else children, else the registry
 * icon), not a replacement. `<Toolbar.Bold hidden />` is
 * `<Toolbar.Slot slotId="text.bold" hidden />`. @public
 */
export type ToolbarPartProps = Omit<ToolbarSlotBaseProps, 'slotId'> &
  (
    | {
        /** Render the control as the single child element. */
        readonly asChild: true;
        /** The host element that renders the control. */
        readonly children: ReactElement;
      }
    | {
        readonly asChild?: false;
        /** Icon content, when there is no `icon`. */
        readonly children?: ReactNode;
      }
  );

/**
 * Props for the named picker parts (`Toolbar.FontFamily`, `Toolbar.FontColor`, …), as DOCX
 * Editor's slot parts: a picker is not one button, so it takes no `icon`, `asChild`, or
 * children. @public
 */
export interface ToolbarSlotPartProps {
  /** Remove the control from the toolbar. Its command keeps working elsewhere. */
  readonly hidden?: boolean;
  /** A class name added to the control's trigger. */
  readonly className?: string;
}

/**
 * Props for the named color split parts (`Toolbar.FontColor`, `Toolbar.Highlight`,
 * `Toolbar.ShapeFill`, `Toolbar.LineColor`), as DOCX Editor's `ToolbarColorSplitProps`: the
 * main half takes an `icon` (the glyph over the color bar). @public
 */
export interface ToolbarColorSplitProps extends ToolbarSlotPartProps {
  /** A host glyph over the color bar, in place of the registry icon. */
  readonly icon?: ReactNode;
}

/** A named color split part: a component with its slot as a static. @public */
export interface ToolbarColorSplitComponent {
  /** Changes, replaces, or hides the color control the part stands for. */
  (props: ToolbarColorSplitProps): ReactNode;
  /** The slot the part stands for. The toolbar reads it to place the part. */
  readonly pptxSlot: ChromeSlotId;
}

/** A named toolbar part: a `Toolbar.Slot` with its slot pinned. @public */
export interface ToolbarPartComponent {
  /** Changes, replaces, or hides the control the part stands for. */
  (props: ToolbarPartProps): ReactNode;
  /** The slot the part stands for. The toolbar reads it to place the part. */
  readonly pptxSlot: ChromeSlotId;
}

/** A named picker part: a component with its slot as a static. @public */
export interface ToolbarSlotPartComponent {
  /** Changes, replaces, or hides the control the part stands for. */
  (props: ToolbarSlotPartProps): ReactNode;
  /** The slot the part stands for. The toolbar reads it to place the part. */
  readonly pptxSlot: ChromeSlotId;
}

/**
 * The slot props a named part stands for: children of a button part are icon content (as in
 * DOCX Editor), and the part's name goes along for warning text. @internal
 */
export function namedPartSlotProps(
  slot: ChromeSlotId,
  props: Record<string, unknown>,
  partName: string | undefined
): ToolbarSlotProps {
  const shape = chromeControlForSlot(slot)?.shape;
  // A picker part keeps its props as given, so a stray `icon` or `asChild` still warns.
  if (shape !== undefined && shape !== 'icon')
    return { ...props, slotId: slot, [NAMED_PART]: partName } as unknown as ToolbarSlotProps;
  const { children, icon, ...rest } = props as { children?: ReactNode; icon?: ReactNode };
  const asChild = (props as { asChild?: boolean }).asChild === true;
  return {
    ...rest,
    slotId: slot,
    ...(asChild ? { icon, children } : { icon: icon ?? children }),
    [NAMED_PART]: partName,
  } as unknown as ToolbarSlotProps;
}

/** Props for `PptxEditor.Toolbar.Group`. @public */
export interface ToolbarGroupProps {
  /** A built-in group id changes that group; any other id adds a host group. */
  readonly id: ChromeGroupId | (string & {});
  /** Remove the whole group and its separator, from the bar and the More menu. */
  readonly hidden?: boolean;
  /** The accessible name of a host group. Wins over `labelKey`. */
  readonly label?: string;
  /** A translation key for the name of a host group, used when there is no `label`. */
  readonly labelKey?: string;
  /**
   * Collapse priority: a lower number leaves the bar for the More menu first. Built-in groups
   * use 10 (the last group) up to 90 (history). A host group without one leaves first.
   */
  readonly priority?: number;
  /** Keep the group on the bar at every width. */
  readonly pinned?: boolean;
  /**
   * The group a host group follows. Without it, a host group follows every built-in group.
   * A built-in id ignores it, with a development warning.
   */
  readonly after?: ChromeGroupId | (string & {});
  /** A class name added to the group's element. */
  readonly className?: string;
  /**
   * Content for the group in the More menu: after the controls of a built-in group, as in
   * DOCX Editor, so a built-in id adds a menu-only row; in place of a host group's controls.
   */
  readonly overflowContent?: () => ReactNode;
  /** Controls added after the group's built-in controls, or the controls of a host group. */
  readonly children?: ReactNode;
}

/** The overrides declared as `Toolbar` children when the preset is on. */
export interface ToolbarOverrides {
  readonly slots: ReadonlyMap<string, ToolbarSlotProps>;
  readonly groups: ReadonlyMap<string, ToolbarGroupProps>;
  /** Other children: a loose `Toolbar.Action` or host element, after the groups. */
  readonly extras: readonly ReactNode[];
}

/** The one element child of an `asChild` slot, or null (with a warning) when there is not one. */
function asChildElement(props: ToolbarSlotProps): ReactElement<HostElementProps> | null {
  const children = Children.toArray(props.children);
  const [child] = children;
  if (children.length === 1 && isValidElement<HostElementProps>(child)) return child;
  warnAsChildChild(props.slotId, slotPartNames(props));
  return null;
}

/** The props a slot part reads. Anything else (`style`, `data-*`, a listener) never reaches the control. */
const SLOT_PROPS: ReadonlySet<string> = new Set([
  'slotId',
  'hidden',
  'icon',
  'className',
  'asChild',
  'children',
]);

/** The control a slot renders, with the slot's overrides applied. */
export function SlotControl({
  props,
  render,
}: {
  props: ToolbarSlotProps | undefined;
  render: () => ReactNode;
}): ReactNode {
  if (!props) return render();
  if (props.hidden) return null;
  const parts = slotPartNames(props);
  const ignored = Object.keys(props).filter((key) => !SLOT_PROPS.has(key));
  warnIgnoredSlotAttributes(props.slotId, ignored, 'className', parts);
  // The children of a compound named part (`Toolbar.FontFamily`) compose its sub-parts inside
  // the control's root; any other slot's children replace the control.
  const named = (props as { [NAMED_PART]?: string })[NAMED_PART];
  const compose = !!named && TOOLBAR_COMPOUND_SLOTS.has(props.slotId);
  if (!props.asChild && props.children !== undefined && !compose) {
    // The children replace the control, so a setting for the control does nothing: say so.
    if (props.className) warnReplacedControlSetting(props.slotId, 'className', parts);
    if (props.icon !== undefined) warnReplacedControlSetting(props.slotId, 'icon', parts);
    return <>{props.children}</>;
  }
  const control = compose ? chromeControlForSlot(props.slotId) : null;
  const body =
    control && !props.asChild && props.children !== undefined
      ? () => (
          <FontFamilyControl slot={props.slotId} control={control}>
            {props.children}
          </FontFamilyControl>
        )
      : render;
  const child = props.asChild ? asChildElement(props) : null;
  return (
    <SlotOverrideContext.Provider
      value={{
        slotId: props.slotId,
        ...(props.icon !== undefined ? { icon: props.icon } : {}),
        ...(props.className ? { className: props.className } : {}),
        ...(child ? { asChild: child } : {}),
        // The named part's name, for warning text.
        [NAMED_PART]: named,
      }}
    >
      {body()}
    </SlotOverrideContext.Provider>
  );
}
