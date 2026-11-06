// The toolbar's host parts and named parts: the Vue twin of the React adapter's
// `toolbar-action.tsx` and `toolbar-named-parts.tsx`, after DOCX Editor's `ToolbarAction`,
// `ToolbarSeparator`, and named parts. A named part (`PptxEditorToolbar.Bold`) is a
// `ToolbarSlot` with its slot pinned; the names come from core (`TOOLBAR_NAMED_PARTS`), so
// both adapters use the same ones.

import {
  camelize,
  cloneVNode,
  defineComponent,
  h,
  isVNode,
  type DefineComponent,
  type ExtractPublicPropTypes,
  type PropType,
  type VNode,
  type VNodeChild,
  Text,
} from 'vue';
import {
  TOOLBAR_NAMED_PARTS,
  chromeControlForSlot,
  tooltipAttributes,
  warnAsChildChild,
  type ChromeSlotId,
} from '@pptx-editor.dev/core/editor';
import { keepFocus } from './chrome.tsx';
import {
  FONT_FAMILY_PARTS,
  type FontFamilyItemProps,
  type FontFamilyPartProps,
} from './FontFamilyControl.tsx';
import {
  SlotControl,
  ToolbarGroup,
  ToolbarRoot,
  ToolbarSlot,
  slotList,
  slotSpec,
} from './toolbar.ts';

const toolbarActionProps = {
  /** The accessible name and tooltip: a resolved string from the host's own catalog. */
  label: { type: String, required: true as const },
  /** Icon content, about 18 pixels, as DOCX Editor's Vue `ToolbarAction`. The `icon` slot also works. */
  icon: { type: Object as PropType<VNode>, default: undefined },
  /** A class name added to the button, as React's `className`. */
  className: { type: String, default: undefined },
  /** The pressed state, for an action that toggles: `aria-pressed` and `data-active`. */
  active: { type: Boolean as PropType<boolean | undefined>, default: undefined },
  /** Turn the action off: the button stays focusable (`aria-disabled`) and `select` is not emitted. */
  disabled: { type: Boolean, default: false },
  /** Why the action is disabled, added to the tooltip. */
  disabledReason: { type: String, default: undefined },
  /** Merge the behavior onto the single default-slot element instead of a `<button>`. */
  asChild: { type: Boolean, default: false },
} as const;

/** Props of `ToolbarAction`, the twin of React's `ToolbarActionProps`. @public */
export type ToolbarActionProps = ExtractPublicPropTypes<typeof toolbarActionProps>;

/** Warning names for `ToolbarAction`: it is a part of its own, not a slot. */
const ACTION_PARTS = { slot: 'ToolbarAction', group: 'ToolbarGroup' } as const;

/**
 * A host-owned toolbar action, styled and behaved like the packaged controls, as React's
 * `Toolbar.Action`. Emits `select`; the `icon` prop, the `icon` slot, or the default slot is
 * its content, else its label as text. A direct child of a preset toolbar goes after the
 * groups and never collapses. @public
 */
export const ToolbarAction = defineComponent({
  name: 'ToolbarAction',
  props: toolbarActionProps,
  emits: { select: () => true },
  setup(props, { emit, slots }) {
    return () => {
      const shared = {
        class: ['pptx-toolbar__button', props.className],
        'aria-label': props.label,
        ...tooltipAttributes(props.label, props.disabled, props.disabledReason),
        'data-active': props.active ? '' : undefined,
        'data-disabled': props.disabled ? '' : undefined,
        'aria-pressed': props.active,
        onMousedown: keepFocus,
        onClick: () => {
          if (!props.disabled) emit('select');
        },
      };
      const select = (): void => {
        if (!props.disabled) emit('select');
      };
      if (props.asChild) {
        // Exactly one element: text and comments have symbol types.
        const nodes = (slots.default?.() ?? []).filter(isVNode);
        const child = nodes.length === 1 && typeof nodes[0]!.type !== 'symbol' ? nodes[0] : null;
        if (child) {
          const isButton = child.type === 'button';
          const wiring: Record<string, unknown> = {
            ...shared,
            // Runs after the child's own click handler: a child that calls preventDefault()
            // keeps the action from running, as in React's renderAsChild.
            onClick: (event: MouseEvent) => {
              if (!event.defaultPrevented) select();
            },
            ...(isButton
              ? { type: 'button' }
              : {
                  // Not a button: the role, a tab stop, and Enter and Space (WCAG 2.1.1). A
                  // link keeps its own Enter, which already clicks it.
                  role: 'button',
                  tabindex: 0,
                  onKeydown: (event: KeyboardEvent) => {
                    // As the React twin: only a link with an href keeps its own Enter.
                    const target = event.currentTarget as Element & { href?: string };
                    const link = target.localName === 'a' && !!target.href;
                    if (event.key !== ' ' && (event.key !== 'Enter' || link)) return;
                    event.preventDefault();
                    select();
                  },
                }),
            'aria-disabled': props.disabled ? 'true' : undefined,
          };
          // The child's own attributes win (as React's renderAsChild); class and handlers merge.
          const own = child.props ?? {};
          for (const key of Object.keys(wiring))
            if (key !== 'class' && !/^on[A-Z]/.test(key) && key in own) delete wiring[key];
          return cloneVNode(child, wiring);
        }
        warnAsChildChild(props.label, ACTION_PARTS);
      }
      // A misused `asChild` falls back to the packaged button with the label.
      const own = props.asChild ? undefined : slots.default?.();
      const content = props.icon ?? slots.icon?.() ?? (own?.length ? own : undefined);
      // Text (the label, or text children) needs room; one element child is icon content.
      // As the React twin: text (the label, or content with a text node) sits beside any icon.
      const list: unknown[] = Array.isArray(content) ? content : content ? [content] : [];
      const hasText = list.some(
        (n) =>
          (typeof n === 'string' && n.trim() !== '') ||
          (isVNode(n) && n.type === Text && String(n.children ?? '').trim() !== '')
      );
      const text = props.icon == null && !slots.icon && (content === undefined || hasText);
      return h(
        'button',
        {
          ...shared,
          class: [shared.class, text ? 'pptx-toolbar__button--text' : null],
          type: 'button',
          // Focusable when disabled, as every toolbar control.
          'aria-disabled': props.disabled ? 'true' : undefined,
        },
        content ?? props.label
      );
    };
  },
});

const toolbarSeparatorProps = {
  /** A class name added to the rule, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `ToolbarSeparator`, the twin of React's `ToolbarSeparatorProps`. @public */
export type ToolbarSeparatorProps = ExtractPublicPropTypes<typeof toolbarSeparatorProps>;

/** A vertical rule between controls, as React's `Toolbar.Separator`. @public */
export const ToolbarSeparator = defineComponent({
  name: 'ToolbarSeparator',
  props: toolbarSeparatorProps,
  setup(props) {
    return () =>
      h('span', {
        role: 'separator',
        'aria-orientation': 'vertical',
        class: ['pptx-toolbar__sep', props.className],
      });
  },
});

const hostClassProps = {
  /** A class name added to the control, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

const toolbarSlotPartProps = {
  /** Remove the control from the toolbar. Its command keeps working elsewhere. */
  hidden: { type: Boolean, default: undefined },
  ...hostClassProps,
} as const;

const toolbarColorSplitProps = {
  ...toolbarSlotPartProps,
  /** A host glyph over the color bar, in place of the registry icon. The `icon` slot also works. */
  icon: { type: [Object, Array, String] as PropType<VNodeChild>, default: undefined },
} as const;

const toolbarPartProps = {
  ...toolbarColorSplitProps,
  /** Render the control as the single default-slot element. */
  asChild: { type: Boolean, default: undefined },
} as const;

/**
 * Props of the named button parts (`PptxEditorToolbar.Bold`, `PptxEditorToolbar.Undo`, …), as
 * React's and DOCX Editor's: the slot is pinned, and the default slot is icon content (`icon`,
 * else the default slot, else the registry icon), not a replacement. With `asChild`, the one
 * default-slot element renders the control. @public
 */
export type ToolbarPartProps = ExtractPublicPropTypes<typeof toolbarPartProps>;

/**
 * Props of the named picker parts (`PptxEditorToolbar.FontFamily`, …), as DOCX Editor's slot
 * parts: a picker is not one button, so it takes no `icon`, `asChild`, or slot content. @public
 */
export type ToolbarSlotPartProps = ExtractPublicPropTypes<typeof toolbarSlotPartProps>;

/**
 * Props of the named color split parts (`PptxEditorToolbar.FontColor`, `Highlight`,
 * `ShapeFill`, `LineColor`, `TableCellFill`), as DOCX Editor's `ToolbarColorSplitProps`: the
 * main half takes an `icon` (the glyph over the color bar). @public
 */
export type ToolbarColorSplitProps = ExtractPublicPropTypes<typeof toolbarColorSplitProps>;

/** A named button part: a component with its slot as a static. @public */
export type ToolbarPartComponent = DefineComponent<ToolbarPartProps> & {
  readonly pptxSlot: ChromeSlotId;
};

/** A named picker part: a component with its slot as a static. @public */
export type ToolbarSlotPartComponent = DefineComponent<ToolbarSlotPartProps> & {
  readonly pptxSlot: ChromeSlotId;
};

/**
 * Props of `PptxEditorToolbar.FontFamily`, as React's `FontFamilyProps`: a picker part's
 * `className` and `hidden`; its default slot composes the `Trigger`, `Content`, and `Item` parts.
 * @public
 */
export type FontFamilyProps = ToolbarSlotPartProps;

/**
 * `PptxEditorToolbar.FontFamily`: the `font.family` picker, as DOCX Editor's `FontFamily`
 * compound. With no slot content it renders the default trigger and list; with slot content,
 * compose the `Trigger`, `Content`, and `Item` parts around the shared state. @public
 */
export type FontFamilyNamespace = DefineComponent<FontFamilyProps> & {
  readonly pptxSlot: ChromeSlotId;
  /** The button that opens the list. */
  readonly Trigger: DefineComponent<FontFamilyPartProps>;
  /** The list, while it is open. */
  readonly Content: DefineComponent<FontFamilyPartProps>;
  /** One option of the list. */
  readonly Item: DefineComponent<FontFamilyItemProps>;
};

/** A named color split part: a component with its slot as a static. @public */
export type ToolbarColorSplitComponent = DefineComponent<ToolbarColorSplitProps> & {
  readonly pptxSlot: ChromeSlotId;
};

/** The named picker parts, as React types them (`ToolbarSlotPartComponent`). */
const SLOT_PART_NAMES = [
  'FontFamily',
  'FontSize',
  'CharacterSpacing',
  'ChangeCase',
  'Alignment',
  'VerticalAlignment',
  'BulletStyle',
  'NumberStyle',
  'LineSpacing',
  'TextDirection',
  'Autofit',
  'Columns',
  'Line',
  'Shapes',
  'LineWidth',
  'LineStart',
  'LineEnd',
  'LineDash',
  'ShapeStyles',
  'ShapeEffects',
  'Order',
  'AlignObjects',
  'Rotate',
  'TableStyle',
  'TableBorderTarget',
  'TableBorderStyle',
  'TableBorderWidth',
  'TableSelect',
  'TableDelete',
  'TableSplitCell',
  'TableCellMargins',
] as const;

/** The named color split parts, as React types them (`ToolbarColorSplitComponent`). */
const COLOR_SPLIT_NAMES = [
  'FontColor',
  'Highlight',
  'ShapeFill',
  'LineColor',
  'TableCellFill',
  'TableBorderColor',
] as const;

type SlotPartName = (typeof SLOT_PART_NAMES)[number];
type ColorSplitName = (typeof COLOR_SPLIT_NAMES)[number];

function definePart(name: string, slot: ChromeSlotId): ToolbarPartComponent {
  const part = `PptxEditorToolbar.${name}`;
  const props: Record<string, unknown> = (SLOT_PART_NAMES as readonly string[]).includes(name)
    ? toolbarSlotPartProps
    : (COLOR_SPLIT_NAMES as readonly string[]).includes(name)
      ? toolbarColorSplitProps
      : toolbarPartProps;
  return Object.assign(
    defineComponent({
      name: `Toolbar${name}`,
      inheritAttrs: false,
      props: props as typeof toolbarPartProps,
      setup(declared, { attrs, slots }) {
        return () => {
          if (!chromeControlForSlot(slot)) return null;
          // Undeclared `icon` or `asChild` on a picker part reach the spec from attrs, so the
          // control warns that it cannot apply them, as in React.
          const given: Record<string, unknown> = {};
          for (const [key, value] of Object.entries(attrs)) given[camelize(key)] = value;
          Object.assign(given, declared);
          if (given.icon === undefined && slots.icon) given.icon = slots.icon;
          const spec = slotSpec(slot, given, slotList(slots.default), part);
          return h(SlotControl, { spec });
        };
      },
    }),
    { pptxSlot: slot }
  ) as unknown as ToolbarPartComponent;
}

type NamedParts = {
  readonly [K in keyof typeof TOOLBAR_NAMED_PARTS]: K extends 'FontFamily'
    ? FontFamilyNamespace
    : K extends SlotPartName
      ? ToolbarSlotPartComponent
      : K extends ColorSplitName
        ? ToolbarColorSplitComponent
        : ToolbarPartComponent;
};

// `FontFamily` carries its compound parts as statics.
const NAMED_PARTS = Object.fromEntries(
  Object.entries(TOOLBAR_NAMED_PARTS).map(([name, slot]) => {
    const part = definePart(name, slot);
    return [name, name === 'FontFamily' ? Object.assign(part, FONT_FAMILY_PARTS) : part];
  })
) as unknown as NamedParts;

/**
 * The statics of `PptxEditorToolbar`, the twin of React's `ToolbarNamespace`: `Slot`, `Group`,
 * `Action`, `Separator`, and the named parts. @public
 */
export type PptxEditorToolbarNamespace = typeof ToolbarRoot &
  NamedParts & {
    readonly Slot: typeof ToolbarSlot;
    readonly Group: typeof ToolbarGroup;
    readonly Action: typeof ToolbarAction;
    readonly Separator: typeof ToolbarSeparator;
  };

/**
 * The toolbar, with its parts as statics, as React's `PptxEditor.Toolbar`: `Slot`, `Group`,
 * `Action`, `Separator`, and the named parts (`PptxEditorToolbar.Bold`, …). @public
 */
export const PptxEditorToolbar: PptxEditorToolbarNamespace = Object.assign(ToolbarRoot, {
  ...NAMED_PARTS,
  Slot: ToolbarSlot,
  Group: ToolbarGroup,
  Action: ToolbarAction,
  Separator: ToolbarSeparator,
});
