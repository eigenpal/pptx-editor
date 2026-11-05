// The toolbar: the Vue twin of the React adapter's `PptxEditor.Toolbar`, with the same
// customization ladder (DOCX Editor's toolbar.mdx): `class` and state attributes, `icon`,
// `asChild`, `ToolbarSlot slotId` and `ToolbarGroup id` overrides with `hidden`, and
// `preset` false. The arrangement comes from core (`arrangeToolbarGroups`,
// `toolbarGroupsFor`), so a group override means the same in both adapters, and groups that do
// not fit the width move into a More menu by core's rule (`toolbarCollapsedIds`). The markup and
// classes match React's, so the one core stylesheet styles both.
//
// `ToolbarSlot` and `ToolbarGroup` under a preset bar are descriptions: the bar reads their
// props and renders the controls itself, as in DOCX Editor's Vue toolbar.

import {
  Comment,
  Text,
  camelize,
  computed,
  defineComponent,
  h,
  isVNode,
  normalizeClass,
  provide,
  type ExtractPublicPropTypes,
  type PropType,
  type VNode,
  type VNodeArrayChildren,
  type VNodeChild,
} from 'vue';
import {
  TOOLBAR_COMPOUND_SLOTS,
  TOOLBAR_NAMED_PARTS,
  arrangeToolbarGroups,
  chromeControlForSlot,
  isToolbarBarSlot,
  toolbarGroupLabel,
  toolbarGroupPriority,
  toolbarContextOf,
  toolbarGroupsFor,
  translateKeyOrText,
  warnAsChildChild,
  warnDuplicateSlot,
  warnIgnoredSlotAttributes,
  warnReplacedControlSetting,
  warnSlotOutsideBar,
  warnUnknownSlot,
  type ChromeGroup,
  type ChromeGroupId,
  type ChromeSlotId,
  type ToolbarPartNames,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { useNotesTarget } from './notes-target.ts';
import { slotOverrideKey, type SlotOverride } from './slot-override.ts';
import { ToolbarControl } from './toolbar-controls.tsx';
import { FontFamilyControl } from './FontFamilyControl.tsx';
import { HandBar, PresetToolbar, type BarItem } from './toolbar-bar.tsx';

export { ToolbarControl } from './toolbar-controls.tsx';

/** The Vue names of the toolbar parts, for development warnings. */
const PARTS: ToolbarPartNames = { slot: 'ToolbarSlot', group: 'ToolbarGroup' };

type SlotFn = () => VNodeArrayChildren;

/**
 * A slot as a function that always returns a list: a render function may return one vnode, a
 * string, or nothing. @internal
 */
export function slotList(fn: unknown): SlotFn | undefined {
  if (typeof fn !== 'function') return undefined;
  return () => {
    const out = (fn as () => unknown)();
    if (Array.isArray(out)) return out as VNodeArrayChildren;
    return out === null || out === undefined ? [] : [out as VNodeArrayChildren[number]];
  };
}

/** One slot override, read from a `ToolbarSlot` or a named part. @internal */
export interface SlotSpec {
  readonly slotId: ChromeSlotId;
  readonly hidden: boolean;
  /** Icon content in place of the registry icon: a vnode, or a slot that renders it. */
  readonly icon: VNodeChild | SlotFn | undefined;
  readonly className: string | undefined;
  readonly asChild: boolean;
  /** The default slot: content that replaces the control, or the `asChild` element. */
  readonly children: SlotFn | undefined;
  /** The named part (`PptxEditorToolbar.Bold`) the spec came from, for warning text. */
  readonly part: string | undefined;
}

/** A boolean prop: present with no value (`hidden`) means true, as Vue casts it. */
const bool = (value: unknown): boolean => value === '' || value === true;

/**
 * The spec of a slot override. A named button part (`PptxEditorToolbar.Bold`) takes its default
 * slot as icon content, as in React and DOCX Editor, unless it has `asChild`; a named picker part
 * keeps its props as given, so a stray `icon` or `asChild` still warns. @internal
 */
export function slotSpec(
  slotId: ChromeSlotId,
  props: Record<string, unknown>,
  children: SlotFn | undefined,
  part?: string
): SlotSpec {
  // Only the class, icon, and asChild reach the control; any other attribute (`style`, `id`,
  // `data-*`, a listener) warns, as in React.
  const ignored = Object.keys(props)
    .filter((key) => !SLOT_KEYS.has(key))
    .map((key) => hyphenate(key));
  const names: ToolbarPartNames = part ? { slot: part, group: PARTS.group } : PARTS;
  warnIgnoredSlotAttributes(slotId, ignored, 'class', names);
  const shape = chromeControlForSlot(slotId)?.shape;
  const asChild = bool(props.asChild);
  const iconPart = part !== undefined && (shape === undefined || shape === 'icon') && !asChild;
  const icon = props.icon as VNodeChild | undefined;
  return {
    slotId,
    hidden: bool(props.hidden),
    icon: iconPart ? (icon ?? children) : icon,
    className: normalizeClass([props.class, props.className]) || undefined,
    asChild,
    children: iconPart ? undefined : children,
    part,
  };
}

/** An attribute name as a template writes it: `dataTest` to `data-test`. */
const hyphenate = (key: string): string => key.replace(/\B([A-Z])/g, '-$1').toLowerCase();

/** The props and attributes a slot part reads; anything else never reaches the control. */
const SLOT_KEYS: ReadonlySet<string> = new Set([
  'slotId',
  'hidden',
  'icon',
  'class',
  'className',
  'asChild',
  'key',
  'ref',
  'ref_for',
  'ref_key',
]);

/** The names warnings use for a spec: the named part when there is one. */
const specParts = (spec: SlotSpec): ToolbarPartNames =>
  spec.part ? { slot: spec.part, group: PARTS.group } : PARTS;

/** The one element of an `asChild` slot, or undefined (with a warning) when there is not one. */
function asChildElement(spec: SlotSpec): VNode | undefined {
  const nodes = flatten(spec.children?.());
  const [child] = nodes;
  if (nodes.length === 1 && child && typeof child.type !== 'symbol') return child;
  warnAsChildChild(spec.slotId, specParts(spec));
  return undefined;
}

const slotControlProps = {
  spec: { type: Object as PropType<SlotSpec>, required: true as const },
} as const;

/**
 * The control a slot renders, with the slot's overrides applied: hidden renders nothing, slot
 * content replaces the control, and otherwise the control renders with the host's icon, class,
 * and `asChild` element through `slotOverrideKey`. @internal
 */
export const SlotControl = defineComponent({
  name: 'PptxToolbarSlotControl',
  props: slotControlProps,
  setup(props) {
    // Getters, so the control reads the current spec in its own render.
    const override: SlotOverride = {
      get slotId() {
        return props.spec.slotId;
      },
      get icon() {
        const icon = props.spec.icon;
        return typeof icon === 'function' ? (icon as SlotFn)() : icon;
      },
      get className() {
        return props.spec.className;
      },
      get asChild() {
        return props.spec.asChild ? asChildElement(props.spec) : undefined;
      },
      get part() {
        return props.spec.part;
      },
    };
    provide(slotOverrideKey, override);
    return () => {
      const spec = props.spec;
      if (spec.hidden) return null;
      // The slot content of a compound named part (`PptxEditorToolbar.FontFamily`) composes its
      // sub-parts inside the control's root; any other slot's content replaces the control.
      const compose = !!spec.part && TOOLBAR_COMPOUND_SLOTS.has(spec.slotId);
      if (!spec.asChild && spec.children && !compose) {
        // The content replaces the control, so a setting for the control does nothing: say so.
        if (spec.className) warnReplacedControlSetting(spec.slotId, 'class', specParts(spec));
        if (spec.icon !== undefined)
          warnReplacedControlSetting(spec.slotId, 'icon', specParts(spec));
        return spec.children();
      }
      const control = chromeControlForSlot(spec.slotId);
      if (!control) return null;
      if (compose && !spec.asChild && spec.children)
        return h(FontFamilyControl, { slot: spec.slotId, control }, { default: spec.children });
      const group = spec.slotId.slice(0, spec.slotId.indexOf('.'));
      return h(ToolbarControl, { group, control });
    };
  },
});

const toolbarSlotProps = {
  /** The registered control, for example `text.bold`. */
  slotId: { type: String as PropType<ChromeSlotId>, required: true as const },
  /** Remove the control from the toolbar. Its command keeps working elsewhere. */
  hidden: { type: Boolean, default: false },
  /** A host icon in place of the registry icon. Button controls only; the `icon` slot also works. */
  icon: { type: [Object, Array, String] as PropType<VNodeChild>, default: undefined },
  /** Render the control as the single default-slot element. Button controls only. */
  asChild: { type: Boolean, default: false },
  /** A class name added to the control, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `ToolbarSlot`, the twin of React's `ToolbarSlotProps`. @public */
export type ToolbarSlotProps = ExtractPublicPropTypes<typeof toolbarSlotProps>;

/**
 * One registered control. Inside a `PptxEditorToolbar` with the preset on, it declares an
 * override; anywhere else it renders the control. Default slot content replaces the control;
 * with `asChild`, its one element renders the control. @public
 */
export const ToolbarSlot = defineComponent({
  name: 'ToolbarSlot',
  inheritAttrs: false,
  props: toolbarSlotProps,
  setup(props, { attrs, slots }) {
    return () => {
      if (!chromeControlForSlot(props.slotId)) {
        warnUnknownSlot(props.slotId, PARTS);
        return null;
      }
      const spec = slotSpec(
        props.slotId,
        { ...attrs, ...props, icon: props.icon ?? slots.icon, class: attrs.class },
        slotList(slots.default)
      );
      return h(SlotControl, { spec });
    };
  },
});

const toolbarGroupProps = {
  /** A built-in group id changes that group; any other id adds a host group. */
  id: { type: String as PropType<ChromeGroupId | (string & {})>, required: true as const },
  /** Remove the whole group and its separator, from the bar and the More menu. */
  hidden: { type: Boolean, default: false },
  /** The accessible name of a host group. Wins over `labelKey`. */
  label: { type: String, default: undefined },
  /** A translation key for the name of a host group, used when there is no `label`. */
  labelKey: { type: String, default: undefined },
  /**
   * Collapse priority: a lower number leaves the bar for the More menu first. Built-in groups
   * use 10 (the last group) up to 90 (history). A host group without one leaves first.
   */
  priority: { type: Number, default: undefined },
  /** Keep the group on the bar at every width. */
  pinned: { type: Boolean, default: false },
  /**
   * The group a host group follows. Without it, a host group follows every built-in group.
   * A built-in id ignores it, with a development warning.
   */
  after: { type: String as PropType<ChromeGroupId | (string & {})>, default: undefined },
  /** A class name added to the group's element. A host `class` also works. */
  className: { type: String, default: undefined },
  /**
   * Content for the group in the More menu: after the controls of a built-in group, as in
   * DOCX Editor, so a built-in id adds a menu-only row; in place of a host group's controls.
   */
  overflowContent: { type: Function as PropType<() => VNodeChild>, default: undefined },
} as const;

/** Props of `ToolbarGroup`, the twin of React's `ToolbarGroupProps`. @public */
export type ToolbarGroupProps = ExtractPublicPropTypes<typeof toolbarGroupProps>;

/**
 * A group of controls. Inside a `PptxEditorToolbar` with the preset on, a built-in id changes
 * that group and any other id adds a host group; with `preset` false, it renders its slot. A
 * host group's accessible name is its `label`, else its translated `labelKey`, else its `id`.
 * @public
 */
export const ToolbarGroup = defineComponent({
  name: 'ToolbarGroup',
  props: toolbarGroupProps,
  setup(props, { slots }) {
    const { t } = useTranslation();
    return () =>
      props.hidden
        ? null
        : h(
            'div',
            {
              class: ['pptx-toolbar__group', props.className],
              role: 'group',
              'aria-label': toolbarGroupLabel(props, (key) => translateKeyOrText(t, key)),
              'data-group': props.id,
            },
            slots.default?.()
          );
  },
});

/** A group override read from a `ToolbarGroup` child, with its slot content. */
interface GroupOverride {
  readonly id: string;
  readonly hidden: boolean;
  readonly label?: string;
  readonly labelKey?: string;
  readonly after?: string;
  readonly priority: number | undefined;
  readonly pinned: boolean;
  readonly className: string | undefined;
  readonly overflowContent: (() => VNodeChild) | undefined;
  readonly content: VNodeArrayChildren | undefined;
}

/** The slot a named part (`PptxEditorToolbar.Bold`) stands for, read from its static. */
const namedPartSlot = (type: unknown): ChromeSlotId | undefined =>
  (type as { pptxSlot?: ChromeSlotId } | null)?.pptxSlot;

/** A vnode's props with camelCase keys (templates pass `slot-id`, `as-child`). */
function vnodeProps(vnode: VNode): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(vnode.props ?? {})) out[camelize(key)] = value;
  return out;
}

/** Vnodes with fragments opened; comments (a false `v-if`) and blank text left out. */
export function flatten(children: VNodeArrayChildren | undefined, out: VNode[] = []): VNode[] {
  for (const child of children ?? []) {
    if (Array.isArray(child)) flatten(child, out);
    else if (isVNode(child)) {
      if (child.type === Comment) continue;
      if (child.type === Text && String(child.children ?? '').trim() === '') continue;
      if (Array.isArray(child.children) && typeof child.type === 'symbol')
        flatten(child.children as VNodeArrayChildren, out);
      else out.push(child);
    } else if (typeof child === 'string' && child.trim() !== '') out.push(h(Text, child));
    else if (typeof child === 'number') out.push(h(Text, String(child)));
  }
  return out;
}

/** A named part's name by its slot, for warning text: `text.bold` → `PptxEditorToolbar.Bold`. */
const NAMED_PART_BY_SLOT = new Map<string, string>(
  Object.entries(TOOLBAR_NAMED_PARTS).map(([name, slot]) => [slot, `PptxEditorToolbar.${name}`])
);

/** A vnode's named slot, as a function. */
function vnodeSlot(vnode: VNode, name: string): SlotFn | undefined {
  const children = vnode.children as Record<string, unknown> | null;
  if (!children || typeof children !== 'object' || Array.isArray(children)) return undefined;
  return slotList(children[name]);
}

/** The spec a `ToolbarSlot` or named part vnode declares. */
function vnodeSpec(vnode: VNode, named: ChromeSlotId | undefined): SlotSpec {
  const props = vnodeProps(vnode);
  if (props.icon === undefined && vnodeSlot(vnode, 'icon')) props.icon = vnodeSlot(vnode, 'icon');
  const slotId = named ?? (String(props.slotId) as ChromeSlotId);
  return slotSpec(
    slotId,
    props,
    vnodeSlot(vnode, 'default'),
    named ? NAMED_PART_BY_SLOT.get(named) : undefined
  );
}

/** A number prop that a template may pass as text (`priority="15"`). */
function numberProp(value: unknown): number | undefined {
  if (typeof value === 'number') return value;
  if (typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value)))
    return Number(value);
  return undefined;
}

/**
 * Read `ToolbarSlot`, named part, and `ToolbarGroup` children as overrides of the preset. Only
 * direct children declare overrides. Other children (a loose `ToolbarAction`, a host element) go
 * after the groups and never collapse, as in React and DOCX Editor.
 */
function collectOverrides(
  children: VNodeArrayChildren | undefined,
  inBar: (slot: string) => boolean = isToolbarBarSlot
): { slots: Map<string, SlotSpec>; groups: Map<string, GroupOverride>; extras: VNode[] } {
  const slots = new Map<string, SlotSpec>();
  const groups = new Map<string, GroupOverride>();
  const extras: VNode[] = [];
  for (const vnode of flatten(children)) {
    const named = namedPartSlot(vnode.type);
    if (vnode.type === ToolbarSlot || named) {
      const spec = vnodeSpec(vnode, named);
      const parts = specParts(spec);
      if (!chromeControlForSlot(spec.slotId)) warnUnknownSlot(spec.slotId, parts);
      else if (!inBar(spec.slotId)) {
        // A registered control the preset bar does not draw goes after the groups, as in React;
        // a `ToolbarSlot` says so (a named part does not).
        // A hidden slot draws nothing, so it gets no warning.
        if (!named && !spec.hidden) warnSlotOutsideBar(spec.slotId, parts);
        if (!spec.hidden) extras.push(vnode);
        continue;
      }
      if (slots.has(spec.slotId)) warnDuplicateSlot(spec.slotId, parts);
      slots.set(spec.slotId, spec);
    } else if (vnode.type === ToolbarGroup) {
      const props = vnodeProps(vnode);
      const str = (v: unknown): string | undefined => (typeof v === 'string' ? v : undefined);
      const label = str(props.label);
      const labelKey = str(props.labelKey);
      const after = str(props.after);
      groups.set(String(props.id), {
        id: String(props.id),
        hidden: bool(props.hidden),
        ...(label !== undefined ? { label } : {}),
        ...(labelKey !== undefined ? { labelKey } : {}),
        ...(after !== undefined ? { after } : {}),
        priority: numberProp(props.priority),
        pinned: bool(props.pinned),
        className: normalizeClass([props.class, props.className]) || undefined,
        overflowContent:
          typeof props.overflowContent === 'function'
            ? (props.overflowContent as () => VNodeChild)
            : undefined,
        content: vnodeSlot(vnode, 'default')?.(),
      });
    } else extras.push(vnode);
  }
  return { slots, groups, extras };
}

/** Whether `nodes` renders anything: not only comments, blank text, or empty fragments. */
const hasContent = (nodes: VNodeArrayChildren | undefined): boolean => flatten(nodes).length > 0;

const pptxEditorToolbarProps = {
  /**
   * `true` (default): the registry's groups, changed by `ToolbarSlot` and `ToolbarGroup`
   * children. `false`: only the children, in the order given.
   */
  preset: { type: Boolean, default: true },
  /**
   * `true` (default): groups that do not fit the width move into a More menu, lowest priority
   * first. `false`: the bar wraps to more rows instead. A `preset` false bar never measures
   * itself, so it always wraps, as in DOCX Editor.
   */
  overflow: { type: Boolean, default: true },
  /**
   * `true` (default): the bar shows only the groups that act on the selection (text, a shape,
   * a picture, a table, or nothing). `false`: every bar group shows for every selection,
   * disabled with the engine's reason where it cannot run, and the picture or table group
   * joins after them. Hiding a group never turns off its commands.
   */
  contextual: { type: Boolean, default: true },
  /** A class name added to the bar, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `PptxEditorToolbar`, the twin of React's `ToolbarProps`. @public */
export type ToolbarProps = ExtractPublicPropTypes<typeof pptxEditorToolbarProps>;

/**
 * The toolbar root. `toolbar-parts.ts` attaches the parts and exports it as
 * `PptxEditorToolbar`. With the preset (default), its arrangement derives from
 * `CHROME_GROUPS` (the groups that act on the selection, `toolbarGroupsFor`), and
 * `ToolbarSlot`, named part, and `ToolbarGroup` children change it; with `preset`
 * false, the default slot is the whole bar. Hiding a control never disables its command.
 */
export const ToolbarRoot = defineComponent({
  name: 'PptxEditorToolbar',
  props: pptxEditorToolbarProps,
  setup(props, { slots }) {
    const { t } = useTranslation();
    // The bar shows the groups that act on the selection, as in React (`toolbarContextOf`).
    const selection = useEditorState(toolbarContextOf);
    // A focused speaker notes field takes the text formatting (`NOTES_SLOTS`), as in React.
    const notes = useNotesTarget(usePptxEditor());
    const context = computed(() => (notes.value ? 'notes' : selection.value));
    return () =>
      props.preset
        ? h(
            PresetBar,
            {
              groups: toolbarGroupsFor(context.value, { contextual: props.contextual }),
              label: t('toolbar.ariaLabel'),
              overflow: props.overflow,
              inBar: (slot: string) => isToolbarBarSlot(slot, { contextual: props.contextual }),
              class: props.className,
            },
            { default: () => slots.default?.() }
          )
        : h(HandBar, { class: props.className }, { default: () => slots.default?.() });
  },
});

/**
 * A preset bar over `groups`: the registry's controls with the host's `ToolbarSlot`, named
 * part, and `ToolbarGroup` overrides, collapsed into a More menu by priority. The toolbar draws
 * the formatting groups; each ribbon tab draws its own groups. `inBar` says which slots the bar
 * draws (default: the slots of `groups`); another slot goes after the groups. @internal
 */
export const PresetBar = defineComponent({
  name: 'PptxPresetBar',
  props: {
    groups: { type: Array as PropType<readonly ChromeGroup[]>, required: true },
    label: { type: String, required: true },
    /** Whether groups that do not fit move into the More menu; false wraps the bar. */
    overflow: { type: Boolean, default: false },
    inBar: { type: Function as PropType<(slot: string) => boolean>, default: undefined },
  },
  setup(props, { slots }) {
    const { t } = useTranslation();
    return () => {
      const barGroups = props.groups;
      const own = new Set(barGroups.flatMap((g) => g.controls.map((c) => `${g.id}.${c.id}`)));
      const overrides = collectOverrides(
        slots.default?.(),
        props.inBar ?? ((slot) => own.has(slot))
      );
      const items: BarItem[] = [];
      for (const entry of arrangeToolbarGroups(overrides.groups, barGroups, PARTS)) {
        const override = entry.override;
        const id = entry.kind === 'built-in' ? entry.group.id : entry.override.id;
        const placement = {
          priority: toolbarGroupPriority(id, override?.priority, barGroups),
          pinned: override?.pinned === true,
          className: override?.className,
          overflowContent: override?.overflowContent,
        };
        if (entry.kind === 'host') {
          // A group with nothing to render leaves the bar, as in React and DOCX Editor.
          if (!hasContent(entry.override.content)) continue;
          const content = entry.override.content;
          items.push({
            ...placement,
            id: entry.override.id,
            label: toolbarGroupLabel(entry.override, (key) => translateKeyOrText(t, key)),
            content: () => content ?? [],
          });
          continue;
        }
        const group = entry.group;
        const shown = group.controls.filter(
          (c) => !overrides.slots.get(`${group.id}.${c.id}`)?.hidden
        );
        if (!shown.length && !hasContent(override?.content)) continue;
        const controls = (): VNodeArrayChildren =>
          shown.map((c) => {
            const spec = overrides.slots.get(`${group.id}.${c.id}`);
            return spec
              ? h(SlotControl, { key: c.id, spec })
              : h(ToolbarControl, { key: c.id, group: group.id, control: c });
          });
        const extra = override?.content;
        const menu = placement.overflowContent;
        items.push({
          ...placement,
          id: group.id,
          label: t(group.labelKey),
          content: () => [...controls(), ...(extra ?? [])],
          menuContent: menu
            ? () => [...controls(), menu() as VNodeArrayChildren[number]]
            : undefined,
        });
      }
      return h(PresetToolbar, {
        items,
        extras: overrides.extras,
        overflow: props.overflow,
        label: props.label,
      });
    };
  },
});
