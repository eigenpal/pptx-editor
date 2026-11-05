// The ribbon: the Vue twin of the React adapter's `PptxEditor.Ribbon`. Tabs over the toolbar's groups from core
// (`RIBBON_TABS`, `ribbonTabs`): each tab panel is the toolbar's preset bar with that tab's
// groups, so `ToolbarSlot`, named parts, `ToolbarGroup`, and `ToolbarAction` work inside a tab.
// The markup and classes match React's, so the one core stylesheet styles both.

import {
  defineComponent,
  h,
  isVNode,
  normalizeClass,
  ref,
  watch,
  nextTick,
  Comment,
  Text,
  useId,
  type ExtractPublicPropTypes,
  type PropType,
  type VNode,
  type VNodeArrayChildren,
} from 'vue';
import {
  arrangeRibbonTabs,
  ribbonInsertTab,
  ribbonShownTab,
  ribbonTabs,
  translateKeyOrText,
  warnRibbonChild,
  type RibbonTabId,
  type RibbonTabLayout,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, useTranslation } from './composables.ts';
import { PresetBar } from './toolbar.ts';

const ribbonTabProps = {
  /** A built-in tab id changes that tab; any other id adds a host tab. */
  id: { type: String as PropType<RibbonTabId | (string & {})>, required: true as const },
  /** The tab's title. Wins over `labelKey`. */
  label: { type: String, default: undefined },
  /** A translation key for the tab's title. */
  labelKey: { type: String, default: undefined },
  /** Remove the tab. */
  hidden: { type: Boolean, default: false },
  /** For a host tab: the tab it follows. */
  after: { type: String as PropType<RibbonTabId | (string & {})>, default: undefined },
  /** A class name added to the tab's panel, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `RibbonTab`, the twin of React's `RibbonTabProps`. @public */
export type RibbonTabProps = ExtractPublicPropTypes<typeof ribbonTabProps>;

/** One tab of the ribbon: a description `PptxEditorRibbon` reads. Its slot is toolbar children. @public */
export const RibbonTab = defineComponent({
  name: 'RibbonTab',
  props: ribbonTabProps,
  setup() {
    return () => null;
  },
});

const ribbonProps = {
  /** `true` (default): the tabs of `RIBBON_TABS`, changed by `RibbonTab` children. */
  preset: { type: Boolean, default: true },
  /** The selected tab, controlled. Use with `v-model:tab`, or the `tabChange` event. */
  tab: { type: String as PropType<RibbonTabId | (string & {})>, default: undefined },
  /** The tab selected first, uncontrolled. */
  defaultTab: { type: String as PropType<RibbonTabId | (string & {})>, default: 'home' },
  /** `false`: a tab's groups wrap instead of moving into the More menu. Default `true`. */
  overflow: { type: Boolean, default: true },
  /** A class name added to the ribbon, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `PptxEditorRibbon`, the twin of React's `RibbonProps`. @public */
export type RibbonProps = ExtractPublicPropTypes<typeof ribbonProps>;

interface TabSpec {
  readonly id: string;
  readonly label?: string;
  readonly labelKey?: string;
  readonly hidden: boolean;
  readonly after?: string;
  readonly className: string | undefined;
  readonly content: VNodeArrayChildren | undefined;
}

function flatten(children: VNodeArrayChildren | undefined, out: VNode[] = []): VNode[] {
  for (const child of children ?? []) {
    if (Array.isArray(child)) flatten(child, out);
    else if (isVNode(child)) {
      if (Array.isArray(child.children) && typeof child.type === 'symbol')
        flatten(child.children as VNodeArrayChildren, out);
      else out.push(child);
    }
  }
  return out;
}

function readSpec(vnode: VNode): TabSpec {
  const props = (vnode.props ?? {}) as Record<string, unknown>;
  const pick = (camel: string): unknown =>
    camel in props ? props[camel] : props[camel.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`)];
  const children = vnode.children as { default?: () => VNodeArrayChildren } | null;
  const label = pick('label');
  const labelKey = pick('labelKey');
  const after = pick('after');
  const hidden = pick('hidden');
  return {
    id: String(pick('id')),
    ...(typeof label === 'string' ? { label } : {}),
    ...(typeof labelKey === 'string' ? { labelKey } : {}),
    ...(typeof after === 'string' ? { after } : {}),
    hidden: hidden === '' || hidden === true,
    className: normalizeClass([props.class, pick('className')]) || undefined,
    content: children && typeof children === 'object' ? children.default?.() : undefined,
  };
}

const sameLayout = (a: readonly RibbonTabLayout[], b: readonly RibbonTabLayout[]): boolean => {
  const key = (l: readonly RibbonTabLayout[]) =>
    l.map((t) => `${t.tab.id}:${t.groups.map((g) => g.id).join(',')}`).join('|');
  return key(a) === key(b);
};

/**
 * The ribbon: the Quick Access Toolbar, the tab names, and the selected tab's groups, with
 * contextual tabs while the selection needs them, as React's `PptxEditor.Ribbon`. Emits
 * `update:tab` (for `v-model:tab`) and `tabChange`. The tabs are attached as `PptxEditorRibbon.Tab`. @public
 */
export const PptxEditorRibbon = Object.assign(
  defineComponent({
    name: 'PptxEditorRibbon',
    props: ribbonProps,
    emits: {
      'update:tab': (_tab: string) => true,
      tabChange: (_tab: string) => true,
    },
    setup(props, { emit, slots }) {
      const { t } = useTranslation();
      const baseId = useId();
      const layout = useEditorState((s) => ribbonTabs(s), sameLayout);
      const own = ref<string>(props.defaultTab);
      const names = new Map<string, HTMLElement>();
      const root = ref<HTMLElement | null>(null);
      // Whether the keyboard focus is in the ribbon, as in React: a removed panel sends no
      // `focusout` with a target, so the flag stays set until the focus goes elsewhere.
      let focusInside = false;
      const select = (id: string): void => {
        if (props.tab === undefined) own.value = id;
        emit('update:tab', id);
        emit('tabChange', id);
      };
      // The shown tab after the last render, and whether the wanted tab was gone.
      const shown = ref<{ id: string | undefined; fellBack: boolean }>({
        id: undefined,
        fellBack: false,
      });
      // A tab that went away gives way to Home, as in React: the ribbon selects it, tells the
      // host, and moves the focus to its name.
      watch(
        () => `${shown.value.fellBack}|${shown.value.id ?? ''}`,
        async () => {
          const { id, fellBack } = shown.value;
          if (!fellBack || id === undefined) return;
          select(id);
          await nextTick();
          const active = root.value?.ownerDocument.activeElement;
          if (focusInside && (!active || active === active.ownerDocument.body))
            names.get(id)?.focus();
        }
      );
      // An insert opens the new object's Format tab. The tabs
      // of the render after the insert are known in the watch, after the render.
      const insertions = useEditorState((s) => s.insertions);
      let insertTab: string | null = null;
      watch(insertions, async () => {
        await nextTick();
        if (insertTab !== null && insertTab !== (props.tab ?? own.value)) select(insertTab);
      });
      return () => {
        const specs = new Map<string, TabSpec>();
        for (const vnode of flatten(slots.default?.()))
          if (vnode.type === RibbonTab) {
            const spec = readSpec(vnode);
            specs.set(spec.id, spec);
          } else if (
            vnode.type !== Comment &&
            (vnode.type !== Text || String(vnode.children ?? '').trim() !== '')
          )
            // Text warns as in React, where a string child is not a tab; blank text does not.
            warnRibbonChild('PptxEditorRibbon', 'PptxEditorRibbon.Tab');
        const { quickAccess, tabs } = arrangeRibbonTabs(
          layout.value,
          specs,
          props.preset,
          (key) => translateKeyOrText(t, key),
          'PptxEditorRibbon.Tab'
        );
        insertTab = ribbonInsertTab(tabs);
        const wanted = props.tab ?? own.value;
        const picked = ribbonShownTab(tabs, wanted);
        const current = picked.shown;
        if (shown.value.id !== current?.id || shown.value.fellBack !== picked.fellBack)
          shown.value = { id: current?.id, fellBack: picked.fellBack };
        const onKey = (e: KeyboardEvent): void => {
          const index = tabs.findIndex((d) => d.id === current?.id);
          const next =
            e.key === 'ArrowRight'
              ? tabs[(index + 1) % tabs.length]
              : e.key === 'ArrowLeft'
                ? tabs[(index - 1 + tabs.length) % tabs.length]
                : e.key === 'Home'
                  ? tabs[0]
                  : e.key === 'End'
                    ? tabs[tabs.length - 1]
                    : undefined;
          if (!next) return;
          e.preventDefault();
          select(next.id);
          names.get(next.id)?.focus();
        };
        const tabId = (id: string) => `${baseId}-tab-${id}`;
        const panelId = `${baseId}-panel`;
        return h(
          'div',
          {
            ref: root,
            class: ['pptx-ribbon', props.className],
            onFocusin: () => {
              focusInside = true;
            },
            onFocusout: (e: FocusEvent) => {
              const next = e.relatedTarget as Node | null;
              if (next && !(e.currentTarget as HTMLElement).contains(next)) focusInside = false;
            },
          },
          [
            h('div', { class: 'pptx-ribbon__bar' }, [
              quickAccess
                ? h('div', { class: 'pptx-ribbon__qat' }, [
                    h(
                      PresetBar,
                      { groups: quickAccess.groups, label: quickAccess.title, overflow: false },
                      () => quickAccess.override?.content
                    ),
                  ])
                : null,
              h(
                'div',
                {
                  role: 'tablist',
                  'aria-label': t('ribbon.ariaLabel'),
                  class: 'pptx-ribbon__tabs',
                  onKeydown: onKey,
                },
                tabs.map((d) => {
                  const selected = d.id === current?.id;
                  return h(
                    'button',
                    {
                      key: d.id,
                      ref: (el: unknown) => {
                        if (el) names.set(d.id, el as HTMLElement);
                        else names.delete(d.id);
                      },
                      type: 'button',
                      role: 'tab',
                      id: tabId(d.id),
                      class: 'pptx-ribbon__tab',
                      'aria-selected': selected,
                      'aria-controls': selected ? panelId : undefined,
                      tabindex: selected ? 0 : -1,
                      'data-tab': d.id,
                      'data-contextual': d.contextual ? '' : undefined,
                      onMousedown: (e: MouseEvent) => e.preventDefault(),
                      onClick: () => select(d.id),
                    },
                    d.title
                  );
                })
              ),
            ]),
            current
              ? h(
                  'div',
                  {
                    role: 'tabpanel',
                    id: panelId,
                    'aria-labelledby': tabId(current.id),
                    class: ['pptx-ribbon__panel', current.override?.className],
                    'data-tab': current.id,
                  },
                  [
                    h(
                      PresetBar,
                      {
                        key: current.id,
                        groups: current.groups,
                        label: current.title,
                        overflow: props.overflow,
                      },
                      () => current.override?.content
                    ),
                  ]
                )
              : null,
          ]
        );
      };
    },
  }),
  { Tab: RibbonTab }
);
