/** @jsxImportSource vue */
// The named parts of the slide rail's action row: the Vue twins of the React adapter's
// `rail-parts.tsx`, with the same markup and classes. Each part takes `hidden`, `icon`,
// `className` (or a host `class`), and `asChild` (the one element of the default slot renders
// the button). A named part child of `PptxEditorSlideRail` replaces its
// part in place, other children go after Delete, and `preset: false` renders only the
// default slot. The parts read the rail's state from an injection, so they render only inside
// a rail.

import {
  Comment,
  defineComponent,
  inject,
  isVNode,
  normalizeClass,
  provide,
  shallowRef,
  type ExtractPublicPropTypes,
  type InjectionKey,
  type PropType,
  type VNode,
  type VNodeArrayChildren,
} from 'vue';
import type {
  PptxEditorInstance,
  RailOrientationChrome,
  SlideStep,
} from '@pptx-editor.dev/core/editor';
import { railOverflowPoint, warnAsChildChild } from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { ToolButton } from './chrome.tsx';
import { ContextMenu, type ContextMenuPoint } from './ContextMenu.tsx';
import { MenuRow, menuCloseKey } from './menu-row.tsx';

const partProps = {
  /** Remove the button from the rail. */
  hidden: { type: Boolean, default: false },
  /** Icon content in place of the packaged icon, about 18 pixels. */
  icon: { type: Object as PropType<VNode>, default: undefined },
  /** A class name added to the button, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
  /**
   * Render the one element of the default slot as the button: it gets the part's label, state,
   * disabled reason, and click, as `ToolbarAction` does.
   */
  asChild: { type: Boolean, default: false },
} as const;

/**
 * Props of every named part of the slide rail's action row, the twin of React's
 * `SlideRailPartProps`. @public
 */
export type SlideRailPartProps = ExtractPublicPropTypes<typeof partProps>;

/** What the rail hands its action parts; the getters keep the parts reactive. @internal */
export interface RailActions {
  readonly t: (key: TranslationKey) => string;
  readonly editor: () => PptxEditorInstance | null;
  /** Why the slide buttons are off, or null when they work. */
  readonly why: () => string | null;
  readonly selected: () => readonly number[];
  readonly count: () => number;
  readonly chrome: () => RailOrientationChrome;
  readonly step: (step: -1 | 1) => SlideStep | null;
  readonly moveBy: (step: -1 | 1) => void;
  readonly galleryOpen: () => boolean;
  readonly toggleGallery: () => void;
  readonly listId: string;
  /** Whether the rail is open: the action row shows, or the folded strip. */
  readonly open: () => boolean;
  /** The rail's `toggle`: `false` removes Hide thumbnails and Show thumbnails. */
  readonly toggle: () => boolean;
  readonly collapse: () => void;
  readonly expand: () => void;
}

/** Provided by `PptxEditorSlideRail` for its action parts. @internal */
export const railActionsKey: InjectionKey<RailActions> = Symbol('pptxRailActions');

/** Inside the overflow menu of a narrow rail, the parts render as menu rows. */
const railOverflowKey: InjectionKey<boolean> = Symbol('pptxRailOverflow');

interface RailButton {
  readonly base: string;
  readonly label: string;
  readonly icon: string;
  readonly onClick: () => void;
  readonly disabled?: boolean;
  readonly disabledReason?: string | null;
  readonly hasPopup?: 'menu';
  readonly expanded?: boolean;
  readonly controls?: string;
}

/** Defines one part: `button` describes it from the rail's state, as React's part does. */
function definePart(
  name: string,
  button: (rail: RailActions) => RailButton,
  shows: (rail: RailActions) => boolean = () => true
) {
  return defineComponent({
    name,
    inheritAttrs: false,
    props: partProps,
    setup(props, { attrs, slots }) {
      const rail = inject(railActionsKey, null);
      const inMenu = inject(railOverflowKey, false);
      return () => {
        if (!rail || props.hidden || !shows(rail)) return null;
        const { base, icon, ...rest } = button(rail);
        const host = props.icon;
        // In the overflow menu of a narrow rail: the same action as a menu row.
        if (inMenu)
          return (
            <MenuRow
              label={rest.label}
              icon={host === undefined ? icon : host}
              disabled={!!rest.disabled}
              reason={rest.disabled ? (rest.disabledReason ?? null) : null}
              className={normalizeClass([base, attrs.class, props.className])}
              onSelect={rest.onClick}
            />
          );
        // One element in the default slot, or the packaged button (with a warning).
        const only = props.asChild ? (slots.default?.() ?? []).filter(isVNode) : [];
        const element = only.length === 1 && typeof only[0]!.type !== 'symbol' ? only[0]! : null;
        if (props.asChild && !element)
          warnAsChildChild(rest.label, { slot: name, group: 'PptxEditorSlideRail' });
        return (
          <ToolButton
            {...rest}
            host={element}
            icon={host === undefined ? icon : undefined}
            className={normalizeClass([base, attrs.class, props.className])}
          >
            {host === undefined ? undefined : () => host}
          </ToolButton>
        );
      };
    },
  });
}

/** Add slide, after the current slide, as React's `SlideRailAddSlide`. @public */
export const SlideRailAddSlide = definePart('SlideRailAddSlide', (rail) => ({
  base: 'pptx-rail__add',
  label: rail.t('rail.addSlide'),
  icon: 'add',
  disabled: !!rail.why(),
  disabledReason: rail.why(),
  onClick: () => rail.editor()?.exec({ type: 'addSlide' }),
}));

/** The layout gallery, as React's `SlideRailChooseLayout`. @public */
export const SlideRailChooseLayout = definePart('SlideRailChooseLayout', (rail) => ({
  base: 'pptx-rail__layouts',
  label: rail.t('rail.chooseLayout'),
  icon: 'note_add',
  hasPopup: 'menu',
  expanded: rail.galleryOpen(),
  disabled: !!rail.why(),
  disabledReason: rail.why(),
  onClick: rail.toggleGallery,
}));

/** Duplicate the selected slides, as React's `SlideRailDuplicate`. @public */
export const SlideRailDuplicate = definePart('SlideRailDuplicate', (rail) => ({
  base: 'pptx-rail__duplicate',
  label: rail.t('rail.duplicateSlide'),
  icon: 'content_copy',
  disabled: !!rail.why(),
  disabledReason: rail.why(),
  onClick: () => rail.editor()?.exec({ type: 'duplicateSlides', indices: rail.selected() }),
}));

/** Move the selected slides earlier, as React's `SlideRailMoveEarlier`. @public */
export const SlideRailMoveEarlier = definePart('SlideRailMoveEarlier', (rail) => ({
  base: 'pptx-rail__move',
  label: rail.t(rail.chrome().moveEarlierLabel),
  icon: rail.chrome().moveEarlierIcon,
  disabled: !!rail.why() || rail.step(-1) === null,
  disabledReason: rail.why() ?? rail.t('disabled.slideEdge'),
  onClick: () => rail.moveBy(-1),
}));

/** Move the selected slides later, as React's `SlideRailMoveLater`. @public */
export const SlideRailMoveLater = definePart('SlideRailMoveLater', (rail) => ({
  base: 'pptx-rail__move',
  label: rail.t(rail.chrome().moveLaterLabel),
  icon: rail.chrome().moveLaterIcon,
  disabled: !!rail.why() || rail.step(1) === null,
  disabledReason: rail.why() ?? rail.t('disabled.slideEdge'),
  onClick: () => rail.moveBy(1),
}));

/** Delete the selected slides, as React's `SlideRailDelete`. @public */
export const SlideRailDelete = definePart('SlideRailDelete', (rail) => ({
  base: 'pptx-rail__delete',
  label: rail.t('rail.deleteSlide'),
  icon: 'delete',
  disabled: !!rail.why() || rail.selected().length >= rail.count(),
  disabledReason: rail.why() ?? rail.t('disabled.lastSlide'),
  onClick: () => rail.editor()?.exec({ type: 'deleteSlides', indices: rail.selected() }),
}));

/** Hide thumbnails, as React's `SlideRailCollapse`. `toggle: false` removes it. @public */
export const SlideRailCollapse = definePart(
  'SlideRailCollapse',
  (rail) => ({
    base: 'pptx-rail__collapse',
    label: rail.t('rail.collapse'),
    icon: rail.chrome().collapseIcon,
    expanded: true,
    controls: rail.listId,
    onClick: rail.collapse,
  }),
  (rail) => rail.toggle() && rail.open()
);

/**
 * Show thumbnails, the one button of the folded rail strip, as React's `SlideRailExpand`.
 * `toggle: false` removes it. @public
 */
export const SlideRailExpand = definePart(
  'SlideRailExpand',
  (rail) => ({
    base: 'pptx-rail__collapse pptx-rail__expand',
    label: rail.t('rail.expand'),
    icon: rail.chrome().expandIcon,
    expanded: false,
    onClick: rail.expand,
  }),
  (rail) => rail.toggle() && !rail.open()
);

/** The overflow button of the rail whose slide list has the id `listId`. */
function overflowButtonOf(listId: string): HTMLElement | null {
  if (typeof document === 'undefined') return null;
  const rail = document.getElementById(listId)?.closest('.pptx-rail');
  return rail?.querySelector<HTMLElement>('.pptx-rail__more') ?? null;
}

/**
 * The overflow button of a narrow rail: the secondary actions (layouts, duplicate, move,
 * delete) in one menu, as React's `SlideRailOverflow`.
 */
const SlideRailOverflow = defineComponent({
  name: 'PptxSlideRailOverflow',
  setup(_props, { slots }) {
    const rail = inject(railActionsKey, null);
    const menu = shallowRef<{ readonly at: ContextMenuPoint; readonly owner: HTMLElement } | null>(
      null
    );
    const close = (): void => {
      menu.value = null;
    };
    provide(menuCloseKey, close);
    provide(railOverflowKey, true);
    return () => {
      if (!rail) return null;
      const open = menu.value;
      return [
        <ToolButton
          label={rail.t('rail.moreActions')}
          icon="more_vert"
          className="pptx-rail__more"
          hasPopup="menu"
          expanded={open !== null}
          onClick={() => {
            const owner = overflowButtonOf(rail.listId);
            menu.value = menu.value || !owner ? null : { at: railOverflowPoint(owner), owner };
          }}
        />,
        open ? (
          <ContextMenu
            at={open.at}
            owner={open.owner}
            // Focus goes to the first row, as the toolbar's menus take focus when they open.
            keyboard
            label={rail.t('rail.moreActions')}
            onClose={close}
          >
            {{ default: () => slots.default?.() }}
          </ContextMenu>
        ) : null,
      ];
    };
  },
});

/** The named parts, in their packaged order. Collapse is last and follows the host children. */
const PARTS: readonly unknown[] = [
  SlideRailAddSlide,
  SlideRailChooseLayout,
  SlideRailDuplicate,
  SlideRailMoveEarlier,
  SlideRailMoveLater,
  SlideRailDelete,
];

function flatten(children: VNodeArrayChildren | undefined, out: VNode[] = []): VNode[] {
  for (const child of children ?? []) {
    if (Array.isArray(child)) flatten(child, out);
    // A `v-if` that is false leaves a comment: it renders no button, as React's `null`.
    else if (isVNode(child) && child.type !== Comment) {
      if (Array.isArray(child.children) && typeof child.type === 'symbol')
        flatten(child.children as VNodeArrayChildren, out);
      else out.push(child);
    }
  }
  return out;
}

/**
 * A child that renders no button in the action row: a part with `hidden` (`true`, or the bare
 * attribute), Show thumbnails (only in the folded strip), and Hide thumbnails under
 * `toggle: false`.
 */
function isHidden(vnode: VNode, toggle: boolean): boolean {
  if (vnode.type === SlideRailExpand || (!toggle && vnode.type === SlideRailCollapse)) return true;
  const hidden = (vnode.props as Record<string, unknown> | null)?.hidden;
  return hidden !== undefined && hidden !== false;
}

/** The folded strip's content: the host's `SlideRailExpand` child, or the packaged one. @internal */
export function railExpandContent(children: VNodeArrayChildren | undefined): VNode {
  return flatten(children).find((vnode) => vnode.type === SlideRailExpand) ?? <SlideRailExpand />;
}

/**
 * The rail's action row content and the number of buttons it shows, as React's
 * `railActionContent`. A `folded` (narrow) rail shows New slide and one overflow menu for the
 * other parts. @internal
 */
export function railActionContent(
  children: VNodeArrayChildren | undefined,
  preset: boolean,
  toggle: boolean,
  folded = false
): { readonly content: VNode[]; readonly shown: number } {
  const nodes = flatten(children);
  if (!preset) return { content: nodes, shown: nodes.filter((n) => !isHidden(n, toggle)).length };
  const overrides = new Map<unknown, VNode>();
  const extras: VNode[] = [];
  for (const vnode of nodes)
    if (PARTS.includes(vnode.type) || vnode.type === SlideRailCollapse)
      overrides.set(vnode.type, vnode);
    else if (vnode.type !== SlideRailExpand) extras.push(vnode);
  const add = overrides.get(SlideRailAddSlide) ?? <SlideRailAddSlide />;
  const secondary: VNode[] = [
    overrides.get(SlideRailChooseLayout) ?? <SlideRailChooseLayout />,
    overrides.get(SlideRailDuplicate) ?? <SlideRailDuplicate />,
    overrides.get(SlideRailMoveEarlier) ?? <SlideRailMoveEarlier />,
    overrides.get(SlideRailMoveLater) ?? <SlideRailMoveLater />,
    overrides.get(SlideRailDelete) ?? <SlideRailDelete />,
  ];
  // A narrow rail keeps New slide and folds the rest into one menu; a hidden part stays hidden.
  const folds = folded && secondary.some((n) => !isHidden(n, toggle));
  const list: VNode[] = folds
    ? [add, <SlideRailOverflow>{{ default: () => secondary }}</SlideRailOverflow>, ...extras]
    : [add, ...secondary, ...extras];
  if (toggle) list.push(overrides.get(SlideRailCollapse) ?? <SlideRailCollapse />);
  return { content: list, shown: list.filter((n) => !isHidden(n, toggle)).length };
}
