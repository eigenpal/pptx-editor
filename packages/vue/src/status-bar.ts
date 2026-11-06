// The status bar: the Vue twin of the React adapter's `PptxEditor.StatusBar`, with the same
// markup and classes, so the one core stylesheet styles both: "Slide x of y", the Notes button,
// the view buttons, and zoom. The view buttons and zoom are the editor's own `ViewSwitch` and
// `PptxEditorZoomPill`, so the bar adds no second control for the same action. Each part takes
// `hidden` and a host `class`; a named part child replaces its part in place, other children
// follow the Notes button, and `preset: false` renders only the children.

import { chromeTranslateProp, useChromeTranslateScope } from './headless-hooks.ts';
import {
  defineComponent,
  h,
  isVNode,
  normalizeClass,
  type ExtractPublicPropTypes,
  type PropType,
  type VNode,
  type VNodeArrayChildren,
} from 'vue';
import { chromeKeepsFocus, slideStatus, warnAsChildChild } from '@pptx-editor.dev/core/editor';
import { Icon } from './chrome.tsx';
import { useEditorState, useTranslation } from './composables.ts';
import { useNotesPane } from './pane-state.ts';
import { PptxEditorZoomPill } from './panes.tsx';
import { ReadOnlyBadge, ViewSwitch } from './SlideSorter.tsx';
import { ToolbarAction } from './toolbar-parts.ts';

const partProps = {
  /** Remove the part from the status bar. */
  hidden: { type: Boolean, default: false },
  /** A class name added to the part, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of every status bar part: `hidden`, and `className` (or `class`) on the part. @public */
export type StatusBarPartProps = ExtractPublicPropTypes<typeof partProps>;

function keepFocus(e: MouseEvent): void {
  if (chromeKeepsFocus(e.target)) e.preventDefault();
}

/** "Slide 2 of 5", as React's `StatusBarSlideNumber`. @public */
export const StatusBarSlideNumber = defineComponent({
  name: 'StatusBarSlideNumber',
  props: partProps,
  setup(props) {
    const { t } = useTranslation();
    const status = useEditorState(
      (s) => slideStatus(s),
      (a, b) => a?.current === b?.current && a?.total === b?.total
    );
    return () =>
      props.hidden
        ? null
        : h(
            'span',
            {
              class: ['pptx-statusbar__slide', props.className],
              'data-part': 'slideNumber',
              'aria-live': 'polite',
            },
            status.value ? t('statusBar.slide', status.value) : ''
          );
  },
});

/** The part name in warnings about `StatusBarNotes`. */
const NOTES_PART = { slot: 'StatusBar.Notes', group: 'StatusBar' } as const;

const notesPartProps = {
  ...partProps,
  /** Icon content in place of the packaged icon. */
  icon: { type: Object as PropType<VNode>, default: undefined },
  /** Render the one element of the default slot as the button: it gets the label, state, and click. */
  asChild: { type: Boolean, default: false },
} as const;

/**
 * Props of `StatusBarNotes`, the twin of React's `StatusBarNotesProps`: the part props, and the
 * ladder's `icon` and `asChild` (the element goes in the default slot). `className` (or a host
 * `class`) goes on the button. @public
 */
export type StatusBarNotesProps = ExtractPublicPropTypes<typeof notesPartProps>;

/**
 * The Notes button: shows or hides the speaker notes (`useNotesPane`), as React's
 * `StatusBarNotes`. @public
 */
export const StatusBarNotes = defineComponent({
  name: 'StatusBarNotes',
  inheritAttrs: false,
  props: notesPartProps,
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    const notes = useNotesPane();
    return () => {
      if (props.hidden) return null;
      // One element child, or the packaged Notes button with its icon (and a warning that names
      // this part).
      const only = props.asChild ? (slots.default?.() ?? []).filter(isVNode) : [];
      const host = only.length === 1 && typeof only[0]!.type !== 'symbol';
      if (props.asChild && !host) warnAsChildChild('notes', NOTES_PART);
      return h('span', { class: 'pptx-statusbar__notes-part', 'data-part': 'notes' }, [
        h(
          ToolbarAction,
          {
            label: t('statusBar.notes'),
            icon: host ? undefined : (props.icon ?? h(Icon, { name: 'edit_note' })),
            className: normalizeClass(['pptx-statusbar__notes', attrs.class, props.className]),
            active: notes.open.value,
            asChild: host,
            onSelect: notes.toggle,
          },
          host ? { default: () => only } : undefined
        ),
      ]);
    };
  },
});

/**
 * The view buttons: the editor's view switch (`view.presentationView`: Normal and Slide
 * Sorter), the same control as the default editor's status area, as React's `StatusBarViews`.
 * @public
 */
export const StatusBarViews = defineComponent({
  name: 'StatusBarViews',
  props: partProps,
  setup(props) {
    return () =>
      props.hidden
        ? null
        : h('span', { class: ['pptx-statusbar__views', props.className], 'data-part': 'views' }, [
            h(ViewSwitch),
          ]);
  },
});

/** Zoom out, the zoom level with its menu, and zoom in: the editor's `PptxEditorZoomPill`. @public */
export const StatusBarZoom = defineComponent({
  name: 'StatusBarZoom',
  props: partProps,
  setup(props) {
    return () =>
      props.hidden
        ? null
        : h('span', { class: ['pptx-statusbar__zoom', props.className], 'data-part': 'zoom' }, [
            h(PptxEditorZoomPill),
          ]);
  },
});

const PARTS: readonly unknown[] = [
  StatusBarSlideNumber,
  StatusBarNotes,
  StatusBarViews,
  StatusBarZoom,
];

const statusBarProps = {
  ...chromeTranslateProp,
  /** `false`: only the default slot, with no default parts. */
  preset: { type: Boolean, default: true },
  /** A class name added to the bar, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `PptxEditorStatusBar`, the twin of React's `StatusBarProps`. @public */
export type StatusBarProps = ExtractPublicPropTypes<typeof statusBarProps>;

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

/**
 * The status bar: "Slide x of y", the Notes button, the view buttons, and zoom, as React's
 * `PptxEditor.StatusBar`. A named part child replaces its part in place; other children follow
 * the Notes button. The parts are attached as `PptxEditorStatusBar.SlideNumber`, `.Notes`,
 * `.Views`, and `.Zoom`. @public
 */
export const PptxEditorStatusBar = Object.assign(
  defineComponent({
    name: 'PptxEditorStatusBar',
    props: statusBarProps,
    setup(props, { slots }) {
      const { t } = useChromeTranslateScope(() => props.t);
      return () => {
        const children = slots.default?.();
        const bar = (content: VNodeArrayChildren) =>
          h(
            'div',
            {
              role: 'group',
              'aria-label': t('statusBar.ariaLabel'),
              class: ['pptx-statusbar pptx-statusbar--bar', props.className],
              onMousedown: keepFocus,
            },
            content
          );
        if (!props.preset) return bar([...(children ?? []), h(ReadOnlyBadge)]);
        const overrides = new Map<unknown, VNode>();
        const extras: VNode[] = [];
        for (const vnode of flatten(children))
          if (PARTS.includes(vnode.type)) overrides.set(vnode.type, vnode);
          else extras.push(vnode);
        return bar([
          overrides.get(StatusBarSlideNumber) ?? h(StatusBarSlideNumber),
          // Its own part of the bar: hiding the slide number does not hide it.
          h(ReadOnlyBadge),
          h('span', { class: 'pptx-statusbar__spacer' }),
          overrides.get(StatusBarNotes) ?? h(StatusBarNotes),
          ...extras,
          overrides.get(StatusBarViews) ?? h(StatusBarViews),
          overrides.get(StatusBarZoom) ?? h(StatusBarZoom),
        ]);
      };
    },
  }),
  {
    SlideNumber: StatusBarSlideNumber,
    Notes: StatusBarNotes,
    Views: StatusBarViews,
    Zoom: StatusBarZoom,
    Action: ToolbarAction,
  }
);
