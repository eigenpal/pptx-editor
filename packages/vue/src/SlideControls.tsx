/** @jsxImportSource vue */
// Slide chrome of the Home tab: the layout gallery, the New Slide split button, the Layout
// button, and the section buttons. Each layout tile is the engine's own render of a new slide on
// that layout, with its placeholder prompts (`renderLayoutThumbnail`). The Vue twin of the React
// adapter's `SlideControls.tsx`, on the same core rules (`insert-controls.ts`).

import { defineComponent, ref, watch, type ExtractPublicPropTypes, type PropType } from 'vue';
import {
  tooltipAttributes,
  layoutGalleryEnter,
  layoutGalleryKey,
  layoutGalleryTabStop,
  layoutMasterGroups,
  translateKeyOrText,
  type ChromeControl,
  type ChromeSlotId,
  type SlideLayoutSummary,
} from '@pptx-editor.dev/core/editor';
import { Icon, Popover, ToolButton, keepFocus, useAriaDisabled } from './chrome.tsx';
import { useEditorState, usePptxEditor, useSlideLayouts, useTranslation } from './composables.ts';
import { useSlot } from './slot.ts';
import { buttonRoleProps, renderAsChild, useSlotOverride } from './slot-override.ts';

const LayoutThumb = defineComponent({
  name: 'PptxLayoutThumb',
  props: {
    id: { type: String, required: true as const },
    ratio: { type: Number, required: true as const },
  },
  setup(props) {
    const editor = usePptxEditor();
    const box = ref<HTMLSpanElement | null>(null);
    // Repaint when the document changes (a theme or master edit changes every layout) and when
    // fonts load: the slide summaries change with the font generation, as the rail's do. The
    // engine skips a paint whose package, layout, and fonts are the same.
    const revision = useEditorState((s) => s.revision);
    const slides = useEditorState((s) => s.slides);
    watch(
      [editor, () => props.id, revision, slides, box],
      () => {
        if (editor.value && box.value) editor.value.renderLayoutThumbnail(props.id, box.value);
      },
      { flush: 'post', immediate: true }
    );
    return () => (
      <span
        ref={box}
        class="pptx-layout-gallery__thumb"
        style={{ aspectRatio: String(1 / props.ratio) }}
        aria-hidden="true"
      />
    );
  },
});

const layoutGalleryProps = {
  /** The layout to mark as the current one (`menuitemradio`, checked). Absent: plain rows. */
  value: { type: String, default: undefined },
  /** The layouts to show. Default: every layout of the deck (`useSlideLayouts`). */
  layouts: { type: Array as PropType<readonly SlideLayoutSummary[]>, default: undefined },
  /** Show the tiles but take no pick. Default: true while the deck is open for viewing. */
  disabled: { type: Boolean, default: undefined },
  /**
   * Why the gallery takes no pick, for the tiles' tooltips: a catalog key (such as
   * `useEditorCommand(slotId).disabledReason`) or text. Default: the viewing reason.
   */
  disabledReason: { type: String as PropType<string | null>, default: undefined },
  /** A class name added to the gallery, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `LayoutGallery`, the twin of React's `LayoutGalleryProps`. @public */
export type LayoutGalleryProps = ExtractPublicPropTypes<typeof layoutGalleryProps>;

/**
 * A grid of the deck's layouts, as the Layout and New Slide galleries show them: one group per
 * slide master, headed by its name, and each tile the engine's render of a new slide on the
 * layout, named below. The tiles are menu rows (`menuitem`, or `menuitemradio` with `value`);
 * arrow keys, Home, and End move between them, so the gallery works in any menu or popup of the
 * host. Emits `change` with the picked layout's id; `className` (or a host `class`) goes on the
 * gallery. @public
 */
export const LayoutGallery = defineComponent({
  name: 'LayoutGallery',
  // The host class goes on the gallery, not on the empty-deck note, as in React.
  inheritAttrs: false,
  props: layoutGalleryProps,
  emits: { change: (_layoutId: string) => true },
  setup(props, { emit, attrs }) {
    const deck = useSlideLayouts();
    const viewing = useEditorState((s) => s.mode === 'view');
    const ratio = useEditorState((s) => s.slideHeight / s.slideWidth);
    const { t } = useTranslation();
    let entered = false;
    const off = (): boolean => props.disabled ?? viewing.value;
    const reasonText = (): string | null =>
      props.disabledReason
        ? translateKeyOrText(t, props.disabledReason)
        : viewing.value
          ? t('disabled.viewing')
          : null;
    // Every tile says the one reason through `aria-describedby`.
    const offTiles = useAriaDisabled(off, reasonText);
    return () => {
      const layouts = props.layouts ?? deck.value;
      if (!layouts.length)
        return <p class="pptx-layout-gallery__empty">{t('layoutGallery.empty')}</p>;
      const current = props.value;
      // One tile takes Tab focus: the current layout, or the first.
      const focusable = layoutGalleryTabStop(layouts, current);
      const isOff = off();
      const reason = reasonText();
      return (
        <div
          {...attrs}
          class={['pptx-layout-gallery', attrs.class, props.className]}
          onFocusin={(e: FocusEvent) => {
            // A menu opened from the keyboard focuses its first row: start on the current layout.
            if (entered) return;
            entered = true;
            layoutGalleryEnter(e.currentTarget as Element, e.target);
          }}
          onKeydown={(e: KeyboardEvent) => {
            // Left from the first column leaves the gallery: a submenu then closes.
            if (layoutGalleryKey(e.currentTarget as HTMLElement, e.target, e.key)) {
              e.preventDefault();
              e.stopPropagation();
            }
          }}
        >
          {layoutMasterGroups(layouts).map((m) => (
            <div key={m.id} class="pptx-layout-gallery__group" role="group" aria-label={m.name}>
              <div class="pptx-layout-gallery__header" aria-hidden="true">
                {m.name}
              </div>
              <div class="pptx-layout-gallery__grid" data-columns={3}>
                {m.layouts.map((l) => (
                  <button
                    key={l.id}
                    type="button"
                    role={current === undefined ? 'menuitem' : 'menuitemradio'}
                    {...(current === undefined ? {} : { 'aria-checked': l.id === current })}
                    tabindex={l.id === focusable ? 0 : -1}
                    class="pptx-layout-gallery__item"
                    data-layout-type={l.type}
                    data-current={l.id === current ? '' : undefined}
                    {...offTiles.attrs.value}
                    {...tooltipAttributes(l.name, isOff, reason)}
                    onClick={() => {
                      if (!isOff) emit('change', l.id);
                    }}
                  >
                    <LayoutThumb id={l.id} ratio={ratio.value} />
                    <span class="pptx-layout-gallery__name">{l.name}</span>
                  </button>
                ))}
              </div>
            </div>
          ))}
          {offTiles.note()}
        </div>
      );
    };
  },
});

const slotPartProps = {
  slot: { type: String as PropType<ChromeSlotId>, required: true as const },
  control: { type: Object as PropType<ChromeControl>, required: true as const },
} as const;

/** The props of `NewSlideControl`, the twin of React's `NewSlideControlProps`. @internal */
export type NewSlideControlProps = ExtractPublicPropTypes<typeof slotPartProps>;

/**
 * Home > New Slide: the button adds a slide on the default next layout; the arrow opens the
 * layout gallery. The toolbar's part for `slide.add`; hosts place it with `ToolbarSlot`.
 * @internal
 */
export const NewSlideControl = defineComponent({
  name: 'NewSlideControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason } = useSlot(
      () => props.slot,
      () => props.control
    );
    const caretOff = useAriaDisabled(() => !cmd.isEnabled.value, reason);
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const anchor = ref<HTMLButtonElement | null>(null);
    const open = ref(false);
    return () => {
      // The caret is part of the split button: host overrides go to the main button, as for the
      // color pickers, so the caret keeps its own class.
      const choose = t('layoutGallery.choose');
      return (
        <span class="pptx-toolbar__group">
          <ToolButton
            slot={props.slot}
            label={label.value}
            paths={props.control.paths}
            disabled={!cmd.isEnabled.value}
            disabledReason={reason.value}
            onClick={() => editor.value?.exec({ type: 'addSlide' })}
          />
          <button
            ref={anchor}
            type="button"
            class="pptx-toolbar__button pptx-toolbar__caret-button"
            data-disabled={cmd.isEnabled.value ? undefined : ''}
            aria-label={choose}
            aria-haspopup="menu"
            aria-expanded={open.value}
            {...tooltipAttributes(choose, !cmd.isEnabled.value, reason.value)}
            {...caretOff.attrs.value}
            onMousedown={keepFocus}
            onClick={() => {
              if (cmd.isEnabled.value) open.value = !open.value;
            }}
          >
            <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
          </button>
          {caretOff.note()}
          {open.value ? (
            <Popover anchor={anchor.value} onClose={() => (open.value = false)} label={label.value}>
              {{
                default: () => (
                  <>
                    <LayoutGallery
                      onChange={(id: string) => {
                        open.value = false;
                        cmd.execute(id);
                      }}
                    />
                  </>
                ),
              }}
            </Popover>
          ) : null}
        </span>
      );
    };
  },
});

/** The props of `LayoutControl`, the twin of React's `LayoutControlProps`. @internal */
export type LayoutControlProps = ExtractPublicPropTypes<typeof slotPartProps>;

/**
 * Home > Layout: the layout gallery, with the current slide's layout marked. The trigger is one
 * button, so it takes the slot's `icon` and `asChild` as `ToolButton` does. The toolbar's part
 * for `slide.layout`; hosts place it with `ToolbarSlot`. @internal
 */
export const LayoutControl = defineComponent({
  name: 'LayoutControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, tip, reason } = useSlot(
      () => props.slot,
      () => props.control
    );
    const override = useSlotOverride(props.slot);
    const anchor = ref<HTMLElement | null>(null);
    const open = ref(false);
    // A disabled trigger stays focusable and says why (`aria-disabled`).
    const off = useAriaDisabled(() => !cmd.isEnabled.value, reason);
    const toggle = (el: Element): void => {
      if (!cmd.isEnabled.value || !(el instanceof HTMLElement)) return;
      anchor.value = el;
      open.value = !open.value;
    };
    return () => {
      const disabled = !cmd.isEnabled.value;
      const attributes = {
        class: ['pptx-toolbar__button pptx-toolbar__button--wide', override?.className]
          .filter(Boolean)
          .join(' '),
        'data-slot': props.slot,
        'data-disabled': disabled ? '' : undefined,
        'aria-label': label.value,
        'aria-haspopup': 'menu' as const,
        'aria-expanded': open.value,
        ...tip.value,
        ...off.attrs.value,
        onMousedown: keepFocus,
        onClick: (e: MouseEvent) => toggle(e.currentTarget as Element),
      };
      const host = override?.asChild;
      return [
        host ? (
          // A host element that is not a button gets the button role, a tab stop, and Enter and
          // Space, as `ToolButton` gives it.
          renderAsChild(host, {
            ...attributes,
            ...(host.type === 'button' ? { type: 'button' } : buttonRoleProps(toggle)),
          })
        ) : (
          <button type="button" {...attributes}>
            {override?.icon !== undefined ? override.icon : <Icon paths={props.control.paths} />}
            <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
          </button>
        ),
        off.note(),
        open.value ? (
          <Popover anchor={anchor.value} onClose={() => (open.value = false)}>
            {{
              default: () => (
                <>
                  <LayoutGallery
                    value={typeof cmd.value.value === 'string' ? cmd.value.value : ''}
                    onChange={(id: string) => {
                      open.value = false;
                      cmd.execute(id);
                    }}
                  />
                </>
              ),
            }}
          </Popover>
        ) : null,
      ];
    };
  },
});

/** A section button: one `ToolButton` that runs the slot's command. */
function sectionControl(name: string) {
  return defineComponent({
    name,
    props: slotPartProps,
    setup(props) {
      const { cmd, label, reason } = useSlot(
        () => props.slot,
        () => props.control
      );
      return () => (
        <ToolButton
          slot={props.slot}
          label={label.value}
          paths={props.control.paths}
          disabled={!cmd.isEnabled.value}
          disabledReason={reason.value}
          onClick={() => cmd.execute()}
        />
      );
    },
  });
}

/**
 * Home > Section > Add Section: adds a section at the current slide and starts renaming it, so
 * it asks for the name, where a rail or a sorter can show the name field (`runSectionSlot`). The
 * toolbar's part for `section.add`. @internal
 */
export const SectionAddControl = sectionControl('PptxSectionAddControl');

/**
 * Home > Section > Rename Section: opens the name field of the current slide's section in the
 * rail or the sorter. Without one, it is off and says why (`disabled.noSectionPane`). The
 * toolbar's part for `section.rename`. @internal
 */
export const SectionRenameControl = sectionControl('PptxSectionRenameControl');
