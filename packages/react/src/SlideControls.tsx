/** @jsxImportSource react */
// Slide chrome of the Home tab: the layout gallery, the New Slide split button, and the
// Layout button. Each layout tile is the engine's own render of a new slide on that layout,
// with its placeholder prompts (`renderLayoutThumbnail`).

import { memo, useEffect, useRef, useState } from 'react';
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
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { useSlot } from './slot.ts';
import { buttonRoleProps, renderAsChild, useSlotOverride } from './slot-override.ts';

const EMPTY: readonly SlideLayoutSummary[] = [];
const sameLayouts = (a: readonly SlideLayoutSummary[], b: readonly SlideLayoutSummary[]) => a === b;

/**
 * The deck's slide layouts (`EditorSnapshot.layouts`), in the order the galleries list them,
 * reference-stable while they do not change. For a host's own layout picker. @public
 */
export function useSlideLayouts(): readonly SlideLayoutSummary[] {
  return useEditorState((s) => s.layouts ?? EMPTY, sameLayouts);
}

const LayoutThumb = memo(function LayoutThumb({ id, ratio }: { id: string; ratio: number }) {
  const editor = usePptxEditor();
  const box = useRef<HTMLSpanElement>(null);
  // Repaint when the document changes (a theme or master edit changes every layout) and when
  // fonts load: the slide summaries change with the font generation, as the rail's do. The
  // engine skips a paint whose package, layout, and fonts are the same.
  const revision = useEditorState((s) => s.revision);
  const slides = useEditorState((s) => s.slides);
  useEffect(() => {
    if (editor && box.current) editor.renderLayoutThumbnail(id, box.current);
  }, [editor, id, revision, slides]);
  return (
    <span
      ref={box}
      className="pptx-layout-gallery__thumb"
      style={{ aspectRatio: String(1 / ratio) }}
      aria-hidden="true"
    />
  );
});

/** Props of {@link LayoutGallery}, named as DOCX Editor's `StylePicker` names them. @public */
export interface LayoutGalleryProps {
  /** The layout to mark as the current one (`menuitemradio`, checked). Absent: plain rows. */
  readonly value?: string;
  /** Called with the picked layout's id (`SlideLayoutSummary.id`). */
  readonly onChange: (layoutId: string) => void;
  /** A class name added to the gallery's root element. */
  readonly className?: string;
  /** The layouts to show. Default: every layout of the deck (`useSlideLayouts`). */
  readonly layouts?: readonly SlideLayoutSummary[];
  /** Show the tiles but take no pick. Default: true while the deck is open for viewing. */
  readonly disabled?: boolean;
  /**
   * Why the gallery takes no pick, for the tiles' tooltips: a catalog key (such as
   * `useEditorCommand(slotId).disabledReason`) or text. Default: the viewing reason.
   */
  readonly disabledReason?: string | null;
}

/**
 * A grid of the deck's layouts, as the Layout and New Slide galleries show them: one
 * group per slide master, headed by its name, and each tile the engine's render of a new slide
 * on the layout, named below. The tiles are menu rows (`menuitem`, or `menuitemradio` with
 * `value`); arrow keys, Home, and End move between them, so the gallery works in any menu or
 * popup of the host. @public
 */
export function LayoutGallery({
  value: current,
  onChange,
  className,
  layouts: given,
  disabled,
  disabledReason,
}: LayoutGalleryProps) {
  const deck = useSlideLayouts();
  const layouts = given ?? deck;
  const viewing = useEditorState((s) => s.mode === 'view');
  const off = disabled ?? viewing;
  const ratio = useEditorState((s) => s.slideHeight / s.slideWidth);
  const { t } = useTranslation();
  const entered = useRef(false);
  const reasonText = disabledReason
    ? translateKeyOrText(t, disabledReason)
    : viewing
      ? t('disabled.viewing')
      : null;
  // Every tile says the one reason through `aria-describedby`.
  const offTiles = useAriaDisabled(off, reasonText);
  if (!layouts.length)
    return <p className="pptx-layout-gallery__empty">{t('layoutGallery.empty')}</p>;
  const masters = layoutMasterGroups(layouts);
  // One tile takes Tab focus: the current layout, or the first.
  const focusable = layoutGalleryTabStop(layouts, current);
  return (
    <div
      className={['pptx-layout-gallery', className].filter(Boolean).join(' ')}
      onFocus={(e) => {
        // A menu opened from the keyboard focuses its first row: start on the current layout.
        if (entered.current) return;
        entered.current = true;
        layoutGalleryEnter(e.currentTarget, e.target);
      }}
      onKeyDown={(e) => {
        // Left from the first column leaves the gallery: a submenu then closes.
        if (layoutGalleryKey(e.currentTarget, e.target, e.key)) {
          e.preventDefault();
          e.stopPropagation();
        }
      }}
    >
      {masters.map((m) => (
        <div key={m.id} className="pptx-layout-gallery__group" role="group" aria-label={m.name}>
          <div className="pptx-layout-gallery__header" aria-hidden="true">
            {m.name}
          </div>
          <div className="pptx-layout-gallery__grid" data-columns={3}>
            {m.layouts.map((l) => (
              <button
                key={l.id}
                type="button"
                role={current === undefined ? 'menuitem' : 'menuitemradio'}
                {...(current === undefined ? {} : { 'aria-checked': l.id === current })}
                tabIndex={l.id === focusable ? 0 : -1}
                className="pptx-layout-gallery__item"
                data-layout-type={l.type}
                data-current={l.id === current ? '' : undefined}
                {...offTiles.attrs}
                {...tooltipAttributes(l.name, off, reasonText)}
                onClick={() => {
                  if (!off) onChange(l.id);
                }}
              >
                <LayoutThumb id={l.id} ratio={ratio} />
                <span className="pptx-layout-gallery__name">{l.name}</span>
              </button>
            ))}
          </div>
        </div>
      ))}
      {offTiles.note}
    </div>
  );
}

/** The props of `NewSlideControl`. @internal */
export interface NewSlideControlProps {
  /** The chrome slot the control runs. */
  slot: ChromeSlotId;
  /** The registry entry of the control, from `CHROME_GROUPS`. */
  control: ChromeControl;
}

/**
 * Home > New Slide: the button adds a slide on the default next layout; the arrow opens the
 * layout gallery. The toolbar's part for `slide.add`; hosts place it with `Toolbar.Slot`. @internal
 */
export function NewSlideControl({ slot, control }: NewSlideControlProps) {
  const { cmd, label, reason } = useSlot(slot, control);
  const caretOff = useAriaDisabled(!cmd.isEnabled, reason);
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const anchor = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  // The caret is part of the split button: host overrides go to the main button, as for the
  // color pickers, so the caret keeps its own class.
  const choose = t('layoutGallery.choose');
  return (
    <span className="pptx-toolbar__group">
      <ToolButton
        slot={slot}
        label={label}
        paths={control.paths}
        disabled={!cmd.isEnabled}
        disabledReason={reason}
        onClick={() => editor?.exec({ type: 'addSlide' })}
      />
      <button
        ref={anchor}
        type="button"
        className="pptx-toolbar__button pptx-toolbar__caret-button"
        data-disabled={cmd.isEnabled ? undefined : ''}
        aria-label={choose}
        aria-haspopup="menu"
        aria-expanded={open}
        {...tooltipAttributes(choose, !cmd.isEnabled, reason)}
        {...caretOff.attrs}
        onMouseDown={keepFocus}
        onClick={() => {
          if (cmd.isEnabled) setOpen((o) => !o);
        }}
      >
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>
      {caretOff.note}
      {open ? (
        <Popover anchor={anchor.current} onClose={() => setOpen(false)} label={label}>
          <LayoutGallery
            onChange={(id) => {
              setOpen(false);
              cmd.execute(id);
            }}
          />
        </Popover>
      ) : null}
    </span>
  );
}

/** The props of `LayoutControl`. @internal */
export interface LayoutControlProps {
  /** The chrome slot the control runs. */
  slot: ChromeSlotId;
  /** The registry entry of the control, from `CHROME_GROUPS`. */
  control: ChromeControl;
}

/**
 * Home > Layout: the layout gallery, with the current slide's layout marked. The trigger is one
 * button, so it takes the slot's `icon` and `asChild` as `ToolButton` does. The toolbar's part
 * for `slide.layout`; hosts place it with `Toolbar.Slot`. @internal
 */
export function LayoutControl({ slot, control }: LayoutControlProps) {
  const { cmd, label, tip, reason } = useSlot(slot, control);
  const override = useSlotOverride(slot);
  const [anchor, setAnchor] = useState<HTMLElement | null>(null);
  const [open, setOpen] = useState(false);
  const disabled = !cmd.isEnabled;
  // A disabled trigger stays focusable and says why (`aria-disabled`).
  const off = useAriaDisabled(disabled, reason);
  const toggle = (el: Element): void => {
    if (disabled || !(el instanceof HTMLElement)) return;
    setAnchor(el);
    setOpen((o) => !o);
  };
  const attributes = {
    className: ['pptx-toolbar__button pptx-toolbar__button--wide', override?.className]
      .filter(Boolean)
      .join(' '),
    'data-slot': slot,
    'data-disabled': disabled ? '' : undefined,
    'aria-label': label,
    'aria-haspopup': 'menu' as const,
    'aria-expanded': open,
    ...tip,
    ...off.attrs,
    onMouseDown: keepFocus,
    onClick: (e: { currentTarget: Element }) => toggle(e.currentTarget),
  };
  const host = override?.asChild;
  return (
    <>
      {host ? (
        // A host element that is not a button gets the button role, a tab stop, and Enter and
        // Space, as `ToolButton` gives it.
        renderAsChild(host, {
          ...attributes,
          ...(host.type === 'button' ? { type: 'button' } : buttonRoleProps(toggle)),
        })
      ) : (
        <button type="button" {...attributes}>
          {override?.icon !== undefined ? override.icon : <Icon paths={control.paths} />}
          <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
        </button>
      )}
      {off.note}
      {open ? (
        <Popover anchor={anchor} onClose={() => setOpen(false)}>
          <LayoutGallery
            value={typeof cmd.value === 'string' ? cmd.value : ''}
            onChange={(id) => {
              setOpen(false);
              cmd.execute(id);
            }}
          />
        </Popover>
      ) : null}
    </>
  );
}

/**
 * Home > Section > Add Section: adds a section at the current slide and starts renaming it, so it
 * asks for the name, where a rail or a sorter can show the name field
 * (`runSectionSlot`). The toolbar's part for `section.add`. @internal
 */
export function SectionAddControl({
  slot,
  control,
}: {
  slot: ChromeSlotId;
  control: ChromeControl;
}) {
  const { cmd, label, reason } = useSlot(slot, control);
  return (
    <ToolButton
      slot={slot}
      label={label}
      paths={control.paths}
      disabled={!cmd.isEnabled}
      disabledReason={reason}
      onClick={() => cmd.execute()}
    />
  );
}

/**
 * Home > Section > Rename Section: opens the name field of the current slide's section in the
 * rail or the sorter. Without one, it is off and says why (`disabled.noSectionPane`). The
 * toolbar's part for `section.rename`. @internal
 */
export function SectionRenameControl({
  slot,
  control,
}: {
  slot: ChromeSlotId;
  control: ChromeControl;
}) {
  const { cmd, label, reason } = useSlot(slot, control);
  return (
    <ToolButton
      slot={slot}
      label={label}
      paths={control.paths}
      disabled={!cmd.isEnabled}
      disabledReason={reason}
      onClick={() => cmd.execute()}
    />
  );
}
