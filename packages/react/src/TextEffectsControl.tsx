/** @jsxImportSource react */
// Text Effects (`shape.textEffects`): a Transform submenu with No Transform and the text warp
// gallery, one section for the warps that follow a path and one for the warps between paths.
// Each cell strokes the warp's own paths (`textWarpPreview`). The registry says what a pick
// dispatches and when the control is enabled (`chrome-text-effects.ts`).

import { useCallback, useRef, useState, type ReactNode } from 'react';
import {
  TEXT_WARP_GALLERY,
  chromeOptionChecked,
  textWarpPreview,
  type ChromeControl,
  type ChromeSlotId,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { Icon, Popover, keepFocus, useAriaDisabled } from './chrome.tsx';
import { MenuRow, MenuSubmenu } from './ContextMenu.tsx';
import { useTranslation } from './context.ts';
import { useSlot } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';

/** The Text Effects trigger: a letter over an arched baseline, drawn, not an icon path. */
function TextEffectsGlyph(): ReactNode {
  return (
    <svg className="pptx-icon pptx-text-effects-glyph" viewBox="0 0 24 24" aria-hidden="true">
      <path d="M7 15 11 4h2l4 11h-2.2l-.9-2.6H10L9.2 15Zm3.6-4.4h2.8L12 6.5Z" />
      <path d="M3 21q9-6 18 0" className="pptx-text-effects-glyph__arc" />
    </svg>
  );
}

/** One gallery cell: the warp's paths, stroked. */
function WarpSwatch({ preset }: { preset: string }): ReactNode {
  return (
    <svg className="pptx-style-gallery__swatch" viewBox="0 0 40 28" aria-hidden="true">
      {textWarpPreview(preset).map((d, i) => (
        <path key={i} d={d} className="pptx-warp-gallery__path" />
      ))}
    </svg>
  );
}

/**
 * The Transform gallery: a heading and a grid of cells for each section. Up and Down move by a
 * row of cells (`data-columns`).
 */
function WarpGallery({
  slot,
  value,
  onPick,
}: {
  slot: ChromeSlotId;
  value: unknown;
  onPick: (value: string) => void;
}): ReactNode {
  const { t } = useTranslation();
  const columns = 5;
  return (
    <>
      {TEXT_WARP_GALLERY.map(({ section, labelKey, warps }) => (
        <div
          key={section}
          className="pptx-style-gallery pptx-warp-gallery"
          role="group"
          aria-label={t(labelKey)}
          data-columns={columns}
        >
          <div className="pptx-menu__heading" aria-hidden="true">
            {t(labelKey)}
          </div>
          {Array.from({ length: Math.ceil(warps.length / columns) }, (_, r) => (
            <div key={r} className="pptx-style-gallery__row" role="presentation">
              {warps.slice(r * columns, (r + 1) * columns).map((w) => {
                const name = t(`toolbar.textWarps.${w}` as TranslationKey);
                const option = `transform.${w}`;
                return (
                  <button
                    key={w}
                    type="button"
                    role="menuitemradio"
                    aria-checked={chromeOptionChecked(value, option)}
                    tabIndex={-1}
                    className="pptx-menu__item pptx-style-gallery__cell"
                    aria-label={name}
                    data-pptx-tooltip={name}
                    data-slot={slot}
                    data-value={option}
                    onClick={() => onPick(option)}
                  >
                    <WarpSwatch preset={w} />
                  </button>
                );
              })}
            </div>
          ))}
        </div>
      ))}
    </>
  );
}

/**
 * The Text Effects menu (`shape.textEffects`): the Transform submenu, with the selection's warp
 * checked. @internal
 */
export function TextEffectsMenuControl({
  slot,
  control,
}: {
  slot: ChromeSlotId;
  control: ChromeControl;
}): ReactNode {
  const { cmd, label, reason, tip } = useSlot(slot, control);
  // A disabled picker stays in the arrow order and says why.
  const off = useAriaDisabled(!cmd.isEnabled, reason);
  const { t } = useTranslation();
  const anchor = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  const trigger = useTriggerAttributes(slot, 'pptx-toolbar__button pptx-toolbar__button--wide', {
    disabled: !cmd.isEnabled,
  });
  const close = useCallback(() => setOpen(false), []);
  const pick = (value: string): void => {
    cmd.execute(value);
    close();
  };
  return (
    <>
      <button
        ref={anchor}
        type="button"
        {...trigger}
        data-slot={slot}
        aria-label={label}
        aria-haspopup="menu"
        aria-expanded={open}
        {...tip}
        {...off.attrs}
        onMouseDown={keepFocus}
        onClick={() => setOpen((o) => !o)}
      >
        <TextEffectsGlyph />
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>
      {off.note}
      {open ? (
        <Popover anchor={anchor.current} onClose={close}>
          {control.optionGroups?.map((group) => (
            <MenuSubmenu key={group.id} label={t(group.labelKey)}>
              <MenuRow
                slot={slot}
                label={t('toolbar.noTransform')}
                checked={chromeOptionChecked(cmd.value, 'transform.none')}
                radio
                onSelect={() => pick('transform.none')}
              />
              <WarpGallery slot={slot} value={cmd.value} onPick={pick} />
            </MenuSubmenu>
          ))}
        </Popover>
      ) : null}
    </>
  );
}
