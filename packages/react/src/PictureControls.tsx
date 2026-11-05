/** @jsxImportSource react */
// The Picture Format controls whose look is more than a list: the Corrections gallery, a grid
// with a row per contrast and a column per brightness, each cell a sample picture with that
// correction. A slot renderer: the registry says what it dispatches and when it is enabled
// (`chrome-picture.ts`); core draws the grid and the samples (`picture-corrections.ts`).

import { useRef, useState, type ReactNode } from 'react';
import {
  chromeOptionChecked,
  correctionGridRows,
  correctionPercent,
  correctionSwatch,
  type ChromeControl,
  type ChromeSlotId,
} from '@pptx-editor.dev/core/editor';
import { Icon, Popover, keepFocus, useAriaDisabled } from './chrome.tsx';
import { useTranslation } from './context.ts';
import { useSlot } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';

/** One Corrections sample: the sample picture with a brightness and a contrast. */
function CorrectionSwatch({
  brightness,
  contrast,
}: {
  brightness: number;
  contrast: number;
}): ReactNode {
  return (
    <svg className="pptx-corrections-grid__swatch" viewBox="0 0 40 28" aria-hidden="true">
      {correctionSwatch(brightness, contrast).map((p) => (
        <path key={p.d} d={p.d} fill={p.fill} />
      ))}
    </svg>
  );
}

/** The toolbar control of `image.corrections`: a trigger and the grid in a popover. @internal */
export function CorrectionsControl({
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
  const rows = correctionGridRows(control.options ?? []);
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
        <Icon paths={control.paths} />
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>
      {off.note}
      {open ? (
        <Popover anchor={anchor.current} onClose={() => setOpen(false)}>
          <div
            className="pptx-style-gallery pptx-corrections-grid"
            role="group"
            aria-label={t('pictureFormat.brightnessContrast')}
            // A grid for the arrow keys: Up and Down move by a row of brightness values.
            data-columns={rows[0]?.cells.length}
          >
            {/* The axis labels: each cell names its values, so they are hidden from assistive technology. */}
            <div className="pptx-corrections-grid__title" aria-hidden="true">
              {t('pictureFormat.brightness')}
            </div>
            <div className="pptx-corrections-grid__head" aria-hidden="true">
              {rows[0]?.cells.map(({ brightness }) => (
                <span key={brightness} className="pptx-corrections-grid__label">
                  {t('pictureFormat.percent', { value: correctionPercent(brightness) })}
                </span>
              ))}
            </div>
            <div className="pptx-corrections-grid__body" role="presentation">
              <div className="pptx-corrections-grid__side" aria-hidden="true">
                {t('pictureFormat.contrast')}
              </div>
              <div className="pptx-corrections-grid__rows" role="presentation">
                {rows.map(({ contrast, cells }) => (
                  <div key={contrast} className="pptx-style-gallery__row" role="presentation">
                    <span
                      className="pptx-corrections-grid__label pptx-corrections-grid__label--row"
                      aria-hidden="true"
                    >
                      {t('pictureFormat.percent', { value: correctionPercent(contrast) })}
                    </span>
                    {cells.map(({ option: o, brightness, contrast: c }) => {
                      const name = t(o.labelKey, o.labelParams);
                      return (
                        <button
                          key={o.value}
                          type="button"
                          role="menuitemradio"
                          aria-checked={chromeOptionChecked(cmd.value, o.value)}
                          tabIndex={-1}
                          className="pptx-menu__item pptx-style-gallery__cell"
                          aria-label={name}
                          data-pptx-tooltip={name}
                          data-slot={slot}
                          data-value={o.value}
                          onClick={() => {
                            cmd.execute(o.value);
                            setOpen(false);
                          }}
                        >
                          <CorrectionSwatch brightness={brightness} contrast={c} />
                        </button>
                      );
                    })}
                  </div>
                ))}
              </div>
            </div>
          </div>
        </Popover>
      ) : null}
    </>
  );
}
