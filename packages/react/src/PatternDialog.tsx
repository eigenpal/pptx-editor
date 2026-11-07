/** @jsxImportSource react */
// The pattern fill dialog: the 54 preset patterns, the foreground and background colors, and a
// preview. The shape fill menu and the Background dialog open it. The gallery order, names, and
// swatches come from core (`pattern-fields.ts`), so the Vue twin shows the same dialog.

import { useEffect, useId, useRef, useState } from 'react';
import {
  dialogInitialFocus,
  isDialogBackdropPoint,
  PATTERN_CHOICES,
  patternCss,
  patternDialogStart,
  patternLabel,
  type PatternFill,
} from '@pptx-editor.dev/core/editor';
import { ToolButton } from './chrome.tsx';
import { useTranslation } from './context.ts';
import { Swatch, SwatchGroup } from './ColorSwatches.tsx';

/** Props for `PptxEditor.PatternDialog`. @public */
export interface PatternDialogProps {
  /** Whether the dialog shows. */
  readonly open: boolean;
  /** The pattern the dialog starts from, or null for a 50% pattern in `base` on white. */
  readonly initial: PatternFill | null;
  /** The foreground of a new pattern, as `#RRGGBB`. */
  readonly base?: string;
  /** Called when the dialog closes, applied or not. */
  readonly onClose: () => void;
  /** Called with the pattern the person applied. */
  readonly onApply: (pattern: PatternFill) => void;
}

/** The pattern fill dialog: a preset, its two colors, and a preview. @public */
export function PatternDialog({ open, initial, base, onClose, onApply }: PatternDialogProps) {
  const { t } = useTranslation();
  const ref = useRef<HTMLDialogElement>(null);
  const titleId = useId();
  const pressedOutside = useRef(false);
  const [p, setP] = useState<PatternFill>(() => patternDialogStart(initial, base));

  useEffect(() => {
    const dialog = ref.current;
    if (!dialog) return;
    if (open) {
      setP(patternDialogStart(initial, base));
      if (!dialog.open) dialog.showModal();
      // The chosen pattern takes focus, as the first field of the dialog. It is found by name:
      // the pressed state of this render can still show the previous opening.
      const name = CSS.escape(patternLabel(t, patternDialogStart(initial, base).preset));
      const chosen = dialog.querySelector<HTMLElement>(`.pptx-swatch[aria-label="${name}"]`);
      (chosen ?? dialogInitialFocus(dialog))?.focus({ preventScroll: true });
    } else if (dialog.open) dialog.close();
  }, [open, initial, base, t]);

  return (
    <dialog
      ref={ref}
      className="pptx-dialog pptx-pattern"
      data-pptx-dialog="pattern"
      aria-labelledby={titleId}
      onCancel={(e) => {
        e.preventDefault();
        onClose();
      }}
      onMouseDown={(e) => {
        pressedOutside.current = isDialogBackdropPoint(e.currentTarget, e);
      }}
      onClick={(e) => {
        // A press and a release on the backdrop close it; a drag out of a field does not.
        if (pressedOutside.current && isDialogBackdropPoint(e.currentTarget, e)) onClose();
        pressedOutside.current = false;
      }}
    >
      {open ? (
        <>
          <div className="pptx-dialog__header">
            <h2 id={titleId} className="pptx-dialog__title">
              {t('pattern.title')}
            </h2>
            <ToolButton label={t('pattern.close')} icon="close" onClick={onClose} />
          </div>
          <div className="pptx-dialog__body">
            <SwatchGroup label={t('pattern.patterns')} columns={9} roving>
              {PATTERN_CHOICES.map((preset) => (
                <Swatch
                  key={preset}
                  color={p.foreground}
                  background={patternCss({ ...p, preset })}
                  name={patternLabel(t, preset)}
                  menu={false}
                  pressed={p.preset === preset}
                  onPick={() => setP({ ...p, preset })}
                />
              ))}
            </SwatchGroup>
            <div className="pptx-pattern__row">
              <label className="pptx-field">
                {t('pattern.foreground')}
                <input
                  type="color"
                  className="pptx-gradient__color"
                  value={p.foreground.toLowerCase()}
                  onChange={(e) => setP({ ...p, foreground: e.target.value.toUpperCase() })}
                />
              </label>
              <label className="pptx-field">
                {t('pattern.background')}
                <input
                  type="color"
                  className="pptx-gradient__color"
                  value={p.background.toLowerCase()}
                  onChange={(e) => setP({ ...p, background: e.target.value.toUpperCase() })}
                />
              </label>
              <div className="pptx-dialog__section">
                <h3 className="pptx-dialog__label">{t('pattern.preview')}</h3>
                <div
                  className="pptx-pattern__preview"
                  role="img"
                  aria-label={patternLabel(t, p.preset)}
                  style={{ background: patternCss(p, 12) }}
                />
              </div>
            </div>
          </div>
          <div className="pptx-dialog__footer">
            <button type="button" className="pptx-button" onClick={onClose}>
              {t('dialog.cancel')}
            </button>
            <button
              type="button"
              className="pptx-button pptx-button--primary"
              onClick={() => onApply(p)}
            >
              {t('dialog.ok')}
            </button>
          </div>
        </>
      ) : null}
    </dialog>
  );
}
