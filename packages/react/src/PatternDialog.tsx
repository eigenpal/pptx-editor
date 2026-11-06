/** @jsxImportSource react */
// The pattern fill dialog: the 54 preset patterns, the foreground and background colors, and a
// preview. The shape fill menu and the Background dialog open it. The gallery order, names, and
// swatches come from core (`pattern-fields.ts`), so the Vue twin shows the same dialog.

import { useEffect, useRef, useState } from 'react';
import {
  dialogInitialFocus,
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
  /** The pattern the dialog starts from, or null for 50% in the theme's first accent. */
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
  const [p, setP] = useState<PatternFill>(() => patternDialogStart(initial, base));

  useEffect(() => {
    const dialog = ref.current;
    if (!dialog) return;
    if (open) {
      setP(patternDialogStart(initial, base));
      if (!dialog.open) dialog.showModal();
      // The chosen pattern takes focus, as the first field of the dialog.
      const chosen = dialog.querySelector<HTMLElement>('.pptx-swatch[aria-pressed="true"]');
      (chosen ?? dialogInitialFocus(dialog))?.focus({ preventScroll: true });
    } else if (dialog.open) dialog.close();
  }, [open, initial, base]);

  return (
    <dialog
      ref={ref}
      className="pptx-dialog pptx-pattern"
      data-pptx-dialog="pattern"
      aria-labelledby="pptx-pattern-title"
      onCancel={(e) => {
        e.preventDefault();
        onClose();
      }}
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      {open ? (
        <>
          <div className="pptx-dialog__header">
            <h2 id="pptx-pattern-title" className="pptx-dialog__title">
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
