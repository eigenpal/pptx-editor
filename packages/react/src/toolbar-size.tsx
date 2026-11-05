/** @jsxImportSource react */
// The toolbar's Font size control: Decrease, the size box, and Increase. Split from
// `Toolbar.tsx`; the stepper shape of `CHROME_GROUPS` renders it.

import { useEffect, useId, useRef, useState } from 'react';
import { ToolButton, focusText, useAriaDisabled } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { useSlot, type SlotPartProps } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';
import {
  SIZE_ERROR_MS,
  parseToolbarSize,
  sizeErrorPlace,
  toolbarSizeRefusal,
  toolbarSizeText,
} from '@pptx-editor.dev/core/editor';

/** The `font.size` stepper. @internal */
export function SizeControl({ slot, control }: SlotPartProps) {
  const { cmd, label, reason, tip } = useSlot(slot, control);
  const editor = usePptxEditor();
  const { t } = useTranslation();
  // The steps are off over a mixed size, as DOCX Editor turns its steppers off.
  const mixed = useEditorState((s) => s.format?.font?.disagrees.fontSizePt === true);
  const [draft, setDraft] = useState<string | null>(null);
  const box = useRef<HTMLInputElement>(null);
  // Set when Enter or Escape has already handled the box, so the blur that follows does nothing.
  const handled = useRef(false);
  // Over a mixed size the box is blank, as the Font dialog's size field is.
  const value = toolbarSizeText(cmd.value, mixed);
  const off = useAriaDisabled(!cmd.isEnabled, reason);
  // A size out of range is refused with its range, under the box, and announced.
  const errorId = useId();
  const [error, setError] = useState<{ text: string; left: number; top: number } | null>(null);
  useEffect(() => {
    if (!error) return;
    const timer = setTimeout(() => setError(null), SIZE_ERROR_MS);
    return () => clearTimeout(timer);
  }, [error]);
  const commit = (): void => {
    const refusal = toolbarSizeRefusal(draft);
    const n = parseToolbarSize(draft);
    if (!handled.current && refusal)
      setError({ text: t(refusal.key, refusal.params), ...sizeErrorPlace(box.current) });
    else if (!handled.current && n !== null) cmd.execute(n);
    handled.current = false;
    setDraft(null);
  };
  // Enter and Escape hand the caret back to the text, as DOCX Editor's size box does, so a size
  // armed at a caret applies to the text typed next.
  const leave = (apply: boolean): void => {
    if (apply) commit();
    setDraft(null);
    handled.current = true;
    focusText(editor, box.current);
    if (box.current?.ownerDocument.activeElement === box.current) handled.current = false;
  };
  return (
    <span
      {...useTriggerAttributes(slot, 'pptx-toolbar__stepper', { disabled: !cmd.isEnabled })}
      data-slot={slot}
    >
      <ToolButton
        label={t('toolbar.fontSizeDecrease')}
        icon="remove"
        className="pptx-toolbar__step"
        disabled={!cmd.isEnabled || mixed}
        disabledReason={mixed ? t('disabled.mixedSize') : reason}
        onClick={() => cmd.execute({ step: -1 })}
      />
      <input
        ref={box}
        className="pptx-toolbar__size"
        aria-label={label}
        {...tip}
        value={draft ?? value}
        {...(error ? { 'aria-invalid': true, 'aria-describedby': errorId } : {})}
        {...off.attrs}
        readOnly={!cmd.isEnabled}
        onChange={(e) => {
          setDraft(e.target.value);
          setError(null);
        }}
        onBlur={commit}
        onKeyDown={(e) => {
          // The key stops here: Enter must not reach the slide text, and Escape must not close a
          // More card around the box (focus goes to the text instead).
          if (e.key !== 'Enter' && e.key !== 'Escape') return;
          e.preventDefault();
          e.stopPropagation();
          leave(e.key === 'Enter');
        }}
      />
      <ToolButton
        label={t('toolbar.fontSizeIncrease')}
        icon="add"
        className="pptx-toolbar__step"
        disabled={!cmd.isEnabled || mixed}
        disabledReason={mixed ? t('disabled.mixedSize') : reason}
        onClick={() => cmd.execute({ step: 1 })}
      />
      {off.note}
      {error ? (
        <span
          id={errorId}
          role="alert"
          className="pptx-toolbar__size-error"
          style={{ left: error.left, top: error.top }}
        >
          {error.text}
        </span>
      ) : null}
    </span>
  );
}
