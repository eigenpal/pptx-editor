/** @jsxImportSource react */
// The toolbar controls that open a popup or a file picker: the link control and the picture
// file controls (Insert Picture, Change Picture). Split from `Toolbar.tsx` for the line cap.

import { useRef } from 'react';
import { imageFileTypes, insertImageFile, translateKeyOrText } from '@pptx-editor.dev/core/editor';
import { ToolButton } from './chrome.tsx';
import { usePptxEditor, useTranslation } from './context.ts';
import { useSlot, type SlotPartProps } from './slot.ts';
import { requestLinkPopup, useChromeActionHost } from './chrome-actions.tsx';

/** The link control: it opens the link popup (`popups.hyperlink`). @internal */
export function LinkControl({ slot, control }: SlotPartProps) {
  const { cmd, label, reason } = useSlot(slot, control);
  const editor = usePptxEditor();
  const { t } = useTranslation();
  // The button opens the popup, so a `popups.hyperlink` of `false` turns it off; the command
  // (`setLink`) keeps working.
  const off = useChromeActionHost().unavailable?.(slot) ?? null;
  const anchor = useRef<HTMLSpanElement>(null);
  return (
    <span ref={anchor} className="pptx-toolbar__group">
      <ToolButton
        slot={slot}
        label={label}
        paths={control.paths}
        disabled={!!off || !cmd.isEnabled}
        disabledReason={off ? translateKeyOrText(t, off) : reason}
        pressed={typeof cmd.value === 'string' && cmd.value !== ''}
        onClick={() => requestLinkPopup(editor, anchor.current)}
      />
    </span>
  );
}

/** Insert Picture and Change Picture: a button over a hidden file input. @internal */
export function ImageControl({ slot, control }: SlotPartProps) {
  const { cmd, label, reason } = useSlot(slot, control);
  const editor = usePptxEditor();
  const input = useRef<HTMLInputElement>(null);
  // Insert > Picture adds a picture; Change Picture replaces the selected one's image.
  const target = control.state.kind === 'replace' ? 'replaceImage' : undefined;
  const onFile = async (file: File | undefined): Promise<void> => {
    if (file && editor) await insertImageFile(editor, file, target);
  };
  return (
    <>
      <ToolButton
        slot={slot}
        label={label}
        paths={control.paths}
        disabled={!cmd.isEnabled}
        disabledReason={reason}
        onClick={() => input.current?.click()}
      />
      <input
        ref={input}
        type="file"
        accept={imageFileTypes(target)}
        hidden
        onChange={(e) => {
          const file = e.target.files?.[0];
          // Clear the input, so picking the same file again still fires `change`.
          e.target.value = '';
          void onFile(file);
        }}
      />
    </>
  );
}
