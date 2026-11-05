// The demo's "Discard unsaved changes?" question: asked before New, Open, or a drop replaces a
// deck with edits that were not exported. A native modal with the editor's dialog classes
// (`pptx-dialog`, its header, body, footer, and buttons), so it looks like every editor dialog.
// Escape and Keep editing cancel; the safe choice has the focus first.

import { useLayoutEffect, useRef } from 'react';
import { useTranslation } from '@pptx-editor.dev/react';

export function DiscardDialog({
  onDiscard,
  onCancel,
}: {
  onDiscard: () => void;
  onCancel: () => void;
}) {
  const { t } = useTranslation();
  const ref = useRef<HTMLDialogElement>(null);
  const keep = useRef<HTMLButtonElement>(null);
  useLayoutEffect(() => {
    const dialog = ref.current;
    if (!dialog) return;
    const opener = dialog.ownerDocument.activeElement as HTMLElement | null;
    if (typeof dialog.showModal === 'function' && !dialog.open) dialog.showModal();
    else dialog.setAttribute('open', '');
    keep.current?.focus();
    return () => {
      if (dialog.open && typeof dialog.close === 'function') dialog.close();
      if (opener?.isConnected) opener.focus({ preventScroll: true });
    };
  }, []);
  return (
    <dialog
      ref={ref}
      className="pptx-dialog"
      data-pptx-dialog="discardChanges"
      aria-labelledby="demo-discard-title"
      aria-describedby="demo-discard-message"
      onCancel={(e) => {
        e.preventDefault();
        onCancel();
      }}
    >
      <div className="pptx-dialog__header">
        <h2 id="demo-discard-title" className="pptx-dialog__title">
          {t('app.discardTitle')}
        </h2>
      </div>
      <div className="pptx-dialog__body">
        <p id="demo-discard-message">{t('app.discardMessage')}</p>
      </div>
      <div className="pptx-dialog__footer">
        <button ref={keep} type="button" className="pptx-button" onClick={onCancel}>
          {t('app.keepEditing')}
        </button>
        <button type="button" className="pptx-button pptx-button--destructive" onClick={onDiscard}>
          {t('app.discard')}
        </button>
      </div>
    </dialog>
  );
}
