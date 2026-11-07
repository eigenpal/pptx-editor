/** @jsxImportSource react */
// The edit notice: says why the last edit was refused (a keystroke, a paste, a command), so a
// refused edit never looks like an editor that stopped responding. After a Delete of shapes or
// slides it says what went and offers Undo, for ten seconds; a pointer over it or the focus in
// it holds it. It reads the engine's `lastRejection` and `lastRemoval` and holds no state of its
// own besides the dismissal and the hold. The messages and the rules for when they show are
// core's (`overlays.ts`), shared with the Vue `PptxEditorEditNotice`.

import {
  editNoticeShows,
  editNoticeTone,
  editRejectionMessage,
  editRemovalMessage,
  focusSurfaceOf,
  linkPopupShowsRejection,
  fieldShowsRejection,
  REMOVAL_NOTICE_MS,
  removalNoticeShows,
} from '@pptx-editor.dev/core/editor';
import { useContext, useEffect, useState, useSyncExternalStore, type ReactNode } from 'react';
import type { ChromeTranslate, EditorRejection } from '@pptx-editor.dev/core/editor';
import { withChromeTranslate } from './headless-hooks.tsx';
import type { Translate } from '@pptx-editor.dev/i18n';
import { ToolButton, keepFocus } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { usePopupEntry } from './popups.tsx';
import { HyperlinkPopupContext } from './useHyperlinkPopup.ts';

const NO_STORE = (): (() => void) => () => undefined;
const NO_STATE = (): null => null;

/** The message for a refusal: its reason, with the detail where the reason names one. @public */
export function rejectionMessage(t: Translate, rejection: EditorRejection): string {
  return editRejectionMessage(t, rejection);
}

/** The props of `EditNotice`. @public */
export interface EditNoticeProps {
  /**
   * Your translation function for the notice, as DOCX Editor's `t` on its notices. Without it,
   * the root's `locale` gives the labels.
   */
  readonly t?: ChromeTranslate;
  /** Appended after the `pptx-edit-notice` class. */
  className?: string;
}

/**
 * Why the last edit was refused, or what the last Delete removed with an Undo action. Renders
 * nothing while edits apply, and nothing for a command that had nothing to do
 * (`EditorRejection.noop`). @public
 */
export const EditNotice: (props: EditNoticeProps) => ReactNode =
  withChromeTranslate<EditNoticeProps>(EditNoticeRoot, 'EditNotice');

function EditNoticeRoot({ className }: Omit<EditNoticeProps, 't'>) {
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const rejection = useEditorState((s) => s.lastRejection ?? null);
  const removal = useEditorState((s) => s.lastRemoval ?? null);
  const [dismissed, setDismissed] = useState<EditorRejection | null>(null);
  // The serial of the removal the user closed or that timed out.
  const [closed, setClosed] = useState<number | null>(null);
  const [held, setHeld] = useState(false);
  // A click on Undo or Dismiss removes the notice under the pointer, and no mouseleave or blur
  // follows: each new removal starts unheld.
  useEffect(() => setHeld(false), [removal?.serial]);
  const offered = removalNoticeShows(removal, closed);
  useEffect(() => {
    if (!offered || held) return;
    const serial = removal.serial;
    const timer = setTimeout(() => setClosed(serial), REMOVAL_NOTICE_MS);
    return () => clearTimeout(timer);
  }, [offered, held, removal]);
  // A refused link shows its reason in the open link popup: the notice does not repeat it, and
  // stays closed when the popup closes.
  const linkStore = useContext(HyperlinkPopupContext);
  const linkPopup = useSyncExternalStore(
    linkStore?.subscribe ?? NO_STORE,
    linkStore?.getState ?? NO_STATE,
    linkStore?.getState ?? NO_STATE
  );
  // A field that shows its own refusal (a comment field) is not repeated either.
  const covered =
    (usePopupEntry('hyperlink') !== false && linkPopupShowsRejection(linkPopup, rejection)) ||
    fieldShowsRejection(rejection);
  useEffect(() => {
    if (covered) setDismissed(rejection);
  }, [covered, rejection]);
  const classes = `pptx-edit-notice${className ? ` ${className}` : ''}`;
  if (!covered && editNoticeShows(rejection, dismissed))
    return (
      <div
        className={classes}
        // A state the edit cannot run in (view mode) is information, not an error.
        data-tone={editNoticeTone(rejection)}
        role="status"
        onMouseDown={keepFocus}
      >
        <span className="pptx-edit-notice__text">{rejectionMessage(t, rejection)}</span>
        <ToolButton
          label={t('notice.dismiss')}
          icon="close"
          onClick={() => setDismissed(rejection)}
        />
      </div>
    );
  if (!offered) return null;
  return (
    <div
      className={classes}
      data-tone="info"
      role="status"
      onMouseDown={keepFocus}
      onMouseEnter={() => setHeld(true)}
      onMouseLeave={() => setHeld(false)}
      onFocus={() => setHeld(true)}
      onBlur={() => setHeld(false)}
    >
      <span className="pptx-edit-notice__text">{editRemovalMessage(t, removal)}</span>
      <button
        type="button"
        className="pptx-edit-notice__action"
        onClick={() => {
          if (!editor) return;
          editor.exec({ type: 'undo' });
          // The button goes with the notice; the focus goes back to the slide.
          focusSurfaceOf(editor);
        }}
      >
        {t('notice.undo')}
      </button>
      <ToolButton
        label={t('notice.dismiss')}
        icon="close"
        onClick={() => {
          setClosed(removal.serial);
          if (editor) focusSurfaceOf(editor);
        }}
      />
    </div>
  );
}
