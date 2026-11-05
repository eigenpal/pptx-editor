/** @jsxImportSource react */
// The font substitution notice, matching DOCX Editor's: it lists the families that render with
// a substitute that can change line breaks, and offers the fonts installed on this computer
// where the browser supports it. Metric-compatible substitutes (Calibri → Carlito, …) are not
// listed: their line breaks match the original font. The list and the local font request are
// core's (`overlays.ts`), shared with the Vue `PptxEditorFontNotice`.

import { useEffect, useRef, useState, type ReactNode } from 'react';
import {
  applyLocalFonts,
  canQueryLocalFonts,
  fontNoticeFamilies,
  type ChromeTranslate,
} from '@pptx-editor.dev/core/editor';
import { withChromeTranslate } from './headless-hooks.tsx';
import { ToolButton, keepFocus } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';

/** The props of `FontNotice`. @public */
export interface FontNoticeProps {
  /**
   * Your translation function for the notice, as DOCX Editor's `t` on its font notice. Without
   * it, the root's `locale` gives the labels.
   */
  readonly t?: ChromeTranslate;
  /** Appended after the `pptx-font-notice` class. */
  className?: string;
}

/** The substitution notice. Renders nothing when every font is available. @public */
export const FontNotice: (props: FontNoticeProps) => ReactNode =
  withChromeTranslate<FontNoticeProps>(FontNoticeRoot, 'FontNotice');

function FontNoticeRoot({ className }: Omit<FontNoticeProps, 't'>) {
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const substituted = useEditorState((s) => s.substitutedFonts);
  // A newly opened document starts with no outcome of an earlier local font request.
  const loadState = useEditorState((s) => s.status);
  const missing = fontNoticeFamilies(substituted);
  const key = missing.join('|');
  const [dismissed, setDismissed] = useState<string | null>(null);
  // The outcome of a local font request, for the substituted families when it finished.
  const [outcome, setOutcome] = useState<{ text: string; key: string } | null>(null);
  const [busy, setBusy] = useState(false);
  // Each opened document starts a new round; a request from an earlier round reports nothing.
  // A request that adds fonts changes `key` itself, so only a load starts a round.
  const round = useRef(0);
  useEffect(() => {
    round.current += 1;
    setOutcome(null);
  }, [loadState]);
  // A later change to the substituted families makes the outcome stale.
  const status = outcome?.key === key ? outcome.text : null;
  if ((missing.length === 0 && !status) || dismissed === key) return null;
  const addLocalFonts = async (): Promise<void> => {
    if (!editor) return;
    setBusy(true);
    const asked = round.current;
    try {
      const result = await applyLocalFonts(editor, missing);
      const now = fontNoticeFamilies(editor.snapshot().substitutedFonts).join('|');
      if (asked === round.current) setOutcome({ text: t(result), key: now });
    } catch {
      // `setFonts` rejects only after the editor is destroyed; nothing is left to tell.
    } finally {
      setBusy(false);
    }
  };
  return (
    <div
      className={`pptx-font-notice${className ? ` ${className}` : ''}`}
      role="status"
      onMouseDown={keepFocus}
    >
      <span className="pptx-font-notice__text">
        {missing.length ? t('fonts.notice', { families: missing.join(', ') }) : status}
        {missing.length && status ? ` ${status}` : ''}
      </span>
      {missing.length && canQueryLocalFonts() ? (
        <button
          type="button"
          className="pptx-font-notice__action"
          disabled={busy || loadState === 'loading'}
          onClick={() => void addLocalFonts()}
        >
          {t('fonts.useLocal')}
        </button>
      ) : null}
      <ToolButton label={t('fonts.dismiss')} icon="close" onClick={() => setDismissed(key)} />
    </div>
  );
}
