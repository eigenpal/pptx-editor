/** @jsxImportSource react */
// The open-failure notice, after DOCX Editor's `DocxEditor.OpenFailure`: what a person sees when
// the editor refuses a file (a damaged package, a file that is not a presentation). Without it a
// refused open shows nothing, and the deck that was open looks like the new one. It reads
// `snapshot().openFailure`, the same pull state every other consumer reads, so it also covers a
// document passed at construction. The reason groups and their messages are core's
// (`loading.ts`), shared with the Vue `PptxEditorOpenFailure`.
//
// The default text is generic on purpose: it explains the refusal and shows the stable code,
// and names no contact. A host adds its own support action through `actions`.

import { useCallback, useId, useState, type CSSProperties, type ReactNode } from 'react';
import {
  OPEN_FAILURE_MESSAGE_KEYS,
  shownOpenFailure,
  type DocumentOpenFailure,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, useTranslation } from './context.ts';

/** The current refused open, with the packaged message and a way to dismiss it. @public */
export interface UseOpenFailureReturn {
  /** The refused open, or `null` when there is none or the person dismissed it. */
  readonly failure: DocumentOpenFailure | null;
  /** The packaged, translated message for `failure`. */
  readonly message: string | null;
  /** Whether the deck that was open before stays open. */
  readonly kept: boolean;
  /**
   * Hide this refusal. A later refusal shows again, even for the same file. Pass the failure
   * your UI showed, so a late call cannot hide a newer refusal.
   */
  dismiss(failure?: DocumentOpenFailure | null): void;
}

/**
 * Reads the file the editor refused to open, for a host that shows it in its own UI, such as
 * a toast. Pair it with `openFailure={false}` on `<PptxEditor>`, so the packaged notice does not
 * show as well. Call it inside a `PptxEditor.Root`. @public
 */
export function useOpenFailure(): UseOpenFailureReturn {
  const current = useEditorState(shownOpenFailure);
  const kept = useEditorState((s) => s.status === 'ready');
  const { t } = useTranslation();
  // Each refusal is a new object, so remembering the dismissed one by reference lets the next
  // refusal of the same file show again.
  const [dismissed, setDismissed] = useState<DocumentOpenFailure | null>(null);
  const failure = current !== null && current !== dismissed ? current : null;
  const dismiss = useCallback(
    (shown?: DocumentOpenFailure | null) => {
      if (!shown || shown === current) setDismissed(current);
    },
    [current]
  );
  return {
    failure,
    message: failure ? t(OPEN_FAILURE_MESSAGE_KEYS[failure.reason]) : null,
    kept,
    dismiss,
  };
}

/** What `actions` receives. @public */
export interface OpenFailureContext {
  /** The refused open: its stable `code` and its `reason`. */
  readonly failure: DocumentOpenFailure;
  /** The packaged, translated message for the failure. */
  readonly message: string;
  /** Close the notice. It calls `onDismiss`, as **Close** does. */
  dismiss(): void;
}

/** Props for `PptxEditor.OpenFailure`. @public */
export interface PptxEditorOpenFailureProps {
  /** The notice title. Defaults to the translated `openFailure.title`. */
  readonly title?: string;
  /** Your own message in place of the packaged one, as content or per failure. */
  readonly message?: ReactNode | ((failure: DocumentOpenFailure) => ReactNode);
  /** Show the stable error code under the message. Defaults to `true`. */
  readonly showCode?: boolean;
  /** Your own buttons, before **Close**. The notice shows no contact of its own. */
  readonly actions?: (context: OpenFailureContext) => ReactNode;
  /** The label of the **Close** button. Defaults to the translated `openFailure.close`. */
  readonly closeLabel?: string;
  /** Show the packaged **Close** button. Defaults to `true`. */
  readonly showClose?: boolean;
  /** Called after the person closes the notice, by any route. */
  readonly onDismiss?: (failure: DocumentOpenFailure) => void;
  /** Appended after the `pptx-open-failure` class. */
  readonly className?: string;
  /** Inline styles for the notice. */
  readonly style?: CSSProperties;
}

/**
 * A notice that explains why a file did not open, and says when the deck that was open stays
 * open. It shows nothing until the editor refuses a file, and shows again for each later
 * refusal. `<PptxEditor>` renders it by default; compose it yourself inside a positioned box in
 * a `PptxEditor.Root`:
 *
 * ```tsx
 * <PptxEditor.OpenFailure
 *   actions={({ failure }) => (
 *     <a className="pptx-open-failure__action" href={`/support?code=${failure.code}`}>
 *       Get help
 *     </a>
 *   )}
 * />
 * ```
 *
 * For a toast or a banner of your own, use `useOpenFailure`. @public
 */
export function OpenFailure({
  title,
  message,
  showCode = true,
  actions,
  closeLabel,
  showClose = true,
  onDismiss,
  className,
  style,
}: PptxEditorOpenFailureProps) {
  const { t } = useTranslation();
  const { failure, message: packaged, kept, dismiss } = useOpenFailure();
  const titleId = useId();
  if (!failure || packaged === null) return null;
  const close = (): void => {
    dismiss(failure);
    onDismiss?.(failure);
  };
  const context: OpenFailureContext = { failure, message: packaged, dismiss: close };
  const body = typeof message === 'function' ? message(failure) : (message ?? packaged);
  return (
    <div
      className={`pptx-open-failure${className ? ` ${className}` : ''}`}
      style={style}
      role="alert"
      aria-labelledby={titleId}
      data-reason={failure.reason}
    >
      <div className="pptx-open-failure__text">
        <p id={titleId} className="pptx-open-failure__title">
          {title ?? t('openFailure.title')}
        </p>
        <div className="pptx-open-failure__message">{body}</div>
        {kept ? <p className="pptx-open-failure__kept">{t('openFailure.kept')}</p> : null}
        {showCode ? (
          <p className="pptx-open-failure__code">{t('openFailure.code', { code: failure.code })}</p>
        ) : null}
      </div>
      <div className="pptx-open-failure__actions">
        {actions?.(context)}
        {showClose ? (
          <button type="button" className="pptx-open-failure__action" onClick={close}>
            {closeLabel ?? t('openFailure.close')}
          </button>
        ) : null}
      </div>
    </div>
  );
}
