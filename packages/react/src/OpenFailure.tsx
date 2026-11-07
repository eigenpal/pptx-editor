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

import {
  useCallback,
  useEffect,
  useId,
  useRef,
  useState,
  type CSSProperties,
  type ReactNode,
} from 'react';
import {
  copyOpenFailureReport,
  OPEN_FAILURE_MESSAGE_KEYS,
  COPY_RESULT_GAP_MS,
  openFailureKey,
  type OpenFailureCopy,
  shownOpenFailure,
  type ChromeTranslate,
  type DocumentOpenFailure,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { withChromeTranslate } from './headless-hooks.tsx';

/** Options for `useOpenFailure`, as DOCX Editor's `UseOpenFailureOptions`. @public */
export interface UseOpenFailureOptions {
  /** Your translation function for `message`. Without it, the root's `locale` gives it. */
  readonly t?: ChromeTranslate;
}

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
export function useOpenFailure(options: UseOpenFailureOptions = {}): UseOpenFailureReturn {
  const current = useEditorState(shownOpenFailure);
  const kept = useEditorState((s) => s.status === 'ready');
  const { t: ambient } = useTranslation();
  const t = options.t ?? ambient;
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
  /**
   * Show **Copy details**, which copies `openFailureReport(failure)`: the code, the reason,
   * the editor version, the browser, and the time, for a support request. It holds no file
   * content. Defaults to `true`.
   */
  readonly showCopy?: boolean;
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
  /**
   * Your translation function for the notice, as DOCX Editor's `t`. Without it, the root's
   * `locale` gives the labels.
   */
  readonly t?: ChromeTranslate;
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
export const OpenFailure: (props: PptxEditorOpenFailureProps) => ReactNode =
  withChromeTranslate<PptxEditorOpenFailureProps>(OpenFailureRoot, 'OpenFailure');

function OpenFailureRoot({
  title,
  message,
  showCode = true,
  showCopy = true,
  actions,
  closeLabel,
  showClose = true,
  onDismiss,
  className,
  style,
}: Omit<PptxEditorOpenFailureProps, 't'>) {
  const { t } = useTranslation();
  const { failure, message: packaged, kept, dismiss } = useOpenFailure();
  const titleId = useId();
  // The last copy and its result, for the refusal it copied: a later refusal starts fresh.
  const [copy, setCopy] = useState<OpenFailureCopy | null>(null);
  const reportRef = useRef<HTMLTextAreaElement>(null);
  const rootRef = useRef<HTMLDivElement>(null);
  const editor = usePptxEditor();
  // A blocked copy moves focus to the report, which selects it for the copy key.
  useEffect(() => {
    if (copy && !copy.copied) reportRef.current?.focus();
  }, [copy]);
  if (!failure || packaged === null) return null;
  const copyReport = async (): Promise<void> => {
    // Clear the result first, so the same result a second time is read again.
    setCopy(null);
    await new Promise((resolve) => setTimeout(resolve, COPY_RESULT_GAP_MS));
    setCopy(await copyOpenFailureReport(failure));
  };
  const copyShown = copy?.failure === failure ? copy : null;
  const close = (): void => {
    // Focus inside the notice goes back to the editor, not to the page body.
    const focused = rootRef.current?.contains(document.activeElement) ?? false;
    dismiss(failure);
    onDismiss?.(failure);
    if (focused) editor?.focus();
  };
  const context: OpenFailureContext = { failure, message: packaged, dismiss: close };
  const body = typeof message === 'function' ? message(failure) : (message ?? packaged);
  return (
    <div
      key={openFailureKey(failure)}
      ref={rootRef}
      className={`pptx-open-failure${className ? ` ${className}` : ''}`}
      style={style}
      role="group"
      aria-labelledby={titleId}
      data-reason={failure.reason}
    >
      {/* Only the reason is an alert: the copy result and the report read on their own. */}
      <div className="pptx-open-failure__text" role="alert">
        <p id={titleId} className="pptx-open-failure__title">
          {title ?? t('openFailure.title')}
        </p>
        <div className="pptx-open-failure__message">{body}</div>
        {kept ? <p className="pptx-open-failure__kept">{t('openFailure.kept')}</p> : null}
        {showCode ? (
          <p className="pptx-open-failure__code">{t('openFailure.code', { code: failure.code })}</p>
        ) : null}
      </div>
      <p className="pptx-open-failure__status" role="status">
        {copyShown ? t(copyShown.copied ? 'openFailure.copied' : 'openFailure.copyFailed') : null}
      </p>
      {copyShown && !copyShown.copied ? (
        <textarea
          ref={reportRef}
          className="pptx-open-failure__report"
          readOnly
          rows={6}
          aria-label={t('openFailure.report')}
          value={copyShown.report}
          onFocus={(e) => e.currentTarget.select()}
        />
      ) : null}
      <div className="pptx-open-failure__actions">
        {actions?.(context)}
        {showCopy ? (
          <button
            type="button"
            className="pptx-open-failure__action"
            onClick={() => void copyReport()}
          >
            {t('openFailure.copyDetails')}
          </button>
        ) : null}
        {showClose ? (
          <button type="button" className="pptx-open-failure__action" onClick={close}>
            {closeLabel ?? t('openFailure.close')}
          </button>
        ) : null}
      </div>
    </div>
  );
}
