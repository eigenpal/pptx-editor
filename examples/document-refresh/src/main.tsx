// The document refresh example, after DOCX Editor's examples/document-refresh/src/main.tsx: a
// mock processing server returns complete PPTX files, and `createDocumentRefresh(editor)` shows
// them without a new editor, a loading screen, or a jump away from the current slide.

import { StrictMode, useEffect, useRef, useState, useSyncExternalStore } from 'react';
import { createRoot } from 'react-dom/client';
import { packagedFonts } from '@docx-editor.dev/fonts';
import type { EditorModule, PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { reviewModule } from '@pptx-editor.dev/pro';
import {
  DocumentRefreshError,
  PptxEditor,
  createDocumentRefresh,
  describeRefreshChange,
  useEditorEvent,
  useFonts,
  usePptxEditor,
  useTranslation,
  type DocumentRefresh,
  type RefreshChange,
  type RefreshResult,
  type RefreshUpdateMode,
} from '@pptx-editor.dev/react';
import { download } from '../../shared/demo-files.ts';
import { receiveSampleUpdates, runRefreshJob } from './refresh-job.ts';
import { TEXT as t } from './text.ts';
import './styles.css';

// Without `changes`, the review module computes them (EigenPal Pro License). Evaluation use
// needs no key.
const MODULES: readonly EditorModule[] = [
  reviewModule({ licenseKey: import.meta.env.VITE_PPTX_LICENSE }),
];
const HIGHLIGHT_MS = 5000;

function UpdateControls() {
  const editor = usePptxEditor();
  return editor ? <ReadyControls editor={editor} /> : null;
}

/** The status text for one result. */
function resultMessage(result: RefreshResult, sequence: number): string {
  if (result.ok) {
    const fresh = result.changes.filter((change) => change.isNew).length;
    return t.applied(sequence, result.mode === 'edit' ? 'an undoable edit' : 'a reload', fresh);
  }
  if (result.code === 'local-edits') return t.refusedEdits;
  if (result.code === 'out-of-order') return t.refusedLate;
  return t.refused;
}

/** The result code beside the status text: `ok` and the mode, or the refusal code and cause. */
function resultCode(result: RefreshResult | null): string | null {
  if (!result) return null;
  if (result.ok) return `ok (${result.mode})`;
  return result.reason ? `${result.code} (${result.reason})` : result.code;
}

function ReadyControls({ editor }: { editor: PptxEditorInstance }) {
  const refresh = createDocumentRefresh(editor);
  const state = useSyncExternalStore(refresh.subscribe, refresh.snapshot);
  const { t: translate } = useTranslation();
  const describe = (change: RefreshChange) => describeRefreshChange(change, translate);
  const request = useRef<AbortController | null>(null);
  const round = useRef(0);
  const sampleEditedRef = useRef(false);
  // The changes that were new in any result of the running submission. Each result highlights
  // all of them, so a later result does not cut short the flash of an earlier one.
  const submissionChanges = useRef(new Set<string>());
  const [busy, setBusy] = useState(false);
  const [mode, setMode] = useState<RefreshUpdateMode>('reload');
  const [computeChanges, setComputeChanges] = useState(false);
  const [goToChange, setGoToChange] = useState(false);
  const [includeLateResult, setIncludeLateResult] = useState(false);
  const [sampleEdited, setSampleEdited] = useState(false);
  const [message, setMessage] = useState(t.idle);
  // The latest reported result, refusals included. `snapshot().result` keeps the accepted
  // result when a late result is refused.
  const [lastResult, setLastResult] = useState<RefreshResult | null>(null);

  useEffect(
    () => () => {
      request.current?.abort();
      refresh.cancel();
    },
    [refresh]
  );

  // This fixed fixture demo requires a reset after any user edit. Loads and refresh results
  // have a `source`; edits, undo, and redo do not.
  useEditorEvent('change', (_snapshot, change) => {
    if (change.source) return;
    sampleEditedRef.current = true;
    setSampleEdited(true);
    if (!request.current) setMessage(t.sampleEdited);
  });

  function report(result: RefreshResult, sequence: number) {
    setLastResult(result);
    setMessage(resultMessage(result, sequence));
    if (!result.ok) return;
    for (const c of result.changes) if (c.isNew) submissionChanges.current.add(c.id);
    const change = result.changes.find((c) => c.isNew && c.status === 'available');
    const moved =
      goToChange && change && refresh.navigateToChange(change.id, { block: 'centerIfNeeded' });
    refresh.highlightChanges({
      changeIds: [...submissionChanges.current],
      timeoutMs: HIGHLIGHT_MS,
    });
    if (moved) setMessage(`${resultMessage(result, sequence)} ${t.movedToChange}`);
  }

  async function update() {
    if (request.current || sampleEditedRef.current) return;
    const controller = new AbortController();
    request.current = controller;
    submissionChanges.current.clear();
    setBusy(true);
    try {
      setLastResult(null);
      setMessage(t.processing);
      await runRefreshJob(
        refresh,
        controller.signal,
        (submission) =>
          receiveSampleUpdates(submission, controller.signal, ++round.current, includeLateResult),
        (result, { sequence }) => report(result, sequence),
        { mode, computeChanges }
      );
    } catch (error) {
      if (!controller.signal.aborted) {
        setMessage(
          error instanceof DocumentRefreshError ? `${t.failed} (${error.code})` : t.failed
        );
      }
    } finally {
      if (request.current === controller) {
        request.current = null;
        setBusy(false);
      }
    }
  }

  function cancel() {
    if (!request.current) return;
    request.current.abort();
    refresh.cancel();
    setMessage(t.cancelled);
  }

  const code = resultCode(lastResult);
  const failures = state.result?.ok ? state.result.failures : [];
  return (
    <aside className="refresh-controls" aria-labelledby="refresh-title">
      <h1 id="refresh-title">{t.title}</h1>
      <p>{t.description}</p>
      <p>{t.limitations}</p>
      <fieldset disabled={busy}>
        <legend>{t.modeLabel}</legend>
        {(['reload', 'edit'] as const).map((value) => (
          <label key={value}>
            <input
              type="radio"
              name="refresh-mode"
              value={value}
              checked={mode === value}
              onChange={() => setMode(value)}
            />
            {value === 'reload' ? t.modeReload : t.modeEdit}
          </label>
        ))}
      </fieldset>
      <label>
        <input
          type="checkbox"
          checked={computeChanges}
          disabled={busy}
          onChange={(event) => setComputeChanges(event.target.checked)}
        />
        {t.computeChanges}
      </label>
      <label>
        <input
          type="checkbox"
          checked={goToChange}
          disabled={busy}
          onChange={(event) => setGoToChange(event.target.checked)}
        />
        {t.goToChange}
      </label>
      <label>
        <input
          type="checkbox"
          checked={includeLateResult}
          disabled={busy}
          onChange={(event) => setIncludeLateResult(event.target.checked)}
        />
        {t.lateResult}
      </label>
      <div className="refresh-actions">
        {/* `aria-disabled` keeps these buttons focusable, so focus does not fall to the page
            when a job starts or ends. The handlers ignore a press while they are disabled. The
            labels stay the same, so the buttons do not move; the status line shows the job. */}
        <button type="button" aria-disabled={busy || sampleEdited} onClick={() => void update()}>
          {t.start}
        </button>
        <button type="button" aria-disabled={!busy} onClick={cancel}>
          {t.cancel}
        </button>
        <button type="button" disabled={busy} onClick={() => location.reload()}>
          {t.reset}
        </button>
      </div>
      <p role="status" aria-live="polite" className="refresh-status">
        {message}
        {code && (
          <>
            {' '}
            <span className="refresh-code">
              {t.resultCode}: <code>{code}</code>
            </span>
          </>
        )}
      </p>
      {failures.length > 0 && (
        <p>
          {t.partial(failures.length)}{' '}
          {failures.map((id) => (
            <code key={id}>{id}</code>
          ))}
        </p>
      )}
      {state.result?.ok && state.result.changeInformation === 'unavailable' && (
        <p>{t.noChangeInformation}</p>
      )}
      {state.recoveryAvailable && <Recovery refresh={refresh} busy={busy} />}
      {state.changes.length > 0 && (
        <details open className="refresh-review">
          <summary>{t.reviewChanges}</summary>
          <ul className="refresh-change-list">
            {state.changes.map((change) => (
              <li key={change.id}>
                <span className="refresh-change-text">{describe(change)}</span>
                <span className="refresh-badges">
                  <span className="refresh-badge">{t.changeStatus[change.status]}</span>
                  {change.isNew && <span className="refresh-badge">{t.newChange}</span>}
                  {change.computed && <span className="refresh-badge">{t.computed}</span>}
                  {change.diagnostic && change.diagnostic.code !== change.status && (
                    <code>{change.diagnostic.code}</code>
                  )}
                </span>
                {change.status === 'available' && (
                  <button
                    type="button"
                    aria-label={`${t.showChange}: ${describe(change)}`}
                    onClick={() => {
                      if (!refresh.navigateToChange(change.id, { block: 'centerIfNeeded' })) return;
                      setMessage(t.showingSlide(editor.snapshot().currentSlide + 1));
                      refresh.highlightChanges({
                        changeIds: [change.id],
                        timeoutMs: HIGHLIGHT_MS,
                      });
                    }}
                  >
                    {t.showChange}
                  </button>
                )}
              </li>
            ))}
          </ul>
        </details>
      )}
    </aside>
  );
}

function Recovery({ refresh, busy }: { refresh: DocumentRefresh; busy: boolean }) {
  return (
    <div role="group" aria-label={t.recoveryTitle} className="refresh-actions">
      <p>{t.recoveryTitle}</p>
      <button type="button" disabled={busy} onClick={() => void refresh.recover()}>
        {t.retry}
      </button>
      <button
        type="button"
        onClick={() => {
          const bytes = refresh.recoveryBytes();
          if (bytes) download(new Uint8Array(bytes), 'recovery.pptx');
        }}
      >
        {t.download}
      </button>
    </div>
  );
}

function App() {
  const [bytes, setBytes] = useState<Uint8Array | null>(null);
  const [error, setError] = useState('');
  const fonts = useFonts(packagedFonts());
  useEffect(() => {
    const controller = new AbortController();
    async function load() {
      try {
        const response = await fetch('/api/presentation', { signal: controller.signal });
        if (!response.ok) throw new Error('load-failed');
        const presentation = new Uint8Array(await response.arrayBuffer());
        if (!controller.signal.aborted) setBytes(presentation);
      } catch {
        if (!controller.signal.aborted) setError(t.loadError);
      }
    }
    void load();
    return () => controller.abort();
  }, []);
  if (!bytes) return <p role="status">{error || t.loading}</p>;
  return (
    <PptxEditor.Root
      document={bytes}
      colorMode="system"
      fonts={fonts}
      modules={MODULES}
      onReady={(instance: PptxEditorInstance) => {
        // The verification scripts drive the editor through this handle.
        (window as unknown as { __pptxEditor?: unknown }).__pptxEditor = instance;
      }}
    >
      <PptxEditor.Viewport className="refresh-app">
        <PptxEditor.Toolbar />
        <div className="refresh-workspace">
          <PptxEditor.SlideRail />
          <main className="refresh-docpane">
            <div className="refresh-canvas">
              <PptxEditor.Content />
              <PptxEditor.Loading overlay />
              <PptxEditor.FontNotice />
            </div>
            <PptxEditor.NotesPane />
            <PptxEditor.EditNotice />
            <PptxEditor.OpenFailure />
          </main>
          <UpdateControls />
        </div>
      </PptxEditor.Viewport>
    </PptxEditor.Root>
  );
}

const container = document.getElementById('app');
if (container)
  createRoot(container).render(
    <StrictMode>
      <App />
    </StrictMode>
  );
