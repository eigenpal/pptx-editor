// The agent replay page: a scripted agent edits the presentation on the server in steps. Each
// step returns the complete PPTX file, and `createDocumentRefresh(editor)` applies it in place.
// `reviewModule()` compares the files and the editor flashes what changed. Step through the
// edits one at a time, or play them all.

import { StrictMode, useEffect, useRef, useState, useSyncExternalStore } from 'react';
import { createRoot } from 'react-dom/client';
import { packagedFonts } from '@docx-editor.dev/fonts';
import type { EditorModule, PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { reviewModule } from '@pptx-editor.dev/pro';
import {
  PptxEditor,
  createDocumentRefresh,
  describeRefreshChange,
  useFonts,
  usePptxEditor,
  useTranslation,
  type RefreshUpdateMode,
} from '@pptx-editor.dev/react';
import { AGENT_STEPS } from './agent-steps.ts';
import './styles.css';

// The review module computes the changes of each step (EigenPal Pro License). Evaluation use
// needs no key.
const MODULES: readonly EditorModule[] = [
  reviewModule({ licenseKey: import.meta.env.VITE_PPTX_LICENSE }),
];
const THINK_MS = 900;
const PAUSE_MS = 2600;
const HIGHLIGHT_MS = 2400;

const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

async function fetchStep(step: number): Promise<Uint8Array> {
  const response = await fetch(`/api/agent?step=${step}`);
  if (!response.ok) throw new Error(`The agent step ${step} failed: ${response.status}`);
  return new Uint8Array(await response.arrayBuffer());
}

interface StepLog {
  readonly summaries: readonly string[];
  readonly code?: string;
}

function AgentPanel() {
  const editor = usePptxEditor();
  return editor ? <Agent editor={editor} /> : null;
}

function Agent({ editor }: { editor: PptxEditorInstance }) {
  const refresh = createDocumentRefresh(editor);
  useSyncExternalStore(refresh.subscribe, refresh.snapshot);
  const { t } = useTranslation();
  const [step, setStep] = useState(0);
  const [running, setRunning] = useState<number | null>(null);
  const [playing, setPlaying] = useState(false);
  const [mode, setMode] = useState<RefreshUpdateMode>('edit');
  const [follow, setFollow] = useState(true);
  const [logs, setLogs] = useState<Record<number, StepLog>>({});
  const [message, setMessage] = useState('Select Next step to let the agent make its first edit.');
  const playRef = useRef(false);
  const stepRef = useRef(0);
  const busyRef = useRef(false);

  useEffect(() => () => refresh.cancel(), [refresh]);

  /** Run agent step `n` (1-based): capture, let the "server" work, apply, flash. */
  async function runStep(n: number): Promise<boolean> {
    if (busyRef.current || n > AGENT_STEPS.length) return false;
    busyRef.current = true;
    setRunning(n);
    setMessage(`Agent: ${AGENT_STEPS[n - 1]!.detail}`);
    try {
      const submission = await refresh.capture();
      try {
        // A real agent edits `submission.bytes`. This mock returns its scripted result.
        const [bytes] = await Promise.all([fetchStep(n), wait(THINK_MS)]);
        // No `changes`: the review module compares the files and computes them.
        const result = await refresh.applyUpdate({ submission, sequence: 1, bytes, mode });
        if (!result.ok) {
          setLogs((l) => ({ ...l, [n]: { summaries: [], code: result.code } }));
          setMessage(`The editor refused step ${n}: ${result.code}.`);
          return false;
        }
        const fresh = result.changes.filter((c) => c.isNew);
        setLogs((l) => ({
          ...l,
          [n]: { summaries: fresh.map((c) => describeRefreshChange(c, t)) },
        }));
        const first = fresh.find((c) => c.status === 'available');
        if (follow && first) refresh.navigateToChange(first.id, { block: 'centerIfNeeded' });
        refresh.highlightChanges({ timeoutMs: HIGHLIGHT_MS });
        stepRef.current = n;
        setStep(n);
        setMessage(
          `Step ${n} applied as ${mode === 'edit' ? 'an undoable edit' : 'a reload'}. ` +
            `${fresh.length} ${fresh.length === 1 ? 'change' : 'changes'} flashed.`
        );
        return true;
      } finally {
        refresh.finish(submission);
      }
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'The step failed.');
      return false;
    } finally {
      busyRef.current = false;
      setRunning(null);
    }
  }

  async function play() {
    if (playRef.current) return;
    playRef.current = true;
    setPlaying(true);
    while (playRef.current && stepRef.current < AGENT_STEPS.length) {
      if (!(await runStep(stepRef.current + 1))) break;
      if (stepRef.current < AGENT_STEPS.length) await wait(PAUSE_MS);
    }
    playRef.current = false;
    setPlaying(false);
  }

  function pause() {
    playRef.current = false;
    setPlaying(false);
  }

  function undo() {
    if (busyRef.current || stepRef.current === 0 || mode !== 'edit') return;
    if (!editor.exec({ type: 'undo' }).ok) return;
    refresh.clearHighlights({ animation: false });
    const n = stepRef.current;
    setLogs((l) => {
      const next = { ...l };
      delete next[n];
      return next;
    });
    stepRef.current = n - 1;
    setStep(n - 1);
    setMessage(`Undid step ${n}. Select Next step to apply it again.`);
  }

  /** Back to the first file, through the same refresh: no loading screen. */
  async function restart() {
    pause();
    if (busyRef.current) return;
    busyRef.current = true;
    try {
      const submission = await refresh.capture();
      try {
        const bytes = await fetchStep(0);
        const result = await refresh.applyUpdate({ submission, sequence: 1, bytes });
        refresh.clearHighlights({ animation: false });
        if (!result.ok) {
          setMessage(`Restart failed: ${result.code}. Reload the page.`);
          return;
        }
        editor.goToSlide(0);
        stepRef.current = 0;
        setStep(0);
        setLogs({});
        setMessage('Restarted. Select Next step or Play all.');
      } finally {
        refresh.finish(submission);
      }
    } finally {
      busyRef.current = false;
    }
  }

  const busy = running !== null;
  const done = step >= AGENT_STEPS.length;
  return (
    <aside className="refresh-controls" aria-labelledby="agent-title">
      <h1 id="agent-title">Agent replay</h1>
      <p>
        An agent edits the presentation on the server, one step at a time. Each step returns the
        whole PPTX file. The editor applies it in place and flashes what changed.
      </p>
      <fieldset disabled={busy || playing}>
        <legend>Apply each step as</legend>
        {(['edit', 'reload'] as const).map((value) => (
          <label key={value}>
            <input
              type="radio"
              name="agent-mode"
              checked={mode === value}
              onChange={() => setMode(value)}
            />
            {value === 'edit' ? 'Undoable edit' : 'Reload'}
          </label>
        ))}
      </fieldset>
      <label>
        <input
          type="checkbox"
          checked={follow}
          onChange={(event) => setFollow(event.target.checked)}
        />
        Go to each change
      </label>
      <div className="refresh-actions">
        <button
          type="button"
          aria-disabled={busy || playing || done}
          onClick={() => void runStep(stepRef.current + 1)}
        >
          Next step
        </button>
        {playing ? (
          <button type="button" onClick={pause}>
            Pause
          </button>
        ) : (
          <button type="button" aria-disabled={busy || done} onClick={() => void play()}>
            Play all
          </button>
        )}
        <button
          type="button"
          aria-disabled={busy || playing || step === 0 || mode !== 'edit'}
          onClick={undo}
        >
          Undo step
        </button>
        <button type="button" aria-disabled={busy} onClick={() => void restart()}>
          Restart
        </button>
      </div>
      <p role="status" aria-live="polite" className="refresh-status">
        {message}
      </p>
      <ol className="agent-steps">
        {AGENT_STEPS.map((s, i) => {
          const n = i + 1;
          const state = running === n ? 'running' : n <= step ? 'done' : 'pending';
          const log = logs[n];
          return (
            <li key={s.title} className="agent-step" data-state={state}>
              <span className="agent-step__mark" aria-hidden="true">
                {state === 'done' ? '✓' : state === 'running' ? '…' : n}
              </span>
              <div>
                <strong>{s.title}</strong>
                <span className="agent-step__state">
                  {state === 'running' ? ' (working)' : state === 'done' ? ' (applied)' : ''}
                </span>
                <p>{s.detail}</p>
                {log && log.summaries.length > 0 && (
                  <ul className="agent-step__changes">
                    {log.summaries.map((text) => (
                      <li key={text}>{text}</li>
                    ))}
                  </ul>
                )}
                {log?.code && <code>{log.code}</code>}
              </div>
            </li>
          );
        })}
      </ol>
    </aside>
  );
}

function App() {
  const [bytes, setBytes] = useState<Uint8Array | null>(null);
  const [error, setError] = useState('');
  const fonts = useFonts(packagedFonts());
  useEffect(() => {
    fetchStep(0).then(setBytes, () => setError('The presentation did not load.'));
  }, []);
  if (!bytes) return <p role="status">{error || 'Loading the presentation'}</p>;
  return (
    <PptxEditor.Root
      document={bytes}
      colorMode="system"
      fonts={fonts}
      modules={MODULES}
      onReady={(instance: PptxEditorInstance) => {
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
            </div>
            <PptxEditor.NotesPane />
          </main>
          <AgentPanel />
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
