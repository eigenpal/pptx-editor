// The controls of the agent replay page: a scripted agent edits the presentation on the server
// in steps. Each step returns the complete PPTX file, and `createDocumentRefresh(editor)`
// applies it in place. The review module compares the files and the editor flashes what
// changed. `agent.tsx` mounts the page; the tests mount this panel with short timings.

import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  DocumentRefreshError,
  type DocumentHandle,
  createDocumentRefresh,
  describeRefreshChange,
  usePptxEditor,
  useTranslation,
  type RefreshUpdateMode,
} from '@pptx-editor.dev/react';
import { AGENT_STEPS } from './agent-steps.ts';

/** The pauses of the replay, in milliseconds. */
export interface AgentTiming {
  /** The time the "server" works on a step, at least. */
  readonly thinkMs: number;
  /** The pause after a step during Play all. */
  readonly pauseMs: number;
  /** How long the flash of a step shows. */
  readonly highlightMs: number;
}

export const AGENT_TIMING: AgentTiming = { thinkMs: 900, pauseMs: 2600, highlightMs: 2400 };

const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

/** The presentation after `step` agent steps, from the mock server. */
export async function fetchStep(step: number): Promise<Uint8Array> {
  const response = await fetch(`/api/agent?step=${step}`);
  if (!response.ok) throw new Error(`The agent step ${step} failed: ${response.status}`);
  return new Uint8Array(await response.arrayBuffer());
}

export const UNDO_REFUSED =
  'Undo step works only for the last step applied as an undoable edit, with no edits after it.';

interface StepLog {
  readonly changes: readonly { readonly id: string; readonly text: string }[];
  readonly code?: string;
}

/** An applied step: how it applied, and the document handles before and right after it. */
interface StepRecord {
  readonly step: number;
  readonly mode: RefreshUpdateMode;
  readonly before: DocumentHandle;
  readonly handle: DocumentHandle;
}

/** What `runStep()` did: applied the step, found another step running, or stopped. */
type StepOutcome = 'applied' | 'busy' | 'failed';

const sameHandle = (a: DocumentHandle, b: DocumentHandle) =>
  a.loadId === b.loadId && a.revision === b.revision;

export interface AgentPanelProps {
  readonly timing?: AgentTiming;
  /** Returns the presentation after `step` agent steps. Default: `fetchStep`. */
  readonly loadStep?: (step: number) => Promise<Uint8Array>;
}

/** The agent controls of the editor in the surrounding `PptxEditor.Root`. */
export function AgentPanel(props: AgentPanelProps) {
  const editor = usePptxEditor();
  return editor ? <Agent editor={editor} {...props} /> : null;
}

function Agent({
  editor,
  timing = AGENT_TIMING,
  loadStep = fetchStep,
}: AgentPanelProps & { editor: PptxEditorInstance }) {
  const refresh = createDocumentRefresh(editor);
  useSyncExternalStore(refresh.subscribe, refresh.snapshot);
  // Draw again after each edit: Undo step depends on the document handle.
  useSyncExternalStore(
    (listener) => editor.subscribe(listener),
    () => editor.snapshot()
  );
  const { t } = useTranslation();
  const [step, setStep] = useState(0);
  const [running, setRunning] = useState<number | null>(null);
  const [playing, setPlaying] = useState(false);
  const [restarting, setRestarting] = useState(false);
  const [mode, setMode] = useState<RefreshUpdateMode>('edit');
  const [follow, setFollow] = useState(true);
  const [logs, setLogs] = useState<Record<number, StepLog>>({});
  const [message, setMessage] = useState(
    'Select Play all to run every step, or Next step to run one.'
  );
  /** The generation of the play loop: a loop runs while it holds the current value. */
  const playId = useRef(0);
  const stepRef = useRef(0);
  const busyRef = useRef(false);
  const restartingRef = useRef(false);
  /** The step that runs now, so a play loop that finds it running waits for it. */
  const runningRef = useRef<Promise<StepOutcome> | null>(null);
  /** The checkbox takes effect during Play all, in the step that runs next. */
  const followRef = useRef(true);
  /** The applied steps, last on top: Undo step reverts only the top one. */
  const records = useRef<StepRecord[]>([]);
  const panelRef = useRef<HTMLElement>(null);

  useEffect(
    () => () => {
      // An unmount ends the play loop and the refresh.
      playId.current++;
      refresh.cancel();
    },
    [refresh]
  );

  // Keep the active step in view when the panel scrolls: the running step, then the applied
  // step after its changes open (240 ms). Only the panel scrolls, so the slide and its flash
  // stay where the reader put them.
  useEffect(() => {
    const n = running ?? step;
    if (n === 0) return;
    const timer = setTimeout(
      () => {
        const panel = panelRef.current;
        const row = panel?.querySelectorAll<HTMLElement>('.agent-step')[n - 1];
        if (!panel || !row || panel.scrollHeight <= panel.clientHeight) return;
        const p = panel.getBoundingClientRect();
        const r = row.getBoundingClientRect();
        const below = r.bottom + 12 - p.bottom;
        const above = r.top - 12 - p.top;
        const by = below > 0 ? below : above < 0 ? above : 0;
        if (by === 0) return;
        const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
        panel.scrollBy({ top: by, behavior: reduce ? 'auto' : 'smooth' });
      },
      running === null ? 260 : 0
    );
    return () => clearTimeout(timer);
  }, [running, step]);

  /** Run agent step `n` (1-based): capture, let the "server" work, apply, flash. */
  function runStep(n: number): Promise<StepOutcome> {
    if (busyRef.current) return Promise.resolve('busy');
    if (n > AGENT_STEPS.length) return Promise.resolve('failed');
    busyRef.current = true;
    const run = applyStep(n).finally(() => {
      busyRef.current = false;
      runningRef.current = null;
      setRunning(null);
    });
    runningRef.current = run;
    return run;
  }

  async function applyStep(n: number): Promise<StepOutcome> {
    setRunning(n);
    setMessage(`Step ${n} of ${AGENT_STEPS.length}: the agent works on the file.`);
    try {
      const submission = await refresh.capture();
      // An edit after the capture makes the step fail: this is the handle before the step.
      const before = editor.getDocumentHandle();
      try {
        // A real agent edits `submission.bytes`. This mock returns its scripted result.
        const [bytes] = await Promise.all([loadStep(n), wait(timing.thinkMs)]);
        // No `changes`: the review module compares the files and computes them.
        const result = await refresh.applyUpdate({ submission, sequence: 1, bytes, mode });
        if (!result.ok) {
          setLogs((l) => ({ ...l, [n]: { changes: [], code: result.code } }));
          setMessage(`The editor refused step ${n}: ${result.code}.`);
          return 'failed';
        }
        records.current.push({ step: n, mode, before, handle: editor.getDocumentHandle() });
        const fresh = result.changes.filter((c) => c.isNew);
        setLogs((l) => ({
          ...l,
          [n]: { changes: fresh.map((c) => ({ id: c.id, text: describeRefreshChange(c, t) })) },
        }));
        const first = fresh.find((c) => c.status === 'available');
        if (followRef.current && first)
          refresh.navigateToChange(first.id, { block: 'centerIfNeeded' });
        refresh.highlightChanges({ timeoutMs: timing.highlightMs });
        stepRef.current = n;
        setStep(n);
        setMessage(
          `Step ${n} applied as ${mode === 'edit' ? 'an undoable edit' : 'a reload'}. ` +
            `${fresh.length} ${fresh.length === 1 ? 'change' : 'changes'} flashed.`
        );
        return 'applied';
      } finally {
        refresh.finish(submission);
      }
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'The step failed.');
      return 'failed';
    }
  }

  async function play() {
    if (restartingRef.current || stepRef.current >= AGENT_STEPS.length) return;
    // A new generation: an older loop that wakes from its pause sees it and ends.
    const id = ++playId.current;
    setPlaying(true);
    while (playId.current === id && stepRef.current < AGENT_STEPS.length) {
      const outcome = await runStep(stepRef.current + 1);
      if (playId.current !== id || outcome === 'failed') break;
      // Another step runs, for example from before a Pause: wait for it, then go on.
      if (outcome === 'busy') {
        await runningRef.current;
        continue;
      }
      if (stepRef.current < AGENT_STEPS.length) await wait(timing.pauseMs);
    }
    if (playId.current === id) setPlaying(false);
  }

  function pause() {
    playId.current++;
    setPlaying(false);
  }

  const handle = editor.session ? editor.getDocumentHandle() : null;
  const top = records.current.at(-1);
  const undoable = !!top && top.mode === 'edit' && !!handle && sameHandle(top.handle, handle);

  function next() {
    if (busyRef.current || restartingRef.current || playing) return;
    if (stepRef.current >= AGENT_STEPS.length) return;
    void runStep(stepRef.current + 1);
  }

  function undo() {
    if (busyRef.current || restartingRef.current || playing) return;
    if (!undoable) {
      setMessage(UNDO_REFUSED);
      return;
    }
    if (!editor.exec({ type: 'undo' }).ok) {
      setMessage('The editor could not undo the step.');
      return;
    }
    refresh.clearHighlights({ animation: false });
    const undone = records.current.pop()!;
    // With no edit between the two steps, the document is again as the previous step left it:
    // that step is the current one. Else its handle stays old, and Undo step refuses it.
    const below = records.current.at(-1);
    if (below && sameHandle(undone.before, below.handle))
      records.current[records.current.length - 1] = {
        ...below,
        handle: editor.getDocumentHandle(),
      };
    const n = stepRef.current;
    setLogs((l) => {
      const rest = { ...l };
      delete rest[n];
      return rest;
    });
    stepRef.current = n - 1;
    setStep(n - 1);
    setMessage(`Undid step ${n}. Select Next step to apply it again.`);
  }

  /** Back to the first file, through the same refresh: no loading screen. */
  async function restart() {
    if (busyRef.current || restartingRef.current) return;
    pause();
    busyRef.current = true;
    restartingRef.current = true;
    setRestarting(true);
    try {
      const submission = await refresh.capture();
      try {
        const bytes = await loadStep(0);
        const result = await refresh.applyUpdate({ submission, sequence: 1, bytes });
        refresh.clearHighlights({ animation: false });
        if (!result.ok) {
          setMessage(`Restart failed: ${result.code}. Reload the page.`);
          return;
        }
        editor.goToSlide(0);
        records.current = [];
        stepRef.current = 0;
        setStep(0);
        setLogs({});
        setMessage('Restarted. Select Play all or Next step.');
      } finally {
        refresh.finish(submission);
      }
    } catch (error) {
      const reason =
        error instanceof DocumentRefreshError
          ? error.code
          : error instanceof Error
            ? error.message
            : 'unknown error';
      setMessage(`Restart failed: ${reason}. Reload the page.`);
    } finally {
      busyRef.current = false;
      restartingRef.current = false;
      setRestarting(false);
    }
  }

  const busy = running !== null || restarting;
  const done = step >= AGENT_STEPS.length;
  return (
    <aside className="refresh-controls" aria-labelledby="agent-title" ref={panelRef}>
      <div className="refresh-intro">
        <h1 id="agent-title">Agent replay</h1>
        <p>
          An agent edits the presentation on the server, one step at a time. Each step returns the
          whole PPTX file. The editor applies it in place and flashes what changed.
        </p>
      </div>
      <div className="refresh-settings">
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
            onChange={(event) => {
              followRef.current = event.target.checked;
              setFollow(event.target.checked);
            }}
          />
          Go to each change
        </label>
      </div>
      <div className="refresh-actions">
        {/* Play all and Pause share the primary place, so the swap moves no button. */}
        {playing ? (
          <button type="button" className="refresh-primary" onClick={pause}>
            Pause
          </button>
        ) : (
          // Play all during a step that runs, after a Pause, waits for that step.
          <button
            type="button"
            className="refresh-primary"
            aria-disabled={restarting || done}
            onClick={() => void play()}
          >
            Play all
          </button>
        )}
        <button type="button" aria-disabled={busy || playing || done} onClick={next}>
          Next step
        </button>
        <button type="button" aria-disabled={busy || playing || !undoable} onClick={undo}>
          Undo step
        </button>
        <button type="button" aria-disabled={busy} onClick={() => void restart()}>
          Restart
        </button>
      </div>
      <p role="status" aria-live="polite" className="refresh-status">
        {message}
      </p>
      <section className="agent-progress" aria-labelledby="agent-steps-title">
        <h2 id="agent-steps-title">
          Steps
          <span className="refresh-count">
            {step} of {AGENT_STEPS.length} applied
          </span>
        </h2>
        <ol className="agent-steps">
          {AGENT_STEPS.map((s, i) => {
            const n = i + 1;
            const state = running === n ? 'running' : n <= step ? 'done' : 'pending';
            const log = logs[n];
            return (
              <li
                key={s.title}
                className="agent-step"
                data-state={state}
                aria-current={state === 'running' ? 'step' : undefined}
              >
                {/* The key replays the entrance of the glyph when the state changes. */}
                <span className="agent-step__mark" aria-hidden="true">
                  {state === 'done' ? (
                    <Check key="done" />
                  ) : state === 'pending' ? (
                    <span key="pending">{n}</span>
                  ) : null}
                </span>
                <div className="agent-step__body">
                  <span className="agent-step__state">
                    Step {n} · {STATE_TEXT[state]}
                  </span>
                  <strong className="agent-step__title">{s.title}</strong>
                  <span className="agent-step__detail">{s.detail}</span>
                  {log && (log.changes.length > 0 || log.code) && (
                    <div className="agent-step__reveal">
                      <div>
                        {log.changes.length > 0 && (
                          <ul className="agent-step__changes" aria-label="Changes">
                            {log.changes.map((change) => (
                              <li key={change.id}>{change.text}</li>
                            ))}
                          </ul>
                        )}
                        {log.code && <code>{log.code}</code>}
                      </div>
                    </div>
                  )}
                </div>
              </li>
            );
          })}
        </ol>
      </section>
    </aside>
  );
}

const STATE_TEXT = { pending: 'Waiting', running: 'Working', done: 'Applied' } as const;

/** The check mark of an applied step. */
function Check() {
  return (
    <svg width="12" height="12" viewBox="0 0 12 12" fill="none" aria-hidden="true">
      <path
        d="M2.5 6.25 5 8.75 9.5 3.5"
        stroke="currentColor"
        strokeWidth="1.75"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}
